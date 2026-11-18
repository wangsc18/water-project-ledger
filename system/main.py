# -*- coding: utf-8 -*-
"""
水利新建工程全流程电子台账系统 — 主程序入口
支持命令：
  scan   : 立即扫描收件箱归档 + 物理文件夹双向校准 + 刷新双层Excel与HTML看板
  watch  : 常驻后台监听（单实例防重锁、零CPU空闲指纹检测、自动感知配置热更新）
  stop   : 优雅停止正在后台运行的监听进程
  demo   : 向收件箱生成全套水利工程仿真样板文件并执行全自动归档出表
  reset  : 清空所有归档库文件、待确认文件与数据库，恢复纯净空库状态
"""
from __future__ import annotations

import argparse
import os
import shutil
import sys
import time
from pathlib import Path
from typing import Optional, Tuple

# 强制将标准输出/错误流设为 UTF-8 + replace 容错，彻底杜绝英文版 Windows (cp1252) 或传统控制台 (cp936) 下打印中文崩溃
for _sname in ("stdout", "stderr"):
    _s = getattr(sys, _sname, None)
    if _s is not None and hasattr(_s, "reconfigure"):
        try:
            _s.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

# 确保模块搜索路径包含 system/ 目录（兼容外部工作目录调用）
SYS_DIR = Path(__file__).resolve().parent
if str(SYS_DIR) not in sys.path:
    sys.path.insert(0, str(SYS_DIR))

from config_loader import (
    CONFIG_FILE_NAME,
    AppConfig,
    ensure_directory_skeleton,
    load_config,
)
from database import LedgerDB
from excel_builder import build_all_excel_ledgers, flush_deferred_workbooks
from extractor import should_ignore_file
from html_builder import build_dashboard_html
from notifier import send_windows_notification
from progress_calc import compute_portfolio_summary
from reconciler import reconcile_once


def _compute_workspace_fingerprint(cfg: AppConfig) -> Tuple:
    """
    轻量级毫秒级目录状态快照（仅读取目录项 stat，零文件内容读取）：
    监控 收件箱、项目归档库、99_待人工确认、项目配置.yaml 以及台账 Excel 手工保存时间戳
    """
    items = []
    cfg_file = cfg.root_dir / CONFIG_FILE_NAME
    if cfg_file.exists():
        st = cfg_file.stat()
        items.append(("__cfg__", st.st_size, st.st_mtime_ns))

    if cfg.master_excel_path.exists():
        st = cfg.master_excel_path.stat()
        items.append(("__master_xl__", st.st_size, st.st_mtime_ns))

    for base in (cfg.inbox_dir, cfg.archive_lib_dir, cfg.pending_dir):
        if not base.exists():
            continue
        for root, dirs, files in os.walk(base):
            dirs[:] = sorted(d for d in dirs if not d.startswith((".", "_", "~")))
            for fn in sorted(files):
                fp = Path(root) / fn
                if should_ignore_file(fp):
                    continue
                try:
                    st = fp.stat()
                    items.append((str(fp), st.st_size, st.st_mtime_ns))
                except OSError:
                    continue
    return tuple(items)


def _is_pid_alive(pid: int) -> bool:
    if pid <= 0:
        return False
    if sys.platform == "win32":
        import ctypes

        PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
        handle = ctypes.windll.kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
        if handle:
            ctypes.windll.kernel32.CloseHandle(handle)
            return True
        return False
    else:
        try:
            os.kill(pid, 0)
            return True
        except OSError:
            return False


def cmd_scan(explicit_root: Optional[str] = None, force_rebuild: bool = True) -> int:
    cfg = load_config(explicit_root)
    db = LedgerDB(cfg.db_path)
    print(f"=== 水利工程电子台账单次扫描校准 ({cfg.root_dir}) ===")
    rep = reconcile_once(cfg, db, force_rebuild=force_rebuild, verbose=True)
    print(
        f"=== 完成：新归档 {rep.inbox_archived} 份 | 待确认 {rep.inbox_pending} 份 | "
        f"目录同步 {rep.explorer_moved_or_renamed + rep.explorer_direct_added + rep.explorer_deleted} 项 | "
        f"Excel手填保留 {rep.manual_edits_synced} 条 ==="
    )
    return 0


def cmd_watch(explicit_root: Optional[str] = None) -> int:
    cfg = load_config(explicit_root)
    ensure_directory_skeleton(cfg)
    pid_file = cfg.data_dir / "watcher.pid"
    stop_file = cfg.data_dir / "stop.signal"

    if stop_file.exists():
        try:
            stop_file.unlink()
        except OSError:
            pass

    if pid_file.exists():
        try:
            old_pid = int(pid_file.read_text(encoding="utf-8").strip())
            if old_pid != os.getpid() and _is_pid_alive(old_pid):
                print(f"[提示] 后台自动监控已在运行中 (PID={old_pid})，无需重复启动。")
                send_windows_notification(
                    "电子台账后台监控已在运行中",
                    "请直接将工程文件拖入 00_待归档收件箱，无需重复开启。",
                    enabled=cfg.enable_windows_notification,
                )
                return 0
        except Exception:
            pass

    pid_file.write_text(str(os.getpid()), encoding="utf-8")
    db = LedgerDB(cfg.db_path)

    print("============================================================")
    print("  水利新建工程全流程电子台账 — 正在实时自动监听中...")
    print(f"  投递入口：{cfg.inbox_dir}")
    print("  （关闭本窗口或双击《停止后台监控.bat》可停止监听）")
    print("============================================================")
    send_windows_notification(
        "水利工程电子台账监控已启动",
        "现在起将工程文件拖入 00_待归档收件箱 即可全自动识别、重命名归档并刷新台账。",
        enabled=cfg.enable_windows_notification,
    )

    try:
        reconcile_once(cfg, db, force_rebuild=False, verbose=True)
        last_fp = _compute_workspace_fingerprint(cfg)

        while True:
            if stop_file.exists():
                try:
                    stop_file.unlink()
                except OSError:
                    pass
                print("[停止] 收到停止信号，后台监听已安全退出。")
                break

            flush_deferred_workbooks(cfg)
            curr_fp = _compute_workspace_fingerprint(cfg)
            if curr_fp != last_fp:
                # 热加载最新配置（支持用户随时用记事本修改 项目配置.yaml 立即生效）
                cfg = load_config(explicit_root)
                reconcile_once(cfg, db, force_rebuild=False, verbose=True)
                last_fp = _compute_workspace_fingerprint(cfg)

            time.sleep(max(1.0, cfg.watch_interval_seconds))
    except KeyboardInterrupt:
        print("\n[停止] 用户中断退出。")
    finally:
        try:
            if pid_file.exists() and pid_file.read_text(encoding="utf-8").strip() == str(os.getpid()):
                pid_file.unlink()
        except OSError:
            pass

    return 0


def cmd_stop(explicit_root: Optional[str] = None) -> int:
    cfg = load_config(explicit_root)
    cfg.data_dir.mkdir(parents=True, exist_ok=True)
    pid_file = cfg.data_dir / "watcher.pid"
    stop_file = cfg.data_dir / "stop.signal"

    if not pid_file.exists():
        print("[提示] 当前没有正在运行的后台监控进程。")
        return 0

    stop_file.write_text("stop", encoding="utf-8")
    for _ in range(15):
        if not pid_file.exists():
            print("[成功] 后台自动监控进程已安全停止。")
            send_windows_notification("电子台账后台监控已停止", "如需恢复自动归档，请双击启动脚本。", enabled=cfg.enable_windows_notification)
            return 0
        time.sleep(0.4)

    # 超时强制清理陈旧锁文件
    try:
        pid_file.unlink()
        if stop_file.exists():
            stop_file.unlink()
    except OSError:
        pass
    print("[成功] 后台监控状态已重置。")
    return 0


def cmd_demo(explicit_root: Optional[str] = None) -> int:
    from demo_seeder import seed_demo_files_to_inbox

    cfg = load_config(explicit_root)
    ensure_directory_skeleton(cfg)
    count = seed_demo_files_to_inbox(cfg.inbox_dir)
    print(f"[演示生成] 已向 00_待归档收件箱 投入 {count} 份水利工程多格式仿真文件，开始自动识别归档...")
    db = LedgerDB(cfg.db_path)
    # 样板本地刚写完，无需长等待文件锁
    reconcile_once(cfg, db, force_rebuild=True, check_file_ready_delay=0.05, verbose=True)
    return 0


def cmd_reset(explicit_root: Optional[str] = None, assume_yes: bool = False) -> int:
    cfg = load_config(explicit_root)
    if not assume_yes:
        ans = input("确认要清空所有归档演示文件、待确认文件并重置为空白电子台账吗？输入 YES 确认: ").strip()
        if ans.upper() != "YES":
            print("已取消清空操作。")
            return 0

    # 清空收件箱（保留 _ 开头说明文件）、归档库与待确认目录
    if cfg.inbox_dir.exists():
        for item in cfg.inbox_dir.iterdir():
            if item.name.startswith("_"):
                continue
            if item.is_dir():
                shutil.rmtree(item, ignore_errors=True)
            else:
                item.unlink(missing_ok=True)

    if cfg.archive_lib_dir.exists():
        shutil.rmtree(cfg.archive_lib_dir, ignore_errors=True)
    if cfg.pending_dir.exists():
        shutil.rmtree(cfg.pending_dir, ignore_errors=True)

    db = LedgerDB(cfg.db_path)
    db.clear_all()
    ensure_directory_skeleton(cfg)

    summary = compute_portfolio_summary(cfg, [])
    build_dashboard_html(cfg, summary)
    build_all_excel_ledgers(cfg, summary)
    print("[重置完成] 已恢复为纯净空库状态，空白目录骨架、总台账 Excel 与可视化看板已就绪！")
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description="水利新建工程全流程电子台账引擎")
    parser.add_argument(
        "action",
        nargs="?",
        default="scan",
        choices=["scan", "watch", "stop", "demo", "reset"],
        help="执行动作：scan(单次扫描校准) / watch(常驻监听) / stop(停止监听) / demo(注入演示数据) / reset(重置空库)",
    )
    parser.add_argument("--root", type=str, default=None, help="指定台账根目录路径")
    parser.add_argument("--yes", action="store_true", help="reset 时跳过二次确认")
    args = parser.parse_args()

    if args.action == "scan":
        return cmd_scan(args.root, force_rebuild=True)
    elif args.action == "watch":
        return cmd_watch(args.root)
    elif args.action == "stop":
        return cmd_stop(args.root)
    elif args.action == "demo":
        return cmd_demo(args.root)
    elif args.action == "reset":
        return cmd_reset(args.root, assume_yes=args.yes)
    return 0


if __name__ == "__main__":
    sys.exit(main())

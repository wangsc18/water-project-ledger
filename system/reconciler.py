# -*- coding: utf-8 -*-
"""
收件箱自动归档与物理文件夹双向校准核心协调器
完整实现：
1. 无锁回读用户在 Excel 手填的“文号/编号、涉及金额、人工备注”
2. 扫描 00_待归档收件箱 -> 提取/OCR -> 智能分类 -> 三段式规范重命名（同名顺延 -02，不拦截重复件） -> 移入阶段目录
3. 对 99_待人工确认 文件重新检测项目白名单（当用户在配置里补充新项目关键词后自动转正归档）
4. 扫描 项目归档库 与 99_待人工确认 实际物理现状：自动追踪资源管理器中的移动目录、改名、直接放入及删除操作
5. 同步刷新双层超链接 Excel 台账与静态 HTML 状态可视化看板
"""
from __future__ import annotations

import datetime as dt
import os
import shutil
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from classifier import (
    classify_document,
    infer_milestone_from_type,
    parse_standard_filename,
    sanitize_doc_type,
)
from config_loader import (
    ARCHIVE_LIBRARY_DIR_NAME,
    PENDING_DIR_NAME,
    STAGES_ORDER,
    AppConfig,
    ensure_directory_skeleton,
)
from database import LedgerDB, LedgerRecord, compute_sha256
from excel_builder import (
    build_all_excel_ledgers,
    flush_deferred_workbooks,
    read_back_manual_edits_from_excels,
)
from extractor import extract_document, is_file_ready, should_ignore_file
from html_builder import build_dashboard_html
from notifier import send_windows_notification
from progress_calc import compute_portfolio_summary


@dataclass
class ReconcileReport:
    manual_edits_synced: int = 0
    inbox_archived: int = 0
    inbox_pending: int = 0
    pending_promoted: int = 0
    explorer_moved_or_renamed: int = 0
    explorer_direct_added: int = 0
    explorer_deleted: int = 0
    locked_excels: List[Path] = None  # type: ignore

    def __post_init__(self):
        if self.locked_excels is None:
            self.locked_excels = []

    @property
    def has_changes(self) -> bool:
        return any([
            self.manual_edits_synced > 0,
            self.inbox_archived > 0,
            self.inbox_pending > 0,
            self.pending_promoted > 0,
            self.explorer_moved_or_renamed > 0,
            self.explorer_direct_added > 0,
            self.explorer_deleted > 0,
        ])


def _allocate_unique_target_path(target_dir: Path, standard_filename: str) -> Path:
    """
    若目标目录已存在同名文件，绝不覆盖且不丢弃，自动按 -02, -03 ... 顺延命名
    """
    target_dir.mkdir(parents=True, exist_ok=True)
    candidate = target_dir / standard_filename
    if not candidate.exists():
        return candidate

    p = Path(standard_filename)
    stem = p.stem
    ext = p.suffix
    idx = 2
    while True:
        cand = target_dir / f"{stem}-{idx:02d}{ext}"
        if not cand.exists():
            return cand
        idx += 1


def _to_rel_posix(path: Path, root_dir: Path) -> str:
    return path.resolve().relative_to(root_dir.resolve()).as_posix()


def _is_generated_ledger_excel(path: Path) -> bool:
    if path.suffix.lower() != ".xlsx":
        return False
    if path.name.endswith("_电子台账.xlsx") or path.name == "工程全流程总台账.xlsx":
        return True
    return False


def _collect_inbox_files(inbox_dir: Path) -> List[Path]:
    files: List[Path] = []
    if not inbox_dir.exists():
        return files
    for root, dirs, filenames in os.walk(inbox_dir):
        # 过滤隐藏与临时子目录
        dirs[:] = [d for d in dirs if not d.startswith((".", "_", "~"))]
        for fn in sorted(filenames):
            fp = Path(root) / fn
            if not should_ignore_file(fp):
                files.append(fp)
    return files


def _cleanup_empty_inbox_subdirs(inbox_dir: Path) -> None:
    if not inbox_dir.exists():
        return
    for root, dirs, _ in os.walk(inbox_dir, topdown=False):
        rp = Path(root)
        if rp == inbox_dir:
            continue
        try:
            if not any(rp.iterdir()):
                rp.rmdir()
        except OSError:
            pass


def _collect_physical_archive_files(cfg: AppConfig) -> Dict[str, Path]:
    """扫描 项目归档库/ 与 99_待人工确认/ 下所有真实工程文件，返回 {rel_posix_path: Path}"""
    found: Dict[str, Path] = {}
    for base_dir in (cfg.archive_lib_dir, cfg.pending_dir):
        if not base_dir.exists():
            continue
        for root, dirs, filenames in os.walk(base_dir):
            dirs[:] = [d for d in dirs if not d.startswith((".", "_", "~"))]
            for fn in filenames:
                fp = Path(root) / fn
                if should_ignore_file(fp) or _is_generated_ledger_excel(fp):
                    continue
                rel = _to_rel_posix(fp, cfg.root_dir)
                found[rel] = fp
    return found


def _parse_location_context(rel_posix: str) -> Tuple[int, str, str]:
    """
    从相对路径解析物理归属: (is_pending, project_name, stage)
    例如:
    - '项目归档库/清河水库除险加固工程/03_勘察设计/xxx.pdf' -> (0, '清河水库除险加固工程', '03_勘察设计')
    - '99_待人工确认/xxx.pdf' -> (1, '未匹配项目', '99_待人工确认')
    """
    parts = rel_posix.split("/")
    if len(parts) >= 4 and parts[0] == ARCHIVE_LIBRARY_DIR_NAME:
        proj_name = parts[1]
        stage_folder = parts[2]
        if stage_folder in STAGES_ORDER:
            return 0, proj_name, stage_folder
        return 0, proj_name, STAGES_ORDER[0]
    return 1, "未匹配项目", PENDING_DIR_NAME


def reconcile_once(
    cfg: AppConfig,
    db: LedgerDB,
    force_rebuild: bool = False,
    check_file_ready_delay: float = 0.6,
    verbose: bool = True,
) -> ReconcileReport:
    ensure_directory_skeleton(cfg)
    flush_deferred_workbooks(cfg)

    report = ReconcileReport()
    scratch_dir = cfg.data_dir / ".ocr_scratch"

    # 1. 先无锁回读用户在 Excel 中手工填写的“文号/金额/备注”
    report.manual_edits_synced = read_back_manual_edits_from_excels(cfg, db)
    if verbose and report.manual_edits_synced > 0:
        print(f"[Excel回读] 已保留同步 {report.manual_edits_synced} 条手工修订记录")

    # 2. 处理 00_待归档收件箱 中的新文件
    inbox_files = _collect_inbox_files(cfg.inbox_dir)
    for src_path in inbox_files:
        if not is_file_ready(src_path, check_delay=check_file_ready_delay):
            continue

        orig_name = src_path.name
        if verbose:
            print(f"[收件识别] 正在解析: {orig_name} ...")

        content = extract_document(src_path, scratch_dir)
        cls_res = classify_document(orig_name, content, cfg)

        if not cls_res.is_pending and cls_res.stage in STAGES_ORDER:
            target_dir = cfg.archive_lib_dir / cls_res.project_name / cls_res.stage
        else:
            target_dir = cfg.pending_dir

        dest_path = _allocate_unique_target_path(target_dir, cls_res.standard_filename)
        shutil.move(str(src_path), str(dest_path))

        st = dest_path.stat()
        sha = compute_sha256(dest_path)
        rel_posix = _to_rel_posix(dest_path, cfg.root_dir)
        now_ts = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        rec = LedgerRecord(
            id=None,
            rel_path=rel_posix,
            project_name=cls_res.project_name,
            stage=cls_res.stage if not cls_res.is_pending else PENDING_DIR_NAME,
            is_pending=1 if cls_res.is_pending else 0,
            pending_reason=cls_res.pending_reason,
            archived_filename=dest_path.name,
            original_filename=orig_name,
            file_ext=dest_path.suffix.lower(),
            doc_date=cls_res.doc_date,
            date_source=cls_res.date_source,
            doc_type=cls_res.doc_type,
            milestone_key=cls_res.milestone_key if not cls_res.is_pending else "",
            doc_no=cls_res.doc_no,
            doc_no_manual=0,
            amount_wan=cls_res.amount_wan,
            amount_manual=0,
            remark="",
            extract_method=content.extract_method,
            sha256=sha,
            file_size=st.st_size,
            mtime_ns=st.st_mtime_ns,
            archived_at=now_ts,
        )
        db.insert(rec)

        if cls_res.is_pending:
            report.inbox_pending += 1
            if verbose:
                print(f"  -> [待人工确认] 已移入 99_待人工确认\\{dest_path.name} ({cls_res.pending_reason})")
            send_windows_notification(
                "电子台账：待人工确认文件",
                f"{dest_path.name}\n原因：{cls_res.pending_reason}",
                is_warning=True,
                enabled=cfg.enable_windows_notification,
            )
        else:
            report.inbox_archived += 1
            if verbose:
                print(f"  -> [归档完成] {cls_res.project_name} / {cls_res.stage} / {dest_path.name}")
            send_windows_notification(
                f"自动归档成功：{cls_res.project_name}",
                f"阶段：{cls_res.stage}\n文件：{dest_path.name}",
                is_warning=False,
                enabled=cfg.enable_windows_notification,
            )

    _cleanup_empty_inbox_subdirs(cfg.inbox_dir)

    # 3A. 检查 99_待人工确认 中是否有因用户新配置了项目白名单而可以自动转正的文件
    if cfg.projects and cfg.pending_dir.exists():
        for pend_file in list(cfg.pending_dir.iterdir()):
            if not pend_file.is_file() or should_ignore_file(pend_file):
                continue
            old_rel = _to_rel_posix(pend_file, cfg.root_dir)
            old_rec = db.get_by_rel_path(old_rel)
            probe_orig_name = old_rec.original_filename if old_rec else pend_file.name
            content = extract_document(pend_file, scratch_dir)
            re_cls = classify_document(probe_orig_name, content, cfg)
            if not re_cls.is_pending and re_cls.stage in STAGES_ORDER:
                target_dir = cfg.archive_lib_dir / re_cls.project_name / re_cls.stage
                dest_path = _allocate_unique_target_path(target_dir, re_cls.standard_filename)
                shutil.move(str(pend_file), str(dest_path))
                new_rel = _to_rel_posix(dest_path, cfg.root_dir)
                st = dest_path.stat()
                if old_rec and old_rec.id is not None:
                    db.update_moved_or_renamed(
                        rec_id=old_rec.id,
                        rel_path=new_rel,
                        project_name=re_cls.project_name,
                        stage=re_cls.stage,
                        is_pending=0,
                        pending_reason="",
                        archived_filename=dest_path.name,
                        doc_date=re_cls.doc_date,
                        doc_type=re_cls.doc_type,
                        milestone_key=re_cls.milestone_key,
                        mtime_ns=st.st_mtime_ns,
                    )
                report.pending_promoted += 1
                if verbose:
                    print(f"[白名单转正] {pend_file.name} -> {re_cls.project_name}/{re_cls.stage}/{dest_path.name}")

    # 3B. 扫描物理目录现状，反向追踪用户在资源管理器中的挪动目录、改名、直接放入、删除行为
    disk_map = _collect_physical_archive_files(cfg)
    db_records = db.list_all()
    db_by_rel: Dict[str, LedgerRecord] = {r.rel_path: r for r in db_records}

    missing_records: List[LedgerRecord] = [r for rel, r in db_by_rel.items() if rel not in disk_map]
    untracked_rels: List[str] = [rel for rel in disk_map.keys() if rel not in db_by_rel]

    # 按 (sha256, file_size) 碰撞识别“移动了文件夹或修改了文件名”的同一份文件
    if missing_records and untracked_rels:
        missing_by_sig: Dict[Tuple[str, int], List[LedgerRecord]] = {}
        for mr in missing_records:
            missing_by_sig.setdefault((mr.sha256, mr.file_size), []).append(mr)

        still_untracked: List[str] = []
        matched_missing_ids: set = set()

        for u_rel in untracked_rels:
            u_path = disk_map[u_rel]
            st = u_path.stat()
            u_sha = compute_sha256(u_path)
            sig = (u_sha, st.st_size)
            candidates = missing_by_sig.get(sig, [])
            if candidates:
                matched_rec = candidates.pop(0)
                matched_missing_ids.add(matched_rec.id)

                new_is_pending, new_proj, new_stage = _parse_location_context(u_rel)
                p_date, p_proj, p_type = parse_standard_filename(u_path.name)
                final_date = p_date or matched_rec.doc_date
                final_type = (
                    sanitize_doc_type(p_type, cfg.max_doc_type_length)
                    if p_type
                    else sanitize_doc_type(u_path.stem, cfg.max_doc_type_length)
                )
                if new_is_pending and p_proj:
                    new_proj = p_proj
                final_ms = "" if new_is_pending else infer_milestone_from_type(new_stage, final_type)

                db.update_moved_or_renamed(
                    rec_id=matched_rec.id,  # type: ignore
                    rel_path=u_rel,
                    project_name=new_proj,
                    stage=new_stage,
                    is_pending=new_is_pending,
                    pending_reason="" if not new_is_pending else matched_rec.pending_reason,
                    archived_filename=u_path.name,
                    doc_date=final_date,
                    doc_type=final_type,
                    milestone_key=final_ms,
                    mtime_ns=st.st_mtime_ns,
                )
                report.explorer_moved_or_renamed += 1
                if verbose:
                    print(f"[目录同步] 检测到文件位置/名称变更: {matched_rec.rel_path} -> {u_rel}")
            else:
                still_untracked.append(u_rel)

        missing_records = [mr for mr in missing_records if mr.id not in matched_missing_ids]
        untracked_rels = still_untracked

    # 处理用户跳过收件箱、直接手工粘贴进某个阶段目录的新文件（尊重其当前文件名与目录，直接登记入账）
    for u_rel in untracked_rels:
        u_path = disk_map[u_rel]
        if not is_file_ready(u_path, check_delay=0.2):
            continue
        st = u_path.stat()
        sha = compute_sha256(u_path)
        content = extract_document(u_path, scratch_dir)
        cls_res = classify_document(u_path.name, content, cfg)
        new_is_pending, loc_proj, loc_stage = _parse_location_context(u_rel)

        p_date, p_proj, p_type = parse_standard_filename(u_path.name)
        final_proj = loc_proj if not new_is_pending else (p_proj or cls_res.project_name)
        final_stage = loc_stage if not new_is_pending else PENDING_DIR_NAME
        final_date = p_date or cls_res.doc_date
        final_type = sanitize_doc_type(p_type, cfg.max_doc_type_length) if p_type else cls_res.doc_type
        final_ms = "" if new_is_pending else infer_milestone_from_type(final_stage, final_type)

        rec = LedgerRecord(
            id=None,
            rel_path=u_rel,
            project_name=final_proj,
            stage=final_stage,
            is_pending=new_is_pending,
            pending_reason="直接放入待确认目录" if new_is_pending else "",
            archived_filename=u_path.name,
            original_filename=u_path.name,
            file_ext=u_path.suffix.lower(),
            doc_date=final_date,
            date_source=cls_res.date_source,
            doc_type=final_type,
            milestone_key=final_ms,
            doc_no=cls_res.doc_no,
            doc_no_manual=0,
            amount_wan=cls_res.amount_wan,
            amount_manual=0,
            remark="",
            extract_method=content.extract_method,
            sha256=sha,
            file_size=st.st_size,
            mtime_ns=st.st_mtime_ns,
            archived_at=dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        )
        db.insert(rec)
        report.explorer_direct_added += 1
        if verbose:
            print(f"[直接纳管] 发现阶段目录新增文件并纳入台账: {u_rel}")

    # 清理已在物理磁盘被用户删除的文件记录，杜绝死链
    if missing_records:
        del_ids = [mr.id for mr in missing_records if mr.id is not None]
        db.delete_ids(del_ids)
        report.explorer_deleted = len(del_ids)
        if verbose:
            print(f"[死链清理] 已同步移除 {len(del_ids)} 条物理已删除文件的台账记录")

    # 4. 判断是否需要重绘 Excel 与 HTML 看板
    outputs_missing = (not cfg.master_excel_path.exists()) or (not cfg.dashboard_html_path.exists())
    if force_rebuild or report.has_changes or outputs_missing:
        all_recs = db.list_all()
        summary = compute_portfolio_summary(cfg, all_recs)
        build_dashboard_html(cfg, summary)
        locked = build_all_excel_ledgers(cfg, summary)
        report.locked_excels = locked
        if locked:
            locked_names = "、".join(p.name for p in locked)
            if verbose:
                print(f"[占用提示] {locked_names} 正被 Excel/WPS 打开占用，已暂存更新，关闭 Excel 后自动生效。")
            send_windows_notification(
                "台账已更新（等待关闭Excel生效）",
                f"{locked_names} 当前正被打开占用，关闭 Excel 窗口后将自动替换为最新台账。",
                is_warning=True,
                enabled=cfg.enable_windows_notification,
            )
        elif verbose:
            print("[台账刷新] 《工程全流程总台账.xlsx》、各项目独立台账及《项目状态可视化看板.html》已同步更新！")

    return report

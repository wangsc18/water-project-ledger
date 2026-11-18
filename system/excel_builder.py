# -*- coding: utf-8 -*-
"""
双层 Excel 超链接电子台账生成与手工修订无锁回读模块
- 顶层《工程全流程总台账.xlsx》：多项目总览 + 全量总台账 + 各项目专属分表 + 99待人工确认表
- 单项目《<项目简称>_电子台账.xlsx》：独立保存在各项目文件夹内，采用同级相对超链接，单项目打包外发即点即开
- 专属适配 Windows Excel/WPS 独占文件锁：无锁快照回读手工列 + 占用时自动暂存待生效文件 + 关闭 Excel 后自动替换
"""
from __future__ import annotations

import datetime as dt
import os
import re
import shutil
from pathlib import Path
from typing import Dict, List, Optional, Set, Tuple

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from config_loader import ARCHIVE_LIBRARY_DIR_NAME, PENDING_DIR_NAME, STAGES_ORDER, AppConfig
from database import LedgerDB, LedgerRecord
from progress_calc import PortfolioSummary, ProjectProgress, compute_portfolio_summary

FONT_FAMILY = "Microsoft YaHei"

# 样式调色板
FILL_NAVY_TITLE = PatternFill("solid", fgColor="154360")
FILL_HEADER = PatternFill("solid", fgColor="1F618D")
FILL_HEADER_EDIT = PatternFill("solid", fgColor="B7950B")
FILL_SUBHEADER = PatternFill("solid", fgColor="D4E6F1")
FILL_GREEN_DONE = PatternFill("solid", fgColor="D4EFDF")
FILL_BLUE_PROG = PatternFill("solid", fgColor="D6EAF8")
FILL_GRAY_NONE = PatternFill("solid", fgColor="F2F4F4")
FILL_AMBER_WARN = PatternFill("solid", fgColor="FCF3CF")
FILL_ZEBRA = PatternFill("solid", fgColor="F8FBFD")
FILL_EDIT_CELL = PatternFill("solid", fgColor="FFFDF5")

FONT_TITLE = Font(name=FONT_FAMILY, size=14, bold=True, color="FFFFFF")
FONT_SUBTITLE = Font(name=FONT_FAMILY, size=10, italic=False, color="154360")
FONT_TH = Font(name=FONT_FAMILY, size=10, bold=True, color="FFFFFF")
FONT_SECTION = Font(name=FONT_FAMILY, size=11, bold=True, color="154360")
FONT_NORMAL = Font(name=FONT_FAMILY, size=10, color="1C2833")
FONT_BOLD = Font(name=FONT_FAMILY, size=10, bold=True, color="1C2833")
FONT_LINK = Font(name=FONT_FAMILY, size=10, color="0563C1", underline="single")
FONT_LINK_BOLD = Font(name=FONT_FAMILY, size=10, bold=True, color="0563C1", underline="single")
FONT_GREEN = Font(name=FONT_FAMILY, size=10, bold=True, color="196F3D")
FONT_BLUE = Font(name=FONT_FAMILY, size=10, bold=True, color="1B4F72")
FONT_GRAY = Font(name=FONT_FAMILY, size=10, color="7F8C8D")
FONT_WARN = Font(name=FONT_FAMILY, size=10, bold=True, color="9A7D0A")

THIN_SIDE = Side(border_style="thin", color="D5DBDB")
BORDER_ALL = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)

ALIGN_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
ALIGN_LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
ALIGN_RIGHT = Alignment(horizontal="right", vertical="center", wrap_text=True)


def _format_ymd_dash(ymd8: str) -> str:
    if len(ymd8) == 8 and ymd8.isdigit():
        return f"{ymd8[:4]}-{ymd8[4:6]}-{ymd8[6:8]}"
    return ymd8


def _make_safe_sheet_title(raw_title: str, used_titles: Set[str]) -> str:
    """Excel 工作表名称最长 31 字符且不能含 : \\ / ? * [ ]"""
    cleaned = re.sub(r"[:\\/?*\[\]]", "_", raw_title).strip("'")
    base = cleaned[:31] if len(cleaned) > 31 else cleaned
    cand = base
    idx = 2
    while cand in used_titles:
        suffix = f"_{idx}"
        cand = base[: 31 - len(suffix)] + suffix
        idx += 1
    used_titles.add(cand)
    return cand


def _normalize_cell_str(val) -> str:
    if val is None:
        return ""
    if isinstance(val, float):
        if val.is_integer():
            return f"{val:.2f}"
        return f"{val:.2f}"
    s = str(val).strip()
    return "" if s == "-" else s


# ==============================================================================
# 1. 回读用户在 Excel 手工填写的“文号/编号、涉及金额、人工备注”列（支持被 Excel 打开锁死时快照读取）
# ==============================================================================
def read_back_manual_edits_from_excels(cfg: AppConfig, db: LedgerDB) -> int:
    orig_records = db.list_all()
    if not orig_records:
        return 0
    orig_by_id: Dict[int, LedgerRecord] = {r.id: r for r in orig_records if r.id is not None}

    excel_files: List[Path] = []
    if cfg.master_excel_path.exists():
        excel_files.append(cfg.master_excel_path)
    if cfg.archive_lib_dir.exists():
        for p_dir in cfg.archive_lib_dir.iterdir():
            if p_dir.is_dir():
                sub_xl = p_dir / f"{p_dir.name}_电子台账.xlsx"
                if sub_xl.exists():
                    excel_files.append(sub_xl)

    if not excel_files:
        return 0

    # 收集相对于数据库快照发生变化的字段: rec_id -> {"doc_no": ..., "amount_wan": ..., "remark": ...}
    overrides: Dict[int, Dict[str, str]] = {}
    tmp_copy = cfg.data_dir / f".readback_tmp_{os.getpid()}.xlsx"

    for xl_path in excel_files:
        try:
            shutil.copy2(xl_path, tmp_copy)
            wb = openpyxl.load_workbook(str(tmp_copy), read_only=True, data_only=True)
            try:
                for ws in wb.worksheets:
                    id_col = doc_no_col = amount_col = remark_col = None
                    for row in ws.iter_rows(values_only=True):
                        if not any(row):
                            continue
                        row_strs = [str(c).strip() if c is not None else "" for c in row]
                        # 识别表头行
                        if "台账ID" in row_strs and any("人工备注" in s for s in row_strs):
                            id_col = row_strs.index("台账ID")
                            for idx_c, h_text in enumerate(row_strs):
                                if "文号/合同编号" in h_text:
                                    doc_no_col = idx_c
                                elif "涉及金额" in h_text:
                                    amount_col = idx_c
                                elif "人工备注" in h_text:
                                    remark_col = idx_c
                            continue

                        if id_col is not None and id_col < len(row):
                            raw_id = str(row[id_col] or "").strip()
                            m_id = re.fullmatch(r"REC-(\d+)", raw_id)
                            if not m_id:
                                continue
                            rec_id = int(m_id.group(1))
                            orig_rec = orig_by_id.get(rec_id)
                            if not orig_rec:
                                continue

                            if doc_no_col is not None and doc_no_col < len(row):
                                v_doc = _normalize_cell_str(row[doc_no_col])
                                if v_doc != (orig_rec.doc_no or ""):
                                    overrides.setdefault(rec_id, {})["doc_no"] = v_doc

                            if amount_col is not None and amount_col < len(row):
                                v_amt = _normalize_cell_str(row[amount_col])
                                if v_amt != (orig_rec.amount_wan or ""):
                                    overrides.setdefault(rec_id, {})["amount_wan"] = v_amt

                            if remark_col is not None and remark_col < len(row):
                                v_rem = _normalize_cell_str(row[remark_col])
                                if v_rem != (orig_rec.remark or ""):
                                    overrides.setdefault(rec_id, {})["remark"] = v_rem
            finally:
                wb.close()
        except Exception:
            continue
        finally:
            try:
                if tmp_copy.exists():
                    tmp_copy.unlink()
            except OSError:
                pass

    for rec_id, fields in overrides.items():
        db.update_manual_fields(
            rec_id=rec_id,
            doc_no=fields.get("doc_no"),
            amount_wan=fields.get("amount_wan"),
            remark=fields.get("remark"),
        )
    return len(overrides)


# ==============================================================================
# 2. 适配 Windows Excel 独占写锁的安全保存与解锁后自动替换
# ==============================================================================
def _is_locked_by_office(target_path: Path) -> bool:
    lock_file = target_path.parent / f"~${target_path.name}"
    if lock_file.exists():
        return True
    if target_path.exists():
        try:
            with open(target_path, "r+b"):
                pass
        except PermissionError:
            return True
        except OSError:
            pass
    return False


def safe_save_workbook(wb: openpyxl.Workbook, target_path: Path) -> bool:
    """
    安全保存工作簿：
    返回 True 表示直接写入成功；
    返回 False 表示目标文件正被 Windows Excel/WPS 打开独占锁定，已自动写入同目录 `~待生效_xxx.xlsx`
    """
    target_path.parent.mkdir(parents=True, exist_ok=True)
    deferred_path = target_path.parent / f"~待生效_{target_path.name}"
    tmp_write = target_path.parent / f".writing_{os.getpid()}_{target_path.name}"

    wb.save(str(tmp_write))

    if _is_locked_by_office(target_path):
        try:
            os.replace(str(tmp_write), str(deferred_path))
        except OSError:
            if tmp_write.exists():
                tmp_write.unlink()
        return False

    try:
        os.replace(str(tmp_write), str(target_path))
        if deferred_path.exists():
            try:
                deferred_path.unlink()
            except OSError:
                pass
        return True
    except PermissionError:
        os.replace(str(tmp_write), str(deferred_path))
        return False


def flush_deferred_workbooks(cfg: AppConfig) -> int:
    """监听循环中定期调用：一旦用户关闭了 Excel 释放锁，自动将 `~待生效_*.xlsx` 无感替换到位"""
    swapped = 0
    dirs_to_check = [cfg.root_dir]
    if cfg.archive_lib_dir.exists():
        dirs_to_check.extend([d for d in cfg.archive_lib_dir.iterdir() if d.is_dir()])

    for d in dirs_to_check:
        for f in d.glob("~待生效_*.xlsx"):
            real_name = f.name[len("~待生效_"):]
            real_target = d / real_name
            if not _is_locked_by_office(real_target):
                try:
                    os.replace(str(f), str(real_target))
                    swapped += 1
                except OSError:
                    pass
    return swapped


# ==============================================================================
# 3. 台账表格构建辅助函数
# ==============================================================================
def _set_column_widths(ws, widths: List[float]) -> None:
    for idx, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = w


def _write_ledger_table(
    ws,
    start_row: int,
    records: List[LedgerRecord],
    rel_prefix_to_strip: str = "",
    include_project_col: bool = True,
) -> int:
    """
    写入标准化的收文台账明细表（支持总台账与单项目独立台账复用，确保超链接相对路径精准对应）
    返回结束行号
    """
    headers = [
        "序号",
        "台账ID",
        "所属项目" if include_project_col else "项目阶段编码",
        "流程阶段",
        "规范归档文件名（点击直接打开）",
        "文件具体类型",
        "点亮核心要件",
        "成文/签署日期",
        "文号/合同编号（可手填）",
        "涉及金额-万元（可手填）",
        "人工备注（可手填）",
        "快捷打开原件",
        "打开所在文件夹",
        "归档前原始文件名",
        "内容识别方式",
        "归档入库时间",
    ]
    editable_indices = {9, 10, 11}  # 1-based 列号：文号、金额、备注

    ws.row_dimensions[start_row].height = 28
    for c_idx, h_text in enumerate(headers, start=1):
        cell = ws.cell(row=start_row, column=c_idx, value=h_text)
        cell.font = FONT_TH
        cell.fill = FILL_HEADER_EDIT if c_idx in editable_indices else FILL_HEADER
        cell.alignment = ALIGN_CENTER
        cell.border = BORDER_ALL

    if not records:
        empty_row = start_row + 1
        ws.row_dimensions[empty_row].height = 26
        ws.merge_cells(start_row=empty_row, start_column=1, end_row=empty_row, end_column=len(headers))
        c = ws.cell(row=empty_row, column=1, value="暂无归档文件记录（将工程文件拖入 00_待归档收件箱 后自动生成）")
        c.font = FONT_GRAY
        c.alignment = ALIGN_CENTER
        for col_i in range(1, len(headers) + 1):
            ws.cell(row=empty_row, column=col_i).border = BORDER_ALL
        return empty_row

    curr_row = start_row
    for seq, r in enumerate(records, start=1):
        curr_row += 1
        ws.row_dimensions[curr_row].height = 24
        is_even = (seq % 2 == 0)
        row_fill = FILL_ZEBRA if is_even else None

        # 计算相对于当前 Excel 所在目录的相对路径（使用正斜杠兼容 Windows Excel 与 WPS）
        rel_target = r.rel_path.replace("\\", "/")
        if rel_prefix_to_strip and rel_target.startswith(rel_prefix_to_strip):
            rel_target = rel_target[len(rel_prefix_to_strip):]
        rel_folder = str(Path(rel_target).parent).replace("\\", "/")

        amt_val: object = r.amount_wan or ""
        if r.amount_wan:
            try:
                amt_val = round(float(str(r.amount_wan).replace(",", "")), 2)
            except ValueError:
                amt_val = r.amount_wan

        vals = [
            seq,
            f"REC-{r.id:04d}" if r.id else "",
            r.project_name if include_project_col else r.stage[:2],
            r.stage,
            r.archived_filename,
            r.doc_type,
            f"★ {r.milestone_key}" if r.milestone_key else "-",
            _format_ymd_dash(r.doc_date),
            r.doc_no or "",
            amt_val,
            r.remark or (r.pending_reason if r.is_pending else ""),
            "📄 打开原件",
            "📂 打开目录",
            r.original_filename,
            r.extract_method or "",
            r.archived_at,
        ]

        for c_idx, val in enumerate(vals, start=1):
            cell = ws.cell(row=curr_row, column=c_idx, value=val)
            cell.border = BORDER_ALL
            cell.font = FONT_NORMAL

            if c_idx in editable_indices:
                cell.fill = FILL_EDIT_CELL
            elif row_fill:
                cell.fill = row_fill

            if c_idx in (1, 2, 4, 8, 12, 13, 15, 16):
                cell.alignment = ALIGN_CENTER
            elif c_idx == 10:
                cell.alignment = ALIGN_RIGHT
                if isinstance(val, float):
                    cell.number_format = "#,##0.00"
            else:
                cell.alignment = ALIGN_LEFT

            if c_idx == 7 and r.milestone_key:
                cell.font = FONT_GREEN
                cell.fill = FILL_GREEN_DONE

            # 挂载文件与文件夹相对路径超链接
            if c_idx == 5:
                cell.hyperlink = rel_target
                cell.font = FONT_LINK_BOLD
            elif c_idx == 12:
                cell.hyperlink = rel_target
                cell.font = FONT_LINK
            elif c_idx == 13:
                cell.hyperlink = rel_folder
                cell.font = FONT_LINK

    return curr_row


def _write_project_milestone_matrix(
    ws,
    start_row: int,
    proj_prog: ProjectProgress,
    rel_prefix_to_strip: str = "",
) -> int:
    """在项目专属页上半部分写入《八大阶段核心必办要件亮灯核查表》"""
    headers = [
        "流程阶段",
        "阶段齐套状态",
        "本阶段文件数",
        "核心必办亮灯要件名称",
        "要件收齐状态",
        "对应已归档文件（点击直接打开批文/合同原件）",
        "成文/签署日期",
        "文号/合同编号",
    ]
    ws.row_dimensions[start_row].height = 26
    for c_idx, h in enumerate(headers, start=1):
        c = ws.cell(row=start_row, column=c_idx, value=h)
        c.font = FONT_TH
        c.fill = FILL_HEADER
        c.alignment = ALIGN_CENTER
        c.border = BORDER_ALL

    curr_row = start_row
    for stg_code in STAGES_ORDER:
        sp = proj_prog.stages[stg_code]
        ms_list = sp.milestones
        if not ms_list:
            curr_row += 1
            ws.row_dimensions[curr_row].height = 22
            ws.cell(row=curr_row, column=1, value=stg_code).alignment = ALIGN_CENTER
            ws.cell(row=curr_row, column=2, value=sp.status_label).alignment = ALIGN_CENTER
            ws.cell(row=curr_row, column=3, value=sp.file_count).alignment = ALIGN_CENTER
            ws.cell(row=curr_row, column=4, value="（已配置免检无必办项）").font = FONT_GRAY
            for ci in range(1, 9):
                ws.cell(row=curr_row, column=ci).border = BORDER_ALL
            continue

        first_ms_row = curr_row + 1
        for ms in ms_list:
            curr_row += 1
            ws.row_dimensions[curr_row].height = 23

            c_stg = ws.cell(row=curr_row, column=1, value=stg_code)
            c_status = ws.cell(row=curr_row, column=2, value=f"{sp.status_label} ({sp.lit_milestones}/{sp.total_milestones})")
            c_cnt = ws.cell(row=curr_row, column=3, value=sp.file_count)
            c_ms_name = ws.cell(row=curr_row, column=4, value=ms.name)

            if ms.is_lit and ms.matched_files:
                latest_f = ms.matched_files[-1]
                rel_target = latest_f.rel_path.replace("\\", "/")
                if rel_prefix_to_strip and rel_target.startswith(rel_prefix_to_strip):
                    rel_target = rel_target[len(rel_prefix_to_strip):]

                c_lit = ws.cell(row=curr_row, column=5, value="● 已收齐")
                c_lit.font = FONT_GREEN
                c_lit.fill = FILL_GREEN_DONE

                c_file = ws.cell(row=curr_row, column=6, value=latest_f.archived_filename)
                c_file.hyperlink = rel_target
                c_file.font = FONT_LINK

                c_dt = ws.cell(row=curr_row, column=7, value=_format_ymd_dash(latest_f.doc_date))
                c_no = ws.cell(row=curr_row, column=8, value=latest_f.doc_no or "-")
            else:
                c_lit = ws.cell(row=curr_row, column=5, value="○ 待收件")
                c_lit.font = FONT_GRAY
                c_lit.fill = FILL_GRAY_NONE

                c_file = ws.cell(row=curr_row, column=6, value="暂未收到对应文件")
                c_file.font = FONT_GRAY
                c_dt = ws.cell(row=curr_row, column=7, value="-")
                c_no = ws.cell(row=curr_row, column=8, value="-")

            for ci in range(1, 9):
                cell_i = ws.cell(row=curr_row, column=ci)
                cell_i.border = BORDER_ALL
                if ci in (1, 2, 3, 5, 7):
                    cell_i.alignment = ALIGN_CENTER
                else:
                    cell_i.alignment = ALIGN_LEFT
                if ci not in (2, 5, 6) and cell_i.font == openpyxl.styles.Font():
                    cell_i.font = FONT_NORMAL

            c_ms_name.font = FONT_BOLD if ms.is_lit else FONT_NORMAL

        # 合并同阶段前3列单元格并按阶段完成状态着色
        if curr_row > first_ms_row:
            ws.merge_cells(start_row=first_ms_row, start_column=1, end_row=curr_row, end_column=1)
            ws.merge_cells(start_row=first_ms_row, start_column=2, end_row=curr_row, end_column=2)
            ws.merge_cells(start_row=first_ms_row, start_column=3, end_row=curr_row, end_column=3)

        top_stg_cell = ws.cell(row=first_ms_row, column=1)
        top_st_cell = ws.cell(row=first_ms_row, column=2)
        top_cnt_cell = ws.cell(row=first_ms_row, column=3)
        top_stg_cell.font = FONT_BOLD
        top_cnt_cell.font = FONT_BOLD
        if sp.status == "completed":
            top_st_cell.fill = FILL_GREEN_DONE
            top_st_cell.font = FONT_GREEN
        elif sp.status == "in_progress":
            top_st_cell.fill = FILL_BLUE_PROG
            top_st_cell.font = FONT_BLUE
        else:
            top_st_cell.fill = FILL_GRAY_NONE
            top_st_cell.font = FONT_GRAY

    return curr_row


# ==============================================================================
# 4. 生成根目录《工程全流程总台账.xlsx》与各单项目独立《_电子台账.xlsx》
# ==============================================================================
def build_all_excel_ledgers(cfg: AppConfig, summary: PortfolioSummary) -> List[Path]:
    """
    生成总台账与各在管项目独立台账
    返回因被 Excel/WPS 打开锁定而暂存待替换的文件列表（空列表代表全部直接落盘成功）
    """
    locked_files: List[Path] = []
    now_str = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # ---------------- A. 构建根目录《工程全流程总台账.xlsx》 ----------------
    wb_master = openpyxl.Workbook()
    used_titles: Set[str] = set()

    # 预先分配各项目分 Sheet 名称以便总览表生成内部跳转超链接
    proj_sheet_names: Dict[str, str] = {}
    for pp in summary.projects:
        proj_sheet_names[pp.spec.standard_name] = _make_safe_sheet_title(
            f"项目_{pp.spec.standard_name}", used_titles
        )

    # Sheet 1: 00_多项目进度总览
    ws_ov = wb_master.active
    ws_ov.title = "00_多项目进度总览"
    used_titles.add(ws_ov.title)
    ws_ov.views.sheetView[0].showGridLines = True

    total_ov_cols = 6 + len(STAGES_ORDER)
    ws_ov.merge_cells(start_row=1, start_column=1, end_row=1, end_column=total_ov_cols)
    t_cell = ws_ov.cell(row=1, column=1, value="水利新建工程全流程电子台账 — 多项目状态总览")
    t_cell.font = FONT_TITLE
    t_cell.fill = FILL_NAVY_TITLE
    t_cell.alignment = ALIGN_CENTER
    ws_ov.row_dimensions[1].height = 36

    ws_ov.merge_cells(start_row=2, start_column=1, end_row=2, end_column=total_ov_cols)
    sub_msg = (
        f"最后自动同步时间：{now_str}   |   "
        f"在管水利项目：{summary.total_projects} 个   |   "
        f"已规范归档文件：{summary.total_active_files} 份   |   "
        f"待人工确认文件：{summary.total_pending_files} 份（提示：点击表中蓝色下划线文字可直接跳转或打开对应文件夹/原件）"
    )
    s_cell = ws_ov.cell(row=2, column=1, value=sub_msg)
    s_cell.font = FONT_SUBTITLE
    s_cell.fill = FILL_AMBER_WARN if summary.total_pending_files > 0 else FILL_SUBHEADER
    s_cell.alignment = ALIGN_LEFT
    ws_ov.row_dimensions[2].height = 24

    ov_headers = [
        "序号",
        "项目标准名称（点击看分表）",
        "当前所处阶段",
        "要件齐套率",
        "已归档总数",
    ] + STAGES_ORDER + ["打开项目归档目录"]

    ws_ov.row_dimensions[4].height = 30
    for c_i, h_txt in enumerate(ov_headers, start=1):
        hc = ws_ov.cell(row=4, column=c_i, value=h_txt)
        hc.font = FONT_TH
        hc.fill = FILL_HEADER
        hc.alignment = ALIGN_CENTER
        hc.border = BORDER_ALL

    curr_r = 4
    for idx_p, pp in enumerate(summary.projects, start=1):
        curr_r += 1
        ws_ov.row_dimensions[curr_r].height = 32
        p_name = pp.spec.standard_name
        sh_target = proj_sheet_names[p_name]
        proj_folder_rel = f"{ARCHIVE_LIBRARY_DIR_NAME}/{p_name}"

        ws_ov.cell(row=curr_r, column=1, value=idx_p).alignment = ALIGN_CENTER

        c_pn = ws_ov.cell(row=curr_r, column=2, value=p_name)
        c_pn.hyperlink = f"#'{sh_target}'!A1"
        c_pn.font = FONT_LINK_BOLD
        c_pn.alignment = ALIGN_LEFT

        c_cur = ws_ov.cell(row=curr_r, column=3, value=pp.current_stage_label)
        c_cur.font = FONT_BOLD
        c_cur.alignment = ALIGN_CENTER

        c_pct = ws_ov.cell(
            row=curr_r,
            column=4,
            value=f"{pp.completion_pct:.1f}% ({pp.lit_milestones}/{pp.total_milestones})",
        )
        c_pct.font = FONT_BOLD
        c_pct.alignment = ALIGN_CENTER

        c_tot = ws_ov.cell(row=curr_r, column=5, value=f"{pp.total_files} 份")
        c_tot.alignment = ALIGN_CENTER
        c_tot.font = FONT_NORMAL

        for s_idx, stg_code in enumerate(STAGES_ORDER):
            col_pos = 6 + s_idx
            sp = pp.stages[stg_code]
            txt = f"{sp.status_label} ({sp.lit_milestones}/{sp.total_milestones})\n共 {sp.file_count} 份"
            sc = ws_ov.cell(row=curr_r, column=col_pos, value=txt)
            sc.alignment = ALIGN_CENTER
            sc.hyperlink = f"{proj_folder_rel}/{stg_code}"
            if sp.status == "completed":
                sc.fill = FILL_GREEN_DONE
                sc.font = FONT_GREEN
            elif sp.status == "in_progress":
                sc.fill = FILL_BLUE_PROG
                sc.font = FONT_BLUE
            else:
                sc.fill = FILL_GRAY_NONE
                sc.font = FONT_GRAY

        c_dir = ws_ov.cell(row=curr_r, column=total_ov_cols, value="📂 打开项目库")
        c_dir.hyperlink = proj_folder_rel
        c_dir.font = FONT_LINK
        c_dir.alignment = ALIGN_CENTER

        for c_i in range(1, total_ov_cols + 1):
            ws_ov.cell(row=curr_r, column=c_i).border = BORDER_ALL

    _set_column_widths(ws_ov, [6, 28, 22, 16, 12] + [16] * len(STAGES_ORDER) + [16])
    ws_ov.freeze_panes = "C5"

    # Sheet 2: 01_全部项目收文总台账
    ws_all = wb_master.create_sheet("01_全部项目收文总台账")
    ws_all.merge_cells("A1:P1")
    c_all_t = ws_all["A1"]
    c_all_t.value = "全部在管水利项目 — 全流程收文电子总台账（点击文件名直接打开原件）"
    c_all_t.font = FONT_TITLE
    c_all_t.fill = FILL_NAVY_TITLE
    c_all_t.alignment = ALIGN_CENTER
    ws_all.row_dimensions[1].height = 34

    ws_all.merge_cells("A2:P2")
    c_all_s = ws_all["A2"]
    c_all_s.value = (
        f"共 {len(summary.active_files)} 份正式归档文件  |  "
        "金黄色表头列（文号/合同编号、涉及金额、人工备注）支持直接在 Excel 中手工补充修改，下次自动归档重刷时系统会自动保留您的手填内容！"
    )
    c_all_s.font = FONT_SUBTITLE
    c_all_s.fill = FILL_SUBHEADER
    c_all_s.alignment = ALIGN_LEFT
    ws_all.row_dimensions[2].height = 22

    end_all_row = _write_ledger_table(ws_all, start_row=3, records=summary.active_files, include_project_col=True)
    ws_all.auto_filter.ref = f"A3:P{max(4, end_all_row)}"
    ws_all.freeze_panes = "A4"
    _set_column_widths(ws_all, [6, 11, 24, 14, 48, 26, 22, 14, 24, 16, 28, 14, 14, 30, 14, 20])

    # Sheet 3..N: 每个项目独立专属 Sheet
    for pp in summary.projects:
        p_name = pp.spec.standard_name
        ws_p = wb_master.create_sheet(proj_sheet_names[p_name])

        ws_p.merge_cells("A1:P1")
        cp_t = ws_p["A1"]
        cp_t.value = f"【{p_name}】全流程要件亮灯核查与收文台账"
        cp_t.font = FONT_TITLE
        cp_t.fill = FILL_NAVY_TITLE
        cp_t.alignment = ALIGN_CENTER
        ws_p.row_dimensions[1].height = 34

        inv_txt = f"{pp.approved_investment_wan:.2f} 万元" if pp.approved_investment_wan is not None else "待提取批复"
        ws_p.merge_cells("A2:P2")
        cp_s = ws_p["A2"]
        cp_s.value = (
            f"工程类型：{pp.spec.project_type}  |  当前阶段：{pp.current_stage_label}  |  "
            f"核心要件齐套率：{pp.completion_pct:.1f}% ({pp.lit_milestones}/{pp.total_milestones})  |  "
            f"批复总投资：{inv_txt}  |  已签合同累计：{pp.total_amount_contract_wan:.2f} 万元"
        )
        cp_s.font = FONT_SUBTITLE
        cp_s.fill = FILL_SUBHEADER
        cp_s.alignment = ALIGN_LEFT
        ws_p.row_dimensions[2].height = 23

        ws_p.cell(row=4, column=1, value="一、 八大阶段核心必办要件核查清单（收齐自动点亮绿灯并挂载原件超链接）").font = FONT_SECTION
        ws_p.row_dimensions[4].height = 24
        ms_end = _write_project_milestone_matrix(ws_p, start_row=5, proj_prog=pp, rel_prefix_to_strip="")

        tbl_title_row = ms_end + 2
        ws_p.cell(row=tbl_title_row, column=1, value="二、 本项目已归档全流程文件明细台账").font = FONT_SECTION
        ws_p.row_dimensions[tbl_title_row].height = 24
        _write_ledger_table(ws_p, start_row=tbl_title_row + 1, records=pp.all_files, include_project_col=False)
        _set_column_widths(ws_p, [14, 18, 12, 26, 14, 48, 22, 22, 24, 16, 28, 14, 14, 30, 14, 20])

    # Sheet 末: 99_待人工确认
    ws_pend = wb_master.create_sheet("99_待人工确认")
    ws_pend.sheet_properties.tabColor = "E67E22" if summary.pending_files else "95A5A6"
    ws_pend.merge_cells("A1:P1")
    cp_pend = ws_pend["A1"]
    cp_pend.value = f"99_待人工确认文件清单（当前共 {len(summary.pending_files)} 份）"
    cp_pend.font = FONT_TITLE
    cp_pend.fill = FILL_NAVY_TITLE
    cp_pend.alignment = ALIGN_CENTER
    ws_pend.row_dimensions[1].height = 34

    ws_pend.merge_cells("A2:P2")
    cp_pend_tip = ws_pend["A2"]
    cp_pend_tip.value = (
        "处理方式：① 若属于新工程，请在《项目配置.yaml》中添加该项目关键词后双击《立即扫描校准一次.bat》；"
        "② 或直接在 Windows 资源管理器中将文件从 99_待人工确认 拖入对应项目的 01~08 阶段文件夹并按需改名，台账将自动同步转正！"
    )
    cp_pend_tip.font = FONT_WARN if summary.pending_files else FONT_SUBTITLE
    cp_pend_tip.fill = FILL_AMBER_WARN if summary.pending_files else FILL_SUBHEADER
    cp_pend_tip.alignment = ALIGN_LEFT
    ws_pend.row_dimensions[2].height = 26

    c_open_pend = ws_pend.cell(row=3, column=1, value="📂 点击打开 99_待人工确认 文件夹")
    c_open_pend.hyperlink = PENDING_DIR_NAME
    c_open_pend.font = FONT_LINK_BOLD
    ws_pend.row_dimensions[3].height = 22

    _write_ledger_table(ws_pend, start_row=4, records=summary.pending_files, include_project_col=True)
    _set_column_widths(ws_pend, [6, 11, 24, 14, 48, 26, 22, 14, 24, 16, 36, 14, 14, 30, 14, 20])

    if not safe_save_workbook(wb_master, cfg.master_excel_path):
        locked_files.append(cfg.master_excel_path)

    # ---------------- B. 构建每个在管项目文件夹内的独立《<项目简称>_电子台账.xlsx》 ----------------
    for pp in summary.projects:
        p_name = pp.spec.standard_name
        proj_dir = cfg.archive_lib_dir / p_name
        proj_dir.mkdir(parents=True, exist_ok=True)
        sub_excel_path = proj_dir / f"{p_name}_电子台账.xlsx"
        strip_prefix = f"{ARCHIVE_LIBRARY_DIR_NAME}/{p_name}/"

        wb_sub = openpyxl.Workbook()
        ws_sub_ms = wb_sub.active
        ws_sub_ms.title = "00_阶段进度与要件核查"

        ws_sub_ms.merge_cells("A1:H1")
        cs1 = ws_sub_ms["A1"]
        cs1.value = f"【{p_name}】全流程建设进度与核心必办要件核查表"
        cs1.font = FONT_TITLE
        cs1.fill = FILL_NAVY_TITLE
        cs1.alignment = ALIGN_CENTER
        ws_sub_ms.row_dimensions[1].height = 34

        ws_sub_ms.merge_cells("A2:H2")
        cs2 = ws_sub_ms["A2"]
        cs2.value = (
            f"更新时间：{now_str}  |  当前阶段：{pp.current_stage_label}  |  "
            f"核心要件齐套率：{pp.completion_pct:.1f}% ({pp.lit_milestones}/{pp.total_milestones})  |  "
            f"已收归档文件：{pp.total_files} 份（本表采用同目录相对超链接，单独拷贝本项目文件夹外发依然可一键点开原件）"
        )
        cs2.font = FONT_SUBTITLE
        cs2.fill = FILL_SUBHEADER
        cs2.alignment = ALIGN_LEFT
        ws_sub_ms.row_dimensions[2].height = 24

        _write_project_milestone_matrix(ws_sub_ms, start_row=4, proj_prog=pp, rel_prefix_to_strip=strip_prefix)
        _set_column_widths(ws_sub_ms, [15, 18, 14, 28, 14, 50, 16, 26])
        ws_sub_ms.freeze_panes = "A5"

        ws_sub_list = wb_sub.create_sheet("01_项目全流程收文台账")
        ws_sub_list.merge_cells("A1:P1")
        cl1 = ws_sub_list["A1"]
        cl1.value = f"【{p_name}】全流程收文电子台账明细"
        cl1.font = FONT_TITLE
        cl1.fill = FILL_NAVY_TITLE
        cl1.alignment = ALIGN_CENTER
        ws_sub_list.row_dimensions[1].height = 34

        ws_sub_list.merge_cells("A2:P2")
        cl2 = ws_sub_list["A2"]
        cl2.value = "点击蓝色文件名或“打开原件”即可直接打开对应工程文件；金黄色表头列支持手工填写备注，系统刷新自动保留。"
        cl2.font = FONT_SUBTITLE
        cl2.fill = FILL_SUBHEADER
        cl2.alignment = ALIGN_LEFT
        ws_sub_list.row_dimensions[2].height = 22

        end_sub_r = _write_ledger_table(
            ws_sub_list,
            start_row=3,
            records=pp.all_files,
            rel_prefix_to_strip=strip_prefix,
            include_project_col=False,
        )
        ws_sub_list.auto_filter.ref = f"A3:P{max(4, end_sub_r)}"
        ws_sub_list.freeze_panes = "A4"
        _set_column_widths(ws_sub_list, [6, 11, 12, 14, 48, 26, 22, 14, 24, 16, 28, 14, 14, 30, 14, 20])

        if not safe_save_workbook(wb_sub, sub_excel_path):
            locked_files.append(sub_excel_path)

    return locked_files

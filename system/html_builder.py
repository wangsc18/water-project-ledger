# -*- coding: utf-8 -*-
"""
纯静态单文件《项目状态可视化看板.html》生成器
零外部网络依赖（内网断网可用），双击直接用 Edge / Chrome 打开：
- 多项目进度对比总览 + 单项目八大阶段推进灯带 + 核心必办要件亮灯核查卡片 + 实时搜索台账表
- 原生支持点击超链接在新标签页直接打开 PDF / 图片等本地原件，并支持动态计算当前电脑盘符复制 Windows 路径
"""
from __future__ import annotations

import datetime as dt
import json
from pathlib import Path
from urllib.parse import quote

from config_loader import AppConfig, STAGES_ORDER
from database import LedgerRecord
from progress_calc import PortfolioSummary


def _encode_rel_url(rel_path: str) -> str:
    parts = [quote(seg, safe="") for seg in rel_path.replace("\\", "/").split("/") if seg]
    return "./" + "/".join(parts)


def _record_to_dict(r: LedgerRecord) -> dict:
    ymd = r.doc_date
    ymd_fmt = f"{ymd[:4]}-{ymd[4:6]}-{ymd[6:8]}" if len(ymd) == 8 and ymd.isdigit() else ymd
    return {
        "id": f"REC-{r.id:04d}" if r.id else "",
        "rel_path": r.rel_path.replace("\\", "/"),
        "win_rel_path": r.rel_path.replace("/", "\\"),
        "href": _encode_rel_url(r.rel_path),
        "folder_href": _encode_rel_url(str(Path(r.rel_path).parent)),
        "project_name": r.project_name,
        "stage": r.stage,
        "is_pending": bool(r.is_pending),
        "pending_reason": r.pending_reason or "",
        "archived_filename": r.archived_filename,
        "original_filename": r.original_filename,
        "ext": (r.file_ext or "").upper().lstrip("."),
        "doc_date": ymd_fmt,
        "date_source": r.date_source or "",
        "doc_type": r.doc_type,
        "milestone_key": r.milestone_key or "",
        "doc_no": r.doc_no or "",
        "amount_wan": r.amount_wan or "",
        "remark": r.remark or "",
        "extract_method": r.extract_method or "",
        "archived_at": r.archived_at,
    }


def build_dashboard_html(cfg: AppConfig, summary: PortfolioSummary) -> Path:
    now_str = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    tot_ms_all = sum(p.total_milestones for p in summary.projects)
    lit_ms_all = sum(p.lit_milestones for p in summary.projects)
    overall_pct = round((lit_ms_all / tot_ms_all) * 100.0, 1) if tot_ms_all > 0 else 0.0

    projects_payload = []
    for pp in summary.projects:
        stages_list = []
        for stg_code in STAGES_ORDER:
            sp = pp.stages[stg_code]
            ms_items = []
            for ms in sp.milestones:
                latest = ms.matched_files[-1] if ms.matched_files else None
                ms_items.append({
                    "name": ms.name,
                    "is_lit": ms.is_lit,
                    "file_href": _encode_rel_url(latest.rel_path) if latest else "",
                    "filename": latest.archived_filename if latest else "",
                    "doc_date": (
                        f"{latest.doc_date[:4]}-{latest.doc_date[4:6]}-{latest.doc_date[6:8]}"
                        if latest and len(latest.doc_date) == 8
                        else ""
                    ),
                    "doc_no": latest.doc_no if latest else "",
                })
            stages_list.append({
                "code": sp.stage_code,
                "num": sp.stage_code[:2],
                "name": sp.stage_name,
                "description": sp.description,
                "status": sp.status,
                "status_label": sp.status_label,
                "total_milestones": sp.total_milestones,
                "lit_milestones": sp.lit_milestones,
                "file_count": sp.file_count,
                "milestones": ms_items,
            })

        proj_rel_excel = f"项目归档库/{pp.spec.standard_name}/{pp.spec.standard_name}_电子台账.xlsx"
        projects_payload.append({
            "name": pp.spec.standard_name,
            "project_type": pp.spec.project_type,
            "current_stage_code": pp.current_stage_code,
            "current_stage_label": pp.current_stage_label,
            "total_files": pp.total_files,
            "total_milestones": pp.total_milestones,
            "lit_milestones": pp.lit_milestones,
            "completion_pct": pp.completion_pct,
            "approved_investment_wan": pp.approved_investment_wan,
            "total_amount_contract_wan": pp.total_amount_contract_wan,
            "excel_href": _encode_rel_url(proj_rel_excel),
            "stages": stages_list,
            "files": [_record_to_dict(f) for f in pp.all_files],
        })

    data_json = json.dumps(
        {
            "updated_at": now_str,
            "master_excel_href": _encode_rel_url(cfg.master_excel_path.name),
            "total_projects": summary.total_projects,
            "total_active_files": summary.total_active_files,
            "total_pending_files": summary.total_pending_files,
            "overall_pct": overall_pct,
            "stages_order": STAGES_ORDER,
            "projects": projects_payload,
            "pending_files": [_record_to_dict(f) for f in summary.pending_files],
        },
        ensure_ascii=False,
    )

    html = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>水利新建工程全流程状态可视化看板</title>
<style>
:root {{
  --bg: #f1f5f9;
  --card: #ffffff;
  --navy: #0f3d5e;
  --navy-light: #1e5b88;
  --primary: #0284c7;
  --green: #059669;
  --green-bg: #d1fae5;
  --blue: #2563eb;
  --blue-bg: #dbeafe;
  --amber: #d97706;
  --amber-bg: #fef3c7;
  --gray: #64748b;
  --gray-bg: #f1f5f9;
  --border: #e2e8f0;
  --text: #1e293b;
}}
* {{ box-sizing: border-box; margin: 0; padding: 0; }}
body {{
  font-family: "Microsoft YaHei UI", "Microsoft YaHei", "PingFang SC", -apple-system, sans-serif;
  background: var(--bg);
  color: var(--text);
  line-height: 1.5;
  padding-bottom: 48px;
}}
.topbar {{
  background: linear-gradient(135deg, #0b2e47 0%, #15527d 100%);
  color: #fff;
  padding: 20px 32px;
  display: flex;
  justify-content: space-between;
  align-items: center;
  flex-wrap: wrap;
  gap: 16px;
  box-shadow: 0 4px 14px rgba(15, 61, 94, 0.18);
}}
.topbar-title h1 {{ font-size: 22px; font-weight: 700; letter-spacing: 0.5px; }}
.topbar-title p {{ font-size: 13px; color: #bae6fd; margin-top: 4px; }}
.topbar-actions {{ display: flex; gap: 10px; align-items: center; flex-wrap: wrap; }}
.btn {{
  display: inline-flex; align-items: center; gap: 6px;
  padding: 8px 15px; border-radius: 7px; font-size: 13px; font-weight: 600;
  text-decoration: none; cursor: pointer; border: 1px solid transparent; transition: all 0.15s;
}}
.btn-excel {{ background: #10b981; color: #fff; }}
.btn-excel:hover {{ background: #059669; }}
.btn-outline {{ background: rgba(255,255,255,0.12); color: #fff; border-color: rgba(255,255,255,0.28); }}
.btn-outline:hover {{ background: rgba(255,255,255,0.22); }}
.container {{ max-width: 1440px; margin: 0 auto; padding: 22px 28px; }}

/* KPI 统计条 */
.kpi-grid {{
  display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
  gap: 16px; margin-bottom: 20px;
}}
.kpi-card {{
  background: var(--card); border-radius: 10px; padding: 16px 20px;
  border: 1px solid var(--border); box-shadow: 0 1px 3px rgba(0,0,0,0.04);
  display: flex; justify-content: space-between; align-items: center;
}}
.kpi-label {{ font-size: 13px; color: var(--gray); }}
.kpi-val {{ font-size: 26px; font-weight: 700; color: var(--navy); margin-top: 2px; }}
.kpi-badge {{
  width: 44px; height: 44px; border-radius: 10px;
  display: flex; align-items: center; justify-content: center; font-size: 20px;
}}

/* 导航切换标签 */
.tabs-bar {{
  display: flex; gap: 8px; flex-wrap: wrap; margin-bottom: 20px;
  background: var(--card); padding: 8px; border-radius: 10px; border: 1px solid var(--border);
}}
.tab-btn {{
  padding: 9px 18px; border-radius: 7px; border: none; background: transparent;
  font-size: 14px; font-weight: 600; color: var(--gray); cursor: pointer;
  display: inline-flex; align-items: center; gap: 7px; font-family: inherit;
}}
.tab-btn:hover {{ background: #f8fafc; color: var(--navy); }}
.tab-btn.active {{ background: var(--navy); color: #fff; box-shadow: 0 2px 6px rgba(15,61,94,0.2); }}
.badge-pill {{
  padding: 1px 7px; border-radius: 99px; font-size: 12px; font-weight: 700;
}}
.badge-warn {{ background: #ef4444; color: #fff; }}
.badge-light {{ background: rgba(255,255,255,0.2); color: inherit; }}

/* 多项目总览卡片 */
.overview-card {{
  background: var(--card); border-radius: 12px; padding: 20px 24px;
  border: 1px solid var(--border); margin-bottom: 16px;
  transition: transform 0.15s, box-shadow 0.15s;
}}
.overview-card:hover {{ box-shadow: 0 6px 18px rgba(15,61,94,0.08); }}
.ov-header {{ display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 12px; margin-bottom: 14px; }}
.ov-title {{ font-size: 18px; font-weight: 700; color: var(--navy); cursor: pointer; }}
.ov-title:hover {{ color: var(--primary); text-decoration: underline; }}
.ov-meta {{ display: flex; gap: 16px; font-size: 13px; color: var(--gray); align-items: center; flex-wrap: wrap; }}
.mini-stages {{
  display: grid; grid-template-columns: repeat(8, minmax(110px, 1fr)); gap: 8px; overflow-x: auto;
}}
.mini-stg {{
  padding: 10px 8px; border-radius: 8px; text-align: center; font-size: 12px; border: 1px solid var(--border);
}}
.st-completed {{ background: var(--green-bg); border-color: #6ee7b7; color: #065f46; }}
.st-in_progress {{ background: var(--blue-bg); border-color: #93c5fd; color: #1e40af; }}
.st-not_started {{ background: var(--gray-bg); border-color: #e2e8f0; color: var(--gray); }}

/* 单项目横幅与八大阶段横向推进节点灯带 */
.proj-banner {{
  background: var(--card); border-radius: 12px; padding: 20px 24px;
  border: 1px solid var(--border); margin-bottom: 18px;
}}
.proj-banner-top {{ display: flex; justify-content: space-between; align-items: flex-start; flex-wrap: wrap; gap: 14px; }}
.proj-progress-track {{
  height: 10px; background: #e2e8f0; border-radius: 99px; overflow: hidden; margin-top: 14px;
}}
.proj-progress-bar {{
  height: 100%; background: linear-gradient(90deg, #0284c7, #10b981); border-radius: 99px; transition: width 0.35s;
}}

.stepper-wrap {{
  background: var(--card); border-radius: 12px; padding: 22px 20px;
  border: 1px solid var(--border); margin-bottom: 18px; overflow-x: auto;
}}
.stepper {{
  display: grid; grid-template-columns: repeat(8, minmax(130px, 1fr)); gap: 10px; min-width: 980px;
}}
.step-node {{
  position: relative; border-radius: 10px; padding: 14px 12px; cursor: pointer;
  border: 2px solid var(--border); background: #f8fafc; transition: all 0.15s; text-align: center;
}}
.step-node:hover {{ transform: translateY(-2px); }}
.step-node.active-filter {{ outline: 3px solid #0284c7; outline-offset: 2px; }}
.step-node.completed {{ background: #ecfdf5; border-color: #10b981; }}
.step-node.in_progress {{ background: #eff6ff; border-color: #3b82f6; }}
.step-dot {{
  width: 32px; height: 32px; border-radius: 50%; margin: 0 auto 8px;
  display: flex; align-items: center; justify-content: center; font-weight: 700; font-size: 13px;
}}
.completed .step-dot {{ background: #10b981; color: #fff; }}
.in_progress .step-dot {{ background: #2563eb; color: #fff; box-shadow: 0 0 0 4px rgba(37,99,235,0.2); }}
.not_started .step-dot {{ background: #cbd5e1; color: #475569; }}
.step-title {{ font-size: 14px; font-weight: 700; color: var(--text); }}
.step-sub {{ font-size: 12px; margin-top: 4px; font-weight: 600; }}
.completed .step-sub {{ color: var(--green); }}
.in_progress .step-sub {{ color: var(--blue); }}
.not_started .step-sub {{ color: var(--gray); }}

/* 八大阶段要件矩阵卡片 */
.section-heading {{
  font-size: 16px; font-weight: 700; color: var(--navy);
  margin: 18px 0 12px; display: flex; justify-content: space-between; align-items: center;
}}
.stages-grid {{
  display: grid; grid-template-columns: repeat(auto-fill, minmax(320px, 1fr));
  gap: 14px; margin-bottom: 24px;
}}
.stage-card {{
  background: var(--card); border-radius: 10px; border: 1px solid var(--border);
  padding: 16px; display: flex; flex-direction: column; gap: 10px;
}}
.stage-card-hd {{ display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid #f1f5f9; padding-bottom: 8px; }}
.stage-card-title {{ font-weight: 700; font-size: 15px; color: var(--navy); }}
.status-tag {{ font-size: 12px; font-weight: 700; padding: 2px 9px; border-radius: 99px; }}
.ms-list {{ list-style: none; display: flex; flex-direction: column; gap: 7px; }}
.ms-item {{
  display: flex; justify-content: space-between; align-items: center; gap: 8px;
  font-size: 13px; padding: 6px 9px; border-radius: 6px;
}}
.ms-lit {{ background: #f0fdf4; color: #065f46; border: 1px solid #bbf7d0; }}
.ms-unlit {{ background: #f8fafc; color: #64748b; border: 1px dashed #cbd5e1; }}
.ms-link {{
  color: #0284c7; text-decoration: none; font-weight: 600; font-size: 12px; white-space: nowrap;
}}
.ms-link:hover {{ text-decoration: underline; }}

/* 搜索与明细台账表 */
.table-panel {{
  background: var(--card); border-radius: 12px; border: 1px solid var(--border); overflow: hidden;
}}
.table-toolbar {{
  padding: 14px 18px; border-bottom: 1px solid var(--border);
  display: flex; gap: 12px; flex-wrap: wrap; align-items: center; justify-content: space-between;
}}
.search-input, .stage-select {{
  padding: 8px 12px; border: 1px solid #cbd5e1; border-radius: 7px; font-size: 13px; font-family: inherit;
}}
.search-input {{ min-width: 280px; }}
.table-scroll {{ overflow-x: auto; }}
table {{ width: 100%; border-collapse: collapse; font-size: 13px; }}
th {{
  background: #f8fafc; color: #334155; font-weight: 700; text-align: left;
  padding: 11px 12px; border-bottom: 2px solid var(--border); white-space: nowrap;
}}
td {{ padding: 10px 12px; border-bottom: 1px solid #f1f5f9; vertical-align: middle; }}
tr:hover td {{ background: #f8fafc; }}
.file-link {{ color: #0369a1; font-weight: 600; text-decoration: none; }}
.file-link:hover {{ text-decoration: underline; color: #0284c7; }}
.ext-badge {{
  display: inline-block; padding: 1px 6px; border-radius: 4px; font-size: 11px;
  font-weight: 700; background: #e0f2fe; color: #0369a1; margin-right: 5px;
}}
.ms-star {{
  display: inline-block; padding: 2px 7px; border-radius: 4px; font-size: 11px;
  background: #d1fae5; color: #065f46; font-weight: 600;
}}
.action-btn {{
  display: inline-block; padding: 4px 9px; border-radius: 5px; font-size: 12px;
  text-decoration: none; cursor: pointer; border: 1px solid #cbd5e1; background: #fff; color: #334155;
}}
.action-btn:hover {{ background: #f1f5f9; border-color: #94a3b8; }}
.action-open {{ background: #eff6ff; border-color: #bfdbfe; color: #1d4ed8; font-weight: 600; }}
.toast {{
  position: fixed; bottom: 24px; right: 24px; background: #0f172a; color: #fff;
  padding: 10px 18px; border-radius: 8px; font-size: 13px; opacity: 0;
  pointer-events: none; transition: opacity 0.25s; z-index: 999;
}}
.toast.show {{ opacity: 0.95; }}
.pending-alert {{
  background: #fffbeb; border: 1px solid #fcd34d; color: #92400e;
  padding: 12px 16px; border-radius: 8px; margin-bottom: 16px; font-size: 13px;
}}
</style>
</head>
<body>

<div class="topbar">
  <div class="topbar-title">
    <h1>水利新建工程全流程电子台账 · 状态可视化看板</h1>
    <p id="updated-time"></p>
  </div>
  <div class="topbar-actions">
    <a id="master-excel-btn" class="btn btn-excel" href="#">📊 打开《工程全流程总台账.xlsx》</a>
  </div>
</div>

<div class="container">
  <div class="kpi-grid" id="kpi-container"></div>
  <div class="tabs-bar" id="tabs-container"></div>
  <div id="main-view"></div>
</div>

<div class="toast" id="toast"></div>

<script>
const DATA = {data_json};

let currentTab = DATA.projects.length > 0 ? "__overview__" : "__pending__";
let selectedStageFilter = "ALL";
let searchKeyword = "";

function showToast(msg) {{
  const el = document.getElementById("toast");
  el.textContent = msg;
  el.classList.add("show");
  setTimeout(() => el.classList.remove("show"), 2600);
}}

function escapeHtml(s) {{
  return String(s || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}}

function resolveLocalWindowsPath(winRelPath) {{
  try {{
    let loc = decodeURIComponent(window.location.pathname);
    // Windows file:///D:/folder/项目状态可视化看板.html -> pathname 为 /D:/folder/...
    if (/^\\/[A-Za-z]:\\//.test(loc)) {{
      loc = loc.substring(1);
    }}
    const baseDir = loc.substring(0, loc.lastIndexOf("/")).replace(/\\//g, "\\\\");
    if (baseDir.includes(":\\\\")) {{
      return baseDir + "\\\\" + winRelPath;
    }}
  }} catch (e) {{}}
  return winRelPath;
}}

function copyWinPath(winRelPath) {{
  const fullPath = resolveLocalWindowsPath(winRelPath);
  if (navigator.clipboard && navigator.clipboard.writeText) {{
    navigator.clipboard.writeText(fullPath).then(() => {{
      showToast("已复制本地路径：" + fullPath);
    }}).catch(() => fallbackCopy(fullPath));
  }} else {{
    fallbackCopy(fullPath);
  }}
}}

function fallbackCopy(text) {{
  const ta = document.createElement("textarea");
  ta.value = text;
  document.body.appendChild(ta);
  ta.select();
  document.execCommand("copy");
  document.body.removeChild(ta);
  showToast("已复制本地路径：" + text);
}}

function renderHeaderAndKpis() {{
  document.getElementById("updated-time").textContent =
    "数据自动同步时间：" + DATA.updated_at + "  |  日常操作：将收到的工程文件直接拖入【00_待归档收件箱】即可自动识别更新";
  document.getElementById("master-excel-btn").href = DATA.master_excel_href;

  const kpiHtml = `
    <div class="kpi-card">
      <div><div class="kpi-label">在管水利工程项目</div><div class="kpi-val">${{DATA.total_projects}} <span style="font-size:14px;font-weight:500">个</span></div></div>
      <div class="kpi-badge" style="background:#e0f2fe;color:#0284c7">🌊</div>
    </div>
    <div class="kpi-card">
      <div><div class="kpi-label">已规范归档文件总数</div><div class="kpi-val">${{DATA.total_active_files}} <span style="font-size:14px;font-weight:500">份</span></div></div>
      <div class="kpi-badge" style="background:#dcfce7;color:#059669">📁</div>
    </div>
    <div class="kpi-card">
      <div><div class="kpi-label">总体核心要件齐套率</div><div class="kpi-val">${{DATA.overall_pct}}%</div></div>
      <div class="kpi-badge" style="background:#dbeafe;color:#2563eb">🎯</div>
    </div>
    <div class="kpi-card" style="cursor:pointer" onclick="switchTab('__pending__')">
      <div><div class="kpi-label">99_待人工确认文件</div><div class="kpi-val" style="color:${{DATA.total_pending_files > 0 ? '#d97706' : '#0f3d5e'}}">${{DATA.total_pending_files}} <span style="font-size:14px;font-weight:500">份</span></div></div>
      <div class="kpi-badge" style="background:${{DATA.total_pending_files > 0 ? '#fef3c7' : '#f1f5f9'}};color:#d97706">⚠️</div>
    </div>
  `;
  document.getElementById("kpi-container").innerHTML = kpiHtml;
}}

function renderTabs() {{
  let html = `
    <button class="tab-btn ${{currentTab === '__overview__' ? 'active' : ''}}" onclick="switchTab('__overview__')">
      📊 多项目状态总览
    </button>
  `;
  DATA.projects.forEach((p, idx) => {{
    const pKey = "proj_" + idx;
    html += `
      <button class="tab-btn ${{currentTab === pKey ? 'active' : ''}}" onclick="switchTab('${{pKey}}')">
        🏗️ ${{escapeHtml(p.name)}}
        <span class="badge-pill badge-light">${{p.completion_pct}}%</span>
      </button>
    `;
  }});
  html += `
    <button class="tab-btn ${{currentTab === '__pending__' ? 'active' : ''}}" onclick="switchTab('__pending__')">
      ⚠️ 待人工确认
      <span class="badge-pill ${{DATA.total_pending_files > 0 ? 'badge-warn' : 'badge-light'}}">${{DATA.total_pending_files}}</span>
    </button>
  `;
  document.getElementById("tabs-container").innerHTML = html;
}}

function switchTab(tabId, stageCode = "ALL") {{
  currentTab = tabId;
  selectedStageFilter = stageCode;
  renderTabs();
  renderMainView();
}}

function toggleStageFilter(stageCode) {{
  selectedStageFilter = (selectedStageFilter === stageCode) ? "ALL" : stageCode;
  renderMainView();
}}

function renderOverview() {{
  if (DATA.projects.length === 0) {{
    return `<div class="overview-card">尚未配置在管项目，请在《项目配置.yaml》中添加水利工程名称与匹配关键词。</div>`;
  }}
  let html = "";
  if (DATA.total_pending_files > 0) {{
    html += `
      <div class="pending-alert">
        🔔 当前有 <strong>${{DATA.total_pending_files}}</strong> 份文件因未匹配已知项目关键词暂存于 <strong>99_待人工确认</strong> 文件夹，
        <a href="javascript:void(0)" onclick="switchTab('__pending__')" style="color:#b45309;font-weight:700">点击此处查看处理 →</a>
      </div>
    `;
  }}

  DATA.projects.forEach((p, idx) => {{
    const stagesHtml = p.stages.map(s => `
      <div class="mini-stg st-${{s.status}}" style="cursor:pointer" onclick="switchTab('proj_${{idx}}', '${{s.code}}')">
        <div style="font-weight:700">${{s.num}} ${{escapeHtml(s.name)}}</div>
        <div style="margin-top:3px">${{s.status_label}} (${{s.lit_milestones}}/${{s.total_milestones}})</div>
        <div style="opacity:0.85;margin-top:2px">${{s.file_count}} 份文件</div>
      </div>
    `).join("");

    const invStr = p.approved_investment_wan !== null ? `批复总投资 ${{p.approved_investment_wan.toFixed(2)}} 万元` : "批复总投资 待收文提取";
    html += `
      <div class="overview-card">
        <div class="ov-header">
          <div>
            <span class="ov-title" onclick="switchTab('proj_${{idx}}')">${{escapeHtml(p.name)}}</span>
            <span style="margin-left:10px;font-size:12px;padding:3px 9px;border-radius:99px;background:#e0f2fe;color:#0369a1;font-weight:600">${{escapeHtml(p.project_type)}}</span>
          </div>
          <div class="ov-meta">
            <span>当前推进：<strong>${{escapeHtml(p.current_stage_label)}}</strong></span>
            <span>核心要件齐套率：<strong>${{p.completion_pct}}% (${{p.lit_milestones}}/${{p.total_milestones}})</strong></span>
            <span>已归档：<strong>${{p.total_files}} 份</strong></span>
            <span>${{invStr}}</span>
            <a class="action-btn action-open" href="${{p.excel_href}}">📊 本项目独立Excel</a>
            <button class="action-btn" onclick="switchTab('proj_${{idx}}')">进入详情 →</button>
          </div>
        </div>
        <div class="mini-stages">${{stagesHtml}}</div>
      </div>
    `;
  }});

  // 底部附全部项目汇总检索表
  const allActive = [];
  DATA.projects.forEach(p => p.files.forEach(f => allActive.push(f)));
  html += renderLedgerTableSection("全部在管项目归档文件总览检索", allActive, true);
  return html;
}}

function renderSingleProject(projIdx) {{
  const p = DATA.projects[projIdx];
  if (!p) return "";

  const invStr = p.approved_investment_wan !== null ? `${{p.approved_investment_wan.toFixed(2)}} 万元` : "待收文提取";
  const stepperNodes = p.stages.map(s => {{
    const icon = s.status === "completed" ? "✓" : s.num;
    const isSel = selectedStageFilter === s.code ? "active-filter" : "";
    return `
      <div class="step-node ${{s.status}} ${{isSel}}" onclick="toggleStageFilter('${{s.code}}')" title="点击只看【${{s.code}}】阶段">
        <div class="step-dot">${{icon}}</div>
        <div class="step-title">${{s.num}} ${{escapeHtml(s.name)}}</div>
        <div class="step-sub">${{s.status_label}} (${{s.lit_milestones}}/${{s.total_milestones}})</div>
        <div style="font-size:11px;color:#64748b;margin-top:3px">已收 ${{s.file_count}} 份</div>
      </div>
    `;
  }}).join("");

  const visibleStages = selectedStageFilter === "ALL"
    ? p.stages
    : p.stages.filter(s => s.code === selectedStageFilter);

  const cardsHtml = visibleStages.map(s => {{
    const tagCls = s.status === "completed" ? "st-completed" : (s.status === "in_progress" ? "st-in_progress" : "st-not_started");
    const msRows = s.milestones.map(m => {{
      if (m.is_lit) {{
        const tip = [m.doc_date, m.doc_no].filter(Boolean).join(" | ");
        return `
          <li class="ms-item ms-lit" title="${{escapeHtml(m.filename)}} ${{escapeHtml(tip)}}">
            <span>✅ ${{escapeHtml(m.name)}}</span>
            <a class="ms-link" href="${{m.file_href}}" target="_blank">打开批文/原件 ↗</a>
          </li>
        `;
      }} else {{
        return `
          <li class="ms-item ms-unlit">
            <span>○ ${{escapeHtml(m.name)}}</span>
            <span style="font-size:11px">待收件</span>
          </li>
        `;
      }}
    }}).join("");

    return `
      <div class="stage-card">
        <div class="stage-card-hd">
          <span class="stage-card-title">${{escapeHtml(s.code)}}</span>
          <span class="status-tag ${{tagCls}}">${{s.status_label}} · ${{s.lit_milestones}}/${{s.total_milestones}} 要件 · ${{s.file_count}}份</span>
        </div>
        <div style="font-size:12px;color:#64748b">${{escapeHtml(s.description)}}</div>
        <ul class="ms-list">${{msRows}}</ul>
      </div>
    `;
  }}).join("");

  const filterResetBtn = selectedStageFilter !== "ALL"
    ? `<button class="action-btn" onclick="toggleStageFilter('ALL')">✕ 取消阶段筛选（当前仅看 ${{selectedStageFilter}}）</button>`
    : `<span style="font-size:12px;color:#64748b">提示：点击上方任意流程节点可按阶段筛选下方卡片与文件</span>`;

  return `
    <div class="proj-banner">
      <div class="proj-banner-top">
        <div>
          <h2 style="font-size:20px;color:#0f3d5e">${{escapeHtml(p.name)}}</h2>
          <div style="font-size:13px;color:#475569;margin-top:5px">
            工程类别：<strong>${{escapeHtml(p.project_type)}}</strong> &nbsp;|&nbsp;
            当前阶段：<strong>${{escapeHtml(p.current_stage_label)}}</strong> &nbsp;|&nbsp;
            批复总投资：<strong>${{invStr}}</strong> &nbsp;|&nbsp;
            已签合同累计：<strong>${{p.total_amount_contract_wan.toFixed(2)}} 万元</strong>
          </div>
        </div>
        <div style="display:flex;gap:10px;align-items:center">
          <div style="text-align:right;margin-right:8px">
            <div style="font-size:12px;color:#64748b">核心要件完成度</div>
            <div style="font-size:22px;font-weight:700;color:#059669">${{p.completion_pct}}% <span style="font-size:13px;color:#64748b">(${{p.lit_milestones}}/${{p.total_milestones}})</span></div>
          </div>
          <a class="btn btn-excel" href="${{p.excel_href}}">📊 打开本项目独立 Excel 台账</a>
        </div>
      </div>
      <div class="proj-progress-track"><div class="proj-progress-bar" style="width:${{p.completion_pct}}%"></div></div>
    </div>

    <div class="stepper-wrap">
      <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:12px">
        <strong style="color:#0f3d5e;font-size:15px">全生命周期八大阶段推进状态灯带</strong>
        ${{filterResetBtn}}
      </div>
      <div class="stepper">${{stepperNodes}}</div>
    </div>

    <div class="section-heading">
      <span>核心必办亮灯要件核查清单</span>
    </div>
    <div class="stages-grid">${{cardsHtml}}</div>

    ${{renderLedgerTableSection("【" + p.name + "】已归档文件电子台账明细", p.files, false)}}
  `;
}}

function renderPendingView() {{
  return `
    <div class="pending-alert" style="background:#fff7ed;border-color:#fdba74">
      <strong>🛠️ 99_待人工确认文件纠偏指南：</strong><br>
      1. 若属于<strong>新开工水利工程</strong>：用记事本打开根目录《项目配置.yaml》，将该工程名称与匹配关键词加入白名单，保存后双击《立即扫描校准一次.bat》。<br>
      2. 若属于<strong>现有项目偶发未识别</strong>：直接在 Windows 资源管理器中，把文件从 <code>99_待人工确认</code> 拖进对应项目的 <code>01_立项可研 ~ 08_结算归档</code> 文件夹（可顺手修改文件名），系统将自动同步更新台账与超链接！
    </div>
    ${{renderLedgerTableSection("99_待人工确认文件列表", DATA.pending_files, true)}}
  `;
}}

function renderLedgerTableSection(title, filesList, showProjCol) {{
  const filtered = filesList.filter(f => {{
    if (selectedStageFilter !== "ALL" && f.stage !== selectedStageFilter) return false;
    if (!searchKeyword.trim()) return true;
    const kw = searchKeyword.trim().toLowerCase();
    const hay = [
      f.project_name, f.stage, f.archived_filename, f.doc_type,
      f.milestone_key, f.doc_no, f.amount_wan, f.remark, f.original_filename
    ].join(" ").toLowerCase();
    return hay.includes(kw);
  }});

  const stageOptions = ['<option value="ALL">全部流程阶段</option>']
    .concat(DATA.stages_order.map(sc => `<option value="${{sc}}" ${{selectedStageFilter === sc ? 'selected' : ''}}>${{sc}}</option>`))
    .join("");

  const rowsHtml = filtered.length === 0
    ? `<tr><td colspan="11" style="text-align:center;color:#94a3b8;padding:28px">暂无匹配的归档文件记录</td></tr>`
    : filtered.map((f, idx) => `
      <tr>
        <td>${{idx + 1}}</td>
        ${{showProjCol ? `<td><strong>${{escapeHtml(f.project_name)}}</strong></td>` : ""}}
        <td>${{escapeHtml(f.stage)}}</td>
        <td>
          <span class="ext-badge">${{escapeHtml(f.ext)}}</span>
          <a class="file-link" href="${{f.href}}" target="_blank" title="原文件名：${{escapeHtml(f.original_filename)}}">${{escapeHtml(f.archived_filename)}}</a>
        </td>
        <td>${{escapeHtml(f.doc_type)}}</td>
        <td>${{f.milestone_key ? `<span class="ms-star">★ ${{escapeHtml(f.milestone_key)}}</span>` : '<span style="color:#cbd5e1">-</span>'}}</td>
        <td style="white-space:nowrap">${{escapeHtml(f.doc_date)}}</td>
        <td>${{escapeHtml(f.doc_no || "-")}}</td>
        <td style="text-align:right">${{f.amount_wan ? escapeHtml(f.amount_wan) + " 万" : "-"}}</td>
        <td>${{escapeHtml(f.remark || f.pending_reason || "-")}}</td>
        <td style="white-space:nowrap">
          <a class="action-btn action-open" href="${{f.href}}" target="_blank">📄 打开原件</a>
          <button class="action-btn" onclick="copyWinPath('${{f.win_rel_path.replace(/\\\\/g, '\\\\\\\\').replace(/'/g, &quot;\\\\'&quot;)}}')">📋 复制路径</button>
        </td>
      </tr>
    `).join("");

  return `
    <div class="table-panel">
      <div class="table-toolbar">
        <strong style="font-size:15px;color:#0f3d5e">${{escapeHtml(title)}}（共 ${{filtered.length}} 份）</strong>
        <div style="display:flex;gap:10px;flex-wrap:wrap">
          <select class="stage-select" onchange="onStageSelectChange(this.value)">${{stageOptions}}</select>
          <input class="search-input" type="text" placeholder="🔍 搜索文件名、文号、标段、金额、原文件名..." value="${{escapeHtml(searchKeyword)}}" oninput="onSearchInput(this.value)">
        </div>
      </div>
      <div class="table-scroll">
        <table>
          <thead>
            <tr>
              <th>序号</th>
              ${{showProjCol ? "<th>所属项目</th>" : ""}}
              <th>流程阶段</th>
              <th>规范归档文件名（点击直接打开）</th>
              <th>文件具体类型</th>
              <th>点亮核心要件</th>
              <th>成文/签署日期</th>
              <th>文号/合同编号</th>
              <th style="text-align:right">涉及金额</th>
              <th>备注 / 说明</th>
              <th>快捷操作</th>
            </tr>
          </thead>
          <tbody>${{rowsHtml}}</tbody>
        </table>
      </div>
    </div>
  `;
}}

function onStageSelectChange(val) {{
  selectedStageFilter = val;
  renderMainView();
}}

function onSearchInput(val) {{
  searchKeyword = val;
  renderMainView();
  // 保持输入框焦点
  const input = document.querySelector(".search-input");
  if (input) {{
    input.focus();
    input.setSelectionRange(input.value.length, input.value.length);
  }}
}}

function renderMainView() {{
  const container = document.getElementById("main-view");
  if (currentTab === "__overview__") {{
    container.innerHTML = renderOverview();
  }} else if (currentTab === "__pending__") {{
    container.innerHTML = renderPendingView();
  }} else if (currentTab.startsWith("proj_")) {{
    const idx = parseInt(currentTab.split("_")[1], 10);
    container.innerHTML = renderSingleProject(idx);
  }}
}}

renderHeaderAndKpis();
renderTabs();
renderMainView();
</script>
</body>
</html>
"""
    cfg.dashboard_html_path.write_text(html, encoding="utf-8")
    return cfg.dashboard_html_path

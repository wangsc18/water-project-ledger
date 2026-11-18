# -*- coding: utf-8 -*-
"""
项目全流程状态与核心必办要件亮灯计算模块（供 Excel 台账与静态 HTML 看板共用，保证双端数据 100% 一致）
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, List, Optional

from config_loader import (
    AppConfig,
    ProjectSpec,
    STAGE_DESCRIPTIONS,
    STAGE_DISPLAY_NAMES,
    STAGES_ORDER,
)
from database import LedgerRecord


@dataclass
class MilestoneStatus:
    name: str
    is_lit: bool
    matched_files: List[LedgerRecord] = field(default_factory=list)


@dataclass
class StageProgress:
    stage_code: str
    stage_name: str
    description: str
    status: str           # "completed" (已齐备) | "in_progress" (进行中) | "not_started" (未开始)
    status_label: str     # "已齐备" | "进行中" | "未开始"
    total_milestones: int
    lit_milestones: int
    file_count: int
    milestones: List[MilestoneStatus] = field(default_factory=list)
    files: List[LedgerRecord] = field(default_factory=list)


@dataclass
class ProjectProgress:
    spec: ProjectSpec
    current_stage_code: str
    current_stage_label: str
    total_files: int
    total_milestones: int
    lit_milestones: int
    completion_pct: float
    total_amount_contract_wan: float
    approved_investment_wan: Optional[float]
    stages: Dict[str, StageProgress] = field(default_factory=dict)
    all_files: List[LedgerRecord] = field(default_factory=list)


@dataclass
class PortfolioSummary:
    projects: List[ProjectProgress]
    pending_files: List[LedgerRecord]
    active_files: List[LedgerRecord]
    total_projects: int
    total_active_files: int
    total_pending_files: int


def compute_portfolio_summary(cfg: AppConfig, all_records: List[LedgerRecord]) -> PortfolioSummary:
    pending_files = [r for r in all_records if r.is_pending]
    active_files = [r for r in all_records if not r.is_pending]

    # 合并配置中的白名单项目与归档库物理目录可能存在的额外项目名
    proj_map: Dict[str, ProjectSpec] = {p.standard_name: p for p in cfg.projects}
    for r in active_files:
        if r.project_name not in proj_map:
            proj_map[r.project_name] = ProjectSpec(standard_name=r.project_name)

    project_progress_list: List[ProjectProgress] = []

    for proj_name, spec in proj_map.items():
        p_files = [r for r in active_files if r.project_name == proj_name]
        stages_dict: Dict[str, StageProgress] = {}
        proj_total_ms = 0
        proj_lit_ms = 0
        furthest_active_stage = ""
        all_stages_completed = True

        for stg in STAGES_ORDER:
            s_files = [r for r in p_files if r.stage == stg]
            req_ms_names = spec.get_stage_milestones(stg)
            ms_statuses: List[MilestoneStatus] = []
            lit_cnt = 0

            for ms_name in req_ms_names:
                matched = [
                    f for f in s_files
                    if f.milestone_key == ms_name or (ms_name in f.doc_type)
                ]
                is_lit = len(matched) > 0
                if is_lit:
                    lit_cnt += 1
                ms_statuses.append(MilestoneStatus(name=ms_name, is_lit=is_lit, matched_files=matched))

            tot_ms = len(req_ms_names)
            proj_total_ms += tot_ms
            proj_lit_ms += lit_cnt

            if tot_ms > 0 and lit_cnt == tot_ms:
                status = "completed"
                status_label = "已齐备"
            elif len(s_files) > 0 or lit_cnt > 0:
                status = "in_progress"
                status_label = "进行中"
                all_stages_completed = False
            else:
                status = "not_started"
                status_label = "未开始"
                all_stages_completed = False

            if len(s_files) > 0:
                furthest_active_stage = stg

            stages_dict[stg] = StageProgress(
                stage_code=stg,
                stage_name=STAGE_DISPLAY_NAMES.get(stg, stg),
                description=STAGE_DESCRIPTIONS.get(stg, ""),
                status=status,
                status_label=status_label,
                total_milestones=tot_ms,
                lit_milestones=lit_cnt,
                file_count=len(s_files),
                milestones=ms_statuses,
                files=s_files,
            )

        if not p_files:
            curr_code = "00_尚未启动"
            curr_label = "尚未收文启动"
        elif all_stages_completed and proj_total_ms > 0:
            curr_code = "08_结算归档"
            curr_label = "全流程已齐备归档"
        else:
            curr_code = furthest_active_stage or "01_立项可研"
            stg_obj = stages_dict[curr_code]
            curr_label = f"{curr_code}（{stg_obj.status_label}）"

        pct = round((proj_lit_ms / proj_total_ms) * 100.0, 1) if proj_total_ms > 0 else 0.0

        # 统计概算批复金额与合同累计金额参考值
        approved_inv: Optional[float] = None
        contract_sum = 0.0
        for f in p_files:
            if not f.amount_wan:
                continue
            try:
                val = float(str(f.amount_wan).replace(",", ""))
            except ValueError:
                continue
            if f.stage in ("01_立项可研", "03_勘察设计") and "批复" in f.doc_type:
                approved_inv = val
            elif f.stage == "04_招采合同" and "合同" in f.doc_type:
                contract_sum += val

        project_progress_list.append(
            ProjectProgress(
                spec=spec,
                current_stage_code=curr_code,
                current_stage_label=curr_label,
                total_files=len(p_files),
                total_milestones=proj_total_ms,
                lit_milestones=proj_lit_ms,
                completion_pct=pct,
                total_amount_contract_wan=round(contract_sum, 2),
                approved_investment_wan=approved_inv,
                stages=stages_dict,
                all_files=p_files,
            )
        )

    return PortfolioSummary(
        projects=project_progress_list,
        pending_files=pending_files,
        active_files=active_files,
        total_projects=len(project_progress_list),
        total_active_files=len(active_files),
        total_pending_files=len(pending_files),
    )

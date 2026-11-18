# -*- coding: utf-8 -*-
"""
配置加载与水利工程八大阶段标准要件定义模块
支持 Python 源码模式与 PyInstaller 打包后的 .exe 独立运行模式
"""
from __future__ import annotations

import os
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional
import yaml

# 水利新建工程标准八大阶段及默认核心必办亮灯要件
STAGES_ORDER: List[str] = [
    "01_立项可研",
    "02_用地规划",
    "03_勘察设计",
    "04_招采合同",
    "05_施工许可",
    "06_过程管控",
    "07_竣工验收",
    "08_结算归档",
]

STAGE_DISPLAY_NAMES: Dict[str, str] = {
    "01_立项可研": "立项可研",
    "02_用地规划": "用地规划",
    "03_勘察设计": "勘察设计",
    "04_招采合同": "招采合同",
    "05_施工许可": "施工许可",
    "06_过程管控": "过程管控",
    "07_竣工验收": "竣工验收",
    "08_结算归档": "结算归档",
    "99_待人工确认": "待人工确认",
}

STAGE_DESCRIPTIONS: Dict[str, str] = {
    "01_立项可研": "项目建议书、可研报告及批复、水保方案、环评批复、水资源论证、防洪评价等",
    "02_用地规划": "用地预审与选址意见书、建设用地规划许可、征地移民安置批复、不动产权证等",
    "03_勘察设计": "工程地质/水文勘察、初步设计及概算批复、施工图设计文件、施工图审查合格书等",
    "04_招采合同": "招标文件、控制价、评标报告、中标通知书、施工/监理/勘察设计合同、履约保函等",
    "05_施工许可": "项目法人组建、水利工程质量安全监督书、施工许可证或水利开工备案批复、监理规划等",
    "06_过程管控": "总监开工令、设计交底、联系单、设计变更、工程签证、进度款支付证书、质量检测与评定等",
    "07_竣工验收": "分部/单位工程验收、水保/环保/档案专项验收、合同工程完工验收鉴定书、竣工验收鉴定书等",
    "08_结算归档": "工程竣工结算书、结算审核报告、竣工财务决算审计报告、质保金凭证、工程档案移交书等",
}

# 默认通用水利基建核心必办要件（收齐某阶段全部要件即点亮绿灯“已齐备”）
DEFAULT_STAGE_MILESTONES: Dict[str, List[str]] = {
    "01_立项可研": [
        "项目建议书批复",
        "可行性研究报告批复",
        "水土保持方案批复",
        "环境影响评价批复",
    ],
    "02_用地规划": [
        "用地预审与选址意见书",
        "建设用地规划许可证",
        "征地移民安置批复",
    ],
    "03_勘察设计": [
        "工程地质勘察报告",
        "初步设计及概算批复",
        "施工图审查合格书",
    ],
    "04_招采合同": [
        "施工招标文件",
        "中标通知书",
        "施工承包合同",
        "监理合同",
    ],
    "05_施工许可": [
        "水利工程质量安全监督手续",
        "施工许可证或开工批复",
    ],
    "06_过程管控": [
        "总监开工令",
        "设计变更或工程签证",
        "进度款支付证书",
    ],
    "07_竣工验收": [
        "合同工程完工验收鉴定书",
        "专项验收通过文件",
        "竣工验收鉴定书",
    ],
    "08_结算归档": [
        "工程结算审核报告",
        "竣工财务决算审计报告",
        "工程档案移交书",
    ],
}

INBOX_DIR_NAME = "00_待归档收件箱"
ARCHIVE_LIBRARY_DIR_NAME = "项目归档库"
PENDING_DIR_NAME = "99_待人工确认"
MASTER_EXCEL_NAME = "工程全流程总台账.xlsx"
DASHBOARD_HTML_NAME = "项目状态可视化看板.html"
CONFIG_FILE_NAME = "项目配置.yaml"


def resolve_root_dir(explicit_root: Optional[str] = None) -> Path:
    """
    自动定位“水利工程电子台账”根目录，兼容：
    1. 显式传入 --root
    2. Python 源码运行: WaterLedger/system/main.py -> parent.parent
    3. PyInstaller 打包运行: WaterLedger/system/runtime/ledger_engine.exe -> 向上查找 项目配置.yaml
    """
    if explicit_root:
        return Path(explicit_root).resolve()

    if getattr(sys, "frozen", False):
        exe_dir = Path(sys.executable).resolve().parent
        # 向上最多查找 4 层寻找 项目配置.yaml
        curr = exe_dir
        for _ in range(4):
            if (curr / CONFIG_FILE_NAME).exists() or (curr / INBOX_DIR_NAME).exists():
                return curr
            if curr.parent == curr:
                break
            curr = curr.parent
        # 默认回退：若在 system/runtime/ledger_engine.exe 则上退两层
        if exe_dir.name.lower() == "runtime" and exe_dir.parent.name.lower() == "system":
            return exe_dir.parent.parent
        return exe_dir.parent

    # 源码模式：本文件在 <root>/system/config_loader.py
    return Path(__file__).resolve().parent.parent


@dataclass
class ProjectSpec:
    standard_name: str
    project_type: str = "水利工程"
    keywords: List[str] = field(default_factory=list)
    skip_milestones: List[str] = field(default_factory=list)
    add_milestones: Dict[str, List[str]] = field(default_factory=dict)

    def get_stage_milestones(self, stage: str) -> List[str]:
        base = list(DEFAULT_STAGE_MILESTONES.get(stage, []))
        skip_set = set(self.skip_milestones or [])
        filtered = [m for m in base if m not in skip_set]
        extra = (self.add_milestones or {}).get(stage, [])
        for item in extra:
            if item and item not in filtered:
                filtered.append(item)
        return filtered

    def all_match_terms(self) -> List[str]:
        terms = [self.standard_name]
        for kw in self.keywords:
            kw_clean = str(kw).strip()
            if kw_clean and kw_clean not in terms:
                terms.append(kw_clean)
        return terms


@dataclass
class AILLMConfig:
    enabled: bool = False
    mode: str = "fallback"  # "fallback" | "always"
    api_base: str = "https://api.deepseek.com/v1"
    api_key: str = ""
    model: str = "deepseek-chat"
    timeout_seconds: int = 15


@dataclass
class AppConfig:
    root_dir: Path
    projects: List[ProjectSpec] = field(default_factory=list)
    watch_interval_seconds: float = 2.0
    enable_windows_notification: bool = True
    max_doc_type_length: int = 34
    ai_llm: AILLMConfig = field(default_factory=AILLMConfig)

    @property
    def inbox_dir(self) -> Path:
        return self.root_dir / INBOX_DIR_NAME

    @property
    def archive_lib_dir(self) -> Path:
        return self.root_dir / ARCHIVE_LIBRARY_DIR_NAME

    @property
    def pending_dir(self) -> Path:
        return self.root_dir / PENDING_DIR_NAME

    @property
    def master_excel_path(self) -> Path:
        return self.root_dir / MASTER_EXCEL_NAME

    @property
    def dashboard_html_path(self) -> Path:
        return self.root_dir / DASHBOARD_HTML_NAME

    @property
    def system_dir(self) -> Path:
        return self.root_dir / "system"

    @property
    def data_dir(self) -> Path:
        return self.system_dir / "data"

    @property
    def db_path(self) -> Path:
        return self.data_dir / "ledger.db"

    def get_project_map(self) -> Dict[str, ProjectSpec]:
        return {p.standard_name: p for p in self.projects}


def sanitize_windows_name(name: str) -> str:
    """清理 Windows 文件/文件夹非法字符"""
    illegal = '\\/:*?"<>|\r\n\t'
    cleaned = "".join("_" if ch in illegal else ch for ch in str(name).strip())
    return cleaned.strip(" .")


def load_config(explicit_root: Optional[str] = None) -> AppConfig:
    root_dir = resolve_root_dir(explicit_root)
    cfg_path = root_dir / CONFIG_FILE_NAME
    raw = {}
    if cfg_path.exists():
        with open(cfg_path, "r", encoding="utf-8-sig") as f:
            raw = yaml.safe_load(f) or {}

    projects: List[ProjectSpec] = []
    for item in raw.get("projects", []) or []:
        if not isinstance(item, dict):
            continue
        std_name = sanitize_windows_name(item.get("standard_name", "")).replace("-", "_")
        if not std_name:
            continue
        projects.append(
            ProjectSpec(
                standard_name=std_name,
                project_type=str(item.get("project_type", "水利工程")).strip(),
                keywords=[str(k).strip() for k in (item.get("keywords") or []) if str(k).strip()],
                skip_milestones=[str(m).strip() for m in (item.get("skip_milestones") or []) if str(m).strip()],
                add_milestones=item.get("add_milestones") or {},
            )
        )

    settings = raw.get("settings") or {}
    ai_raw = raw.get("ai_llm") or {}
    ai_cfg = AILLMConfig(
        enabled=bool(ai_raw.get("enabled", False)),
        mode=str(ai_raw.get("mode", "fallback")).strip().lower(),
        api_base=str(ai_raw.get("api_base", "https://api.deepseek.com/v1")).strip().rstrip("/"),
        api_key=str(ai_raw.get("api_key", "")).strip(),
        model=str(ai_raw.get("model", "deepseek-chat")).strip(),
        timeout_seconds=int(ai_raw.get("timeout_seconds", 15)),
    )

    return AppConfig(
        root_dir=root_dir,
        projects=projects,
        watch_interval_seconds=float(settings.get("watch_interval_seconds", 2.0)),
        enable_windows_notification=bool(settings.get("enable_windows_notification", True)),
        max_doc_type_length=int(settings.get("max_doc_type_length", 34)),
        ai_llm=ai_cfg,
    )


def ensure_directory_skeleton(cfg: AppConfig) -> None:
    """确保收件箱、待确认目录、数据目录以及所有白名单项目的 01~08 阶段目录全部就绪"""
    cfg.inbox_dir.mkdir(parents=True, exist_ok=True)
    cfg.archive_lib_dir.mkdir(parents=True, exist_ok=True)
    cfg.pending_dir.mkdir(parents=True, exist_ok=True)
    cfg.data_dir.mkdir(parents=True, exist_ok=True)

    readme_inbox = cfg.inbox_dir / "_收件箱投递说明.txt"
    if not readme_inbox.exists():
        content = (
            "【00_待归档收件箱 使用说明】\r\n"
            "------------------------------------------------------------\r\n"
            "1. 日常收到任何工程文件（PDF电子件、红头盖章扫描件、Word合同、Excel报表、现场图片、ZIP整包等），\r\n"
            "   无需手动改名，直接拖入或保存到本文件夹即可！\r\n"
            "2. 若已启动《启动后台自动监控(无黑框).vbs》，约 2~5 秒内系统会自动：\r\n"
            "   - 提取正文或执行离线中文 OCR 识别\r\n"
            "   - 按“年月日-项目名称-文件具体类型”规范重命名\r\n"
            "   - 自动归入《项目归档库\\对应项目\\01~08阶段》文件夹（未匹配已知项目则归入 99_待人工确认）\r\n"
            "   - 自动更新《工程全流程总台账.xlsx》、各项目独立台账及《项目状态可视化看板.html》\r\n"
            "3. 若未开启后台常驻监听，随时双击根目录的《立即扫描校准一次.bat》即可一键处理。\r\n"
            "（注：下划线 _ 开头的文件会被监听程序自动忽略，请勿删除本说明或以 _ 命名正式文件）\r\n"
        )
        readme_inbox.write_text(content, encoding="utf-8")

    for proj in cfg.projects:
        proj_dir = cfg.archive_lib_dir / proj.standard_name
        proj_dir.mkdir(parents=True, exist_ok=True)
        for stage in STAGES_ORDER:
            (proj_dir / stage).mkdir(parents=True, exist_ok=True)

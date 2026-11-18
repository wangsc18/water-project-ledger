# -*- coding: utf-8 -*-
"""
水利新建工程全流程智能分类、要素抽取与三段式规范重命名引擎
支持：
1. 在管项目白名单碰撞归一（未匹配自动标记进入 99_待人工确认）
2. 标题强权重 + 正文辅助的水利八大阶段与核心必办要件精准判定（含水保方案vs水保验收等防混消歧）
3. 标段 / 专业 / 期次 / 水工建筑物部位特征词提取
4. 公文落款成文日期（含中文大写日期“二〇二六年十月二十八日”）+ 文号/合同编号 + 金额(万元)抽取
5. 可选兼容 OpenAI 协议的大模型 API 结构化增强（断网/超时无缝回退本地规则）
"""
from __future__ import annotations

import datetime as dt
import json
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from config_loader import (
    AppConfig,
    DEFAULT_STAGE_MILESTONES,
    STAGES_ORDER,
    sanitize_windows_name,
)
from extractor import ExtractedContent


@dataclass
class ClassificationResult:
    project_name: str
    is_pending: bool
    pending_reason: str
    stage: str
    doc_type: str
    milestone_key: str
    doc_date: str          # 8位 YYYYMMDD
    date_source: str       # "正文成文/签署日期" | "原文件名日期" | "收件当天日期"
    doc_no: str
    amount_wan: str
    confidence: float
    standard_filename: str


# ---------------------------------------------------------------------------
# 水利工程业务文种规则库（按特异性与业务优先级精心编排）
# 字段说明：(阶段, 标准具体文种名, 对应核心亮灯要件key, 必含词组列表[任一满足], 排斥词列表, 基础分)
# 必含词组支持单个词 "中标通知书" 或组合与关系 "可行性研究+批复"
# ---------------------------------------------------------------------------
RAW_DOC_RULES: List[Tuple[str, str, str, List[str], List[str], float]] = [
    # ================= 07_竣工验收（前置高优：防止“水保验收/环保验收/档案验收”误入01立项或08归档） =================
    ("07_竣工验收", "竣工验收鉴定书", "竣工验收鉴定书",
     ["竣工验收鉴定书", "工程竣工验收报告+鉴定"], ["合同工程完工", "分部工程", "单位工程"], 98),
    ("07_竣工验收", "竣工技术预验收报告", "",
     ["技术预验收工作报告", "竣工技术预验收"], [], 92),
    ("07_竣工验收", "合同工程完工验收鉴定书", "合同工程完工验收鉴定书",
     ["合同工程完工验收鉴定书", "完工验收鉴定书", "投入使用验收鉴定书"], [], 96),
    ("07_竣工验收", "单位工程验收鉴定书", "合同工程完工验收鉴定书",
     ["单位工程验收鉴定书", "单位工程外观质量评定"], [], 93),
    ("07_竣工验收", "分部工程验收鉴定书", "",
     ["分部工程验收鉴定书", "法人验收鉴定书+分部"], [], 92),
    ("07_竣工验收", "阶段验收鉴定书", "",
     ["阶段验收鉴定书", "下闸蓄水验收", "通水验收鉴定书", "截流验收鉴定书"], [], 92),
    ("07_竣工验收", "水土保持设施验收鉴定书", "专项验收通过文件",
     ["水土保持设施验收鉴定书", "水土保持设施自主验收", "水保设施验收", "水土保持专项验收"], [], 96),
    ("07_竣工验收", "环境保护设施竣工验收意见", "专项验收通过文件",
     ["环境保护验收意见", "竣工环境保护验收", "环保专项验收", "环保设施验收"], [], 95),
    ("07_竣工验收", "工程档案专项验收意见", "专项验收通过文件",
     ["档案专项验收", "档案验收意见", "档案验收鉴定"], [], 95),
    ("07_竣工验收", "征地移民安置验收意见", "专项验收通过文件",
     ["移民安置验收", "征地移民专项验收"], [], 93),
    ("07_竣工验收", "工程质量监督报告", "",
     ["工程质量监督报告", "竣工验收质量监督报告"], ["监督书", "监督登记"], 88),

    # ================= 08_结算归档 =================
    ("08_结算归档", "竣工财务决算审计报告", "竣工财务决算审计报告",
     ["竣工财务决算审计", "竣工决算审计报告", "决算批复", "财务决算批复"], [], 96),
    ("08_结算归档", "竣工财务决算报告", "",
     ["竣工财务决算报告", "竣工财务决算说明书", "竣工决算报表"], ["审计"], 88),
    ("08_结算归档", "工程结算审核报告", "工程结算审核报告",
     ["结算审核报告", "结算审定单", "结算审核定案表", "竣工结算审计报告", "结算评审报告"], [], 95),
    ("08_结算归档", "工程竣工结算书", "",
     ["竣工结算书", "工程结算书", "完工结算申请"], ["审核", "审计"], 86),
    ("08_结算归档", "工程档案移交书", "工程档案移交书",
     ["档案移交书", "档案移交证明", "档案接收证明", "城建档案移交", "水利档案移交"], [], 94),
    ("08_结算归档", "工程质量保证金退还凭证", "",
     ["质保金退还", "质量保证金返还", "缺陷责任期满证书"], [], 86),

    # ================= 01_立项可研 =================
    ("01_立项可研", "项目建议书批复", "项目建议书批复",
     ["项目建议书+批复", "立项批复", "同意立项", "项目立项审批"], [], 95),
    ("01_立项可研", "项目建议书", "",
     ["项目建议书"], ["批复"], 82),
    ("01_立项可研", "可行性研究报告批复", "可行性研究报告批复",
     ["可行性研究报告+批复", "可研报告+批复", "可行性研究+批复", "可研批复"], [], 96),
    ("01_立项可研", "可行性研究报告", "",
     ["可行性研究报告", "工程可研报告"], ["批复", "招标"], 83),
    ("01_立项可研", "水土保持方案批复", "水土保持方案批复",
     ["水土保持方案+批复", "水土保持方案+审批准予", "水土保持方案+行政许可", "水保方案批复"], ["设施验收"], 96),
    ("01_立项可研", "水土保持方案报告书", "",
     ["水土保持方案报告书", "水土保持方案报告表"], ["批复", "设施验收"], 83),
    ("01_立项可研", "环境影响评价批复", "环境影响评价批复",
     ["环境影响报告书+批复", "环境影响报告表+批复", "环评批复", "环境影响评价+审批意见"], ["竣工环境保护验收"], 95),
    ("01_立项可研", "环境影响报告书", "",
     ["环境影响报告书", "环境影响报告表", "环境影响登记表"], ["批复", "验收"], 82),
    ("01_立项可研", "水资源论证报告及批复", "",
     ["水资源论证", "取水许可批复", "取水许可证"], [], 86),
    ("01_立项可研", "洪水影响评价批复", "",
     ["洪水影响评价", "防洪评价报告", "防洪影响评价"], [], 86),
    ("01_立项可研", "社会稳定风险评估报告", "",
     ["社会稳定风险评估", "社稳评估备案", "稳评报告"], [], 85),
    ("01_立项可研", "大坝安全鉴定报告书", "",
     ["安全鉴定报告书", "安全评价报告+水库", "除险加固安全鉴定"], [], 86),

    # ================= 02_用地规划 =================
    ("02_用地规划", "用地预审与选址意见书", "用地预审与选址意见书",
     ["用地预审与选址意见书", "建设项目选址意见书", "用地预审意见", "建设项目用地预审"], [], 96),
    ("02_用地规划", "建设用地规划许可证", "建设用地规划许可证",
     ["建设用地规划许可证", "建设用地批准书"], [], 96),
    ("02_用地规划", "建设工程规划许可证", "",
     ["建设工程规划许可证", "乡村建设规划许可证"], [], 93),
    ("02_用地规划", "征地移民安置批复", "征地移民安置批复",
     ["移民安置规划+批复", "征地补偿安置方案", "建设征地移民安置", "征收土地批复"], ["验收"], 94),
    ("02_用地规划", "使用林地审核同意书", "",
     ["使用林地审核同意书", "占用林地批复", "临时用地批复"], [], 88),
    ("02_用地规划", "不动产权证书", "",
     ["不动产权证", "国有土地使用证", "土地划拨决定书"], [], 90),

    # ================= 03_勘察设计 =================
    ("03_勘察设计", "初步设计及概算批复", "初步设计及概算批复",
     ["初步设计+批复", "初步设计报告+批复", "初设及概算批复", "概算核定+批复", "实施方案+批复"], [], 97),
    ("03_勘察设计", "初步设计报告", "",
     ["初步设计报告", "初步设计概算书"], ["批复", "招标"], 84),
    ("03_勘察设计", "工程地质勘察报告", "工程地质勘察报告",
     ["地质勘察报告", "岩土工程勘察报告", "水文地质勘察报告", "工程测量报告", "地形测绘报告"], ["招标", "合同"], 92),
    ("03_勘察设计", "施工图审查合格书", "施工图审查合格书",
     ["施工图设计文件审查合格书", "施工图审查合格书", "施工图审查意见", "图审合格证"], [], 95),
    ("03_勘察设计", "施工图设计文件", "",
     ["施工图设计", "全套施工图", "大坝枢纽施工图", "渠道施工图"], ["审查合格", "招标", "合同"], 85),

    # ================= 04_招采合同 =================
    ("04_招采合同", "中标通知书", "中标通知书",
     ["中标通知书", "成交通知书"], ["合同协议书"], 96),
    ("04_招采合同", "施工承包合同", "施工承包合同",
     ["施工合同", "施工总承包合同", "建设工程施工合同", "水利水电工程标准施工+合同协议书"], ["招标文件"], 95),
    ("04_招采合同", "监理合同", "监理合同",
     ["建设工程监理合同", "水利工程施工监理合同", "监理合同"], ["招标文件", "月报"], 95),
    ("04_招采合同", "勘察设计合同", "",
     ["勘察设计合同", "工程设计合同", "工程勘察合同", "全过程工程咨询合同"], ["招标文件"], 92),
    ("04_招采合同", "第三方检测合同", "",
     ["质量检测合同", "第三方检测合同", "水土保持监测合同", "安全监测合同"], [], 90),
    ("04_招采合同", "机电及金属结构采购合同", "",
     ["设备采购合同", "金属结构采购", "闸门启闭机采购合同", "水泵机组采购合同"], [], 90),
    ("04_招采合同", "施工招标文件", "施工招标文件",
     ["施工招标文件", "招标文件", "招标公告", "工程量清单及招标控制价", "最高投标限价"], ["评标报告"], 90),
    ("04_招采合同", "评标报告及中标候选人公示", "",
     ["评标报告", "中标候选人公示", "招投标情况书面报告"], [], 88),
    ("04_招采合同", "履约保函", "",
     ["履约保函", "农民工工资支付保函", "预付款保函"], [], 88),

    # ================= 05_施工许可 =================
    ("05_施工许可", "水利工程质量安全监督书", "水利工程质量安全监督手续",
     ["质量监督书", "质量与安全监督", "质量监督登记表", "安全监督备案", "质量安全监督手续"], ["质量监督报告"], 95),
    ("05_施工许可", "施工许可证", "施工许可证或开工批复",
     ["建筑工程施工许可证", "水利工程开工备案", "开工审批表", "准予开工批复", "开工条件备案"], [], 96),
    ("05_施工许可", "项目法人组建批复", "",
     ["项目法人组建", "成立项目建设管理处", "项目部组建通知"], [], 88),
    ("05_施工许可", "施工组织设计报审表", "",
     ["施工组织设计报审", "施工组织设计方案", "度汛方案批复", "导流度汛方案"], [], 86),
    ("05_施工许可", "监理规划及实施细则", "",
     ["监理规划", "监理实施细则"], [], 86),

    # ================= 06_过程管控 =================
    ("06_过程管控", "总监开工令", "总监开工令",
     ["工程开工令", "总监开工令", "开工通知+总监理工程师"], [], 96),
    ("06_过程管控", "设计变更通知及批复", "设计变更或工程签证",
     ["设计变更单", "设计变更通知单", "重大设计变更批复", "工程变更单"], [], 94),
    ("06_过程管控", "工程现场签证单", "设计变更或工程签证",
     ["工程签证单", "现场签证单", "工程量签证单", "索赔意向通知"], [], 94),
    ("06_过程管控", "进度款支付证书", "进度款支付证书",
     ["进度款支付证书", "工程款支付申请表", "进度付款申请单", "工程进度款审批表", "预付款支付证书"], [], 95),
    ("06_过程管控", "图纸会审与设计交底纪要", "",
     ["图纸会审纪要", "设计交底纪要", "技术交底记录"], [], 89),
    ("06_过程管控", "工程联系单", "",
     ["工程联系单", "工作联系单", "监理通知单", "监理通知回复单"], [], 88),
    ("06_过程管控", "监理月报及工地例会纪要", "",
     ["监理月报", "工地例会纪要", "监理日志汇总", "工程建设月报"], [], 87),
    ("06_过程管控", "工程质量检测报告", "",
     ["检测报告", "试验报告", "压实度检测", "混凝土抗压强度", "原材料进场复检"], ["检测合同"], 86),
    ("06_过程管控", "单元及分部工程质量评定表", "",
     ["单元工程质量评定", "分部工程质量评定", "隐蔽工程验收记录", "重要隐蔽单元工程"], ["验收鉴定书"], 87),
]


CN_NUM_MAP = {
    "〇": "0", "零": "0", "Ｏ": "0", "O": "0", "o": "0",
    "一": "1", "二": "2", "三": "3", "四": "4",
    "五": "5", "六": "6", "七": "7", "八": "8", "九": "9",
    "元": "1",
}


def _parse_cn_month_or_day(s: str) -> Optional[int]:
    s = s.strip()
    if not s:
        return None
    if s.isdigit():
        return int(s)
    if s == "十":
        return 10
    if s.startswith("十"):
        tail = CN_NUM_MAP.get(s[1:], "")
        return 10 + int(tail) if tail.isdigit() else None
    if s.endswith("十"):
        head = CN_NUM_MAP.get(s[:-1], "")
        return int(head) * 10 if head.isdigit() else None
    if "十" in s:
        parts = s.split("十", 1)
        h = CN_NUM_MAP.get(parts[0], "")
        t = CN_NUM_MAP.get(parts[1], "")
        if h.isdigit() and t.isdigit():
            return int(h) * 10 + int(t)
    digits = "".join(CN_NUM_MAP.get(ch, "") for ch in s)
    return int(digits) if digits.isdigit() else None


def _validate_ymd(y: int, m: int, d: int) -> Optional[str]:
    try:
        if not (1990 <= y <= 2099):
            return None
        dt_obj = dt.date(y, m, d)
        return dt_obj.strftime("%Y%m%d")
    except ValueError:
        return None


def _extract_dates_from_text(text: str) -> List[Tuple[str, int]]:
    """从文本中提取全部合法日期及其在文本中的偏移位置"""
    found: List[Tuple[str, int]] = []
    if not text:
        return found

    # 1. 阿拉伯数字日期: 2026年10月28日 / 2026-10-28 / 2026.10.28
    for m in re.finditer(r"(20\d{2}|19\d{2})\s*[年\-\./]\s*(\d{1,2})\s*[月\-\./]\s*(\d{1,2})\s*日?", text):
        ymd = _validate_ymd(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        if ymd:
            found.append((ymd, m.start()))

    # 2. 中文公文成文日期: 二〇二六年十月二十八日
    cn_pat = r"([一二〇零ＯOo九][〇零ＯOo九0-9一二三四五六七八]{3})\s*年\s*([一二三四五六七八九十元]{1,3})\s*月\s*([一二三四五六七八九十]{1,3})\s*日"
    for m in re.finditer(cn_pat, text):
        yr_str = "".join(CN_NUM_MAP.get(c, c) for c in m.group(1))
        mo = _parse_cn_month_or_day(m.group(2))
        da = _parse_cn_month_or_day(m.group(3))
        if yr_str.isdigit() and mo and da:
            ymd = _validate_ymd(int(yr_str), mo, da)
            if ymd:
                found.append((ymd, m.start()))

    return found


def extract_doc_date(original_filename: str, content: ExtractedContent) -> Tuple[str, str]:
    """
    按优先级提取业务日期：
    1. 正文显式标签日期（成文日期/签订日期/批复日期/签发日期/验收日期）
    2. 末页/尾部签名落款日期（红头批文与鉴定书标准落款位）
    3. 首页/正文其他合法日期
    4. 原文件名中的日期（若原文件名不是微信随机导出名则优先，或者是合法8位日期）
    5. 收件当天日期
    """
    combined_for_label = (content.head_text or "") + "\n" + (content.tail_text or "")
    label_pat = r"(?:签订日期|签署日期|成文日期|批复日期|批准日期|签发日期|发证日期|验收日期|审定日期|报告日期|日期)[\s|：:]*.{0,28}"
    for lm in re.finditer(label_pat, combined_for_label):
        snippet = lm.group(0)
        dates = _extract_dates_from_text(snippet)
        if dates:
            return dates[0][0], "正文成文/签署日期"

    # 末页最后 500 字落款日期（取最后一个合法日期，通常是盖章落款日）
    tail_zone = (content.tail_text or "")[-600:]
    tail_dates = _extract_dates_from_text(tail_zone)
    if tail_dates:
        return tail_dates[-1][0], "正文成文/签署日期"

    # 正文其他日期
    head_dates = _extract_dates_from_text(content.head_text or "")
    if head_dates:
        return head_dates[-1][0], "正文成文/签署日期"

    # 原文件名中的日期（支持 20261028 或 2026-10-28）
    stem = Path(original_filename).stem
    fn_dates = _extract_dates_from_text(stem)
    if fn_dates:
        return fn_dates[0][0], "原文件名日期"

    for m8 in re.finditer(r"(?<!\d)(20\d{2})(0[1-9]|1[0-2])(0[1-9]|[12]\d|3[01])(?!\d{7,})", stem):
        ymd = _validate_ymd(int(m8.group(1)), int(m8.group(2)), int(m8.group(3)))
        if ymd:
            return ymd, "原文件名日期"

    return dt.date.today().strftime("%Y%m%d"), "收件当天日期"


def _is_generic_junk_filename(stem: str) -> bool:
    """判断原文件名是否是毫无业务含义的系统/微信/扫描仪默认名"""
    s = stem.strip().lower()
    junk_prefixes = (
        "微信文件", "微信图片", "微信转发", "扫描转发", "转发",
        "wxfile", "mmexport", "img_", "dsc_", "scan", "扫描件",
        "新建", "未命名", "附件", "文档", "document",
    )
    if any(s.startswith(p) for p in junk_prefixes) and len(re.findall(r"[\u4e00-\u9fff]", s)) <= 5:
        return True
    if re.fullmatch(r"[0-9a-f\-_\(\)\s]+", s):
        return True
    return False


def match_project(original_filename: str, content: ExtractedContent, cfg: AppConfig) -> Tuple[str, bool, str]:
    """
    匹配在管项目白名单：
    返回 (project_name, is_pending, pending_reason)
    """
    stem = Path(original_filename).stem
    head_top = (content.head_text or "")[:600]
    full_txt = content.full_text or ""

    best_proj = ""
    best_score = 0.0

    for proj in cfg.projects:
        score = 0.0
        for term in proj.all_match_terms():
            if not term:
                continue
            weight_len = max(1.0, len(term) / 4.0)
            if term in stem:
                score += 6.0 * weight_len
            if term in head_top:
                score += 4.5 * weight_len
            elif term in full_txt:
                score += 1.8 * weight_len
        if score > best_score:
            best_score = score
            best_proj = proj.standard_name

    if best_proj and best_score > 0:
        return best_proj, False, ""

    # 未命中任何在管项目白名单 -> 尝试抽取疑似项目名作为重命名参考，标记放入 99_待人工确认
    candidate = _extract_candidate_unknown_project(stem, head_top)
    reason = f"未匹配《项目配置.yaml》白名单（识别到疑似工程：{candidate}）" if candidate != "未匹配项目" else "未匹配《项目配置.yaml》中的已知项目关键词"
    return candidate, True, reason


def _extract_candidate_unknown_project(stem: str, head_top: str) -> str:
    combined = f"{stem}\n{head_top}"
    patterns = [
        r"关于[《\s]*([\u4e00-\u9fff0-9]{3,20}?(?:水库|灌区|河道|堤防|水闸|泵站|引水|供水|除险加固|续建配套|治理)[\u4e00-\u9fff0-9]{0,10}?工程)",
        r"工程名称\s*[：:]\s*([\u4e00-\u9fff0-9]{4,24}?工程)",
        r"([\u4e00-\u9fff]{2,14}?(?:水库|灌区|河道|堤防|水闸|泵站)[\u4e00-\u9fff]{0,10}?工程)",
    ]
    for pat in patterns:
        m = re.search(pat, combined)
        if m:
            cand = sanitize_windows_name(m.group(1)).replace("-", "")
            if 4 <= len(cand) <= 24:
                return cand
    return "未匹配项目"


def _normalize_cjk_spaces(text: str) -> str:
    """消除公文标题常见的汉字排版字间距空格（如 '中 标 通 知 书' -> '中标通知书'）"""
    if not text:
        return ""
    return re.sub(r"(?<=[\u4e00-\u9fff])[ \t\u3000]+(?=[\u4e00-\u9fff])", "", text)


def _pattern_matches(expr: str, text: str) -> bool:
    parts = [p.strip() for p in expr.split("+") if p.strip()]
    if len(parts) == 1:
        return parts[0] in text
    # 合并 PDF 长标题自动折行（保留句号、分号、感叹号作为语句边界），且要求多词在同一语句中相距不超过 60 字
    compact = re.sub(r"[\r\n]+", " ", text)
    for clause in re.split(r"[。；！？]+", compact):
        if all(p in clause for p in parts):
            positions = [clause.find(p) for p in parts]
            if max(positions) - min(positions) <= 60:
                return True
    return False


def extract_qualifier(title_zone: str) -> str:
    """提取标段、专业、期次、部位等限定特征词"""
    qualifiers: List[str] = []

    # 1. 标段 / 专业标
    bid_patterns = [
        r"(施工[一二三四五六七八九十0-9]+标段?)",
        r"(监理[一二三四0-9]*标段?)",
        r"(勘察设计[一二0-9]*标段?)",
        r"(金属结构[一二0-9]*标段?)",
        r"(机电设备[一二0-9]*标段?)",
        r"(信息化[一二0-9]*标段?)",
        r"(第?[一二三四五六七八九十0-9]+标段)",
    ]
    for pat in bid_patterns:
        m = re.search(pat, title_zone)
        if m:
            q = m.group(1).rstrip("段") if m.group(1).endswith("标段") and "施工" in m.group(1) else m.group(1)
            qualifiers.append(q)
            break

    # 2. 期次 / 编号（如 第3期、第2号）
    m_seq = re.search(r"第\s*([0-9一二三四五六七八九十]{1,3})\s*([期批号次])", title_zone)
    if m_seq:
        raw_num = m_seq.group(1)
        unit = m_seq.group(2)
        parsed = _parse_cn_month_or_day(raw_num)
        num_str = str(parsed) if parsed is not None else raw_num
        qualifiers.append(f"第{num_str}{unit}")

    # 3. 水工建筑物关键部位（尤其对分部工程验收、变更、签证区分度极高）
    part_m = re.search(
        r"(主坝|副坝|大坝枢纽|溢洪道|输水隧洞|输水洞|放水洞|泄洪闸|进水闸|分水闸|节制闸|渡槽|倒虹吸|总干渠|干渠|支渠|泵房)",
        title_zone,
    )
    if part_m and len(qualifiers) < 2:
        qualifiers.append(part_m.group(1))

    return "".join(qualifiers)


def classify_stage_and_type(
    original_filename: str,
    ext: str,
    content: ExtractedContent,
    max_type_len: int = 34,
) -> Tuple[str, str, str, float]:
    """
    判定 (stage, doc_type, milestone_key, confidence)
    标题封面区权重 5.0x，全文区权重 1.0x，彻底避免长合同内文条款干扰主分类
    """
    stem = Path(original_filename).stem
    meaningful_stem = "" if _is_generic_junk_filename(stem) else _normalize_cjk_spaces(stem)
    norm_head = _normalize_cjk_spaces(content.head_text or "")
    title_zone = f"{meaningful_stem}\n{norm_head[:360]}"
    body_zone = _normalize_cjk_spaces(content.full_text or "")

    best_rule = None
    best_score = 0.0

    for stage, doc_type, milestone, pos_exprs, neg_words, base_score in RAW_DOC_RULES:
        # 排斥词若出现在标题区则直接排除
        if any(nw in title_zone for nw in neg_words):
            continue

        in_title = any(_pattern_matches(expr, title_zone) for expr in pos_exprs)
        in_body = any(_pattern_matches(expr, body_zone) for expr in pos_exprs)

        if not in_title and not in_body:
            continue

        # 正文匹配时，检查正文是否受排斥词压制
        if not in_title and any(nw in body_zone[:500] for nw in neg_words):
            continue

        score = base_score * (5.0 if in_title else 1.0)
        if meaningful_stem and any(_pattern_matches(expr, meaningful_stem) for expr in pos_exprs):
            score += 60.0

        if score > best_score:
            best_score = score
            best_rule = (stage, doc_type, milestone)

    if best_rule is None:
        # 宽泛阶段词兜底
        broad_stage_keywords = [
            ("01_立项可研", "立项前期文件", ["立项", "可研", "可行性", "水保", "环评", "论证"]),
            ("02_用地规划", "用地规划文件", ["用地", "选址", "规划许可", "征地", "移民", "林地"]),
            ("03_勘察设计", "勘察设计文件", ["勘察", "地勘", "初设", "初步设计", "概算", "施工图", "图纸"]),
            ("04_招采合同", "招采合同文件", ["招标", "投标", "中标", "合同", "协议书", "工程量清单", "保函"]),
            ("05_施工许可", "施工准备报建文件", ["施工许可", "开工备案", "质量监督", "安全监督", "施工组织"]),
            ("06_过程管控", "施工过程管控文件", ["开工令", "变更", "签证", "进度款", "联系单", "监理", "检测", "评定"]),
            ("07_竣工验收", "工程验收文件", ["验收", "完工", "竣工"]),
            ("08_结算归档", "结算决算归档文件", ["结算", "决算", "审计", "质保金", "移交"]),
        ]
        combined = f"{meaningful_stem}\n{body_zone}"
        for stg, fallback_type, kws in broad_stage_keywords:
            if any(kw in combined for kw in kws):
                best_rule = (stg, meaningful_stem[:20] if meaningful_stem else fallback_type, "")
                best_score = 45.0
                break

    if best_rule is None:
        fallback_doc = meaningful_stem[:24] if meaningful_stem else "待分类工程文件"
        return "99_待人工确认", sanitize_doc_type(fallback_doc, max_type_len), "", 0.15

    stage, base_doc_type, milestone = best_rule
    qualifier = extract_qualifier(title_zone)
    final_type = base_doc_type
    if qualifier and qualifier not in final_type:
        final_type = f"{qualifier}{final_type}"
    if ext.lower() == ".zip" and not final_type.endswith(("整包", "汇总", "压缩包")):
        final_type = f"{final_type}整包"

    confidence = min(0.99, round(best_score / 500.0, 2)) if best_score >= 100 else 0.65
    return stage, sanitize_doc_type(final_type, max_type_len), milestone, confidence


def sanitize_doc_type(doc_type: str, max_len: int = 34) -> str:
    s = sanitize_windows_name(doc_type)
    # 去除内部横杠，保证“年月日-项目名称-类型”严格三段式结构
    s = s.replace("-", "_").replace(" ", "")
    s = re.sub(r"_+", "_", s).strip("_")
    return (s[:max_len] if len(s) > max_len else s) or "工程文件"


def extract_doc_no(content: ExtractedContent) -> str:
    """提取红头批复发文字号或工程合同编号"""
    text = f"{content.head_text or ''}\n{(content.tail_text or '')[:500]}"
    if not text.strip():
        return ""

    # 1. 政府公文标准发文字号：如 江水许〔2026〕48号、东发改农经[2026]15号
    gov_pat = r"([\u4e00-\u9fff]{2,12}[〔\[【（(]\s*20\d{2}\s*[〕\]】）)]\s*\d{1,4}\s*号)"
    m = re.search(gov_pat, text)
    if m:
        val = re.sub(r"\s+", "", m.group(1))
        val = re.sub(r"[\[【（(](20\d{2})[\]】）)]", r"〔\1〕", val)
        return val

    # 2. 许可证号 / 合同编号 / 证书编号（支持 Word/Excel 表格竖线分隔）
    code_pat = r"(?:合同编号|证书编号|许可证号|报告编号|文件编号|支付证书号|编号)[\s|：:]+([A-Za-z0-9\-\u4e00-\u9fff〔〕\[\]]{4,32})"
    m2 = re.search(code_pat, text)
    if m2:
        cand = m2.group(1).strip().rstrip("。，；")
        if re.search(r"[0-9]", cand):
            return cand

    return ""


def extract_amount_wan(content: ExtractedContent) -> str:
    """提取核心工程金额并统一换算/格式化为“万元”（保留2位小数）"""
    text = f"{content.head_text or ''}\n{content.tail_text or ''}"
    if not text.strip():
        return ""

    # 1. 优先匹配带明确业务前缀的“万元”
    wan_pat = r"(?:概算总投资|核定总投资|工程总投资|总投资|签约合同价|合同总价|合同金额|中标价|审定结算价|审定金额|决算总投资|本期支付金额|本次申请支付|支付金额)[^\d\n]{0,22}?([0-9]{1,7}(?:,[0-9]{3})*(?:\.[0-9]{1,4})?)\s*万元"
    m = re.search(wan_pat, text)
    if m:
        num = float(m.group(1).replace(",", ""))
        return f"{num:.2f}"

    # 2. 匹配带明确业务前缀的“元”（支持千分位逗号如 45,200,000.00元 或纯数字 45200000元，自动除以 10000 转为万元）
    yuan_pat = r"(?:签约合同价|合同总价|合同金额|中标价|审定结算价|审定金额|本期支付金额|本次申请支付|小写)[\s|：:人民币（()）小写￥¥]*((?:[0-9]{1,3}(?:,[0-9]{3})+|[0-9]{4,12})(?:\.[0-9]{1,2})?)\s*元"
    m2 = re.search(yuan_pat, text)
    if m2:
        num_yuan = float(m2.group(1).replace(",", ""))
        return f"{num_yuan / 10000.0:.2f}"

    # 3. 通用兜底“XX万元”
    m3 = re.search(r"([0-9]{2,6}(?:\.[0-9]{1,2})?)\s*万元", text[:1200])
    if m3:
        return f"{float(m3.group(1)):.2f}"

    return ""


def _enhance_with_ai_if_configured(
    local_res: ClassificationResult,
    original_filename: str,
    content: ExtractedContent,
    cfg: AppConfig,
) -> ClassificationResult:
    ai = cfg.ai_llm
    if not (ai.enabled and ai.api_key):
        return local_res
    if ai.mode == "fallback" and (not local_res.is_pending and local_res.confidence >= 0.75 and local_res.doc_no):
        return local_res

    try:
        import requests

        proj_names = [p.standard_name for p in cfg.projects]
        prompt = (
            "你是中国水利工程档案台账专家。请根据文件名与提取文本返回严格 JSON：\n"
            f"已知在管项目列表: {json.dumps(proj_names, ensure_ascii=False)}\n"
            f"可选阶段列表: {json.dumps(STAGES_ORDER, ensure_ascii=False)}\n"
            "返回字段：{\"project_name\": \"匹配的已知项目名或空字符串\", \"stage\": \"01_立项可研..08_结算归档\", "
            "\"doc_type\": \"含标段或期次的简明具体文件类型\", \"doc_date\": \"YYYYMMDD或空\", "
            "\"doc_no\": \"发文字号或合同编号\", \"amount_wan\": \"万元纯数字字符串或空\"}\n\n"
            f"原文件名：{original_filename}\n"
            f"文档首部文本：\n{(content.head_text or '')[:1800]}\n"
            f"文档尾部文本：\n{(content.tail_text or '')[-800:]}"
        )
        resp = requests.post(
            f"{ai.api_base}/chat/completions",
            headers={"Authorization": f"Bearer {ai.api_key}", "Content-Type": "application/json"},
            json={
                "model": ai.model,
                "messages": [{"role": "user", "content": prompt}],
                "temperature": 0.1,
                "response_format": {"type": "json_object"},
            },
            timeout=ai.timeout_seconds,
        )
        if resp.status_code == 200:
            raw_content = resp.json()["choices"][0]["message"]["content"]
            data = json.loads(raw_content)
            if data.get("project_name") in proj_names:
                local_res.project_name = data["project_name"]
                local_res.is_pending = False
                local_res.pending_reason = ""
            if data.get("stage") in STAGES_ORDER:
                local_res.stage = data["stage"]
            if data.get("doc_type"):
                local_res.doc_type = sanitize_doc_type(data["doc_type"], cfg.max_doc_type_length)
                local_res.milestone_key = infer_milestone_from_type(local_res.stage, local_res.doc_type)
            if re.fullmatch(r"20\d{6}", str(data.get("doc_date", ""))):
                local_res.doc_date = str(data["doc_date"])
            if data.get("doc_no") and not local_res.doc_no:
                local_res.doc_no = str(data["doc_no"]).strip()
            if data.get("amount_wan") and not local_res.amount_wan:
                local_res.amount_wan = str(data["amount_wan"]).strip()
    except Exception:
        pass

    return local_res


def infer_milestone_from_type(stage: str, doc_type: str) -> str:
    """当人工在资源管理器改名或 AI 返回类型时，根据阶段与文件类型反推其点亮的核心要件"""
    for r_stage, r_type, r_milestone, pos_exprs, neg_words, _ in RAW_DOC_RULES:
        if r_stage != stage or not r_milestone:
            continue
        if any(nw in doc_type for nw in neg_words):
            continue
        if r_type in doc_type or any(_pattern_matches(expr, doc_type) for expr in pos_exprs):
            return r_milestone
    # 直接与默认要件名称包含匹配
    for ms in DEFAULT_STAGE_MILESTONES.get(stage, []):
        if ms in doc_type or doc_type in ms:
            return ms
    return ""


def parse_standard_filename(filename: str) -> Tuple[Optional[str], Optional[str], Optional[str]]:
    """
    解析符合 `YYYYMMDD-项目名称-文件类型(-02).ext` 的规范文件名
    返回 (doc_date, project_name, doc_type)
    """
    stem = Path(filename).stem
    # 去除末尾防重名序号如 -02, -03
    stem_clean = re.sub(r"-0?\d{1,2}$", "", stem)
    parts = stem_clean.split("-", 2)
    if len(parts) == 3 and re.fullmatch(r"(19|20)\d{6}", parts[0]):
        return parts[0], parts[1].strip(), parts[2].strip()
    return None, None, None


def classify_document(original_filename: str, content: ExtractedContent, cfg: AppConfig) -> ClassificationResult:
    ext = Path(original_filename).suffix.lower()
    proj_name, is_pending_proj, pending_reason = match_project(original_filename, content, cfg)
    stage, doc_type, milestone_key, confidence = classify_stage_and_type(
        original_filename, ext, content, cfg.max_doc_type_length
    )
    doc_date, date_source = extract_doc_date(original_filename, content)
    doc_no = extract_doc_no(content)
    amount_wan = extract_amount_wan(content)

    is_pending = is_pending_proj or (stage == "99_待人工确认")
    if stage == "99_待人工确认" and not pending_reason:
        pending_reason = "文件内容缺乏水利阶段特征词，需人工拖入对应阶段文件夹"

    res = ClassificationResult(
        project_name=proj_name,
        is_pending=is_pending,
        pending_reason=pending_reason,
        stage=stage,
        doc_type=doc_type,
        milestone_key=milestone_key if not is_pending else "",
        doc_date=doc_date,
        date_source=date_source,
        doc_no=doc_no,
        amount_wan=amount_wan,
        confidence=confidence,
        standard_filename="",
    )

    res = _enhance_with_ai_if_configured(res, original_filename, content, cfg)
    if not res.is_pending and res.stage in STAGES_ORDER and not res.milestone_key:
        res.milestone_key = infer_milestone_from_type(res.stage, res.doc_type)

    res.standard_filename = f"{res.doc_date}-{res.project_name}-{res.doc_type}{ext}"
    return res

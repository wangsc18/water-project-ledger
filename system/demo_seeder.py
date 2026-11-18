# -*- coding: utf-8 -*-
"""
水利工程仿真多格式测试样板数据生成器
向 00_待归档收件箱 投递命名杂乱的真实感水利工程文件（含纯图像无文字层扫描件PDF、多页公文PDF、Word合同、Excel进度款表、ZIP图纸整包及白名单外待确认文件），
验证全自动 OCR 识别、三段式规范重命名、标段/期次提取、双层 Excel 超链接与可视化看板点亮效果。
"""
from __future__ import annotations

import zipfile
from pathlib import Path

import fitz  # PyMuPDF
import openpyxl
from docx import Document


def _make_text_pdf(path: Path, pages_html: list[str]) -> None:
    doc = fitz.open()
    for html in pages_html:
        page = doc.new_page(width=595, height=842)
        page.insert_htmlbox(fitz.Rect(52, 56, 543, 785), html)
    doc.save(str(path))
    doc.close()


def _make_scanned_image_pdf(path: Path, html: str) -> None:
    """生成真正无内嵌文字层的纯图片扫描件 PDF，强制触发离线中文 OCR 引擎"""
    tmp_doc = fitz.open()
    p = tmp_doc.new_page(width=595, height=842)
    p.insert_htmlbox(fitz.Rect(48, 52, 547, 780), html)
    pix = p.get_pixmap(dpi=200, alpha=False)
    img_bytes = pix.tobytes("png")
    tmp_doc.close()

    scan_doc = fitz.open()
    sp = scan_doc.new_page(width=595, height=842)
    sp.insert_image(fitz.Rect(0, 0, 595, 842), stream=img_bytes)
    scan_doc.save(str(path))
    scan_doc.close()


def _make_contract_docx(
    path: Path,
    title: str,
    proj_name: str,
    bid_section: str,
    contract_no: str,
    amount_yuan_str: str,
    sign_date: str,
) -> None:
    doc = Document()
    doc.add_heading(title, level=1)
    tbl = doc.add_table(rows=5, cols=2)
    rows_data = [
        ("工程名称", f"{proj_name}{bid_section}"),
        ("合同编号", contract_no),
        ("发包人", "江州市水利建设投资开发有限公司"),
        ("签约合同价", f"人民币（小写）：{amount_yuan_str}元"),
        ("签订日期", sign_date),
    ]
    for idx, (k, v) in enumerate(rows_data):
        tbl.rows[idx].cells[0].text = k
        tbl.rows[idx].cells[1].text = v

    doc.add_paragraph(
        "通用合同条款摘要：承包人应严格按照初步设计及施工图组织施工，配合发包人办理进度款支付、设计变更、合同工程完工验收、竣工验收及最终工程结算审核。"
    )
    doc.add_paragraph(f"签署日期：{sign_date}")
    doc.save(str(path))


def _make_payment_xlsx(path: Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "工程进度款支付证书"
    rows = [
        ["工程进度款支付证书（第2期）", "", "", ""],
        ["工程名称", "清河水库除险加固工程（施工一标）", "支付证书号", "QHSK-ZF-2026-02"],
        ["承包单位", "省水利水电第二工程局有限公司", "监理单位", "江州禹顺水利工程监理有限公司"],
        ["本期完成工程量合计", "756.10万元", "本期支付金额", "680.50万元"],
        ["总监理工程师签字", "张明远", "签发日期", "2026年11月16日"],
    ]
    for r in rows:
        ws.append(r)
    wb.save(str(path))
    wb.close()


def _make_design_zip(path: Path) -> None:
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("01_清河水库除险加固工程_大坝枢纽全套施工图设计说明.txt", "清河水库主坝防渗加固与溢洪道施工图设计说明".encode("utf-8"))
        zf.writestr("02_主坝帷幕灌浆平面布置图.dwg", b"AC1027_DUMMY_DWG_CONTENT")
        zf.writestr("03_溢洪道控制段配筋图.dwg", b"AC1027_DUMMY_DWG_CONTENT")


def seed_demo_files_to_inbox(inbox_dir: Path) -> int:
    """生成 18 份典型杂乱命名的多格式水利工程收文投入收件箱"""
    inbox_dir.mkdir(parents=True, exist_ok=True)
    font_css = "font-family: sans-serif; line-height: 1.65;"

    # ---------------- 项目 1：清河水库除险加固工程（推进至 06_过程管控，多个前置阶段亮绿灯） ----------------
    # 01_立项可研 (4/4 齐备 -> 绿灯)
    _make_text_pdf(
        inbox_dir / "发改委收文_001.pdf",
        [f'<div style="{font_css}"><h2 style="color:#c00;text-align:center">江州市发展和改革委员会文件</h2>'
         '<p style="text-align:center">江发改农经〔2026〕31号</p>'
         '<h3 style="text-align:center">关于清河水库除险加固工程项目建议书的批复</h3>'
         '<p>同意清河水库除险加固工程立项建设，估算总投资 8900.00 万元。</p>'
         '<p style="text-align:right;margin-top:120px">2026年06月12日</p></div>'],
    )
    _make_text_pdf(
        inbox_dir / "扫描转发(2).pdf",
        [
            f'<div style="{font_css}"><h2 style="color:#c00;text-align:center">江州市发展和改革委员会文件</h2>'
            '<p style="text-align:center">江发改农经〔2026〕89号</p>'
            '<h3 style="text-align:center">关于清河水库除险加固工程可行性研究报告的批复</h3>'
            '<p>经研究，原则同意《清河水库除险加固工程可行性研究报告》，估算总投资 8780.00 万元。</p></div>',
            f'<div style="{font_css}"><p>请据此抓紧开展初步设计工作。</p>'
            '<p style="text-align:right;margin-top:220px">江州市发展和改革委员会<br>二〇二六年七月二十日</p></div>',
        ],
    )
    _make_text_pdf(
        inbox_dir / "水保许可决定书.pdf",
        [f'<div style="{font_css}"><h2 style="color:#c00;text-align:center">江州市水利局行政许可决定书</h2>'
         '<p style="text-align:center">江水保许〔2026〕19号</p>'
         '<h3 style="text-align:center">关于清河水库除险加固工程水土保持方案的批复</h3>'
         '<p>准予水土保持方案行政许可。</p><p style="text-align:right;margin-top:120px">2026年08月05日</p></div>'],
    )
    _make_text_pdf(
        inbox_dir / "生态环境局批文_scan.pdf",
        [f'<div style="{font_css}"><h2 style="color:#c00;text-align:center">江州市生态环境局文件</h2>'
         '<p style="text-align:center">江环审〔2026〕44号</p>'
         '<h3 style="text-align:center">关于清河水库除险加固工程环境影响报告表的批复</h3>'
         '<p>同意该项目环境影响评价结论。</p><p style="text-align:right;margin-top:120px">2026年08月18日</p></div>'],
    )

    # 02_用地规划 (2/3 进行中 -> 蓝灯，缺“征地移民安置批复”展示缺件提醒效果)
    _make_text_pdf(
        inbox_dir / "选址意见.pdf",
        [f'<div style="{font_css}"><h3 style="text-align:center">建设项目用地预审与选址意见书</h3>'
         '<p>证书编号：用字第3601202600028号</p><p>项目名称：清河水库除险加固工程</p>'
         '<p style="text-align:right;margin-top:120px">发证日期：2026年08月29日</p></div>'],
    )
    _make_text_pdf(
        inbox_dir / "用地规划证_副本.pdf",
        [f'<div style="{font_css}"><h3 style="text-align:center">中华人民共和国建设用地规划许可证</h3>'
         '<p>许可证号：地字第3601202600051号</p><p>项目名称：清河水库除险加固工程</p>'
         '<p style="text-align:right;margin-top:120px">发证日期：2026年09月10日</p></div>'],
    )

    # 03_勘察设计 (3/3 齐备 -> 绿灯，含 1 份无文字层纯图片扫描件 PDF + 1 份 ZIP 整包图纸)
    _make_text_pdf(
        inbox_dir / "地勘院最终版.pdf",
        [f'<div style="{font_css}"><h2 style="text-align:center">清河水库除险加固工程地质勘察报告</h2>'
         '<p>报告编号：JZK-2026-091</p><p>提交单位：省水利水电勘测设计研究院</p>'
         '<p style="text-align:right;margin-top:120px">报告日期：2026年09月18日</p></div>'],
    )
    # 重点：纯图片扫描件 PDF，测试离线中文 OCR
    _make_scanned_image_pdf(
        inbox_dir / "微信文件_20261118102345.pdf",
        f'<div style="{font_css}"><h2 style="color:#d00;text-align:center">江州市水利局文件</h2>'
        '<p style="text-align:center">江水许〔2026〕48号</p>'
        '<h3 style="text-align:center">关于清河水库除险加固工程初步设计报告的批复</h3>'
        '<p>市水利建设投资公司：你单位报来的《清河水库除险加固工程初步设计报告》收悉。经审查，核定工程概算总投资 8650.40 万元。</p>'
        '<p style="text-align:right;margin-top:140px">2026年09月28日</p></div>',
    )
    _make_text_pdf(
        inbox_dir / "图审合格证0930.pdf",
        [f'<div style="{font_css}"><h3 style="text-align:center">水利工程施工图设计文件审查合格书</h3>'
         '<p>证书编号：江水图审〔2026〕16号</p><p>工程名称：清河水库除险加固工程</p>'
         '<p style="text-align:right;margin-top:120px">签发日期：2026年10月08日</p></div>'],
    )
    _make_design_zip(inbox_dir / "设计院发来图纸汇总包.zip")

    # 04_招采合同 (4/4 齐备 -> 绿灯，含多标段 Word 合同)
    _make_text_pdf(
        inbox_dir / "招标代理定稿_v3.pdf",
        [f'<div style="{font_css}"><h2 style="text-align:center">清河水库除险加固工程施工一标段施工招标文件</h2>'
         '<p>文件编号：QHSK-ZB-2026-01</p><p>招标控制价：4680.00万元</p>'
         '<p style="text-align:right;margin-top:120px">日期：2026年10月12日</p></div>'],
    )
    _make_text_pdf(
        inbox_dir / "中标通知_盖章版.pdf",
        [f'<div style="{font_css}"><h2 style="text-align:center">中 标 通 知 书</h2>'
         '<p>编号：江公易建〔2026〕208号</p>'
         '<p>省水利水电第二工程局有限公司：你方于清河水库除险加固工程施工一标段投标中标，中标价 4520.00 万元。</p>'
         '<p style="text-align:right;margin-top:120px">签发日期：2026年11月02日</p></div>'],
    )
    _make_contract_docx(
        inbox_dir / "新建 Microsoft Word 文档 (2).docx",
        title="水利水电工程标准施工承包合同协议书",
        proj_name="清河水库除险加固工程",
        bid_section="施工一标段",
        contract_no="QHSK-SG-2026-01",
        amount_yuan_str="45,200,000.00",
        sign_date="2026年11月06日",
    )
    _make_contract_docx(
        inbox_dir / "监理合同正本扫描前电子稿.docx",
        title="水利工程施工监理合同",
        proj_name="清河水库除险加固工程",
        bid_section="监理标",
        contract_no="QHSK-JL-2026-01",
        amount_yuan_str="1,180,000.00",
        sign_date="2026年11月07日",
    )

    # 05_施工许可 & 06_过程管控
    _make_text_pdf(
        inbox_dir / "质监书.pdf",
        [f'<div style="{font_css}"><h2 style="text-align:center">水利工程质量与安全监督书</h2>'
         '<p>编号：江水质监〔2026〕22号</p><p>工程名称：清河水库除险加固工程</p>'
         '<p style="text-align:right;margin-top:120px">签发日期：2026年11月09日</p></div>'],
    )
    _make_text_pdf(
        inbox_dir / "开工备案批复.pdf",
        [f'<div style="{font_css}"><h2 style="text-align:center">水利工程准予开工批复（施工许可证）</h2>'
         '<p>编号：江水开工备〔2026〕15号</p><p>工程名称：清河水库除险加固工程</p>'
         '<p style="text-align:right;margin-top:120px">批准日期：2026年11月10日</p></div>'],
    )
    _make_text_pdf(
        inbox_dir / "总监1号令.pdf",
        [f'<div style="{font_css}"><h2 style="text-align:center">工程开工令（总监开工令）</h2>'
         '<p>编号：QHSK-JL-KGL-01</p><p>工程名称：清河水库除险加固工程施工一标段</p>'
         '<p>兹指令你单位于 2026年11月12日 正式破土开工。</p>'
         '<p style="text-align:right;margin-top:120px">总监理工程师签发日期：2026年11月11日</p></div>'],
    )
    _make_payment_xlsx(inbox_dir / "附件3_报审表_final.xlsx")

    # ---------------- 项目 2：东渠灌区续建配套工程（前期推进至 03_勘察设计） ----------------
    _make_text_pdf(
        inbox_dir / "东渠立项可研批文.pdf",
        [f'<div style="{font_css}"><h2 style="color:#c00;text-align:center">江州市发展和改革委员会文件</h2>'
         '<p style="text-align:center">江发改农经〔2026〕105号</p>'
         '<h3 style="text-align:center">关于东渠灌区续建配套与节水改造工程可行性研究报告的批复</h3>'
         '<p>核定可研估算总投资 12600.00 万元。</p>'
         '<p style="text-align:right;margin-top:120px">2026年09月05日</p></div>'],
    )
    _make_text_pdf(
        inbox_dir / "东渠初设批复.pdf",
        [f'<div style="{font_css}"><h2 style="color:#c00;text-align:center">江州市水利局文件</h2>'
         '<p style="text-align:center">江水许〔2026〕63号</p>'
         '<h3 style="text-align:center">关于东渠灌区续建配套工程总干渠初步设计报告的批复</h3>'
         '<p>核定工程概算总投资 12180.00 万元。</p>'
         '<p style="text-align:right;margin-top:120px">2026年11月08日</p></div>'],
    )

    # ---------------- 待人工确认样本（故意不在白名单内的陌生工程） ----------------
    _make_text_pdf(
        inbox_dir / "微信转发_陌生工程批文.pdf",
        [f'<div style="{font_css}"><h2 style="color:#c00;text-align:center">江州市水利局文件</h2>'
         '<p style="text-align:center">江水保许〔2026〕35号</p>'
         '<h3 style="text-align:center">关于西山河道生态治理工程水土保持方案的批复</h3>'
         '<p>同意西山河道生态治理工程水土保持方案。</p>'
         '<p style="text-align:right;margin-top:120px">2026年11月14日</p></div>'],
    )

    return 21

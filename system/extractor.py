# -*- coding: utf-8 -*-
"""
多格式工程文件内容提取与离线中文 OCR 识别模块
支持：
- PDF（电子文字层首尾页提取 + 无字盖章扫描件渲染高清图走离线中文 OCR）
- Word（.docx 段落与封面表格提取 + Windows Word/WPS COM 及纯 Python 兜底提取旧版 .doc）
- Excel（.xlsx / .xlsm 各工作表内容提取）
- 图片（.jpg / .jpeg / .png / .bmp / .tif 直接离线中文 OCR）
- 压缩包（.zip 不解压只读内部目录文件名，兼容 Windows 中文 GBK 编码）
- 图纸及其他格式（.dwg / .rar / .7z 等基于原文件名识别）
"""
from __future__ import annotations

import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import List, Optional, Tuple

IGNORED_NAMES = {"thumbs.db", "desktop.ini", ".ds_store"}
TEMP_EXTENSIONS = {".tmp", ".part", ".crdownload", ".download", ".partial", ".lock"}

_RAPID_OCR_ENGINE = None
_RAPID_OCR_CHECKED = False


@dataclass
class ExtractedContent:
    head_text: str = ""
    tail_text: str = ""
    full_text: str = ""
    ocr_used: bool = False
    extract_method: str = "文件名识别"


def should_ignore_file(path: Path) -> bool:
    name = path.name
    if not name:
        return True
    lower = name.lower()
    if lower in IGNORED_NAMES:
        return True
    if name.startswith((".", "_", "~")):
        return True
    if path.suffix.lower() in TEMP_EXTENSIONS:
        return True
    return False


def is_file_ready(path: Path, check_delay: float = 0.6) -> bool:
    """
    检测文件是否已由 Windows 资源管理器 / 微信电脑版彻底写入完成：
    1. 间隔 check_delay 秒检查文件大小与修改时间稳定
    2. 尝试以读写模式打开探测是否仍被写入进程独占锁定
    """
    try:
        if not path.exists() or not path.is_file():
            return False
        st1 = path.stat()
        time.sleep(check_delay)
        if not path.exists():
            return False
        st2 = path.stat()
        if st1.st_size != st2.st_size or st1.st_mtime_ns != st2.st_mtime_ns:
            return False
        # 探测独占写入锁（Windows 下正在拷贝的大文件无法以 r+b 打开）
        try:
            with open(path, "r+b"):
                pass
        except PermissionError:
            return False
        except OSError:
            # 只读属性文件可能拒绝 r+b，改试 rb
            with open(path, "rb") as f:
                f.read(16)
        return True
    except OSError:
        return False


def clean_ocr_text(text: str) -> str:
    """清洗 OCR 结果：去除汉字之间的多余空格，规范化公文六角括号与标点"""
    if not text:
        return ""
    # 合并连续空白（保留换行）
    lines = []
    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line:
            continue
        # 去除中文字符与中文字符/数字/括号之间的识别碎空格
        line = re.sub(r"(?<=[\u4e00-\u9fff])\s+(?=[\u4e00-\u9fff])", "", line)
        line = re.sub(r"(?<=[\u4e00-\u9fff])\s+(?=[0-9〔\[【（(])", "", line)
        line = re.sub(r"(?<=[0-9〕\]】）)])\s+(?=[\u4e00-\u9fff])", "", line)
        line = re.sub(r"\[\s*(\d{4})\s*\]", r"〔\1〕", line)
        lines.append(line)
    return "\n".join(lines)


def _get_rapid_ocr():
    global _RAPID_OCR_ENGINE, _RAPID_OCR_CHECKED
    if _RAPID_OCR_CHECKED:
        return _RAPID_OCR_ENGINE
    _RAPID_OCR_CHECKED = True
    try:
        from rapidocr_onnxruntime import RapidOCR  # type: ignore

        _RAPID_OCR_ENGINE = RapidOCR()
    except Exception:
        _RAPID_OCR_ENGINE = None
    return _RAPID_OCR_ENGINE


def ocr_image_bytes(img_bytes: bytes, scratch_dir: Path) -> str:
    """
    多级离线中文 OCR：
    1. RapidOCR (Windows 打包绿色版与 pip 环境首选：直接传 bytes 彻底规避 Windows 中文路径问题)
    2. Windows 10/11 原生内置 Windows.Media.Ocr (零依赖后备)
    3. Tesseract chi_sim+eng (本地开发与跨平台后备)
    """
    if not img_bytes:
        return ""

    # 1. RapidOCR 优先
    engine = _get_rapid_ocr()
    if engine is not None:
        try:
            result, _ = engine(img_bytes)
            if result:
                texts = [str(item[1]) for item in result if len(item) >= 2 and item[1]]
                return clean_ocr_text("\n".join(texts))
        except Exception:
            pass

    scratch_dir.mkdir(parents=True, exist_ok=True)
    tmp_img = scratch_dir / f"ocr_{os.getpid()}_{int(time.time() * 1000)}.png"
    try:
        tmp_img.write_bytes(img_bytes)

        # 2. Windows 10/11 原生 OCR 后备
        if sys.platform == "win32":
            text_win = _run_windows_native_ocr(tmp_img)
            if text_win.strip():
                return clean_ocr_text(text_win)

        # 3. Tesseract 后备
        tess_bin = shutil.which("tesseract")
        if tess_bin:
            proc = subprocess.run(
                [tess_bin, str(tmp_img), "stdout", "-l", "chi_sim+eng"],
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="ignore",
                timeout=30,
            )
            if proc.returncode == 0 and proc.stdout.strip():
                return clean_ocr_text(proc.stdout)
    except Exception:
        pass
    finally:
        try:
            if tmp_img.exists():
                tmp_img.unlink()
        except OSError:
            pass

    return ""


def _run_windows_native_ocr(img_path: Path) -> str:
    """调用 Windows 10/11 自带简体中文 OCR 组件作为零依赖兜底"""
    ps_script = f"""
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
Add-Type -AssemblyName System.Runtime.WindowsRuntime
$null = [Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType = WindowsRuntime]
$null = [Windows.Graphics.Imaging.BitmapDecoder, Windows.Foundation, ContentType = WindowsRuntime]
$null = [Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime]

function Await($WinRtTask, $ResultType) {{
    $asTaskGeneric = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {{ $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' }})[0]
    $asTask = $asTaskGeneric.MakeGenericMethod($ResultType)
    $netTask = $asTask.Invoke($null, @($WinRtTask))
    $netTask.Wait(-1) | Out-Null
    return $netTask.Result
}}

$file = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync('{str(img_path)}')) ([Windows.Storage.StorageFile])
$stream = Await ($file.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])
$decoder = Await ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
$bitmap = Await ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])
$lang = New-Object Windows.Globalization.Language('zh-Hans-CN')
$engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage($lang)
if ($null -eq $engine) {{ $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages() }}
$ocrResult = Await ($engine.RecognizeAsync($bitmap)) ([Windows.Media.Ocr.OcrResult])
Write-Output $ocrResult.Text
$stream.Dispose()
"""
    try:
        creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        proc = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-Command", ps_script],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="ignore",
            timeout=25,
            creationflags=creationflags,
        )
        if proc.returncode == 0:
            return proc.stdout.strip()
    except Exception:
        pass
    return ""


def _count_meaningful_chars(text: str) -> int:
    return len(re.findall(r"[\u4e00-\u9fffa-zA-Z0-9]", text or ""))


def _extract_pdf(path: Path, scratch_dir: Path) -> ExtractedContent:
    import fitz  # PyMuPDF

    head_parts: List[str] = []
    tail_parts: List[str] = []
    ocr_used = False

    with fitz.open(path) as doc:
        page_count = len(doc)
        if page_count == 0:
            return ExtractedContent()

        head_indices = list(range(min(5, page_count)))
        tail_indices = [i for i in range(max(0, page_count - 2), page_count) if i not in head_indices]

        for idx in head_indices:
            txt = doc[idx].get_text("text") or ""
            if txt.strip():
                head_parts.append(txt.strip())

        for idx in tail_indices:
            txt = doc[idx].get_text("text") or ""
            if txt.strip():
                tail_parts.append(txt.strip())

        combined_text = "\n".join(head_parts + tail_parts)
        if _count_meaningful_chars(combined_text) >= 28:
            # 若只有1~2页，tail_text 也赋上最后一页文本便于提取落款成文日期
            last_page_txt = (doc[page_count - 1].get_text("text") or "").strip()
            return ExtractedContent(
                head_text="\n".join(head_parts)[:4000],
                tail_text=("\n".join(tail_parts) if tail_parts else last_page_txt)[-2000:],
                full_text=combined_text[:7000],
                ocr_used=False,
                extract_method="PDF电子文本",
            )

        # 文字过少，判定为扫描件 PDF -> 渲染第1页、第2页(若有)及最后一页(若有)执行中文 OCR
        ocr_used = True
        ocr_head_pages = [0] + ([1] if page_count > 1 else [])
        ocr_tail_pages = [page_count - 1] if page_count > 2 else []

        ocr_head_texts: List[str] = []
        ocr_tail_texts: List[str] = []

        for idx in ocr_head_pages:
            pix = doc[idx].get_pixmap(dpi=200, alpha=False)
            ocr_txt = ocr_image_bytes(pix.tobytes("png"), scratch_dir)
            if ocr_txt:
                ocr_head_texts.append(ocr_txt)

        for idx in ocr_tail_pages:
            pix = doc[idx].get_pixmap(dpi=200, alpha=False)
            ocr_txt = ocr_image_bytes(pix.tobytes("png"), scratch_dir)
            if ocr_txt:
                ocr_tail_texts.append(ocr_txt)

        all_ocr = "\n".join(ocr_head_texts + ocr_tail_texts)
        tail_final = "\n".join(ocr_tail_texts) if ocr_tail_texts else (ocr_head_texts[-1] if ocr_head_texts else "")
        return ExtractedContent(
            head_text="\n".join(ocr_head_texts)[:4000],
            tail_text=tail_final[-2000:],
            full_text=all_ocr[:7000],
            ocr_used=ocr_used,
            extract_method="扫描件OCR" if all_ocr.strip() else "文件名识别",
        )


def _extract_docx(path: Path) -> ExtractedContent:
    from docx import Document

    doc = Document(str(path))
    paras: List[str] = []
    for p in doc.paragraphs:
        t = p.text.strip()
        if t:
            paras.append(t)

    # 提取前 8 个表格（工程合同封面、签证单、报审表核心要素多在表格内）
    table_lines: List[str] = []
    for tbl in doc.tables[:8]:
        for row in tbl.rows[:30]:
            seen_cells = []
            for cell in row.cells:
                ct = re.sub(r"\s+", " ", cell.text.strip())
                if ct and (not seen_cells or seen_cells[-1] != ct):
                    seen_cells.append(ct)
            if seen_cells:
                table_lines.append(" | ".join(seen_cells))

    # 头部融合前部段落与封面表格，尾部取末尾段落与表格（含签订日期）
    head_chunk = "\n".join(paras[:40] + table_lines[:30])
    tail_chunk = "\n".join(paras[-25:] + table_lines[-15:])
    full_chunk = "\n".join(paras[:120] + table_lines + paras[-30:])

    return ExtractedContent(
        head_text=head_chunk[:4000],
        tail_text=tail_chunk[-2000:],
        full_text=full_chunk[:7000],
        ocr_used=False,
        extract_method="Word电子文本",
    )


def _extract_legacy_doc(path: Path) -> ExtractedContent:
    """提取旧版 .doc / .rtf：Windows 优先调用 Word/WPS COM，跨平台支持 textutil 与二进制宽字符流兜底"""
    text = ""
    if sys.platform == "win32":
        try:
            import win32com.client  # type: ignore
            import pythoncom  # type: ignore

            pythoncom.CoInitialize()
            app = None
            for prog_id in ("Word.Application", "KWPS.Application", "Ket.Application"):
                try:
                    app = win32com.client.DispatchEx(prog_id)
                    break
                except Exception:
                    continue
            if app is not None:
                try:
                    app.Visible = False
                    app.DisplayAlerts = 0
                    doc = app.Documents.Open(
                        str(path.resolve()),
                        ConfirmConversions=False,
                        ReadOnly=True,
                        AddToRecentFiles=False,
                    )
                    text = doc.Content.Text or ""
                    doc.Close(False)
                finally:
                    try:
                        app.Quit()
                    except Exception:
                        pass
                    pythoncom.CoUninitialize()
        except Exception:
            pass
    elif shutil.which("textutil"):
        try:
            proc = subprocess.run(
                ["textutil", "-convert", "txt", "-stdout", str(path)],
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="ignore",
                timeout=15,
            )
            if proc.returncode == 0:
                text = proc.stdout
        except Exception:
            pass

    if _count_meaningful_chars(text) < 15:
        # 纯 Python 二进制 UTF-16LE 中文连续片段挖掘兜底（无需安装 Office）
        try:
            raw = path.read_bytes()
            decoded_u16 = raw.decode("utf-16le", errors="ignore")
            chunks = re.findall(r"[\u4e00-\u9fff0-9A-Za-z，。、《》：；（）〔〕\[\]\-\.\s]{6,}", decoded_u16)
            valid = [c.strip() for c in chunks if len(re.findall(r"[\u4e00-\u9fff]", c)) >= 4]
            text = "\n".join(valid)
        except Exception:
            pass

    text = text.strip()
    return ExtractedContent(
        head_text=text[:4000],
        tail_text=text[-2000:],
        full_text=text[:7000],
        ocr_used=False,
        extract_method="Word旧版文本" if text else "文件名识别",
    )


def _extract_xlsx(path: Path) -> ExtractedContent:
    import openpyxl

    wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
    lines: List[str] = []
    try:
        lines.append("工作表名称：" + "、".join(wb.sheetnames))
        for sheet_name in wb.sheetnames[:6]:
            ws = wb[sheet_name]
            lines.append(f"【{sheet_name}】")
            for row_idx, row in enumerate(ws.iter_rows(values_only=True)):
                if row_idx >= 50:
                    break
                vals = [str(v).strip() for v in row if v is not None and str(v).strip()]
                if vals:
                    lines.append(" | ".join(vals[:15]))
    finally:
        wb.close()

    full = "\n".join(lines)
    return ExtractedContent(
        head_text=full[:4000],
        tail_text=full[-2000:],
        full_text=full[:7000],
        ocr_used=False,
        extract_method="Excel表格文本",
    )


def _decode_zip_member_name(info: zipfile.ZipInfo) -> str:
    """修复 Windows 中文压缩包在 Python zipfile 下的 CP437 -> GBK 乱码问题"""
    name = info.filename
    if info.flag_bits & 0x800:
        return name
    try:
        raw = name.encode("cp437")
        for enc in ("gb18030", "utf-8"):
            try:
                return raw.decode(enc)
            except UnicodeDecodeError:
                continue
    except Exception:
        pass
    return name


def _extract_zip(path: Path) -> ExtractedContent:
    """只读窥探压缩包内部文件名称清单，不解压破坏原压缩包"""
    names: List[str] = []
    with zipfile.ZipFile(path, "r") as zf:
        for info in zf.infolist()[:120]:
            decoded = _decode_zip_member_name(info).strip("/")
            if decoded and not decoded.startswith("__MACOSX"):
                names.append(decoded)
    listing = "压缩包内含文件清单：\n" + "\n".join(names)
    return ExtractedContent(
        head_text=listing[:4000],
        tail_text=listing[-1500:],
        full_text=listing[:6000],
        ocr_used=False,
        extract_method="压缩包清单识别",
    )


def extract_document(path: Path, scratch_dir: Path) -> ExtractedContent:
    ext = path.suffix.lower()
    try:
        if ext == ".pdf":
            return _extract_pdf(path, scratch_dir)
        elif ext == ".docx":
            return _extract_docx(path)
        elif ext in {".doc", ".rtf"}:
            return _extract_legacy_doc(path)
        elif ext in {".xlsx", ".xlsm"}:
            return _extract_xlsx(path)
        elif ext in {".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff", ".webp"}:
            img_txt = ocr_image_bytes(path.read_bytes(), scratch_dir)
            return ExtractedContent(
                head_text=img_txt[:4000],
                tail_text=img_txt[-2000:],
                full_text=img_txt[:6000],
                ocr_used=bool(img_txt.strip()),
                extract_method="图片OCR" if img_txt.strip() else "文件名识别",
            )
        elif ext == ".zip":
            return _extract_zip(path)
        elif ext in {".txt", ".md", ".csv"}:
            raw = path.read_bytes()
            txt = ""
            for enc in ("utf-8-sig", "gb18030", "utf-16"):
                try:
                    txt = raw.decode(enc)
                    break
                except UnicodeDecodeError:
                    continue
            return ExtractedContent(
                head_text=txt[:4000],
                tail_text=txt[-2000:],
                full_text=txt[:7000],
                ocr_used=False,
                extract_method="纯文本读取",
            )
    except Exception:
        pass

    return ExtractedContent(extract_method="文件名识别")

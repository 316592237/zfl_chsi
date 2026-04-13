#!/usr/bin/env python3
"""从 CHSI-Desktp/index.html 提取 #chesat_bigReport 部分生成独立的 HTML 报告，
再通过 Chrome headless 打印为 PDF（report_only.pdf），
并将其作为第一页合并进 8446969_zh.pdf。

用法：
    python3 extract_report.py

依赖：
    pip install pypdf
    系统需安装 google-chrome / chromium
"""
import re
import subprocess
from pathlib import Path

ROOT = Path(__file__).resolve().parent
SRC = ROOT / "CHSI-Desktp" / "index.html"
TMP_HTML = ROOT / "CHSI-Desktp" / "_report_only.html"
OUT_PDF = ROOT / "report_only.pdf"
FINAL_PDF = ROOT / "8446969_zh.pdf"

# 纸张尺寸（pt）。与 8446969_zh.pdf 原第二页 CropBox 保持一致（阅读器实际显示尺寸）
PAGE_W_PT = 613.66
PAGE_H_PT = 793.85
# 内容缩放比例：让报告在页面上填充更紧凑，四周留白约为原 letter 留白的 1/3
SCALE = 1.10


def build_report_html() -> None:
    html = SRC.read_text()
    head_links = re.findall(r'<link[^>]*rel=["\']stylesheet["\'][^>]*>', html)
    head_styles = re.findall(r'<style[^>]*>[\s\S]*?</style>', html)

    # 提取 #chesat_bigReport 整块（含外层 div）
    start = html.find('<div id="chesat_bigReport">')
    if start < 0:
        raise RuntimeError("Cannot find #chesat_bigReport in source HTML")
    i, depth = start, 0
    while i < len(html):
        if html[i : i + 4] == "<div":
            depth += 1
            i = html.find(">", i) + 1
        elif html[i : i + 6] == "</div>":
            depth -= 1
            i += 6
            if depth == 0:
                break
        else:
            i += 1
    report_block = html[start:i]

    out = f"""<!DOCTYPE html>
<html><head><meta charset="UTF-8">
<title>中国高等教育学历认证报告</title>
{chr(10).join(head_links)}
{chr(10).join(head_styles)}
<style>
html, body {{ margin: 0; padding: 0; width: {PAGE_W_PT}pt; height: {PAGE_H_PT}pt; }}
body {{ display: flex; align-items: center; justify-content: center; overflow: hidden; }}
#chesat_bigReport {{
    margin: 0 auto;
    transform: scale({SCALE});
    transform-origin: center center;
}}
.cj-page {{ page-break-inside: avoid; }}
@page {{ size: {PAGE_W_PT}pt {PAGE_H_PT}pt; margin: 0; }}
</style>
</head><body>
{report_block}
</body></html>"""
    TMP_HTML.write_text(out)


def render_pdf() -> None:
    subprocess.run(
        [
            "google-chrome",
            "--headless",
            "--disable-gpu",
            "--no-sandbox",
            f"--print-to-pdf={OUT_PDF}",
            "--no-pdf-header-footer",
            f"file://{TMP_HTML}",
        ],
        check=True,
        stderr=subprocess.DEVNULL,
    )


def merge_into_final() -> None:
    from pypdf import PdfReader, PdfWriter

    report = PdfReader(str(OUT_PDF))
    original = PdfReader(str(FINAL_PDF))

    # 以原第二页的 CropBox（阅读器实际可见尺寸）为基准对齐第一页
    ref_crop = original.pages[1].cropbox
    ref_w, ref_h = float(ref_crop.width), float(ref_crop.height)

    writer = PdfWriter()
    new_first = report.pages[0]
    # 强制第一页的 mediabox 和 cropbox 都匹配第二页 cropbox
    for box_name in ("mediabox", "cropbox", "trimbox", "bleedbox", "artbox"):
        try:
            getattr(new_first, box_name).upper_right = (ref_w, ref_h)
            getattr(new_first, box_name).lower_left = (0, 0)
        except Exception:
            pass
    writer.add_page(new_first)
    # 把原第二页的 mediabox 也裁到 cropbox 大小，保证阅读器一致
    second = original.pages[1]
    for box_name in ("mediabox", "cropbox"):
        try:
            getattr(second, box_name).upper_right = (ref_w, ref_h)
            getattr(second, box_name).lower_left = (0, 0)
        except Exception:
            pass
    writer.add_page(second)
    with open(FINAL_PDF, "wb") as f:
        writer.write(f)


def main() -> None:
    build_report_html()
    render_pdf()
    merge_into_final()
    print(f"OK: {OUT_PDF}")
    print(f"OK: {FINAL_PDF}")


if __name__ == "__main__":
    main()

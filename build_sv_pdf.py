#!/usr/bin/env python3
"""
Generate the Faculty Performance Criteria SV PDF.

Combines:
  1. First two pages of the source .docx (intro text + flowchart on one page)
  2. PNG table images generated from the SV Excel workbook

Usage:
    python build_sv_pdf.py
    python build_sv_pdf.py --input criteria_SV.xlsx --docx source.docx --date "March 16, 2026"
"""
import argparse
import io
import os
import subprocess
import sys
import tempfile
from pathlib import Path

from pdf2image import convert_from_path
from PIL import Image
from playwright.sync_api import sync_playwright
from reportlab.lib.colors import Color
from reportlab.lib.pagesizes import letter
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas

from build_table_pngs import build_html
from criteria_data import load_tables_from_excel

SCRIPT_DIR = Path(__file__).parent
LETTER_W, LETTER_H = letter  # 612, 792
MARGIN = 36  # 0.5 inch

DEFAULT_INPUT = SCRIPT_DIR / "Clinical criteria_input_SV.xlsx"
DEFAULT_DOCX = SCRIPT_DIR / "Dept of Psychiatry & Behav Sciences_ Faculty Performance Criteria_03.01.26.docx"
DEFAULT_OUTPUT = SCRIPT_DIR / "Faculty_Performance_Criteria_SV.pdf"
DEFAULT_DATE = "March 16, 2026"

LIBREOFFICE = "/Applications/LibreOffice.app/Contents/MacOS/soffice"

MISSION_ORDER = ["Clinical", "Education", "Research", "Community_Service", "Departmental_Culture"]


def docx_to_cover_image(docx_path: Path, flowchart_crop_frac: float = 0.60) -> Image.Image:
    """Convert docx to PDF, rasterize pages 1–2, and combine into a single image."""
    with tempfile.TemporaryDirectory() as tmpdir:
        subprocess.run(
            [LIBREOFFICE, "--headless", "--convert-to", "pdf", "--outdir", tmpdir, str(docx_path)],
            check=True, capture_output=True,
        )
        pdf_path = Path(tmpdir) / (docx_path.stem + ".pdf")
        pages = convert_from_path(str(pdf_path), dpi=300)

    page1 = pages[0]
    page2 = pages[1]

    # Crop page 2 to flowchart only (top portion)
    p2_w, p2_h = page2.size
    page2_cropped = page2.crop((0, 0, p2_w, int(p2_h * flowchart_crop_frac)))

    # Scale flowchart to match page 1 width and stack vertically
    p1_w, p1_h = page1.size
    fc_w, fc_h = page2_cropped.size
    fc_scale = p1_w / fc_w
    fc_resized = page2_cropped.resize((p1_w, int(fc_h * fc_scale)), Image.LANCZOS)

    combined_h = p1_h + fc_resized.size[1]
    combined = Image.new("RGB", (p1_w, combined_h), "white")
    combined.paste(page1, (0, 0))
    combined.paste(fc_resized, (0, p1_h))
    return combined


def generate_table_pngs(excel_path: Path) -> dict[str, Path]:
    """Generate PNG table images from the Excel workbook, return {name: path} mapping."""
    tables = load_tables_from_excel(excel_path)
    png_paths = {}

    with sync_playwright() as p:
        browser = p.chromium.launch()
        for sheet_name, data in tables.items():
            html_content = build_html(sheet_name, data)
            page = browser.new_page(viewport={"width": 1100, "height": 800}, device_scale_factor=2)
            page.set_content(html_content, wait_until="networkidle")
            content_height = page.evaluate("document.body.scrollHeight")
            page.set_viewport_size({"width": 1100, "height": content_height + 48})

            filename = sheet_name.replace(" ", "_") + ".png"
            filepath = Path(tempfile.gettempdir()) / filename
            page.screenshot(path=str(filepath), full_page=True)
            page.close()
            png_paths[sheet_name.replace(" ", "_")] = filepath
        browser.close()

    return png_paths


def draw_image_on_page(c: canvas.Canvas, img_or_path, top_align: bool = True):
    """Draw an image scaled to fit within margins on the current page."""
    avail_w = LETTER_W - 2 * MARGIN
    avail_h = LETTER_H - 2 * MARGIN

    if isinstance(img_or_path, Image.Image):
        buf = io.BytesIO()
        img_or_path.save(buf, format="PNG")
        buf.seek(0)
        img_w, img_h = img_or_path.size
        reader = ImageReader(buf)
    else:
        img = Image.open(img_or_path)
        img_w, img_h = img.size
        reader = ImageReader(str(img_or_path))

    scale = min(avail_w / img_w, avail_h / img_h)
    draw_w = img_w * scale
    draw_h = img_h * scale
    x = (LETTER_W - draw_w) / 2
    y = LETTER_H - MARGIN - draw_h if top_align else (LETTER_H - draw_h) / 2

    c.drawImage(reader, x, y, draw_w, draw_h)


def add_footers(c: canvas.Canvas, page_num: int, total_pages: int, date_str: str):
    """Add date and page number footer to the current page."""
    c.setFont("Helvetica", 9)
    c.setFillColor(Color(0.3, 0.3, 0.3))
    y = 30
    c.drawString(MARGIN, y, date_str)
    c.drawRightString(LETTER_W - MARGIN, y, f"Page {page_num} of {total_pages}")


def build_pdf(cover_img: Image.Image, png_paths: dict[str, Path], output: Path, date_str: str):
    """Assemble the final PDF."""
    total_pages = 1 + len(MISSION_ORDER)
    c = canvas.Canvas(str(output), pagesize=letter)

    # Page 1: cover (intro + flowchart)
    draw_image_on_page(c, cover_img)
    add_footers(c, 1, total_pages, date_str)
    c.showPage()

    # Pages 2–6: mission area tables
    for i, name in enumerate(MISSION_ORDER, start=2):
        draw_image_on_page(c, png_paths[name])
        add_footers(c, i, total_pages, date_str)
        c.showPage()

    c.save()
    print(f"Saved {output} ({total_pages} pages)")


def main():
    parser = argparse.ArgumentParser(description="Build Faculty Performance Criteria SV PDF")
    parser.add_argument("--input", type=Path, default=DEFAULT_INPUT, help="Excel workbook with criteria")
    parser.add_argument("--docx", type=Path, default=DEFAULT_DOCX, help="Source .docx for cover pages")
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT, help="Output PDF path")
    parser.add_argument("--date", default=DEFAULT_DATE, help="Date string for footer")
    args = parser.parse_args()

    if not args.input.exists():
        print(f"Error: input file not found: {args.input}", file=sys.stderr)
        sys.exit(1)
    if not args.docx.exists():
        print(f"Error: docx file not found: {args.docx}", file=sys.stderr)
        sys.exit(1)

    print("Generating cover page from docx...")
    cover_img = docx_to_cover_image(args.docx)

    print("Generating table PNGs from Excel...")
    png_paths = generate_table_pngs(args.input)

    print("Assembling PDF...")
    build_pdf(cover_img, png_paths, args.output, args.date)


if __name__ == "__main__":
    main()

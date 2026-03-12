#! /usr/bin/env python3
"""
Generate a PNG table image for each of the 5 mission areas.
Uses Playwright (headless Chromium) to render styled HTML tables.

Usage:
    python build_table_pngs.py          # outputs to png/ directory
"""
import os
import html as html_mod
from playwright.sync_api import sync_playwright
from criteria_data import load_tables

OUTPUT_DIR = os.path.join(os.path.dirname(__file__), "png")

tables = load_tables()

COL_HEADERS = ["Criterion", "1 \u2013 Unsatisfactory", "2 \u2013 Low Satisfactory", "3 \u2013 High Satisfactory", "4 \u2013 Outstanding"]

# ── HTML / CSS ──

CSS = """\
* { margin: 0; padding: 0; box-sizing: border-box; }
body {
    font-family: 'Segoe UI', Calibri, Arial, sans-serif;
    background: white;
    padding: 24px 28px;
    -webkit-font-smoothing: antialiased;
}
.title {
    color: #7A0019;
    font-size: 20px;
    font-weight: 700;
    text-align: center;
    margin-bottom: 6px;
    letter-spacing: 0.5px;
}
.note {
    font-size: 11px;
    color: #444;
    font-style: italic;
    text-align: center;
    margin-bottom: 14px;
    line-height: 1.4;
    max-width: 960px;
    margin-left: auto;
    margin-right: auto;
}
table {
    width: 100%;
    border-collapse: collapse;
    table-layout: fixed;
    font-size: 12px;
}
col.criterion { width: 18%; }
col.rating    { width: 20.5%; }
th {
    background: #7A0019;
    color: white;
    font-weight: 600;
    font-size: 11.5px;
    padding: 8px 10px;
    text-align: center;
    border: 1px solid #5B0013;
}
/* Separator rows */
tr.separator td {
    background: #7A0019;
    color: white;
    font-weight: 600;
    font-size: 11.5px;
    padding: 6px 10px;
    border: 1px solid #5B0013;
    letter-spacing: 0.3px;
}
/* Data rows */
td {
    padding: 7px 10px;
    vertical-align: top;
    border: 1px solid #D0D0D0;
    line-height: 1.45;
    font-size: 11.5px;
}
td.criterion {
    font-weight: 600;
    font-size: 12px;
    background: #F0D5DC;
}
/* Promotion rows */
tr.promo td {
    background: #FFF2CC;
    font-weight: 600;
}
tr.promo td.criterion {
    background: #FFF2CC;
    color: #5B0013;
}
"""


def esc(text: str) -> str:
    return html_mod.escape(text)


def build_html(sheet_name: str, data: dict) -> str:
    rows_html = []
    data_row_count = 0

    for group_idx, (group_label, rows) in enumerate(data["groups"]):
        # Separator row
        rows_html.append(
            f'<tr class="separator"><td colspan="5">{esc(group_label)}</td></tr>'
        )
        data_row_count = 0

        for criterion, values, is_promo in rows:
            cls = []
            if is_promo:
                cls.append("promo")
            cls_str = f' class="{" ".join(cls)}"' if cls else ""

            cells = f'<td class="criterion">{esc(criterion)}</td>'
            for val in values:
                cells += f"<td>{esc(val)}</td>"
            rows_html.append(f"<tr{cls_str}>{cells}</tr>")
            data_row_count += 1

    headers = "".join(f"<th>{esc(h)}</th>" for h in COL_HEADERS)

    return f"""<!DOCTYPE html>
<html><head><meta charset="utf-8"><style>{CSS}</style></head>
<body>
<div class="title">{esc(sheet_name.upper())} PERFORMANCE CRITERIA</div>
<div class="note">{esc(data['note'])}</div>
<table>
<colgroup>
  <col class="criterion">
  <col class="rating"><col class="rating"><col class="rating"><col class="rating">
</colgroup>
<thead><tr>{headers}</tr></thead>
<tbody>
{"".join(rows_html)}
</tbody>
</table>
</body></html>"""


def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    with sync_playwright() as p:
        browser = p.chromium.launch()

        for sheet_name, data in tables.items():
            html_content = build_html(sheet_name, data)
            page = browser.new_page(
                viewport={"width": 1100, "height": 800},
                device_scale_factor=2,
            )
            page.set_content(html_content, wait_until="networkidle")

            # Resize viewport to full content height so screenshot captures everything
            content_height = page.evaluate("document.body.scrollHeight")
            page.set_viewport_size({"width": 1100, "height": content_height + 48})

            filename = sheet_name.replace(" ", "_") + ".png"
            filepath = os.path.join(OUTPUT_DIR, filename)
            page.screenshot(path=filepath, full_page=True)
            page.close()
            print(f"  Saved {filepath}")

        browser.close()
    print(f"\nDone! {len(tables)} PNGs saved to {OUTPUT_DIR}/")


if __name__ == "__main__":
    main()

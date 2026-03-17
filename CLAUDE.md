# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This project restructures the Dept of Psychiatry & Behavioral Sciences Faculty Performance Criteria document (a .docx file) into formatted, horizontally-aligned Excel spreadsheets. The criteria cover 5 mission areas rated 1–4: Clinical, Education, Research, Community Service, and Departmental Culture.

## Key Files

- `criteria_input.xlsx` — User-editable Excel workbook (one sheet per mission area) with Yes/No dropdowns for promotion relevance. Preferred input source.
- `criteria_tables.yaml` — Original YAML source for all mission criteria content. Used as fallback when the Excel workbook is absent.
- `criteria_data.py` — Shared loader/validator that reads either Excel or YAML and returns the normalized table structure expected by generators. Auto-detects format (prefers Excel when present).
- `build_criteria_input.py` — Generates the editable Excel input workbook from YAML (migration/reset tool).
- `build_aligned_criteria.py` — Main script that generates the Excel output. Contains formatting/styling logic using openpyxl.
- `build_table_pngs.py` — Generates PNG table images (one per mission area) using Playwright to render styled HTML tables.
- `build_png_powerpoint.py` — Generates a PowerPoint deck with one PNG table image per slide.
- `build_sv_pdf.py` — Generates the SV variant PDF: combines docx cover pages (intro + flowchart) with PNG tables from an alternate Excel workbook into a single portrait PDF with footers.
- `convert_numbers_to_xlsx.py` — Converts Apple Numbers files to Excel (.xlsx) format.
- `Faculty_Performance_Criteria_Source_2026-03-01.docx` — Original source document.
- `Dept of Psychiatry & Behav Sciences_ Faculty Performance Criteria_03.01.26.docx` — Alternate source document used for SV PDF cover pages.
- `Faculty_Performance_Criteria_Aligned.xlsx` — Primary output file (landscape, UMN branding).
- `Faculty_Performance_Criteria_SV.pdf` — SV variant PDF output (portrait, with docx cover + table PNGs).
- `Faculty_Performance_Criteria_PNGs.pptx` — Generated PowerPoint deck using the PNG table outputs.
- `Clinical criteria_input_SV.xlsx` — Alternate criteria workbook (converted from Numbers) used by the SV PDF builder.
- `png/` — Directory of generated PNG table images (2x retina resolution).
- `png_SV/` — Directory of generated PNG table images for the SV variant.

## Build & Run

```bash
# Activate venv
source .venv/bin/activate

# Generate/regenerate the editable input workbook from YAML
python build_criteria_input.py                     # default: criteria_input.xlsx
python build_criteria_input.py custom_name.xlsx    # custom filename

# Validate criteria structure (auto-detects Excel or YAML)
python validate_criteria.py

# Regenerate the Excel output
python build_aligned_criteria.py

# Regenerate PNG table images
python build_table_pngs.py

# Regenerate PowerPoint from PNG table images
python build_png_powerpoint.py

# Convert Apple Numbers to Excel
python convert_numbers_to_xlsx.py "Clinical criteria_input_SV.numbers"

# Generate the SV variant PDF (cover pages + table PNGs)
python build_sv_pdf.py
python build_sv_pdf.py --input other.xlsx --docx other.docx --date "April 1, 2026"
```

Scripts are executable and can also be run directly (for example, `./build_table_pngs.py`).

## Architecture

`build_aligned_criteria.py` is structured as:
1. **Style definitions** — UMN brand colors (Maroon #7A0019, Gold #FFCC33), fonts, borders, fills for headers/zebra striping/promotion rows/separators.
2. **Shared criteria data** — Can live in `criteria_input.xlsx` (preferred, user-editable) or `criteria_tables.yaml` (original). Each mission area has a `note` and `groups`. Groups separate "Core" vs "Scholarship/Leadership/Promotion" criteria with a visual separator row. Each row has `criterion`, `levels` (4 values), and `promotion_relevant`.
3. **Shared loader/validator (`criteria_data.py`)** — Auto-detects Excel or YAML input. Validates structure and normalizes it to the tuple structure expected by the generators.
4. **Sheet/image generation loops** — Excel, PNG, and PowerPoint scripts iterate over the same loaded tables so editorial updates happen in one place only.

## Design Decisions

- Criteria text can be edited in `criteria_input.xlsx` (user-friendly) or `criteria_tables.yaml`. The loader auto-detects: if the Excel file exists it is used, otherwise YAML. This allows editorial control over wording and horizontal alignment across rating levels while keeping a single source for all outputs.
- Promotion-relevant rows (`is_promotion_relevant=True`) get gold highlighting and dark maroon bold text.
- Page setup uses `fitToHeight=0` to allow multi-page flow with repeating header rows rather than shrinking fonts to force single-page fit.
- The PowerPoint generator expects the PNG outputs to exist and places one mission table per slide for review/presentation use.

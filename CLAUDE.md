# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This project restructures the Dept of Psychiatry & Behavioral Sciences Faculty Performance Criteria document (a .docx file) into formatted, horizontally-aligned Excel spreadsheets. The criteria cover 5 mission areas rated 1–4: Clinical, Education, Research, Community Service, and Departmental Culture.

## Key Files

- `criteria_tables.yaml` — Single source of truth for all mission criteria content.
- `criteria_data.py` — Shared loader/validator that reads YAML and returns the normalized table structure expected by generators.
- `build_aligned_criteria.py` — Main script that generates the Excel output. Contains formatting/styling logic using openpyxl.
- `build_table_pngs.py` — Generates PNG table images (one per mission area) using Playwright to render styled HTML tables.
- `build_png_powerpoint.py` — Generates a PowerPoint deck with one PNG table image per slide.
- `Faculty_Performance_Criteria_Source_2026-03-01.docx` — Original source document.
- `Faculty_Performance_Criteria_Aligned.xlsx` — Primary output file (landscape, UMN branding).
- `Faculty_Performance_Criteria_PNGs.pptx` — Generated PowerPoint deck using the PNG table outputs.
- `png/` — Directory of generated PNG table images (2x retina resolution).

## Build & Run

```bash
# Activate venv
source .venv/bin/activate

# Validate criteria YAML structure only
python validate_criteria.py

# Regenerate the Excel output
python build_aligned_criteria.py

# Regenerate PNG table images
python build_table_pngs.py

# Regenerate PowerPoint from PNG table images
python build_png_powerpoint.py
```

## Architecture

`build_aligned_criteria.py` is structured as:
1. **Style definitions** — UMN brand colors (Maroon #7A0019, Gold #FFCC33), fonts, borders, fills for headers/zebra striping/promotion rows/separators.
2. **Shared criteria data (`criteria_tables.yaml`)** — Each mission area has a `note` and `groups`. Groups separate "Core" vs "Scholarship/Leadership/Promotion" criteria with a visual separator row. Each row has `criterion`, `levels` (4 values), and `promotion_relevant`.
3. **Shared loader/validator (`criteria_data.py`)** — Validates YAML shape and normalizes it to the tuple structure expected by the generators.
4. **Sheet/image generation loops** — Excel, PNG, and PowerPoint scripts iterate over the same loaded tables so editorial updates happen in one place only.

## Design Decisions

- All criteria text lives in `criteria_tables.yaml`, not parsed from the .docx at runtime. This allows editorial control over wording and horizontal alignment across rating levels while keeping a single source for Excel and PNG outputs.
- Promotion-relevant rows (`is_promotion_relevant=True`) get gold highlighting and dark maroon bold text.
- Page setup uses `fitToHeight=0` to allow multi-page flow with repeating header rows rather than shrinking fonts to force single-page fit.
- The PowerPoint generator expects the PNG outputs to exist and places one mission table per slide for review/presentation use.

# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This project restructures the Dept of Psychiatry & Behavioral Sciences Faculty Performance Criteria document (a .docx file) into formatted, horizontally-aligned Excel spreadsheets. The criteria cover 5 mission areas rated 1–4: Clinical, Education, Research, Community Service, and Departmental Culture.

## Key Files

- `build_aligned_criteria.py` — Main script that generates the Excel output. Contains all criteria data as Python dicts and all formatting/styling logic using openpyxl. This is the single source of truth for criteria content and spreadsheet layout.
- `build_table_pngs.py` — Generates PNG table images (one per mission area) using Playwright to render styled HTML tables. Contains its own copy of the criteria data dict to avoid import side-effects.
- `Faculty_Performance_Criteria_Source_2026-03-01.docx` — Original source document.
- `Faculty_Performance_Criteria_Aligned.xlsx` — Primary output file (landscape, UMN branding).
- `png/` — Directory of generated PNG table images (2x retina resolution).

## Build & Run

```bash
# Activate venv
source .venv/bin/activate

# Regenerate the Excel output
python build_aligned_criteria.py

# Regenerate PNG table images
python build_table_pngs.py
```

## Architecture

`build_aligned_criteria.py` is structured as:
1. **Style definitions** — UMN brand colors (Maroon #7A0019, Gold #FFCC33), fonts, borders, fills for headers/zebra striping/promotion rows/separators.
2. **Data dict (`tables`)** — Each mission area has a `note` and `groups`. Groups separate "Core" vs "Scholarship/Leadership/Promotion" criteria with a visual separator row. Each row is `(criterion_name, [level_1, level_2, level_3, level_4], is_promotion_relevant)`.
3. **Sheet generation loop** — Iterates over tables, creates one sheet per mission area with page setup (landscape, letter, fit-to-width, repeating headers), separator rows between groups, zebra striping, and promotion-row highlighting in gold.

## Design Decisions

- All criteria text lives inline in the Python dict, not parsed from the .docx at runtime. This allows editorial control over wording and horizontal alignment across rating levels.
- Promotion-relevant rows (`is_promotion_relevant=True`) get gold highlighting and dark maroon bold text.
- Page setup uses `fitToHeight=0` to allow multi-page flow with repeating header rows rather than shrinking fonts to force single-page fit.

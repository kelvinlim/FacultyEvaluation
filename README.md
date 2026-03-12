# Faculty Performance Criteria — Excel & PNG Generator

Restructures the Dept of Psychiatry & Behavioral Sciences Faculty Performance Criteria document into formatted, horizontally-aligned Excel spreadsheets and high-resolution PNG table images with UMN branding (Maroon & Gold).

The output covers **5 mission areas** rated on levels 1–4:

- Clinical
- Education
- Research
- Community Service
- Departmental Culture

## Features

- Landscape orientation with repeating headers for multi-page printing
- UMN brand colors (Maroon `#7A0019`, Gold `#FFCC33`)
- Color-coded rating columns and zebra-striped rows in Excel output
- PNG output uses non-zebra body rows and pale-gold highlight for promotion-relevant criteria
- PowerPoint output can place one PNG table per slide for presentation review
- Separator rows between core duties and scholarship/leadership/promotion criteria

## Setup

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Usage

All generator/validator scripts include a shebang and executable bit, so you can run either `python script.py` or `./script.py`.

```bash
# Validate criteria YAML structure only
python validate_criteria.py

# Generate Excel spreadsheet
python build_aligned_criteria.py

# Generate PNG table images (one per mission area)
python build_table_pngs.py

# Generate PowerPoint with one PNG table per slide
python build_png_powerpoint.py
```

This generates `Faculty_Performance_Criteria_Aligned.xlsx`, 5 PNG files in `png/`, and `Faculty_Performance_Criteria_PNGs.pptx`.

## Criteria Workflow

All criteria content is authored in `criteria_tables.yaml`.

- `build_aligned_criteria.py` loads criteria through `criteria_data.py` and generates the Excel workbook.
- `build_table_pngs.py` loads the same criteria source and generates PNG tables.
- `build_png_powerpoint.py` places the generated PNG tables into a PowerPoint presentation with one slide per mission area.
- `validate_criteria.py` validates the YAML structure only (no output generation).

### YAML Schema (per mission area)

- `note`: string
- `groups`: list of group objects
- group object:
	- `label`: string
	- `rows`: list of row objects
- row object:
	- `criterion`: string
	- `levels`: list of exactly 4 strings
	- `promotion_relevant`: boolean

## Editing Guidelines

Use this checklist when updating criteria text.

- Edit only `criteria_tables.yaml` for criteria content changes.
- Keep mission names and group labels stable unless you intentionally want output headings to change.
- Keep every `levels` entry at exactly 4 strings (ratings 1 through 4).
- Use `promotion_relevant: true/false` (boolean), not quoted strings.
- Prefer short, consistent phrasing across levels so horizontal comparisons stay readable.

Do:

- Keep parallel sentence style within each row's four levels.
- Keep clinical/education/research terminology explicit and measurable when possible.
- Run `python validate_criteria.py` before generating outputs.

Don't:

- Do not add extra fields to rows unless loader validation is updated.
- Do not use tabs or malformed indentation in YAML.
- Do not edit duplicated criteria text in generator scripts; they are loaded from YAML.

Recommended update flow:

1. Edit `criteria_tables.yaml`.
2. Run `python validate_criteria.py`.
3. Run `python build_aligned_criteria.py`.
4. Run `python build_table_pngs.py`.
5. Run `python build_png_powerpoint.py`.
6. Review generated workbook, PNGs, and PowerPoint.

## Outputs

- Excel output: `Faculty_Performance_Criteria_Aligned.xlsx` (written to the repository root)
- PNG outputs: files in `png/` (one image per mission area)
- PowerPoint output: `Faculty_Performance_Criteria_PNGs.pptx` (written to the repository root)

## Project Structure

| File | Description |
|------|-------------|
| `criteria_tables.yaml` | Single source of truth for all mission criteria text and ratings |
| `criteria_data.py` | Shared loader + validation for criteria YAML consumed by both generators |
| `build_aligned_criteria.py` | Main Excel generator and formatting logic |
| `build_table_pngs.py` | PNG table image generator using Playwright (headless Chromium) |
| `build_png_powerpoint.py` | PowerPoint generator that places each PNG table on its own slide |
| `requirements.txt` | Python dependencies (openpyxl, playwright, PyYAML, python-docx, python-pptx) |
| `Faculty_Performance_Criteria_Source_2026-03-01.docx` | Original source document |
| `Faculty_Performance_Criteria_Aligned.xlsx` | Generated Excel output |
| `png/` | Generated PNG table images (one per mission area, 2x retina resolution) |

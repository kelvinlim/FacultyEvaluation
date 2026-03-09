# Faculty Performance Criteria — Excel Generator

Restructures the Dept of Psychiatry & Behavioral Sciences Faculty Performance Criteria document into formatted, horizontally-aligned Excel spreadsheets with UMN branding (Maroon & Gold).

The output covers **5 mission areas** rated on levels 1–4:

- Clinical
- Education
- Research
- Community Service
- Departmental Culture

## Features

- Landscape orientation with repeating headers for multi-page printing
- UMN brand colors (Maroon `#7A0019`, Gold `#FFCC33`)
- Color-coded rating columns and zebra-striped rows
- Gold-highlighted promotion-relevant criteria
- Separator rows between core duties and scholarship/leadership/promotion criteria

## Setup

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Usage

```bash
python build_aligned_criteria.py
```

This generates `Faculty_Performance_Criteria_Aligned.xlsx`.

## Project Structure

| File | Description |
|------|-------------|
| `build_aligned_criteria.py` | Main script — contains all criteria data and Excel formatting logic |
| `requirements.txt` | Python dependencies (openpyxl, python-docx) |
| `Faculty_Performance_Criteria_Source_2026-03-01.docx` | Original source document |
| `Faculty_Performance_Criteria_Aligned.xlsx` | Generated output |

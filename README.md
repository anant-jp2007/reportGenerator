# reportGenerator

Generates **test certificates for HDPE pipes (IS 4984 : 2016)** from a production register.
Upload the register as it was last time and as it is now; the app finds the new production entries, fills a
certificate for each one, and downloads them as **PDF, Excel and/or Word**.

It runs as a small local web app (Flask) on the same Windows PC, with no internet connection or login.

```
Production register (.xls/.xlsx)        Certificate template (.xlsx)       IS 4984 weight chart
   old copy + new copy                    layout, fonts, fixed texts          OD / wall / ovality limits
          │                                         │                                 │
          ▼                                         ▼                                 ▼
   find new entries  ──►  resolve size (OD, PN, SDR)  ──►  fill every cell  ──►  Excel / PDF / Word
   (grade + year + Sr. No.)     against the chart            (+ test values)        one file each or combined
```

## Features

- **New-entry detection.** Compares two copies of the register sheet by sheet (`Production 63`, `PE 63`, ...), matching rows by Sr. No.
  A row is identified by grade + financial year + Sr. No. (Sr. No. restarts every April). Rows edited after
  their certificate was made, and rows that only have a pre-typed report number, are ignored.
- **Columns found by heading, not position.** Extra or moved columns don't matter. Headings are matched
  loosely ("Sr. No.", "SR.NO", "S.No."…). A missing or ambiguous heading stops that sheet with a clear
  message instead of guessing.
- **Size resolution against the IS 4984 chart.** `63 X 10 ( SDR 13.6 )` gives OD 63, PN 10 and SDR 13.6. A
  missing SDR is worked out from grade + PN. Sizes the standard doesn't define for that grade are listed
  as not certifiable, with the reason.
- **Certificate contents.** Header values from the row, testing end date = start + 3 days, specification
  limits from the chart, and test values drawn fresh each time inside configurable ranges. For example, OD and
  wall thickness stay strictly inside the chart limits, and MFR always runs PE 100 < PE 80 < PE 63. Every rule
  is written down in [docs/report-rules.md](docs/report-rules.md).
- **Three output formats that look the same as Excel's print.** The Excel file fills the template. The PDF is
  drawn directly from the template's layout (no Excel needed). The Word file is an editable table on the same
  grid. Both PDF and Word were measured against Excel's own print: across all 174 text pieces the median
  position difference is about 0.25 pt, and every point of every border line is printed.
- **Several companies.** Each company has a profile in the settings: its header (name, address, licence number),
  its own template (same table, its own header layout) and optional pictures such as a logo or stamp. The first
  screen has a company menu.
- **Review before generating.** Untick entries, change the report number (taken from the register's Report No.
  column) if needed, then choose the formats and whether to have one file per certificate or one file per grade.
  The ZIP is sorted by grade: a `PE 63` / `PE 80` / `PE 100` folder per grade, or `PE 63.pdf`, `PE 80.pdf`... when
  combined. Every download is also saved under `output/`.

## Requirements

- Windows 10 or 11. The PDF uses the Windows fonts named in the template: Times New Roman, and Cambria for
  one cell. Other systems fall back to built-in Times.
- Python 3.10 or newer.
- The Python packages in [requirements.txt](requirements.txt). `pymupdf` is only used by the tests.
- Excel and Word are **not** needed to run the app. They're only needed for the comparison tools in `tools/compare/`.

## Getting started

```bash
git clone https://github.com/anant-jp2007/reportGenerator.git
cd reportGenerator
python -m pip install -r requirements.txt
```

Start the app:

```bash
python -m app.web
```

The app opens a browser at <http://127.0.0.1:5050/>; only this computer can reach it. On Windows, double-clicking
**`Start.bat`** does the same: it installs the packages the first time, closes any copy of the app that's already
running, and starts it.

To try it without a real register, make two sample registers (4 entries, then 10) and upload them:

```bash
python tools/make_sample_registers.py samples
```

## The production register

One sheet per grade, named like `Production 63` or `PE 63` (case, spaces and punctuation ignored); other sheets are
ignored and a grade a company doesn't make simply has no sheet. The sheet name gives the grade. Only five
columns are read, found by their heading anywhere in the first 30 rows:

| Heading | Used for |
|---|---|
| Sr. No. | identifying the row (with grade and financial year) |
| Date | date of manufacturing; testing start; testing end = +3 days |
| Size | `OD X PN ( SDR x )`, e.g. `32 X 6 ( SDR 13.6 )`; the SDR part may be left out |
| Report No. | the certificate's file name (`MP100/019` → `MP100-019.pdf`) |
| Batch Number | printed on the certificate |

Month title rows, TOTAL rows and rows without a date and size are skipped.

## Configuration

Everything that may change lives in [config/settings.json](config/settings.json). It's checked on every start,
and problems are reported in plain words (a range with min above max, MFR ranges that overlap, two values
aimed at the same cell, a company with an empty header line…).

| Section | What it holds |
|---|---|
| `company` | company profiles (template, header cells, pictures, header layout) and which one is the default |
| `report` | report numbering, grade order, date format, testing start/end offsets |
| `header_cells`, `chart_cells` | which template cell gets which value |
| `random_values` | ranges, decimals and steps for every test value |
| `empty_cells` | test values deliberately left blank |
| `fixed_text` | every fixed text on the certificate (`^{2}` marks a superscript) |

**Adding a company:** copy a profile and give it a new id, a `name`, its header texts by cell and a `template`
path. Header-only changes (same rows as the base) can reuse `templates/test_report.xlsx`; for a different header
layout add `template_layout` (merged cells, fonts, row heights) and run
`python tools/make_company_template.py <id>` to build its template. The table (rows 6 onwards) always stays the
same. Pictures (`images`) are placed by cell plus an offset in points. Keep the page within A4 at 100%
(`tests/test_settings.py` checks this), so the PDF and Word copies stay identical to Excel's print.

## The IS 4984 weight chart

[data/is4984_chart.json](data/is4984_chart.json) holds, for 36 nominal ODs (16–2500 mm) and 10 SDRs, the mean OD
limits, maximum ovality, pipe-end tolerance, wall thickness limits and weight per metre, plus the grade + SDR → PN
table. It's imported from a weight chart workbook laid out like the standard's tables:

```bash
python tools/import_chart.py path/to/weight_chart.xls
```

## The certificate template

[templates/test_report.xlsx](templates/test_report.xlsx) is built from a sample certificate by
`tools/make_template.py`, which empties the value cells, fixes the page setup and keeps all formatting.
The PDF and Word writers read their layout from it (column widths, row heights, merged cells, borders, fonts,
alignment), so a change to the template changes all three formats.

## Project structure

```
app/
  production.py      read the register, find new rows, resolve sizes
  chart.py           IS 4984 chart lookups
  settings.py        load and check config/settings.json
  report.py          build every cell of a certificate (value + printed text)
  layout.py          template layout, shared by the PDF and Word writers
  render_excel.py    Excel output (fills the template)
  render_pdf.py      PDF output (reportlab)
  render_word.py     Word output (python-docx)
  web.py             Flask app: choose files → check entries → download
  templates/ static/ web pages and styles (design notes: docs/web-design.md)
config/settings.json
data/is4984_chart.json
templates/test_report.xlsx
docs/                report rules, web design notes
tools/               chart import, template build, sample registers, preview, compare/
tests/               test files and run_all.py
Start.bat            Windows launcher
```

## Tests

```bash
python tests/run_all.py
```

The test files cover the chart (against issued certificates), the settings, reading registers with altered
layouts, the report engine (1000+ generated certificates checked against every rule), each output format, the
website, and an end-to-end run on generated sample registers. Most test files take a register to read as
their first argument (default `data/source/production_sample.xls`); `test_end_to_end.py` makes its own and needs none.

## Development tools

| Tool | Purpose |
|---|---|
| `tools/preview_report.py` | print a certificate's values for a register row, without making files |
| `tools/make_sample_registers.py` | before/after sample registers covering the special cases |
| `tools/compare/` | measure our PDF / Word against Excel's own print (needs Excel and Word) |

## Local files

Files that belong to an installation stay on that computer and aren't part of the repository: an optional
company pictures (`assets/`), input files (`data/source/`, `samples/`), uploaded registers (`uploads/`, deleted after a day)
and generated certificates (`output/`).

## Limitations

- Built for one PC: no accounts, and only reachable from the same computer.
- PDF fonts come from `C:\Windows\Fonts`. On other systems the PDF uses built-in Times, and positions match less exactly.
- Sizes of 710 mm and above have no ovality in the chart; those certificates leave ovality blank.

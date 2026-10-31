# reportGenerator

Turns a pipe factory's **production register** into **IS 4984 : 2016 HDPE pipe test certificates**, as PDF, Excel or Word.

> Upload today's register. The app remembers what was already certified, finds the new entries and makes a certificate for each one.

Runs locally on one Windows PC: a small Flask web app, with no internet and no login.

---

## How it works

```mermaid
flowchart LR
    A[("Record of certified<br/>entries, per company")] --> C
    B[/"Register<br/>(.xls / .xlsx)"/] --> C
    C["1 · Find new entries<br/>by Sr. No."] --> D["2 · Read size<br/>OD · PN · SDR"]
    D --> E["3 · Look up limits<br/>IS 4984 chart"]
    E --> F["4 · Fill certificate<br/>+ test values"]
    F --> G["5 · Review<br/>in the browser"]
    G --> H[("ZIP<br/>PDF · Excel · Word")]
    H -. "entries added" .-> A
```

## Using it

| Step | Screen | What you do |
|:-:|---|---|
| 1 | **Choose files** | Pick the company, then the register as it is now (the first time for a company: the old one too) |
| 2 | **Check entries** | See the new entries; untick any, edit the Report No.; **View** a certificate as it prints and change any value |
| 3 | **Download** | In the side panel: Report 1 and/or 2, formats, *one file per certificate* or *one file per grade* |

The ZIP comes out sorted by grade:

```
Test certificates 2026-10-30.zip
├── PE 63/   MP63-171.pdf, MP63-171.xlsx …
├── PE 80/   …
└── PE 100/  …
```

**Report 2** (testing record) can be made instead of or with the certificates: the same entries and values, 4 side by side per sheet, all in one Excel file. An earlier, incomplete Report 2 can be uploaded so its free columns are filled first.

Rows the app can't certify (an unknown size, an empty size…) are listed with the reason, never silently dropped.

### Look

| | |
|---|---|
| **Theme** | Light: warm paper, ink, a tangerine key for the one main action, green for done |
| **Check screen** | Entries on a ruled sheet; the job sheet (tally, reports, formats, download) stays beside them |
| **Fonts** | Red Hat Display / Text / Mono, bundled in `app/static/fonts/`, so it works offline |
| **CSS** | `tokens.css` (colours, sizes) → `base.css` (shared rules) → `style.css` (the screens) |

Design notes: [docs/web-design.md](docs/web-design.md).

---

## Finding new entries

**Against the record** (normal): each company has a record of what was certified, `ledger/<company>.json`, kept on the PC and not in git. After every download its entries are added, with the Report No. used. An entry not in the record is new.

| Situation | What happens |
|---|---|
| First time for a company | The old register is needed once; its entries count as done and start the record |
| *Compare with an old register instead* | The two registers are compared as before; the old one's entries are added to the record |
| Same entries downloaded again | Nothing is added twice |
| Record damaged | Kept aside (renamed); the old register is asked for, and a new record starts |
| A grade the record hasn't seen (new or renamed sheet) | That grade is skipped with a message, like a sheet missing from one of two registers |

Both ways use the same matching:

```mermaid
flowchart TD
    R["Row in the register"] --> K{"Same grade + year + Sr. No.<br/>in the record (or old register)?"}
    K -- "no" --> N["✅ New → certificate"]
    K -- "yes" --> O["Already done → skip<br/>(even if edited)"]
```

- **Key = grade + financial year + Sr. No.** (Sr. No. restarts every April).
- **Repeated Sr. No.:** each recorded (or old) row covers one row, so only the extra rows count as new.
- **Sheets:** one per grade, named like `Production 63` or `PE 63`. Other sheets are ignored.
- **Columns** are found by **heading**, not by position:

| Heading | Used for |
|---|---|
| Sr. No. | identifying the row |
| Date | manufacturing date → testing dates |
| Size | `32 X 6 ( SDR 13.6 )`: the SDR part is optional |
| Report No. | report number and file name (`MP100/019` → `MP100-019.pdf`) |
| Batch Number | printed on the certificate |

---

## What goes on a certificate

```mermaid
flowchart LR
    subgraph Sources
      P["Company profile"]
      R["Register row"]
      C["IS 4984 chart"]
      S["settings.json ranges"]
    end
    P --> H["Header<br/>name · address · logo · CM/L"]
    R --> I["Details<br/>Report No. · dates · batch · size"]
    C --> L["Spec limits<br/>OD · wall · ovality · pipe end"]
    S --> T["Test values<br/>random, always within limits"]
```

All the fill rules are in [docs/report-rules.md](docs/report-rules.md).

---

## Quick start

```bash
git clone https://github.com/anant-jp2007/reportGenerator.git
cd reportGenerator
python -m pip install -r requirements.txt
python -m app.web
```

It opens at **http://127.0.0.1:5050**. On Windows you can double-click **`Start.bat`** instead.

To try it with sample data:

```bash
python tools/make_sample_registers.py samples
```

**Needs:** Windows 10/11, Python 3.10+. Excel and Word are *not* needed to run the app.

### On a USB stick (no install on the other PC)

```bash
python tools/make_portable.py
```

Run once, with internet. It puts a portable Python and the libraries in `python/` (~66 MB, not in git). Copy the whole folder to a USB stick: on any 64-bit Windows PC, double-click **`Start.bat`**. Nothing is installed there and no internet is needed. Without `python/`, `Start.bat` uses the PC's own Python as before.

---

## Code map

```mermaid
flowchart LR
    W["web.py<br/>Flask UI"] --> P["production.py<br/>read + diff registers"]
    P --> CH["chart.py<br/>IS 4984 lookups"]
    W --> LG["ledger.py<br/>record of certified"]
    LG --> P
    W --> PV["preview.py<br/>certificate on screen"]
    W --> NC["new_company.py<br/>add a company"]
    W --> R["report.py<br/>build every cell"]
    R --> X["render_excel.py"]
    R --> PDF["render_pdf.py"]
    R --> WD["render_word.py"]
    R --> R2["report2.py<br/>4 entries per sheet"]
    L["layout.py<br/>template grid"] --> PDF
    L --> WD
    L --> PV
    ST["settings.py"] -.-> R
```

| Path | What's there |
|---|---|
| `app/` | the code above |
| `app/templates/`, `app/static/` | the three screens, CSS, fonts, logo |
| `config/settings.json` | **everything editable**: companies, cells, ranges, texts, Report 2 rows |
| `templates/*.xlsx` | certificate templates (one per company header) and `report2.xlsx` |
| `data/is4984_chart.json` | the IS 4984 weight chart |
| `tools/` | template builders, chart import, sample data, comparison tools |
| `tests/` | `python tests/run_all.py` |

---

## Adding a company

**In the app:** on the first screen, **Add a company**. Type the name, choose one of the company's own test certificates (.xls or .xlsx) and, if its logo isn't found, the logo (a stamp is optional).

```mermaid
flowchart LR
    S[/"Company's certificate"/] --> R["Read its header<br/>rows above Test Report No.<br/>+ logo, stamp"]
    R --> C["Check screen<br/>5 rows: text · size · bold<br/>sample certificate beside"]
    C -- "Show again" --> C
    C -- "Add" --> A[("settings.json profile<br/>+ template + pictures")]
    A --> M["In the company menu"]
```

- The header is laid out like the other companies' (logo on the left, rows 1–3 beside it, contact row, title row); sizes shrink by themselves until the page fits A4 at 100%.
- Nothing is saved until **Add**; settings.json is checked first and the old one kept as `config/settings.previous.json`. If anything fails, nothing is left behind.
- The company menu reads settings.json on every page, so the new company is there straight away.

**By hand:**

1. In `config/settings.json`, copy a profile under `company.profiles`.
2. Set its `name`, header texts (by cell), `images` (logo/stamp) and `template_layout`.
3. Build its template:
   ```bash
   python tools/make_company_template.py <id>
   ```
4. Run `python tests/run_all.py`. It checks that the page still fits A4 at 100%.

Only the header changes. The table is the same for every company.

---

## Good to know

- **The three formats match.** PDF and Word are drawn from the Excel template's grid and were measured against Excel's own print, to within about 0.25 pt.
- **Settings are checked on start.** Mistakes such as min above max or overlapping ranges are reported in plain words.
- **Local only.** Uploads, output, the record of certified entries and company pictures stay on the PC and are not in this repo.
- **Limitation:** sizes of 710 mm and above have no ovality in the chart, so that value is left blank.

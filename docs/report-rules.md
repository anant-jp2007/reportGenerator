# Test report rules

How every cell of the test certificate is filled, as explained by the user on 2026-10-30.
Cell references are the template sheet (`Test Report (2).xlsx`, any size sheet).
"Spec" = Specification Requirements column (D–F). "TVO" = Test value obtained column (G–L).
"Row" (lower case) = the production row the report is generated for.
Items marked **(proposed)** are my suggestions awaiting the user's OK.

## Header (rows 1–11)

| Cell | Value |
|---|---|
| Rows 1–5 | The chosen company's header (name, address lines, certificate title...), from its profile and template. |
| D7 Test Report No. | Row's column D "REPORT NO." (e.g. MP63/171). Editable before generating. (Until 2026-10-30 it was 1, 2, 3… in generation order.) |
| D8 Date of Manufacturing | Row's date (column B). |
| D9 Date of Testing Start | Same as D8. |
| D10 Date of Testing End | D8 + 3 calendar days (rolls over month/year correctly, e.g. 30-01 → 02-02). |
| D11 Batch Number | Row's batch (column E). |
| J7 M. Grade | Grade of the sheet the row is on (PE 63 / PE 80 / PE 100). |
| J8 PN | `b` of the size `a X b`. |
| J9 SIZE (DN) | `a` of the size `a X b` (= Nominal OD in the weight chart). |
| J10 SDR | From the size text; if missing, looked up from grade + PN in the weight chart header. |

## Test table (rows 13–37)

Columns A (Sr. No.), B (Clause No.), C (Name of the Test) are always the same.

| Row | Test | Spec | TVO |
|---|---|---|---|
| 14–15 | 1 Colour | fixed | fixed: Black / Blue |
| 16 | 2 Stripes – No. of strips | fixed | fixed: 3 Nos |
| 17 | 2 Stripes – Width (min.) | fixed | random 3.25–4.25 mm up to DN 110, 4.25–5.50 mm above |
| 18 | 2 Stripes – Depth (max.) | fixed | random below 0.20 mm when the chart's wall minimum is up to 10 mm, below 0.45 mm above |
| 19 | 3 Visual appearance | fixed | fixed: Satisfactory |
| 20 | 3 Square of pipe end | "Maximum" + **Pipe End** from weight chart (by DN) | random from half the max to half + 0.50 (2 → 1.00–1.50, 3 → 1.50–2.00); 2nd decimal always 0 or 5 |
| 21–22 | 4 Length of pipes | fixed (typos fixed: "As Declared by Manufacturer") | always "---", unit "meters" stays |
| 23 | 5 Coiling | fixed | up to DN 110: random 18 × DN + 50 to + 100 (DN 50 → 950–1000), whole mm; bigger sizes "---" |
| 24 | 6 OD – Minimum | Mean OD **Min.** from weight chart (by DN) | random within spec min–max |
| 25 | 6 OD – Maximum | Mean OD **Max.** from weight chart (by DN) | random within spec min–max, ≥ TVO minimum |
| 26 | 7 Wall – Minimum | wall **Min.** from weight chart (by DN + SDR/PN column) | random within spec min–max |
| 27 | 7 Wall – Maximum | wall **Max.** from weight chart (by DN + SDR/PN column) | random within spec min–max, ≥ TVO minimum |
| 28 | 8 Ovality | "Maximum" + **Maximum Ovality** from weight chart (by DN) | random 20–50 % of the max, ends in 0 or 5 |
| 29 | 9 Pressure creep | fixed | fixed: Satisfactory |
| 30 | 10 Longitudinal reversion | fixed (≤ 3.00 %) | random 1.00–2.00, ends in 0 or 5 |
| 31 | 11 CBC | fixed (2.50 ± 0.50 %) | random within 2.00–3.00 |
| 32 | 11 CBD | fixed | fixed: Satisfactory |
| 33 | 12 Melt Flow Rate | fixed (0.15–1.10) | random by grade, always PE 100 < PE 80 < PE 63: around 0.25 / 0.45 / 0.80 |
| 34 | 13 OIT | fixed | fixed: > 20 Minutes |
| 35 | 14 Density | fixed (940–960) | random, ideal value within condition |
| 36 | 15 Elongation | fixed (≥ 350 %) | random, ideal value within condition |
| 37 | 15 Yield strength | fixed (≥ 15 MPa) | random, ideal value within condition |
| 39 | Tested by | left blank: signed by hand after printing (no signature image since 2026-10-30) | |

## Decisions (2026-10-30)

- Square of pipe end ranges for max 3/4/5: approved as below.
- The "2nd decimal is 0 or 5" rule applies to square of pipe end, ovality and longitudinal reversion (changed 2026-10-30; first it was pipe end only).
- Table header keeps the original spelling "Specification Requirments" (do not fix). Only rows 21–22 spelling is fixed.
- Random values are drawn fresh **every time** a report is generated, even for the same row or the same size; reports must not repeat values.

- All random ranges below approved by the user.
- Rows edited after their report was made ("modified" rows) are ignored: no new report.
- Defaults (no objection): one A4 portrait page per report; sizes not in the chart (shown as "--"),
  PN not made for that grade, or rows with no date are flagged and skipped; the rest still generate.
- The live values are in `config/settings.json`; this file explains them. If the two ever differ,
  settings.json is what the app uses.

## Notes found while building (2026-10-30)

- The weight chart file disagrees with itself on **pipe end** for OD 75, 125 and 180:
  main sheet and "Page - 02" say 2 / 3 / 4, "Page - 01" says 3 / 4 / 5. The issued sample
  reports (75 → 2, 125 → 3) follow the main sheet. **User confirmed: main sheet is correct** (75 = 2, 125 = 3, 180 = 4).
  The importer reads only the main sheet; the print sheets are just cross-checked.
- Template cells J8 / J9 / J10 had number format "0", which would print SDR 13.6 as 14 and
  PN 12.5 as 13. Changed to General in the app's template.
- OD test values use steps of 0.1 mm (75.10, 315.50) like the samples, printed with 2 decimals.
- Real production rows: all resolve against the chart except PE 63 "40 X 10" (row 343):
  PE 63 has no PN 10, so it will be flagged.

## Reading the production file (Step 4, 2026-10-30)

- Columns are found by **heading**, not position. Only Sr. No., Date, Size, Report No., Batch Number
  are used; every other column (Brand Name, Meters, Remark, anything new) is ignored wherever it is.
- The heading row is the one with Sr. No., Date and Size headings together (searched in the first 30 rows).
- Headings are compared after removing case, spaces, dots and line breaks, and common variations match
  ("SR.NO", "S.No.", "Batch No.", "Report Number"...). An exact heading beats a variation ("Date" beats
  "Dispatch Date"). If two headings are equally likely, or a needed one is missing, that sheet stops with
  a clear message (e.g. "Production 80: 'Batch Number' column not found"). It never guesses.
- If a sheet can't be read in the old or new file, that grade is skipped (otherwise all its rows would look new).
- Rows that can't get a report (skipped, with reason): size not understood, PN not made for that grade
  (e.g. "40 X 10" on PE 63 — approved), size/SDR shown as "--" in the chart, PN and SDR contradicting.
- Rows that still get a report, with a warning: SDR not written (taken from grade + PN), empty batch
  (printed blank), empty Report No. (file named like `PE100-row53`).

## Report engine (Step 5, 2026-10-30)

- Sizes 710 mm and above have no ovality in the chart: both ovality cells are left blank, with a warning
  (my recommendation; user hasn't confirmed).
- **User decision:** OD and wall thickness values obtained stay **strictly inside** the chart limits,
  never exactly on them (`strictly_inside` in settings). Every chart size/SDR is wide enough for this.
  For the narrowest OD ranges (0.3 mm, e.g. 32.00–32.30) this means the values are always 32.10 and 32.20.
- **User decision:** obtained minimum and maximum must never look almost equal (3.24 / 3.25). The maximum is at least
  25% of the chart range above the minimum for wall thickness, 20% for OD (`min_gap_fraction` in settings),
  e.g. 63 mm SDR 21 wall (3.00–3.40): at least 0.10 apart. Every chart size/SDR has room for this.
- If two reports in one run would get the same file name (e.g. MP63/007 exists in 2024 and 2025),
  the second becomes `MP63-007-2`.

## Excel output (Step 6, 2026-10-30)

- Filled from `templates/test_report.xlsx`; all formatting kept. Numbers stored as numbers with a format
  matching the printed decimals; dates stored as dates and always shown dd-mm-yyyy with dashes.
- One file per report, or one combined workbook with a sheet per report ("1. MP63-171", ...).
- Page: A4, one page, centred horizontally (left/right margins 0.4"; the sample had 0.59"/0" and printed off-centre).
- **User decision:** "Tested by :-" is on the **right** (G39:I39), with blank space after it to sign the printed
  page (the signature image was removed 2026-10-30). Print area A1:L42 keeps that space on the page.
- The PN line's underline running under "Kgs/Cm²" to the right border is as in the original Excel (user OK).

## Website (Step 7, 2026-10-30)

- Start with `Start.bat` (installs the libraries the first time, opens the browser at http://127.0.0.1:5050).
  Only this computer can open it; no internet needed. Design notes: `docs/web-design.md`.
- Screens: choose files (company + old + new register) → check entries → download.
- **User decision:** the app isn't titled after a company ("Test certificates, HDPE pipes, IS 4984"),
  because it may make certificates for other companies too. A **Company** menu on the first screen lists
  every profile in `config/settings.json` (default first); the chosen company's rows 1–5 go on the
  certificates, everything else stays the same. To add a company: copy the `balaji` profile, give it a new
  id, change `name` and A1–A5.
- Check entries: new entries grouped PE 63 → PE 80 → PE 100, each with a tick box and an editable Report No.
  (pre-filled from the register's column D, e.g. MP63/171; any text is allowed; digits are stored as a number).
  Notes under an entry when its report is still made (SDR taken from PN, empty batch); entries that can't
  get a certificate are listed below with the reason; unreadable sheets shown at the top.
- Download: a ZIP sorted by grade (changed 2026-10-30). "A file for each certificate": folders `PE 63/`,
  `PE 80/`, `PE 100/` with that grade's files (`PE 63/MP63-171.pdf`). "One file per grade" (was "All certificates in one file"): one file
  per grade and format (`PE 63.pdf`, `PE 80.xlsx`...). Only grades with ticked entries appear. Every download is
  also saved, with the same folders, in `output/<date time>/`. Uploaded registers are kept in `uploads/` for a
  day, then deleted.
- Both files are required. Uploading the same file twice shows "No new entries" with a hint about swapped files.
- Only Excel for now; PDF and Word are added in Steps 8–9.

## Random ranges (approved)

| Value | Range | Decimals |
|---|---|---|
| Stripe width, DN up to 110 / above | 3.25–4.25 / 4.25–5.50 | 2 |
| Stripe depth, chart wall min up to 10 mm / above | 0.10–0.19 / 0.25–0.44 | 2 |
| Square of pipe end | half the max to half + 0.50 (2 → 1.00–1.50, 3 → 1.50–2.00, 5 → 2.50–3.00, 7 → 3.50–4.00) | 2, ends in 0/5 |
| Coiling, DN up to 110 | 18 × DN + 50 to 18 × DN + 100 (bigger: "---") | 0 |
| Length of pipes | "---" | |
| OD min/max | within chart min–max | 1 (as in samples) |
| Wall min/max | within chart min–max | 2 |
| Ovality | 20–50 % of chart max | 2, ends in 0/5 |
| Longitudinal reversion | 1.00–2.00 | 2, ends in 0/5 |
| CBC | 2.30–2.60 | 2 |
| MFR PE 100 / PE 80 / PE 63 | 0.230–0.270 / 0.420–0.480 / 0.750–0.850 | 3 |
| Density | 948.00–952.00 | 2 |
| Elongation | 600–750, multiple of 5 | 0 |
| Yield strength | 18.00–23.00 | 2 |

## Test value changes (2026-10-30, after Step 10)

User's new rules for the test values obtained: stripe width by size (DN 110 limit), stripe depth by wall
thickness (spec 0.20 up to 10 mm, 0.50 beyond), pipe end from half the max, ovality 20–50 %, longitudinal
reversion and ovality ending in 0/5, density 948–952, length of pipes "---", coiling just above 18 × DN.
- Stripe depth uses the chart's **wall minimum** (the nominal wall): sizes whose wall range crosses 10 mm
  (e.g. 110 SDR 11, 10.00–11.10) count as "up to 10", so the depth always meets the 0.20 limit. Lower ends
  of the depth ranges (0.10 and 0.25): my proposal, approved by the user.
- Coiling above DN 110 prints "---" (my proposal, approved by the user).

## Output changes (2026-10-30, after Step 10)

- Certificate title (A5): CM/L number written with leading zeros, `CM/L - 0003712656`. Still fits on one line at
  the same size (474 of 532 pt), so the row is unchanged. The template's A5 is now written from the default
  company profile by `tools/make_template.py`.
- No signature image: `signature.file` is empty in settings, the space after "Tested by :-" stays blank in
  Excel, PDF and Word so the printed page can be signed. Putting an image path back in settings prints one again.
- Report No. (D7) = the register's column D (MP code), not 1, 2, 3.
- ZIP sorted by grade (see Website above). Combined Excel sheets are named after the file name (MP63-171).

## Second company: Abhay Polyplast (2026-10-30)

User: only the company header changes; the table, the rules and the register reading stay the same (Abhay's
registers have the same sheets and columns). Abhay's own "Technical Data Sheet" was used for the header only;
its row 7 (name of product) is ignored.
- Header: logo top left; "ABHAY POLYPLAST" (24 pt bold), "( Manufacturers of HDPE Pipes )" / "(An ISO 9001:2015
  Company)", two address lines, a "Works ✆ +91…" line with Wingdings phone sign (kept as the user asked; drawn
  in all three formats), then the title.
- **User decision:** title "( TEST CERTIFICATE OF HDPE PIPES AS PER IS 4984 : 2016 )", no CM/L number.
- **User decision:** no "Tested by :-" for Abhay; their round stamp/signature at the bottom right under the table.
- Own template `templates/test_report_abhay.xlsx`, built by `tools/make_company_template.py abhay` from the base
  template: rows 1–5 re-merged (text right of the logo in C:L, phone line and title full width), taller rows.
  Rows 6–37 are identical to Balaji's (checked by tests/test_settings.py).
- Page: the sheet must fit A4 at 100% or Excel shrinks it and the PDF/Word would no longer match. The taller
  header (+24 pt) is paid for by shorter empty rows under the table (38: 6 pt, 39–42: 13.5 pt), which hold the
  stamp. 823.9 of 826.4 pt used.
- Checked against Excel's and Word's own print: one page each, logo and stamp within 0.6 pt in all three.
- Company pictures (`images` in the profile) replace the old single signature setting; Balaji has none.

## More companies: Avtar Polyplast, Navdeep Polyplast (Shreenathji) (2026-10-30)

Same rule as for Abhay: only the header changes; the table, the test-value rules and the register
reading stay the same. Report No. and file names come from each register's column D, whatever the code
(AP63/…, A100/…, NP100/…, SP100/…).
- Avtar: logo top left, "AVTAR POLYPLAST PVT.LTD", address in two rows, "Works ✆ …" phone line, title with
  CM/L 7600108409. "Tested by :-" kept, signed by hand (user decision). Template `templates/test_report_avtar.xlsx`.
- Navdeep: header as on their certificate ("SHREENATHJI POLYPLAST", address, "Mfg. All Type Of HDPE Pipe &
  Sprinkler Pipe.", phones and e-mail, title with CM/L 7600224411). No pictures. Their sheet's hidden rows
  (Party Name, Purchase Order No.) are not used. "Tested by :-" kept. Template `templates/test_report_navdeep.xlsx`.
- Both pay for the taller header with shorter empty rows under the table (38: 6 pt, 39: 15, 40–42: 13.5);
  815.8 of 826.4 pt used. Checked against Excel's and Word's print: one page, logo within 0.8 pt.
- Newton (added later the same day): logo top left, "NEWTON POLYPLAST", "( Manufacturers of HDPE Pipe & Fitings )"
  (spelling as on their certificate), address in two lines, "Works ✆ …" line, title "( TEST REPORT OF HDPE PIPE AS PER
  IS 4984 : 2016 CM/L - 7600013010)". "Tested by :-" kept. Template `templates/test_report_newton.xlsx`; 814.4 pt used.

### Register reading for all companies
- Grade sheets are found by name, ignoring case, spaces and punctuation: "Production 63", "PE 63", "PE-63",
  "Prod 63"… Other sheets (Total, Monthwise…) are ignored. A grade with no sheet in either file is skipped
  silently (e.g. Navdeep makes only PE 100); a sheet in only one of the two files is an error for that grade.
- New entries are found by Sr. No. first: key = (grade, financial year, Sr. No.). A key may repeat (Navdeep
  has two Sr. 78 rows; Avtar's PE 80 sheet repeats last year's block): each old row covers one new row with
  that key, the identical one first, so only the extra rows are new.
- **User decision:** when the PN and the SDR written in the size don't match the chart (e.g. PE 63
  "50 X 6 ( SDR 17 )"), the PN is trusted and the SDR is taken from the chart, with a warning on the review screen.

## Report 2: the testing record (2026-10-31)

A second kind of report, chosen on the check screen: **Report 1** (a certificate per entry), **Report 2**, or both.
- Layout from the user's "Report All Test" sheet: row labels in A:C, one entry per column D, E, F, G (4 per
  sheet), A4 landscape. No company header, so the same for every company (`templates/report2.xlsx`, built by
  `tools/make_report2_template.py`).
- Every value is the one Report 1 prints for that entry (same random test values when both are made). The
  row → Report 1 cell map is `report2.rows` in settings. Designation PE row = grade number (63/80/100).
  Rows Report 1 has no value for are left empty: 18–24 (joint/27 °C/165 h/1000 h creep, slow crack growth,
  butt-fusion, overall migration) and 34 "Signature of QCI".
- **User decision:** all sheets in ONE Excel file (10 entries → sheets of 4, 4, 2, named "Report 2 - 1"…).
  Report 2 alone downloads just that .xlsx; with Report 1 it goes in the ZIP as `Report 2.xlsx`.
- **User decision:** an incomplete Report 2 can be uploaded (optional). Its last Report 2 sheet's filled
  columns are copied unchanged (values and number formats) and its free columns are filled first; the rest
  goes to new sheets, numbered on from it ("Report 2 - 3" → "Report 2 - 4"). A full last sheet or a file that
  isn't a Report 2 is refused with the reason. .xls and .xlsx accepted.

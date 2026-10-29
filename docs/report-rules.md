# Test report rules

How every cell of the test certificate is filled, as explained by the user on 2026-10-30.
Cell references are the template sheet (`Test Report (2).xlsx`, any size sheet).
"Spec" = Specification Requirements column (D–F). "TVO" = Test value obtained column (G–L).
"Row" (lower case) = the production row the report is generated for.
Items marked **(proposed)** are my suggestions awaiting the user's OK.

## Header (rows 1–11)

| Cell | Value |
|---|---|
| Rows 1–5 | Company profile (name, ISO line, 2 address lines, certificate title). Same on every report. |
| D7 Test Report No. | 1, 2, 3… in generation order (PE 63 rows, then PE 80, then PE 100, top to bottom). Editable before generating. |
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
| 17 | 2 Stripes – Width (min.) | fixed | random 3.00–7.00 mm |
| 18 | 2 Stripes – Depth (max.) | fixed | random 0.10–0.16 mm |
| 19 | 3 Visual appearance | fixed | fixed: Satisfactory |
| 20 | 3 Square of pipe end | "Maximum" + **Pipe End** from weight chart (by DN) | random in a reasonable range below the max; 2nd decimal always 0 or 5 (e.g. 1.55, 2.35, 6.70) |
| 21–22 | 4 Length of pipes | fixed (typos fixed: "As Declared by Manufacturer") | value empty, unit "meters" stays |
| 23 | 5 Coiling | fixed | value empty, unit "mm" stays |
| 24 | 6 OD – Minimum | Mean OD **Min.** from weight chart (by DN) | random within spec min–max |
| 25 | 6 OD – Maximum | Mean OD **Max.** from weight chart (by DN) | random within spec min–max, ≥ TVO minimum |
| 26 | 7 Wall – Minimum | wall **Min.** from weight chart (by DN + SDR/PN column) | random within spec min–max |
| 27 | 7 Wall – Maximum | wall **Max.** from weight chart (by DN + SDR/PN column) | random within spec min–max, ≥ TVO minimum |
| 28 | 8 Ovality | "Maximum" + **Maximum Ovality** from weight chart (by DN) | random, a decent distance from both 0 and the max |
| 29 | 9 Pressure creep | fixed | fixed: Satisfactory |
| 30 | 10 Longitudinal reversion | fixed (≤ 3.00 %) | random within condition |
| 31 | 11 CBC | fixed (2.50 ± 0.50 %) | random within 2.00–3.00 |
| 32 | 11 CBD | fixed | fixed: Satisfactory |
| 33 | 12 Melt Flow Rate | fixed (0.15–1.10) | random by grade, always PE 100 < PE 80 < PE 63: around 0.25 / 0.45 / 0.80 |
| 34 | 13 OIT | fixed | fixed: > 20 Minutes |
| 35 | 14 Density | fixed (940–960) | random, ideal value within condition |
| 36 | 15 Elongation | fixed (≥ 350 %) | random, ideal value within condition |
| 37 | 15 Yield strength | fixed (≥ 15 MPa) | random, ideal value within condition |
| 39 | Tested by | signature image (`assets/signature.png`) | |

## Decisions (2026-10-30)

- Square of pipe end ranges for max 3/4/5: approved as below.
- The "2nd decimal is 0 or 5" rule applies **only** to square of pipe end. Other random values use normal decimals.
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

## Random ranges (approved)

| Value | Range | Decimals |
|---|---|---|
| Stripe width | 3.00–7.00 | 2 |
| Stripe depth | 0.10–0.16 | 2 |
| Square of pipe end, max 2 | 0.75–1.50 | 2, ends in 0/5 |
| Square of pipe end, max 3 | 1.00–2.25 | 2, ends in 0/5 |
| Square of pipe end, max 4 | 1.25–3.00 | 2, ends in 0/5 |
| Square of pipe end, max 5 | 1.50–3.75 | 2, ends in 0/5 |
| Square of pipe end, max 7 | 2.00–6.00 | 2, ends in 0/5 |
| OD min/max | within chart min–max | 1 (as in samples) |
| Wall min/max | within chart min–max | 2 |
| Ovality | 25–75 % of chart max | 2 |
| Longitudinal reversion | 1.00–2.00 | 2 |
| CBC | 2.30–2.60 | 2 |
| MFR PE 100 / PE 80 / PE 63 | 0.230–0.270 / 0.420–0.480 / 0.750–0.850 | 3 |
| Density | 947.00–951.00 | 2 |
| Elongation | 600–750, multiple of 5 | 0 |
| Yield strength | 18.00–23.00 | 2 |

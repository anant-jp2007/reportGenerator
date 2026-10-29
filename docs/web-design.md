# Website design notes

## Brief
- **Subject:** test certificates for HDPE pipes (IS 4984) at Balaji Polyplast, Rajkot.
- **Audience:** the owner's father. Not technical, uses it on the office PC, offline.
- **Primary job:** upload the old and new production register, check the new entries, download their certificates.
- Minimalist (user's request). Three screens in a fixed order: choose files → check entries → download.

## Tokens

| Name | Hex | Use |
|---|---|---|
| Paper | `#F6F8FA` | page background (cool, faint blue-grey, like a QC office sheet) |
| Sheet | `#FFFFFF` | the working area (table, file boxes) |
| Pipe | `#1B232B` | text. HDPE charcoal, not neutral black |
| Stripe | `#1D5FC4` | the one accent: the blue identification stripes on the pipe. Primary button, focus, the mark |
| Rule | `#D3DAE1` | table rules, box outlines |
| Quiet | `#5B6772` | secondary text |
| Caution | `#9A5B00` on `#FFF6E5` | warnings (report still made) |
| Stop | `#B3261E` on `#FDECEA` | skipped rows, unreadable sheets |

**Type:** one family, **Bahnschrift** (DIN-style, ships with Windows 10/11, so it works offline). DIN is the lettering of
engineering drawings and standards, which fits a certificate tool. Regular for body, SemiBold for headings, the
condensed width for the dense table. Tabular figures for sizes, dates and numbers. Fallbacks: "DIN Alternate",
"Arial Narrow", sans-serif.
Scale (1.25): 13 / 15 / 18 / 23 / 29 px. Body 15/1.5.

## Layout
One left-aligned column, max 860px, generous top margin. No cards, no shadows: content sits on rules.

```
[pipe] Test certificates
       HDPE pipes, IS 4984

Choose files      Check entries      Download       <- the three steps, current one in Stripe blue

Company [ Balaji Polyplast  v ]

┌ Old register ──────────┐  ┌ New register ──────────┐
│ PRODUCTION (old).xls   │  │ PRODUCTION (2).xls     │
└────────────────────────┘  └────────────────────────┘
[ Find new entries ]
```

Review: one table, grouped by grade with the grade as a plain row heading, a tick box and an editable report
number per row, warnings as a second line under the row. Skipped rows listed below the table with the reason.

- The header never names a company: the app may serve several. The company is a choice on the first screen.

## Principles
- **The one bold thing:** three thin blue stripes, the pipe's identification stripes, as the mark and as the step
  indicator underline. Nothing else decorative.
- Words name what the father sees in his register: "entries", "register", "certificate", "Report No.".
  Buttons say exactly what happens: "Find new entries", "Download certificates".
- Problems say what's wrong and how to fix it, e.g. "Production 80: the 'Batch Number' column wasn't found.
  Check the heading in the new register."
- No motion except the button changing to "Making certificates…" while it works.

## Checked against generic defaults
- Not cream + terracotta, not black + neon: the palette comes from the pipe itself (charcoal + blue stripe).
- Not the SaaS card kit: one sheet with rules, no rounded shadow cards.
- No all-caps labels, no "→" on buttons, no middle-dot meta strings, no monospace data labels.
- Numbered steps are used because the screens really are a sequence.

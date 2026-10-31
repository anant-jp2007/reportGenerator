# Website design notes

## Brief
- **Job:** upload the old and new production register, check the new entries, download Report 1 and/or Report 2.
- **Audience:** the owner's father, on the office PC, offline. Repetitive work, so the screens should feel calm
  and a little rewarding, never dark or busy.
- Three screens in a fixed order: choose files → check entries → download.

## Look (redesign 2026-10-31, Hallmark theme "Furrow")
- **Paper and ink:** warm manila paper (`--color-paper`), near-black ink, white sheets. Light throughout.
- **Tangerine** (`--color-accent`) is a fill with ink on it: the primary key, the current step, the mark's bore.
  **Green** is "done": finished steps, a chosen file, the downloaded tick. Caution (amber) and stop (red) only for
  notes and entries without a certificate.
- **Type:** Red Hat Display 600 for heads, Red Hat Text 400/500 for reading, Red Hat Mono for what gets copied
  onto a form (Report No., dates, batch numbers, file names). The font files ship in `app/static/fonts/`
  (OFL), so nothing is fetched from the internet.
- **Mark:** a pipe seen end-on (ink wall, tangerine bore) with a tick through it: a certified pipe.
- Tokens live in `app/static/tokens.css`; `base.css` holds the shared base rules; `style.css` the screens.

## Screens
1. **Choose files:** company field (label inside the box), a line saying what the new entries are checked against,
   then the "New register" notebook card. The first time for a company (no record yet), or after "Compare with an old
   register instead", two cards: "1 Old register" → "2 New
   register". A chosen file turns its card's name green with a tick.
2. **Check entries:** the ledger (a white ruled sheet, grouped by grade) beside the **job sheet**: an ink-headed
   card with the live tally ("4 of 5 ticked"), Report 1 / Report 2, the optional incomplete Report 2, format,
   files and the download key. On a wide screen the job sheet stays in view while you scroll and tick. A red
   chip under the title points to the entries without a certificate, which stay listed below the table.
   Each entry has a **View** key: a window with the certificate drawn as it prints (white page, same grid,
   fonts and line breaks as the PDF). The values that can be changed are tinted tangerine boxes; Previous / Next
   go through the entries, "New random values" and "Undo changes" sit above the page. A changed entry gets an
   "Edited" chip in the list.
3. **Done:** one card with a green tick, what was downloaded, where the copy is saved, "Make more certificates".
4. **Add a company** (link beside the company field; no step bar): name field, the certificate card, and two small
   optional cards for logo and stamp. Then **Check the header**: the five header rows on the left (text, size,
   bold; pictures to replace or remove; "Tested by" kept or not) and the sample certificate on the right, as it
   prints. Typing disables "Add" until "Show again" redraws the page, so what is added is what was seen. After
   adding, the first screen opens with the new company chosen and a green notice.

## Rules kept
- Words name what is in the register: entries, register, certificate, Report No. Buttons say what happens.
- No dashes in copy, no tracked capitals, no eyebrow labels; one filled (tangerine) key per screen.
- Motion only as feedback: key press, hover, a card lifting 1px under the pointer. Nothing moves on its own.
- On a phone the tick and Report No. columns stay pinned while the rest of the table scrolls sideways.

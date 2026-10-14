# Logbook Scanner (Google Apps Script)

A free tool that lives inside a Google Sheet. It shows the Sales Journal photos
uploaded through the Google Form next to an entry grid. A person types in
what the handwritten logbook says, and the tool compares it with
Spreadsheet 2, by date.

No paid APIs and no automatic handwriting reading: everything runs on
Apps Script, which is free with a Google account.

```
Google Form upload ─► Spreadsheet 1 (Form responses)
                        F = Sales Journal photo, G = extra page (optional)
                                   │
                   Scanner panel (inside the sheet)
                   photo viewer: zoom · rotate · "scan look"
                   entry grid: DATE · NAME · CI# · CASH · CARD · VAT SALE · 12% VAT
                   live row check: VAT ≈ 12% of VAT SALE, CASH+CARD ≈ VAT SALE+VAT
                                   │ Save & compare
                                   ▼
             "Journal Entries" tab ──by date──► Spreadsheet 2 (C, D, X–AF)
                                   ▼
             "Comparison" tab: ✅ Match / ⚠️ Check / ❌ Mismatch
```

## Install (about 5 minutes)

1. Open the spreadsheet where you want the tool to live. Spreadsheet 1 is
   the natural home, but any sheet works. You need **edit** access to it
   and **view** access to Spreadsheet 2 and to the Form's uploaded photos.
2. **Extensions → Apps Script.**
3. Create these files and paste in the contents from `apps-script/`:
   - `Code.gs`, `Config.gs`, `Compare.gs` (use **+ → Script** and give each the same name)
   - `Panel.html` (use **+ → HTML**, name it `Panel`)
   - `appsscript.json`: under **Project Settings**, tick *Show "appsscript.json"*,
     then replace that file's contents.
4. Save, then reload the spreadsheet. A **Scanner** menu appears.
5. **Scanner → Set up / refresh column matching.** Approve the permissions
   prompt the first time.
   - The tool creates three tabs: **Scanner Config**, **Journal Entries** and **Comparison**.

## Configure (Scanner Config tab)

| Setting | Default |
|---|---|
| Spreadsheet 2 (URL or ID) | the Spreadsheet 2 link; its `gid` picks the tab |
| Spreadsheet 2 header row | 1 |
| Spreadsheet 2 columns to compare | `C,D,X:AF` |
| Spreadsheet 1: Form responses (URL or ID) | the Spreadsheet 1 link (`gid=921021716`); leave it blank if the tool lives in Spreadsheet 1 |
| Sales Journal / extra page photo columns | `F` / `G` |
| Branch column in Spreadsheet 1 | blank = the column whose header contains "Branch" |
| Tolerance for amounts (±) | 1 |
| VAT rate | 0.12 |

Below the settings is the **column matching** table. The tool reads
Spreadsheet 2's header for each of C, D and X–AF and guesses which logbook
field each one holds. Check every row and fix it with the dropdown.
**Exactly one column must be `DATE`.** Set the column that holds the
branch name to `BRANCH`, so each page is compared only with its own branch's
rows. Mark columns that have no logbook
equivalent `(ignore)`. Run *Set up / refresh column matching* again any
time the columns change; choices you already made are kept.

## Daily use

1. **Scanner → Open scanner.**
2. Pick a **branch** in the top-left dropdown. Branch names come from
   Spreadsheet 1 and are listed A–Z. *All branches* groups the
   submissions by branch, A–Z. The choice is remembered next time.
3. Pick a submission. The newest pending one is shown first, and ✓ marks
   the ones already done.
4. Read the photo. Use zoom, rotate and *Scan look* (grayscale with high
   contrast) to make the handwriting easier to read. Switch between
   *Journal* and *Extra page* photos with the buttons above the photo.
5. **Auto-fill:** type a date (`10/12/26`) into an empty row and press
   **Enter**. If Spreadsheet 2 already has rows for that date and branch,
   they're filled into the grid in **blue**. Check each blue cell against
   the photo: a cell turns white once you edit it, or click
   **Mark all checked**. Saving with blue cells left needs a second click
   (*Save anyway*), and those rows are marked "filled from Spreadsheet 2,
   not checked" in the Journal Entries tab. A filled value always matches
   Spreadsheet 2, so the comparison is only as good as this check.
   Otherwise, type each logbook row. **Enter** moves to the next cell, a new row
   copies the date from the row above, and rows whose numbers don't add up
   turn yellow.
6. **Save & compare.** The results for that page's dates appear in the
   panel, and the full list is in the **Comparison** tab.

Re-opening a done submission loads what was typed before, so it can be
corrected and saved again.

## How the comparison works ("by date")

For each date typed in, every logbook row with that date is compared with
**every Spreadsheet 2 row that has the same date and branch**. The branch
filter applies only when a Spreadsheet 2 column is set to `BRANCH`. Branch
names must match apart from capital letters, spaces and punctuation.
Field by field:

- **Amount fields** (CASH, CARD, VAT SALE, 12% VAT, TOTAL): the day's
  totals must agree within the tolerance.
- **Text fields** (NAME, CI#): each value must appear on both sides.
  A near-miss, such as "Shaina Vilena" against "Shaina Villena", shows as
  ⚠️ *Check* so that typos can be told apart from real gaps.
- A date that isn't in Spreadsheet 2 at all shows as ❌ *Date missing*.

Text dates in Spreadsheet 2 are read **month first** (`09/16/2026`).

## Optional: a button on the sheet

Use **Insert → Drawing** to draw a button, then open the drawing's ⋮ menu →
**Assign script** and enter `openScanner`. Clicking the button opens the
scanner.

## Notes

- The photos stay in the Form owner's Google Drive. The person using the
  scanner must be able to open them. iPhone HEIC photos and PDFs are shown
  using Drive's preview image.
- The script's time zone is set to `Asia/Manila` in `appsscript.json`.
  Change it there if that's wrong.

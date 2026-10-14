# Logbook Scanner (Google Apps Script)

A free tool that lives inside a Google Sheet. It shows a logbook photo
uploaded through the Google Form next to the matching records from
Spreadsheet 2. A person checks the records against the photo, marks
anything that doesn't match, and the result is saved in a **Review** tab.

No paid APIs and no automatic handwriting reading: everything runs on
Apps Script, which is free with a Google account. The checking itself is
done by a person.

```
Google Form upload ─► Spreadsheet 1 (Form responses)
                        F = Sales Journal photo, G = extra page (optional), Branch
                                   │
                   Scanner panel (inside the sheet)
      ┌────────────────────────────┴───────────────────────────────┐
      photo viewer                         Spreadsheet 2 records
      zoom · drag · rotate · scan look     filtered by branch + logbook date + search
                                           click a value that doesn't match the photo
                                           ✓ match · ✕ not in photo · + only in photo
                                   │ Save review
                                   ▼
             "Review" tab: Match / Not match / Not in photo / Only in photo
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

   Copy each file from GitHub with the **Copy raw file** button so nothing
   gets cut off.
4. Save, then reload the spreadsheet. A **Scanner** menu appears.
5. **Scanner → Set up / refresh column matching.** Approve the permissions
   prompt the first time. The tool creates two tabs: **Scanner Config** and
   **Review**.

## Configure (Scanner Config tab)

| Setting | Default |
|---|---|
| Spreadsheet 2 (URL or ID) | the Spreadsheet 2 link; its `gid` picks the tab |
| Spreadsheet 2 header row | 1 |
| Spreadsheet 2 columns to compare | `C,D,X:AF`. Add the invoice/CI# and customer columns here too |
| Spreadsheet 1: Form responses (URL or ID) | the Spreadsheet 1 link (`gid=921021716`); leave it blank if the tool lives in Spreadsheet 1 |
| Sales Journal / extra page photo columns | `F` / `G` |
| Branch column in Spreadsheet 1 | blank = the column whose header contains "Branch" |
| Count identical Spreadsheet 2 rows once | `yes`: rows that repeat in every used column are shown once |
| Only list submissions from the last N days | `30`; blank = the whole history |
| Logbook date column in Spreadsheet 1 | blank = the column whose header contains "Date" (not the Timestamp). Fills the date box automatically |
| Name column in Spreadsheet 1 | blank = the column whose header contains "Name". Shown in the submission picker |

Below the settings is the **column matching** table. It lists every
Spreadsheet 2 column from the setting above with its header. Set each one
to how it is used:

| Use as | Meaning |
|---|---|
| `DATE` | the date of the sale. **Exactly one column.** |
| `BRANCH` | the branch name, so only the submission's branch is listed |
| `SHOW` | shown in the scanner and checked against the photo |
| `(ignore)` | not used |

Run *Set up / refresh column matching* again whenever the columns change;
choices you already made are kept.

## Daily use

1. **Scanner → Open scanner.**
2. Pick a **branch** at the top left (names come from Spreadsheet 1, A–Z),
   then a **submission**. The list/🔍 switch in the top bar sets how you
   browse: normal dropdowns, or search boxes where you type part of a row
   number, date, email or branch (press **/** to jump to the search).
3. Read the photo. Zoom stays centered. When zoomed in, **hold the left
   mouse button and drag** to move around the photo. Use *Scan look*
   (grayscale with high contrast) for faint writing, and switch it off for
   photos with dark shadows.
4. The **logbook date** is filled in from the Form answers (check it
   against the page; ‹ › step a day). Dates are typed the way your sheets
   write them: day first (`13/10/26`) when the sheets show `13/10/2026`.
   The scanner detects which one your sheets use. The right side lists Spreadsheet 2's records for that date and the
   submission's branch. Use **Search** to narrow them down (name, CI#,
   amount), or to find records on any date.
5. Go through the records against the photo:
   - **A value is different** → click it. It turns red, and you can type
     what the photo says underneath.
     Click the struck-out value to undo.
   - **The whole row matches** → ✓. **The row isn't in the photo** → ✕.
   - **Everything left matches** → **Rest match**.
   - **The photo has a row Spreadsheet 2 doesn't** → *Add a row that's in
     the photo…*
6. The tiles at the bottom count **Not match**, **Not checked**, **Match**
   and **Only in photo**. Click a tile to show only that category; click it
   again to show everything.
7. **Save review.** If some records aren't checked yet, the first click
   asks you to confirm (*Save anyway*). Results go to the **Review** tab,
   one line per matched row, per value that doesn't match (with what the
   photo says), per row missing from the photo and per row only in the
   photo, color-coded.

Reopening a submission loads its saved review, so it can be finished or
corrected later. Submissions with a saved review show ✓ in the picker.

The submission picker shows the logbook date, the name, when it was sent
and the row number. The row number is the row in the Form responses tab
itself. A **sorted filter view** shows different numbers, so go by the
name and date.

## Optional: a button on the sheet

Use **Insert → Drawing** to draw a button, then open the drawing's ⋮ menu →
**Assign script** and enter `openScanner`. Clicking the button opens the
scanner.

## Notes

- The photos stay in the Form owner's Google Drive. The person using the
  scanner must be able to open them. iPhone HEIC photos and PDFs are shown
  using Drive's preview image.
- The old **Journal Entries** and **Comparison** tabs from the previous
  version are no longer used. You can delete them.
- The script's time zone is set to `Asia/Manila` in `appsscript.json`.
  Change it there if that's wrong.

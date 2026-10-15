# ACS Hire Report Check

Weekly check of what plant hire suppliers are charging ACS against the ACS in-house hire report.

## The job

1. **ACS in-house hire report**: one Excel workbook, one sheet per job (plus sheets like Yard and Vehicles). Rows are already colour-coded by me (green on hire, amber/red, key for off-hired and suspended): keep those colours.
2. **Supplier hire reports**: one Excel file per supplier, showing what we're being charged for.
3. For every line on each supplier report, look up its **plant code** on all sheets of the ACS report.
   - **Found**: highlight that supplier line green and make it bold, and give the ACS sheet and cell it matched (e.g. `M3 J9!C16`) in an added "ACS match" column.
   - **Not found**: notify me, with a suggestion of what it probably is on the ACS report. Do not mark it.
4. The other direction: ACS lines for that supplier that aren't on the supplier report (probably off-hired or swapped). The supplier's name on the ACS report (e.g. "Sunbelt") is worked out from which ACS lines its codes matched. Lines marked off-hire in the Status column are left out.
5. After everything is checked, take the ACS report's £0 lines that nothing paired with and look for each one (e.g. a ladder) on the supplier reports.

Always work from the files given for this check. Nothing about a particular week's report is built into the script or the web app, and there is no example data.

Call the company **ACS** everywhere (pages, README, code, commits). Never "ACS Civils".

Plant codes are identical in both reports (e.g. a 4.5m manhole box is `BHRAIL4.5` in both).

## How to run

```
python hire_check.py <ACS_hire_report.xlsx> <supplier1.xlsx> [<supplier2.xlsx> ...]
```

- The ACS report always goes first. Every file after it is treated as a supplier report.
- When I say "check this week's hire reports", find the ACS report in the folder (the one with a sheet per job / "ACS" in the name, never the ACS `_checked.xlsx`), and pass every other supplier `.xlsx` as the rest. Never pass a `_checked.xlsx` file back in. Ask only if it's genuinely unclear which file is the ACS report.
- If a supplier file's plant-code column is picked wrong, rerun with `--code-col <letter>` (e.g. `--code-col B`).
- First-time setup: `pip install openpyxl`.

## What the script does

- Reads every sheet of the ACS report and collects every plant code. This is the same as Ctrl+F across the whole workbook.
- On each supplier report, finds the plant-code column automatically: whichever column has the most codes that are also on the ACS report. Each supplier can have a different layout.
- Matching ignores capitals and spaces (`dump9t ` = `DUMP9T`).
- Skips headers, blank rows, titles and total lines. Only single-word, code-like values count as charge lines.
- Saves the marked-up copy as `<supplier>_checked.xlsx` next to the original: matched lines green + bold, plus an "ACS match" column after the last column giving every ACS sheet!cell the code appears in.
- Also saves `<ACS report>_checked.xlsx`: every ACS plant line being charged on the supplier reports in **bold with a thick border** (only the report's own columns), keeping my existing colours. A line listing several assets is only marked when all of them are charged.
- Prints the result on screen for each supplier: either OK with the match count, or every plant code NOT on the ACS report, with its row number and the full charged line, plus a grand total.
- Then prints **suggestions**, worked out once every report is checked:
  - Each charged line whose code isn't on the ACS report is compared with the ACS lines no supplier has charged for: on-hire date, rate (also rate × quantity for "2no" lines), quantity, description words and sizes (4m, 350mm, 8 CFM). A suggestion needs the description to agree plus the date or rate, or both date and rate. Sizes that disagree (4m vs 5m) rule a pair out.
  - **£0 end pass**: ACS lines at £0 (or blank rate) that are still unpaired are looked for among the supplier lines nothing else claimed. These need a shared description word (a ladder is a ladder).
  - Suggestions are text only. Nothing is marked for them.
  - A charged line with no suggestion that went on hire after the ACS report's date (read from its filename, e.g. `01.09.26`) is flagged as probably new plant to add.
- The "on ACS but not charged" list (point 4 above) is shown in the web app only. The script doesn't print it yet.

## Rules

- **Never edit the original files.** Output is always a separate `_checked.xlsx` copy.
- Supplier copies: **green fill + bold on matched lines** and the one added "ACS match" column. ACS copy: **bold + thick border** on charged lines, colours untouched. No new sheets, no red marking, no deleted rows, no other added columns.
- Unmatched plant is **reported, not marked**.
- Don't change the matching logic without asking me first.

## When reporting back to me

Keep it short. For each supplier, give either "all on ACS report" or the list of plant codes not on the ACS report (code, description, charge), each with its suggestion if there is one. Then the £0 end pass. End with the total count. Don't recap the steps.

## Known limits

- Old `.xls` files need LibreOffice to convert. If it's not installed, open the file in Excel and use Save As `.xlsx` first.
- Plant codes containing spaces won't be recognised as codes. On the ACS report, an asset cell with spaces (`P LADDERS 4M CQ`) is still used for suggestions and the £0 pass, but not for matching.
- Rates like `54/HR` are read as £54.
- ACS cells listing several assets (`A1459741 + A1156216`, `A1, A2, A3`) are split on `+ , ; &` and each part counts as its own code (parts must contain a digit).
- PDF supplier reports are not supported. They must be Excel.

## Web app

- https://hire-checker.vercel.app: public, works on any device including phones, no login. Also published privately on claude.ai: https://claude.ai/artifact/2DR6XUR3QU2BZEJmDcWLpH (republish it after each change).
- GitHub: Pearse13/hire-checker (public). Pushing to `main` redeploys the Vercel site in about a minute.
- Files are read in the viewer's browser and never uploaded.

### How the page should look and work
- Yellow "HIRE CHECK" header with hazard stripe, then nothing else above the upload boxes: no intro text, no notes at the bottom of the page.
- Starts empty. Before a check has run, the step 1 (ACS report) and step 2 (supplier reports) boxes sit in the middle of the screen, with their text and buttons centred. Once both are added they move to the top.
- All buttons are solid blue with white text.
- After a check: the **Suggestions** area (plain text: plant codes not on the ACS report, then the £0 end pass), then **Step 3: Marked-up spreadsheets** (ACS copy first, one download per supplier, plus "Download all"), then one tab per supplier with totals, "charged but not on ACS", "on ACS but not charged", "on both reports" (with the ACS cell) and the no-charge lines folded away.
- Each supplier tab has a plant-code column picker, the web version of `--code-col`.
- Downloads use Claude's download prompt on claude.ai and an ordinary browser download on the Vercel site.

### Code
- Source is `src/engine.js` (matching, suggestions, mark-up), `src/app.js` (screens) and `src/page.html`. Run `python3 build.py` to rebuild `web/index.html` after editing them. Never edit `web/index.html` by hand.
- `hire_check.py` and `src/engine.js` must give the same answers. Change both together.
- Spreadsheets are never committed or deployed (`.gitignore`, `.vercelignore`). They hold real supplier contacts and names.
- The web app reads files with ExcelJS, which treats `<b val="0"/>` ("bold off") as bold. `cleanXlsx` in `src/engine.js` strips those before reading, or every unmatched supplier line comes out bold. Keep it.
- To check the two versions agree, run both on the same files and compare the `_checked.xlsx` files cell by cell (values, bold, fill, borders).

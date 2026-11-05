# ACS Hire Report Check

Weekly check of what plant hire suppliers are charging ACS against the ACS in-house hire report.

## The job

1. **ACS in-house hire report**: one Excel workbook, one sheet per job (plus sheets like Yard and Vehicles). Rows are already colour-coded by me (green on hire, amber/red, key for off-hired and suspended): keep those colours.
2. **Supplier hire reports**: one Excel file per supplier, showing what we're being charged for.
3. For every line on each supplier report, look up its **plant code** on all sheets of the ACS report.
   - **Found**: highlight that supplier line green and make it bold, and give the ACS sheet and cell it matched (e.g. `M3 J9!C16`) in an added "ACS match" column. The matching ACS line goes green and bold too.
   - In the web app, nothing is marked until I approve it: I click through every pairing one at a time and approve or reject it.
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
- Also saves `<ACS report>_checked.xlsx`: every ACS plant line being charged on the supplier reports in **green fill + bold** (only the report's own columns). The green replaces that line's own colour; every other line keeps my colours. A line listing several assets is only marked when all of them are charged.
- Prints the result on screen for each supplier: either OK with the match count, or every plant code NOT on the ACS report, with its row number and the full charged line, plus a grand total.
- Then prints **suggestions**, worked out once every report is checked:
  - Each charged line whose code isn't on the ACS report is compared with the ACS lines no supplier has charged for: on-hire date, rate (also rate × quantity for "2no" lines), quantity, description words and sizes (4m, 350mm, 8 CFM). A suggestion needs the description to agree plus the date or rate, or both date and rate. Sizes that disagree (4m vs 5m) rule a pair out.
  - **£0 end pass**: ACS lines at £0 (or blank rate) that are still unpaired are looked for among the supplier lines nothing else claimed. These need a shared description word (a ladder is a ladder).
  - In the script, suggestions are text only and nothing is marked for them. In the web app, each suggestion with an ACS line is a pairing to approve, and is marked like a match once approved.
  - A charged line with no suggestion that went on hire after the ACS report's date (read from its filename, e.g. `01.09.26`) is flagged as probably new plant to add.
- "On ACS but not charged" (point 4 above) is in the web app's second look only. The script doesn't print it yet.

## Rules

- **Never edit the original files.** Output is always a separate `_checked.xlsx` copy.
- Supplier copies: **green fill + bold on matched lines** and the one added "ACS match" column. ACS copy: **green fill + bold** on charged lines, every other line's colours untouched. No borders, no new sheets, no red marking, no deleted rows, no other added columns (apart from the web app's Outcome column, below).
- Web app: only **approved** pairings are marked, on both copies. Rejected and skipped ones are left alone.
- Web app: second-look outcomes go in an added **Outcome** column on both copies, on the lines they belong to. Supplier copy: after "ACS match". ACS copy: the first empty column after the report's own, so notes beside it aren't overwritten. It's only added to sheets that have an outcome, and writing one doesn't mark a line green. The script has no outcomes, so its copies never have this column.
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
- Files are read in the viewer's browser. The only upload is "Open in Excel Online", which puts a marked-up copy in the viewer's own OneDrive. Nothing goes to any other server.

### How the page should look and work
- Yellow "HIRE CHECK" header with hazard stripe, then nothing else above the upload boxes: no intro text, no notes at the bottom of the page.
- Starts empty. Before a check has run, the step 1 (ACS report) and step 2 (supplier reports) boxes sit in the middle of the screen, with their text and buttons centred. Once both are added they move to the top and a **Run check** button appears. Nothing is checked until I press it. Adding another file hides the results until I press it again.
- Keep the explanatory text short: one line at most per section.
- All buttons are solid blue with white text.
- After Run check, the pairing card comes first (upload boxes shrink to one line each), above everything else. Under it is the current pairing's supplier tab, which updates live as I go: approved lines turn green and the line on the card is outlined. Suggestions and downloads stay hidden until every pairing has an answer.
- The card's title is the ACS job (the sheet the line is on, e.g. M3 J9), with the count (4 of 82) beside it. Under it is a four-row table, Supplier next to ACS: **Item** (item name, always, with the plant code / asset number under it), **Cost** (always), **Company** and **On hire**. No ticks, crosses or red in this first click-through. Keep it this short: no explanations, reasons or extra fields on the card.
- An ACS line listing several assets (a kit, "4no Radios and Charging Station"), or one plant code charged on several supplier lines (a hose split into 4 lines), is **one card**, titled with the job and the ACS description. Lines that pair one-to-one keep the normal four-row card. It shows a **running tally**: every supplier line charged for those assets with its cost and the total so far, ending on a Tally row (£90.00 = ACS £90.00 ✓, or ≠ and red). Company and on-hire date follow as two short rows. Assets the supplier hasn't charged show as "Not charged". A suggestion for one of its assets joins that card (tagged "Different code") instead of being its own. A code on several such lines (a charger in each radio set) gives one supplier line to each ACS line. The ACS line is only marked when every asset is covered.
- Buttons: Approve and Reject (big, keys A and R), Back and Skip (keys ← and →). Order: plant-code matches, then suggestions (tagged "Different code"), then £0 end pass finds. Decisions are kept in the browser per ACS report, so a reload doesn't lose them.
- **Second look**: when every pairing is answered, the next click-through starts, one card at a time. First the pairings with something wrong (cost, dates, company or sizes that disagree, or a part not charged), unless I already rejected them: each row now gets ✓ or red ✗, and I can approve or reject again. Then what nothing paired with: charged but not on ACS (with "probably new plant" if it went on hire after the report date), on ACS but not charged, and £0 ACS lines not found on a supplier report. These are not lists anywhere on the page.
- Every second-look card has outcome buttons saying what happened: Off-hired, Moved site, Loss charge applied, Query with supplier (plus New plant on charged lines), keys 1 to 5, and Other for a typed note. Back and Skip as well. On a pairing an outcome keeps its approval as it was. Outcomes are counted when the check finishes and kept in the browser with the decisions. A chosen outcome shows on its button when I go back.
- When the second look is finished, both marked-up copies download by themselves (only approved pairings marked), and **Step 4: Marked-up spreadsheets** appears (ACS copy first, one download per supplier, plus "Download all"). Under it is one tab per supplier with totals, "on both reports" (with the ACS cell and my Approved / Rejected / Waiting check) and the no-charge lines folded away.
- Each supplier tab has a plant-code column picker, the web version of `--code-col`.
- Downloads use Claude's download prompt on claude.ai and an ordinary browser download on the Vercel site.
- Excel Online (Vercel site only): "Sign in to Microsoft" in Step 4, then each copy gets an "Open in Excel Online" button. It saves to `Hire Checks/<ACS report>/` in my OneDrive and opens in a new tab; pressing it again replaces that copy. The account is my personal Microsoft account (`tenantId: "consumers"`). The buttons stay hidden until `MS.clientId` is set in `src/app.js` (from the Microsoft app registration, see README). `web/auth.html` is the sign-in return page.

### Code
- Source is `src/engine.js` (matching, suggestions, mark-up), `src/app.js` (screens) and `src/page.html`. Run `python3 build.py` to rebuild `web/index.html` after editing them. Never edit `web/index.html` by hand.
- `hire_check.py` and `src/engine.js` must give the same answers. Change both together.
- Spreadsheets are never committed or deployed (`.gitignore`, `.vercelignore`). They hold real supplier contacts and names.
- The web app reads files with ExcelJS, which treats `<b val="0"/>` ("bold off") as bold. `cleanXlsx` in `src/engine.js` strips those before reading, or every unmatched supplier line comes out bold. Keep it.
- To check the two versions agree, run both on the same files and compare the `_checked.xlsx` files cell by cell (values, bold, fill, borders).

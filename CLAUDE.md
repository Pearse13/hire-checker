# ACS Hire Report Check

Weekly check of what plant hire suppliers are charging ACS against the ACS in-house hire report.

## The job

1. **ACS in-house hire report**: one Excel workbook, one sheet per job (currently 3 jobs).
2. **Supplier hire reports**: one Excel file per supplier, showing what we're being charged for.
3. For every line on each supplier report, look up its **plant code** on all sheets of the ACS report.
   - **Found**: highlight that supplier line green and make it bold.
   - **Not found**: notify me. Do not mark, add or remove anything for it.

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
- Saves the marked-up copy as `<supplier>_checked.xlsx` next to the original.
- Also saves `<ACS report>_checked.xlsx`: every ACS plant line being charged on the supplier reports in green + bold (only the report's own columns). A line listing several assets is only marked when all of them are charged.
- Prints the result on screen for each supplier: either OK with the match count, or every plant code NOT on the ACS report, with its row number and the full charged line, plus a grand total.

## Rules

- **Never edit the original files.** Output is always a separate `_checked.xlsx` copy.
- The only change made to any copy is **green fill + bold on matched lines**. No new sheets, no red marking, no deleted rows, no added columns.
- Unmatched plant is **reported, not marked**.
- Don't change the matching logic without asking me first.

## When reporting back to me

Keep it short. For each supplier, give either "all on ACS report" or the list of plant codes not on the ACS report (code, description, charge). End with the total count. Don't recap the steps.

## Known limits

- Old `.xls` files need LibreOffice to convert. If it's not installed, open the file in Excel and use Save As `.xlsx` first.
- Plant codes containing spaces won't be recognised as codes.
- ACS cells listing several assets (`A1459741 + A1156216`, `A1, A2, A3`) are split on `+ , ; &` and each part counts as its own code (parts must contain a digit).
- PDF supplier reports are not supported. They must be Excel.

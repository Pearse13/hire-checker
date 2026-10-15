# Hire Checker

Weekly check of what plant hire suppliers are charging ACS against the ACS on-hire report.

## Web app

`web/index.html`: open it, drop in the ACS on-hire report and the supplier hire reports. Files are read in the browser and never uploaded. It shows, for each supplier:

- charged lines not on the ACS report
- ACS plant from that supplier that isn't being charged
- lines on both reports, grouped by job, with the ACS cell each matched

Then, once everything is checked, suggestions in plain text: what each unknown plant code probably is on the ACS report (by on-hire date, rate, quantity and description), and where the ACS report's unpaired £0 items are on the supplier reports.

It downloads marked-up copies: each supplier report with matched lines in green and bold plus an "ACS match" column, and the ACS report with plant being charged in bold with a thick border (its own colours kept).

Edit `src/`, then run `python3 build.py` to rebuild `web/index.html`.

## Script

```
pip3 install openpyxl
python3 hire_check.py <ACS_hire_report.xlsx> <supplier1.xlsx> [<supplier2.xlsx> ...] [--code-col C]
```

Writes `<file>_checked.xlsx` copies next to the originals and prints every plant code not on the ACS report. See `CLAUDE.md` for the full rules.

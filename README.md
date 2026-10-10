# Hire Checker

Weekly check of what plant hire suppliers are charging ACS against the ACS on-hire report.

## Web app

`web/index.html`: open it, drop in the ACS on-hire report and the supplier hire reports. Files are read in the browser and never uploaded. It shows, for each supplier:

- charged lines not on the ACS report (with a possible ACS match when the asset number looks swapped)
- ACS plant from that supplier that isn't being charged
- lines on both reports, grouped by job

and downloads marked-up copies: matched lines in green and bold on each supplier report, and plant being charged in green and bold on the ACS report.

## Script

```
pip3 install openpyxl
python3 hire_check.py <ACS_hire_report.xlsx> <supplier1.xlsx> [<supplier2.xlsx> ...] [--code-col C]
```

Writes `<file>_checked.xlsx` copies next to the originals and prints every plant code not on the ACS report. See `CLAUDE.md` for the full rules.

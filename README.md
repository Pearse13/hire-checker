# Hire Checker

Weekly check of what plant hire suppliers are charging ACS against the ACS on-hire report.

## Web app

`web/index.html`: open it, drop in the ACS on-hire report and the supplier hire reports. Files are read in the browser and never uploaded. It shows, for each supplier:

- charged lines not on the ACS report
- ACS plant from that supplier that isn't being charged
- lines on both reports, grouped by job, with the ACS cell each matched

Then, once everything is checked, suggestions in plain text: what each unknown plant code probably is on the ACS report (by on-hire date, rate, quantity and description), and where the ACS report's unpaired £0 items are on the supplier reports.

Then you check each pairing one at a time (every plant-code match, plus each suggestion that has an ACS line) on a short card: item, cost, company and on-hire date side by side, ticked where they agree. Approve or reject it. It downloads marked-up copies where only the approved pairings are marked: green and bold on both the supplier report (plus an "ACS match" column) and the ACS report. After the pairings, a second look goes through anything that needs it, and what you say happened goes in an "Outcome" column.

Edit `src/`, then run `python3 build.py` to rebuild `web/index.html`.

## Script

```
pip3 install openpyxl
python3 hire_check.py <ACS_hire_report.xlsx> <supplier1.xlsx> [<supplier2.xlsx> ...] [--code-col C]
```

Writes `<file>_checked.xlsx` copies next to the originals and prints every plant code not on the ACS report. The script marks every plant-code match without asking; approving pairings one by one is in the web app only. See `CLAUDE.md` for the full rules.

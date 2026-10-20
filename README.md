# Jarred Project

A single repo holding the College Football betting model and the data-pull
tools for football and basketball. Each piece keeps its own spreadsheet — the
Excel files are intentionally separate, not merged.

## What's here

### CFB Betting Model (repo root)
The College Football predictive betting model.

- `CFB_Betting_Model.xlsx` — the model workbook.
- `Open CFB Model.bat` — pulls the latest data (once per day) then opens the
  workbook. This is what the desktop shortcut runs.
- `UpdateData.bat` — refreshes the data only, without opening Excel.
- `update_data.py` — the data feed (SP+, FPI, talent, PPA, schedule, etc.)
  from CollegeFootballData.com.
- `update_data_with_extra_stats.py` — the same feed plus extra play-by-play
  columns (red-zone conversions, 20+ yard plays, blocked kicks, kickoff yards).
- `setup_python.bat` — finds or installs a private Python copy for the scripts.
- `bt.html` — saved reference page used alongside the model.

### DataPull/
Combined one-click updaters for both spreadsheets:
`College Football.xlsx` and `College Basketball.xlsx`. See `DataPull/README.txt`.

### Milestone 1 - Football/
Standalone football data-pull tool and spreadsheet (`College Football.xlsx`).

### Milestone 2 - Basketball/
Standalone basketball data-pull tool and spreadsheet (`College Basketball.xlsx`).

### Milestone 3 - KenPom/
The basketball tool from Milestone 2 with the four KenPom columns filled in
automatically (requires a KenPom login — see `Milestone 3 - KenPom/README.txt`).

## Notes

- Nothing updates on its own; the numbers only change when an updater is run.
- Runtime state the updaters create (`.last_update`, `.plays_cache.json`,
  the downloaded `python/` folder, Excel lock files) is git-ignored.

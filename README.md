# Jarred Project

A single repo holding the College Football betting model and the data-pull
tools for football and basketball.

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

### Milestone 1+2 - Football and Basketball/
The football (Milestone 1) and basketball (Milestone 2) data-pull tools in one
folder, sharing a single spreadsheet, `College Sports.xlsx`, with a `CFB` tab
and a `CBB` tab. `Update Football.bat` and `Update Basketball.bat` each refresh
their own tab; `Update All.bat` runs both. See the folder's `README.txt`.

### Milestone 3 - KenPom/
The Milestone 1+2 tool (same `College Sports.xlsx`) with the four KenPom
basketball columns (SOS, Ovr, ORtg, DRtg) filled in automatically from the
free kenpom.com home page — no subscription or login. If kenpom.com can't be
read, it falls back to a copy of that page saved into the folder from a
browser. See `Milestone 3 - KenPom/README.txt`.

## Notes

- Nothing updates on its own; the numbers only change when an updater is run.
- Runtime state the updaters create (`.last_update`, `.plays_cache.json`,
  the downloaded `python/` folder, Excel lock files) is git-ignored.

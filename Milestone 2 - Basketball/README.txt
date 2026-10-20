===========================================================================
 MILESTONE 2  -  COLLEGE BASKETBALL DATA PULL
===========================================================================

WHAT'S IN THIS FOLDER
  College Basketball.xlsx   your basketball sheet, 365 teams, in your layout
  Update Basketball.bat     double-click this to refresh the numbers
  (the other files are the engine - leave them alone)

HOW TO UPDATE
  1. Close College Basketball.xlsx if it's open in Excel.
  2. Double-click "Update Basketball.bat".
  3. A black window shows progress and says "Done." when finished.
     (The very first time it spends a minute setting itself up - that's
      normal and only happens once.)
  4. Open College Basketball.xlsx.

  It only runs when you click it. Nothing updates on its own.

WHAT GETS FILLED IN
  ESPN ....... SOR, SOS, BPI, BPI Off, BPI Def
  Barttorvik . Roster rating, ADJOE, ADJDE

  The four KENPOM columns (SOS, Ovr, ORtg, DRtg) are left for you to type
  in by hand for now - the updater does NOT touch them, so whatever you
  enter stays put. Automating those is the next milestone, once you've got
  a KenPom login.

  Your team names (column A) and all formatting are never changed.

NOTES
  * Pre-season: until ESPN posts the new season's BPI it still serves last
    season's final BPI; the updater switches over automatically once ESPN
    publishes. Barttorvik pre-season numbers are already live.
  * West Florida: ESPN doesn't carry a BPI for them yet, so their ESPN
    columns stay blank for now (Barttorvik still fills in).
  * The update window lists any team it couldn't find data for, at the end.
===========================================================================

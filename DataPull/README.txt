===========================================================================
 COLLEGE FOOTBALL & BASKETBALL  -  DATA PULL
===========================================================================

WHAT THIS IS
  Two spreadsheets and two one-click updaters.  Nothing updates on its own -
  the numbers only change when YOU run an updater.

  College Football.xlsx      <- your football sheet (138 teams)
  College Basketball.xlsx     <- your basketball sheet (365 teams)

  Your team names (the left-hand column) and the layout/formatting are never
  changed.  The updaters only fill in the stat columns.


HOW TO UPDATE  (do this whenever you want fresh numbers)
  1. Close the spreadsheet in Excel if it's open.
  2. Double-click:
        "Update Football.bat"     to refresh College Football.xlsx
        "Update Basketball.bat"   to refresh College Basketball.xlsx
  3. A black window shows the progress and says "Done." when finished.
     (The very first time, it spends a minute setting up - that's normal and
      only happens once.)
  4. Open the spreadsheet.

  Run them as often or as rarely as you like - they do nothing until clicked.


WHERE EACH NUMBER COMES FROM
  FOOTBALL
    ESPN SOR / SOS / FPI, SP+ (Ovr/Off/Def/ST), PPA, Yards-per-play, Havoc,
    Success Rate, Turnover Margin ....... CollegeFootballData.com
    Control Rating (column S) ........... mcillecesports.com  (see note below)

  BASKETBALL
    Roster rating, ADJOE, ADJDE ......... Barttorvik (barttorvik.com)
    ESPN SOR / SOS / BPI / BPI Off/Def .. ESPN
    KenPom SOS / Ovr / ORtg / DRtg ...... KenPom  (see note below)


CONTROL RATING  (football, column S)
  Pulled automatically from mcillecesports.com for the most recent week shown
  there.  Any team that did NOT play that week is set to 25, as you asked.
  The update window prints which teams got the 25 so you can eyeball them.
  If you'd rather type this column in yourself, open config.py and set
  AUTO_CONTROL = False - the updater will then leave column S alone.


KENPOM  (basketball columns F, H, J, M)
  KenPom has no public data feed and sits behind a paid login, so these four
  columns are left for you to fill in by hand - the updater does NOT touch
  them, so whatever you type stays put.
  If you have a KenPom subscription and want them pulled automatically too:
     1. open config.py and put your KenPom email + password in
     2. one-time: run  "Update Basketball.bat"  after installing the helper:
        open the black window once and run:  pip install kenpompy
  (Leave the login blank to keep these columns manual.)


A FEW THINGS TO KNOW
  * Pre-season basketball: until ESPN publishes the new season's BPI, it
    still serves last season's final BPI.  The updater automatically switches
    to the new numbers the moment ESPN posts them - no change needed here.
  * West Florida: ESPN doesn't carry a BPI for them yet, so their ESPN columns
    stay blank until ESPN adds them.  (Barttorvik numbers still fill in.)
  * If a team ever comes back blank, it means that source spelled the team
    differently; tell Max and it's a one-line fix.
  * The update window lists any team it couldn't find data for, at the end.


SETTINGS  (config.py - only if you need them)
  CFBD_API_KEY   the free CollegeFootballData key (already set)
  CFB_YEAR       football season year (blank = current automatically)
  CBB_SEASON     basketball season (blank = current automatically)
  AUTO_CONTROL   True = pull Control Rating; False = you type column S
  KENPOM_EMAIL / KENPOM_PASSWORD   optional, for auto KenPom (see above)
===========================================================================

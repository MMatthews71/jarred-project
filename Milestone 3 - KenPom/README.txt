===========================================================================
 MILESTONE 3  -  KENPOM AUTOMATION  (football + basketball, final step)
===========================================================================

WHAT THIS IS
  The same tool as Milestone 1 + 2, now with the four KenPom basketball
  columns filled in automatically as well - for FREE (no KenPom
  subscription or login needed):

      KenPom SOS, Overall Rating, Offense (ORtg), Defense (DRtg)

  Use this folder in place of your Milestone 1 + 2 folder - everything else
  works exactly the same.  (If you typed anything into your old
  College Sports.xlsx that you want to keep, copy that file over the one in
  this folder.)

  One spreadsheet with two tabs, and one-click updaters that fill it in.
  Nothing updates on its own - the numbers only change when YOU run an
  updater.

  College Sports.xlsx
     CFB tab   <- your football sheet (138 teams)
     CBB tab   <- your basketball sheet (365 teams)

  Your team names (the left-hand column) and the layout/formatting are never
  changed.  The updaters only fill in the stat columns, and each one only
  touches its own tab - updating football leaves basketball alone and vice
  versa.


HOW TO UPDATE  (do this whenever you want fresh numbers)
  1. Close College Sports.xlsx in Excel if it's open.
  2. Double-click:
        "Update All.bat"          to refresh both tabs
        "Update Football.bat"     to refresh just the CFB tab
        "Update Basketball.bat"   to refresh just the CBB tab
  3. A black window shows the progress and says "Done." when finished.
     (The very first time, it spends a minute setting up - that's normal and
      only happens once.)
  4. Open College Sports.xlsx.

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


CONTROL RATING  (football, CFB tab column S)
  Pulled automatically from mcillecesports.com for the most recent week shown
  there.  Any team that did NOT play that week is set to 25, as you asked.
  The update window prints which teams got the 25 so you can eyeball them.
  If you'd rather type this column in yourself, open config.py and set
  AUTO_CONTROL = False - the updater will then leave column S alone.


KENPOM  (basketball, CBB tab columns F, H, J, M)  -  FREE
  These come from the free kenpom.com home page - the same ratings table
  anyone can see without logging in:
     F  KenPom SOS ......... "Strength of Schedule - NetRtg"
     H  KenPom Ovr Rating .. "NetRtg"
     J  KenPom Offense ..... "ORtg"
     M  KenPom Defense ..... "DRtg"
  Every row is double-checked (NetRtg must equal ORtg minus DRtg), so if
  KenPom ever rearranges its page, nothing wrong gets written - the columns
  are just left as they were and the window tells you.

  IF THE WINDOW SAYS "KenPom not updated"
  KenPom sometimes turns away automatic requests.  If it does, the free
  backup takes 10 seconds:
     1. Open  kenpom.com  in your normal web browser.
     2. Press Ctrl+S (Save page), go to THIS folder, type the name  kenpom
        and click Save.  (Leave "Save as type" on "Webpage, Complete" or
        "Webpage, HTML only".)
     3. Run "Update Basketball.bat" (or "Update All.bat") again.
  The updater uses the newest saved page whenever it can't reach the site,
  and warns you if that page is more than a week old - just save a fresh
  one each time you want new KenPom numbers.

  To type these four columns in yourself instead, open config.py and set
  AUTO_KENPOM = False - the updater will then leave them alone.


A FEW THINGS TO KNOW
  * Pre-season basketball: until ESPN publishes the new season's BPI, it
    still serves last season's final BPI.  The updater automatically switches
    to the new numbers the moment ESPN posts them - no change needed here.
    KenPom's home page works the same way.
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
  AUTO_KENPOM    True = pull the KenPom columns; False = you type them
===========================================================================

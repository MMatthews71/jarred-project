"""
config.py  -  settings for the College Football / Basketball data pulls
=======================================================================

Edit the values below if needed, then save.  Nothing else in here needs
changing from week to week - you just run the updater (see README.txt).
"""

# ---------------------------------------------------------------------------
# CollegeFootballData.com API key  (free, needed for the FOOTBALL pull only)
# Get your own instantly at:  https://collegefootballdata.com/key
# ---------------------------------------------------------------------------
CFBD_API_KEY = "MUUUNFo1ICHMRpyBwImFWe6SKjJMMST6Zd9gmP2QHf/Wc+ykaj54K/0F6ECZNXfL"

# ---------------------------------------------------------------------------
# Seasons.  Leave as None to use the current season automatically.
#   CFB_YEAR     e.g. 2026  (the calendar year the season starts in)
#   CBB_SEASON   e.g. 2027  (ESPN labels 2026-27 as 2027)
# ---------------------------------------------------------------------------
CFB_YEAR = None
CBB_SEASON = None

# ---------------------------------------------------------------------------
# Control Rating (football, column S) - source: mcillecesports.com
#   AUTO_CONTROL = True   -> pull it automatically each run; teams that did
#                            not play that week are set to 25 (your rule).
#   AUTO_CONTROL = False  -> leave column S exactly as you typed it (manual).
# ---------------------------------------------------------------------------
AUTO_CONTROL = True
CONTROL_BYE_VALUE = 25

# ---------------------------------------------------------------------------
# KenPom (basketball columns F, H, J, M) - source: the FREE kenpom.com home
# page (no subscription or login needed).
#   AUTO_KENPOM = True   -> fill SOS / Ovr / ORtg / DRtg each run.  If the
#                           site can't be read, the updater uses a copy of the
#                           page you saved yourself (see README.txt), and if
#                           there isn't one either, leaves the columns as-is.
#   AUTO_KENPOM = False  -> leave those four columns exactly as you typed them.
# ---------------------------------------------------------------------------
AUTO_KENPOM = True
KENPOM_SAVED_PAGE = "kenpom.html"     # the saved-page fallback, in this folder

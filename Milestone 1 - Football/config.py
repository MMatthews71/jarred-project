

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
# KenPom (basketball columns F, H, J, M).  KenPom has NO public API and needs
# a paid login, so these four columns are left for you to type in by hand by
# default.  If you have a KenPom subscription and want them pulled too, put
# your login here AND install the helper once:   pip install kenpompy
# Leaving these blank keeps those columns manual (the updater won't touch them).
# ---------------------------------------------------------------------------
KENPOM_EMAIL = ""
KENPOM_PASSWORD = ""

#!/usr/bin/env python3
"""
update_basketball.py  -  fill the CBB tab of 'College Sports.xlsx' with the latest numbers
==========================================================================================

Run this (double-click "Update Basketball.bat", or: python update_basketball.py)
whenever you want fresh data.  It never runs on its own.

It writes only the stat columns.  Your team names in column A and all of the
sheet's formatting are left exactly as they are.

Columns and sources:
  C  Roster (Barttorvik Talent) ....... Barttorvik
  D  ESPN SOR .......................... ESPN
  E  ESPN SOS .......................... ESPN
  F  KenPom SOS ........................ KenPom  (manual unless login set - see config.py)
  G  ESPN BPI .......................... ESPN
  H  KenPom Ovr Rating ................. KenPom  (manual unless login set)
  I  ESPN BPI Off ...................... ESPN
  J  KenPom Offense (ORtg) ............. KenPom  (manual unless login set)
  K  Barttorvik Off (ADJOE) ............ Barttorvik
  L  ESPN BPI Def ...................... ESPN
  M  KenPom Defense (DRtg) ............. KenPom  (manual unless login set)
  N  Barttorvik Def (ADJDE) ............ Barttorvik

Note: in the pre-season ESPN still serves last season's final BPI until they
publish the new season's numbers; this updater automatically picks up the new
figures as soon as ESPN posts them.
"""

import os
import sys
from datetime import datetime

# The embedded Python that setup_python.bat installs uses a ._pth file, which
# stops Python from adding this script's own folder to the import path. Add it
# back so the local config.py / common.py next to this file are found.
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import config
from common import (requests, TeamMatcher, open_sheet, sheet_rows, write_values,
                    save_sheet, report, barttorvik_get, parse_tables)
from openpyxl.utils import column_index_from_string

SEASON = config.CBB_SEASON or (datetime.now().year + (1 if datetime.now().month >= 5 else 0))
WORKBOOK = os.path.join(os.path.dirname(os.path.abspath(__file__)), "College Sports.xlsx")
SHEET = "CBB"
_UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
       "(KHTML, like Gecko) Chrome/129.0 Safari/537.36")

# sheet name -> ESPN abbreviation, only where normal matching can't find it.
CBB_ESPN_ALIAS = {
    "Albany": "UALB", "Boston University": "BU", "Cal": "CAL", "Cal Baptist": "CBU",
    "Citadel": "CIT", "Illinois Chicago": "UIC", "K-State": "KSU", "Miami (FL)": "MIA",
    "Mizzou": "MIZ", "N. Carolina": "UNC", "New Orleans": "NOLA", "NM State": "NMSU",
    "OK State": "OKST", "Sac State": "SAC", "Seattle": "SEA", "UMKC": "KC",
    "USC Upstate": "UPST", "UTRGV": "RGV", "VA Tech": "VT", "Wash St.": "WSU",
    "Nebraska Omaha": "OMA", "Southeastern LA": "SELA", "TAMU Corpus Christi": "AMCC",
    "S. Carolina": "SC", "S. Carolina St.": "SCST", "S. Alabama": "USA",
    "Tennessee Martin": "UTM", "UL-Lafayette": "UL", "Central Mich": "CMU",
    "Eastern Mich": "EMU", "Eastern Wash": "EWU", "Grambling St.": "GRAM",
    "McNeese St.": "MCN", "MD Eastern Shore": "UMES", "Nicholls St.": "NICH",
    "Sam Houston St.": "SHSU", "San Jose St.": "SJSU", "Western Mich": "WMU",
    "West Florida": "__NONE__",   # not carried by ESPN
}

# sheet name -> the name Barttorvik uses, only where they differ.
CBB_BT_ALIAS = {
    "AR-Pine Bluff": "Arkansas Pine Bluff", "App State": "Appalachian St.",
    "Cal": "California", "Central Mich": "Central Michigan", "Citadel": "The Citadel",
    "ETSU": "East Tennessee St.", "Eastern Mich": "Eastern Michigan",
    "Eastern Wash": "Eastern Washington", "FAU": "Florida Atlantic",
    "GA Southern": "Georgia Southern", "JMU": "James Madison",
    "Jax State": "Jacksonville St.", "K-State": "Kansas St.",
    "MD Eastern Shore": "Maryland Eastern Shore", "MTSU": "Middle Tennessee",
    "MVSU": "Mississippi Valley St.", "Mizzou": "Missouri", "N. Carolina": "North Carolina",
    "NC Central": "North Carolina Central", "NDSU": "North Dakota St.",
    "NIU": "Northern Illinois", "NM State": "New Mexico St.", "OK State": "Oklahoma St.",
    "Ole Miss": "Mississippi", "PITT": "Pittsburgh", "S. Alabama": "South Alabama",
    "S. Carolina": "South Carolina", "S. Carolina St.": "South Carolina St.",
    "SEMO": "Southeast Missouri St.", "Sac State": "Sacramento St.",
    "Southeastern LA": "Southeastern Louisiana",
    "TAMU Corpus Christi": "Texas A&M Corpus Chris", "UConn": "Connecticut",
    "UL-Lafayette": "Louisiana", "UL-Monroe": "Louisiana Monroe",
    "UTRGV": "UT Rio Grande Valley", "VA Tech": "Virginia Tech",
    "WKU": "Western Kentucky", "Wash St.": "Washington St.",
    "Western Mich": "Western Michigan",
}


def pull_espn():
    """{espn_id: {stat: value}} with team-name candidates for matching."""
    teams = {}
    page = 1
    while True:
        url = (f"https://site.api.espn.com/apis/fitt/v3/sports/basketball/"
               f"mens-college-basketball/powerindex?season={SEASON}&limit=50&page={page}")
        r = requests.get(url, headers={"User-Agent": _UA}, timeout=60)
        r.raise_for_status()
        d = r.json()
        for t in d["teams"]:
            cat = {c["name"]: c.get("values") or [] for c in t["categories"]}
            bpi = cat.get("bpi", [])
            res = cat.get("resume", [])

            def g(arr, i):
                return arr[i] if len(arr) > i else None

            teams[t["team"]["id"]] = {
                "team": t["team"],
                "sor": g(res, 0),          # SOR rank
                "sos": g(res, 5),          # SOS rank
                "bpi": g(bpi, 0),
                "bpi_off": g(bpi, 3),
                "bpi_def": g(bpi, 4),
            }
        if page >= d["pagination"]["pages"]:
            break
        page += 1
    return teams


def _num(x):
    try:
        return float(str(x).replace(",", "").strip())
    except (TypeError, ValueError):
        return None


def pull_barttorvik():
    """{bt_team_name: {'talent':, 'adjoe':, 'adjde':}} from the pre-season page."""
    html = barttorvik_get("https://barttorvik.com/trankpre.php")
    out = {}
    for tbl in parse_tables(html):
        for row in tbl:
            # a data row looks like: [rank, Team, Conf, AdjOE, AdjDE, ...]
            if len(row) < 11:
                continue
            adjoe, adjde, talent = _num(row[3]), _num(row[4]), _num(row[10])
            team = (row[1] or "").strip()
            if not team or adjoe is None or not (60 <= adjoe <= 150):
                continue                      # skip header / non-data rows
            out[team] = {"talent": talent, "adjoe": adjoe, "adjde": adjde}
    return out


def pull_kenpom():
    """{kp_team: {'sos':,'ovr':,'ortg':,'drtg':}} or {} if no login/available."""
    if not (config.KENPOM_EMAIL and config.KENPOM_PASSWORD):
        return {}, "manual"
    try:
        import kenpompy.utils as kpu
        import kenpompy.summary as kps
        from kenpompy.misc import get_pomeroy_ratings
        browser = kpu.login(config.KENPOM_EMAIL, config.KENPOM_PASSWORD)
        df = get_pomeroy_ratings(browser, season=str(SEASON))
        out = {}
        for _, r in df.iterrows():
            out[str(r.get("Team", "")).strip()] = {
                "ovr": _num(r.get("AdjEM")),
                "ortg": _num(r.get("AdjO")),
                "drtg": _num(r.get("AdjD")),
                "sos": _num(r.get("AdjEM.1")),     # SOS AdjEM column
            }
        return out, "ok"
    except Exception as e:
        return {}, f"error: {e}"


def main():
    print(f"College Basketball - pulling {SEASON} data ...")
    wb, ws = open_sheet(WORKBOOK, SHEET)
    rows = sheet_rows(ws)
    total = 0

    # ---- ESPN ----
    print("  ESPN BPI / SOR / SOS ...")
    espn = pull_espn()
    em = TeamMatcher(rows.keys(), alias=CBB_ESPN_ALIAS)
    for _id, rec in espn.items():
        tm = rec["team"]
        em.add(_id, [tm.get("shortDisplayName"), tm.get("nickname"), tm.get("displayName")],
               abbr=tm.get("abbreviation"), payload=rec)
    emap, emiss = em.sheet_to_source()
    for letter, field, nd in [("D", "sor", 0), ("E", "sos", 0), ("G", "bpi", 2),
                              ("I", "bpi_off", 2), ("L", "bpi_def", 2)]:
        vals = {t: d.get(field) for t, d in emap.items()}
        total += write_values(ws, rows, column_index_from_string(letter), vals, nd)

    # ---- Barttorvik ----
    print("  Barttorvik Talent / ADJOE / ADJDE ...")
    try:
        bt = pull_barttorvik()
    except Exception as e:
        print(f"  ! Barttorvik unavailable ({e}) - columns C/K/N left unchanged.")
        bt, bmiss = {}, []
    if bt:
        bm = TeamMatcher(rows.keys(), alias=CBB_BT_ALIAS)
        for name, rec in bt.items():
            bm.add(name, [name], payload=rec)
        bmap, bmiss = bm.sheet_to_source()
        for letter, field, nd in [("C", "talent", 1), ("K", "adjoe", 1), ("N", "adjde", 1)]:
            vals = {t: d.get(field) for t, d in bmap.items()}
            total += write_values(ws, rows, column_index_from_string(letter), vals, nd)
    else:
        bmiss = []

    # ---- KenPom (optional) ----
    kp, status = pull_kenpom()
    if status == "manual":
        print("  KenPom: columns F/H/J/M left for manual entry (no login in config.py).")
        kmiss = []
    elif kp:
        km = TeamMatcher(rows.keys())
        for name, rec in kp.items():
            km.add(name, [name], payload=rec)
        kmap, kmiss = km.sheet_to_source()
        for letter, field, nd in [("F", "sos", 2), ("H", "ovr", 2), ("J", "ortg", 1), ("M", "drtg", 1)]:
            vals = {t: d.get(field) for t, d in kmap.items()}
            total += write_values(ws, rows, column_index_from_string(letter), vals, nd)
        print(f"  KenPom: pulled for {len(kmap)} teams.")
    else:
        print(f"  KenPom: not updated ({status}) - columns F/H/J/M left as-is.")
        kmiss = []

    save_sheet(wb, WORKBOOK)
    print(f"\nDone.  Wrote {total} stat values across {len(rows)} teams.")
    report("ESPN", emiss)
    report("Barttorvik", bmiss)


if __name__ == "__main__":
    main()

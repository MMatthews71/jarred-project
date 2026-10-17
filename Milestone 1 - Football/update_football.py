#!/usr/bin/env python3
"""
update_football.py  -  fill 'College Football.xlsx' with the latest numbers
===========================================================================

Run this (double-click "Update Football.bat", or: python update_football.py)
whenever you want fresh data.  It never runs on its own.

It writes only the stat columns (C-S).  Your team names in column A and all
of the sheet's formatting are left exactly as they are.  Any team it could
not find data for is listed at the end.

Sources:
  * CollegeFootballData.com  - SP+, PPA, yards/play, havoc, success rate,
                               turnover margin, ESPN FPI / SOR / SOS
  * mcillecesports.com       - Control Rating (column S)
"""

import os
import re
import sys
from datetime import datetime

# The embedded Python that setup_python.bat installs uses a ._pth file, which
# stops Python from adding this script's own folder to the import path. Add it
# back so the local config.py / common.py next to this file are found.
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import config
from common import (requests, TeamMatcher, open_sheet, sheet_rows, write_values,
                    save_sheet, report, http_get, parse_tables, norm)

YEAR = config.CFB_YEAR or datetime.now().year
BASE = "https://api.collegefootballdata.com"
WORKBOOK = os.path.join(os.path.dirname(os.path.abspath(__file__)), "College Football.xlsx")
SHEET = "CFB"

# sheet name -> the name CollegeFootballData uses (only where they differ)
CFB_ALIAS = {
    "Cal": "California", "Central Mich": "Central Michigan",
    "Eastern Mich": "Eastern Michigan", "FAU": "Florida Atlantic",
    "FIU": "Florida International", "GA Southern": "Georgia Southern",
    "Jax State": "Jacksonville State", "JMU": "James Madison",
    "K-State": "Kansas State", "LA Tech": "Louisiana Tech", "Miami (FL)": "Miami",
    "Mizzou": "Missouri", "MTSU": "Middle Tennessee", "N. Carolina": "North Carolina",
    "NDSU": "North Dakota State", "NIU": "Northern Illinois",
    "NM State": "New Mexico State", "OK State": "Oklahoma State",
    "PITT": "Pittsburgh", "S. Alabama": "South Alabama", "S. Carolina": "South Carolina",
    "Sac State": "Sacramento State", "San Jose St.": "San José State",
    "UL-Lafayette": "Louisiana", "UMass": "Massachusetts", "VA Tech": "Virginia Tech",
    "Wash St.": "Washington State", "Western Mich": "Western Michigan",
    "WKU": "Western Kentucky",
}

# sheet name -> the name mcillecesports.com uses, ONLY where they differ.
# (mcille spells many teams exactly like the sheet - Cal, JMU, MTSU, UMass -
#  so this is intentionally a different, smaller list than CFB_ALIAS.)
MCILLE_ALIAS = {
    "Central Mich": "Central Michigan", "Eastern Mich": "Eastern Michigan",
    "FAU": "Florida Atlantic", "GA Southern": "Georgia Southern",
    "LA Tech": "Louisiana Tech", "Miami (FL)": "Miami", "Mizzou": "Missouri",
    "N. Carolina": "North Carolina", "NDSU": "North Dakota State",
    "NM State": "New Mexico State", "PITT": "Pittsburgh", "S. Alabama": "South Alabama",
    "S. Carolina": "South Carolina", "South Florida": "USF", "UL-Lafayette": "Louisiana",
    "UL-Monroe": "ULM", "VA Tech": "Virginia Tech", "WKU": "Western Kentucky",
    "Wash St.": "Washington State", "Western Mich": "Western Michigan",
}


def get(path, **params):
    headers = {"Authorization": f"Bearer {config.CFBD_API_KEY}", "Accept": "application/json"}
    r = requests.get(f"{BASE}{path}", headers=headers, params=params, timeout=60)
    if r.status_code == 401:
        sys.exit("CollegeFootballData rejected the API key (401).  Check CFBD_API_KEY in config.py.")
    if r.status_code == 429:
        sys.exit("CollegeFootballData monthly call limit used up (429).  Try later or use your own key.")
    r.raise_for_status()
    return r.json()


def pull_cfbd():
    """Return {cfbd_team_name: {field: value}} for every FBS team."""
    data = {}

    def slot(t):
        return data.setdefault(t, {})

    for row in get("/ratings/fpi", year=YEAR):
        t = row.get("team")
        if not t:
            continue
        d = slot(t)
        d["fpi"] = row.get("fpi")
        rr = row.get("resumeRanks") or {}
        d["sor"] = rr.get("strengthOfRecord")
        d["sos"] = rr.get("strengthOfSchedule")

    for row in get("/ratings/sp", year=YEAR):
        t = row.get("team")
        if not t:
            continue
        d = slot(t)
        d["sp_ovr"] = row.get("rating")
        d["sp_off"] = (row.get("offense") or {}).get("rating")
        d["sp_def"] = (row.get("defense") or {}).get("rating")
        d["sp_st"] = (row.get("specialTeams") or {}).get("rating")

    for row in get("/ppa/teams", year=YEAR):
        d = slot(row.get("team", ""))
        d["off_ppa"] = (row.get("offense") or {}).get("overall")
        d["def_ppa"] = (row.get("defense") or {}).get("overall")

    for row in get("/stats/season/advanced", year=YEAR):
        d = slot(row.get("team", ""))
        off, dfn = row.get("offense") or {}, row.get("defense") or {}
        d["off_success"] = off.get("successRate")
        d["def_success"] = dfn.get("successRate")
        d["off_havoc"] = (off.get("havoc") or {}).get("total")
        d["def_havoc"] = (dfn.get("havoc") or {}).get("total")

    # season totals -> yards/play (off & def) and turnover margin
    season = {}
    for row in get("/stats/season", year=YEAR):
        t = row.get("team", "")
        if row.get("statName") is not None:
            season.setdefault(t, {})[row["statName"]] = row.get("statValue")
    for t, s in season.items():
        d = slot(t)
        off_plays = (s.get("rushingAttempts") or 0) + (s.get("passAttempts") or 0)
        def_plays = (s.get("rushingAttemptsOpponent") or 0) + (s.get("passAttemptsOpponent") or 0)
        if off_plays:
            d["off_ypp"] = (s.get("totalYards") or 0) / off_plays
        if def_plays:
            d["def_ypp"] = (s.get("totalYardsOpponent") or 0) / def_plays
        if s.get("turnovers") is not None and s.get("turnoversOpponent") is not None:
            d["to_margin"] = s["turnoversOpponent"] - s["turnovers"]   # takeaways - giveaways

    return data


def pull_control():
    """{mcille_team: control_value} for the most recent week on the site,
    plus that week number.  Returns ({}, None) on any failure."""
    try:
        html = http_get("https://mcillecesports.com/game-ratings/")
    except Exception as e:
        print(f"  ! could not reach mcillecesports.com ({e}) - Control left unchanged.")
        return None, None
    best = {}
    for tbl in parse_tables(html):
        if not tbl or len(tbl[0]) < 6:
            continue
        head = [h.lower() for h in tbl[0]]
        try:
            wk_i = next(i for i, h in enumerate(head) if h.startswith("week"))
            tm_i = next(i for i, h in enumerate(head) if h == "team")
            ct_i = next(i for i, h in enumerate(head) if h.startswith("control"))
        except StopIteration:
            continue
        for row in tbl[1:]:
            if len(row) <= max(wk_i, tm_i, ct_i):
                continue
            try:
                wk = int(re.sub(r"[^\d]", "", row[wk_i]))
            except ValueError:
                continue
            team = row[tm_i].strip()
            try:
                ctrl = float(re.sub(r"[^\d.\-]", "", row[ct_i]))
            except ValueError:
                continue
            # keep the latest week seen for each team
            if team not in best or wk > best[team][0]:
                best[team] = (wk, ctrl)
    if not best:
        return {}, None
    latest = max(wk for wk, _ in best.values())
    control = {t: c for t, (wk, c) in best.items() if wk == latest}
    return control, latest


def main():
    print(f"College Football - pulling {YEAR} data from CollegeFootballData ...")
    data = pull_cfbd()

    wb, ws = open_sheet(WORKBOOK, SHEET)
    rows = sheet_rows(ws)

    matcher = TeamMatcher(rows.keys(), alias=CFB_ALIAS)
    for team, d in data.items():
        matcher.add(team, [team], payload=d)
    mapping, unmatched = matcher.sheet_to_source()   # {sheet team: cfbd dict}

    # column letter -> (field, decimals)
    COLS = {
        "C": ("sor", 0), "D": ("sos", 0), "E": ("fpi", 1),
        "F": ("sp_ovr", 1), "G": ("to_margin", 0), "H": ("sp_off", 1),
        "I": ("off_ppa", 2), "J": ("off_ypp", 2), "K": ("off_havoc", 3),
        "L": ("off_success", 3), "M": ("sp_def", 1), "N": ("def_ppa", 2),
        "O": ("def_ypp", 2), "P": ("def_havoc", 3), "Q": ("def_success", 3),
        "R": ("sp_st", 1),
    }
    from openpyxl.utils import column_index_from_string
    total = 0
    for letter, (field, nd) in COLS.items():
        vals = {team: d.get(field) for team, d in mapping.items()}
        total += write_values(ws, rows, column_index_from_string(letter), vals, nd)

    # ---- Control Rating (column S) ----
    if config.AUTO_CONTROL:
        control, week = pull_control()
        if control is None:
            pass                                   # unreachable - leave as-is
        else:
            cm = TeamMatcher(rows.keys(), alias=MCILLE_ALIAS)
            for t in control:
                cm.add(t, [t], payload=control[t])
            cmap, cmiss = cm.sheet_to_source()
            s_col = column_index_from_string("S")
            filled = write_values(ws, rows, s_col, cmap)
            # teams with no row this week -> bye -> your default value
            byes = 0
            for team, row in rows.items():
                if team not in cmap:
                    ws.cell(row, s_col, config.CONTROL_BYE_VALUE)
                    byes += 1
            print(f"Control Rating: week {week} - {filled} teams, "
                  f"{byes} set to {config.CONTROL_BYE_VALUE} (no game / not found).")
            if cmiss:
                print("  (verify these got the default correctly:",
                      ", ".join(cmiss[:12]) + (" ..." if len(cmiss) > 12 else "") + ")")

    save_sheet(wb, WORKBOOK)
    print(f"\nDone.  Wrote {total} stat values for {len(mapping)} of {len(rows)} teams.")
    report("CollegeFootballData", unmatched)


if __name__ == "__main__":
    main()

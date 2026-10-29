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
  F  KenPom SOS (SOS NetRtg) .......... KenPom  (free home page - no login)
  G  ESPN BPI .......................... ESPN
  H  KenPom Ovr Rating (NetRtg) ........ KenPom
  I  ESPN BPI Off ...................... ESPN
  J  KenPom Offense (ORtg) ............. KenPom
  K  Barttorvik Off (ADJOE) ............ Barttorvik
  L  ESPN BPI Def ...................... ESPN
  M  KenPom Defense (DRtg) ............. KenPom
  N  Barttorvik Def (ADJDE) ............ Barttorvik

Note: in the pre-season ESPN still serves last season's final BPI until they
publish the new season's numbers; this updater automatically picks up the new
figures as soon as ESPN posts them.  KenPom's home page likewise shows last
season's final ratings until KenPom posts the new season's.
"""

import glob
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
                    save_sheet, report, barttorvik_get, parse_tables, http_get, norm)
from openpyxl.utils import column_index_from_string

SEASON = config.CBB_SEASON or (datetime.now().year + (1 if datetime.now().month >= 5 else 0))
HERE = os.path.dirname(os.path.abspath(__file__))
WORKBOOK = os.path.join(HERE, "College Sports.xlsx")
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

# sheet name -> the name(s) KenPom uses.  KenPom spells teams almost exactly
# like Barttorvik, so this starts from that list; a tuple lists spellings to
# try in order.  ("X St." vs plain "X", e.g. McNeese, is handled automatically.)
CBB_KP_ALIAS = {
    **CBB_BT_ALIAS,
    "Cal St. Northridge": ("CSUN", "Cal St. Northridge"),
    "SIU Edwardsville": ("SIUE",),
    "UMKC": ("Kansas City",),
    "UMass": ("Massachusetts",),
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


def parse_kenpom(html):
    """{kp_team: {'sos':,'ovr':,'ortg':,'drtg':}} from the kenpom.com ratings table.

    A data row is: Rk, Team, Conf, W-L, NetRtg, ORtg, rank, DRtg, rank, AdjT,
    rank, Luck, rank, SOS NetRtg, rank, SOS ORtg, rank, SOS DRtg, rank, ...
    Every row is checked (NetRtg = ORtg - DRtg, and the same for SOS), so if
    KenPom ever moves its columns nothing wrong gets written.
    """
    out = {}
    for tbl in parse_tables(html):
        for row in tbl:
            if len(row) < 18 or not row[0].strip().isdigit():
                continue                      # header / non-data rows
            team = re.sub(r"\s*\d+$", "", row[1]).strip()   # drop a tournament seed
            net, ortg, drtg, sos, opp_o, opp_d = (_num(row[i]) for i in (4, 5, 7, 13, 15, 17))
            if not team or None in (net, ortg, drtg, sos, opp_o, opp_d):
                continue
            if abs(net - (ortg - drtg)) > 0.15 or abs(sos - (opp_o - opp_d)) > 0.15:
                continue
            out[team] = {"sos": sos, "ovr": net, "ortg": ortg, "drtg": drtg}
    return out


def _saved_kenpom_page():
    """Newest copy of the KenPom home page saved into this folder, or None."""
    named = os.path.join(HERE, config.KENPOM_SAVED_PAGE)
    pages = [named] if os.path.exists(named) else []
    pages += [p for p in glob.glob(os.path.join(HERE, "*.htm*"))
              if re.search(r"kenpom|pomeroy", os.path.basename(p), re.I)]
    return max(pages, key=os.path.getmtime) if pages else None


def pull_kenpom():
    """({kp_team: stats}, where it came from) - or ({}, why not)."""
    try:
        data = parse_kenpom(http_get("https://kenpom.com/"))
        if len(data) >= 300:
            return data, "kenpom.com"
        why = "kenpom.com did not return its ratings table"
    except Exception as e:
        why = f"could not read kenpom.com ({e})"

    saved = _saved_kenpom_page()
    if not saved:
        return {}, why
    with open(saved, encoding="utf-8", errors="replace") as f:
        data = parse_kenpom(f.read())
    if not data:
        return {}, (f"{why}; and no ratings table found in {os.path.basename(saved)} - "
                    f"save it again as 'Webpage, Complete'")
    when = datetime.fromtimestamp(os.path.getmtime(saved))
    age = (datetime.now() - when).days
    note = f" - {age} days old, save a fresh copy" if age >= 7 else ""
    print(f"  ({why}; using your saved page instead)")
    return data, f"your saved page {os.path.basename(saved)}, saved {when:%b %d}{note}"


def match_kenpom(sheet_names, kp):
    """(mapping, unmatched) like TeamMatcher.sheet_to_source(), using CBB_KP_ALIAS."""
    exact = {norm(n): n for n in kp}
    loose = {}                                # name without a trailing "St."
    for n in kp:
        loose.setdefault(re.sub(r" state$", "", norm(n)), []).append(n)

    found = {}
    for team in sheet_names:
        alias = CBB_KP_ALIAS.get(team, ())
        tries = ((alias,) if isinstance(alias, str) else tuple(alias)) + (team,)
        hit = next((exact[norm(t)] for t in tries if norm(t) in exact), None)
        if hit is None:
            for t in tries:
                cands = loose.get(re.sub(r" state$", "", norm(t)), [])
                if len(cands) == 1:
                    hit = cands[0]
                    break
        found[team] = hit

    # never let one KenPom team fill two rows on the sheet
    claims = {}
    for team, hit in found.items():
        if hit is not None:
            claims.setdefault(hit, []).append(team)
    mapping, unmatched = {}, []
    for team, hit in found.items():
        if hit is None or len(claims[hit]) > 1:
            unmatched.append(team)
        else:
            mapping[team] = kp[hit]
    return mapping, unmatched


def main():
    print(f"College Basketball - pulling {SEASON} data ...")
    wb, ws = open_sheet(WORKBOOK, SHEET)
    rows = sheet_rows(ws)
    total = 0

    # ---- ESPN ----
    print("  ESPN BPI / SOR / SOS ...")
    try:
        espn = pull_espn()
    except Exception as e:
        print(f"  ! ESPN unavailable ({e}) - columns D/E/G/I/L left unchanged.")
        espn = {}
    em = TeamMatcher(rows.keys(), alias=CBB_ESPN_ALIAS)
    for _id, rec in espn.items():
        tm = rec["team"]
        em.add(_id, [tm.get("shortDisplayName"), tm.get("nickname"), tm.get("displayName")],
               abbr=tm.get("abbreviation"), payload=rec)
    emap, emiss = em.sheet_to_source()
    if not espn:
        emiss = []                    # already reported above
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

    # ---- KenPom (free home page) ----
    kmiss = []
    if not config.AUTO_KENPOM:
        print("  KenPom: AUTO_KENPOM is off - columns F/H/J/M left as you typed them.")
    else:
        print("  KenPom SOS / Ovr / ORtg / DRtg ...")
        kp, source = pull_kenpom()
        if kp:
            kmap, kmiss = match_kenpom(rows.keys(), kp)
            for letter, field, nd in [("F", "sos", 2), ("H", "ovr", 2), ("J", "ortg", 1), ("M", "drtg", 1)]:
                vals = {t: d.get(field) for t, d in kmap.items()}
                total += write_values(ws, rows, column_index_from_string(letter), vals, nd)
            print(f"  KenPom: {len(kmap)} teams, from {source}.")
        else:
            print(f"  ! KenPom not updated ({source}) - columns F/H/J/M left unchanged.")
            print("    To fill them anyway, save the kenpom.com home page into this folder")
            print("    (see README.txt) and run the updater again.")

    save_sheet(wb, WORKBOOK)
    print(f"\nDone.  Wrote {total} stat values across {len(rows)} teams.")
    report("ESPN", emiss)
    report("Barttorvik", bmiss)
    report("KenPom", kmiss)


if __name__ == "__main__":
    main()

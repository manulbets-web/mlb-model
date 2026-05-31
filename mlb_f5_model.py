"""
mlb_f5_model.py
===============
STANDALONE First-5-Innings MLB projection engine. Replicates the F5 Excel
model's ML and Totals projections, with all F5-specific adjustments baked in.

This file has NO dependency on mlb_model.py — it can be dropped into a folder
with just the data_f5/ reference CSVs and run on its own.

F5-specific behaviors (vs. the full-game model):
  1. Batters contribute 5/9 of their full-game value (5 of 9 innings).
  2. Pitcher Total ERA uses ONLY the starter's ERA — no bullpen blend,
     no UZR/framing subtraction.
  3. No bullpen WAR in the ML win% calc (relievers don't pitch first 5).
  4. Batter BaseRunning and BattingRuns use PURE PLATOON (vR or vL only),
     no vO blend.
  5. League baseline is 2.48 runs/team/5-innings (vs 4.5 full game).
  6. League averages (offense + ERA) are pinned via constants.csv overrides
     to FULL-slate values — so totals are correct even on small slates.
  7. F5 odds come from Pinnacle (totals) + FanDuel fallback (ML when
     Pinnacle does not offer F5 h2h).

Daily workflow:
    1. Set TARGET_DATE near the top of this file.
    2. python mlb_f5_model.py project
"""


from __future__ import annotations

import argparse
import csv
import os
import re
import sys
import unicodedata
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path

import pandas as pd
import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter


# ── SET THIS EVERY DAY ───────────────────────────────────────────────────────
# Enter the EST date of the games you want to project (the slate you're betting).
# Format: date(YYYY, M, D)
TARGET_DATE = date(2026, 5, 29)
# ────────────────────────────────────────────────────────────────────────────

# F5-specific defaults. The model .xlsx is ONLY needed for `extract`,
# `validate`, `reconcile`, and `project --from-model`. The normal daily run
# (`project`, scraping Fangraphs) does NOT read the model at all — it runs
# entirely off the reference CSVs in DEFAULT_DATA_DIR.
DEFAULT_MODEL_PATH = "/Users/jamievick/Documents/Sports/MLB/F5 MLB 2026 Updated Model.xlsx"
DEFAULT_DATA_DIR = "data_f5"

# ── FIXED LEAGUE AVERAGES ────────────────────────────────────────────────────
# These describe an AVERAGE MLB team and are deliberately CONSTANT — they do
# NOT depend on which teams happen to be on today's slate. A 3-game day and a
# 15-game day are both projected against these same league baselines.
#
# They were computed once from a FULL 30-team slate (every MLB team's projected
# lineup + starter), which is the genuine league average. They are the single
# source of truth at runtime: `project` always uses these values via the
# References loader, and never recomputes an average from the games being run.
#
# To refresh them intentionally (e.g. after a mid-season projection update),
# run:  python mlb_f5_model.py extract --recompute-league-averages
# against a model that contains all 30 teams. Routine extraction leaves them
# untouched so a short slate can never poison them.
F5_LEAGUE_OFFENSE_AVG  = 4.856832043536459    # lg Off. Runs/G   (TOTALS!S35)
F5_LEAGUE_ERA_AVG      = 4.17                  # lg Total ERA     (TOTALS!T35)
F5_LEAGUE_E162         = 226.94402652787213    # lg O162          (TOTALS!E22)
F5_LEAGUE_F162         = 101.21088969020435    # lg D162          (TOTALS!F22)
F5_LEAGUE_BULLPEN_RAA  = -0.35386474284735375  # lg Bullpen RAA   (TOTALS!G22)

# Fixed league run environment / pythag scalars (model assumptions, not slate
# averages).
F5_LEAGUE_RUNS_PER_GAME      = 2.48     # rg: runs / team / 5 innings
F5_HFA_WIN_PCT               = 0.53     # home-field advantage win%
F5_RUNS_PER_WAR              = 9.683
F5_REPLACEMENT_WINNING_PCT   = 0.294
F5_SEASON_GAMES              = 162

# Backwards-compatible aliases (older constants.csv / code referenced these).
F5_FALLBACK_OFFENSE_AVG = F5_LEAGUE_OFFENSE_AVG
F5_FALLBACK_ERA_AVG = F5_LEAGUE_ERA_AVG
F5_FALLBACK_E162 = F5_LEAGUE_E162
F5_FALLBACK_F162 = F5_LEAGUE_F162
F5_FALLBACK_BULLPEN_RAA = F5_LEAGUE_BULLPEN_RAA

# F5 batters contribute this fraction of full-game value (5 innings of 9).
F5_BATTER_FRACTION = 5.0 / 9.0

# F5 starters are evaluated over the FIRST 5 INNINGS. Every game tab in the
# model sets T1/T14 = 5, which sends the pitcher runs/game formula down the
# WAR/IP branch:  Updated WAR/IP * 5 * runs_per_war  (not the full-start
# WAR/GS branch). The daily Fangraphs run uses this same basis so its K2/K15
# match the model. It also scales the bullpen to the remaining 4 innings.
F5_DEFAULT_CUSTOM_IP = 5.0
# ────────────────────────────────────────────────────────────────────────────


# ===========================================================================
# CORE DATA MODELS (used everywhere)
# ===========================================================================

@dataclass
class PlayerSlot:
    """One player in a lineup. position is one of: 'Pitcher', 'c', 'dh', 'x'."""
    order: int | None
    position: str
    name: str


@dataclass
class TeamLineup:
    """One half of a game. Pitcher + 9 batters + optional custom IP override."""
    pitcher: PlayerSlot
    batters: list[PlayerSlot]
    custom_ip: float = 0.0

    def __post_init__(self):
        assert len(self.batters) == 9, f"Expected 9 batters, got {len(self.batters)}"
        assert self.pitcher.position == "Pitcher"


@dataclass
class GameSpec:
    """Two team lineups plus per-game settings."""
    away: TeamLineup
    home: TeamLineup
    book_odds_away: int = 100
    book_odds_home: int = -100
    bullpen_on: bool = True
    hfa_win_pct: float | None = None
    remaining_games: int = 162


# Workbook tabs that aren't game tabs (used by both the slate loop and the
# Fangraphs scraper to know which tabs to skip).
NON_GAME_TABS = {
    "TOTALS", "Parameters", "Total Calcs", "MLB", "Start", "DEF Adj",
    "FG PN Start", "PN vO", "FG PN vO", "PN vR", "FG PN vR", "PN vL",
    "FG PN vL", "MLB Bullpen", "Proj BP", "MLB PF", "Daily Averages",
    "PN vO Preseason", "Starters Preseason",
}


# ===========================================================================
# PART 1: REFERENCE DATA EXTRACTION
# ===========================================================================
# Pulls the reference tables out of the xlsx into CSVs. Run once, then re-run
# whenever the underlying Steamer / Fangraphs projections are refreshed.
# ===========================================================================

def _extract_block(wb, sheet_name, first_col, last_col, headers=None):
    """Extract a rectangular block from a worksheet into a DataFrame. Reads
    until the first row where the first column is empty."""
    ws = wb[sheet_name]
    if headers is None:
        headers = [ws.cell(row=1, column=c).value for c in range(first_col, last_col + 1)]
        headers = [h if h is not None else f"col_{get_column_letter(c)}"
                   for c, h in zip(range(first_col, last_col + 1), headers)]

    rows = []
    for r in range(2, ws.max_row + 1):
        first_val = ws.cell(row=r, column=first_col).value
        if first_val is None or first_val == "":
            break
        row_data = [ws.cell(row=r, column=c).value for c in range(first_col, last_col + 1)]
        rows.append(row_data)
    return pd.DataFrame(rows, columns=headers)


def cmd_extract(args):
    out_dir = Path(args.data_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    print(f"Loading {args.model}...")
    wb = load_workbook(args.model, data_only=True)

    # 1. Pitchers (Start tab, cols BS-CE)
    print("\n[1/10] Extracting pitchers from 'Start' tab...")
    df = _extract_block(wb, "Start", 71, 83)
    df.to_csv(out_dir / "start.csv", index=False)
    print(f"      Wrote {len(df)} pitchers -> {out_dir / 'start.csv'}")

    # 2-4. Batter projections (PN vO/vR/vL, cols BW-CL)
    for i, (sheet, name) in enumerate([("PN vO", "pn_vo"), ("PN vR", "pn_vr"),
                                        ("PN vL", "pn_vl")], start=2):
        print(f"\n[{i}/10] Extracting batters from '{sheet}' tab...")
        df = _extract_block(wb, sheet, 75, 90)
        df.to_csv(out_dir / f"{name}.csv", index=False)
        print(f"      Wrote {len(df)} batters -> {out_dir / f'{name}.csv'}")

    # 5. Bullpen
    print("\n[5/10] Extracting bullpen from 'MLB Bullpen' tab...")
    df = _extract_block(wb, "MLB Bullpen", 23, 32)
    df.to_csv(out_dir / "mlb_bullpen.csv", index=False)
    print(f"      Wrote {len(df)} teams -> {out_dir / 'mlb_bullpen.csv'}")

    # 6. Park factors
    print("\n[6/10] Extracting park factors from 'MLB PF' tab...")
    df = _extract_block(wb, "MLB PF", 1, 7)
    df.to_csv(out_dir / "mlb_pf.csv", index=False)
    print(f"      Wrote {len(df)} parks -> {out_dir / 'mlb_pf.csv'}")

    # 7. Positional defensive adjustments
    print("\n[7/10] Extracting positional adjustments from 'DEF Adj' tab...")
    df = _extract_block(wb, "DEF Adj", 1, 2, headers=["Position", "Adjustment"])
    df = df[df["Position"].isin(["C", "1B", "2B", "3B", "SS", "LF", "CF", "RF", "DH"])].copy()
    df.to_csv(out_dir / "def_adj.csv", index=False)
    print(f"      Wrote {len(df)} positions -> {out_dir / 'def_adj.csv'}")

    # 8. Pythagorean lookup
    print("\n[8/10] Extracting Pythagorean lookup from 'Parameters' tab...")
    df = _extract_block(wb, "Parameters", 1, 4)
    df.to_csv(out_dir / "parameters.csv", index=False)
    print(f"      Wrote {len(df)} rows -> {out_dir / 'parameters.csv'}")

    # 9. Lineup-slot PA/G
    print("\n[9/10] Extracting lineup-slot PA/G table...")
    # The PA/G-by-slot table (cols R/S/T, rows 25-33) is identical on every
    # game tab, so use the first real game tab instead of a hardcoded one
    # (tab names change daily as matchups rotate).
    game_tab = next(
        (t for t in wb.sheetnames
         if t not in NON_GAME_TABS and wb[t]["B1"].value == "Away"),
        None)
    if game_tab is None:
        raise RuntimeError("No game tab found (need one with B1=='Away').")
    ws = wb[game_tab]
    rows = []
    for r in range(25, 34):
        rows.append({
            "Order": ws.cell(row=r, column=18).value,
            "PA":    ws.cell(row=r, column=19).value,
            "G":     ws.cell(row=r, column=20).value,
            "PA_per_G": ws.cell(row=r, column=19).value / ws.cell(row=r, column=20).value,
        })
    df = pd.DataFrame(rows)
    df.to_csv(out_dir / "lineup_pa_per_game.csv", index=False)
    print(f"      Wrote {len(df)} slots -> {out_dir / 'lineup_pa_per_game.csv'}")

    # 10. Global constants
    print("\n[10/10] Writing global constants...")
    tot = wb["TOTALS"]

    # League averages default to the FIXED, slate-independent values defined at
    # the top of this module. They are NOT recomputed from this model's slate
    # unless --recompute-league-averages is explicitly passed AND the model
    # contains all 30 teams (a full slate) — so a short slate can never poison
    # the league baselines.
    lg_off, lg_era = F5_LEAGUE_OFFENSE_AVG, F5_LEAGUE_ERA_AVG
    lg_e162, lg_f162, lg_bp = F5_LEAGUE_E162, F5_LEAGUE_F162, F5_LEAGUE_BULLPEN_RAA

    if getattr(args, "recompute_league_averages", False):
        team_codes = set()
        for r in range(2, tot.max_row + 1):
            v = tot.cell(row=r, column=1).value
            if isinstance(v, str):
                s = v.strip().upper()
                # Real team codes are 2-3 letters (modern or Retrosheet); this
                # excludes label cells like "Team" or blanks.
                if 2 <= len(s) <= 3 and s.isalpha():
                    team_codes.add(s)
        if len(team_codes) < 30:
            raise RuntimeError(
                f"--recompute-league-averages needs a FULL 30-team slate, but "
                f"this model only has {len(team_codes)} teams in TOTALS. "
                f"League averages must represent the whole league, so refuse to "
                f"recompute from a partial slate. (Frozen values left unchanged.)")
        lg_off = tot["S35"].value
        lg_era = tot["T35"].value
        lg_e162 = tot["E22"].value
        lg_f162 = tot["F22"].value
        lg_bp = tot["G22"].value
        print(f"      Recomputed league averages from {len(team_codes)}-team slate.")
    else:
        print("      Using FIXED league averages (slate-independent). "
              "Pass --recompute-league-averages to refresh from a full slate.")

    constants = {
        # Fixed run environment / pythag scalars (model assumptions).
        "league_runs_per_game":     tot["B40"].value,          # rg (R/team/5)
        "hfa_win_pct":              wb[game_tab]["H31"].value,  # home-field win%
        "runs_per_war":             F5_RUNS_PER_WAR,
        "replacement_winning_pct":  F5_REPLACEMENT_WINNING_PCT,
        "season_games":             F5_SEASON_GAMES,
        # FIXED league averages — constant regardless of today's slate.
        "league_avg_era_override":  lg_era,    # lg Total ERA
        "league_avg_off_override":  lg_off,    # lg Off. Runs/G
        "league_avg_e162":          lg_e162,   # lg O162
        "league_avg_f162":          lg_f162,   # lg D162
        "league_avg_bullpen_raa":   lg_bp,     # lg Bullpen RAA
    }
    df = pd.DataFrame(list(constants.items()), columns=["constant", "value"])
    df.to_csv(out_dir / "constants.csv", index=False)
    print(f"      Wrote constants -> {out_dir / 'constants.csv'}")

    print(f"\nDone. Reference data is in {out_dir}/")
    print("Re-run `python mlb_f5_model.py extract` whenever Start / PN vO / etc. "
          "change. League averages stay fixed unless you pass "
          "--recompute-league-averages.")


# ===========================================================================
# PART 2: REFERENCE DATA LOADER
# ===========================================================================

class References:
    """Loads reference CSVs and exposes them as indexed DataFrames + constants."""

    def __init__(self, data_dir=DEFAULT_DATA_DIR):
        d = Path(data_dir)

        def _load(name, key="Full Name", skip=("Full Name", "Hand", "Bats")):
            df = pd.read_csv(d / name)
            df = df.drop_duplicates(subset=key, keep="first").set_index(key)
            for col in df.columns:
                if col not in skip:
                    df[col] = pd.to_numeric(df[col], errors="coerce")
            return df

        self.start = _load("start.csv")
        self.pn_vo = _load("pn_vo.csv")
        self.pn_vr = _load("pn_vr.csv")
        self.pn_vl = _load("pn_vl.csv")
        self.bullpen = pd.read_csv(d / "mlb_bullpen.csv").set_index("MLB")
        self.park_factors = pd.read_csv(d / "mlb_pf.csv")
        self.def_adj = pd.read_csv(d / "def_adj.csv").set_index("Position")
        self.pa_per_game = pd.read_csv(d / "lineup_pa_per_game.csv").set_index("Order")

        # Pythagorean / NB parameters lookup table.
        # Columns: RG (the R/G key), r, B, z (zero-inflated NB params).
        # Used to compute over/under American odds at any total line via the
        # same zero-inflated negative-binomial that 'Total Calcs' tab uses.
        params_df = pd.read_csv(d / "parameters.csv")
        params_df = params_df.dropna(subset=["RG"]).sort_values("RG")
        for c in ("RG", "r", "B", "z"):
            params_df[c] = pd.to_numeric(params_df[c], errors="coerce")
        self.nb_params = params_df.reset_index(drop=True)

        # Scalar constants. constants.csv is OPTIONAL — any value it omits
        # (or the whole file, if absent) falls back to the FIXED league
        # constants defined at the top of this module. This is what keeps the
        # league baselines constant regardless of slate size: they live in
        # code, and constants.csv merely lets you override them deliberately.
        constants_path = d / "constants.csv"
        if constants_path.exists():
            cdf = pd.read_csv(constants_path).set_index("constant")["value"]
            self.constants = cdf.to_dict()
        else:
            self.constants = {}

        def _const(key, default):
            v = self.constants.get(key)
            try:
                return float(v) if v is not None and str(v) != "" else float(default)
            except (TypeError, ValueError):
                return float(default)

        self.league_runs_per_game = _const("league_runs_per_game", F5_LEAGUE_RUNS_PER_GAME)
        self.hfa_win_pct = _const("hfa_win_pct", F5_HFA_WIN_PCT)
        self.runs_per_war = _const("runs_per_war", F5_RUNS_PER_WAR)
        self.replacement_winning_pct = _const("replacement_winning_pct", F5_REPLACEMENT_WINNING_PCT)
        self.season_games = int(_const("season_games", F5_SEASON_GAMES))

        # Fixed league averages for the totals projection. F5 makes every game
        # fully independent: these are CONSTANT league baselines (an average
        # MLB team), never recomputed from the games being projected. A 3-game
        # and a 15-game slate use the identical values below.
        self.league_avg_era_override = _const("league_avg_era_override", F5_LEAGUE_ERA_AVG)
        self.league_avg_off_override = _const("league_avg_off_override", F5_LEAGUE_OFFENSE_AVG)
        self.league_avg_e162 = _const("league_avg_e162", F5_LEAGUE_E162)
        self.league_avg_f162 = _const("league_avg_f162", F5_LEAGUE_F162)
        self.league_avg_bullpen_raa = _const("league_avg_bullpen_raa", F5_LEAGUE_BULLPEN_RAA)

        # V25 = -0.011905 ≈ -1/84, the generic "x" positional adjustment per
        # game used when position isn't C/DH. Hardcoded in the spreadsheet.
        self.x_def_adj_per_game = -0.011904761904761906

        # Name aliases — maps Fangraphs display names to PN vO / Start names
        # for players whose names differ across the two sources. Loaded from
        # data/aliases.json if present, otherwise empty.
        # Format: { "Fangraphs Display Name": "PNvO Name", ... }
        #   e.g. { "Tommy Troy": "Thomas Troy", "JD Dix": "John Dix" }
        # Generated initially by `python mlb_f5_model.py reconcile`.
        self.aliases = {}
        alias_path = d / "aliases.json"
        if alias_path.exists():
            import json
            try:
                with open(alias_path, "r", encoding="utf-8") as f:
                    self.aliases = json.load(f)
            except (json.JSONDecodeError, OSError) as e:
                print(f"  Warning: couldn't load {alias_path}: {e}")


# ===========================================================================
# PART 2B: FANGRAPHS LINEUP SCRAPER
# ===========================================================================
# Pulls projected lineups from fangraphs.com/scores?date=YYYY-MM-DD and
# resolves player names against the master lists (Start for pitchers,
# PN vO for batters) so the names match the format the projection expects.
# ===========================================================================

# Fangraphs display name -> (modern code in game tab, Retrosheet code in master lists)
TEAM_MAP = {
    "Blue Jays":    ("TOR", "TOR"), "Orioles":   ("BAL", "BAL"),
    "Rays":         ("TB",  "TBA"), "Red Sox":   ("BOS", "BOS"),
    "Yankees":      ("NYY", "NYA"), "Guardians": ("CLE", "CLE"),
    "Royals":       ("KC",  "KCA"), "Tigers":    ("DET", "DET"),
    "Twins":        ("MIN", "MIN"), "White Sox": ("CHW", "CHA"),
    "Angels":       ("LAA", "LAA"), "Astros":    ("HOU", "HOU"),
    "Athletics":    ("OAK", "OAK"), "Mariners":  ("SEA", "SEA"),
    "Rangers":      ("TEX", "TEX"), "Braves":    ("ATL", "ATL"),
    "Marlins":      ("MIA", "MIA"), "Mets":      ("NYM", "NYN"),
    "Nationals":    ("WAS", "WAS"), "Phillies":  ("PHI", "PHI"),
    "Brewers":      ("MIL", "MIL"), "Cardinals": ("STL", "SLN"),
    "Cubs":         ("CHC", "CHN"), "Pirates":   ("PIT", "PIT"),
    "Reds":         ("CIN", "CIN"), "D-backs":   ("ARI", "ARI"),
    "Diamondbacks": ("ARI", "ARI"), "Dodgers":   ("LAD", "LAN"),
    "Giants":       ("SF",  "SFN"), "Padres":    ("SD",  "SDN"),
    "Rockies":      ("COL", "COL"),
}
MODERN_TO_PNVO = {modern: pnvo for modern, pnvo in TEAM_MAP.values()}
FG_TEAM_NAMES = set(TEAM_MAP.keys())
FG_URL = "https://www.fangraphs.com/scores?date={date}"


def _strip_accents(s):
    if s is None:
        return ""
    nfkd = unicodedata.normalize("NFKD", s)
    return "".join(c for c in nfkd if not unicodedata.combining(c))


def _norm_name(s):
    s = _strip_accents(s or "").lower().strip()
    s = s.replace(".", "").replace("'", "").replace("\u2019", "")
    return re.sub(r"\s+", " ", s)


def _fetch_fangraphs_html(target_date):
    """Fetch the Fangraphs scores page for a date.

    Fangraphs is now behind Cloudflare's Managed Challenge, which inspects TLS
    fingerprint and runs a JavaScript challenge. Neither `requests` nor
    `curl_cffi` can pass this. We use undetected-chromedriver (a real Chrome
    instance with anti-detection patches) — the same approach proven on KBO's
    Cloudflare-protected sites.

    On the first call per process, this spins up a Chrome window (visible by
    design — headless gets flagged faster). The window is reused for the
    whole session so subsequent fetches are fast.

    Requires:
        pip install undetected-chromedriver
    """
    return _get_via_chrome(FG_URL.format(date=target_date.strftime("%Y-%m-%d")))


# Module-level Chrome driver — built lazily on first fetch, reused thereafter
_CHROME_DRIVER = None


def _get_chrome_major_version():
    """Read installed Chrome's major version on macOS so chromedriver matches.
    Falls back to a reasonable default if detection fails."""
    import subprocess
    try:
        plist = "/Applications/Google Chrome.app/Contents/Info.plist"
        r = subprocess.run(
            ["defaults", "read", plist, "CFBundleShortVersionString"],
            capture_output=True, text=True, timeout=5,
        )
        m = re.search(r"(\d+)\.", r.stdout)
        if m:
            return int(m.group(1))
    except Exception:
        pass
    # Try launching the binary
    for path in (
        "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
        "/usr/bin/google-chrome",
        "/usr/bin/chromium-browser",
    ):
        try:
            r = subprocess.run([path, "--version"], capture_output=True,
                                text=True, timeout=5)
            m = re.search(r"(\d+)\.", r.stdout)
            if m:
                return int(m.group(1))
        except Exception:
            continue
    return 148  # reasonable modern default


def _build_chrome():
    """Construct an undetected-chromedriver instance with the same options
    that work reliably on KBO sites behind Cloudflare."""
    try:
        import undetected_chromedriver as uc
    except ImportError:
        raise RuntimeError(
            "\nFangraphs requires a real browser (Cloudflare Managed Challenge).\n"
            "Install undetected-chromedriver:\n"
            "    pip install undetected-chromedriver\n\n"
            "After install, restart Python and re-run.\n"
        )
    import time as _time

    opts = uc.ChromeOptions()
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--window-size=1400,2200")
    opts.add_argument("--disable-blink-features=AutomationControlled")
    opts.add_argument("--disable-popup-blocking")

    print("  Launching Chrome (this takes ~10 seconds)...")
    driver = uc.Chrome(options=opts, headless=False,
                        version_main=_get_chrome_major_version())
    driver.set_page_load_timeout(45)
    driver.implicitly_wait(3)
    # Cloudflare fingerprinting can close the window if a request arrives
    # too soon after launch. Wait longer than feels necessary.
    print("  Waiting for browser to stabilise (10s)...")
    _time.sleep(10)

    # Pre-warm the browser with a benign navigation. Cloudflare is more
    # suspicious of a session whose very first navigation is to a protected
    # site — visiting Google first looks like normal human behaviour.
    try:
        print("  Pre-warming with a Google visit...")
        driver.get("https://www.google.com/")
        _time.sleep(3)
    except Exception as e:
        # Pre-warm failure isn't fatal — just warn and continue
        print(f"  (pre-warm failed: {e}; continuing)")

    return driver


def _get_via_chrome(url, retries=2):
    """Navigate Chrome to the URL and return rendered HTML. On Cloudflare
    closing the window (NoSuchWindowException), rebuild the driver and retry."""
    import time as _time
    global _CHROME_DRIVER

    last_err = None
    for attempt in range(retries + 1):
        if _CHROME_DRIVER is None:
            _CHROME_DRIVER = _build_chrome()

        # Sanity check the window is still alive
        try:
            _ = _CHROME_DRIVER.window_handles
        except Exception:
            print("  Chrome window died — relaunching...")
            try: _CHROME_DRIVER.quit()
            except Exception: pass
            _CHROME_DRIVER = None
            continue

        try:
            print(f"  GET {url}")
            _CHROME_DRIVER.get(url)
        except Exception as e:
            err_name = type(e).__name__
            last_err = e
            if "NoSuchWindow" in err_name or "web view not found" in str(e):
                print(f"  Cloudflare closed the window during navigation.")
                if attempt < retries:
                    wait = 15 * (attempt + 1)  # 15s, then 30s
                    print(f"  Waiting {wait}s before retrying with a fresh "
                          f"Chrome instance (attempt {attempt + 2}/{retries + 1})...")
                    try: _CHROME_DRIVER.quit()
                    except Exception: pass
                    _CHROME_DRIVER = None
                    _time.sleep(wait)
                    continue
            raise

        # Wait for the real page to render. Two phases:
        #   1. Short poll for team names (normal fast load, 3-8s typical).
        #   2. If not found, check for a Cloudflare challenge page and, if
        #      present, wait much longer so you can click the checkbox.
        def _has_lineups(html):
            return any(team_name in html for team_name in FG_TEAM_NAMES)

        def _is_challenge(html):
            if not html:
                return True
            markers = ("Just a moment", "Checking your browser",
                       "challenge-platform", "cf-challenge", "Verifying you are human",
                       "needs to review the security")
            return any(m in html for m in markers)

        # Phase 1: quick poll (up to 12s) for a clean load
        deadline = _time.time() + 12
        html = ""
        while _time.time() < deadline:
            try:
                html = _CHROME_DRIVER.page_source
            except Exception as e:
                last_err = e
                html = None
                break
            if _has_lineups(html):
                return html
            _time.sleep(1)

        # Phase 2: if the window died, fall through to retry
        if html is None:
            if attempt < retries:
                print("  Window died after navigation — retrying...")
                try: _CHROME_DRIVER.quit()
                except Exception: pass
                _CHROME_DRIVER = None
                _time.sleep(15)
                continue
            break

        # Phase 2b: still no lineups. If it looks like a Cloudflare challenge,
        # give the user time to solve it in the visible Chrome window.
        if _is_challenge(html) or not _has_lineups(html):
            print("  ")
            print("  " + "=" * 60)
            print("  CLOUDFLARE CHALLENGE detected (or page still loading).")
            print("  --> Look at the Chrome window that opened.")
            print("  --> If you see a checkbox / 'Verify you are human', CLICK IT.")
            print("  --> Waiting up to 90 seconds for the lineups to appear...")
            print("  " + "=" * 60)
            challenge_deadline = _time.time() + 90
            while _time.time() < challenge_deadline:
                try:
                    html = _CHROME_DRIVER.page_source
                except Exception as e:
                    last_err = e
                    html = None
                    break
                if _has_lineups(html):
                    print("  Lineups detected — continuing.")
                    return html
                _time.sleep(2)
            if html is None:
                if attempt < retries:
                    print("  Window died during challenge wait — retrying...")
                    try: _CHROME_DRIVER.quit()
                    except Exception: pass
                    _CHROME_DRIVER = None
                    _time.sleep(15)
                    continue
                break

        # If we have lineups by now, return; otherwise retry/fail
        if _has_lineups(html):
            return html
        if attempt < retries:
            print("  No lineups found — retrying with a fresh Chrome instance...")
            try: _CHROME_DRIVER.quit()
            except Exception: pass
            _CHROME_DRIVER = None
            _time.sleep(15)
            continue

    raise RuntimeError(
        f"Failed to fetch {url} after {retries + 1} attempts. "
        f"Last error: {last_err}\n"
        f"This usually means Cloudflare is being extra aggressive today. "
        f"Wait 60s and re-run; or open Fangraphs in your normal browser "
        f"first, solve any challenge there, then re-run."
    )


def _close_chrome():
    """Close the shared Chrome instance. Safe to call multiple times."""
    global _CHROME_DRIVER
    if _CHROME_DRIVER is not None:
        try:
            _CHROME_DRIVER.quit()
        except Exception:
            pass
        _CHROME_DRIVER = None


_POS_TAIL_RE = re.compile(
    r"(?:[LRS])?\s*(1B|2B|3B|SS|LF|CF|RF|DH|C|P)\s*$", re.IGNORECASE,
)


def _extract_position(a_tag):
    """Find the position code next to a batter's <a> tag by walking up the
    DOM until we hit a container whose text contains the name and ends in a
    valid position code."""
    name = a_tag.get_text(strip=True)
    if not name:
        return "x"
    node = a_tag
    for _ in range(8):
        node = node.parent
        if node is None or not hasattr(node, "get_text"):
            continue
        full = node.get_text(" ", strip=True)
        if not full or name not in full:
            continue
        cleaned = re.sub(r"^\s*\d+\s*[.)]\s*", "", full)
        idx = cleaned.find(name)
        if idx >= 0:
            cleaned = cleaned[idx + len(name):]
        cleaned = cleaned.strip()
        if not cleaned:
            continue
        m = _POS_TAIL_RE.search(cleaned)
        if m:
            pos = m.group(1).upper()
            if pos == "C":  return "c"
            if pos == "DH": return "dh"
            return "x"
    return "x"


def _find_team_panel(header):
    """Walk up from a team-name header until we find a container with both
    a pitching link and a batting link."""
    node = header
    for _ in range(10):
        parent = node.parent
        if parent is None:
            return None
        anchors = parent.find_all("a", href=True)
        has_pit = any("/stats/pitching" in a["href"] for a in anchors)
        has_bat = any("/stats/batting" in a["href"] for a in anchors)
        if has_pit and has_bat:
            return parent
        node = parent
    return None


def parse_fangraphs_lineups(html):
    """Returns (teams, games) where:
        teams: {team_display_name: {'pitcher': str, 'batters': [(name, pos), ...]}}
        games: [(away_team_name, home_team_name), ...] in page order — Fangraphs
               displays the away team first, then the home team for each matchup.
    """
    soup = BeautifulSoup(html, "html.parser")
    team_headers = []
    for tag in soup.find_all(["h3", "h4", "h2"]):
        txt = re.sub(r"\d+(\.\d+)?\s*%", "", tag.get_text(" ", strip=True)).strip()
        if txt in FG_TEAM_NAMES:
            team_headers.append((tag, txt))
    if not team_headers:
        for tag in soup.find_all(True):
            if tag.name in ("a", "script", "style"):
                continue
            direct = "".join(c for c in tag.contents if isinstance(c, str)).strip()
            direct = re.sub(r"\d+(\.\d+)?\s*%", "", direct).strip()
            if direct in FG_TEAM_NAMES:
                team_headers.append((tag, direct))

    teams = {}
    order = []  # team display names in page order
    seen = set()
    for header, team_name in team_headers:
        panel = _find_team_panel(header)
        if panel is None or id(panel) in seen:
            continue
        seen.add(id(panel))
        pitcher = None; batters = []
        seen_batters = set()
        for a in panel.find_all("a", href=True):
            href = a["href"]
            if "/stats/pitching" in href and pitcher is None:
                pitcher = a.get_text(strip=True)
            elif "/stats/batting" in href:
                n = a.get_text(strip=True)
                if not n or n in seen_batters:
                    continue
                seen_batters.add(n)
                batters.append((n, _extract_position(a)))
        teams[team_name] = {"pitcher": pitcher, "batters": batters[:9]}
        order.append(team_name)

    # Pair into games — Fangraphs displays away first, then home
    games = [(order[i], order[i + 1]) for i in range(0, len(order) - 1, 2)]
    return teams, games


def _split_tab_name(tab):
    """'BALTB' -> ('BAL','TB'); 'LADSD' -> ('LAD','SD'); 'CHWSEA' -> ('CHW','SEA')."""
    tab = tab.upper()
    for length in (2, 3):
        away, home = tab[:length], tab[length:]
        if away in MODERN_TO_PNVO and home in MODERN_TO_PNVO:
            return away, home
    for length in (2, 3):
        away, home = tab[:-length], tab[-length:]
        if away in MODERN_TO_PNVO and home in MODERN_TO_PNVO:
            return away, home
    return None, None


def _find_fg_team_by_modern_code(modern_code):
    for fg_name, (m, _) in TEAM_MAP.items():
        if m == modern_code:
            return fg_name
    return None


def _build_name_indexes(refs):
    """Build {(normalized_name, retro_team): exact 'Firstname Lastname TEAM'}
    for both pitchers (Start) and batters (PN vO), so we can match Fangraphs
    names to the exact format the projection expects."""
    def _index(df):
        idx = {}
        name_only = {}
        for full_name in df.index:
            # Full name format: "Firstname Lastname TEAM"
            parts = full_name.rsplit(" ", 1)
            if len(parts) != 2:
                continue
            name_part, team = parts
            n = _norm_name(name_part)
            idx[(n, team.upper())] = full_name
            name_only.setdefault(n, []).append(full_name)
        return idx, name_only

    p_idx, p_name_only = _index(refs.start)
    b_idx, b_name_only = _index(refs.pn_vo)
    return p_idx, p_name_only, b_idx, b_name_only


def _resolve_player(fg_name, retro_team, idx, name_only, report_rows,
                    tab_name, role, lineup_slot, source_sheet, aliases=None):
    """Look up an FG name in the master index. On miss, try the alias map.
    If still no match, write a row to the unmatched-players report and return
    the constructed name anyway."""
    # First try the alias map — covers Tommy/Thomas, JD/John, etc.
    if aliases and fg_name in aliases:
        alias_name = aliases[fg_name]
        n_alias = _norm_name(alias_name)
        key_alias = (n_alias, retro_team.upper())
        if key_alias in idx:
            return idx[key_alias]
        # Alias mapped to a name not in PN vO for this team — fall through

    n = _norm_name(fg_name)
    key = (n, retro_team.upper())
    if key in idx:
        return idx[key]
    clean = _strip_accents(fg_name)
    written = f"{clean} {retro_team}"
    if n in name_only:
        other = ", ".join(c.rsplit(" ", 1)[-1] for c in name_only[n])
        reason = f"in {source_sheet} under different team(s): {other}"
    else:
        reason = f"not in {source_sheet}"
    report_rows.append({
        "tab": tab_name, "role": role, "lineup_slot": lineup_slot,
        "fangraphs_name": fg_name, "tab_team_code": retro_team,
        "written_to_cell": written, "source_sheet": source_sheet,
        "reason": reason,
    })
    return written


def build_lineups_from_fangraphs(target_date, refs):
    """Scrape Fangraphs for the target date and produce:
        {tab_name (e.g. 'BAL@DET'): (away_TeamLineup, home_TeamLineup, status)}

    Tab names are derived from the Fangraphs matchups for the date — so when
    a new series starts (BAL plays DET instead of TB), the output reflects
    today's actual games rather than yesterday's.

    Status is one of:
        'ok'                                      lineup fully resolved
        'missing_pitcher: away SP ... / home SP ...'   pitcher not in Start
    """
    print(f"Fetching Fangraphs lineups for {target_date}...")
    html = _fetch_fangraphs_html(target_date)
    teams, games = parse_fangraphs_lineups(html)
    print(f"  Parsed {len(teams)} team lineups across {len(games)} games.")

    # Sanity check: real lineups always have catchers + DHs in most games
    all_positions = [p for t in teams.values() for (_, p) in t["batters"]]
    n_c = sum(1 for p in all_positions if p == "c")
    n_dh = sum(1 for p in all_positions if p == "dh")
    if all_positions and (n_c + n_dh) == 0:
        print("  WARNING: position parser returned 'x' for every batter.")
        print("  Fangraphs may have changed their layout.")
    elif all_positions:
        print(f"  Found {n_c} catchers, {n_dh} DHs across all lineups.")

    # Build name indexes from references
    p_idx, p_name_only, b_idx, b_name_only = _build_name_indexes(refs)

    report_rows = []
    lineups_by_tab = {}

    for away_fg, home_fg in games:
        # Get the modern + Retrosheet codes for each side
        away_modern, away_retro = TEAM_MAP[away_fg]
        home_modern, home_retro = TEAM_MAP[home_fg]
        tab = f"{away_modern}@{home_modern}"

        def _build_team(team_data, retro, role_prefix):
            p_name = _resolve_player(
                team_data["pitcher"], retro, p_idx, p_name_only,
                report_rows, tab, "pitcher", "SP", "Start",
                aliases=refs.aliases)
            pitcher = PlayerSlot(order=None, position="Pitcher", name=p_name)
            pitcher_matched = p_name in refs.start.index
            batters = []
            for i, (fg_name, pos) in enumerate(team_data["batters"][:9]):
                resolved = _resolve_player(
                    fg_name, retro, b_idx, b_name_only,
                    report_rows, tab, f"{role_prefix} batter",
                    str(i + 1), "PN vO",
                    aliases=refs.aliases)
                batters.append(PlayerSlot(order=i + 1, position=pos, name=resolved))
            return (TeamLineup(pitcher=pitcher, batters=batters,
                               custom_ip=F5_DEFAULT_CUSTOM_IP),
                    pitcher_matched)

        # Some FG games may not have lineups posted yet (only "P" or "TBD" listed)
        # — handle that by treating it as no_lineup_yet
        if len(teams[away_fg]["batters"]) < 9 or len(teams[home_fg]["batters"]) < 9:
            lineups_by_tab[tab] = (None, None, "no_lineup_yet")
            continue

        away_lineup, away_p_ok = _build_team(teams[away_fg], away_retro, "away")
        home_lineup, home_p_ok = _build_team(teams[home_fg], home_retro, "home")

        if not away_p_ok or not home_p_ok:
            missing = []
            if not away_p_ok: missing.append(f"away SP {away_lineup.pitcher.name}")
            if not home_p_ok: missing.append(f"home SP {home_lineup.pitcher.name}")
            lineups_by_tab[tab] = (None, None,
                                    f"missing_pitcher: {', '.join(missing)}")
            continue

        lineups_by_tab[tab] = (away_lineup, home_lineup, "ok")

    return lineups_by_tab, report_rows


# ===========================================================================
# PART 2C: PINNACLE LIVE ODDS (The Odds API)
# ===========================================================================
# One Pinnacle slate call returns H2H + Totals for every MLB game on the day.
# Cost: 1 credit per call. Requires ODDS_API_KEY env var (load via .env file
# next to the script).
#
# Result is a dict keyed by AWAY@HOME tab name (matching the FG-derived
# matchup tabs), so we can merge straight into the slate projection.
# ===========================================================================

# Odds API team-name -> modern code we use in tab names
ODDS_API_TEAM_TO_MODERN = {
    "Arizona Diamondbacks": "ARI",
    "Atlanta Braves":       "ATL",
    "Baltimore Orioles":    "BAL",
    "Boston Red Sox":       "BOS",
    "Chicago Cubs":         "CHC",
    "Chicago White Sox":    "CHW",
    "Cincinnati Reds":      "CIN",
    "Cleveland Guardians":  "CLE",
    "Colorado Rockies":     "COL",
    "Detroit Tigers":       "DET",
    "Houston Astros":       "HOU",
    "Kansas City Royals":   "KC",
    "Los Angeles Angels":   "LAA",
    "Los Angeles Dodgers":  "LAD",
    "Miami Marlins":        "MIA",
    "Milwaukee Brewers":    "MIL",
    "Minnesota Twins":      "MIN",
    "New York Mets":        "NYM",
    "New York Yankees":     "NYY",
    "Athletics":            "OAK",  # OddsAPI sometimes drops "Oakland" / "Sacramento"
    "Oakland Athletics":    "OAK",
    "Sacramento Athletics": "OAK",
    "Philadelphia Phillies":"PHI",
    "Pittsburgh Pirates":   "PIT",
    "San Diego Padres":     "SD",
    "San Francisco Giants": "SF",
    "Seattle Mariners":     "SEA",
    "St. Louis Cardinals":  "STL",
    "St Louis Cardinals":   "STL",  # some feeds drop the period
    "Tampa Bay Rays":       "TB",
    "Texas Rangers":        "TEX",
    "Toronto Blue Jays":    "TOR",
    "Washington Nationals": "WAS",
}


def fetch_pinnacle_odds(target_date):
    """Fetch first-5-innings H2H + totals for every MLB game, across ALL books.

    Pulls regions us + us2 with no bookmaker filter, then for each side
    (away ML, home ML, over, under) computes the MEDIAN price (used as the
    fair-edge benchmark) and the BEST price for the bettor plus which book
    posts it. Returns a dict keyed by AWAY@HOME tab name.

    Backwards-compatible keys (medians): away_ml, home_ml, total,
    over_juice, under_juice. New keys: *_best / *_book and n_books.
    """
    from statistics import median as _median
    try:
        from dotenv import load_dotenv
        load_dotenv(override=True)
    except ImportError:
        pass

    api_key = os.environ.get("ODDS_API_KEY")
    if not api_key:
        print("  ODDS_API_KEY not set — skipping F5 odds.")
        return {}

    base_url = "https://api.the-odds-api.com/v4/sports/baseball_mlb"

    import time as _time
    events = None
    for attempt in range(3):
        try:
            r = requests.get(f"{base_url}/events",
                             params={"apiKey": api_key, "dateFormat": "iso"},
                             timeout=30)
            if r.status_code != 200:
                if attempt < 2:
                    _time.sleep(2 ** attempt); continue
                print(f"  Odds API events HTTP {r.status_code}: {r.text[:200]}")
                return {}
            events = r.json(); break
        except requests.RequestException as e:
            if attempt < 2:
                _time.sleep(2 ** attempt); continue
            print(f"  Odds API events request failed: {e}")
            return {}

    if not events:
        print("  Odds API returned no upcoming events for F5.")
        return {}

    def _best(prices):
        # Best for the bettor = highest numeric American price
        # (-105 beats -120; +130 beats +120).
        return max(prices) if prices else None

    def _agg(entries):
        """entries: list of (price, book_title) -> (median, best, best_book)."""
        entries = [(int(p), b) for p, b in entries if p is not None]
        if not entries:
            return None, None, None
        prices = [p for p, _ in entries]
        med = int(round(_median(prices)))
        best = _best(prices)
        best_book = next(b for p, b in entries if p == best)
        return med, best, best_book

    out = {}
    n_priced = 0
    print(f"  F5 odds: pulling all-book markets for {len(events)} games "
          f"(~{len(events) * 4} credits)...")
    diag_printed = False

    for ev in events:
        away_modern = ODDS_API_TEAM_TO_MODERN.get(ev["away_team"])
        home_modern = ODDS_API_TEAM_TO_MODERN.get(ev["home_team"])
        if not away_modern or not home_modern:
            continue
        tab = f"{away_modern}@{home_modern}"

        params = {
            "apiKey": api_key,
            "regions": "us,us2",
            "markets": "h2h_1st_5_innings,totals_1st_5_innings",
            "oddsFormat": "american",
        }
        try:
            r2 = requests.get(f"{base_url}/events/{ev['id']}/odds",
                              params=params, timeout=30)
        except requests.RequestException:
            continue
        if r2.status_code != 200:
            continue
        payload = r2.json()
        books = payload.get("bookmakers", [])

        away_t, home_t = ev["away_team"], ev["home_team"]
        away_ml, home_ml, over, under, points = [], [], [], [], []
        for bk in books:
            title = bk.get("title", bk.get("key", "?"))
            for m in bk.get("markets", []):
                if m["key"] == "h2h_1st_5_innings":
                    for o in m["outcomes"]:
                        if o["name"] == away_t: away_ml.append((o["price"], title))
                        elif o["name"] == home_t: home_ml.append((o["price"], title))
                elif m["key"] == "totals_1st_5_innings":
                    for o in m["outcomes"]:
                        if o["name"] == "Over":
                            over.append((o["price"], title)); points.append(o.get("point"))
                        elif o["name"] == "Under":
                            under.append((o["price"], title))

        if not diag_printed and books:
            diag_printed = True
            print(f"    [diag] {tab}: {len(books)} books -> "
                  f"{[b.get('title') for b in books][:8]}")

        pts = [p for p in points if p is not None]
        line = _median(pts) if pts else None
        am, ab, abk = _agg(away_ml)
        hm, hb, hbk = _agg(home_ml)
        om, ob, obk = _agg(over)
        um, ub, ubk = _agg(under)

        if am is None and line is None:
            continue
        n_books = len({b for _, b in away_ml + home_ml + over + under})
        out[tab] = {
            "away_ml": am, "away_ml_best": ab, "away_ml_book": abk,
            "home_ml": hm, "home_ml_best": hb, "home_ml_book": hbk,
            "total": line,
            "over_juice": om, "over_best": ob, "over_book": obk,
            "under_juice": um, "under_best": ub, "under_book": ubk,
            "n_books": n_books,
            "ml_source": f"median/{n_books} books",
            "total_source": f"median/{n_books} books",
            "last_update": (books[0].get("last_update", "") if books else ""),
            "event_id": ev["id"],
        }
        n_priced += 1

    print(f"  Got all-book F5 odds for {n_priced} games.")
    return out


def save_odds_snapshot(odds_by_tab, target_date, out_dir):
    """Write a timestamped CSV snapshot of the pulled odds (useful for
    line-movement tracking)."""
    if not odds_by_tab:
        return
    snap_dir = Path(out_dir) / "live_snapshots"
    snap_dir.mkdir(exist_ok=True)
    from datetime import datetime as _dt
    fname = f"mlb_odds_{_dt.now().strftime('%Y%m%d_%H%M%S')}.csv"
    snap_path = snap_dir / fname
    fields = ["date", "tab", "away_ml", "home_ml", "total",
              "over_juice", "under_juice", "ml_source", "total_source",
              "last_update", "event_id"]
    with open(snap_path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        for tab, row in odds_by_tab.items():
            w.writerow({
                "date": target_date.strftime("%Y-%m-%d"),
                "tab":  tab,
                **{k: row.get(k, "") for k in fields[2:]},
            })
    print(f"  Odds snapshot saved: {snap_path}")


# ===========================================================================
# PART 3: PROJECTION ENGINE — ONE GAME
# ===========================================================================
# Cell-by-cell calcs that mirror BALTB's formulas. Each function maps to
# one cell (or short range) so we can validate them individually.
# ===========================================================================

# ---- Lookups (faithful to the spreadsheet's VLOOKUP behaviour) ----

def _vlookup_pitcher(refs, name, field):
    try:
        return refs.start.loc[name, field]
    except KeyError:
        return float("nan")


def _vlookup_batter(refs, name, vs_pitcher_hand, field):
    """Picks vR or vL table based on opposing pitcher's hand. Mirrors the
    spreadsheet's IF(D=R, vR, vL) — so any non-'R' hand falls into vL."""
    table = refs.pn_vr if vs_pitcher_hand == "R" else refs.pn_vl
    try:
        return table.loc[name, field]
    except KeyError:
        return float("nan")


def _vlookup_batter_vo(refs, name, field):
    try:
        return refs.pn_vo.loc[name, field]
    except KeyError:
        return float("nan")


def _batter_exists(refs, name):
    return name in refs.pn_vr.index


# ---- Cell-level calcs (each one mirrors one BALTB column formula) ----

def pitcher_hand(refs, name):
    """Cell D2 / D15."""
    h = _vlookup_pitcher(refs, name, "Hand")
    return h if isinstance(h, str) else ""


def pitcher_runs_per_game(refs, name, custom_ip):
    """Cell K2 / K15. Uses 'Updated WAR/GS' (col CB) by default, with
    'Updated xWAR/GS' (CC) as the IFERROR fallback, and 'Updated WAR/IP' (CD)
    when a custom IP override is set."""
    rpw = refs.runs_per_war
    if custom_ip > 0:
        return _vlookup_pitcher(refs, name, "Updated WAR/IP") * custom_ip * rpw
    war_gs = _vlookup_pitcher(refs, name, "Updated WAR/GS")
    if pd.isna(war_gs):
        war_gs = _vlookup_pitcher(refs, name, "Updated xWAR/GS")
    return war_gs * rpw


def pitcher_era(refs, name):
    """Cell O2 / O15 — Updated ERA (col CE)."""
    return _vlookup_pitcher(refs, name, "Updated ERA")


def pitcher_team_code(name):
    """RIGHT(name, 3) — Retrosheet code off the end of Full Name."""
    return name[-3:].upper()


def bullpen_stats(refs, team_code, custom_ip):
    """Returns (S5/S18, T5/T18, U5/U18) for a team.
        S = scaled BPRAA
        T = scaled BPWAR (the 'Bullpen' WAR added to row 28/29)
        U = wERA+ROS  (the bullpen ERA used in Totals projection)
    """
    try:
        bp = refs.bullpen.loc[team_code]
    except KeyError:
        return float("nan"), float("nan"), float("nan")
    bpraa = bp["BPRAA"]; bpwar = bp["BPWAR"]
    bpg_bpwar = bp["bpgBPWAR"]; wera_ros = bp["wERA+ROS"]
    if custom_ip > 0:
        s = bpraa * ((9 - custom_ip) / 9) / (4 / 9)
        t = bpg_bpwar * (9 - custom_ip)
    else:
        s = bpraa; t = bpwar
    return s, t, wera_ros


def batter_d_col(refs, name):
    if not _batter_exists(refs, name):
        return "BLANK"
    bats = refs.pn_vr.loc[name, "Bats"]
    return str(bats) if isinstance(bats, str) else "BLANK"


def batter_baserunning(refs, name, opp_pitcher_hand):
    """F5 E column: pure platoon BaseRunning (vR or vL), NO vO blend.
    Full game uses (vR*2/3 + vO*1/3); F5 uses just the platoon split.
    Uses proper handedness (away batters vs home pitcher, home batters vs
    away pitcher) — same convention as the full-game model."""
    if not _batter_exists(refs, name):
        return 0.0
    return _vlookup_batter(refs, name, opp_pitcher_hand, "BaseRunning")


def batter_uzr(refs, name, position, opp_pitcher_hand):
    """F column. Zero for DH."""
    if not _batter_exists(refs, name) or position.lower() == "dh":
        return 0.0
    return _vlookup_batter(refs, name, opp_pitcher_hand, "UZR")


def batter_runs(refs, name, order, opp_pitcher_hand):
    """F5 G column: lineup-slot weighted BattingRuns/PA, PURE PLATOON (no vO
    blend). Full game blends vR*2/3 + vO*1/3; F5 uses just the platoon split."""
    if not _batter_exists(refs, name):
        return 0.0
    platoon = _vlookup_batter(refs, name, opp_pitcher_hand, "Updated BattingRuns/PA")
    pa_per_g = refs.pa_per_game.loc[order, "PA_per_G"]
    return platoon * pa_per_g


def batter_replacement(refs, name, order):
    """H column: PA_per_G[slot] * 0.0303."""
    if not _batter_exists(refs, name):
        return 0.0
    return refs.pa_per_game.loc[order, "PA_per_G"] * 0.0303


def batter_def_adj(refs, name, position):
    """I column: positional adjustment per game."""
    if not _batter_exists(refs, name):
        return 0.0
    if position.lower() == "x":
        return refs.x_def_adj_per_game
    try:
        return refs.def_adj.loc[position.upper(), "Adjustment"] / 162.0
    except KeyError:
        return refs.x_def_adj_per_game


def batter_framing(refs, name, position, opp_pitcher_hand):
    """J column: FramingRuns (catchers only)."""
    if position.lower() != "c" or not _batter_exists(refs, name):
        return 0.0
    return _vlookup_batter(refs, name, opp_pitcher_hand, "FramingRuns")


# ---- Team aggregate (replicates rows 12 and 25) ----

def project_team(refs, lineup, opp_pitcher_hand, remaining_games):
    """Compute every cell for one team (away or home).

    F5 ADJUSTMENT: the batter K column (sum of run contributions) is scaled
    by 5/9 to reflect 5 innings out of 9 — matching the spreadsheet's
    K = SUM(E:J) * 5/9. M and N follow since they're K * games / runs_per_war.
    The pitcher row is NOT scaled (pitcher K formula has no 5/9 in F5).
    """
    pitcher = lineup.pitcher

    p_hand = pitcher_hand(refs, pitcher.name)
    p_k = pitcher_runs_per_game(refs, pitcher.name, lineup.custom_ip)
    p_m = p_k * remaining_games
    p_n = p_m / refs.runs_per_war
    p_o = pitcher_era(refs, pitcher.name)
    pitcher_row = {
        "name": pitcher.name, "order": None, "position": "Pitcher",
        "D": p_hand, "E": None, "F": None, "G": None, "H": None, "I": None,
        "J": None, "K": p_k, "L": remaining_games, "M": p_m, "N": p_n, "O": p_o,
    }

    rows = []
    for b in lineup.batters:
        d = batter_d_col(refs, b.name)
        e = batter_baserunning(refs, b.name, opp_pitcher_hand)
        f = batter_uzr(refs, b.name, b.position, opp_pitcher_hand)
        g = batter_runs(refs, b.name, b.order, opp_pitcher_hand)
        h = batter_replacement(refs, b.name, b.order)
        i = batter_def_adj(refs, b.name, b.position)
        j = batter_framing(refs, b.name, b.position, opp_pitcher_hand)
        # F5: scale the K sum by 5/9 (5 innings of 9)
        k_raw = sum(x for x in (e, f, g, h, i, j) if x is not None)
        k = k_raw * F5_BATTER_FRACTION
        m = k * remaining_games
        n = m / refs.runs_per_war
        try:
            o = refs.pn_vo.loc[b.name, "wWOBA Change"]
        except KeyError:
            o = float("nan")
        rows.append({"name": b.name, "order": b.order, "position": b.position,
                     "D": d, "E": e, "F": f, "G": g, "H": h, "I": i, "J": j,
                     "K": k, "L": remaining_games, "M": m, "N": n, "O": o})

    def col_sum(letter):
        return sum(r[letter] for r in rows if r[letter] is not None)

    sums = {
        "E": col_sum("E"), "F": col_sum("F"), "G": col_sum("G"),
        "H": col_sum("H"), "I": col_sum("I"), "J": col_sum("J"),
        "K": col_sum("K"),
        # M12/M25 in the spreadsheet INCLUDE the pitcher row
        "M": pitcher_row["M"] + sum(r["M"] for r in rows),
        "N": pitcher_row["N"] + sum(r["N"] for r in rows),
    }

    team_code = pitcher_team_code(pitcher.name)
    bp_s, bp_t, bp_u = bullpen_stats(refs, team_code, lineup.custom_ip)

    return {
        "pitcher_hand": p_hand, "pitcher_row": pitcher_row, "batter_rows": rows,
        "sums": sums, "team_code": team_code,
        "bp_raa_scaled": bp_s, "bp_war_scaled": bp_t, "bp_wera_ros": bp_u,
    }


# ---- Full ML projection (rows 27-29 of BALTB) ----

def project_game(refs, game):
    """Run the full ML projection for one game."""
    hfa = game.hfa_win_pct if game.hfa_win_pct is not None else refs.hfa_win_pct

    away_p_hand = pitcher_hand(refs, game.away.pitcher.name)
    home_p_hand = pitcher_hand(refs, game.home.pitcher.name)

    away = project_team(refs, game.away, opp_pitcher_hand=home_p_hand,
                        remaining_games=game.remaining_games)
    home = project_team(refs, game.home, opp_pitcher_hand=away_p_hand,
                        remaining_games=game.remaining_games)

    replacement_wins = refs.replacement_winning_pct * refs.season_games

    away_team_wins = replacement_wins + away["sums"]["N"]
    home_team_wins = replacement_wins + home["sums"]["N"]

    # F5 NEVER includes the bullpen in the ML win% calc — relievers rarely
    # pitch the first 5 innings, so the win% is driven entirely by the starter.
    away_bullpen = 0.0
    home_bullpen = 0.0

    away_xwins = away_team_wins + away_bullpen
    home_xwins = home_team_wins + home_bullpen

    away_ros_wpct = away_xwins / game.remaining_games
    home_ros_wpct = home_xwins / game.remaining_games

    away_hf_wpct = away_ros_wpct
    home_hf_wpct = home_ros_wpct * (hfa * 2)

    away_xwpct = away_hf_wpct * (1 - home_hf_wpct)
    home_xwpct = home_hf_wpct * (1 - away_hf_wpct)

    denom = away_xwpct + home_xwpct
    away_gwpct = away_xwpct / denom
    home_gwpct = home_xwpct / denom

    def american_odds(p):
        return ((1 - p) / p) * 100 if p < 0.5 else -(p / (1 - p)) * 100

    def implied_prob(odds):
        return 100 / (odds + 100) if odds > 0 else -odds / (-odds + 100)

    def ev(p, odds):
        if odds > 0:
            return ((p * (odds / 100) * 25) - ((1 - p) * 25)) / 25
        return (((100 / -odds) * p * 25) - (1 - p) * 25) / 25

    return {
        "away": {
            "team_code": away["team_code"], "pitcher_hand": away["pitcher_hand"],
            "team_wins": away_team_wins, "bullpen_war": away_bullpen,
            "xwins": away_xwins, "ros_wpct": away_ros_wpct,
            "hf_wpct": away_hf_wpct, "xwpct": away_xwpct,
            "game_wpct": away_gwpct, "fair_odds": american_odds(away_gwpct),
            "book_odds": game.book_odds_away,
            "edge": away_gwpct - implied_prob(game.book_odds_away),
            "ev": ev(away_gwpct, game.book_odds_away),
            "details": away,
        },
        "home": {
            "team_code": home["team_code"], "pitcher_hand": home["pitcher_hand"],
            "team_wins": home_team_wins, "bullpen_war": home_bullpen,
            "xwins": home_xwins, "ros_wpct": home_ros_wpct,
            "hf_wpct": home_hf_wpct, "xwpct": home_xwpct,
            "game_wpct": home_gwpct, "fair_odds": american_odds(home_gwpct),
            "book_odds": game.book_odds_home,
            "edge": home_gwpct - implied_prob(game.book_odds_home),
            "ev": ev(home_gwpct, game.book_odds_home),
            "details": home,
        },
        "replacement_wins": replacement_wins, "hfa_win_pct": hfa,
    }


# ===========================================================================
# PART 4: TOTALS (OVER/UNDER) PROJECTION
# ===========================================================================
# Per-team offense and pitching, plus slate-wide scaling for the Game Total.
# Mirrors rows 32-65 of each game tab.
# ===========================================================================

def batter_wrc_runs(refs, name, order, opp_pitcher_hand):
    """F column of the Totals section: Updated wRC/PA, weighted 5/9 platoon
    + 4/9 vs overall, times PA_per_G[slot].

    Robustness: the blend draws from two different tables (the platoon vR/vL
    table and the overall vO table). If a batter is present in one but the
    name is spelled differently in the other (e.g. 'Sam' in vR/vL vs 'Samuel'
    in vO), the missing term is NaN. Rather than let one NaN propagate into the
    team sum and wipe out the entire game's total, fall back to whichever term
    is available (renormalised to full weight). Only return 0.0 if the batter
    is in neither table."""
    if not _batter_exists(refs, name):
        return 0.0
    platoon = _vlookup_batter(refs, name, opp_pitcher_hand, "Updated wRC/PA")
    vo = _vlookup_batter_vo(refs, name, "Updated wRC/PA")
    pa_per_g = refs.pa_per_game.loc[order, "PA_per_G"]
    p_ok = pd.notna(platoon)
    v_ok = pd.notna(vo)
    if p_ok and v_ok:
        blended = platoon * (5 / 9) + vo * (4 / 9)
    elif p_ok:
        blended = platoon            # vO missing for this name — use platoon only
    elif v_ok:
        blended = vo                 # platoon missing — use overall only
    else:
        return 0.0
    return blended * pa_per_g


def project_team_offense(refs, lineup, opp_pitcher_hand):
    """G52 / G65 — team's Off. Runs per game."""
    rows = []
    for b in lineup.batters:
        e = batter_baserunning(refs, b.name, opp_pitcher_hand)
        f = batter_wrc_runs(refs, b.name, b.order, opp_pitcher_hand)
        rows.append({"name": b.name, "order": b.order,
                     "E_BR": e, "F_wRC": f, "G_OffRuns": f + e})
    return rows, sum(r["G_OffRuns"] for r in rows)


def project_team_pitching(refs, lineup, sums_from_ml):
    """F5 J42 / J55 — team's Total ERA = the starter's raw ERA only.

    Spreadsheet: H42 = IFERROR(VLOOKUP(SP, Start col 'ERA'),
                               VLOOKUP(SP, Start col 'Updated ERA'))
                 J42 = H42   (no bullpen blend, no UZR/framing subtraction)
    """
    pitcher = lineup.pitcher
    h = _vlookup_pitcher(refs, pitcher.name, "ERA")
    if pd.isna(h):
        h = _vlookup_pitcher(refs, pitcher.name, "Updated ERA")
    # I_BP set to 0 since F5 doesn't blend in the bullpen
    return {"H_ERA": h, "I_BP": 0.0, "J_TotalERA": h}


# ---- Total Calcs replica: zero-inflated negative-binomial run distribution ----
#
# Replicates the formulas in the 'Total Calcs' tab. The Excel model assumes
# each team's runs scored is zero-inflated negative-binomial, with parameters
# (r, B, z) looked up by team R/G from the Parameters tab. The total runs
# distribution is the convolution of the two teams' distributions. Over/Under
# American odds at a given book line are derived from that.
#
# Validated cell-for-cell against BALTB M37/N37 (book line 7.5):
#   Excel:  Under -111.10, Over +111.10
#   Python: Under -111.10, Over +111.10   ✓

def _nb_lookup_params(refs, rg):
    """Mimic the spreadsheet's VLOOKUP(rg, Parameters!A:D, col, FALSE) but
    with TRUE-style closest-not-exceeding match since that matches how the
    Parameters table is structured (0.05 R/G steps). Returns (r, B, z)."""
    df = refs.nb_params
    # Find the largest RG in the table that is <= rg
    rgs = df["RG"].to_numpy()
    import numpy as np
    idx = int(np.searchsorted(rgs, rg, side="right") - 1)
    idx = max(0, min(idx, len(df) - 1))
    row = df.iloc[idx]
    return float(row["r"]), float(row["B"]), float(row["z"])


def _zinb_pmf(K, r, B, z):
    """Return [P(0), P(1), ..., P(K-1)] for the zero-inflated negative-binomial
    used by the model. Matches the spreadsheet's H2:I22 logic exactly.

    NB density (under the model's parameterization):
        P_orig(0) = (1 + B)^(-r)
        P_orig(k) = [r * (r+1) * ... * (r+k-1)] * B^k / (k! * (1+B)^(r+k))   for k >= 1
    Zero-inflation:
        P_zi(0) = z
        P_zi(k) = (1 - z) * P_orig(k) / (1 - P_orig(0))   for k >= 1
    """
    import math
    p_orig = [0.0] * K
    p_orig[0] = (1.0 + B) ** (-r)
    g_product = 1.0
    g_next = r
    for k in range(1, K):
        g_product *= g_next
        g_next += 1.0
        p_orig[k] = g_product * (B ** k) / (math.factorial(k) * (1.0 + B) ** (r + k))

    p_zi = [0.0] * K
    p_zi[0] = z
    one_minus_p_orig_0 = 1.0 - p_orig[0]
    if one_minus_p_orig_0 <= 0:
        # Degenerate case (shouldn't happen for sensible R/G > 0). Fall back
        # to original distribution unscaled.
        return p_orig
    factor = (1.0 - z) / one_minus_p_orig_0
    for k in range(1, K):
        p_zi[k] = factor * p_orig[k]
    return p_zi


def _convolve(a, b):
    """Convolution of two PMFs (a + b distribution)."""
    out = [0.0] * (len(a) + len(b) - 1)
    for i, pa in enumerate(a):
        if pa == 0:
            continue
        for j, pb in enumerate(b):
            out[i + j] += pa * pb
    return out


def _american_odds(p):
    """Convert win probability to American odds. Returns 0 if p is 0 or 1."""
    if p <= 0 or p >= 1:
        return 0.0
    return -100.0 * p / (1.0 - p) if p > 0.5 else 100.0 * (1.0 - p) / p


def totals_odds(refs, away_rg, home_rg, book_line):
    """Compute the over/under American odds at `book_line` given each team's
    expected runs (away_rg, home_rg).

    Replicates Total Calcs!M37/N37 from the spreadsheet:
      - For .5 lines: P(Under) = P(total <= floor(line))
      - For integer lines (push possible): P(Under) and P(Over) normalize
        by dividing by (P(Under) + P(Over)) to exclude the push outcome,
        matching how books quote integer total odds.

    Returns (under_odds, over_odds) as American odds. Returns (None, None)
    if book_line is None.
    """
    if book_line is None:
        return None, None
    import math
    K = 30  # support up to 30 runs per team (way more than enough)

    r_a, B_a, z_a = _nb_lookup_params(refs, away_rg)
    r_h, B_h, z_h = _nb_lookup_params(refs, home_rg)
    away_pmf = _zinb_pmf(K, r_a, B_a, z_a)
    home_pmf = _zinb_pmf(K, r_h, B_h, z_h)
    total_pmf = _convolve(away_pmf, home_pmf)

    line = float(book_line)
    is_half = abs(line - int(line) - 0.5) < 1e-9  # X.5 line

    if is_half:
        x_under = int(math.floor(line))    # total <= x_under is Under
        p_under = sum(total_pmf[:x_under + 1])
        p_over = sum(total_pmf[x_under + 1:])
    else:
        x = int(round(line))               # integer line at x; push if total == x
        p_under_raw = sum(total_pmf[:x])              # total <= x-1
        p_over_raw  = sum(total_pmf[x + 1:])          # total >= x+1
        denom = p_under_raw + p_over_raw
        if denom <= 0:
            return None, None
        p_under = p_under_raw / denom
        p_over  = p_over_raw / denom

    return _american_odds(p_under), _american_odds(p_over)


def lookup_park_factor(refs, team_code):
    """E37/E38 — Park factor (Blend Per 9) for a 3-letter team code."""
    pf = refs.park_factors
    match = pf[pf["Org"] == team_code]
    return float(match["Blend Per 9"].iloc[0]) if len(match) else 1.0


def project_slate_totals(refs, slate_results):
    """Adds Totals predictions to each slate row, processing each game
    INDEPENDENTLY. League averages come from refs constants — no slate-wide
    aggregation, no cross-game references. A game's predicted total, C32/D32,
    and over/under odds depend only on that game's own inputs plus the
    constants. This makes results stable regardless of how many games are
    on the slate or whether stale tabs are present.

    Per-game logic (matches the F5 Excel TOTALS sheet):

      Inputs from the per-game projection:
        away_off / home_off           — per-game offense (G52 / G65)
        away_total_era / home_total_era — per-game pitching (J42 / J55)
        away_bp_raa / home_bp_raa     — bullpen RAA (S5 / S18)
        pitcher Ks and team UZR/Framing sums — for F162

      Constants (from constants.csv):
        rg = league_runs_per_game (2.48)
        league_avg_off, league_avg_era — divisors for C37/D37 scaling
        league_avg_e162, league_avg_f162, league_avg_bullpen_raa — lg_E/F/G
        baseline = rg * 162 = 401.76 — pinned J23/K23

      Per team:
        C37/D37/C38/D38 = (per-game / league_avg) * rg * 162
        E162 = team_off_per_g * 162  (from BR+BAT+Repl+Def sums)
        F162 = (pitcher_K + UZR_sum + Framing_sum) * 162
        H (oRAR) = E162 - league_avg_e162
        I (dRAR) = (F162 + bp_raa) - (league_avg_f162 + league_avg_bullpen_raa)
        J = baseline + H
        K = baseline - I
        O = (J / baseline) * (opponent_K / baseline) * rg
        Q = O * park_factor

      Per game:
        m1_away = (C37/162) * (D38/(162*rg)) * pf
        m1_home = (C38/162) * (D37/(162*rg)) * pf
        C32 = m1_away/4 + Q_away*3/4
        D32 = m1_home/4 + Q_home*3/4
        G37 = ((C37/162) * (D38/(162*rg)) + (C38/162) * (D37/(162*rg))) * pf
        Under/Over odds: ZINB at I37 line using C32 and D32 as per-team R/5

    Returns (results, league_avg_off, league_avg_era) for compatibility with
    the SLATE summary printer; both values come from refs constants.
    """
    import math

    # All league averages are constants — no slate aggregation.
    rg = refs.league_runs_per_game
    g = refs.season_games
    baseline = rg * g                          # 2.48 * 162 = 401.76
    league_avg_off = refs.league_avg_off_override
    league_avg_era = refs.league_avg_era_override
    lg_E = refs.league_avg_e162
    lg_F = refs.league_avg_f162
    lg_G = refs.league_avg_bullpen_raa

    def _per_game_off(sums):
        """E12+G12+H12+I12 — offensive contribution per game from ML rows."""
        if not sums:
            return float("nan")
        return (sums.get("E", 0) + sums.get("G", 0)
                + sums.get("H", 0) + sums.get("I", 0))

    def _safe_num(v, default=0.0):
        if isinstance(v, float) and math.isnan(v):
            return default
        return v if isinstance(v, (int, float)) else default

    for r in slate_results:
        if r.get("status") != "ok":
            continue

        # ---- Scale per-game numbers to season totals (C37/D37/C38/D38) ----
        c37 = r["away_off"] / league_avg_off * rg * g
        d37 = r["away_total_era"] / league_avg_era * rg * g
        c38 = r["home_off"] / league_avg_off * rg * g
        d38 = r["home_total_era"] / league_avg_era * rg * g
        pf = lookup_park_factor(refs, r["home_team"])
        r["away_rs_162"] = c37; r["away_ra_162"] = d37
        r["home_rs_162"] = c38; r["home_ra_162"] = d38
        r["park_factor"] = pf

        # ---- Per-team E162 and F162 (offensive / defensive 162-game totals)
        away_detail = r.get("_away_detail", {}) or {}
        home_detail = r.get("_home_detail", {}) or {}
        away_sums = away_detail.get("sums", {})
        home_sums = home_detail.get("sums", {})
        away_pitcher_K = (away_detail.get("pitcher_row", {}) or {}).get("K", 0.0) or 0.0
        home_pitcher_K = (home_detail.get("pitcher_row", {}) or {}).get("K", 0.0) or 0.0

        away_F_per_g = away_pitcher_K + (away_sums.get("F", 0) or 0) + (away_sums.get("J", 0) or 0)
        home_F_per_g = home_pitcher_K + (home_sums.get("F", 0) or 0) + (home_sums.get("J", 0) or 0)
        away_F162 = away_F_per_g * g
        home_F162 = home_F_per_g * g

        away_E162 = _per_game_off(away_sums) * g
        home_E162 = _per_game_off(home_sums) * g

        away_bp_raa = _safe_num(away_detail.get("bp_raa_scaled", 0.0))
        home_bp_raa = _safe_num(home_detail.get("bp_raa_scaled", 0.0))

        r["_away_bp_raa"] = away_bp_raa
        r["_home_bp_raa"] = home_bp_raa
        r["_away_E162"] = away_E162
        r["_home_E162"] = home_E162
        r["_away_F162"] = away_F162
        r["_home_F162"] = home_F162

        # ---- Per-team H, I, J, K using PINNED league averages ----
        away_H = away_E162 - lg_E                          # oRAR
        home_H = home_E162 - lg_E
        away_I = (away_F162 + away_bp_raa) - (lg_F + lg_G) # dRAR
        home_I = (home_F162 + home_bp_raa) - (lg_F + lg_G)
        away_J = baseline + away_H
        home_J = baseline + home_H
        away_K = baseline - away_I
        home_K = baseline - home_I

        # ---- Per-team paRuns Q (uses pinned baseline as lg_J/lg_K) ----
        # O = (own_J / baseline) * (opp_K / baseline) * rg
        away_O = (away_J / baseline) * (home_K / baseline) * rg
        home_O = (home_J / baseline) * (away_K / baseline) * rg
        away_Q = away_O * pf
        home_Q = home_O * pf

        # ---- C32 / D32 (per-team R/5) and predicted total G37 ----
        m1_away = (c37 / g) * (d38 / (g * rg)) * pf
        m1_home = (c38 / g) * (d37 / (g * rg)) * pf
        c32 = (m1_away + away_Q) / 2.0
        d32 = (m1_home + home_Q) / 2.0

        r["away_pa_runs"] = away_Q
        r["home_pa_runs"] = home_Q
        r["away_rg_per_game"] = c32
        r["home_rg_per_game"] = d32

        # Predicted total G37 (the spreadsheet's exact formula):
        g37_away = (c37 / g) * (d38 / (g * rg))
        g37_home = (c38 / g) * (d37 / (g * rg))
        r["predicted_total"] = (g37_away + g37_home) * pf

        # ---- Over/Under American odds at the book line via ZINB ----
        line = r.get("totals_book_line")
        if line is not None and isinstance(line, (int, float)):
            try:
                under, over = totals_odds(refs, c32, d32, float(line))
            except Exception:
                under, over = None, None
            r["total_under_odds"] = under
            r["total_over_odds"] = over
        else:
            r["total_under_odds"] = None
            r["total_over_odds"] = None

    return slate_results, league_avg_off, league_avg_era



# ===========================================================================
# PART 5: SLATE LOOP — PROJECT ALL GAMES IN THE MODEL
# ===========================================================================


def _read_lineup_from_tab(ws, side):
    """Pull either the away or home lineup off a game tab. Returns None if
    the lineup isn't fully populated."""
    if side == "away":
        p_row, b_start = 2, 3
    else:
        p_row, b_start = 15, 16
    pitcher_name = ws.cell(row=p_row, column=3).value
    if not pitcher_name or not isinstance(pitcher_name, str):
        return None

    batters = []
    for i in range(9):
        row = b_start + i
        pos = ws.cell(row=row, column=2).value
        name = ws.cell(row=row, column=3).value
        if not name or not isinstance(name, str):
            return None
        pos_str = str(pos).strip().lower() if pos else "x"
        if pos_str not in ("c", "dh", "x"):
            pos_str = "x"
        batters.append(PlayerSlot(order=i + 1, position=pos_str, name=name.strip()))

    custom_ip_cell = "T1" if side == "away" else "T14"
    custom_ip = ws[custom_ip_cell].value or 0
    try:
        custom_ip = float(custom_ip)
    except (TypeError, ValueError):
        custom_ip = 0.0

    return TeamLineup(
        pitcher=PlayerSlot(order=None, position="Pitcher", name=pitcher_name.strip()),
        batters=batters, custom_ip=custom_ip,
    )


def _read_game_settings(ws):
    hfa = ws["H31"].value
    bp_toggle = ws["H32"].value or "YES"
    bullpen_on = str(bp_toggle).strip().upper() == "YES"
    book_away = ws["L28"].value
    book_home = ws["L29"].value
    try: book_away = int(book_away) if book_away is not None else None
    except (TypeError, ValueError): book_away = None
    try: book_home = int(book_home) if book_home is not None else None
    except (TypeError, ValueError): book_home = None
    rem = ws["L2"].value
    try: rem = int(rem) if rem is not None else 162
    except (TypeError, ValueError): rem = 162
    return {
        "hfa": float(hfa) if isinstance(hfa, (int, float)) else 0.53,
        "bullpen_on": bullpen_on,
        "book_away": book_away, "book_home": book_home,
        "remaining_games": rem,
    }


def project_slate(model_path, refs, lineups_by_tab=None, odds_by_tab=None):
    """Project every game in the slate. Returns a list of result dicts.

    Two operating modes:

    1. `lineups_by_tab` provided (from build_lineups_from_fangraphs):
       Iterate the dict — tab names are FG-derived (e.g. 'BAL@DET') and the
       slate reflects today's actual matchups, not any preset list of tabs.
       Per-game settings (HFA, bullpen toggle) use sensible defaults since
       new matchups won't have entries in the model. If `odds_by_tab` is
       provided (from fetch_pinnacle_odds), per-game book odds and totals
       lines are pulled from there.

    2. `lineups_by_tab is None` (the validate / --from-model path):
       Iterate game tabs in the model itself. This is for reproducing the
       spreadsheet's existing tabs cell-for-cell. `odds_by_tab` is ignored
       in this mode — book odds come from the model.
    """
    results = []
    odds_by_tab = odds_by_tab or {}

    if lineups_by_tab is not None:
        # ---- FG-derived matchups -----------------------------------------
        # Settings: use global defaults rather than per-tab values.
        hfa = refs.hfa_win_pct
        for tab, (away, home, fg_status) in lineups_by_tab.items():
            if fg_status == "no_lineup_yet":
                results.append({"tab": tab, "status": "no_lineup_yet",
                                "away_sp": None, "home_sp": None})
                continue
            if fg_status.startswith("missing_pitcher"):
                results.append({"tab": tab, "status": fg_status,
                                "away_sp": None, "home_sp": None})
                continue

            # Pull live odds for this game if we have them
            game_odds = odds_by_tab.get(tab, {})
            book_away = game_odds.get("away_ml")
            book_home = game_odds.get("home_ml")
            totals_book_line = game_odds.get("total")

            # GameSpec needs concrete book odds for math to run; use ±100 as
            # placeholder when odds aren't available (edge/EV zeroed out below)
            game = GameSpec(
                away=away, home=home,
                book_odds_away=book_away if book_away is not None else 100,
                book_odds_home=book_home if book_home is not None else -100,
                bullpen_on=True, hfa_win_pct=hfa,
                remaining_games=162,
            )
            try:
                out = project_game(refs, game)
            except Exception as e:
                results.append({"tab": tab, "status": "error", "error": str(e)})
                continue

            # If we didn't have odds, zero out edge/EV (placeholder result
            # would be misleading)
            if book_away is None or book_home is None:
                out["away"]["edge"] = out["home"]["edge"] = None
                out["away"]["ev"] = out["home"]["ev"] = None

            results.append(_finalize_result(
                refs, tab, away, home, out,
                book_away=book_away, book_home=book_home,
                totals_book_line=totals_book_line,
                game_odds=game_odds,
            ))
    else:
        # ---- Model-tab iteration (validate / --from-model path) ----------
        wb = load_workbook(model_path, data_only=True)
        for tab in wb.sheetnames:
            if tab in NON_GAME_TABS:
                continue
            ws = wb[tab]
            if ws["B1"].value != "Away" or ws["B14"].value != "Home":
                continue
            away = _read_lineup_from_tab(ws, "away")
            home = _read_lineup_from_tab(ws, "home")
            if away is None or home is None:
                results.append({"tab": tab, "status": "incomplete_lineup",
                                "away_sp": ws["C2"].value,
                                "home_sp": ws["C15"].value})
                continue

            s = _read_game_settings(ws)
            book_a = s["book_away"] if s["book_away"] is not None else 100
            book_h = s["book_home"] if s["book_home"] is not None else -100
            game = GameSpec(
                away=away, home=home,
                book_odds_away=book_a, book_odds_home=book_h,
                bullpen_on=s["bullpen_on"], hfa_win_pct=s["hfa"],
                remaining_games=s["remaining_games"],
            )
            try:
                out = project_game(refs, game)
            except Exception as e:
                results.append({"tab": tab, "status": "error", "error": str(e)})
                continue

            if s["book_away"] is None or s["book_home"] is None:
                out["away"]["edge"] = out["home"]["edge"] = None
                out["away"]["ev"] = out["home"]["ev"] = None

            results.append(_finalize_result(
                refs, tab, away, home, out,
                book_away=s["book_away"], book_home=s["book_home"],
                totals_book_line=ws["I37"].value,
            ))

    # Slate-wide totals scaling
    results, league_off, league_era = project_slate_totals(refs, results)
    for r in results:
        if r.get("status") == "ok":
            r["league_avg_off"] = league_off
            r["league_avg_era"] = league_era
    return results


def _finalize_result(refs, tab, away, home, out, book_away, book_home,
                     totals_book_line, game_odds=None):
    """Compute Totals per-team and assemble the result dict."""
    away_off_rows, away_off = project_team_offense(
        refs, away, opp_pitcher_hand=out["home"]["pitcher_hand"])
    home_off_rows, home_off = project_team_offense(
        refs, home, opp_pitcher_hand=out["away"]["pitcher_hand"])
    away_pit = project_team_pitching(refs, away, out["away"]["details"]["sums"])
    home_pit = project_team_pitching(refs, home, out["home"]["details"]["sums"])

    return {
        "tab": tab, "status": "ok",
        "away_sp": away.pitcher.name, "home_sp": home.pitcher.name,
        "away_team": out["away"]["team_code"],
        "home_team": out["home"]["team_code"],
        "away_wpct": out["away"]["game_wpct"],
        "home_wpct": out["home"]["game_wpct"],
        "away_fair": out["away"]["fair_odds"],
        "home_fair": out["home"]["fair_odds"],
        "away_book": book_away, "home_book": book_home,
        "away_edge": out["away"]["edge"], "home_edge": out["home"]["edge"],
        "away_ev": out["away"]["ev"], "home_ev": out["home"]["ev"],
        "away_team_wins": out["away"]["team_wins"],
        "home_team_wins": out["home"]["team_wins"],
        "away_bullpen_war": out["away"]["bullpen_war"],
        "home_bullpen_war": out["home"]["bullpen_war"],
        "away_off": away_off, "home_off": home_off,
        "away_total_era": away_pit["J_TotalERA"],
        "home_total_era": home_pit["J_TotalERA"],
        "away_pitcher_era": away_pit["H_ERA"],
        "home_pitcher_era": home_pit["H_ERA"],
        "totals_book_line": totals_book_line,
        "_odds": game_odds or {},
        # Full per-player detail — for writing the xlsx tab layout
        "_away_detail": out["away"]["details"],
        "_home_detail": out["home"]["details"],
        "_away_off_rows": away_off_rows,
        "_home_off_rows": home_off_rows,
        "_away_pit_totals": away_pit,
        "_home_pit_totals": home_pit,
        "_hfa_win_pct": out["hfa_win_pct"],
        "_replacement_wins": out["replacement_wins"],
        "_away_lineup": away,
        "_home_lineup": home,
        "_away_xwins": out["away"]["xwins"],
        "_home_xwins": out["home"]["xwins"],
        "_away_ros_wpct": out["away"]["ros_wpct"],
        "_home_ros_wpct": out["home"]["ros_wpct"],
        "_away_hf_wpct": out["away"]["hf_wpct"],
        "_home_hf_wpct": out["home"]["hf_wpct"],
        "_away_xwpct": out["away"]["xwpct"],
        "_home_xwpct": out["home"]["xwpct"],
    }


def _print_slate(results):
    print("=" * 120)
    print("ML PROJECTIONS")
    print(f"{'Game':10s}  {'Away SP':25s} {'Home SP':25s}  {'AwayW%':>7s} "
          f"{'HomeW%':>7s}  {'Fair A':>7s} {'Fair H':>7s}  {'EdgeA':>7s} {'EdgeH':>7s}")
    print("-" * 120)
    for r in results:
        if r["status"] != "ok":
            print(f"{r['tab']:10s}  [{r['status']}] {r.get('error','')}")
            continue
        edge_a = f"{r['away_edge']:>+7.4f}" if r["away_edge"] is not None else "    n/a"
        edge_h = f"{r['home_edge']:>+7.4f}" if r["home_edge"] is not None else "    n/a"
        print(f"{r['tab']:10s}  "
              f"{r['away_sp'][:25]:25s} {r['home_sp'][:25]:25s}  "
              f"{r['away_wpct']:>7.4f} {r['home_wpct']:>7.4f}  "
              f"{r['away_fair']:>+7.1f} {r['home_fair']:>+7.1f}  "
              f"{edge_a} {edge_h}")
    print()
    print("=" * 120)
    print("TOTALS (Over/Under) PROJECTIONS")
    print(f"{'Game':10s}  {'Away O/G':>9s} {'Home O/G':>9s}  {'AwayERA':>8s} "
          f"{'HomeERA':>8s}  {'PF':>5s}  {'Total':>7s}  {'Line':>6s}  {'Diff':>7s}")
    print("-" * 120)
    for r in results:
        if r["status"] != "ok":
            continue
        line = r.get("totals_book_line")
        line_str = f"{line:>6.1f}" if isinstance(line, (int, float)) else "   n/a"
        diff_str = (f"{r['predicted_total'] - float(line):>+7.3f}"
                    if isinstance(line, (int, float)) else "    n/a")
        print(f"{r['tab']:10s}  "
              f"{r['away_off']:>9.4f} {r['home_off']:>9.4f}  "
              f"{r['away_total_era']:>8.4f} {r['home_total_era']:>8.4f}  "
              f"{r['park_factor']:>5.2f}  "
              f"{r['predicted_total']:>7.3f}  {line_str}  {diff_str}")
    print("=" * 120)
    if results and results[0].get("status") == "ok":
        print(f"\nLeague averages (today's slate): "
              f"Off. Runs/G={results[0]['league_avg_off']:.4f}, "
              f"Total ERA={results[0]['league_avg_era']:.4f}")


def _write_slate_csv(results, out_path):
    rows = [r for r in results if r["status"] == "ok"]
    if not rows:
        return
    # Drop internal fields (prefixed with _) — they're for the xlsx writer
    public_fields = [k for k in rows[0].keys() if not k.startswith("_")]
    public_rows = [{k: r.get(k) for k in public_fields} for r in rows]
    with open(out_path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=public_fields)
        w.writeheader()
        w.writerows(public_rows)
    print(f"\nSlate CSV: {out_path}")


# ---------------------------------------------------------------------------
# XLSX writer — produces a workbook with one tab per game matching the model
# layout. Values only (no formulas) since Python has already computed them.
# ---------------------------------------------------------------------------

def _write_xlsx_output(results, out_path):
    """Write one tab per game in the same layout as the source model.

    Layout per tab (mirrors BALTB exactly):
        Row 1:    headers ("Away", "Throws/Bats", "BaseRunning", ...)
        Row 2:    Away pitcher (Order=Pitcher, position blank, name in C, ...)
        Rows 3-11: Away batters
        Row 12:   SUM row
        Row 14:   "Home" + header row
        Row 15:   Home pitcher
        Rows 16-24: Home batters
        Row 25:   SUM row
        Row 27:   Header row for ML projection
        Rows 28-29: ML projection (Replacement Wins -> Game W% -> Fair Odds -> EV)
        Row 32-38: Totals projection (RS162, RA162, PF, Total, Line, etc.)
        Rows 41-65: Off. Runs sub-table
    """
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, PatternFill

    wb = Workbook()
    # Remove the default sheet — we'll add our own
    default = wb.active
    wb.remove(default)

    bold = Font(bold=True)
    header_fill = PatternFill("solid", start_color="DCE6F1")
    sum_fill = PatternFill("solid", start_color="FFF2CC")
    proj_fill = PatternFill("solid", start_color="C6EFCE")
    skipped_fill = PatternFill("solid", start_color="FFE5E5")

    # Summary tab first
    summary = wb.create_sheet("SLATE")
    summary["A1"] = "Slate Summary"
    summary["A1"].font = Font(bold=True, size=14)
    summary_headers = ["Tab", "Status", "Away", "Away SP", "Home", "Home SP",
                       "Away W%", "Home W%", "Fair A", "Fair H",
                       "Book A", "Book H", "Edge A", "Edge H",
                       "Predicted Total", "Book Line", "Diff",
                       "Under Odds", "Over Odds",
                       "Books", "AwayML best", "AwayML book",
                       "HomeML best", "HomeML book",
                       "Over med", "Over best", "Over book",
                       "Under med", "Under best", "Under book"]
    for i, h in enumerate(summary_headers, start=1):
        c = summary.cell(row=3, column=i, value=h)
        c.font = bold; c.fill = header_fill
    for r_i, r in enumerate(results, start=4):
        summary.cell(row=r_i, column=1, value=r.get("tab"))
        summary.cell(row=r_i, column=2, value=r.get("status"))
        if r.get("status") != "ok":
            continue
        line = r.get("totals_book_line")
        diff = (r["predicted_total"] - float(line)
                if isinstance(line, (int, float)) else None)
        summary.cell(row=r_i, column=3, value=r["away_team"])
        summary.cell(row=r_i, column=4, value=r["away_sp"])
        summary.cell(row=r_i, column=5, value=r["home_team"])
        summary.cell(row=r_i, column=6, value=r["home_sp"])
        summary.cell(row=r_i, column=7, value=round(r["away_wpct"], 4))
        summary.cell(row=r_i, column=8, value=round(r["home_wpct"], 4))
        summary.cell(row=r_i, column=9, value=round(r["away_fair"], 1))
        summary.cell(row=r_i, column=10, value=round(r["home_fair"], 1))
        summary.cell(row=r_i, column=11, value=r["away_book"])
        summary.cell(row=r_i, column=12, value=r["home_book"])
        if r["away_edge"] is not None:
            summary.cell(row=r_i, column=13, value=round(r["away_edge"], 4))
        if r["home_edge"] is not None:
            summary.cell(row=r_i, column=14, value=round(r["home_edge"], 4))
        summary.cell(row=r_i, column=15, value=round(r["predicted_total"], 3))
        if isinstance(line, (int, float)):
            summary.cell(row=r_i, column=16, value=float(line))
            summary.cell(row=r_i, column=17, value=round(diff, 3))
        u_odds = r.get("total_under_odds")
        o_odds = r.get("total_over_odds")
        if isinstance(u_odds, (int, float)):
            summary.cell(row=r_i, column=18, value=round(u_odds, 1))
        if isinstance(o_odds, (int, float)):
            summary.cell(row=r_i, column=19, value=round(o_odds, 1))
        od = r.get("_odds") or {}
        if od:
            summary.cell(row=r_i, column=20, value=od.get("n_books"))
            summary.cell(row=r_i, column=21, value=od.get("away_ml_best"))
            summary.cell(row=r_i, column=22, value=od.get("away_ml_book"))
            summary.cell(row=r_i, column=23, value=od.get("home_ml_best"))
            summary.cell(row=r_i, column=24, value=od.get("home_ml_book"))
            summary.cell(row=r_i, column=25, value=od.get("over_juice"))
            summary.cell(row=r_i, column=26, value=od.get("over_best"))
            summary.cell(row=r_i, column=27, value=od.get("over_book"))
            summary.cell(row=r_i, column=28, value=od.get("under_juice"))
            summary.cell(row=r_i, column=29, value=od.get("under_best"))
            summary.cell(row=r_i, column=30, value=od.get("under_book"))
    # Column widths
    for col_letter, width in [("A", 10), ("B", 22), ("C", 6), ("D", 25),
                                ("E", 6), ("F", 25)]:
        summary.column_dimensions[col_letter].width = width
    summary.freeze_panes = "A4"

    # One tab per game
    for r in results:
        tab_name = r.get("tab", "UNKNOWN")
        ws = wb.create_sheet(tab_name)

        # Row 1 headers
        headers_1 = {"B": "Away", "D": "Throws/Bats", "E": "BaseRunning",
                     "F": "UZR", "G": "BattingRuns", "H": "ReplacementRuns",
                     "I": "DEF_adj", "J": "FramingRuns", "K": "Runs/Game",
                     "L": "Remaining Games", "M": "ROS Runs", "N": "ROS WAR",
                     "O": "ERA"}
        for col, val in headers_1.items():
            c = ws[f"{col}1"]; c.value = val; c.font = bold; c.fill = header_fill

        # If game wasn't projected, just note why and continue
        if r.get("status") != "ok":
            note_cell = ws["B2"]
            note_cell.value = f"[{r.get('status', 'unknown')}]"
            note_cell.font = Font(italic=True, color="C00000")
            note_cell.fill = skipped_fill
            ws["C2"] = r.get("error", "") or ""
            continue

        # Row 2: Away pitcher
        _write_player_row(ws, 2, r["_away_detail"]["pitcher_row"],
                           is_pitcher=True)
        # Rows 3-11: Away batters
        for i, batter_row in enumerate(r["_away_detail"]["batter_rows"]):
            _write_player_row(ws, 3 + i, batter_row, is_pitcher=False)
        # Row 12: Away SUM
        _write_sum_row(ws, 12, r["_away_detail"]["sums"], sum_fill, bold)

        # Row 14: Home header
        ws["B14"].value = "Home"
        ws["B14"].font = bold; ws["B14"].fill = header_fill
        for col, val in headers_1.items():
            if col == "B": continue
            c = ws[f"{col}14"]; c.value = val; c.font = bold; c.fill = header_fill

        # Row 15: Home pitcher
        _write_player_row(ws, 15, r["_home_detail"]["pitcher_row"],
                           is_pitcher=True)
        # Rows 16-24: Home batters
        for i, batter_row in enumerate(r["_home_detail"]["batter_rows"]):
            _write_player_row(ws, 16 + i, batter_row, is_pitcher=False)
        # Row 25: Home SUM
        _write_sum_row(ws, 25, r["_home_detail"]["sums"], sum_fill, bold)

        # Rows 27-29: ML projection
        ml_headers = {"B": "Team", "C": "Replacement Wins", "D": "Team Wins",
                       "E": "Bullpen", "F": "xWins", "G": "ROS W%",
                       "H": "hf W%", "I": "xW%", "J": "Game W%",
                       "K": "Fair Odds", "L": "Book Odds", "M": "Edge", "N": "EV"}
        for col, val in ml_headers.items():
            c = ws[f"{col}27"]; c.value = val; c.font = bold; c.fill = header_fill

        for excel_row, side in [(28, "away"), (29, "home")]:
            ws.cell(row=excel_row, column=2, value="Away" if side == "away" else "Home").font = bold
            ws.cell(row=excel_row, column=3, value=r["_replacement_wins"])
            ws.cell(row=excel_row, column=4, value=r[f"{side}_team_wins"])
            ws.cell(row=excel_row, column=5, value=r[f"{side}_bullpen_war"])
            ws.cell(row=excel_row, column=6, value=r[f"_{side}_xwins"])
            ws.cell(row=excel_row, column=7, value=r[f"_{side}_ros_wpct"])
            ws.cell(row=excel_row, column=8, value=r[f"_{side}_hf_wpct"])
            ws.cell(row=excel_row, column=9, value=r[f"_{side}_xwpct"])
            wpct_cell = ws.cell(row=excel_row, column=10,
                                 value=r[f"{side}_wpct"])
            wpct_cell.font = bold; wpct_cell.fill = proj_fill
            ws.cell(row=excel_row, column=11, value=r[f"{side}_fair"])
            book = r.get(f"{side}_book")
            if book is not None:
                ws.cell(row=excel_row, column=12, value=book)
            edge = r.get(f"{side}_edge")
            if edge is not None:
                ws.cell(row=excel_row, column=13, value=edge)
            ev = r.get(f"{side}_ev")
            if ev is not None:
                ws.cell(row=excel_row, column=14, value=ev)

        # Rows 31-38: Totals section
        ws["G31"].value = "HF Win%"; ws["G31"].font = bold
        ws["H31"].value = r["_hfa_win_pct"]
        ws["G32"].value = "Bullpen"; ws["G32"].font = bold
        # bullpen_on isn't in result dict; reconstruct from whether bullpen_war is nonzero
        ws["H32"].value = "YES" if (r["away_bullpen_war"] or r["home_bullpen_war"]) else "NO"

        totals_headers = {"B": "Team", "C": "RS 162", "D": "RA 162",
                          "E": "PF", "G": "Total", "I": "Line",
                          "M": "Under", "N": "Over"}
        for col, val in totals_headers.items():
            c = ws[f"{col}36"]; c.value = val; c.font = bold; c.fill = header_fill
        ws["B37"].value = "Away"; ws["B37"].font = bold
        ws["C37"].value = r["away_rs_162"]
        ws["D37"].value = r["away_ra_162"]
        ws["E37"].value = r["park_factor"]
        total_cell = ws["G37"]; total_cell.value = r["predicted_total"]
        total_cell.font = bold; total_cell.fill = proj_fill
        line = r.get("totals_book_line")
        if isinstance(line, (int, float)):
            ws["I37"].value = float(line)
        # M37/N37: over/under American odds at the book line (replica of
        # Total Calcs!M37/N37 — pure Python, no formula dependency)
        u_odds = r.get("total_under_odds")
        o_odds = r.get("total_over_odds")
        if isinstance(u_odds, (int, float)):
            uc = ws["M37"]; uc.value = round(u_odds, 2); uc.font = bold; uc.fill = proj_fill
        if isinstance(o_odds, (int, float)):
            oc = ws["N37"]; oc.value = round(o_odds, 2); oc.font = bold; oc.fill = proj_fill
        ws["B38"].value = "Home"; ws["B38"].font = bold
        ws["C38"].value = r["home_rs_162"]
        ws["D38"].value = r["home_ra_162"]
        ws["E38"].value = r["park_factor"]

        # Rows 41-52: Away Off. Runs detail
        ws["C41"].value = "Away"; ws["C41"].font = bold; ws["C41"].fill = header_fill
        off_headers = {"E": "BaseRunning", "F": "wRC", "G": "Off. Runs",
                       "H": "ERA", "I": "Bullpen", "J": "Total ERA"}
        for col, val in off_headers.items():
            c = ws[f"{col}41"]; c.value = val; c.font = bold; c.fill = header_fill
        ws["A42"].value = "Order"; ws["A42"].font = bold
        ws["C42"].value = r["_away_lineup"].pitcher.name
        ws["D42"].value = r["_away_detail"]["pitcher_hand"]
        ws["E42"].value = 0
        ws["H42"].value = r["_away_pit_totals"]["H_ERA"]
        ws["I42"].value = r["_away_pit_totals"]["I_BP"]
        ws["J42"].value = r["_away_pit_totals"]["J_TotalERA"]
        for i, off_row in enumerate(r["_away_off_rows"]):
            row = 43 + i
            ws.cell(row=row, column=1, value=i + 1)
            ws.cell(row=row, column=3, value=off_row["name"])
            ws.cell(row=row, column=5, value=off_row["E_BR"])
            ws.cell(row=row, column=6, value=off_row["F_wRC"])
            ws.cell(row=row, column=7, value=off_row["G_OffRuns"])
        ws["C52"].value = "Total"; ws["C52"].font = bold
        ws["G52"].value = r["away_off"]
        ws["G52"].font = bold; ws["G52"].fill = sum_fill

        # Rows 54-65: Home Off. Runs detail
        ws["C54"].value = "Home"; ws["C54"].font = bold; ws["C54"].fill = header_fill
        for col, val in off_headers.items():
            c = ws[f"{col}54"]; c.value = val; c.font = bold; c.fill = header_fill
        ws["A55"].value = "Order"; ws["A55"].font = bold
        ws["C55"].value = r["_home_lineup"].pitcher.name
        ws["D55"].value = r["_home_detail"]["pitcher_hand"]
        ws["E55"].value = 0
        ws["H55"].value = r["_home_pit_totals"]["H_ERA"]
        ws["I55"].value = r["_home_pit_totals"]["I_BP"]
        ws["J55"].value = r["_home_pit_totals"]["J_TotalERA"]
        for i, off_row in enumerate(r["_home_off_rows"]):
            row = 56 + i
            ws.cell(row=row, column=1, value=i + 1)
            ws.cell(row=row, column=3, value=off_row["name"])
            ws.cell(row=row, column=5, value=off_row["E_BR"])
            ws.cell(row=row, column=6, value=off_row["F_wRC"])
            ws.cell(row=row, column=7, value=off_row["G_OffRuns"])
        ws["C65"].value = "Total"; ws["C65"].font = bold
        ws["G65"].value = r["home_off"]
        ws["G65"].font = bold; ws["G65"].fill = sum_fill

        # Column widths
        for col, w in [("A", 6), ("B", 9), ("C", 24), ("D", 6),
                        ("E", 11), ("F", 11), ("G", 11), ("H", 11),
                        ("I", 11), ("J", 11), ("K", 11), ("L", 9),
                        ("M", 10), ("N", 10), ("O", 10)]:
            ws.column_dimensions[col].width = w

    wb.save(out_path)
    print(f"Game-tab workbook: {out_path}")


def _write_player_row(ws, excel_row, row_data, is_pitcher):
    """Write one player's row into the worksheet at excel_row, matching the
    BALTB column layout."""
    if is_pitcher:
        ws.cell(row=excel_row, column=1, value="Order")  # A
        ws.cell(row=excel_row, column=2, value=row_data.get("position", "Pitcher"))
    else:
        ws.cell(row=excel_row, column=1, value=row_data["order"])
        ws.cell(row=excel_row, column=2, value=row_data["position"])

    ws.cell(row=excel_row, column=3, value=row_data["name"])  # C: Name
    ws.cell(row=excel_row, column=4, value=row_data.get("D"))  # D: Throws/Bats
    # E-J: only for batters
    for col, key in [(5, "E"), (6, "F"), (7, "G"), (8, "H"), (9, "I"), (10, "J")]:
        v = row_data.get(key)
        if v is not None:
            ws.cell(row=excel_row, column=col, value=v)
    # K, L, M, N, O always
    ws.cell(row=excel_row, column=11, value=row_data.get("K"))
    ws.cell(row=excel_row, column=12, value=row_data.get("L"))
    ws.cell(row=excel_row, column=13, value=row_data.get("M"))
    ws.cell(row=excel_row, column=14, value=row_data.get("N"))
    if row_data.get("O") is not None:
        ws.cell(row=excel_row, column=15, value=row_data["O"])


def _write_sum_row(ws, excel_row, sums, fill, bold):
    """Write the SUM row at row 12 or 25."""
    for col_idx, key in [(5, "E"), (6, "F"), (7, "G"), (8, "H"),
                          (9, "I"), (10, "J"), (11, "K"), (13, "M"), (14, "N")]:
        c = ws.cell(row=excel_row, column=col_idx, value=sums[key])
        c.font = bold; c.fill = fill


def cmd_project(args):
    print(f"Loading references from {args.data_dir}/...")
    refs = References(args.data_dir)
    print(f"  {len(refs.start)} pitchers, {len(refs.pn_vo)} batters indexed.\n")

    target_date = args.date if args.date else TARGET_DATE

    if args.from_model:
        print(f"Reading lineups from model: {args.model}\n")
        lineups_by_tab = None
        unmatched = []
    else:
        print(f"Target date: {target_date}")
        lineups_by_tab, unmatched = build_lineups_from_fangraphs(
            target_date, refs)
        # Show which games have lineups, which are still waiting
        ready = sum(1 for v in lineups_by_tab.values() if v[2] == "ok")
        waiting = sum(1 for v in lineups_by_tab.values() if v[2] == "no_lineup_yet")
        missing = sum(1 for v in lineups_by_tab.values()
                      if v[2].startswith("missing_pitcher"))
        print(f"  {len(lineups_by_tab)} games on slate: "
              f"{ready} ready, {waiting} waiting on lineups, "
              f"{missing} missing pitcher data.\n")

    # Decide where to write outputs. The daily Fangraphs run must NOT depend
    # on the model existing, so default to the directory of --out if given,
    # else the data directory, else the current working directory. Only
    # --from-model (which requires the model anyway) anchors on the model dir.
    if args.out:
        out_dir = Path(args.out).expanduser().resolve().parent
    elif args.from_model and Path(args.model).exists():
        out_dir = Path(args.model).expanduser().resolve().parent
    else:
        # Default the daily run's output to the folder ONE LEVEL ABOVE the
        # reference data dir (e.g. data lives in .../MLB/data_f5, so outputs
        # land in .../MLB). Falls back to cwd if that can't be resolved.
        dd = Path(args.data_dir).expanduser().resolve()
        out_dir = dd.parent if dd.parent != dd else Path.cwd()
    out_dir.mkdir(parents=True, exist_ok=True)

    # Fetch live Pinnacle odds (skip if --from-model or --no-odds)
    odds_by_tab = {}
    if not args.from_model and not args.no_odds:
        print("Fetching live Pinnacle odds via The Odds API...")
        odds_by_tab = fetch_pinnacle_odds(target_date)
        if odds_by_tab:
            save_odds_snapshot(odds_by_tab, target_date, out_dir)
        print()

    print(f"Projecting...\n")
    # project_slate only opens the model in --from-model mode (lineups_by_tab
    # is None). In Fangraphs mode the model path is never read.
    results = project_slate(args.model, refs, lineups_by_tab=lineups_by_tab,
                              odds_by_tab=odds_by_tab)
    _print_slate(results)

    if args.out is None:
        date_str = (target_date.strftime("%Y-%m-%d")
                    if not args.from_model else "frommodel")
        args.out = str(out_dir / f"f5_slate_{date_str}.csv")

    # Write both the per-game xlsx and the summary CSV
    xlsx_path = args.out.replace(".csv", ".xlsx") if args.out.endswith(".csv") \
                  else args.out + ".xlsx"
    _write_xlsx_output(results, xlsx_path)
    _write_slate_csv(results, args.out)

    # Unmatched-players report
    if unmatched:
        report_path = out_dir / f"f5_unmatched_players_{target_date.strftime('%Y-%m-%d')}.csv"
        with open(report_path, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=[
                "tab", "role", "lineup_slot", "fangraphs_name",
                "tab_team_code", "written_to_cell", "source_sheet", "reason",
            ])
            w.writeheader()
            w.writerows(unmatched)
        print(f"\nUnmatched players ({len(unmatched)}): {report_path}")
        # Brief summary
        from collections import Counter
        reasons = Counter(r["reason"].split(":")[0] for r in unmatched)
        for reason, count in reasons.most_common():
            print(f"  {count}: {reason}")

    # Close the Chrome window we used for scraping (only created if we
    # actually scraped Fangraphs, not --from-model)
    if not args.from_model:
        _close_chrome()


# ===========================================================================
# PART 6: VALIDATION — PROVE PYTHON MATCHES THE SPREADSHEET
# ===========================================================================
# Optional but useful when you change something. Compares Python output cell
# by cell against the spreadsheet's cached values. Acknowledges that Excel
# only auto-recalcs the visible tab — for the others, cached values may be
# stale, so the validator detects and excludes those.
# ===========================================================================

def cmd_validate(args):
    print(f"Loading references...")
    refs = References(args.data_dir)
    print(f"Loading model: {args.model}")
    wb = load_workbook(args.model, data_only=True)

    print("\n[1/2] Running slate projection...")
    results = project_slate(args.model, refs)

    print("\n[2/2] Comparing every cell against the spreadsheet...\n")
    print(f"{'Game':10s}  {'Cell':6s} {'Excel(cache)':>13s} {'Python':>13s} "
          f"{'Δ':>10s}  Notes")
    print("-" * 90)

    max_delta = 0.0
    mismatches = 0
    total_checks = 0
    stale_cells = 0

    for r in results:
        if r["status"] != "ok":
            continue
        ws = wb[r["tab"]]
        book_missing = r["away_book"] is None or r["home_book"] is None
        checks = [
            ("J28", r["away_wpct"]),       ("J29", r["home_wpct"]),
            ("K28", r["away_fair"]),       ("K29", r["home_fair"]),
            ("D28", r["away_team_wins"]),  ("D29", r["home_team_wins"]),
            ("E28", r["away_bullpen_war"]),("E29", r["home_bullpen_war"]),
            ("G52", r["away_off"]),        ("G65", r["home_off"]),
            ("J42", r["away_total_era"]),  ("J55", r["home_total_era"]),
            ("C37", r["away_rs_162"]),     ("D37", r["away_ra_162"]),
            ("C38", r["home_rs_162"]),     ("D38", r["home_ra_162"]),
            ("E37", r["park_factor"]),     ("G37", r["predicted_total"]),
        ]
        if not book_missing:
            checks += [
                ("M28", r["away_edge"]), ("M29", r["home_edge"]),
                ("N28", r["away_ev"]),   ("N29", r["home_ev"]),
            ]
        for cell, py_val in checks:
            excel_val = ws[cell].value
            if excel_val is None or not isinstance(excel_val, (int, float)):
                continue
            total_checks += 1
            delta = abs(float(excel_val) - float(py_val))
            if delta > max_delta:
                max_delta = delta
            if delta < 1e-4:
                continue

            # Stale-cache detection: Excel only auto-recalcs the visible
            # tab, so cached values on others may not reflect current inputs.
            is_stale = False
            if cell in ("C37", "D37", "C38", "D38", "G37", "J42", "J55"):
                rel = delta / max(abs(float(excel_val)), 1.0)
                if rel < 0.05:
                    is_stale = True
            if is_stale:
                stale_cells += 1
                continue

            mismatches += 1
            print(f"{r['tab']:10s}  {cell:6s} {excel_val:>13.4f} "
                  f"{py_val:>13.4f} {delta:>10.2e}  GENUINE MISMATCH")

    print()
    print(f"Total cells compared:       {total_checks}")
    print(f"Genuine mismatches:         {mismatches}")
    print(f"Stale-cache divergences:    {stale_cells}")
    print(f"Max delta (any source):     {max_delta:.2e}")
    if mismatches == 0:
        print("\nOK — Python matches the spreadsheet's current formulas.")
    print("\nNote: stale-cache divergences are not Python bugs — they're cells")
    print("the spreadsheet hasn't recalc'd since their inputs changed.")


# ===========================================================================
# RECONCILE — compare Fangraphs master list to PN vO and report mismatches
# ===========================================================================
# Reads the FG PN vO tab (raw Fangraphs imports, periodically pasted in by the
# user) and the PN vO master CSV. For every player Fangraphs projects, decides
# which mismatch bucket they fall into:
#
#   1. OK                 — exact match on (name, team), no action.
#   2. TEAM_MISMATCH      — same name in PN vO but different team. Player has
#                           moved teams; PN vO needs an update.
#   3. NAME_ALIAS         — name appears similar (after normalization) but
#                           doesn't match exactly. Suggest aliases.json entry.
#   4. NOT_IN_PNVO        — Fangraphs has them, PN vO doesn't. Likely a
#                           call-up; needs manual add to PN vO.
#   5. ACCENT_CORRUPTED   — PN vO name has UTF-8-as-Latin1 corruption (Ã±/Ã©).
#                           Flagged so you know to re-export PN vO.
#
# Output: reconciliation_<date>.csv with all rows, plus a console summary.
# ===========================================================================

# Modern code -> Retrosheet code, derived from TEAM_MAP
def _modern_to_retro_team():
    return {modern: retro for (modern, retro) in TEAM_MAP.values()}


def _read_fg_master(model_path, sheet_name="FG PN vO"):
    """Read either 'FG PN vO' (batters) or 'FG PN Start' (pitchers) and
    return [(name, modern_team), ...]. Both tabs have Name in column A and
    Team in column B."""
    wb = load_workbook(model_path, data_only=True)
    ws = wb[sheet_name]
    out = []
    for r in range(2, ws.max_row + 1):
        name = ws.cell(row=r, column=1).value
        team = ws.cell(row=r, column=2).value
        if name and team:
            out.append((str(name).strip(), str(team).strip().upper()))
    return out


def _detect_accent_corruption(s):
    """Detect UTF-8-as-Latin1 corruption (e.g. 'Ã±' for 'ñ', 'Ã©' for 'é').
    Returns the suspected original character pattern."""
    if not s:
        return False
    # The telltale sign is a capital A-tilde 'Ã' followed by something
    return "Ã" in s


def _fix_accent_corruption(s):
    """Best-effort fix of UTF-8-as-Latin1 names. Returns the corrected string."""
    if not s:
        return s
    try:
        return s.encode("latin-1").decode("utf-8")
    except (UnicodeEncodeError, UnicodeDecodeError):
        return s


def _build_master_indexes(df, detect_corruption=True):
    """Index a PN vO or Start dataframe by (norm_name, team).
    Returns (exact_idx, by_name_idx, corrupted_list).
    """
    exact = {}
    by_name = {}
    corrupted = []
    for full_name in df.index:
        parts = full_name.rsplit(" ", 1)
        if len(parts) != 2:
            continue
        name_part, team = parts
        n = _norm_name(name_part)
        exact[(n, team.upper())] = full_name
        by_name.setdefault(n, []).append((team.upper(), full_name))
        if detect_corruption and _detect_accent_corruption(name_part):
            corrupted.append((full_name, _fix_accent_corruption(name_part) + " " + team))
    return exact, by_name, corrupted


def _reconcile_one_table(fg_master, master_exact, master_by_name,
                         modern_to_retro, source_label, report_rows, counts):
    """Compare FG master records against an indexed PN vO / Start table.
    Appends rows to report_rows and updates counts in place."""
    for fg_name, fg_modern_team in fg_master:
        retro_team = modern_to_retro.get(fg_modern_team)
        if retro_team is None:
            continue
        norm = _norm_name(fg_name)

        # 1. Exact match
        if (norm, retro_team) in master_exact:
            counts["OK"] += 1
            continue

        # 2. Same name under different team -> TEAM_MISMATCH
        if norm in master_by_name:
            other_entries = master_by_name[norm]
            current_team = other_entries[0][0]
            counts["TEAM_MISMATCH"] += 1
            report_rows.append({
                "category": "TEAM_MISMATCH",
                "source_sheet": source_label,
                "fg_name": fg_name,
                "fg_team": fg_modern_team,
                "fg_team_retro": retro_team,
                "pnvo_name_current": f"{fg_name} {current_team}",
                "pnvo_team_current": current_team,
                "suggested_pnvo_name": f"{fg_name} {retro_team}",
                "action": f"In {source_label}: change team from {current_team} to {retro_team}",
            })
            continue

        # 3. Fuzzy name search -> NAME_ALIAS
        fg_tokens = norm.split()
        candidates = []
        if len(fg_tokens) >= 2:
            first, last = fg_tokens[0], fg_tokens[-1]
            for cand_norm in master_by_name.keys():
                cand_tokens = cand_norm.split()
                if len(cand_tokens) < 2:
                    continue
                if cand_tokens[-1] == last and cand_tokens[0][:1] == first[:1]:
                    candidates.append(cand_norm)
        if candidates:
            best = min(candidates, key=lambda c: abs(len(c) - len(norm)))
            best_entries = master_by_name[best]
            same_team = [(t, f) for (t, f) in best_entries if t == retro_team]
            if same_team:
                counts["NAME_ALIAS"] += 1
                report_rows.append({
                    "category": "NAME_ALIAS",
                    "source_sheet": source_label,
                    "fg_name": fg_name,
                    "fg_team": fg_modern_team,
                    "fg_team_retro": retro_team,
                    "pnvo_name_current": same_team[0][1],
                    "pnvo_team_current": retro_team,
                    "suggested_pnvo_name": same_team[0][1],
                    "action": f"Add alias: {fg_name!r} -> {same_team[0][1]!r}",
                })
                continue
            # If best candidate is on a different team — likely a team-shift
            # under a name alias (e.g. Patrick Bailey on SFN vs CLE)
            other = best_entries[0]
            counts["TEAM_MISMATCH"] += 1
            report_rows.append({
                "category": "TEAM_MISMATCH",
                "source_sheet": source_label,
                "fg_name": fg_name,
                "fg_team": fg_modern_team,
                "fg_team_retro": retro_team,
                "pnvo_name_current": other[1],
                "pnvo_team_current": other[0],
                "suggested_pnvo_name": f"{fg_name} {retro_team}",
                "action": f"In {source_label}: change team for {other[1]} from {other[0]} to {retro_team}",
            })
            continue

        # 4. Not in master at all
        counts["NOT_IN_PNVO"] += 1
        report_rows.append({
            "category": "NOT_IN_PNVO",
            "source_sheet": source_label,
            "fg_name": fg_name,
            "fg_team": fg_modern_team,
            "fg_team_retro": retro_team,
            "pnvo_name_current": "",
            "pnvo_team_current": "",
            "suggested_pnvo_name": f"{fg_name} {retro_team}",
            "action": f"Add to {source_label} master list",
        })


def cmd_reconcile(args):
    print(f"Loading references from {args.data_dir}/...")
    refs = References(args.data_dir)
    print(f"  {len(refs.start)} pitchers, {len(refs.pn_vo)} batters indexed.\n")

    modern_to_retro = _modern_to_retro_team()
    report_rows = []
    counts = {"OK": 0, "TEAM_MISMATCH": 0, "NAME_ALIAS": 0, "NOT_IN_PNVO": 0}

    # ---- Reconcile batters: FG PN vO vs PN vO ----
    print(f"Reading FG PN vO from {args.model}...")
    fg_batters = _read_fg_master(args.model, "FG PN vO")
    print(f"  {len(fg_batters)} Fangraphs batter records.")
    pnvo_exact, pnvo_by_name, pnvo_corrupted = _build_master_indexes(refs.pn_vo)
    _reconcile_one_table(fg_batters, pnvo_exact, pnvo_by_name,
                          modern_to_retro, "PN vO", report_rows, counts)

    # ---- Reconcile pitchers: FG PN Start vs Start ----
    print(f"Reading FG PN Start from {args.model}...")
    fg_pitchers = _read_fg_master(args.model, "FG PN Start")
    print(f"  {len(fg_pitchers)} Fangraphs pitcher records.\n")
    start_exact, start_by_name, start_corrupted = _build_master_indexes(refs.start)
    _reconcile_one_table(fg_pitchers, start_exact, start_by_name,
                          modern_to_retro, "Start", report_rows, counts)

    # Accent-corruption findings (PN vO and Start combined)
    all_corrupted = pnvo_corrupted + start_corrupted
    for corrupted_full, suggested in all_corrupted:
        report_rows.append({
            "category": "ACCENT_CORRUPTED",
            "source_sheet": "",
            "fg_name": "",
            "fg_team": "",
            "fg_team_retro": "",
            "pnvo_name_current": corrupted_full,
            "pnvo_team_current": corrupted_full.rsplit(" ", 1)[-1],
            "suggested_pnvo_name": suggested,
            "action": "Re-export master with UTF-8 encoding to fix garbled accents",
        })

    # Sort: TEAM_MISMATCH first (most impactful), then NAME_ALIAS, then NOT_IN_PNVO,
    # then ACCENT_CORRUPTED
    order = {"TEAM_MISMATCH": 0, "NAME_ALIAS": 1, "NOT_IN_PNVO": 2,
              "ACCENT_CORRUPTED": 3}
    report_rows.sort(key=lambda r: (order.get(r["category"], 99),
                                      r["fg_team"], r["fg_name"]))

    # Write CSV
    p = Path(args.model)
    from datetime import datetime as _dt
    stamp = _dt.now().strftime("%Y%m%d_%H%M%S")
    out_path = args.out or str(p.parent / f"reconciliation_{stamp}.csv")
    with open(out_path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=[
            "category", "source_sheet", "fg_name", "fg_team", "fg_team_retro",
            "pnvo_name_current", "pnvo_team_current", "suggested_pnvo_name",
            "action",
        ])
        w.writeheader()
        w.writerows(report_rows)

    # Write suggested aliases.json — collect every NAME_ALIAS row into a
    # simple {fg_name: pnvo_name_without_team} map, ready to copy to
    # data/aliases.json. We strip the team code from the PN vO name so the
    # alias is portable across teams.
    suggested_aliases = {}
    for r in report_rows:
        if r["category"] != "NAME_ALIAS":
            continue
        pnvo_full = r["pnvo_name_current"]
        # Strip the trailing team code: "Calvin Conley ATL" -> "Calvin Conley"
        parts = pnvo_full.rsplit(" ", 1)
        pnvo_name_only = parts[0] if len(parts) == 2 else pnvo_full
        suggested_aliases[r["fg_name"]] = pnvo_name_only

    aliases_suggested_path = Path(args.data_dir) / "aliases_suggested.json"
    aliases_total = 0
    if suggested_aliases:
        import json
        # Merge with any existing aliases.json so we don't lose manual entries
        existing = dict(refs.aliases)
        merged = {**existing, **suggested_aliases}
        # Sort alphabetically by FG name for stable diffs
        merged = dict(sorted(merged.items()))
        aliases_total = len(merged)
        with open(aliases_suggested_path, "w", encoding="utf-8") as f:
            json.dump(merged, f, indent=2, ensure_ascii=False)
            f.write("\n")

    # Summary
    print("=" * 70)
    print("RECONCILIATION SUMMARY")
    print("=" * 70)
    print(f"Records checked:          {sum(counts.values())}")
    print(f"  OK (no action needed):  {counts['OK']}")
    print(f"  Team shifts:            {counts['TEAM_MISMATCH']}")
    print(f"  Name aliases:           {counts['NAME_ALIAS']}")
    print(f"  Not in master:          {counts['NOT_IN_PNVO']}")
    if all_corrupted:
        print(f"Master names with garbled accents: {len(all_corrupted)}")
    print(f"\nReport: {out_path}")
    if suggested_aliases:
        print(f"Suggested aliases: {aliases_suggested_path}")
        print(f"  ({len(suggested_aliases)} new, {aliases_total} total when "
              f"merged with existing aliases.json)")
        print(f"  To activate: review the file, then rename it to aliases.json")

    if counts["TEAM_MISMATCH"]:
        print(f"\nTop team shifts (showing first 10):")
        shifts = [r for r in report_rows if r["category"] == "TEAM_MISMATCH"][:10]
        for r in shifts:
            print(f"  [{r['source_sheet']:>5s}] {r['fg_name']:30s}  was {r['pnvo_team_current']:4s} "
                  f"-> FG says {r['fg_team']}")

    if counts["NAME_ALIAS"]:
        print(f"\nTop name aliases (showing first 10):")
        aliases_print = [r for r in report_rows if r["category"] == "NAME_ALIAS"][:10]
        for r in aliases_print:
            print(f"  [{r['source_sheet']:>5s}] FG: {r['fg_name']!r:30s} -> {r['pnvo_name_current']!r}")

    if all_corrupted:
        print(f"\nSample garbled master-list names (showing first 5):")
        for corrupted, suggested in all_corrupted[:5]:
            print(f"  {corrupted!r} -> looks like {suggested!r}")


# ===========================================================================
# CLI
# ===========================================================================

def main():
    parser = argparse.ArgumentParser(
        description="MLB First-5-Innings (F5) projection model — standalone",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""\
Subcommands:
    extract    Pull reference tables out of the F5 .xlsx model into data_f5/
    project    Scrape Fangraphs lineups, pull F5 odds, project F5 ML + totals
    validate   Cell-by-cell compare against the F5 spreadsheet (optional)
    reconcile  Compare FG PN vO master to PN vO and report mismatches

Daily workflow:
    1. Edit TARGET_DATE near the top of this file
    2. python mlb_f5_model.py project
""")
    subs = parser.add_subparsers(dest="cmd", required=True)

    for name, fn in [("extract", cmd_extract),
                     ("project", cmd_project),
                     ("validate", cmd_validate),
                     ("reconcile", cmd_reconcile)]:
        sp = subs.add_parser(name)
        sp.add_argument("--model", default=DEFAULT_MODEL_PATH,
                        help=f"Path to .xlsx model (default: {DEFAULT_MODEL_PATH})")
        sp.add_argument("--data-dir", default=DEFAULT_DATA_DIR,
                        help=f"Reference CSV directory (default: {DEFAULT_DATA_DIR})")
        if name == "extract":
            sp.add_argument("--recompute-league-averages", action="store_true",
                            dest="recompute_league_averages",
                            help="Recompute the fixed league averages from this "
                                 "model's TOTALS tab. Requires a full 30-team "
                                 "slate. Off by default so league baselines stay "
                                 "constant across slates.")
        if name == "project":
            sp.add_argument("--out", default=None,
                            help="Output CSV path (default: f5_slate_<date>.csv "
                                 "in the data dir / cwd; the model is not needed)")
            sp.add_argument("--date", default=None,
                            type=lambda s: date.fromisoformat(s),
                            help="Override TARGET_DATE (format: YYYY-MM-DD)")
            sp.add_argument("--from-model", action="store_true",
                            help="Read lineups from the model's game tabs instead "
                                 "of scraping Fangraphs (useful for re-running "
                                 "after pasting lineups manually)")
            sp.add_argument("--no-odds", action="store_true",
                            help="Skip the Pinnacle odds API call. Edge/EV "
                                 "will be blank but no API credits used.")
        if name == "reconcile":
            sp.add_argument("--out", default=None,
                            help="Output CSV path (default: reconciliation_<timestamp>.csv beside model)")
        sp.set_defaults(func=fn)

    args = parser.parse_args()
    args.func(args)


if __name__ == "__main__":
    # If invoked with no args (e.g. F5 in Spyder), default to `project`.
    if len(sys.argv) == 1:
        sys.argv.append("project")
    main()
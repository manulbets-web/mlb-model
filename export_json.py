"""
export_json.py
==============
Runs the F5 projection engine (mlb_f5_model.project_slate) and writes
docs/data.json for the static dashboard.

This replaces the old R pipeline (mlb_daily.R + export_json.R + Google Sheets).
The F5 model already computes ML + Totals projections off the reference CSVs,
so this script just (1) gets today's lineups + odds, (2) projects, and
(3) serializes the results into a clean JSON the website reads.

USAGE
-----
Daily (live FanGraphs scrape + The Odds API):
    python export_json.py

Offline / test (no scrape) using a sample slate:
    python export_json.py --sample

Re-serialize an already-produced slate xlsx (from `mlb_f5_model.py project`):
    python export_json.py --from-model "/path/F5 slate.xlsx"

After running, commit & push docs/data.json:
    git add docs/data.json && git commit -m "data: $(date +%F)" && git push
"""

from __future__ import annotations

import argparse
import datetime as _dt
import json
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
DEFAULT_DATA_DIR = str(HERE / "data_f5")
OUT_PATH = HERE / "docs" / "data.json"

# Retrosheet-style code normalisation (mirrors the model's ABBR_MAP).
ABBR_MAP = {
    "WSH": "WAS", "WSN": "WAS", "WAS": "WAS", "CWS": "CHA", "CHW": "CHA", "CHA": "CHA",
    "SD": "SDN", "SDP": "SDN", "SDN": "SDN", "SF": "SFN", "SFG": "SFN", "SFN": "SFN",
    "KC": "KCA", "KCR": "KCA", "KCA": "KCA", "TB": "TBA", "TBR": "TBA", "TBA": "TBA",
    "AZ": "ARI", "ARI": "ARI", "NYM": "NYN", "NYN": "NYN", "NYY": "NYA", "NYA": "NYA",
    "ATH": "OAK", "OAK": "OAK", "LAD": "LAN", "LAN": "LAN", "LAA": "LAA", "ANA": "LAA",
    "STL": "SLN", "SLN": "SLN", "CHC": "CHN", "CHN": "CHN",
    "ATL": "ATL", "BAL": "BAL", "BOS": "BOS", "CIN": "CIN", "CLE": "CLE", "COL": "COL",
    "DET": "DET", "HOU": "HOU", "MIA": "MIA", "MIL": "MIL", "MIN": "MIN", "PHI": "PHI",
    "PIT": "PIT", "SEA": "SEA", "TEX": "TEX", "TOR": "TOR",
}


def norm_abbr(a: str) -> str:
    a = (a or "").strip().upper()
    return ABBR_MAP.get(a, a)


def american_to_prob(ml):
    if ml is None:
        return None
    ml = float(ml)
    return (-ml) / (-ml + 100) if ml < 0 else 100 / (ml + 100)


def ml_edge_tier(edge):
    """Same thresholds the old exporter used."""
    if edge is None:
        return "none"
    if edge >= 0.05:
        return "strong"
    if edge >= 0.02:
        return "moderate"
    if edge > 0:
        return "slight"
    return "none"


def _round(x, n=4):
    return round(float(x), n) if isinstance(x, (int, float)) else None


# ---------------------------------------------------------------------------
# Serialisation — turns a model `results` list into the website payload.
# Kept as a pure function so it can be unit-tested with synthetic results.
# ---------------------------------------------------------------------------

def _lineup_to_list(team_lineup):
    if team_lineup is None:
        return []
    out = []
    for b in getattr(team_lineup, "batters", []):
        out.append({
            "order": getattr(b, "order", None),
            "pos": getattr(b, "position", "x"),
            "name": getattr(b, "name", ""),
        })
    return out


def totals_tier(run_gap):
    """Tier a totals lean by how far the model is off the posted line (runs)."""
    if run_gap is None:
        return "none"
    if run_gap >= 0.5:
        return "strong"
    if run_gap >= 0.25:
        return "moderate"
    if run_gap > 0.08:
        return "slight"
    return "none"


def _totals_lean(predicted, line):
    """Model leans toward the side its projected total sits on, vs the line.
    Returns (side, run_gap)."""
    if predicted is None or line is None:
        return None, None
    try:
        diff = float(predicted) - float(line)
    except (TypeError, ValueError):
        return None, None
    if diff > 0:
        return "over", round(diff, 3)
    if diff < 0:
        return "under", round(-diff, 3)
    return None, 0.0


def build_payload(results, slate_date, game_times=None):
    game_times = game_times or {}
    games = []
    bets = []

    league_off = None
    league_era = None

    for r in results:
        if r.get("status") != "ok":
            continue
        away = norm_abbr(r.get("away_team", ""))
        home = norm_abbr(r.get("home_team", ""))
        matchup = f"{away} @ {home}"
        time_str = (game_times.get(f"{away}{home}")
                    or game_times.get(f"{away}@{home}")
                    or game_times.get(frozenset((away, home)), "")) if game_times else ""

        league_off = r.get("league_avg_off", league_off)
        league_era = r.get("league_avg_era", league_era)

        a_tier = ml_edge_tier(r.get("away_edge"))
        h_tier = ml_edge_tier(r.get("home_edge"))
        lean, lean_gap = _totals_lean(r.get("predicted_total"),
                                      r.get("totals_book_line"))
        t_tier = totals_tier(lean_gap)

        game = {
            "tab": r.get("tab"),
            "matchup": matchup,
            "away_abbr": away,
            "home_abbr": home,
            "time": time_str,
            "away_pitcher": {"name": r.get("away_sp"),
                             "throws": r.get("away_throws", ""),
                             "era": _round(r.get("away_pitcher_era"))},
            "home_pitcher": {"name": r.get("home_sp"),
                             "throws": r.get("home_throws", ""),
                             "era": _round(r.get("home_pitcher_era"))},
            "away_lineup": _lineup_to_list(r.get("_away_lineup")),
            "home_lineup": _lineup_to_list(r.get("_home_lineup")),
            "ml": {
                "away_wpct": _round(r.get("away_wpct")),
                "home_wpct": _round(r.get("home_wpct")),
                "away_fair": _round(r.get("away_fair"), 0),
                "home_fair": _round(r.get("home_fair"), 0),
                "away_book": r.get("away_book"),
                "home_book": r.get("home_book"),
                "away_edge": _round(r.get("away_edge")),
                "home_edge": _round(r.get("home_edge")),
                "away_ev": _round(r.get("away_ev")),
                "home_ev": _round(r.get("home_ev")),
                "away_tier": a_tier,
                "home_tier": h_tier,
            },
            "totals": {
                "predicted": _round(r.get("predicted_total"), 3),
                "line": r.get("totals_book_line"),
                "over_fair": r.get("total_over_odds"),
                "under_fair": r.get("total_under_odds"),
                "lean": lean,
                "lean_runs": lean_gap,
                "tier": t_tier,
            },
            "runs": {
                "away_off": _round(r.get("away_off"), 3),
                "home_off": _round(r.get("home_off"), 3),
                "park_factor": _round(r.get("park_factor"), 2),
                "away_total_era": _round(r.get("away_total_era")),
                "home_total_era": _round(r.get("home_total_era")),
            },
        }
        games.append(game)

        # Flat bet list (ML + totals lean) for the "Top edges" board.
        for side, abbr, edge, ev, fair, book, tier in (
            ("away", away, r.get("away_edge"), r.get("away_ev"),
             r.get("away_fair"), r.get("away_book"), a_tier),
            ("home", home, r.get("home_edge"), r.get("home_ev"),
             r.get("home_fair"), r.get("home_book"), h_tier),
        ):
            if edge is not None:
                bets.append({
                    "game": r.get("tab"), "matchup": matchup, "time": time_str,
                    "market": "ML", "side": abbr,
                    "line": book, "fair": _round(fair, 0),
                    "edge": _round(edge), "ev": _round(ev), "tier": tier,
                })

        if game["totals"]["line"] is not None and lean:
            tot = game["totals"]
            bets.append({
                "game": r.get("tab"), "matchup": matchup, "time": time_str,
                "market": "TOTAL",
                "side": f"{lean.upper()} {tot['line']}",
                "line": tot["over_fair"] if lean == "over" else tot["under_fair"],
                "fair": None,
                "edge": lean_gap,          # in runs, not probability
                "ev": None,
                "tier": t_tier,
            })

    tier_rank = {"strong": 3, "moderate": 2, "slight": 1, "none": 0}
    bets.sort(key=lambda b: (tier_rank.get(b["tier"], 0), b["edge"] or 0),
              reverse=True)

    n_strong = sum(1 for b in bets if b["tier"] == "strong")
    n_moderate = sum(1 for b in bets if b["tier"] == "moderate")
    n_slight = sum(1 for b in bets if b["tier"] == "slight")

    return {
        "meta": {
            "date": slate_date,
            "updated_at": _dt.datetime.now(_dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
            "n_games": len(games),
            "n_strong": n_strong,
            "n_moderate": n_moderate,
            "n_slight": n_slight,
            "league_avg_off": _round(league_off),
            "league_avg_era": _round(league_era),
        },
        "bets": bets,
        "games": games,
    }


# ---------------------------------------------------------------------------
# Game-time enrichment — headless MLB StatsAPI call (no Chrome needed).
# ---------------------------------------------------------------------------

def fetch_game_times(slate_date):
    try:
        import requests
    except ImportError:
        return {}
    url = ("https://statsapi.mlb.com/api/v1/schedule?sportId=1"
           f"&date={slate_date}&gameType=R&hydrate=team")
    try:
        r = requests.get(url, timeout=20,
                         headers={"User-Agent": "MLB-F5/1.0"})
        if r.status_code != 200:
            return {}
        data = r.json()
    except Exception:
        return {}
    times = {}
    for d in data.get("dates", []):
        for g in d.get("games", []):
            try:
                away = norm_abbr(g["teams"]["away"]["team"]["abbreviation"])
                home = norm_abbr(g["teams"]["home"]["team"]["abbreviation"])
            except (KeyError, TypeError):
                continue
            gd = g.get("gameDate", "")
            t = ""
            if gd:
                try:
                    dt = _dt.datetime.fromisoformat(gd.replace("Z", "+00:00"))
                    # Convert UTC -> US Eastern (no zoneinfo dependency assumed)
                    dt_et = dt - _dt.timedelta(hours=4)
                    t = dt_et.strftime("%-I:%M %p ET")
                except Exception:
                    t = ""
            times[f"{away}{home}"] = t
    return times


# ---------------------------------------------------------------------------
# Entry points
# ---------------------------------------------------------------------------

def run_live(data_dir, target_date, no_odds=False):
    sys.path.insert(0, str(HERE))
    import mlb_f5_model as m

    refs = m.References(data_dir)
    print(f"Loaded references: {len(refs.start)} pitchers, "
          f"{len(refs.pn_vo)} batters.")

    print(f"Building lineups from FanGraphs for {target_date} ...")
    lineups_by_tab, _ = m.build_lineups_from_fangraphs(target_date, refs)

    odds_by_tab = {}
    if not no_odds:
        print("Fetching live odds ...")
        odds_by_tab = m.fetch_pinnacle_odds(target_date)

    print("Projecting slate ...")
    results = m.project_slate(None, refs, lineups_by_tab=lineups_by_tab,
                              odds_by_tab=odds_by_tab)
    if hasattr(m, "_close_chrome"):
        m._close_chrome()
    return results


def run_from_model(model_path, data_dir):
    sys.path.insert(0, str(HERE))
    import mlb_f5_model as m
    refs = m.References(data_dir)
    print(f"Projecting from model: {model_path}")
    return m.project_slate(model_path, refs, lineups_by_tab=None)


def main():
    p = argparse.ArgumentParser(description="Export F5 projections to docs/data.json")
    p.add_argument("--data-dir", default=DEFAULT_DATA_DIR)
    p.add_argument("--date", default=None, help="YYYY-MM-DD (default: model TARGET_DATE)")
    p.add_argument("--out", default=str(OUT_PATH))
    p.add_argument("--no-odds", action="store_true")
    p.add_argument("--from-model", metavar="XLSX", default=None,
                   help="Project from a model workbook instead of scraping")
    p.add_argument("--sample", action="store_true",
                   help="Write a synthetic sample slate (no scraping) for testing")
    args = p.parse_args()

    if args.sample:
        from sample_slate import sample_results, SAMPLE_DATE, SAMPLE_TIMES
        results = sample_results()
        slate_date = SAMPLE_DATE
        game_times = SAMPLE_TIMES
    elif args.from_model:
        results = run_from_model(args.from_model, args.data_dir)
        slate_date = (args.date or _dt.date.today().isoformat())
        game_times = fetch_game_times(slate_date)
    else:
        if args.date:
            y, mo, d = map(int, args.date.split("-"))
            target_date = _dt.date(y, mo, d)
        else:
            sys.path.insert(0, str(HERE))
            import mlb_f5_model as m
            target_date = m.TARGET_DATE
        slate_date = target_date.isoformat()
        results = run_live(args.data_dir, target_date, no_odds=args.no_odds)
        game_times = fetch_game_times(slate_date)

    payload = build_payload(results, slate_date, game_times)

    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(payload, separators=(",", ":")), encoding="utf-8")
    kb = out.stat().st_size / 1024
    print(f"\nWrote {out}  ({kb:.1f} KB)")
    print(f"  {payload['meta']['n_games']} games | "
          f"strong {payload['meta']['n_strong']}, "
          f"moderate {payload['meta']['n_moderate']}, "
          f"slight {payload['meta']['n_slight']}")


if __name__ == "__main__":
    main()

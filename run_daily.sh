#!/usr/bin/env bash
# run_daily.sh — one step: R drives the Python engine, builds data.json, publishes.
#
#   export_json.R will:
#     1. run the Python engine (scrape FanGraphs + odds -> f5_slate_<DATE>.xlsx)
#        Chrome opens — solve any Cloudflare check in that window
#     2. read that workbook -> docs/data.json
#   then this script commits & pushes (GitHub Pages redeploys).
#
# Usage:
#   ./run_daily.sh                 # uses TARGET_DATE inside mlb_f5_model.py
#   ./run_daily.sh 2026-05-29      # override the slate date
set -euo pipefail
cd "$(dirname "$0")"

echo "== Project + build data.json (R drives the engine) =="
if [ $# -ge 1 ]; then Rscript export_json.R "$1"; else Rscript export_json.R; fi

echo "== Publishing =="
git add docs/data.json
git commit -m "data: $(date +%F)" || { echo "nothing to commit"; exit 0; }
git push
echo "Done -> https://manulbets-web.github.io/mlb-model"

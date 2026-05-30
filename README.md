# MLB First-5 (F5) Model — Site

Static dashboard for the Python F5 projection engine. The **Python engine**
projects the slate; **R** builds the site's `data.json`; GitHub Pages serves it.
You run it locally and push (same publish step as your old R pipeline).

## Files

| File | Role |
|------|------|
| `mlb_f5_model.py` | projection engine (unchanged) — writes `f5_slate_<DATE>.xlsx` |
| `export_json.R` | **runs the Python engine, then** reads its workbook → `docs/data.json` |
| `setup_site.R` | bakes the group password into `docs/index.html` |
| `run_daily.sh` | project → export → commit → push, in one step |
| `docs/index.html` | the dashboard (password-gated, light/minimal) |
| `docs/data.json` | today's projections (regenerated daily) |
| `sample_f5_slate.xlsx` | a fixture so you can test `export_json.R` offline |
| `data_f5/*.csv` | reference data the engine reads |
| `export_json.py`, `sample_slate.py` | optional Python equivalents (not required) |

## One-time setup

1. **R packages**

   ```r
   install.packages(c("openxlsx","jsonlite","httr","stringr"))
   ```

2. **Python deps** (for the engine)

   ```bash
   pip install pandas requests beautifulsoup4 openpyxl undetected-chromedriver
   export ODDS_API_KEY="...."          # for live F5 odds
   ```

3. **Password**

   ```bash
   Rscript setup_site.R                # edit SITE_PASSWORD inside first
   git add docs/index.html && git commit -m "set password" && git push
   ```

   Until you do this the page is open (handy for the first deploy).

4. **GitHub Pages**: Settings → Pages → Deploy from branch `main` `/docs`.
   Live at `https://manulbets-web.github.io/mlb-model`.

## Daily

```bash
./run_daily.sh                 # or  ./run_daily.sh 2026-05-29
```

`export_json.R` runs the engine itself (Chrome opens — solve any Cloudflare
check there), rebuilds `docs/data.json`, and `run_daily.sh` commits & pushes.

Run the export on its own (no git):

```bash
Rscript export_json.R                 # engine + export for today
Rscript export_json.R 2026-05-29      # engine + export for a date
Rscript export_json.R --no-engine     # skip engine, use newest f5_slate_*.xlsx
```

### Test the R exporter offline (no scraping)

Passing an explicit `.xlsx` skips the engine, so you can test against the
fixture without scraping:

```bash
Rscript export_json.R sample_f5_slate.xlsx 2026-05-29
cd docs && python3 -m http.server 8000     # open http://localhost:8000
```

You should see 5 sample games render on the dashboard.

## Notes

- **Why local, not CI?** The lineup scrape uses a real Chrome window that can
  need a manual Cloudflare solve, so it can't run unattended. Publishing is a
  git push, same as before.
- **Totals signal** = model's projected F5 total vs the posted line, in runs.
  The O/U odds shown are the model's *fair* odds at the line (not book prices).
  ML edge = model win% minus the book's implied %.
- `export_json.R` reads the `SLATE` tab for projections and each game tab for
  pitchers (throws + ERA), park factor, and lineups. Offense runs/5 and league
  averages aren't in the workbook, so those card fields show "—".

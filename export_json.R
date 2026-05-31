# =============================================================================
# export_json.R
#
# Reads the Python F5 model's output workbook (the f5_slate_<DATE>.xlsx produced
# by `python mlb_f5_model.py project`) and writes docs/data.json for the static
# dashboard. This is the "website half" of the pipeline; the Python engine does
# all the projecting.
#
# The workbook has:
#   - a "SLATE" summary tab (one row per game, projections + odds)
#   - one tab per game in the model layout (pitchers + lineups + park factor)
#
# DEPENDENCIES:
#   install.packages(c("openxlsx","jsonlite","httr","stringr"))
#
# USAGE:
#   Rscript export_json.R                         # run engine, then export
#   Rscript export_json.R 2026-05-29              # run engine for a date, export
#   Rscript export_json.R --no-engine             # skip engine, use newest xlsx
#   Rscript export_json.R path/to/f5_slate.xlsx   # use this xlsx (engine skipped)
#   Rscript export_json.R path/to/slate.xlsx 2026-05-29   # + override game-time date
# =============================================================================

suppressPackageStartupMessages({
  library(openxlsx); library(jsonlite); library(httr); library(stringr)
})

# ── Config ───────────────────────────────────────────────────────────────────
OUT_PATH       <- "docs/data.json"
# Where the Python model drops f5_slate_*.xlsx (default: repo root / one level
# above data_f5). Adjust if your model writes elsewhere.
MODEL_OUT_DIRS <- c(".", "..", "~/Documents/Sports/MLB")

# Python engine config — used when this script runs the projection itself.
PYTHON_BIN   <- Sys.getenv("PYTHON_BIN", "python3")
MODEL_SCRIPT <- "mlb_f5_model.py"
ENGINE_ARGS  <- c("project")          # extra flags, e.g. c("project","--no-odds")

# ── Args ──────────────────────────────────────────────────────────────────────
# Flags (start with "--") are separated from positionals. Positionals:
#   1 = explicit .xlsx path (optional)   2 = date YYYY-MM-DD (optional)
# A 1st positional that looks like a date is treated as the date, not a path.
raw_args  <- commandArgs(trailingOnly = TRUE)
flags     <- raw_args[startsWith(raw_args, "--")]
pos       <- raw_args[!startsWith(raw_args, "--")]
is_date   <- function(x) grepl("^\\d{4}-\\d{2}-\\d{2}$", x)

xlsx_arg <- NA_character_
date_arg <- NA_character_
if (length(pos) >= 1) {
  if (is_date(pos[[1]])) date_arg <- pos[[1]] else xlsx_arg <- pos[[1]]
}
if (length(pos) >= 2 && is.na(date_arg)) date_arg <- pos[[2]]
# Treat an empty-string xlsx placeholder (used by run_daily.sh) as "none".
if (!is.na(xlsx_arg) && xlsx_arg == "") xlsx_arg <- NA_character_

no_engine <- "--no-engine" %in% flags
# Run the engine by default, UNLESS the user passed an explicit workbook or
# asked to skip it.
run_engine_now <- !no_engine && is.na(xlsx_arg)

# ── Run the Python engine (optional) ──────────────────────────────────────────
run_engine <- function(date_arg) {
  if (!file.exists(MODEL_SCRIPT))
    stop("Engine not found: ", MODEL_SCRIPT,
         " (run from the repo root, or set --no-engine).")
  eng_args <- ENGINE_ARGS
  if (!is.na(date_arg)) eng_args <- c(eng_args, "--date", date_arg)
  message("Running Python engine: ", PYTHON_BIN, " ", MODEL_SCRIPT, " ",
          paste(eng_args, collapse = " "))
  message("  (Chrome will open to scrape FanGraphs — solve any Cloudflare check.)")
  status <- system2(PYTHON_BIN, c(MODEL_SCRIPT, eng_args),
                    stdout = "", stderr = "")   # stream output to console
  if (!identical(status, 0L) && !identical(status, 0))
    stop("Python engine exited with status ", status,
         " — fix the error above, or pass an existing .xlsx to skip it.")
  invisible(TRUE)
}

find_workbook <- function() {
  if (!is.na(xlsx_arg) && file.exists(xlsx_arg)) return(normalizePath(xlsx_arg))
  cands <- character(0)
  for (d in MODEL_OUT_DIRS) {
    d <- path.expand(d)
    if (dir.exists(d))
      cands <- c(cands, list.files(d, pattern = "^f5_slate_.*\\.xlsx$",
                                   full.names = TRUE))
  }
  if (!length(cands)) stop("No f5_slate_*.xlsx found. Pass the path as arg 1.")
  cands[order(file.mtime(cands), decreasing = TRUE)][1]
}

if (run_engine_now) run_engine(date_arg)

WB_PATH <- find_workbook()
message("Reading workbook: ", WB_PATH)

# ── Helpers (mirror export_json.py exactly) ───────────────────────────────────
ABBR_MAP <- c(
  WSH="WAS",WSN="WAS",WAS="WAS",CWS="CHA",CHW="CHA",CHA="CHA",
  SD="SDN",SDP="SDN",SDN="SDN",SF="SFN",SFG="SFN",SFN="SFN",
  KC="KCA",KCR="KCA",KCA="KCA",TB="TBA",TBR="TBA",TBA="TBA",
  AZ="ARI",ARI="ARI",NYM="NYN",NYN="NYN",NYY="NYA",NYA="NYA",
  ATH="OAK",OAK="OAK",LAD="LAN",LAN="LAN",LAA="LAA",ANA="LAA",
  STL="SLN",SLN="SLN",CHC="CHN",CHN="CHN",
  ATL="ATL",BAL="BAL",BOS="BOS",CIN="CIN",CLE="CLE",COL="COL",
  DET="DET",HOU="HOU",MIA="MIA",MIL="MIL",MIN="MIN",PHI="PHI",
  PIT="PIT",SEA="SEA",TEX="TEX",TOR="TOR"
)
norm_abbr <- function(a) {
  a <- toupper(trimws(as.character(a %||% "")))
  m <- ABBR_MAP[a]; if (is.na(m)) a else unname(m)
}
`%||%` <- function(a, b) if (is.null(a) || length(a) == 0 || (length(a)==1 && is.na(a))) b else a
as_num <- function(x) suppressWarnings(as.numeric(gsub("%", "", as.character(x))))
rnd    <- function(x, n=4) if (is.null(x)||length(x)==0||is.na(x)) NULL else round(as.numeric(x), n)

ml_tier <- function(edge) {
  if (is.null(edge) || length(edge)==0 || is.na(edge)) return("none")
  if (edge >= 0.05) return("strong")
  if (edge >= 0.02) return("moderate")
  if (edge > 0)     return("slight")
  "none"
}
totals_tier <- function(g) {
  if (is.null(g) || is.na(g)) return("none")
  if (g >= 0.5)  return("strong")
  if (g >= 0.25) return("moderate")
  if (g > 0.08)  return("slight")
  "none"
}
strip_team <- function(name, abbr) {
  if (is.na(name)) return(NA_character_)
  str_trim(gsub(paste0("\\s+", abbr, "\\s*$"), "", name))
}

# Read a sheet as a raw character/numeric matrix preserving row/col positions.
read_grid <- function(path, sheet) {
  openxlsx::read.xlsx(path, sheet = sheet, colNames = FALSE,
                      skipEmptyRows = FALSE, skipEmptyCols = FALSE,
                      sep.names = " ")
}
cell <- function(grid, r, c) {
  if (is.null(grid) || r < 1 || r > nrow(grid) || c < 1 || c > ncol(grid)) return(NA)
  grid[[r, c]]
}

# ── Read SLATE summary tab ────────────────────────────────────────────────────
sheets <- openxlsx::getSheetNames(WB_PATH)
if (!"SLATE" %in% sheets) stop("Workbook has no SLATE tab — wrong file?")
slate <- read_grid(WB_PATH, "SLATE")
# Data rows start at row 4 (row 1 title, row 3 header). Columns:
# 1 Tab 2 Status 3 Away 4 AwaySP 5 Home 6 HomeSP 7 AwayW% 8 HomeW%
# 9 FairA 10 FairH 11 BookA 12 BookH 13 EdgeA 14 EdgeH
# 15 PredTotal 16 BookLine 17 Diff 18 UnderOdds 19 OverOdds
slate_rows <- list()
for (r in 4:nrow(slate)) {
  tab <- cell(slate, r, 1)
  if (is.na(tab) || tab == "") next
  slate_rows[[as.character(tab)]] <- list(
    tab=tab, status=cell(slate,r,2),
    away=cell(slate,r,3), away_sp=cell(slate,r,4),
    home=cell(slate,r,5), home_sp=cell(slate,r,6),
    away_wpct=as_num(cell(slate,r,7)), home_wpct=as_num(cell(slate,r,8)),
    away_fair=as_num(cell(slate,r,9)), home_fair=as_num(cell(slate,r,10)),
    away_book=as_num(cell(slate,r,11)), home_book=as_num(cell(slate,r,12)),
    away_edge=as_num(cell(slate,r,13)), home_edge=as_num(cell(slate,r,14)),
    predicted=as_num(cell(slate,r,15)), line=as_num(cell(slate,r,16)),
    under_fair=as_num(cell(slate,r,18)), over_fair=as_num(cell(slate,r,19))
  )
}

# ── Game times (MLB StatsAPI, headless) ───────────────────────────────────────
slate_date <- if (!is.na(date_arg)) date_arg else format(Sys.Date(), "%Y-%m-%d")
fetch_times <- function(d) {
  url <- paste0("https://statsapi.mlb.com/api/v1/schedule?sportId=1&date=", d,
                "&gameType=R&hydrate=team")
  out <- list()
  r <- tryCatch(httr::GET(url, httr::timeout(20),
                          httr::user_agent("MLB-F5/1.0")), error=function(e) NULL)
  if (is.null(r) || httr::status_code(r) != 200) return(out)
  j <- tryCatch(jsonlite::fromJSON(httr::content(r,"text",encoding="UTF-8"),
                                   flatten=TRUE), error=function(e) NULL)
  if (is.null(j) || length(j$dates) == 0) return(out)
  gr <- j$dates$games[[1]]
  for (i in seq_len(nrow(gr))) {
    away <- norm_abbr(gr[["teams.away.team.abbreviation"]][i])
    home <- norm_abbr(gr[["teams.home.team.abbreviation"]][i])
    t <- tryCatch(format(as.POSIXct(gr$gameDate[i], tz="UTC"),
                         "%I:%M %p ET", tz="America/New_York"),
                  error=function(e) "")
    out[[paste0(away, home)]] <- sub("^0", "", t)
  }
  out
}
times <- tryCatch(fetch_times(slate_date), error=function(e) list())

# ── Build games + bets ────────────────────────────────────────────────────────
games <- list(); bets <- list()
lg_off <- NULL; lg_era <- NULL  # not in workbook; left null (shown as —)

for (key in names(slate_rows)) {
  s <- slate_rows[[key]]
  if (is.na(s$status) || s$status != "ok") next
  away <- norm_abbr(s$away); home <- norm_abbr(s$home)
  matchup <- paste0(away, " @ ", home)
  tkey <- paste0(away, home)
  gtime <- times[[tkey]] %||% ""

  # Per-game tab: pitchers (C2/C15, throws D2/D15, ERA O2/O15), PF (E37), lineups
  g <- if (s$tab %in% sheets) read_grid(WB_PATH, s$tab) else NULL
  read_lineup <- function(grid, start) {
    if (is.null(grid)) return(list())
    out <- list()
    for (i in 0:8) {
      r <- start + i
      nm <- cell(grid, r, 3); if (is.na(nm) || nm=="") next
      pos <- tolower(as.character(cell(grid, r, 2) %||% "x"))
      if (!pos %in% c("c","dh","x")) pos <- "x"
      out[[length(out)+1]] <- list(order=i+1L, pos=pos,
                                   name=strip_team(nm, "")) # FG names have no suffix
    }
    out
  }
  away_throws <- as.character(cell(g, 2, 4) %||% "")
  home_throws <- as.character(cell(g, 15, 4) %||% "")
  away_era <- as_num(cell(g, 2, 15)); home_era <- as_num(cell(g, 15, 15))
  pf       <- as_num(cell(g, 37, 5))

  a_tier <- ml_tier(s$away_edge); h_tier <- ml_tier(s$home_edge)

  lean <- NULL; lean_runs <- NULL; t_tier <- "none"
  if (!is.null(s$predicted) && !is.na(s$predicted) &&
      !is.null(s$line) && !is.na(s$line)) {
    diff <- s$predicted - s$line
    if (diff > 0) { lean <- "over";  lean_runs <- round(diff, 3) }
    else if (diff < 0) { lean <- "under"; lean_runs <- round(-diff, 3) }
    else { lean_runs <- 0 }
    t_tier <- totals_tier(lean_runs)
  }

  games[[length(games)+1]] <- list(
    tab=s$tab, matchup=matchup, away_abbr=away, home_abbr=home, time=gtime,
    away_pitcher=list(name=strip_team(s$away_sp, away), throws=away_throws, era=rnd(away_era)),
    home_pitcher=list(name=strip_team(s$home_sp, home), throws=home_throws, era=rnd(home_era)),
    away_lineup=read_lineup(g, 3), home_lineup=read_lineup(g, 16),
    ml=list(
      away_wpct=rnd(s$away_wpct), home_wpct=rnd(s$home_wpct),
      away_fair=rnd(s$away_fair,0), home_fair=rnd(s$home_fair,0),
      away_book=s$away_book, home_book=s$home_book,
      away_edge=rnd(s$away_edge), home_edge=rnd(s$home_edge),
      away_ev=NULL, home_ev=NULL,
      away_tier=a_tier, home_tier=h_tier
    ),
    totals=list(
      predicted=rnd(s$predicted,3), line=s$line,
      over_fair=s$over_fair, under_fair=s$under_fair,
      lean=lean, lean_runs=lean_runs, tier=t_tier
    ),
    runs=list(away_off=NULL, home_off=NULL, park_factor=rnd(pf,2),
              away_total_era=NULL, home_total_era=NULL)
  )

  # bets: ML both sides + the totals lean
  for (sd in list(
    list(abbr=away, edge=s$away_edge, line=s$away_book, fair=s$away_fair, tier=a_tier),
    list(abbr=home, edge=s$home_edge, line=s$home_book, fair=s$home_fair, tier=h_tier))) {
    if (!is.null(sd$edge) && !is.na(sd$edge)) {
      bets[[length(bets)+1]] <- list(
        game=s$tab, matchup=matchup, time=gtime, market="ML",
        side=sd$abbr, line=sd$line, fair=rnd(sd$fair,0),
        edge=rnd(sd$edge), ev=NULL, tier=sd$tier)
    }
  }
  if (!is.null(s$line) && !is.na(s$line) && !is.null(lean)) {
    bets[[length(bets)+1]] <- list(
      game=s$tab, matchup=matchup, time=gtime, market="TOTAL",
      side=paste0(toupper(lean), " ", s$line),
      line=if (lean=="over") s$over_fair else s$under_fair,
      fair=NULL, edge=lean_runs, ev=NULL, tier=t_tier)
  }
}

# Sort bets by tier strength then magnitude
tier_rank <- c(strong=3, moderate=2, slight=1, none=0)
if (length(bets)) {
  ord <- order(
    sapply(bets, function(b) unname(tier_rank[b$tier])),
    sapply(bets, function(b) b$edge %||% 0),
    decreasing = TRUE)
  bets <- bets[ord]
}
count_tier <- function(tt) sum(vapply(bets, function(b) b$tier == tt, logical(1)))

payload <- list(
  meta = list(
    date = slate_date,
    updated_at = format(as.POSIXct(Sys.time(), tz="UTC"), "%Y-%m-%dT%H:%M:%SZ", tz="UTC"),
    n_games = length(games),
    n_strong = if (length(bets)) count_tier("strong") else 0,
    n_moderate = if (length(bets)) count_tier("moderate") else 0,
    n_slight = if (length(bets)) count_tier("slight") else 0,
    league_avg_off = lg_off, league_avg_era = lg_era
  ),
  bets = bets,
  games = games
)

dir.create(dirname(OUT_PATH), showWarnings = FALSE, recursive = TRUE)
writeLines(jsonlite::toJSON(payload, auto_unbox = TRUE, null = "null",
                            na = "null", digits = NA), OUT_PATH)
message(sprintf("Wrote %s  (%.1f KB) | %d games | strong %d, moderate %d, slight %d",
                OUT_PATH, file.size(OUT_PATH)/1024, length(games),
                payload$meta$n_strong, payload$meta$n_moderate, payload$meta$n_slight))

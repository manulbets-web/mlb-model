# =============================================================================
# mlb_daily.R
#
# Run each morning to produce a dated local Excel model pre-filled with
# today's pitchers and batting lineups.
#
# WHAT IT DOES
# ------------
#   1. Authenticates to Google Sheets (one-time browser prompt, then cached)
#   2. Reads all reference tabs from your model Google Sheet
#   3. Fetches today's schedule + probable starters from the MLB Stats API
#   4. Fetches lineups from FanGraphs (projected) with MLB API boxscore fallback
#   5. Builds a local .xlsx with one new game tab per matchup (e.g. "NYYBAL")
#
# SHEET LAYOUT (per game tab) — TWO identical lineup blocks
# ----------------------------------------------------------
#   TOP block  (WAR / projection section)
#     pitcher header row  → col C = "Name TEAM", col D = hand
#     rows 1-9 away       → col A = order, B = pos, C = "Name TEAM", D = bats
#     pitcher header row  → home pitcher
#     rows 1-9 home
#
#   BOTTOM block  (Off. Runs / wRC section) — same players, same format
#     pitcher header row  → away pitcher again
#     rows 1-9 away
#     pitcher header row  → home pitcher again
#     rows 1-9 home
#
# The script writes ALL FOUR lineup blocks so both sections show the
# correct game's players.
#
# WHAT STILL NEEDS MANUAL ENTRY
# ------------------------------
#   • Book odds  (moneyline + total) in each game tab
#   • HF Win%    toggle
#   • TBD starters once announced
#   • Custom Starter IP overrides if desired
#
# DEPENDENCIES
# ------------
#   install.packages(c("googlesheets4", "openxlsx", "httr", "jsonlite",
#                      "dplyr", "purrr", "tibble", "stringr"))
#
# USAGE
#   source("mlb_daily.R")
#   Rscript mlb_daily.R
# =============================================================================

suppressPackageStartupMessages({
  library(googlesheets4)
  library(openxlsx)
  library(httr)
  library(jsonlite)
  library(dplyr)
  library(purrr)
  library(tibble)
  library(stringr)
})

# =============================================================================
# CONFIGURATION
# =============================================================================

SHEET_ID   <- "1F3hHYptA-lvD3o8a5P3y46yN_uLn5BlqwSjDA4yeUC0"
OUTPUT_DIR <- "~/Downloads/data/output"
SLATE_DATE <- format(Sys.Date(), "%Y-%m-%d")   # override: "2026-05-14"

GAME_TAB_PATTERN <- "^[A-Z]{4,8}$"
PRESERVE_TABS    <- c("TOTALS")

REF_TABS <- c(
  "Parameters", "Total Calcs", "MLB", "Start", "DEF Adj",
  "FG PN Start", "PN vO", "FG PN vO", "PN vR", "FG PN vR",
  "PN vL", "FG PN vL", "MLB Bullpen", "Proj BP", "TOTALS"
)

ABBR_MAP <- c(
  WSN="WSH", WAS="WSH", WSH="WSH",
  CHW="CWS", CHA="CWS", CWS="CWS",
  SDP="SD",  SD="SD",   SFG="SF",  SF="SF",
  KCR="KC",  KC="KC",   TBR="TB",  TBA="TB",  TB="TB",
  ARI="AZ",  AZ="AZ",   NYN="NYM", NYM="NYM", NYY="NYY",
  OAK="ATH", ATH="ATH", LAN="LAD", LAD="LAD", LAA="LAA",
  SLN="STL", STL="STL", CHN="CHC", CHC="CHC",
  PIT="PIT", CIN="CIN", ATL="ATL", MIA="MIA", PHI="PHI",
  COL="COL", MIL="MIL", MIN="MIN", CLE="CLE",
  DET="DET", HOU="HOU", SEA="SEA", TEX="TEX",
  TOR="TOR", BOS="BOS", BAL="BAL"
)

norm_abbr <- function(a) {
  a <- toupper(trimws(a))
  dplyr::coalesce(ABBR_MAP[a], a)
}

`%||%` <- function(a, b) {
  if (!is.null(a) && length(a) > 0 && !all(is.na(a)) && !all(a == "")) a else b
}

MLB_BASE <- "https://statsapi.mlb.com/api/v1"

# =============================================================================
# SECTION 1: Authenticate + read Google Sheet metadata
# =============================================================================

message("\n", strrep("=", 60))
message("  MLB DAILY — ", SLATE_DATE)
message(strrep("=", 60))

message("\n── Authenticating to Google Sheets...")
googlesheets4::gs4_auth()

message("── Reading sheet metadata...")
ss       <- googlesheets4::gs4_get(SHEET_ID)
all_tabs <- ss$sheets$name
message("  Tabs found : ", length(all_tabs))

game_tabs <- all_tabs[
  grepl(GAME_TAB_PATTERN, all_tabs) & !all_tabs %in% PRESERVE_TABS
]
message("  Game tabs  : ", paste(game_tabs, collapse = ", "))

# =============================================================================
# SECTION 2: Read template tab + detect all four lineup block row positions
# =============================================================================
#
# Each game tab has batting order "1.0" in col A exactly four times:
#   hit 1 → away lineup start, TOP block
#   hit 2 → home lineup start, TOP block
#   hit 3 → away lineup start, BOTTOM block  (wRC / Off.Runs section)
#   hit 4 → home lineup start, BOTTOM block
#
# The pitcher header row is always (lineup_start_row - 1).

message("\n── Reading template tab: ", game_tabs[1], " ...")
template_tab  <- game_tabs[1]
template_data <- googlesheets4::read_sheet(
  SHEET_ID,
  sheet     = template_tab,
  col_names = FALSE,
  col_types = "c"
)

find_all_lineup_rows <- function(df) {
  col_a <- as.character(df[[1]])
  hits  <- which(col_a %in% c("1", "1.0"))
  
  if (length(hits) >= 4) {
    return(list(
      away_top = hits[1], home_top = hits[2],
      away_bot = hits[3], home_bot = hits[4]
    ))
  }
  
  # Fewer than 4 hits — build fallback from what we have
  message("  [WARN] Found only ", length(hits),
          " lineup-start rows (expected 4); using positional fallback.")
  if (length(hits) >= 2) {
    gap <- hits[2] - hits[1]          # top-away → top-home spacing (~11)
    # Bottom block typically starts ~25-30 rows below home_top
    bot_offset <- gap * 3L
    return(list(
      away_top = hits[1],
      home_top = hits[2],
      away_bot = hits[1] + bot_offset,
      home_bot = hits[2] + bot_offset
    ))
  }
  # Hard fallback matching the known model layout
  list(away_top = 3L, home_top = 14L, away_bot = 39L, home_bot = 51L)
}

rows       <- find_all_lineup_rows(template_data)
away_top   <- rows$away_top;   home_top <- rows$home_top
away_bot   <- rows$away_bot;   home_bot <- rows$home_bot
p_away_top <- away_top - 1L;   p_home_top <- home_top - 1L
p_away_bot <- away_bot - 1L;   p_home_bot <- home_bot - 1L

message("  TOP block — away row: ", away_top, "  home row: ", home_top)
message("  BOT block — away row: ", away_bot, "  home row: ", home_bot)

# =============================================================================
# SECTION 3: Read reference tabs from Google Sheet
# =============================================================================

message("\n── Reading reference tabs...")
ref_data <- list()
for (tab in REF_TABS) {
  if (!tab %in% all_tabs) { message("  [SKIP] '", tab, "'"); next }
  ref_data[[tab]] <- tryCatch(
    googlesheets4::read_sheet(SHEET_ID, sheet = tab, col_types = "c"),
    error = function(e) {
      message("  [WARN] '", tab, "': ", conditionMessage(e)); NULL
    }
  )
  n <- if (!is.null(ref_data[[tab]])) nrow(ref_data[[tab]]) else 0
  message("  ✓  ", tab, "  (", n, " rows)")
}

# =============================================================================
# SECTION 4: Fetch schedule from MLB Stats API
# =============================================================================

message("\n── Fetching schedule: ", SLATE_DATE, " ...")

mlb_get <- function(url, max_tries = 3L, pause = 2) {
  for (i in seq_len(max_tries)) {
    resp <- tryCatch(
      httr::GET(url, httr::timeout(20),
                httr::user_agent("MLB-Model-R/1.0")),
      error = function(e) { message("  [error] ", e$message); NULL }
    )
    if (!is.null(resp) && httr::status_code(resp) == 200L)
      return(jsonlite::fromJSON(
        httr::content(resp, as = "text", encoding = "UTF-8"),
        flatten = TRUE
      ))
    if (i < max_tries) Sys.sleep(pause)
  }
  NULL
}

raw_sched <- mlb_get(paste0(
  MLB_BASE, "/schedule?sportId=1&date=", SLATE_DATE,
  "&gameType=R&hydrate=probablePitcher(note),team,venue,status"
))
if (is.null(raw_sched) || length(raw_sched$dates) == 0)
  stop("No games found for ", SLATE_DATE)

games_raw <- raw_sched$dates$games[[1]]
message("  Games on slate: ", nrow(games_raw))

games <- tibble::tibble(
  game_pk    = as.integer(games_raw$gamePk),
  away_abbr  = norm_abbr(games_raw[["teams.away.team.abbreviation"]]),
  home_abbr  = norm_abbr(games_raw[["teams.home.team.abbreviation"]]),
  venue      = games_raw[["venue.name"]],
  status     = games_raw[["status.detailedState"]],
  away_pid   = as.integer(games_raw[["teams.away.probablePitcher.id"]]),
  away_pname = games_raw[["teams.away.probablePitcher.fullName"]],
  home_pid   = as.integer(games_raw[["teams.home.probablePitcher.id"]]),
  home_pname = games_raw[["teams.home.probablePitcher.fullName"]]
) |>
  dplyr::mutate(
    away_pname = dplyr::if_else(is.na(away_pname), "TBD", away_pname),
    home_pname = dplyr::if_else(is.na(home_pname), "TBD", home_pname)
  )

# Pitcher handedness
pitcher_ids <- unique(na.omit(c(games$away_pid, games$home_pid)))
throws_map  <- list()
if (length(pitcher_ids) > 0) {
  for (batch in split(pitcher_ids, ceiling(seq_along(pitcher_ids) / 50))) {
    d <- mlb_get(paste0(MLB_BASE, "/people?personIds=",
                        paste(batch, collapse = ","), "&hydrate=none"))
    if (!is.null(d$people))
      for (j in seq_len(nrow(d$people)))
        throws_map[[as.character(d$people$id[j])]] <-
          d$people$pitchHand.code[j] %||% "R"
  }
}
games <- games |>
  dplyr::mutate(
    away_throws = purrr::map_chr(as.character(away_pid),
                                 ~ throws_map[[.x]] %||% "R"),
    home_throws = purrr::map_chr(as.character(home_pid),
                                 ~ throws_map[[.x]] %||% "R")
  )

# =============================================================================
# SECTION 5: Fetch lineups — FanGraphs primary, MLB API fallback
# =============================================================================

message("\n── Fetching lineups from FanGraphs...")

fetch_fg_lineups <- function(slate_date) {
  resp <- tryCatch(
    httr::GET(
      paste0("https://www.fangraphs.com/scores?date=", slate_date),
      httr::timeout(25),
      httr::user_agent(paste0("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) ",
                              "AppleWebKit/537.36 Chrome/124.0 Safari/537.36")),
      httr::add_headers(
        Accept            = "text/html,application/xhtml+xml",
        `Accept-Language` = "en-US,en;q=0.9",
        Referer           = "https://www.fangraphs.com/"
      )
    ),
    error = function(e) { message("  FG error: ", e$message); NULL }
  )
  if (is.null(resp) || httr::status_code(resp) != 200L) {
    message("  FanGraphs HTTP ",
            httr::status_code(resp %||% list(status_code = 0)))
    return(list())
  }
  page <- httr::content(resp, as = "text", encoding = "UTF-8")
  m    <- regmatches(page,
                     regexpr(
                       '<script id="__NEXT_DATA__" type="application/json">(.*?)</script>',
                       page, perl = TRUE
                     ))
  if (length(m) == 0L) { message("  FG: __NEXT_DATA__ not found"); return(list()) }
  
  parsed <- tryCatch(
    jsonlite::fromJSON(
      gsub('<script id="__NEXT_DATA__" type="application/json">|</script>', "", m),
      simplifyVector = FALSE
    ),
    error = function(e) { message("  FG JSON error: ", e$message); NULL }
  )
  if (is.null(parsed)) return(list())
  
  games_fg <- tryCatch(
    parsed$props$pageProps$dehydratedState$queries[[1L]]$state$data,
    error = function(e) NULL
  )
  if (is.null(games_fg) || length(games_fg) == 0L) {
    message("  FG: no games in payload"); return(list())
  }
  
  parse_side <- function(player_list) {
    if (is.null(player_list) || length(player_list) == 0L) return(list())
    player_list <- player_list[order(
      sapply(player_list, function(p) as.integer(p$BatOrder %||% 99L))
    )]
    lapply(player_list, function(p) list(
      mlbam_id = as.integer(p$MLBAMID     %||% 0L),
      name     = as.character(p$PlayerName %||% "Unknown"),
      bats     = as.character(p$Bats       %||% "R"),
      position = as.character(p$DisplayPosition %||% p$Position %||% ""),
      projected= isTRUE(p$IsProjected)
    ))
  }
  
  result <- list()
  for (g in games_fg) {
    s  <- g$schedule;  lu <- g$lineups
    if (is.null(s)) next
    pk   <- as.character(s$MLBGameId %||% "")
    away <- norm_abbr(s$AwayTeamAbbName %||% "")
    if (!nzchar(pk) || !nzchar(away)) next
    result[[pk]] <- list(
      away_lineup = if (!is.null(lu)) parse_side(lu$lineupAway) else list(),
      home_lineup = if (!is.null(lu)) parse_side(lu$lineupHome) else list()
    )
  }
  message("  FanGraphs: parsed ", length(result), " games")
  result
}

fetch_lineup_mlb <- function(game_pk) {
  d <- mlb_get(paste0(MLB_BASE, "/game/", game_pk, "/boxscore"))
  if (is.null(d)) return(list(away = list(), home = list()))
  parse_side <- function(td) {
    batters <- td$batters;  players <- td$players
    if (is.null(batters) || length(batters) == 0) return(list())
    purrr::map(seq_along(batters), function(i) {
      p <- players[[paste0("ID", batters[[i]])]]
      list(
        mlbam_id = as.integer(batters[[i]]),
        name     = tryCatch(p$person$fullName,       error = function(e) "Unknown"),
        bats     = tryCatch(p$person$batSide$code,   error = function(e) "R"),
        position = tryCatch(p$position$abbreviation, error = function(e) "")
      )
    })
  }
  list(away = parse_side(d$teams$away), home = parse_side(d$teams$home))
}

fg_games     <- fetch_fg_lineups(SLATE_DATE)
lineups_list <- purrr::map(games$game_pk, function(pk) {
  fg <- fg_games[[as.character(pk)]]
  if (!is.null(fg) &&
      (length(fg$away_lineup) > 0 || length(fg$home_lineup) > 0)) {
    proj <- any(sapply(c(fg$away_lineup, fg$home_lineup),
                       function(p) isTRUE(p$projected)))
    return(list(away = fg$away_lineup, home = fg$home_lineup,
                src  = if (proj) "FG-projected" else "FG-confirmed"))
  }
  r   <- fetch_lineup_mlb(pk)
  src <- if (length(r$away) > 0 || length(r$home) > 0) "MLB-API" else "none"
  list(away = r$away, home = r$home, src = src)
})
names(lineups_list) <- as.character(games$game_pk)

for (i in seq_len(nrow(games))) {
  pk <- as.character(games$game_pk[i])
  lu <- lineups_list[[pk]]
  message("  ", games$away_abbr[i], "@", games$home_abbr[i],
          "  away=", length(lu$away), " home=", length(lu$home),
          " [", lu$src, "]")
}

# =============================================================================
# SECTION 6: Build local workbook
# =============================================================================

message("\n── Building local workbook...")
wb <- openxlsx::createWorkbook()

# Write reference tabs
for (tab in names(ref_data)) {
  df <- ref_data[[tab]]
  if (is.null(df)) next
  openxlsx::addWorksheet(wb, tab)
  openxlsx::writeData(wb, sheet = tab, x = df, colNames = TRUE)
}

# Write template layout tab
openxlsx::addWorksheet(wb, "TEMPLATE_GAME")
openxlsx::writeData(wb, sheet = "TEMPLATE_GAME",
                    x = template_data, colNames = FALSE)

# Helper: write a scalar to one cell
wc <- function(wb, sheet, row, col, val)
  openxlsx::writeData(wb, sheet = sheet, x = val,
                      startRow = row, startCol = col, colNames = FALSE)

# Helper: player list → 4-column data frame (order, position, name+team, hand)
lineup_df <- function(player_list, abbr) {
  purrr::map_dfr(seq_len(min(length(player_list), 9L)), function(i) {
    p   <- player_list[[i]]
    pos <- tolower(p$position %||% "x")
    tibble::tibble(
      order    = as.numeric(i),
      position = if (pos == "c") "c" else "x",
      name     = paste(p$name %||% "Unknown", abbr),
      hand     = p$bats %||% "R"
    )
  })
}

# Write one game sheet per matchup
for (i in seq_len(nrow(games))) {
  g         <- games[i, ]
  pk        <- as.character(g$game_pk)
  away_abbr <- g$away_abbr
  home_abbr <- g$home_abbr
  sname     <- substr(paste0(away_abbr, home_abbr), 1L, 31L)
  lu        <- lineups_list[[pk]]
  
  openxlsx::addWorksheet(wb, sname)
  openxlsx::writeData(wb, sheet = sname,
                      x = template_data, colNames = FALSE)
  
  away_label <- paste(g$away_pname, away_abbr)
  home_label <- paste(g$home_pname, home_abbr)
  
  # ── TOP block: pitcher headers ─────────────────────────────────────────
  wc(wb, sname, p_away_top, 3L, away_label);  wc(wb, sname, p_away_top, 4L, g$away_throws)
  wc(wb, sname, p_home_top, 3L, home_label);  wc(wb, sname, p_home_top, 4L, g$home_throws)
  
  # ── TOP block: batting lineups ──────────────────────────────────────────
  if (length(lu$away) > 0)
    openxlsx::writeData(wb, sheet = sname, x = lineup_df(lu$away, away_abbr),
                        startRow = away_top, startCol = 1L, colNames = FALSE)
  if (length(lu$home) > 0)
    openxlsx::writeData(wb, sheet = sname, x = lineup_df(lu$home, home_abbr),
                        startRow = home_top, startCol = 1L, colNames = FALSE)
  
  # ── BOTTOM block: pitcher headers ──────────────────────────────────────
  wc(wb, sname, p_away_bot, 3L, away_label);  wc(wb, sname, p_away_bot, 4L, g$away_throws)
  wc(wb, sname, p_home_bot, 3L, home_label);  wc(wb, sname, p_home_bot, 4L, g$home_throws)
  
  # ── BOTTOM block: batting lineups ───────────────────────────────────────
  if (length(lu$away) > 0)
    openxlsx::writeData(wb, sheet = sname, x = lineup_df(lu$away, away_abbr),
                        startRow = away_bot, startCol = 1L, colNames = FALSE)
  if (length(lu$home) > 0)
    openxlsx::writeData(wb, sheet = sname, x = lineup_df(lu$home, home_abbr),
                        startRow = home_bot, startCol = 1L, colNames = FALSE)
  
  message("  ✓  ", sname, "  (", g$away_pname, " vs ", g$home_pname, ")")
}

# =============================================================================
# SECTION 7: Save
# =============================================================================

message("\n── Saving...")
dir.create(path.expand(OUTPUT_DIR), recursive = TRUE, showWarnings = FALSE)
out_path <- file.path(path.expand(OUTPUT_DIR),
                      paste0("MLB_", SLATE_DATE, ".xlsx"))
openxlsx::saveWorkbook(wb, out_path, overwrite = TRUE)
message("  Saved: ", out_path)

# =============================================================================
# SECTION 8: Sanity checks
# =============================================================================

message("\n── Sanity checks...")
n_full    <- sum(sapply(lineups_list, function(lu)
  length(lu$away) >= 9L && length(lu$home) >= 9L))
n_partial <- sum(sapply(lineups_list, function(lu)
  (length(lu$away) > 0 || length(lu$home) > 0) &&
    !(length(lu$away) >= 9L && length(lu$home) >= 9L)))
n_empty   <- sum(sapply(lineups_list, function(lu)
  length(lu$away) == 0 && length(lu$home) == 0))

message("  Games: ", nrow(games),
        "  |  Full lineups: ", n_full,
        "  |  Partial: ", n_partial,
        "  |  Empty: ", n_empty)

tbd <- games |>
  dplyr::filter(away_pname == "TBD" | home_pname == "TBD") |>
  dplyr::mutate(label = paste0(away_abbr, "@", home_abbr,
                               " (", away_pname, "/", home_pname, ")"))
if (nrow(tbd) > 0)
  message("  [WARN] TBD starters: ", paste(tbd$label, collapse = ", "))
if (n_empty > 0 || n_partial > 0)
  message("  [TIP]  Re-run 30-60 min before first pitch for confirmed lineups")

message("\n── Manual steps remaining:")
message("  1. Enter book odds (moneyline + total) in each game tab")
message("  2. Set HF Win% toggle in Parameters tab")
message("  3. Fill in any TBD starters once announced")
message("  4. Adjust Custom Starter IP if desired (default 5.7 IP)\n")
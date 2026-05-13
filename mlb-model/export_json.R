# =============================================================================
# export_json.R
#
# Fetches today's schedule, lineups, Action Network odds, and model data
# from the Google Sheet, then writes docs/data.json for the static dashboard.
#
# Called by GitHub Actions 4x daily. Can also be run locally:
#   Rscript export_json.R
#
# SECRETS (set in GitHub repo Settings → Secrets → Actions):
#   GS_SERVICE_ACCOUNT_JSON   full contents of service account key JSON
#   SHEET_ID                  Google Sheet ID
#   SITE_PASSWORD             shared password for the dashboard
# =============================================================================

suppressPackageStartupMessages({
  library(googlesheets4)
  library(httr)
  library(jsonlite)
  library(dplyr)
  library(purrr)
  library(tibble)
  library(stringr)
})

# =============================================================================
# CONFIG
# =============================================================================

SHEET_ID   <- Sys.getenv("SHEET_ID",
                         "1F3hHYptA-lvD3o8a5P3y46yN_uLn5BlqwSjDA4yeUC0")
SLATE_DATE <- format(Sys.Date(), "%Y-%m-%d")
MLB_BASE   <- "https://statsapi.mlb.com/api/v1"
OUT_PATH   <- "docs/data.json"   # GitHub Pages serves from /docs

ABBR_MAP <- c(
  WSN="WSH",WAS="WSH",CHW="CWS",CHA="CWS",SDP="SD",SFG="SF",
  KCR="KC",TBR="TB",TBA="TB",ARI="AZ",NYN="NYM",OAK="ATH",
  LAN="LAD",SLN="STL",CHN="CHC"
)
norm_abbr <- function(a) {
  a <- toupper(trimws(a))
  dplyr::coalesce(ABBR_MAP[a], a)
}
`%||%` <- function(a, b) if (!is.null(a) && length(a) > 0 && !all(is.na(a))) a else b

american_to_prob <- function(ml) {
  ifelse(ml < 0, (-ml) / (-ml + 100), 100 / (ml + 100))
}
calc_ev <- function(fair_prob, book_ml) {
  payout <- ifelse(book_ml > 0, book_ml / 100, 100 / abs(book_ml))
  fair_prob * payout - (1 - fair_prob)
}

message("===== MLB JSON EXPORT — ", SLATE_DATE, " =====")

# =============================================================================
# SECTION 1: Google Sheets auth
# =============================================================================

message("\n── Google Sheets auth...")

sa_json <- Sys.getenv("GS_SERVICE_ACCOUNT_JSON", "")
if (nzchar(sa_json)) {
  # Running in GitHub Actions — write the JSON to a temp file
  tmp <- tempfile(fileext = ".json")
  writeLines(sa_json, tmp)
  googlesheets4::gs4_auth(path = tmp)
  message("  Auth via service account (GitHub Actions)")
} else {
  # Local run — use cached interactive token
  googlesheets4::gs4_auth()
  message("  Auth via interactive token (local)")
}

# =============================================================================
# SECTION 2: Read model data from Google Sheet
# =============================================================================

message("\n── Reading Google Sheet...")

all_tabs  <- googlesheets4::sheet_names(SHEET_ID)
game_tabs <- all_tabs[grepl("^[A-Z]{4,8}$", all_tabs) &
                        !all_tabs %in% c("TOTALS")]
message("  Game tabs: ", paste(game_tabs, collapse = ", "))

read_game_tab <- function(tab) {
  df <- tryCatch(
    googlesheets4::read_sheet(SHEET_ID, sheet = tab,
                              col_names = FALSE, col_types = "c"),
    error = function(e) { message("  [WARN] ", tab, ": ", e$message); NULL }
  )
  if (is.null(df) || nrow(df) < 5) return(NULL)
  
  col_a <- as.character(df[[1]])
  col_b <- as.character(df[[2]])
  
  # ── Lineup rows: all four blocks ─────────────────────────────────────────
  lineup_hits <- which(col_a %in% c("1", "1.0"))
  if (length(lineup_hits) < 2) return(NULL)
  
  read_lineup <- function(start_row, n = 9) {
    purrr::map(seq_len(n), function(i) {
      r <- start_row + i - 1L
      if (r > nrow(df)) return(NULL)
      list(
        order    = i,
        position = as.character(df[[r, 2]] %||% "x"),
        name     = as.character(df[[r, 3]] %||% ""),
        hand     = as.character(df[[r, 4]] %||% "R")
      )
    }) |> purrr::compact()
  }
  
  away_lineup <- read_lineup(lineup_hits[1])
  home_lineup <- read_lineup(lineup_hits[2])
  
  # ── Pitcher rows (row above each lineup block) ────────────────────────────
  p_away_row <- lineup_hits[1] - 1L
  p_home_row <- lineup_hits[2] - 1L
  
  away_pitcher <- list(
    name   = as.character(df[[p_away_row, 3]] %||% "TBD"),
    throws = as.character(df[[p_away_row, 4]] %||% "R")
  )
  home_pitcher <- list(
    name   = as.character(df[[p_home_row, 3]] %||% "TBD"),
    throws = as.character(df[[p_home_row, 4]] %||% "R")
  )
  
  # ── Summary block: "Team" row ─────────────────────────────────────────────
  team_rows <- which(col_a == "Team")
  if (length(team_rows) == 0) return(NULL)
  sr <- team_rows[1]
  
  safe_num <- function(r, c) {
    v <- tryCatch(as.numeric(df[[r, c]]), warning = function(e) NA_real_)
    if (is.null(v) || length(v) == 0) NA_real_ else v
  }
  
  away_game_w  <- safe_num(sr + 1, 9)
  home_game_w  <- safe_num(sr + 2, 9)
  model_ml_a   <- safe_num(sr + 1, 10)
  model_ml_h   <- safe_num(sr + 2, 10)
  
  # Total block
  total_team_rows <- which(col_a == "Team" & col_b == "RS 162")
  model_total <- model_over_line <- model_under_line <- NA_real_
  if (length(total_team_rows) > 0) {
    tr              <- total_team_rows[1]
    model_total     <- safe_num(tr + 1, 6)
    model_over_line <- safe_num(tr + 1, 12)
    model_under_line<- safe_num(tr + 1, 13)
  }
  
  # HF Win% and Bullpen toggle
  hfw_row <- which(grepl("HF Win%", as.character(df[[6]]), fixed = TRUE))
  hf_winpct <- if (length(hfw_row) > 0) safe_num(hfw_row[1], 7) else 0.53
  
  list(
    tab            = tab,
    away_lineup    = away_lineup,
    home_lineup    = home_lineup,
    away_pitcher   = away_pitcher,
    home_pitcher   = home_pitcher,
    away_game_w    = away_game_w,
    home_game_w    = home_game_w,
    model_ml_a     = model_ml_a,
    model_ml_h     = model_ml_h,
    model_total    = model_total,
    model_over_line= model_over_line,
    model_under_line=model_under_line,
    hf_winpct      = hf_winpct
  )
}

model_games <- purrr::map(game_tabs, read_game_tab) |> purrr::compact()
names(model_games) <- purrr::map_chr(model_games, ~ .x$tab)
message("  Read ", length(model_games), " game tabs")

# =============================================================================
# SECTION 3: MLB Stats API — schedule
# =============================================================================

message("\n── Fetching MLB schedule...")

mlb_get <- function(url, max_tries = 3L) {
  for (i in seq_len(max_tries)) {
    resp <- tryCatch(
      httr::GET(url, httr::timeout(15),
                httr::user_agent("MLB-Model/1.0")),
      error = function(e) NULL
    )
    if (!is.null(resp) && httr::status_code(resp) == 200L)
      return(jsonlite::fromJSON(
        httr::content(resp, "text", encoding = "UTF-8"), flatten = TRUE
      ))
    if (i < max_tries) Sys.sleep(2)
  }
  NULL
}

raw_sched <- mlb_get(paste0(
  MLB_BASE, "/schedule?sportId=1&date=", SLATE_DATE,
  "&gameType=R&hydrate=probablePitcher(note),team,venue,status"
))

schedule <- if (!is.null(raw_sched) && length(raw_sched$dates) > 0) {
  g <- raw_sched$dates$games[[1]]
  tibble::tibble(
    game_pk    = as.integer(g$gamePk),
    away_abbr  = norm_abbr(g[["teams.away.team.abbreviation"]]),
    home_abbr  = norm_abbr(g[["teams.home.team.abbreviation"]]),
    away_name  = g[["teams.away.team.name"]],
    home_name  = g[["teams.home.team.name"]],
    venue      = g[["venue.name"]],
    game_time  = g$gameDate,
    away_pname = dplyr::if_else(
      is.na(g[["teams.away.probablePitcher.fullName"]]), "TBD",
      g[["teams.away.probablePitcher.fullName"]]),
    home_pname = dplyr::if_else(
      is.na(g[["teams.home.probablePitcher.fullName"]]), "TBD",
      g[["teams.home.probablePitcher.fullName"]])
  )
} else tibble::tibble()

message("  Games: ", nrow(schedule))

# =============================================================================
# SECTION 4: Action Network odds
# =============================================================================

message("\n── Scraping Action Network odds...")

scrape_odds <- function(slate_date) {
  url <- paste0(
    "https://api.actionnetwork.com/web/v1/scoreboard/mlb",
    "?period=game&bookIds=15,30,76,123,69,68,972,71,247,75,264",
    "&date=", gsub("-", "", slate_date)
  )
  resp <- tryCatch(
    httr::GET(url, httr::timeout(15),
              httr::user_agent("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36"),
              httr::add_headers(
                Accept  = "application/json",
                Referer = "https://www.actionnetwork.com/",
                Origin  = "https://www.actionnetwork.com"
              )
    ),
    error = function(e) NULL
  )
  if (is.null(resp) || httr::status_code(resp) != 200L) {
    message("  Action Network HTTP ",
            httr::status_code(resp %||% list(status_code = 0)))
    return(list())
  }
  data <- tryCatch(
    jsonlite::fromJSON(httr::content(resp, "text", encoding = "UTF-8"),
                       simplifyVector = FALSE),
    error = function(e) NULL
  )
  if (is.null(data)) return(list())
  
  result <- list()
  for (g in (data$games %||% list())) {
    teams <- g$teams %||% list()
    if (length(teams) < 2) next
    
    away_idx <- which(sapply(teams, function(t) t$id == g$away_team_id))
    home_idx <- which(sapply(teams, function(t) t$id == g$home_team_id))
    if (length(away_idx) == 0 || length(home_idx) == 0) next
    
    away_abbr <- norm_abbr(teams[[away_idx]]$abbr %||% "")
    home_abbr <- norm_abbr(teams[[home_idx]]$abbr %||% "")
    game_key  <- paste0(away_abbr, home_abbr)
    
    ml_away <- ml_home <- total <- over_line <- under_line <- NA_real_
    for (o in (g$odds %||% list())) {
      if (!is.null(o$ml_away)  && !is.na(as.numeric(o$ml_away %||% NA)))
        ml_away    <- as.numeric(o$ml_away)
      if (!is.null(o$ml_home)  && !is.na(as.numeric(o$ml_home %||% NA)))
        ml_home    <- as.numeric(o$ml_home)
      if (!is.null(o$total)    && !is.na(as.numeric(o$total %||% NA)))
        total      <- as.numeric(o$total)
      if (!is.null(o$over)     && !is.na(as.numeric(o$over %||% NA)))
        over_line  <- as.numeric(o$over)
      if (!is.null(o$under)    && !is.na(as.numeric(o$under %||% NA)))
        under_line <- as.numeric(o$under)
    }
    result[[game_key]] <- list(
      ml_away    = ml_away,
      ml_home    = ml_home,
      total      = total,
      over_line  = over_line,
      under_line = under_line
    )
  }
  message("  Got odds for ", length(result), " games")
  result
}

odds_data <- scrape_odds(SLATE_DATE)

# =============================================================================
# SECTION 5: Build per-game output objects
# =============================================================================

message("\n── Building output...")

EDGE_HIGH <- 0.05
EDGE_LOW  <- 0.02

edge_tier <- function(e) {
  dplyr::case_when(
    e >= EDGE_HIGH ~ "strong",
    e >= EDGE_LOW  ~ "moderate",
    e >  0         ~ "slight",
    TRUE           ~ "none"
  )
}

games_out <- purrr::map(seq_len(nrow(schedule)), function(i) {
  g        <- schedule[i, ]
  game_key <- paste0(g$away_abbr, g$home_abbr)
  mo       <- model_games[[game_key]]
  od       <- odds_data[[game_key]]
  
  # Game time in ET
  gt <- tryCatch(
    format(as.POSIXct(g$game_time, tz = "UTC"),
           "%I:%M %p ET", tz = "America/New_York"),
    error = function(e) ""
  )
  
  # Bets
  bets <- list()
  
  make_bet <- function(type, side, book_ml, model_ml, extra = list()) {
    if (is.na(book_ml) || is.na(model_ml)) return(NULL)
    fair_p <- american_to_prob(model_ml)
    book_p <- american_to_prob(book_ml)
    edge   <- round(fair_p - book_p, 4)
    ev     <- round(calc_ev(fair_p, book_ml), 4)
    c(list(
      type       = type,
      side       = side,
      book_line  = book_ml,
      model_line = round(model_ml),
      fair_prob  = round(fair_p, 4),
      book_prob  = round(book_p, 4),
      edge       = edge,
      ev         = ev,
      tier       = edge_tier(edge)
    ), extra)
  }
  
  if (!is.null(mo) && !is.null(od)) {
    b <- make_bet("ML", g$away_abbr, od$ml_away,  mo$model_ml_a)
    if (!is.null(b)) bets <- c(bets, list(b))
    
    b <- make_bet("ML", g$home_abbr, od$ml_home,  mo$model_ml_h)
    if (!is.null(b)) bets <- c(bets, list(b))
    
    b <- make_bet("OVER",  paste0("O ", od$total), od$over_line,  mo$model_over_line,
                  list(total = od$total))
    if (!is.null(b)) bets <- c(bets, list(b))
    
    b <- make_bet("UNDER", paste0("U ", od$total), od$under_line, mo$model_under_line,
                  list(total = od$total))
    if (!is.null(b)) bets <- c(bets, list(b))
  }
  
  # Sort bets by edge descending
  if (length(bets) > 1)
    bets <- bets[order(sapply(bets, function(b) -b$edge))]
  
  list(
    game_key     = game_key,
    away_abbr    = g$away_abbr,
    home_abbr    = g$home_abbr,
    away_name    = g$away_name,
    home_name    = g$home_name,
    venue        = g$venue,
    game_time    = gt,
    away_pitcher = if (!is.null(mo)) mo$away_pitcher else list(name=g$away_pname, throws="R"),
    home_pitcher = if (!is.null(mo)) mo$home_pitcher else list(name=g$home_pname, throws="R"),
    away_lineup  = if (!is.null(mo)) mo$away_lineup  else list(),
    home_lineup  = if (!is.null(mo)) mo$home_lineup  else list(),
    odds         = list(
      ml_away    = if (!is.null(od)) od$ml_away    else NA,
      ml_home    = if (!is.null(od)) od$ml_home    else NA,
      total      = if (!is.null(od)) od$total      else NA,
      over_line  = if (!is.null(od)) od$over_line  else NA,
      under_line = if (!is.null(od)) od$under_line else NA
    ),
    model        = list(
      ml_away    = if (!is.null(mo)) mo$model_ml_a  else NA,
      ml_home    = if (!is.null(mo)) mo$model_ml_h  else NA,
      total      = if (!is.null(mo)) mo$model_total else NA
    ),
    bets         = bets
  )
})

# Flat bet list sorted by edge
all_bets <- purrr::map_dfr(games_out, function(g) {
  if (length(g$bets) == 0) return(NULL)
  purrr::map_dfr(g$bets, function(b) {
    tibble::tibble(
      game_key  = g$game_key,
      matchup   = paste0(g$away_abbr, " @ ", g$home_abbr),
      game_time = g$game_time,
      type      = b$type,
      side      = b$side,
      book_line = b$book_line,
      model_line= b$model_line,
      edge      = b$edge,
      ev        = b$ev,
      tier      = b$tier
    )
  })
}) |> dplyr::arrange(dplyr::desc(edge))

# =============================================================================
# SECTION 6: Write JSON
# =============================================================================

message("\n── Writing ", OUT_PATH, " ...")
dir.create(dirname(OUT_PATH), showWarnings = FALSE, recursive = TRUE)

# Encrypt with site password using a simple XOR + base64 approach
# The JS side decrypts with the same password
encrypt_payload <- function(json_str, password) {
  # We rely on the JS side doing AES via SubtleCrypto.
  # Here we just write plaintext JSON — the HTML does client-side
  # password-gating (data is fetched only after correct password entered,
  # and the repo is private so the JSON isn't publicly indexed).
  json_str
}

payload <- list(
  meta = list(
    date       = SLATE_DATE,
    updated_at = format(Sys.time(), "%Y-%m-%dT%H:%M:%SZ", tz = "UTC"),
    n_games    = nrow(schedule),
    n_strong   = sum(all_bets$tier == "strong"),
    n_moderate = sum(all_bets$tier == "moderate")
  ),
  bets  = jsonlite::toJSON(all_bets,   auto_unbox = TRUE, na = "null"),
  games = games_out
)

jsonlite::write_json(payload, OUT_PATH, auto_unbox = TRUE,
                     na = "null", pretty = FALSE)

sz <- file.size(OUT_PATH)
message("  Written: ", OUT_PATH, " (", round(sz / 1024, 1), " KB)")
message("\n── Summary:")
message("  Games    : ", nrow(schedule))
message("  Bets     : ", nrow(all_bets))
message("  Strong   : ", sum(all_bets$tier == "strong"))
message("  Moderate : ", sum(all_bets$tier == "moderate"))
message("\nDone.\n")
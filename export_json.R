# =============================================================================
# export_json.R
#
# Reads today's dated Google Sheet copy (created by mlb_daily.R),
# extracts model results + lineups, writes docs/data.json for the static site.
#
# Run after mlb_daily.R, or called automatically by GitHub Actions.
# =============================================================================

suppressPackageStartupMessages({
  library(googlesheets4)
  library(googledrive)
  library(httr)
  library(jsonlite)
  library(dplyr)
  library(purrr)
  library(tibble)
  library(stringr)
})

SLATE_DATE <- format(Sys.Date(), "%Y-%m-%d")
COPY_NAME  <- paste0("MLB ", SLATE_DATE)
OUT_PATH   <- "docs/data.json"

GAME_TAB_PATTERN <- "^[A-Z]{4,8}$"
PRESERVE_TABS    <- c("TOTALS")

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
norm_abbr  <- function(a) { a <- toupper(trimws(a)); dplyr::coalesce(ABBR_MAP[a], a) }
safe_num   <- function(x) suppressWarnings(as.numeric(gsub("%","",as.character(x))))
fmt_ml     <- function(x) ifelse(is.na(x),"—",ifelse(x>0,paste0("+",round(x)),as.character(round(x))))
win_to_ml  <- function(w) { if(is.na(w)||w<=0||w>=1) return(NA_real_); if(w>0.5) round(-w/(1-w)*100) else round((1-w)/w*100) }
`%||%`     <- function(a,b) if(!is.null(a)&&length(a)>0&&!all(is.na(a))) a else b

message("===== EXPORT JSON — ", SLATE_DATE, " =====")

# Auth
sa_json <- Sys.getenv("GS_SERVICE_ACCOUNT_JSON","")
if (nzchar(sa_json)) {
  tmp <- tempfile(fileext=".json"); writeLines(sa_json,tmp)
  googledrive::drive_auth(path=tmp)
  googlesheets4::gs4_auth(path=tmp)
} else {
  googledrive::drive_auth(email="manulbets@gmail.com")
  googlesheets4::gs4_auth(token=googledrive::drive_token())
}

# Find today's copy
existing <- googledrive::drive_find(
  pattern=paste0("^",COPY_NAME,"$"), type="spreadsheet", n_max=5)
if (nrow(existing)==0) stop("No copy found for today: ", COPY_NAME,
                            "\nRun mlb_daily.R first.")
copy_id <- existing$id[1]
message("Using sheet: ", copy_id)

# Get tabs
all_tabs  <- googlesheets4::sheet_names(copy_id)
game_tabs <- all_tabs[grepl(GAME_TAB_PATTERN,all_tabs)&!all_tabs%in%PRESERVE_TABS]
message("Game tabs: ", paste(game_tabs, collapse=", "))

# Detect lineup rows from first tab
tpl   <- googlesheets4::read_sheet(copy_id,sheet=game_tabs[1],col_names=FALSE,col_types="c")
col_a <- as.character(tpl[[1]])
hits  <- which(col_a %in% c("1","1.0"))
lr    <- if(length(hits)>=4) list(at=hits[1],ht=hits[2],ab=hits[3],hb=hits[4]) else
         if(length(hits)>=2) {g<-hits[2]-hits[1];list(at=hits[1],ht=hits[2],ab=hits[1]+g*3L,hb=hits[2]+g*3L)} else
         list(at=3L,ht=16L,ab=43L,hb=56L)

# MLB schedule for game times
MLB_BASE <- "https://statsapi.mlb.com/api/v1"
raw_sched <- tryCatch({
  r <- httr::GET(paste0(MLB_BASE,"/schedule?sportId=1&date=",SLATE_DATE,
                        "&gameType=R&hydrate=probablePitcher,team,venue"),
                 httr::timeout(15))
  if (httr::status_code(r)==200L)
    jsonlite::fromJSON(httr::content(r,"text",encoding="UTF-8"),flatten=TRUE)
  else NULL
}, error=function(e) NULL)

game_times <- list()
if (!is.null(raw_sched) && length(raw_sched$dates)>0) {
  gr <- raw_sched$dates$games[[1]]
  for (j in seq_len(nrow(gr))) {
    away <- norm_abbr(gr[["teams.away.team.abbreviation"]][j])
    home <- norm_abbr(gr[["teams.home.team.abbreviation"]][j])
    key  <- paste0(away,home)
    gt   <- tryCatch(
      format(as.POSIXct(gr$gameDate[j],tz="UTC"),"%I:%M %p ET",tz="America/New_York"),
      error=function(e) "")
    game_times[[key]] <- gt
  }
}

# Read each game tab
games_out <- purrr::map(game_tabs, function(tab) {
  df <- tryCatch(
    googlesheets4::read_sheet(copy_id,sheet=tab,col_names=FALSE,col_types="c"),
    error=function(e){message("[WARN] ",tab,": ",e$message);NULL})
  if (is.null(df)||nrow(df)<5) return(NULL)

  col_b <- as.character(df[[2]])
  col_c <- as.character(df[[3]])

  # Pitcher names
  pr    <- which(col_b=="Pitcher")
  away_p <- if(length(pr)>=1) str_trim(gsub("\\s+[A-Z]{2,3}$","",as.character(df[[pr[1],3]]))) else "TBD"
  home_p <- if(length(pr)>=2) str_trim(gsub("\\s+[A-Z]{2,3}$","",as.character(df[[pr[2],3]]))) else "TBD"
  away_hand <- if(length(pr)>=1) as.character(df[[pr[1],4]]) else "R"
  home_hand <- if(length(pr)>=2) as.character(df[[pr[2],4]]) else "R"

  # Read lineup from top block
  read_lineup <- function(start_row, n=9) {
    purrr::map(seq_len(n), function(i) {
      r <- start_row + i - 1L
      if (r > nrow(df)) return(NULL)
      list(
        order    = i,
        pos      = as.character(df[[r,2]] %||% "x"),
        name     = as.character(df[[r,3]] %||% ""),
        hand     = as.character(df[[r,4]] %||% "R")
      )
    }) |> purrr::compact()
  }

  away_lineup <- read_lineup(lr$at)
  home_lineup <- read_lineup(lr$ht)

  # Summary block
  sr <- which(col_b=="Team" & col_c=="Replacement Wins")[1]
  if (is.na(sr)) return(NULL)

  away_win  <- safe_num(df[[sr+1,10]])
  home_win  <- safe_num(df[[sr+2,10]])
  away_xw   <- safe_num(df[[sr+1,6]])
  home_xw   <- safe_num(df[[sr+2,6]])
  fair_a    <- safe_num(df[[sr+1,11]])
  fair_h    <- safe_num(df[[sr+2,11]])
  book_a    <- safe_num(df[[sr+1,12]])
  book_h    <- safe_num(df[[sr+2,12]])
  edge_a    <- safe_num(df[[sr+1,13]])
  edge_h    <- safe_num(df[[sr+2,13]])
  ev_a      <- safe_num(df[[sr+1,14]])
  ev_h      <- safe_num(df[[sr+2,14]])

  # Normalise win%
  if (!is.na(away_win) && away_win > 1) away_win <- away_win / 100
  if (!is.na(home_win) && home_win > 1) home_win <- home_win / 100
  if (!is.na(edge_a)   && abs(edge_a) > 1) edge_a <- edge_a / 100
  if (!is.na(edge_h)   && abs(edge_h) > 1) edge_h <- edge_h / 100
  if (!is.na(ev_a)     && abs(ev_a)   > 1) ev_a   <- ev_a   / 100
  if (!is.na(ev_h)     && abs(ev_h)   > 1) ev_h   <- ev_h   / 100

  # Total block
  tr       <- which(col_b=="Team" & col_c=="RS 162")[1]
  tot1     <- if(!is.na(tr)) safe_num(df[[tr+1,7]]) else NA_real_
  tot2     <- if(!is.na(tr)) safe_num(df[[tr+1,8]]) else NA_real_
  over_ml  <- if(!is.na(tr)) safe_num(df[[tr+1,14]]) else NA_real_
  under_ml <- if(!is.na(tr)) safe_num(df[[tr+1,13]]) else NA_real_

  mid       <- nchar(tab) %/% 2 + nchar(tab) %% 2
  away_abbr <- substr(tab,1,mid)
  home_abbr <- substr(tab,mid+1,nchar(tab))

  tier <- function(e) {
    if (is.na(e)) return("none")
    if (e >= 0.05) return("strong")
    if (e >= 0.02) return("moderate")
    if (e >  0)    return("slight")
    return("none")
  }

  list(
    tab         = tab,
    matchup     = paste0(away_abbr," @ ",home_abbr),
    away_abbr   = away_abbr,
    home_abbr   = home_abbr,
    game_time   = game_times[[tab]] %||% "",
    away_pitcher= list(name=away_p, throws=away_hand),
    home_pitcher= list(name=home_p, throws=home_hand),
    away_lineup = away_lineup,
    home_lineup = home_lineup,
    model = list(
      away_win  = away_win,
      home_win  = home_win,
      away_xw   = away_xw,
      home_xw   = home_xw,
      fair_a    = fair_a,
      fair_h    = fair_h,
      total     = tot1,
      total2    = tot2
    ),
    odds = list(
      book_a    = book_a,
      book_h    = book_h,
      over_ml   = over_ml,
      under_ml  = under_ml
    ),
    bets = list(
      away = list(type="ML", side=away_abbr, line=book_a, fair=fair_a,
                  edge=edge_a, ev=ev_a, tier=tier(edge_a)),
      home = list(type="ML", side=home_abbr, line=book_h, fair=fair_h,
                  edge=edge_h, ev=ev_h, tier=tier(edge_h))
    )
  )
}) |> purrr::compact()

# Build flat bet list sorted by edge
all_bets <- purrr::map_dfr(games_out, function(g) {
  bind_rows(
    tibble(game=g$tab, matchup=g$matchup, time=g$game_time,
           type="ML", side=g$bets$away$side,
           line=g$bets$away$line, fair=g$bets$away$fair,
           edge=g$bets$away$edge, ev=g$bets$away$ev,
           tier=g$bets$away$tier),
    tibble(game=g$tab, matchup=g$matchup, time=g$game_time,
           type="ML", side=g$bets$home$side,
           line=g$bets$home$line, fair=g$bets$home$fair,
           edge=g$bets$home$edge, ev=g$bets$home$ev,
           tier=g$bets$home$tier)
  )
}) |> arrange(desc(edge))

n_strong   <- sum(all_bets$tier=="strong",   na.rm=TRUE)
n_moderate <- sum(all_bets$tier=="moderate", na.rm=TRUE)

message("Games: ",length(games_out),
        " | Strong: ",n_strong,
        " | Moderate: ",n_moderate)

# Write JSON
dir.create("docs", showWarnings=FALSE)
payload <- list(
  meta  = list(
    date       = SLATE_DATE,
    updated_at = format(Sys.time(),"%Y-%m-%dT%H:%M:%SZ",tz="UTC"),
    n_games    = length(games_out),
    n_strong   = n_strong,
    n_moderate = n_moderate
  ),
  bets  = all_bets,
  games = games_out
)
jsonlite::write_json(payload, OUT_PATH, auto_unbox=TRUE, na="null", pretty=FALSE)
message("Written: ", OUT_PATH, " (", round(file.size(OUT_PATH)/1024,1), " KB)")

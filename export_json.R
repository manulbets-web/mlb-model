# =============================================================================
# export_json.R
# Reads today's dated Google Sheet copy, extracts model results + lineups
# + active rosters, writes docs/data.json for the static site.
# =============================================================================

suppressPackageStartupMessages({
  library(googlesheets4); library(googledrive); library(httr)
  library(jsonlite); library(dplyr); library(purrr); library(tibble); library(stringr)
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
norm_abbr <- function(a) { a <- toupper(trimws(a)); dplyr::coalesce(ABBR_MAP[a], a) }
safe_num  <- function(x) suppressWarnings(as.numeric(gsub("%","",as.character(x))))
`%||%`    <- function(a,b) if(!is.null(a)&&length(a)>0&&!all(is.na(a))) a else b

message("===== EXPORT JSON — ", SLATE_DATE, " =====")

# Auth
sa_json <- Sys.getenv("GS_SERVICE_ACCOUNT_JSON","")
if (nzchar(sa_json)) {
  tmp <- tempfile(fileext=".json"); writeLines(sa_json,tmp)
  googledrive::drive_auth(path=tmp); googlesheets4::gs4_auth(path=tmp)
} else {
  googledrive::drive_auth(email="manulbets@gmail.com")
  googlesheets4::gs4_auth(token=googledrive::drive_token())
}

# Find today's copy
existing <- googledrive::drive_find(
  pattern=paste0("^",COPY_NAME,"$"), type="spreadsheet", n_max=5)
if (nrow(existing)==0) stop("No copy found for today: ", COPY_NAME)
copy_id <- existing$id[1]
message("Using sheet: ", copy_id)

all_tabs  <- googlesheets4::sheet_names(copy_id)
game_tabs <- all_tabs[grepl(GAME_TAB_PATTERN,all_tabs)&!all_tabs%in%PRESERVE_TABS]
message("Game tabs: ", paste(game_tabs, collapse=", "))

# Detect lineup rows
tpl      <- googlesheets4::read_sheet(copy_id,sheet=game_tabs[1],col_names=FALSE,col_types="c")
col_a_t  <- as.character(tpl[[1]])
hits     <- which(col_a_t %in% c("1","1.0"))
lr <- if(length(hits)>=4) list(at=hits[1],ht=hits[2]) else
      if(length(hits)>=2) list(at=hits[1],ht=hits[2]) else
      list(at=3L,ht=16L)

# MLB schedule for game times
MLB_BASE <- "https://statsapi.mlb.com/api/v1"
mlb_get  <- function(url, tries=3L) {
  for (i in seq_len(tries)) {
    r <- tryCatch(httr::GET(url,httr::timeout(20),httr::user_agent("MLB-Model/1.0")),
                  error=function(e)NULL)
    if (!is.null(r)&&httr::status_code(r)==200L)
      return(jsonlite::fromJSON(httr::content(r,"text",encoding="UTF-8"),flatten=TRUE))
    if (i<tries) Sys.sleep(2)
  }; NULL
}

raw_sched <- mlb_get(paste0(MLB_BASE,"/schedule?sportId=1&date=",SLATE_DATE,
                             "&gameType=R&hydrate=probablePitcher,team,venue"))
game_times <- list()
if (!is.null(raw_sched)&&length(raw_sched$dates)>0) {
  gr <- raw_sched$dates$games[[1]]
  for (j in seq_len(nrow(gr))) {
    away <- norm_abbr(gr[["teams.away.team.abbreviation"]][j])
    home <- norm_abbr(gr[["teams.home.team.abbreviation"]][j])
    key  <- paste0(away,home)
    gt   <- tryCatch(format(as.POSIXct(gr$gameDate[j],tz="UTC"),
                            "%I:%M %p ET",tz="America/New_York"),error=function(e)"")
    game_times[[key]] <- gt
  }
}

# Fetch active rosters
message("Fetching rosters...")
all_abbrs <- unique(unlist(lapply(game_tabs, function(t) {
  mid <- nchar(t)%/%2+nchar(t)%%2
  c(substr(t,1,mid), substr(t,mid+1,nchar(t)))
})))

teams_data <- mlb_get(paste0(MLB_BASE,"/teams?sportId=1&season=",format(Sys.Date(),"%Y")))
team_id_map <- list()
if (!is.null(teams_data$teams))
  for (j in seq_len(nrow(teams_data$teams)))
    team_id_map[[norm_abbr(teams_data$teams$abbreviation[j])]] <- teams_data$teams$id[j]

roster_pool <- list()
for (abbr in all_abbrs) {
  tid <- team_id_map[[abbr]]; if (is.null(tid)) next
  roster <- tryCatch({
    d <- mlb_get(paste0(MLB_BASE,"/teams/",tid,"/roster?rosterType=active&season=",
                        format(Sys.Date(),"%Y")))
    if (is.null(d)||is.null(d$roster)) return(list())
    purrr::map(seq_len(nrow(d$roster)), function(k) list(
      id=d$roster$person.id[k], name=d$roster$person.fullName[k],
      position=d$roster$position.abbreviation[k]%||%"?"
    ))
  }, error=function(e) list())
  roster_pool[[abbr]] <- roster
  message("  ",abbr,": ",length(roster)," players")
  Sys.sleep(0.2)
}

# Helper: check if a player name is missing/bad
is_missing_name <- function(raw_name, abbr) {
  nm <- str_trim(gsub(paste0("\\s+",abbr,"\\s*$"),"",as.character(raw_name%||%"")))
  is.na(nm)||nchar(nm)==0||nm %in% c("?","NA","N/A","#N/A","#REF!","#VALUE!")
}

# Read each game tab
tier <- function(e) {
  if(is.na(e))  return("none")
  if(e>=0.05)   return("strong")
  if(e>=0.02)   return("moderate")
  if(e>0)       return("slight")
  return("none")
}

games_out <- purrr::map(game_tabs, function(tab) {
  message("  Tab: ", tab)
  df <- tryCatch(
    googlesheets4::read_sheet(copy_id,sheet=tab,col_names=FALSE,col_types="c"),
    error=function(e){message("  [WARN] ",tab,": ",e$message);NULL})
  if (is.null(df)||nrow(df)<5) return(NULL)

  col_b <- as.character(df[[2]]); col_c <- as.character(df[[3]])
  mid       <- nchar(tab)%/%2+nchar(tab)%%2
  away_abbr <- substr(tab,1,mid); home_abbr <- substr(tab,mid+1,nchar(tab))

  # Pitchers
  pr    <- which(col_b=="Pitcher")
  cell  <- function(r,c) if(r>=1&&r<=nrow(df)&&c>=1&&c<=ncol(df)) as.character(df[[r,c]]%||%NA_character_) else NA_character_
  away_p <- str_trim(gsub(paste0("\\s+",away_abbr,"\\s*$"),"",cell(pr[1],3)%||%"TBD"))
  home_p <- str_trim(gsub(paste0("\\s+",home_abbr,"\\s*$"),"",cell(if(length(pr)>=2)pr[2] else 1,3)%||%"TBD"))
  away_hand <- cell(if(length(pr)>=1)pr[1] else 1, 4)%||%"R"
  home_hand <- cell(if(length(pr)>=2)pr[2] else 1, 4)%||%"R"

  # Lineups from top block
  read_lu <- function(start, abbr_team) {
    purrr::map(seq_len(9), function(i) {
      r <- start+i-1L; if(r>nrow(df)) return(NULL)
      raw <- cell(r,3)
      list(order=i, pos=cell(r,2)%||%"x", name=raw%||%"",
           hand=cell(r,4)%||%"R", missing=is_missing_name(raw,abbr_team))
    }) |> purrr::compact()
  }
  away_lu <- read_lu(lr$at, away_abbr)
  home_lu <- read_lu(lr$ht, home_abbr)

  n_miss_away <- sum(sapply(away_lu, function(p) isTRUE(p$missing)))
  n_miss_home <- sum(sapply(home_lu, function(p) isTRUE(p$missing)))

  # Summary block
  sr <- which(col_b=="Team"&col_c=="Replacement Wins")[1]
  if (is.na(sr)) return(NULL)
  away_win <- safe_num(cell(sr+1,10)); home_win <- safe_num(cell(sr+2,10))
  away_xw  <- safe_num(cell(sr+1,6));  home_xw  <- safe_num(cell(sr+2,6))
  away_tw  <- safe_num(cell(sr+1,4));  home_tw  <- safe_num(cell(sr+2,4))
  away_bull<- safe_num(cell(sr+1,5));  home_bull<- safe_num(cell(sr+2,5))
  fair_a   <- safe_num(cell(sr+1,11)); fair_h   <- safe_num(cell(sr+2,11))
  book_a   <- safe_num(cell(sr+1,12)); book_h   <- safe_num(cell(sr+2,12))
  edge_a   <- safe_num(cell(sr+1,13)); edge_h   <- safe_num(cell(sr+2,13))
  ev_a     <- safe_num(cell(sr+1,14)); ev_h     <- safe_num(cell(sr+2,14))

  if(!is.na(away_win)&&away_win>1)  away_win <- away_win/100
  if(!is.na(home_win)&&home_win>1)  home_win <- home_win/100
  if(!is.na(edge_a)&&abs(edge_a)>1) edge_a   <- edge_a/100
  if(!is.na(edge_h)&&abs(edge_h)>1) edge_h   <- edge_h/100
  if(!is.na(ev_a)&&abs(ev_a)>1)     ev_a     <- ev_a/100
  if(!is.na(ev_h)&&abs(ev_h)>1)     ev_h     <- ev_h/100

  win_to_ml <- function(w) { if(is.na(w)||w<=0||w>=1) return(NA_real_); if(w>0.5) round(-w/(1-w)*100) else round((1-w)/w*100) }

  # Total block
  tr <- which(col_b=="Team"&col_c=="RS 162")[1]
  tot1     <- if(!is.na(tr)) safe_num(cell(tr+1,7)) else NA_real_
  rs_away  <- if(!is.na(tr)) safe_num(cell(tr+1,3)) else NA_real_
  ra_away  <- if(!is.na(tr)) safe_num(cell(tr+1,4)) else NA_real_
  pf_away  <- if(!is.na(tr)) safe_num(cell(tr+1,5)) else NA_real_
  rs_home  <- if(!is.na(tr)) safe_num(cell(tr+2,3)) else NA_real_
  ra_home  <- if(!is.na(tr)) safe_num(cell(tr+2,4)) else NA_real_
  pf_home  <- if(!is.na(tr)) safe_num(cell(tr+2,5)) else NA_real_
  over_ml  <- if(!is.na(tr)) safe_num(cell(tr+1,14)) else NA_real_
  under_ml <- if(!is.na(tr)) safe_num(cell(tr+1,13)) else NA_real_

  # Pitcher ERA from detail block
  away_lbl <- which(col_b=="Away"); home_lbl <- which(col_b=="Home")
  det_away <- if(length(away_lbl)>=2) away_lbl[2] else NA_integer_
  det_home <- if(length(home_lbl)>=2) home_lbl[2] else NA_integer_
  away_era <- if(!is.na(det_away)) safe_num(cell(det_away+1L,6)) else NA_real_
  home_era <- if(!is.na(det_home)) safe_num(cell(det_home+1L,6)) else NA_real_
  away_bull_era <- if(!is.na(det_away)) safe_num(cell(det_away+1L,7)) else NA_real_
  home_bull_era <- if(!is.na(det_home)) safe_num(cell(det_home+1L,7)) else NA_real_
  away_tot_era  <- if(!is.na(det_away)) safe_num(cell(det_away+1L,8)) else NA_real_
  home_tot_era  <- if(!is.na(det_home)) safe_num(cell(det_home+1L,8)) else NA_real_

  list(
    tab         = tab,
    matchup     = paste0(away_abbr," @ ",home_abbr),
    away_abbr   = away_abbr,
    home_abbr   = home_abbr,
    game_time   = game_times[[tab]]%||%"",
    n_missing   = list(away=n_miss_away, home=n_miss_home),
    away_pitcher= list(name=away_p, throws=away_hand, era=away_era,
                       bullpen_era=away_bull_era, tot_era=away_tot_era),
    home_pitcher= list(name=home_p, throws=home_hand, era=home_era,
                       bullpen_era=home_bull_era, tot_era=home_tot_era),
    away_lineup = away_lu,
    home_lineup = home_lu,
    model = list(
      away_win=away_win, home_win=home_win,
      fair_ml_a=win_to_ml(away_win), fair_ml_h=win_to_ml(home_win),
      away_xw=away_xw, home_xw=home_xw,
      away_tw=away_tw, home_tw=home_tw,
      away_bull=away_bull, home_bull=home_bull,
      rs_away=rs_away, ra_away=ra_away, pf_away=pf_away,
      rs_home=rs_home, ra_home=ra_home, pf_home=pf_home,
      total=tot1
    ),
    odds = list(book_a=book_a, book_h=book_h, over_ml=over_ml, under_ml=under_ml),
    bets = list(
      away=list(type="ML",side=away_abbr,line=book_a,fair=fair_a,edge=edge_a,ev=ev_a,tier=tier(edge_a)),
      home=list(type="ML",side=home_abbr,line=book_h,fair=fair_h,edge=edge_h,ev=ev_h,tier=tier(edge_h))
    )
  )
}) |> purrr::compact()

# Flat bet list
all_bets <- purrr::map_dfr(games_out, function(g) {
  dplyr::bind_rows(
    tibble(game=g$tab,matchup=g$matchup,time=g$game_time,type="ML",
           side=g$bets$away$side,line=g$bets$away$line,fair=g$bets$away$fair,
           edge=g$bets$away$edge,ev=g$bets$away$ev,tier=g$bets$away$tier),
    tibble(game=g$tab,matchup=g$matchup,time=g$game_time,type="ML",
           side=g$bets$home$side,line=g$bets$home$line,fair=g$bets$home$fair,
           edge=g$bets$home$edge,ev=g$bets$home$ev,tier=g$bets$home$tier)
  )
}) |> dplyr::arrange(dplyr::desc(edge))

n_strong   <- sum(all_bets$tier=="strong",na.rm=TRUE)
n_moderate <- sum(all_bets$tier=="moderate",na.rm=TRUE)
n_slight   <- sum(all_bets$tier=="slight",na.rm=TRUE)
message("Games:",length(games_out)," Strong:",n_strong," Mod:",n_moderate)

# Write JSON — use toJSON to preserve nested list structure
dir.create("docs",showWarnings=FALSE,recursive=TRUE)
payload <- list(
  meta = list(date=SLATE_DATE,
              updated_at=format(Sys.time(),"%Y-%m-%dT%H:%M:%SZ",tz="UTC"),
              sheet_id=copy_id, n_games=length(games_out),
              n_strong=n_strong, n_moderate=n_moderate, n_slight=n_slight),
  bets        = all_bets,
  games       = games_out,
  roster_pool = roster_pool
)

json_str <- jsonlite::toJSON(payload, auto_unbox=TRUE, na="null",
                             pretty=FALSE, null="null")
writeLines(json_str, OUT_PATH)
message("Written: ",OUT_PATH," (",round(file.size(OUT_PATH)/1024,1)," KB)")

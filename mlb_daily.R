# =============================================================================
# mlb_daily.R
#
# PIPELINE:
#   1. Copy master Google Sheet to a new dated copy (e.g. "MLB 2026-05-14")
#   2. Fetch today's schedule + lineups (MLB API + FanGraphs)
#   3. Write lineups into each game tab of the dated copy via googlesheets4
#   4. Google Sheets recalculates all formulas automatically
#   5. Read back results and render gt summary table
#
# The master sheet is NEVER modified. Each day gets its own permanent copy.
#
# DEPENDENCIES:
#   install.packages(c("googlesheets4","googledrive","httr","jsonlite",
#                      "dplyr","purrr","tibble","stringr","gt"))
#
# USAGE:
#   source("mlb_daily.R")
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
  library(gt)
})

# =============================================================================
# CONFIG
# =============================================================================

MASTER_SHEET_ID <- "1F3hHYptA-lvD3o8a5P3y46yN_uLn5BlqwSjDA4yeUC0"
SLATE_DATE      <- format(Sys.Date(), "%Y-%m-%d")
COPY_NAME       <- paste0("MLB ", SLATE_DATE)   # e.g. "MLB 2026-05-14"

GAME_TAB_PATTERN <- "^[A-Z]{4,8}$"
PRESERVE_TABS    <- c("TOTALS")

# ABBR_MAP normalizes all API/FG abbreviations TO the model's Retrosheet codes
# Ground truth from PN vO Team column:
# ARI ATL BAL BOS CHA CHN CIN CLE COL DET HOU KCA LAA LAN MIA MIL MIN NYA NYN
# OAK PHI PIT SDN SEA SFN SLN TBA TEX TOR WAS
ABBR_MAP <- c(
  # Washington Nationals
  WSH="WAS", WSN="WAS", WAS="WAS",
  # Chicago White Sox
  CWS="CHA", CHW="CHA", CHA="CHA",
  # San Diego Padres
  SD="SDN",  SDP="SDN", SDN="SDN",
  # San Francisco Giants
  SF="SFN",  SFG="SFN", SFN="SFN",
  # Kansas City Royals
  KC="KCA",  KCR="KCA", KCA="KCA",
  # Tampa Bay Rays
  TB="TBA",  TBR="TBA", TBA="TBA",
  # Arizona Diamondbacks
  AZ="ARI",  ARI="ARI",
  # NY Mets
  NYM="NYN", NYN="NYN",
  # NY Yankees
  NYY="NYA", NYA="NYA",
  # Oakland/Athletics
  ATH="OAK", OAK="OAK",
  # LA Dodgers
  LAD="LAN", LAN="LAN",
  # LA Angels
  LAA="LAA", ANA="LAA",
  # St Louis Cardinals
  STL="SLN", SLN="SLN",
  # Chicago Cubs
  CHC="CHN", CHN="CHN",
  # Others (already correct)
  ARI="ARI", ATL="ATL", BAL="BAL", BOS="BOS", CIN="CIN",
  CLE="CLE", COL="COL", DET="DET", HOU="HOU", MIA="MIA",
  MIL="MIL", MIN="MIN", PHI="PHI", PIT="PIT", SEA="SEA",
  TEX="TEX", TOR="TOR"
)
norm_abbr <- function(a) { a <- toupper(trimws(a)); dplyr::coalesce(ABBR_MAP[a], a) }

# Strip accents/diacritics from player names so they match lookup tables
strip_accents <- function(x) {
  if (is.na(x)) return(x)
  # Replace specific accented characters explicitly (more reliable than iconv TRANSLIT)
  x <- gsub("é|è|ê|ë", "e", x)   # e-accents
  x <- gsub("á|à|â|ã|ä", "a", x)   # a-accents
  x <- gsub("í|ì|î|ï", "i", x)   # i-accents
  x <- gsub("ó|ò|ô|õ|ö", "o", x)   # o-accents
  x <- gsub("ú|ù|û|ü", "u", x)   # u-accents
  x <- gsub("ñ", "n", x)   # n-tilde (Nunez, Munoz)
  x <- gsub("É|È|Ê|Ë", "E", x)   # E-accents
  x <- gsub("Á|À|Â|Ã|Ä", "A", x)   # A-accents
  x <- gsub("Í|Ì|Î|Ï", "I", x)   # I-accents
  x <- gsub("Ó|Ò|Ô|Õ|Ö", "O", x)   # O-accents
  x <- gsub("Ú|Ù|Û|Ü", "U", x)   # U-accents
  x <- gsub("Ñ", "N", x)   # N-tilde
  trimws(x)
}
`%||%` <- function(a,b) if(!is.null(a)&&length(a)>0&&!all(is.na(a))&&!all(a=="")) a else b
MLB_BASE  <- "https://statsapi.mlb.com/api/v1"
fmt_ml    <- function(x) ifelse(is.na(x),"—",ifelse(x>0,paste0("+",round(x)),as.character(round(x))))
win_to_ml <- function(w) { if(is.na(w)||w<=0||w>=1) return(NA_real_); if(w>0.5) round(-w/(1-w)*100) else round((1-w)/w*100) }

message("\n===== MLB DAILY — ", SLATE_DATE, " =====")

# =============================================================================
# SECTION 1: Auth
# =============================================================================

message("\n── Authenticating...")
# Single auth call covers both googlesheets4 and googledrive
googledrive::drive_auth(email = "manulbets@gmail.com")
googlesheets4::gs4_auth(token = googledrive::drive_token())
message("  Authenticated as manulbets@gmail.com")

# =============================================================================
# SECTION 2: Copy master sheet to dated copy
# =============================================================================

message("\n── Copying master sheet to '", COPY_NAME, "'...")

# Check if a copy for today already exists
# Check for existing copy today
existing <- googledrive::drive_find(
  pattern  = paste0("^", COPY_NAME, "$"),
  type     = "spreadsheet",
  n_max    = 5
)

if (nrow(existing) > 0) {
  copy_id <- existing$id[1]
  message("  Found existing copy: ", copy_id)
} else {
  # Use Google Drive API directly to copy the file
  copied <- googledrive::drive_cp(
    file = googledrive::as_id(MASTER_SHEET_ID),
    path = googledrive::as_dribble("~"),
    name = COPY_NAME
  )
  copy_id <- copied$id
  message("  Created new copy: ", copy_id)
}

# Get game tabs from the copy
all_tabs  <- googlesheets4::sheet_names(copy_id)
game_tabs <- all_tabs[grepl(GAME_TAB_PATTERN, all_tabs) & !all_tabs %in% PRESERVE_TABS]
message("  Game tabs: ", paste(game_tabs, collapse=", "))

# Detect lineup row positions from first game tab
message("\n── Detecting lineup row positions...")
template_data <- googlesheets4::read_sheet(copy_id, sheet=game_tabs[1],
                                            col_names=FALSE, col_types="c")
col_a <- as.character(template_data[[1]])
hits  <- which(col_a %in% c("1","1.0"))
if (length(hits)>=4) {
  lr <- list(at=hits[1],ht=hits[2],ab=hits[3],hb=hits[4])
} else if (length(hits)>=2) {
  g  <- hits[2]-hits[1]
  lr <- list(at=hits[1],ht=hits[2],ab=hits[1]+g*3L,hb=hits[2]+g*3L)
} else {
  lr <- list(at=3L,ht=16L,ab=43L,hb=56L)
}
message("  TOP: away=",lr$at," home=",lr$ht,
        " | BOT: away=",lr$ab," home=",lr$hb)

# =============================================================================
# SECTION 3: Fetch schedule + lineups
# =============================================================================

message("\n── Fetching schedule: ", SLATE_DATE, "...")
mlb_get <- function(url, tries=3L) {
  for (i in seq_len(tries)) {
    r <- tryCatch(httr::GET(url,httr::timeout(20),httr::user_agent("MLB-Model/1.0")),
                  error=function(e)NULL)
    if (!is.null(r)&&httr::status_code(r)==200L)
      return(jsonlite::fromJSON(httr::content(r,"text",encoding="UTF-8"),flatten=TRUE))
    if (i<tries) Sys.sleep(2)
  }; NULL
}

raw <- mlb_get(paste0(MLB_BASE,"/schedule?sportId=1&date=",SLATE_DATE,
                      "&gameType=R&hydrate=probablePitcher(note),team,venue,status"))
if (is.null(raw)||length(raw$dates)==0) stop("No games for ",SLATE_DATE)
gr <- raw$dates$games[[1]]
message("  ",nrow(gr)," games on slate")

games <- tibble(
  game_pk    = as.integer(gr$gamePk),
  away_abbr  = norm_abbr(gr[["teams.away.team.abbreviation"]]),
  home_abbr  = norm_abbr(gr[["teams.home.team.abbreviation"]]),
  venue      = gr[["venue.name"]],
  away_pid   = as.integer(gr[["teams.away.probablePitcher.id"]]),
  away_pname = if_else(is.na(gr[["teams.away.probablePitcher.fullName"]]),"TBD",
                       gr[["teams.away.probablePitcher.fullName"]]),
  home_pid   = as.integer(gr[["teams.home.probablePitcher.id"]]),
  home_pname = if_else(is.na(gr[["teams.home.probablePitcher.fullName"]]),"TBD",
                       gr[["teams.home.probablePitcher.fullName"]])
)

# Pitcher handedness
pids <- unique(na.omit(c(games$away_pid,games$home_pid)))
thm  <- list()
if (length(pids)>0)
  for (b in split(pids,ceiling(seq_along(pids)/50))) {
    d <- mlb_get(paste0(MLB_BASE,"/people?personIds=",paste(b,collapse=","),"&hydrate=none"))
    if (!is.null(d$people)) for (j in seq_len(nrow(d$people)))
      thm[[as.character(d$people$id[j])]] <- d$people$pitchHand.code[j] %||% "R"
  }
games <- games |> mutate(
  away_throws = map_chr(as.character(away_pid),~thm[[.x]]%||%"R"),
  home_throws = map_chr(as.character(home_pid),~thm[[.x]]%||%"R")
)

message("\n── Fetching lineups (FanGraphs + MLB API fallback)...")
fg_lu <- tryCatch({
  r <- httr::GET(
    paste0("https://www.fangraphs.com/scores?date=",SLATE_DATE),
    httr::timeout(25),
    httr::user_agent("Mozilla/5.0 AppleWebKit/537.36"),
    httr::add_headers(Accept="text/html",`Accept-Language`="en-US",
                      Referer="https://www.fangraphs.com/"))
  if (httr::status_code(r)!=200L) return(list())
  pg <- httr::content(r,"text",encoding="UTF-8")
  m  <- regmatches(pg,regexpr('<script id="__NEXT_DATA__"[^>]+>(.*?)</script>',pg,perl=TRUE))
  if (!length(m)) return(list())
  p  <- jsonlite::fromJSON(gsub('<script[^>]+>|</script>',"",m),simplifyVector=FALSE)
  gf <- p$props$pageProps$dehydratedState$queries[[1]]$state$data
  if (is.null(gf)) return(list())
  ps <- function(pl){
    if (!length(pl)) return(list())
    pl <- pl[order(sapply(pl,function(x)as.integer(x$BatOrder%||%99)))]
    lapply(pl,function(x)list(name=x$PlayerName%||%"?",bats=x$Bats%||%"R",
                               position=x$DisplayPosition%||%"",
                               projected=isTRUE(x$IsProjected)))
  }
  res <- list()
  for (g in gf) {
    s<-g$schedule;lu<-g$lineups;if(is.null(s)) next
    pk<-as.character(s$MLBGameId%||%"");if(!nzchar(pk)) next
    res[[pk]]<-list(away=if(!is.null(lu))ps(lu$lineupAway)else list(),
                    home=if(!is.null(lu))ps(lu$lineupHome)else list())
  }
  message("  FanGraphs: ",length(res)," games"); res
}, error=function(e){message("  FG error: ",e$message);list()})

mlb_lu <- function(pk) {
  d <- mlb_get(paste0(MLB_BASE,"/game/",pk,"/boxscore"))
  if(is.null(d)) return(list(away=list(),home=list()))
  ps <- function(td){
    b<-td$batters;p<-td$players;if(!length(b)) return(list())
    map(seq_along(b),function(i){
      pp<-p[[paste0("ID",b[[i]])]]
      list(name=tryCatch(pp$person$fullName,error=function(e)"?"),
           bats=tryCatch(pp$person$batSide$code,error=function(e)"R"),
           position=tryCatch(pp$position$abbreviation,error=function(e)""))
    })
  }; list(away=ps(d$teams$away),home=ps(d$teams$home))
}

lus <- map(games$game_pk,function(pk){
  fg <- fg_lu[[as.character(pk)]]
  if (!is.null(fg)&&(length(fg$away)||length(fg$home))){
    proj <- any(sapply(c(fg$away,fg$home),function(p)isTRUE(p$projected)))
    return(list(away=fg$away,home=fg$home,src=if(proj)"FG-proj"else"FG-conf"))
  }
  r <- mlb_lu(pk)
  list(away=r$away,home=r$home,src=if(length(r$away)||length(r$home))"MLB"else"none")
})
names(lus) <- as.character(games$game_pk)

for (i in seq_len(nrow(games)))
  message("  ",games$away_abbr[i],"@",games$home_abbr[i],
          " away=",length(lus[[as.character(games$game_pk[i])]]$away),
          " home=",length(lus[[as.character(games$game_pk[i])]]$home),
          " [",lus[[as.character(games$game_pk[i])]]$src,"]")

# =============================================================================
# SECTION 4: Match today's games to sheet tabs + write lineups
# =============================================================================

message("\n── Writing lineups to Google Sheet copy...")

# Build lookup: sheet tab name -> game row
# Try both the exact tab names in the copy and today's game combos
tab_to_game <- list()
for (i in seq_len(nrow(games))) {
  g     <- games[i,]
  sname <- substr(paste0(g$away_abbr,g$home_abbr),1L,31L)
  if (sname %in% game_tabs) {
    tab_to_game[[sname]] <- i
  }
}

# For tabs in the copy that don't match today exactly, rename them
# First handle perfect matches, then handle mismatches
today_names <- substr(paste0(games$away_abbr,games$home_abbr),1L,31L)
matched_tabs <- intersect(game_tabs, today_names)
unmatched_source <- setdiff(game_tabs, today_names)
unmatched_today  <- setdiff(today_names, game_tabs)

message("  Matched tabs: ", paste(matched_tabs, collapse=", "))
if (length(unmatched_today)>0)
  message("  New tabs needed: ", paste(unmatched_today, collapse=", "))

# Rename unmatched source tabs to today's unmatched game names
# (pair them up by schedule order)
if (length(unmatched_source)>0 && length(unmatched_today)>0) {
  n_rename <- min(length(unmatched_source), length(unmatched_today))
  for (k in seq_len(n_rename)) {
    old_name <- unmatched_source[k]
    new_name <- unmatched_today[k]
    tryCatch({
      googlesheets4::sheet_rename(copy_id, sheet=old_name, new_name=new_name)
      message("  Renamed: ", old_name, " → ", new_name)
      Sys.sleep(0.5)  # avoid rate limiting
    }, error=function(e) message("  [WARN] Rename failed: ",e$message))
  }
  # Update game_tabs after renames
  game_tabs <- googlesheets4::sheet_names(copy_id)
  game_tabs <- game_tabs[grepl(GAME_TAB_PATTERN,game_tabs)&!game_tabs%in%PRESERVE_TABS]
}

# Helper: build a data frame for a lineup block (9 rows x 4 cols)
lu_range_data <- function(pl, abbr, n=9) {
  purrr::map_dfr(seq_len(min(length(pl),n)), function(i) {
    p   <- pl[[i]]
    pos <- tolower(p$position%||%"x")
    tibble(
      order    = as.numeric(i),
      position = dplyr::case_when(pos=="c" ~ "c", pos %in% c("dh","d") ~ "dh", TRUE ~ "x"),
      name     = paste(strip_accents(p$name%||%"?"), abbr),
      hand     = p$bats%||%"R"
    )
  })
}

# Write lineups into each game tab
for (i in seq_len(nrow(games))) {
  g     <- games[i,]
  pk    <- as.character(g$game_pk)
  sname <- substr(paste0(g$away_abbr,g$home_abbr),1L,31L)
  lu    <- lus[[pk]]

  if (!sname %in% game_tabs) {
    message("  [SKIP] ",sname," not in sheet"); next
  }

  al <- paste(strip_accents(g$away_pname), g$away_abbr)
  hl <- paste(strip_accents(g$home_pname), g$home_abbr)

  # Helper: write to a specific cell range
  write_cell <- function(row, col, value) {
    range <- paste0(LETTERS[col], row)
    tryCatch(
      googlesheets4::range_write(
        copy_id, data=tibble(x=value),
        sheet=sname, range=range, col_names=FALSE
      ),
      error=function(e) message("    [warn] ",range,": ",e$message)
    )
  }

  # Pitcher header rows
  write_cell(lr$at-1L, 3L, al); write_cell(lr$at-1L, 4L, g$away_throws)
  write_cell(lr$ht-1L, 3L, hl); write_cell(lr$ht-1L, 4L, g$home_throws)
  write_cell(lr$ab-1L, 3L, al); write_cell(lr$ab-1L, 4L, g$away_throws)
  write_cell(lr$hb-1L, 3L, hl); write_cell(lr$hb-1L, 4L, g$home_throws)

  # Write lineups as blocks (4 cols wide)
  write_block <- function(pl, abbr, start_row) {
    if (!length(pl)) return(invisible(NULL))
    df    <- lu_range_data(pl, abbr)
    range <- paste0("A",start_row,":D",start_row+nrow(df)-1L)
    tryCatch(
      googlesheets4::range_write(
        copy_id, data=df, sheet=sname,
        range=range, col_names=FALSE
      ),
      error=function(e) message("    [warn] block write: ",e$message)
    )
    Sys.sleep(0.3)  # respect API rate limits
  }

  write_block(lu$away, g$away_abbr, lr$at)
  write_block(lu$home, g$home_abbr, lr$ht)

  # Bottom block: mirror top block using USER_ENTERED input so formulas evaluate
  write_mirror_block_formula <- function(top_start, bot_start, n=9) {
    # Build a list of cell:formula pairs and write via the Sheets API
    # valueInputOption="USER_ENTERED" makes Google Sheets parse =C3 as a formula
    for (i in seq_len(n)) {
      top_row <- top_start + i - 1L
      bot_row <- bot_start + i - 1L
      row_formulas <- data.frame(
        A = paste0("=A", top_row),
        B = paste0("=B", top_row),
        C = paste0("=C", top_row),
        D = paste0("=D", top_row)
      )
      tryCatch(
        googlesheets4::range_write(
          ss        = copy_id,
          data      = row_formulas,
          sheet     = sname,
          range     = paste0("A", bot_row),
          col_names = FALSE,
          reformat  = FALSE
        ),
        error = function(e) NULL
      )
    }
    Sys.sleep(0.5)
  }

  # For formulas to be interpreted, use sheets_edit with USER_ENTERED
  # googlesheets4 range_write uses RAW by default — we need a workaround:
  # Write the formula as a named range using the low-level API
  write_formula_cell <- function(row, col_letter, formula) {
    req <- list(
      spreadsheetId = copy_id,
      range         = paste0("'", sname, "'!", col_letter, row),
      valueInputOption = "USER_ENTERED",
      resource      = list(values = list(list(formula)))
    )
    tryCatch(
      googlesheets4::request_generate("sheets.spreadsheets.values.update",
        params = req) |> googlesheets4::request_make(),
      error = function(e) NULL
    )
  }

  # Mirror pitcher headers
  write_formula_cell(lr$ab-1L, "C", paste0("=C", lr$at-1L))
  write_formula_cell(lr$ab-1L, "D", paste0("=D", lr$at-1L))
  write_formula_cell(lr$hb-1L, "C", paste0("=C", lr$ht-1L))
  write_formula_cell(lr$hb-1L, "D", paste0("=D", lr$ht-1L))

  # Mirror batter rows
  for (i in seq_len(9)) {
    for (col_letter in c("A","B","C","D")) {
      write_formula_cell(lr$ab+i-1L, col_letter, paste0("=",col_letter, lr$at+i-1L))
      write_formula_cell(lr$hb+i-1L, col_letter, paste0("=",col_letter, lr$ht+i-1L))
    }
    Sys.sleep(0.1)
  }

  message("  ✓  ",sname,"  (",strip_accents(g$away_pname)," vs ",strip_accents(g$home_pname),")")
}

message("\n  Google Sheet updated: https://docs.google.com/spreadsheets/d/",copy_id)
message("  Waiting 10s for Sheets to finish recalculating...")
Sys.sleep(10)

# =============================================================================
# SECTION 5: Read back recalculated results + build gt table
# =============================================================================

message("\n── Reading recalculated results...")

safe_num <- function(x) suppressWarnings(as.numeric(gsub("%","",as.character(x))))

# Refresh game_tabs after all renames
game_tabs <- googlesheets4::sheet_names(copy_id)
game_tabs <- game_tabs[grepl(GAME_TAB_PATTERN,game_tabs)&!game_tabs%in%PRESERVE_TABS]
today_tabs <- substr(paste0(games$away_abbr,games$home_abbr),1L,31L)
read_tabs  <- intersect(today_tabs, game_tabs)

model_data <- purrr::map_dfr(read_tabs, function(tab) {
  df <- tryCatch(
    googlesheets4::read_sheet(copy_id, sheet=tab,
                              col_names=FALSE, col_types="c"),
    error=function(e){message("  [WARN] ",tab,": ",e$message);NULL}
  )
  if (is.null(df)||nrow(df)<5) return(NULL)

  col_b <- as.character(df[[2]])
  col_c <- as.character(df[[3]])

  # Pitcher names
  pitcher_rows <- which(col_b=="Pitcher")
  away_p <- if(length(pitcher_rows)>=1)
    str_trim(gsub("\\s+[A-Z]{2,3}$","",as.character(df[[pitcher_rows[1],3]])))
  else "TBD"
  home_p <- if(length(pitcher_rows)>=2)
    str_trim(gsub("\\s+[A-Z]{2,3}$","",as.character(df[[pitcher_rows[2],3]])))
  else "TBD"

  # Summary rows — "Team" / "Replacement Wins" block
  sr <- which(col_b=="Team" & col_c=="Replacement Wins")[1]
  if (is.na(sr)) return(NULL)

  # Total block
  tr <- which(col_b=="Team" & col_c=="RS 162")[1]

  mid <- nchar(tab) %/% 2 + nchar(tab) %% 2
  tibble(
    tab          = tab,
    matchup      = paste0(substr(tab,1,mid)," @ ",substr(tab,mid+1,nchar(tab))),
    away_pitcher = away_p,
    home_pitcher = home_p,
    away_win_pct = safe_num(df[[sr+1,10]]),
    home_win_pct = safe_num(df[[sr+2,10]]),
    away_xwins   = safe_num(df[[sr+1,6]]),
    home_xwins   = safe_num(df[[sr+2,6]]),
    book_ml_away = safe_num(df[[sr+1,12]]),
    book_ml_home = safe_num(df[[sr+2,12]]),
    model_total  = if(!is.na(tr)) safe_num(df[[tr+1,7]]) else NA_real_,
    model_total2 = if(!is.na(tr)) safe_num(df[[tr+1,8]]) else NA_real_,
    over_ml      = if(!is.na(tr)) safe_num(df[[tr+1,14]]) else NA_real_,
    under_ml     = if(!is.na(tr)) safe_num(df[[tr+1,13]]) else NA_real_
  ) |> mutate(
    away_win_pct = if_else(away_win_pct>1, away_win_pct/100, away_win_pct),
    home_win_pct = if_else(home_win_pct>1, home_win_pct/100, home_win_pct),
    away_fair_ml = win_to_ml(away_win_pct),
    home_fair_ml = win_to_ml(home_win_pct)
  )
})

message("  Read ", nrow(model_data), " games")
for (i in seq_len(nrow(model_data))) {
  m <- model_data[i,]
  message("    ",m$tab,": away=",round(m$away_win_pct*100,1),
          "% (",fmt_ml(m$away_fair_ml),
          ") home=",round(m$home_win_pct*100,1),
          "% (",fmt_ml(m$home_fair_ml),
          ") total=",round(m$model_total,1))
}

# =============================================================================
# SECTION 6: gt summary table
# =============================================================================

message("\n── Building gt table...")

model_data |>
  arrange(desc(away_win_pct)) |>
  mutate(
    pitchers  = paste0(away_pitcher,"\n",home_pitcher),
    away_pct  = paste0(round(away_win_pct*100,1),"%"),
    home_pct  = paste0(round(home_win_pct*100,1),"%"),
    away_ml   = fmt_ml(away_fair_ml),
    home_ml   = fmt_ml(home_fair_ml),
    book_away = fmt_ml(book_ml_away),
    book_home = fmt_ml(book_ml_home),
    total_fmt = ifelse(is.na(model_total),"—",
                  paste0(round(model_total,1)," / ",round(model_total2,1))),
    over_fmt  = fmt_ml(over_ml),
    under_fmt = fmt_ml(under_ml)
  ) |>
  select(matchup,pitchers,away_pct,away_ml,book_away,
         home_pct,home_ml,book_home,total_fmt,over_fmt,under_fmt,
         away_xwins,home_xwins,away_win_pct,home_win_pct) |>
  gt() |>
  tab_header(
    title    = md(paste0("**MLB Model — ",format(as.Date(SLATE_DATE),"%B %d, %Y"),"**")),
    subtitle = md(paste0("*",nrow(model_data)," games · model fair lines vs book*"))
  ) |>
  cols_label(
    matchup="GAME", pitchers="PITCHERS",
    away_pct="WIN%", away_ml="FAIR", book_away="BOOK",
    home_pct="WIN%", home_ml="FAIR", book_home="BOOK",
    total_fmt="TOTAL", over_fmt="OVER", under_fmt="UNDER",
    away_xwins="xW", home_xwins="xW",
    away_win_pct="", home_win_pct=""
  ) |>
  cols_hide(columns=c(away_win_pct,home_win_pct)) |>
  tab_spanner(label="AWAY",  columns=c(away_pct,away_ml,book_away)) |>
  tab_spanner(label="HOME",  columns=c(home_pct,home_ml,book_home)) |>
  tab_spanner(label="TOTAL", columns=c(total_fmt,over_fmt,under_fmt)) |>
  tab_spanner(label="xWINS", columns=c(away_xwins,home_xwins)) |>
  fmt_number(columns=c(away_xwins,home_xwins),decimals=1) |>
  tab_style(style=list(cell_fill(color="#ecfdf3"),cell_text(weight="bold",color="#0a7c3e")),
            locations=cells_body(columns=away_pct,rows=away_win_pct>=0.55)) |>
  tab_style(style=list(cell_fill(color="#ecfdf3"),cell_text(weight="bold",color="#0a7c3e")),
            locations=cells_body(columns=home_pct,rows=home_win_pct>=0.55)) |>
  tab_style(style=list(cell_fill(color="#fff8e6"),cell_text(weight="bold",color="#92600a")),
            locations=cells_body(columns=away_pct,rows=away_win_pct>=0.50&away_win_pct<0.55)) |>
  tab_style(style=list(cell_fill(color="#fff8e6"),cell_text(weight="bold",color="#92600a")),
            locations=cells_body(columns=home_pct,rows=home_win_pct>=0.50&home_win_pct<0.55)) |>
  tab_style(style=cell_text(size=px(12),weight="bold"),
            locations=cells_body(columns=matchup)) |>
  tab_style(style=cell_text(size=px(10),color="gray45"),
            locations=cells_body(columns=pitchers)) |>
  tab_style(style=cell_fill(color="#f7f7f5"),
            locations=cells_body(rows=seq(2,nrow(model_data),2))) |>
  tab_style(style=cell_text(weight="bold",size=px(11)),
            locations=cells_column_spanners()) |>
  tab_style(style=cell_text(size=px(10),color="gray40"),
            locations=cells_column_labels()) |>
  tab_style(style=cell_borders(sides="bottom",color="#e2e2dc",weight=px(1)),
            locations=cells_body()) |>
  cols_width(
    matchup~px(105), pitchers~px(160),
    away_pct~px(52), away_ml~px(58), book_away~px(58),
    home_pct~px(52), home_ml~px(58), book_home~px(58),
    total_fmt~px(80), over_fmt~px(52), under_fmt~px(52),
    away_xwins~px(44), home_xwins~px(44)
  ) |>
  tab_options(
    table.font.names="IBM Plex Sans, sans-serif",
    table.font.size=px(12),
    table.border.top.style="hidden",
    table.border.bottom.style="hidden",
    column_labels.border.bottom.width=px(1),
    column_labels.border.bottom.color="#0a0a08",
    heading.border.bottom.style="hidden",
    table.width=pct(100),
    data_row.padding=px(7)
  )

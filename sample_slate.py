"""
sample_slate.py
===============
Synthetic results that match the shape of mlb_f5_model.project_slate output,
so export_json.py --sample can produce a representative docs/data.json without
scraping. Used for building/testing the website offline.
"""
from dataclasses import dataclass, field

SAMPLE_DATE = "2026-05-29"
SAMPLE_TIMES = {
    "PITBOS": "7:10 PM ET", "TBADET": "6:40 PM ET", "SLNMIL": "8:15 PM ET",
    "NYAKCA": "8:10 PM ET", "ARISFN": "9:45 PM ET", "HOUTEX": "8:05 PM ET",
    "LANCOL": "8:40 PM ET",
}


@dataclass
class _Slot:
    order: int
    position: str
    name: str


@dataclass
class _Lineup:
    pitcher: object
    batters: list = field(default_factory=list)


def _lu(names, poss=None):
    poss = poss or ["x"] * 9
    return _Lineup(pitcher=None, batters=[
        _Slot(i + 1, poss[i], names[i]) for i in range(9)])


def sample_results():
    G = []

    G.append(dict(
        tab="PIT@BOS", status="ok",
        away_team="PIT", home_team="BOS",
        away_sp="Paul Skenes", home_sp="Garrett Crochet",
        away_throws="R", home_throws="L",
        away_pitcher_era=2.83, home_pitcher_era=2.89,
        away_wpct=0.471, home_wpct=0.529,
        away_fair=112, home_fair=-119,
        away_book=128, home_book=-142,
        away_edge=0.061, home_edge=-0.061,
        away_ev=0.071, home_ev=-0.052,
        away_off=2.41, home_off=2.66,
        away_total_era=3.05, home_total_era=3.12,
        predicted_total=4.92, totals_book_line=5.0,
        total_over_odds=104, total_under_odds=-122,
        park_factor=1.02,
        league_avg_off=4.857, league_avg_era=4.17,
        _away_lineup=_lu(["Oneil Cruz", "Bryan Reynolds", "Andrew McCutchen",
                          "Ke'Bryan Hayes", "Joey Bart", "Nick Gonzales",
                          "Spencer Horwitz", "Henry Davis", "Jared Triolo"],
                         ["x", "x", "dh", "x", "c", "x", "x", "x", "x"]),
        _home_lineup=_lu(["Jarren Duran", "Rafael Devers", "Trevor Story",
                          "Wilyer Abreu", "Triston Casas", "Connor Wong",
                          "Ceddanne Rafaela", "David Hamilton", "Rob Refsnyder"],
                         ["x", "dh", "x", "x", "x", "c", "x", "x", "x"]),
    ))

    G.append(dict(
        tab="TBA@DET", status="ok",
        away_team="TBA", home_team="DET",
        away_sp="Shane Baz", home_sp="Tarik Skubal",
        away_throws="R", home_throws="L",
        away_pitcher_era=3.80, home_pitcher_era=2.76,
        away_wpct=0.402, home_wpct=0.598,
        away_fair=149, home_fair=-167,
        away_book=152, home_book=-170,
        away_edge=0.005, home_edge=-0.005,
        away_ev=0.006, home_ev=-0.004,
        away_off=2.30, home_off=2.58,
        away_total_era=3.40, home_total_era=2.95,
        predicted_total=4.55, totals_book_line=4.5,
        total_over_odds=-108, total_under_odds=-104,
        park_factor=0.97,
        league_avg_off=4.857, league_avg_era=4.17,
        _away_lineup=_lu(["Chandler Simpson", "Junior Caminero", "Jonathan Aranda",
                          "Yandy Diaz", "Brandon Lowe", "Josh Lowe",
                          "Christopher Morel", "Ben Rortvedt", "Taylor Walls"],
                         ["x", "x", "x", "dh", "x", "x", "x", "c", "x"]),
        _home_lineup=_lu(["Riley Greene", "Kerry Carpenter", "Spencer Torkelson",
                          "Colt Keith", "Matt Vierling", "Jake Rogers",
                          "Trey Sweeney", "Zach McKinstry", "Parker Meadows"],
                         ["x", "dh", "x", "x", "x", "c", "x", "x", "x"]),
    ))

    G.append(dict(
        tab="SLN@MIL", status="ok",
        away_team="SLN", home_team="MIL",
        away_sp="Sonny Gray", home_sp="Freddy Peralta",
        away_throws="R", home_throws="R",
        away_pitcher_era=3.55, home_pitcher_era=3.41,
        away_wpct=0.512, home_wpct=0.488,
        away_fair=-105, home_fair=105,
        away_book=118, home_book=-132,
        away_edge=0.082, home_edge=-0.082,
        away_ev=0.104, home_ev=-0.068,
        away_off=2.52, home_off=2.49,
        away_total_era=3.20, home_total_era=3.10,
        predicted_total=4.78, totals_book_line=4.5,
        total_over_odds=-130, total_under_odds=110,
        park_factor=0.94,
        league_avg_off=4.857, league_avg_era=4.17,
        _away_lineup=_lu(["Masyn Winn", "Alec Burleson", "Willson Contreras",
                          "Nolan Arenado", "Ivan Herrera", "Lars Nootbaar",
                          "Jordan Walker", "Nolan Gorman", "Victor Scott II"],
                         ["x", "x", "dh", "x", "c", "x", "x", "x", "x"]),
        _home_lineup=_lu(["Jackson Chourio", "William Contreras", "Christian Yelich",
                          "Rhys Hoskins", "Brice Turang", "Sal Frelick",
                          "Joey Ortiz", "Garrett Mitchell", "Andruw Monasterio"],
                         ["x", "c", "dh", "x", "x", "x", "x", "x", "x"]),
    ))

    G.append(dict(
        tab="NYA@KCA", status="ok",
        away_team="NYA", home_team="KCA",
        away_sp="Carlos Rodon", home_sp="Cole Ragans",
        away_throws="L", home_throws="L",
        away_pitcher_era=3.30, home_pitcher_era=3.15,
        away_wpct=0.546, home_wpct=0.454,
        away_fair=-120, home_fair=120,
        away_book=-115, home_book=105,
        away_edge=0.022, home_edge=-0.022,
        away_ev=0.026, home_ev=-0.019,
        away_off=2.61, home_off=2.44,
        away_total_era=3.05, home_total_era=2.98,
        predicted_total=4.61, totals_book_line=4.5,
        total_over_odds=-115, total_under_odds=-105,
        park_factor=1.00,
        league_avg_off=4.857, league_avg_era=4.17,
        _away_lineup=_lu(["Gleyber Torres", "Juan Soto", "Aaron Judge",
                          "Giancarlo Stanton", "Anthony Rizzo", "Jose Trevino",
                          "Anthony Volpe", "Alex Verdugo", "Oswaldo Cabrera"],
                         ["x", "x", "x", "dh", "x", "c", "x", "x", "x"]),
        _home_lineup=_lu(["Bobby Witt Jr.", "Vinnie Pasquantino", "Salvador Perez",
                          "MJ Melendez", "Hunter Renfroe", "Maikel Garcia",
                          "Michael Massey", "Kyle Isbel", "Dairon Blanco"],
                         ["x", "x", "dh", "x", "x", "x", "x", "x", "c"]),
    ))

    G.append(dict(
        tab="ARI@SFN", status="ok",
        away_team="ARI", home_team="SFN",
        away_sp="Zac Gallen", home_sp="Logan Webb",
        away_throws="R", home_throws="R",
        away_pitcher_era=3.65, home_pitcher_era=3.25,
        away_wpct=0.486, home_wpct=0.514,
        away_fair=106, home_fair=-112,
        away_book=110, home_book=-124,
        away_edge=0.018, home_edge=-0.018,
        away_ev=0.020, home_ev=-0.015,
        away_off=2.55, home_off=2.39,
        away_total_era=3.30, home_total_era=3.00,
        predicted_total=4.70, totals_book_line=5.0,
        total_over_odds=118, total_under_odds=-138,
        park_factor=0.92,
        league_avg_off=4.857, league_avg_era=4.17,
        _away_lineup=_lu(["Ketel Marte", "Corbin Carroll", "Eugenio Suarez",
                          "Christian Walker", "Lourdes Gurriel Jr.", "Joc Pederson",
                          "Gabriel Moreno", "Jake McCarthy", "Geraldo Perdomo"],
                         ["x", "x", "x", "x", "x", "dh", "c", "x", "x"]),
        _home_lineup=_lu(["LaMonte Wade Jr.", "Matt Chapman", "Heliot Ramos",
                          "Jung Hoo Lee", "Michael Conforto", "Mike Yastrzemski",
                          "Patrick Bailey", "Tyler Fitzgerald", "Brett Wisely"],
                         ["x", "x", "x", "dh", "x", "x", "c", "x", "x"]),
    ))

    # A game still waiting on lineups — should be skipped by the serializer.
    G.append(dict(tab="HOU@TEX", status="no_lineup_yet",
                  away_sp=None, home_sp=None))

    return G

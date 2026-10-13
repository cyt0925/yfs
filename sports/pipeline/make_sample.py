#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""產生本機測試用的範例資料 sports/data/sample.json（不連網）。
網頁加上 ?data=sample 就會讀這份。"""
import json
import os
import sys
from datetime import datetime, timedelta, timezone

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import fetch as F  # noqa: E402

TW = F.TW
now = datetime.now(timezone.utc)


def team(lg, abbr, name, short, tid, record, form, hot=None, injuries=None, color=None):
    logo_lg = {"nba": "nba", "nhl": "nhl", "eng.1": "soccer", "nfl": "nfl"}[lg]
    t = dict(id=tid, abbr=abbr, name=name, short=short, name_zh=F.zh_name(lg, abbr, name, short),
             logo=f"https://a.espncdn.com/i/teamlogos/{logo_lg}/500/{abbr.lower()}.png" if logo_lg != "soccer" else f"https://a.espncdn.com/i/teamlogos/soccer/500/{tid}.png",
             color=color, record=record, record_home=None, record_away=None, score=None, winner=False,
             form_str=None, leaders=[], injuries=injuries or [], key_out=[], key_maybe=[], out_count=0, maybe_count=0,
             hot=hot or [], form=[], form_note=None, streak=None, last10=None, last10_w=0,
             avg_for5=None, avg_against5=None, avg_total5=None, avg_margin5=None)
    rows = []
    base = now - timedelta(days=len(form) * 2)
    for i, (res, pf, pa, opp) in enumerate(form):
        d = base + timedelta(days=i * 2)
        rows.append(dict(date=d.astimezone(TW).strftime("%m/%d"), opp=opp, opp_zh=F.zh_name(lg, opp, opp, opp), opp_logo=None,
                         home=(i % 2 == 0), pf=pf, pa=pa, result=res))
    t["form"] = rows
    w = sum(1 for r in rows if r["result"] == "W"); l = sum(1 for r in rows if r["result"] == "L"); dr = sum(1 for r in rows if r["result"] == "D")
    t["last10"] = f"{w}勝{l}敗" + (f"{dr}和" if dr else "")
    t["last10_w"] = w
    s = 0; first = rows[-1]["result"]
    for r in reversed(rows):
        if r["result"] == first: s += 1
        else: break
    t["streak"] = {"W": "連勝", "L": "連敗", "D": "連和"}[first] + str(s) if s >= 2 else None
    last5 = rows[-5:]
    t["avg_for5"] = round(sum(r["pf"] for r in last5) / 5, 1)
    t["avg_against5"] = round(sum(r["pa"] for r in last5) / 5, 1)
    t["avg_total5"] = round(t["avg_for5"] + t["avg_against5"], 1)
    t["avg_margin5"] = round(t["avg_for5"] - t["avg_against5"], 1)
    t["out_count"] = sum(1 for x in t["injuries"] if x["severity"] == "out")
    t["maybe_count"] = sum(1 for x in t["injuries"] if x["severity"] == "maybe")
    return t


def inj(pid, player, pos, status, detail, lg="nba", ret=None):
    return dict(id=pid, player=player, headshot=f"https://a.espncdn.com/i/headshots/{lg}/players/full/{pid}.png",
                position=pos, position_zh=F.POS_ZH.get(pos), status=status, status_zh=F.STATUS_ZH.get(status, status),
                severity="out" if status in F.OUT_LIKE else ("maybe" if status in F.MAYBE_LIKE else "other"),
                detail=detail, side=None, return_date=ret, comment=None, date=None)


def hot(pid, player, pos, stat, last5, season, games, lg="nba"):
    delta = round((last5 - season) / season * 100)
    label = "火燙" if delta >= 15 else ("冰冷" if delta <= -15 else "平穩")
    return dict(id=pid, player=player, headshot=f"https://a.espncdn.com/i/headshots/{lg}/players/full/{pid}.png", position=pos,
                stat=stat, season_name="2026-27 例行賽", last5=last5, season=season, delta_pct=delta, label=label,
                games=[dict(date=f"10/{d}", opp=o, v=v) for d, o, v in games], n_games=12)


def odds(ml_h, ml_a, ml_d=None, spread=None, total=None, details=None):
    probs = {k: F.implied(v) for k, v in (("home", ml_h), ("away", ml_a), ("draw", ml_d)) if v is not None}
    s = sum(probs.values())
    fair = {k: round(v / s, 3) for k, v in probs.items()}
    return dict(provider="ESPN BET", details=details, spread=spread, total=total, ml_home=ml_h, ml_away=ml_a, ml_draw=ml_d,
                dec_home=F.dec(ml_h), dec_away=F.dec(ml_a), dec_draw=F.dec(ml_d),
                prob_home=fair.get("home"), prob_away=fair.get("away"), prob_draw=fair.get("draw"),
                fav="home" if ml_h is not None and ml_a is not None and ml_h < ml_a else "away")


def game(gid, lg, sport, hours_from_now, home, away, o, state="pre", scores=None, detail=None, news=None, venue=None, bc=None, headline=None):
    start = now + timedelta(hours=hours_from_now)
    stw = start.astimezone(TW)
    st = dict(state=state, completed=(state == "post"), detail=detail or ("未開賽" if state == "pre" else "已結束"), clock=None, period=None)
    if scores:
        home["score"], away["score"] = scores
    g = dict(id=gid, league=lg, league_zh=F.LEAGUE_ZH[lg], sport=sport, tennis=False, tournament=None, round=None,
             name=f"{away['abbr']} @ {home['abbr']}", start=F.iso(start), start_tw=F.iso(stw), day_tw=stw.strftime("%Y-%m-%d"),
             time_tw=stw.strftime("%H:%M"), status=st, venue=venue, broadcast=bc, home=home, away=away, odds=o,
             headline=headline, news=news or [], suggestion=None)
    for side in ("home", "away"):
        t = g[side]
        leaders = {p["player"] for p in t["hot"]}
        t["key_out"] = [x["player"] for x in t["injuries"] if x["severity"] == "out" and x["player"] in leaders]
        t["key_maybe"] = [x["player"] for x in t["injuries"] if x["severity"] == "maybe" and x["player"] in leaders]
    g["suggestion"] = F.build_suggestion(g, F.LEAGUES_BY_KEY[lg])
    return g


def main():
    lal = team("nba", "LAL", "Los Angeles Lakers", "Lakers", "13", "4-1",
               [("W", 118, 109, "PHX"), ("W", 112, 104, "SAC"), ("L", 101, 115, "DEN"), ("W", 124, 117, "MIN"), ("W", 110, 102, "POR"), ("W", 121, 113, "GS")],
               hot=[hot("1966", "LeBron James", "SF", "PTS", 28.6, 24.1, [(3, "PHX", 27), (5, "SAC", 31), (7, "DEN", 24), (9, "MIN", 33), (11, "POR", 28)]),
                    hot("3136195", "Austin Reaves", "SG", "PTS", 15.2, 19.8, [(3, "PHX", 14), (5, "SAC", 18), (7, "DEN", 12), (9, "MIN", 16), (11, "POR", 16)])],
               injuries=[inj("6583", "Anthony Davis", "PF", "Out", "膝蓋", ret="2026-10-20"), inj("4066648", "Jarred Vanderbilt", "PF", "Questionable", "腳踝")])
    bos = team("nba", "BOS", "Boston Celtics", "Celtics", "2", "3-2",
               [("W", 115, 108, "NY"), ("L", 99, 105, "MIA"), ("W", 120, 110, "TOR"), ("L", 104, 112, "PHI"), ("W", 118, 101, "CHA"), ("L", 108, 116, "CLE")],
               hot=[hot("4065648", "Jayson Tatum", "SF", "PTS", 33.4, 27.2, [(2, "NY", 35), (4, "MIA", 29), (6, "TOR", 38), (8, "PHI", 31), (10, "CHA", 34)])],
               injuries=[inj("3917376", "Jaylen Brown", "SG", "Day-To-Day", "腿後肌")])
    gsw = team("nba", "GS", "Golden State Warriors", "Warriors", "9", "2-3",
               [("L", 104, 118, "LAL"), ("W", 122, 108, "UTAH"), ("L", 100, 109, "OKC"), ("L", 111, 120, "DEN"), ("W", 115, 110, "SAC"), ("L", 98, 112, "HOU")],
               hot=[hot("3975", "Stephen Curry", "PG", "PTS", 21.0, 26.5, [(2, "LAL", 18), (4, "UTAH", 25), (6, "OKC", 19), (8, "DEN", 22), (10, "SAC", 21)])],
               injuries=[inj("3975", "Stephen Curry", "PG", "Out", "腳踝"), inj("6589", "Draymond Green", "PF", "Out", "禁賽")])
    okc = team("nba", "OKC", "Oklahoma City Thunder", "Thunder", "25", "5-0",
               [("W", 120, 100, "DAL"), ("W", 115, 103, "GS"), ("W", 128, 110, "HOU"), ("W", 112, 106, "MEM"), ("W", 119, 104, "SA"), ("W", 123, 109, "POR")],
               hot=[hot("4278073", "Shai Gilgeous-Alexander", "PG", "PTS", 34.8, 30.1, [(1, "DAL", 36), (3, "GS", 31), (5, "HOU", 40), (7, "MEM", 33), (9, "SA", 34)])],
               injuries=[])

    mci = team("eng.1", "MCI", "Manchester City", "Man City", "382", "5-2-1",
               [("W", 3, 0, "BUR"), ("W", 2, 1, "ARS"), ("D", 1, 1, "LIV"), ("W", 4, 0, "FUL"), ("W", 2, 0, "EVE"), ("W", 3, 1, "BHA")],
               injuries=[inj("166676", "Rodri", "M", "Out", "膝蓋", lg="soccer")])
    ars = team("eng.1", "ARS", "Arsenal", "Arsenal", "359", "6-1-1",
               [("W", 2, 0, "TOT"), ("W", 1, 0, "NEW"), ("W", 3, 1, "CHE"), ("W", 2, 1, "AVL"), ("L", 1, 2, "MCI"), ("W", 2, 0, "WHU")],
               injuries=[inj("250586", "Bukayo Saka", "FW", "Doubtful", "腿後肌", lg="soccer")])
    liv = team("eng.1", "LIV", "Liverpool", "Liverpool", "364", "5-1-2",
               [("W", 2, 1, "BOU"), ("D", 1, 1, "MCI"), ("L", 0, 2, "CHE"), ("W", 3, 0, "EVE"), ("W", 2, 0, "CRY"), ("L", 1, 2, "MUN")])
    tot = team("eng.1", "TOT", "Tottenham Hotspur", "Tottenham", "367", "3-2-3",
               [("L", 0, 2, "ARS"), ("W", 2, 1, "BRE"), ("D", 2, 2, "WOL"), ("L", 1, 3, "AVL"), ("W", 3, 0, "LEE"), ("D", 1, 1, "NFO")])

    edm = team("nhl", "EDM", "Edmonton Oilers", "Oilers", "6", "3-1-1",
               [("W", 4, 2, "CGY"), ("W", 5, 3, "VAN"), ("L", 2, 4, "COL"), ("W", 3, 1, "SEA"), ("W", 6, 2, "SJ")],
               hot=[hot("3114778", "Connor McDavid", "C", "PTS", 2.4, 1.6, [(5, "CGY", 3), (7, "VAN", 2), (9, "COL", 1), (11, "SEA", 3), (13, "SJ", 3)], lg="nhl")])
    tor = team("nhl", "TOR", "Toronto Maple Leafs", "Maple Leafs", "21", "2-3-0",
               [("L", 1, 3, "MTL"), ("W", 4, 1, "OTT"), ("L", 2, 5, "DET"), ("L", 2, 3, "BUF"), ("W", 3, 2, "BOS")],
               injuries=[inj("4024123", "Auston Matthews", "C", "Day-To-Day", "上半身", lg="nhl")])

    news_lal = [dict(title="Anthony Davis 膝傷持續缺席，湖人禁區輪替吃緊", desc="", url="https://www.espn.com/", image="https://a.espncdn.com/photo/2024/0109/r1274140_1296x729_16-9.jpg", published="2026-10-13T02:10:00Z", team="home"),
                dict(title="LeBron 近五場場均 28.6 分，湖人四連勝", desc="", url="https://www.espn.com/", image=None, published="2026-10-12T20:00:00Z", team="home"),
                dict(title="Jaylen Brown 腿後肌不適，賽前決定是否出賽", desc="", url="https://www.espn.com/", image=None, published="2026-10-12T18:30:00Z", team="away")]

    games = [
        game("g1", "nba", "basketball", 6.5, lal, bos, odds(-165, 140, spread=-3.5, total=224.5, details="LAL -3.5"), news=news_lal, venue="Crypto.com Arena", bc="ESPN"),
        game("g2", "nba", "basketball", 9, okc, gsw, odds(-420, 330, spread=-9.5, total=226.0, details="OKC -9.5"), venue="Paycom Center", bc="TNT"),
        game("g3", "nba", "basketball", -1.2, bos, okc, odds(120, -140, spread=2.5, total=228.5, details="OKC -2.5"), state="in", scores=("78", "81"), detail="第 3 節 05:12"),
        game("g4", "eng.1", "soccer", 27, mci, ars, odds(105, 240, ml_d=250, spread=-0.5, total=2.5, details="MCI -0.5"), venue="Etihad Stadium"),
        game("g5", "eng.1", "soccer", 29.5, liv, tot, odds(-160, 400, ml_d=300, total=3.0, details="LIV -1"), venue="Anfield"),
        game("g6", "nhl", "hockey", 11, edm, tor, odds(-150, 125, spread=-1.5, total=6.5, details="EDM -1.5"), venue="Rogers Place"),
        game("g7", "nba", "basketball", -7, gsw, lal, odds(150, -180), state="post", scores=("104", "118"), detail="已結束"),
    ]
    games.sort(key=lambda g: g["start"])
    # 範例：第一場掛上台灣運彩盤口，看介面
    g1 = next(g for g in games if g["id"] == "g1")
    g1["tw"] = dict(source="玩運彩", url="https://www.playsport.cc/predict/scale?allianceid=3", updated=now.astimezone(TW).isoformat(timespec="minutes"), no="327",
                    spread=dict(line="湖人 -3.5", line_home=-3.5, home=1.75, away=1.68), ml=dict(home=1.5, away=1.96, draw=None),
                    total=dict(line=228.5, over=1.7, under=1.73), intl=None)
    g1["suggestion"] = F.build_suggestion(g1, F.LEAGUES_BY_KEY["nba"])
    feed, hotp = F.aggregate(games)
    standings = {
        "nba": dict(label="NBA", sport="basketball", groups=[dict(name="西區", rows=[
            dict(pos=1, team_id="25", team="Oklahoma City Thunder", abbr="OKC", name_zh="雷霆", logo="https://a.espncdn.com/i/teamlogos/nba/500/okc.png", w=5, l=0, pct=".1000", streak="W5"),
            dict(pos=2, team_id="13", team="Los Angeles Lakers", abbr="LAL", name_zh="湖人", logo="https://a.espncdn.com/i/teamlogos/nba/500/lal.png", w=4, l=1, pct=".800", streak="W4"),
            dict(pos=3, team_id="9", team="Golden State Warriors", abbr="GS", name_zh="勇士", logo="https://a.espncdn.com/i/teamlogos/nba/500/gs.png", w=2, l=3, pct=".400", streak="L1"),
        ]), dict(name="東區", rows=[
            dict(pos=1, team_id="2", team="Boston Celtics", abbr="BOS", name_zh="塞爾提克", logo="https://a.espncdn.com/i/teamlogos/nba/500/bos.png", w=3, l=2, pct=".600", streak="L1"),
        ])]),
        "eng.1": dict(label="英超", sport="soccer", groups=[dict(name="", rows=[
            dict(pos=1, team_id="359", team="Arsenal", abbr="ARS", name_zh="兵工廠", logo="https://a.espncdn.com/i/teamlogos/soccer/500/359.png", gp=8, w=6, d=1, l=1, points=19),
            dict(pos=2, team_id="382", team="Manchester City", abbr="MCI", name_zh="曼城", logo="https://a.espncdn.com/i/teamlogos/soccer/500/382.png", gp=8, w=5, d=2, l=1, points=17),
            dict(pos=3, team_id="364", team="Liverpool", abbr="LIV", name_zh="利物浦", logo="https://a.espncdn.com/i/teamlogos/soccer/500/364.png", gp=8, w=5, d=1, l=2, points=16),
        ])]),
    }
    out = dict(generated_at=now.astimezone(TW).isoformat(timespec="minutes"), generated_ts=int(now.timestamp()),
               today_tw=now.astimezone(TW).strftime("%Y-%m-%d"), sports=F.SPORTS,
               leagues={k: F.LEAGUE_ZH.get(k, k) for k in F.LEAGUES_BY_KEY}, games=games, injury_feed=feed, hot_players=hotp,
               standings=standings, source="範例資料（非即時）", errors=[], calls=0)
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "data", "sample.json")
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=1)
    print("寫入", path, len(games), "場")


if __name__ == "__main__":
    main()

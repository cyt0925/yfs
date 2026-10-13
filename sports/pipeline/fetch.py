#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
抓今日賽事、盤口、傷兵、球員近況、球隊近況與新聞，輸出 sports/data/today.json。
只用 ESPN 的公開接口，不需要任何 key，也不用錢。

用法：python sports/pipeline/fetch.py            # 正式輸出
      python sports/pipeline/fetch.py --out x.json
"""
import json
import os
import sys
import time
import argparse
from datetime import datetime, timedelta, timezone
from statistics import mean
from zoneinfo import ZoneInfo

import requests

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from names_zh import LEAGUE_ZH, STATUS_ZH, OUT_LIKE, MAYBE_LIKE, POS_ZH, zh_name  # noqa: E402

TW = ZoneInfo("Asia/Taipei")
SITE = "https://site.api.espn.com/apis/site/v2/sports"
SITE_V2 = "https://site.api.espn.com/apis/v2/sports"
WEB = "https://site.web.api.espn.com/apis/common/v3/sports"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
    "Accept": "application/json, text/plain, */*",
}
WINDOW_BACK_H = 10     # 保留幾小時前開打的比賽（含剛結束的）
WINDOW_AHEAD_H = 40    # 往後看幾小時

# total_th：兩隊近五場平均總分 與 盤口大小分 差多少才建議大/小
LEAGUES = [
    dict(key="nba", sport="basketball", path="basketball/nba", injuries=True, gamelog=True,
         stat="PTS", hot_min=10, total_th=6, headshot="nba"),
    dict(key="nhl", sport="hockey", path="hockey/nhl", injuries=True, gamelog=True,
         stat="PTS", hot_min=0.5, total_th=0.6, headshot="nhl"),
    dict(key="nfl", sport="football", path="football/nfl", injuries=True, gamelog=False,
         total_th=4, headshot="nfl"),
    dict(key="eng.1", sport="soccer", path="soccer/eng.1", injuries=True, gamelog=False, total_th=0.5, headshot="soccer"),
    dict(key="esp.1", sport="soccer", path="soccer/esp.1", injuries=True, gamelog=False, total_th=0.5, headshot="soccer"),
    dict(key="ita.1", sport="soccer", path="soccer/ita.1", injuries=True, gamelog=False, total_th=0.5, headshot="soccer"),
    dict(key="ger.1", sport="soccer", path="soccer/ger.1", injuries=True, gamelog=False, total_th=0.5, headshot="soccer"),
    dict(key="fra.1", sport="soccer", path="soccer/fra.1", injuries=True, gamelog=False, total_th=0.5, headshot="soccer"),
    dict(key="uefa.champions", sport="soccer", path="soccer/uefa.champions", injuries=False, gamelog=False, total_th=0.5, headshot="soccer"),
    dict(key="uefa.europa", sport="soccer", path="soccer/uefa.europa", injuries=False, gamelog=False, total_th=0.5, headshot="soccer"),
    dict(key="atp", sport="tennis", path="tennis/atp", injuries=False, gamelog=False, tennis=True, total_th=99, headshot="tennis"),
    dict(key="wta", sport="tennis", path="tennis/wta", injuries=False, gamelog=False, tennis=True, total_th=99, headshot="tennis"),
]
LEAGUES_BY_KEY = {lg["key"]: lg for lg in LEAGUES}
SPORTS = [
    dict(key="all", label="全部"),
    dict(key="basketball", label="籃球"),
    dict(key="soccer", label="足球"),
    dict(key="hockey", label="冰球"),
    dict(key="tennis", label="網球"),
    dict(key="football", label="美足"),
]

SESSION = requests.Session()
SESSION.headers.update(HEADERS)
ERRORS = []
CALLS = {"n": 0}


def log(msg):
    print(msg, flush=True)


def get(url, params=None, tries=3, quiet=False):
    last = None
    for i in range(tries):
        try:
            CALLS["n"] += 1
            r = SESSION.get(url, params=params, timeout=25)
            if r.status_code == 200:
                time.sleep(0.12)
                return r.json()
            if r.status_code in (400, 404):
                if not quiet:
                    ERRORS.append(f"{r.status_code} {url}")
                return None
            last = f"http {r.status_code}"
        except Exception as e:  # noqa: BLE001
            last = repr(e)
        time.sleep(1.5 * (i + 1))
    if not quiet:
        ERRORS.append(f"fail {url} {last}")
    return None


# ---------- 小工具 ----------
def parse_dt(s):
    if not s:
        return None
    s = s.replace("Z", "+00:00")
    try:
        dt = datetime.fromisoformat(s)
    except ValueError:
        return None
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    return dt


def iso(dt):
    return dt.isoformat(timespec="minutes") if dt else None


def to_american(v):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().upper().replace("+", "")
    if s in ("EVEN", "EV", "PK", "PICK"):
        return 100.0
    try:
        return float(s)
    except ValueError:
        return None


def dec(a):
    if a is None:
        return None
    return round(1 + (a / 100 if a > 0 else 100 / abs(a)), 2)


def implied(a):
    if a is None:
        return None
    return (100 / (a + 100)) if a > 0 else (abs(a) / (abs(a) + 100))


def num(v):
    try:
        if v is None or v == "":
            return None
        return float(v)
    except (TypeError, ValueError):
        return None


def headshot_url(athlete, lg):
    h = athlete.get("headshot")
    if isinstance(h, dict):
        h = h.get("href")
    if h:
        return h
    aid = athlete.get("id")
    if aid and lg.get("headshot") in ("nba", "nhl", "nfl", "soccer"):
        return f"https://a.espncdn.com/i/headshots/{lg['headshot']}/players/full/{aid}.png"
    return None


def team_label(t):
    return t.get("name_zh") or t.get("short") or t.get("name") or ""


# ---------- 盤口 ----------
RAW = {}


def keep_raw(name, obj):
    """保留第一筆原始資料，--debug-dir 時寫出來，方便看 ESPN 的欄位長什麼樣。"""
    if name not in RAW and obj is not None:
        RAW[name] = obj


def dec_to_american(d):
    d = num(d)
    if not d or d <= 1:
        return None
    return float(round((d - 1) * 100)) if d >= 2 else float(round(-100 / (d - 1)))


def _ml_from_dict(ml):
    """moneyLine 可能是數字、字串，或 {american, decimal, displayValue, ...} 的物件。"""
    if isinstance(ml, dict):
        v = to_american(ml.get("american") or ml.get("alternateDisplayValue") or ml.get("displayValue"))
        if v is None and ml.get("decimal") is not None:
            v = dec_to_american(ml.get("decimal"))
        return v
    return to_american(ml)


def _ml_from_block(block):
    """homeTeamOdds / awayTeamOdds / drawOdds：舊版直接有 moneyLine，新版藏在 current/close/open 裡。"""
    if not isinstance(block, dict):
        return None
    v = _ml_from_dict(block.get("moneyLine"))
    if v is not None:
        return v
    for k in ("current", "close", "open"):
        sub = block.get(k)
        if isinstance(sub, dict):
            v = _ml_from_dict(sub.get("moneyLine"))
            if v is not None:
                return v
    return None


def _ml_from_market(o, side):
    """另一種版本：o["moneyline"] = {home:{close:{odds}}, away:{...}, draw:{...}}"""
    ml = o.get("moneyline")
    if not isinstance(ml, dict):
        return None
    s = ml.get(side)
    if not isinstance(s, dict):
        return None
    for k in ("close", "current", "open"):
        sub = s.get(k)
        if isinstance(sub, dict):
            v = _ml_from_dict(sub.get("odds") if sub.get("odds") is not None else sub.get("american"))
            if v is not None:
                return v
    return None


def _line_from_market(o, market, side, key="line"):
    m = o.get(market)
    if not isinstance(m, dict):
        return None
    s = m.get(side)
    if not isinstance(s, dict):
        return None
    for k in ("close", "current", "open"):
        sub = s.get(k)
        if isinstance(sub, dict) and sub.get(key) is not None:
            return num(str(sub.get(key)).replace("+", ""))
    return None


SPREAD_SIGMA = {"basketball": 11.5, "football": 13.5, "hockey": 2.0}


def spread_prob(spread, sport):
    """沒有獨贏賠率時，用讓分盤估主隊勝率（常態分布近似）。spread 為主隊讓分，負數代表主隊讓。"""
    sigma = SPREAD_SIGMA.get(sport)
    if spread is None or not sigma:
        return None
    from math import erf, sqrt
    return 0.5 * (1 + erf((-spread) / (sigma * sqrt(2))))


def parse_odds(comp, sport=None):
    arr = comp.get("odds") or []
    o = next((x for x in arr if isinstance(x, dict)), None)
    if not o:
        return None
    keep_raw(f"odds_{sport}", o)
    ml_h = _ml_from_block(o.get("homeTeamOdds"))
    if ml_h is None:
        ml_h = _ml_from_market(o, "home")
    ml_a = _ml_from_block(o.get("awayTeamOdds"))
    if ml_a is None:
        ml_a = _ml_from_market(o, "away")
    ml_d = _ml_from_block(o.get("drawOdds"))
    if ml_d is None:
        ml_d = _ml_from_market(o, "draw")
    spread = num(o.get("spread"))
    if spread is None:
        spread = _line_from_market(o, "pointSpread", "home")
    total = num(o.get("overUnder"))
    if total is None:
        total = _line_from_market(o, "total", "over")
    probs = {k: implied(v) for k, v in (("home", ml_h), ("away", ml_a), ("draw", ml_d)) if v is not None}
    source = "moneyline"
    if "home" not in probs or "away" not in probs:
        p = spread_prob(spread, sport)
        if p is not None:
            probs = {"home": p, "away": 1 - p}
            source = "spread"
    s = sum(probs.values())
    fair = {k: round(v / s, 3) for k, v in probs.items()} if s else {}
    fav = None
    if fair.get("home") is not None and fair.get("away") is not None:
        fav = "home" if fair["home"] >= fair["away"] else "away"
    elif (o.get("homeTeamOdds") or {}).get("favorite"):
        fav = "home"
    elif (o.get("awayTeamOdds") or {}).get("favorite"):
        fav = "away"
    return dict(
        provider=(o.get("provider") or {}).get("name"),
        details=o.get("details"),
        spread=spread,
        total=total,
        ml_home=ml_h, ml_away=ml_a, ml_draw=ml_d,
        dec_home=dec(ml_h), dec_away=dec(ml_a), dec_draw=dec(ml_d),
        prob_home=fair.get("home"), prob_away=fair.get("away"), prob_draw=fair.get("draw"),
        prob_source=source if fair else None,
        fav=fav,
    )


# ---------- 賽事解析 ----------
def status_text(lg, st, typ):
    state = typ.get("state")
    if state == "pre":
        return "未開賽"
    if state == "post":
        return "已結束" if typ.get("completed") else (typ.get("description") or "已結束")
    clock = st.get("displayClock") or ""
    p = st.get("period") or 0
    sp = lg["sport"]
    if typ.get("name") == "STATUS_HALFTIME":
        return "中場休息"
    if sp == "soccer":
        return f"進行中 {clock}".strip()
    if sp == "tennis":
        return f"第 {p} 盤" if p else "進行中"
    if p and p <= 4:
        label = f"第 {p} 節"
    elif p and p > 4:
        label = f"延長 {p - 4}"
    else:
        label = "進行中"
    if sp == "hockey" and p and p <= 3:
        label = f"第 {p} 節"
    elif sp == "hockey" and p and p > 3:
        label = "延長賽"
    return f"{label} {clock}".strip()


def parse_team(c, lg):
    t = c.get("team") or {}
    abbr = t.get("abbreviation") or ""
    name = t.get("displayName") or t.get("name") or ""
    short = t.get("shortDisplayName") or t.get("name") or name
    logo = t.get("logo")
    if not logo:
        logos = t.get("logos") or []
        logo = logos[0].get("href") if logos else None
    recs = {}
    for r in c.get("records") or []:
        recs[(r.get("name") or r.get("type") or "").lower()] = r.get("summary")
    score = c.get("score")
    if isinstance(score, dict):
        score = score.get("displayValue")
    leaders = []
    seen = set()
    for cat in c.get("leaders") or []:
        for l in (cat.get("leaders") or [])[:1]:
            a = l.get("athlete") or {}
            if not a.get("id") or a["id"] in seen:
                continue
            seen.add(a["id"])
            leaders.append(dict(
                id=str(a.get("id")),
                player=a.get("displayName") or a.get("fullName") or a.get("shortName"),
                headshot=headshot_url(a, lg),
                position=(a.get("position") or {}).get("abbreviation"),
                stat=cat.get("shortDisplayName") or cat.get("abbreviation") or cat.get("name"),
                value=l.get("displayValue"),
            ))
    return dict(
        id=str(t.get("id") or ""),
        abbr=abbr,
        name=name,
        short=short,
        name_zh=zh_name(lg["key"], abbr, name, short),
        logo=logo,
        color=t.get("color"),
        record=recs.get("overall") or recs.get("total") or (next(iter(recs.values()), None) if recs else None),
        record_home=recs.get("home"),
        record_away=recs.get("road") or recs.get("away"),
        score=score,
        winner=bool(c.get("winner")),
        form_str=c.get("form"),
        leaders=leaders,
        injuries=[],
        key_out=[],
        out_count=0,
        maybe_count=0,
        hot=[],
        form=[],
        form_note=None,
        streak=None,
        last10=None,
        avg_for5=None,
        avg_against5=None,
        avg_total5=None,
        avg_margin5=None,
    )


def parse_player(c, lg):
    a = c.get("athlete") or {}
    flag = a.get("flag") or {}
    sets = []
    for ls in c.get("linescores") or []:
        v = ls.get("value")
        if v is not None:
            sets.append(str(int(v)) if float(v).is_integer() else str(v))
    return dict(
        id=str(a.get("id") or c.get("id") or ""),
        abbr=(a.get("shortName") or "")[:3].upper(),
        name=a.get("displayName") or a.get("fullName") or "",
        short=a.get("shortName") or a.get("displayName") or "",
        name_zh=None,
        logo=flag.get("href") if isinstance(flag, dict) else None,
        headshot=headshot_url(a, lg) if a.get("headshot") else None,
        color=None,
        record=None, record_home=None, record_away=None,
        score=" ".join(sets) if sets else (c.get("score") if isinstance(c.get("score"), str) else None),
        winner=bool(c.get("winner")),
        form_str=None, leaders=[], injuries=[], key_out=[], out_count=0, maybe_count=0,
        hot=[], form=[], form_note=None, streak=None, last10=None,
        avg_for5=None, avg_against5=None, avg_total5=None, avg_margin5=None,
        rank=a.get("rank") or (a.get("ranks") or [{}])[0].get("rank") if isinstance(a.get("ranks"), list) else a.get("rank"),
    )


def make_game(lg, ev, comp, is_tennis=False, group_name=None):
    comps = comp.get("competitors") or []
    if len(comps) < 2:
        return None
    if is_tennis:
        if group_name and "Doubles" in group_name:
            return None
        home_c, away_c = comps[0], comps[1]
        home, away = parse_player(home_c, lg), parse_player(away_c, lg)
        if not home["name"] or not away["name"]:
            return None
    else:
        home_c = next((c for c in comps if c.get("homeAway") == "home"), comps[0])
        away_c = next((c for c in comps if c.get("homeAway") == "away"), comps[1])
        home, away = parse_team(home_c, lg), parse_team(away_c, lg)
    start = parse_dt(comp.get("date") or ev.get("date"))
    if not start:
        return None
    st = comp.get("status") or ev.get("status") or {}
    typ = st.get("type") or {}
    state = typ.get("state") or "pre"
    start_tw = start.astimezone(TW)
    headline = None
    for h in comp.get("headlines") or []:
        if h.get("description"):
            headline = h.get("description")
            break
    bc = []
    for b in comp.get("broadcasts") or []:
        bc.extend(b.get("names") or [])
    venue = (comp.get("venue") or {}).get("fullName")
    name = ev.get("shortName") or ev.get("name")
    if is_tennis:
        name = f"{home['name']} vs {away['name']}"
    return dict(
        id=str(comp.get("id") or ev.get("id")),
        league=lg["key"],
        league_zh=LEAGUE_ZH.get(lg["key"], lg["key"]),
        sport=lg["sport"],
        tennis=is_tennis,
        tournament=(ev.get("name") if is_tennis else None),
        round=group_name,
        name=name,
        start=iso(start.astimezone(timezone.utc)),
        start_tw=iso(start_tw),
        day_tw=start_tw.strftime("%Y-%m-%d"),
        time_tw=start_tw.strftime("%H:%M"),
        status=dict(state=state, completed=bool(typ.get("completed")), detail=status_text(lg, st, typ),
                    clock=st.get("displayClock"), period=st.get("period")),
        venue=venue,
        broadcast=", ".join(dict.fromkeys(bc)) or None,
        home=home,
        away=away,
        odds=parse_odds(comp, lg["sport"]),
        headline=headline,
        news=[],
        suggestion=None,
    )


def fetch_games(lg, dates):
    games = {}
    for d in dates:
        params = {"dates": d.strftime("%Y%m%d"), "limit": 300}
        data = get(f"{SITE}/{lg['path']}/scoreboard", params=params)
        if not data:
            continue
        if data.get("events"):
            keep_raw(f"scoreboard_event_{lg['key']}", data["events"][0])
        for ev in data.get("events") or []:
            try:
                if lg.get("tennis"):
                    groupings = ev.get("groupings") or []
                    if groupings:
                        for gp in groupings:
                            gname = (gp.get("grouping") or {}).get("displayName")
                            for comp in gp.get("competitions") or []:
                                g = make_game(lg, ev, comp, is_tennis=True, group_name=gname)
                                if g:
                                    games[g["id"]] = g
                    else:
                        for comp in ev.get("competitions") or []:
                            g = make_game(lg, ev, comp, is_tennis=True)
                            if g:
                                games[g["id"]] = g
                else:
                    for comp in (ev.get("competitions") or [])[:1]:
                        g = make_game(lg, ev, comp)
                        if g:
                            games[g["id"]] = g
            except Exception as e:  # noqa: BLE001
                ERRORS.append(f"{lg['key']} 解析賽事 {ev.get('id')} 失敗: {e!r}")
    return list(games.values())


# ---------- 傷兵 ----------
def fetch_injuries(lg):
    data = get(f"{SITE}/{lg['path']}/injuries")
    out = {}
    if not data:
        return out
    if data.get("injuries"):
        keep_raw(f"injuries_{lg['key']}", data["injuries"][0])
    for team in data.get("injuries") or []:
        tid = str(team.get("id") or (team.get("team") or {}).get("id") or "")
        lst = []
        for inj in team.get("injuries") or []:
            a = inj.get("athlete") or {}
            status = inj.get("status") or (inj.get("type") or {}).get("description") or ""
            det = inj.get("details") or {}
            if status in MAYBE_LIKE:
                sev = "maybe"
            elif status.startswith("Injured Reserve") or status == "Out For Season":
                sev = "ir"
            elif status in OUT_LIKE:
                sev = "out"
            else:
                sev = "other"
            lst.append(dict(
                id=str(a.get("id") or ""),
                player=a.get("displayName") or a.get("fullName") or a.get("shortName"),
                headshot=headshot_url(a, lg),
                position=(a.get("position") or {}).get("abbreviation"),
                position_zh=POS_ZH.get((a.get("position") or {}).get("abbreviation") or ""),
                status=status,
                status_zh=STATUS_ZH.get(status, status),
                severity=sev,
                detail=det.get("type") or det.get("detail"),
                side=det.get("side"),
                return_date=det.get("returnDate"),
                comment=inj.get("shortComment") or inj.get("longComment"),
                date=inj.get("date"),
            ))
        if tid:
            out[tid] = lst
    return out


# ---------- 球隊近況 ----------
def _score_val(c):
    s = c.get("score")
    if isinstance(s, dict):
        return num(s.get("value") if s.get("value") is not None else s.get("displayValue"))
    return num(s)


def parse_schedule(data, team_id, lg, comp=None):
    """把球隊賽程裡已完成的比賽整理成近況列。comp 是賽事代號（足球跨賽事合併時用）。"""
    rows = []
    for ev in (data or {}).get("events") or []:
        comp = (ev.get("competitions") or [{}])[0]
        typ = ((comp.get("status") or ev.get("status") or {}).get("type") or {})
        if not typ.get("completed"):
            continue
        comps = comp.get("competitors") or []
        me = next((c for c in comps if str((c.get("team") or {}).get("id") or c.get("id")) == str(team_id)), None)
        opp = next((c for c in comps if c is not me), None)
        if not me or not opp:
            continue
        dt = parse_dt(comp.get("date") or ev.get("date"))
        if not dt:
            continue
        pf, pa = _score_val(me), _score_val(opp)
        if pf is None or pa is None:
            continue
        if me.get("winner") is True:
            res = "W"
        elif opp.get("winner") is True:
            res = "L"
        else:
            res = "W" if pf > pa else ("L" if pf < pa else "D")
        ot = (opp.get("team") or {})
        rows.append(dict(
            date=dt.astimezone(TW).strftime("%m/%d"),
            ts=dt.timestamp(),
            opp=ot.get("abbreviation") or ot.get("shortDisplayName") or ot.get("displayName"),
            opp_zh=zh_name(lg["key"], ot.get("abbreviation"), ot.get("displayName"), ot.get("shortDisplayName")),
            opp_logo=ot.get("logo") or ((ot.get("logos") or [{}])[0].get("href") if ot.get("logos") else None),
            home=(me.get("homeAway") == "home"),
            pf=pf, pa=pa, result=res,
            comp=(LEAGUE_ZH.get(comp, comp) if comp and lg["sport"] == "soccer" else None),
        ))
    rows.sort(key=lambda r: r["ts"])
    return rows


# 足球：一支球隊同時踢國內聯賽和歐戰，近況要跨賽事合併才有意義
SOCCER_FORM_LEAGUES = ["eng.1", "esp.1", "ita.1", "ger.1", "fra.1", "ned.1", "por.1", "tur.1", "bel.1", "sco.1",
                       "nor.1", "aut.1", "sui.1", "gre.1", "den.1", "cze.1", "ukr.1", "swe.1", "pol.1", "cro.1",
                       "srb.1", "hun.1", "aze.1", "svk.1", "cyp.1", "kaz.1", "isr.1"]
SOCCER_CUPS = ["uefa.champions", "uefa.europa", "uefa.europa.conf"]
_SOCCER_TEAM_LEAGUE = None   # team_id -> 國內聯賽
_SOCCER_TEAM_CUPS = None     # team_id -> {盃賽}


def _standings_team_ids(data):
    children = (data or {}).get("children") or []
    if not children and (data or {}).get("standings"):
        children = [dict(standings=data["standings"])]
    for ch in children:
        for e in ((ch.get("standings") or {}).get("entries")) or []:
            tid = str((e.get("team") or {}).get("id") or "")
            if tid:
                yield tid


def soccer_team_index():
    global _SOCCER_TEAM_LEAGUE, _SOCCER_TEAM_CUPS
    if _SOCCER_TEAM_LEAGUE is not None:
        return
    _SOCCER_TEAM_LEAGUE, _SOCCER_TEAM_CUPS = {}, {}
    for lgk in SOCCER_FORM_LEAGUES:
        for tid in _standings_team_ids(get(f"{SITE_V2}/soccer/{lgk}/standings", quiet=True)):
            _SOCCER_TEAM_LEAGUE.setdefault(tid, lgk)
    for cup in SOCCER_CUPS:
        for tid in _standings_team_ids(get(f"{SITE_V2}/soccer/{cup}/standings", quiet=True)):
            _SOCCER_TEAM_CUPS.setdefault(tid, set()).add(cup)
    log(f"足球球隊索引：國內聯賽 {len(_SOCCER_TEAM_LEAGUE)} 隊、歐戰 {len(_SOCCER_TEAM_CUPS)} 隊")


def fetch_form(lg, team):
    if not team.get("id"):
        return
    data = get(f"{SITE}/{lg['path']}/teams/{team['id']}/schedule")
    if data and data.get("events"):
        keep_raw(f"schedule_event_{lg['key']}", data["events"][0])
    rows = parse_schedule(data, team["id"], lg, comp=lg["key"])
    note = None
    if lg["sport"] == "soccer":
        soccer_team_index()
        others = set()
        dom = _SOCCER_TEAM_LEAGUE.get(team["id"])
        if dom and dom != lg["key"]:
            others.add(dom)
        for cup in _SOCCER_TEAM_CUPS.get(team["id"], ()):
            if cup != lg["key"]:
                others.add(cup)
        for other in sorted(others):
            d2 = get(f"{SITE}/soccer/{other}/teams/{team['id']}/schedule", quiet=True)
            rows += parse_schedule(d2, team["id"], lg, comp=other)
        seen, merged = set(), []
        for r in sorted(rows, key=lambda r: r["ts"]):
            k = (round(r["ts"] / 3600), r["opp"])
            if k in seen:
                continue
            seen.add(k)
            merged.append(r)
        rows = merged
        if others:
            note = "含國內聯賽與歐戰"
    if not rows and lg["sport"] in ("basketball", "hockey", "football"):
        data = get(f"{SITE}/{lg['path']}/teams/{team['id']}/schedule", params={"seasontype": 1})
        rows = parse_schedule(data, team["id"], lg)
        if rows:
            note = "季前賽"
    if not rows:
        return
    for r in rows:
        r.pop("ts", None)
    last = rows[-10:]
    team["form"] = last
    team["form_note"] = note
    w = sum(1 for r in last if r["result"] == "W")
    l = sum(1 for r in last if r["result"] == "L")
    d = sum(1 for r in last if r["result"] == "D")
    team["last10"] = f"{w}勝{l}敗" + (f"{d}和" if d else "")
    team["last10_w"] = w
    # 連勝連敗
    s = 0
    first = last[-1]["result"]
    for r in reversed(last):
        if r["result"] == first:
            s += 1
        else:
            break
    team["streak"] = {"W": "連勝", "L": "連敗", "D": "連和"}[first] + str(s) if s >= 2 else None
    last5 = last[-5:]
    team["avg_for5"] = round(mean(r["pf"] for r in last5), 1)
    team["avg_against5"] = round(mean(r["pa"] for r in last5), 1)
    team["avg_total5"] = round(team["avg_for5"] + team["avg_against5"], 1)
    team["avg_margin5"] = round(team["avg_for5"] - team["avg_against5"], 1)


# ---------- 球員近況（火燙/冰冷） ----------
def fetch_player_form(lg, athlete):
    data = get(f"{WEB}/{lg['path']}/athletes/{athlete['id']}/gamelog")
    if not data:
        return None
    keep_raw(f"gamelog_{lg['key']}", dict(labels=data.get("labels"), seasonTypes=(data.get("seasonTypes") or [])[:1],
                                         events=dict(list((data.get("events") or {}).items())[:2])))
    labels = data.get("labels") or []
    stat = lg["stat"]
    if stat not in labels:
        return None
    idx = labels.index(stat)
    meta = data.get("events") or {}
    rows = []
    season_name = None
    for st in data.get("seasonTypes") or []:
        season_name = season_name or st.get("displayName")
        for cat in st.get("categories") or []:
            for e in cat.get("events") or []:
                stats = e.get("stats") or []
                if idx >= len(stats):
                    continue
                v = num(stats[idx])
                if v is None:
                    continue
                m = meta.get(str(e.get("eventId"))) or {}
                dt = parse_dt(m.get("gameDate"))
                opp = (m.get("opponent") or {})
                rows.append(dict(ts=dt.timestamp() if dt else 0,
                                 date=dt.astimezone(TW).strftime("%m/%d") if dt else "",
                                 opp=opp.get("abbreviation") or opp.get("displayName"),
                                 result=m.get("gameResult"), v=v))
    if len(rows) < 3:
        return None
    rows.sort(key=lambda r: r["ts"], reverse=True)
    last5 = rows[:5]
    l5 = mean(r["v"] for r in last5)
    season = mean(r["v"] for r in rows)
    delta = ((l5 - season) / season * 100) if season > 0 else 0
    label = "平穩"
    if delta >= 15 and l5 >= lg["hot_min"]:
        label = "火燙"
    elif delta <= -15 and season >= lg["hot_min"]:
        label = "冰冷"
    return dict(
        id=athlete["id"], player=athlete["player"], headshot=athlete.get("headshot"),
        position=athlete.get("position"), stat=stat, season_name=season_name,
        last5=round(l5, 1), season=round(season, 1), delta_pct=round(delta),
        label=label, games=[dict(date=r["date"], opp=r["opp"], v=r["v"]) for r in reversed(last5)],
        n_games=len(rows),
    )


# ---------- 新聞 ----------
def fetch_news(lg):
    data = get(f"{SITE}/{lg['path']}/news", params={"limit": 60})
    arts = []
    for a in (data or {}).get("articles") or []:
        teams = set()
        for c in a.get("categories") or []:
            if c.get("type") == "team":
                tid = c.get("teamId") or (c.get("team") or {}).get("id")
                if tid:
                    teams.add(str(tid))
        imgs = a.get("images") or []
        link = ((a.get("links") or {}).get("web") or {}).get("href")
        arts.append(dict(
            title=a.get("headline"), desc=a.get("description"), url=link,
            image=imgs[0].get("url") if imgs else None,
            published=a.get("published"), teams=teams,
        ))
    arts.sort(key=lambda x: x.get("published") or "", reverse=True)
    return arts


# ---------- 排名 ----------
def fetch_standings(lg):
    data = get(f"{SITE_V2}/{lg['path']}/standings")
    if not data:
        return None
    keep_raw(f"standings_{lg['key']}", dict(children=[dict(name=c.get("name"), standings=dict(entries=((c.get("standings") or {}).get("entries") or [])[:1])) for c in (data.get("children") or [])[:1]]))
    groups = []
    children = data.get("children") or []
    if not children and data.get("standings"):
        children = [dict(name=data.get("name") or LEAGUE_ZH.get(lg["key"]), standings=data["standings"])]
    conf_zh = {"Eastern Conference": "東區", "Western Conference": "西區",
               "American Football Conference": "美聯", "National Football Conference": "國聯"}
    for ch in children:
        entries = ((ch.get("standings") or {}).get("entries")) or []
        rows = []
        for e in entries:
            t = e.get("team") or {}
            st = {}
            for s in e.get("stats") or []:
                st[s.get("name")] = s
            def val(n):
                s = st.get(n) or {}
                return s.get("value")
            def disp(n):
                s = st.get(n) or {}
                return s.get("displayValue")
            logos = t.get("logos") or []
            rows.append(dict(
                team_id=str(t.get("id") or ""),
                team=t.get("displayName"),
                abbr=t.get("abbreviation"),
                name_zh=zh_name(lg["key"], t.get("abbreviation"), t.get("displayName"), t.get("shortDisplayName")),
                logo=logos[0].get("href") if logos else None,
                w=val("wins"), l=val("losses"), d=val("ties") if val("ties") is not None else val("draws"),
                otl=val("otLosses"),
                pct=disp("winPercent"), gb=disp("gamesBehind"), streak=disp("streak"),
                points=val("points"), gp=val("gamesPlayed"),
                rank=val("rank") or val("playoffSeed"),
            ))
        if lg["sport"] == "soccer":
            rows.sort(key=lambda r: (-(r["points"] or 0), (r["rank"] or 99)))
        elif lg["sport"] == "hockey":
            rows.sort(key=lambda r: (-(r["points"] or 0)))
        else:
            rows.sort(key=lambda r: ((r["rank"] or 99), -(num(r["pct"]) or 0)))
        for i, r in enumerate(rows, 1):
            r["pos"] = i
        groups.append(dict(name=conf_zh.get(ch.get("name"), ch.get("name") or ""), rows=rows))
    return dict(groups=groups) if groups else None


# ---------- 建議 ----------
def build_suggestion(g, lg):
    o = g.get("odds") or {}
    sp = lg["sport"]
    res = dict(pick=None, pick_label=None, play="觀望", confidence="低", summary="", reasons=[], cautions=[], totals=None)
    if g["status"]["state"] == "post":
        res["summary"] = "比賽已結束"
        return res
    if g.get("tennis"):
        if o.get("prob_home") is not None and o.get("prob_away") is not None:
            fav = "home" if o["prob_home"] >= o["prob_away"] else "away"
            pf = max(o["prob_home"], o["prob_away"])
            res["reasons"].append(f"國際盤 {g[fav]['name']} 勝率約 {round(pf * 100)}%")
            if pf >= 0.68:
                res.update(pick=fav, pick_label=g[fav]["name"], play="不讓分", confidence="中",
                           summary=f"傾向 {g[fav]['name']}，盤口明顯看好，信心中")
            else:
                res["summary"] = "兩人實力接近，建議觀望"
        else:
            res["summary"] = "目前沒有盤口資料，先觀望"
        return res
    if o.get("prob_home") is None or o.get("prob_away") is None:
        res["summary"] = "目前沒有盤口資料，先觀望"
        return res
    ph, pa = o["prob_home"], o["prob_away"]
    fav, dog = ("home", "away") if ph >= pa else ("away", "home")
    pf = max(ph, pa)
    ft, dt = g[fav], g[dog]
    fn, dn = team_label(ft), team_label(dt)
    score = 0.0
    reasons, cautions = [], []
    side = "主" if fav == "home" else "客"
    if o.get("prob_source") == "spread":
        reasons.append(f"依讓分盤推估{side}隊 {fn} 勝率約 {round(pf * 100)}%（無獨贏賠率）")
    else:
        reasons.append(f"國際盤看好{side}隊 {fn}，勝率約 {round(pf * 100)}%")
    if o.get("prob_draw"):
        reasons.append(f"和局機率約 {round(o['prob_draw'] * 100)}%")
    if ft["key_out"]:
        cautions.append(f"{fn} 主力缺陣：{'、'.join(ft['key_out'][:3])}")
        score -= 1.0
    elif ft["out_count"] >= 3:
        cautions.append(f"{fn} 有 {ft['out_count']} 名球員缺陣")
        score -= 0.3
    if dt["key_out"]:
        reasons.append(f"{dn} 主力缺陣：{'、'.join(dt['key_out'][:3])}")
        score += 0.7
    elif dt["out_count"] >= 3:
        reasons.append(f"{dn} 有 {dt['out_count']} 名球員缺陣")
        score += 0.2
    if ft.get("form"):
        n = len(ft["form"])
        fw = ft.get("last10_w", 0)
        if n >= 5 and fw / n >= 0.7:
            reasons.append(f"{fn} 近 {n} 場 {fw} 勝，狀態好")
            score += 0.5
        elif n >= 5 and fw / n <= 0.3:
            cautions.append(f"{fn} 近 {n} 場只有 {fw} 勝")
            score -= 0.5
        if ft.get("streak", "") and ft["streak"].startswith("連敗"):
            cautions.append(f"{fn} {ft['streak']}")
            score -= 0.2
    if dt.get("form"):
        n = len(dt["form"])
        dw = dt.get("last10_w", 0)
        if dt.get("streak", "") and dt["streak"].startswith("連勝") and int(dt["streak"][2:]) >= 3:
            cautions.append(f"{dn} {dt['streak']}，手感正熱")
            score -= 0.3
        elif n >= 5 and dw / n >= 0.7:
            cautions.append(f"{dn} 近 {n} 場 {dw} 勝，不弱")
            score -= 0.2
    for p in ft.get("hot") or []:
        if p["label"] == "火燙":
            reasons.append(f"{p['player']} 近 5 場場均 {p['last5']} {p['stat']}，比季平均高 {p['delta_pct']}%")
            score += 0.2
        elif p["label"] == "冰冷":
            cautions.append(f"{p['player']} 近 5 場場均 {p['last5']} {p['stat']}，比季平均低 {abs(p['delta_pct'])}%")
            score -= 0.2
    for p in dt.get("hot") or []:
        if p["label"] == "火燙":
            cautions.append(f"對手 {p['player']} 手感火燙，近 5 場場均 {p['last5']} {p['stat']}")
            score -= 0.15
    if sp == "soccer":
        hi, mid = 0.60, 0.50
    else:
        hi, mid = 0.66, 0.56
    adj = pf + score * 0.05
    if adj >= hi and score >= 0:
        conf = "高"
    elif adj >= mid:
        conf = "中"
    else:
        conf = "低"
    if conf == "低":
        res.update(pick=None, pick_label=None, play="觀望", confidence="低",
                   summary="兩隊實力接近或訊號互相矛盾，建議觀望")
    else:
        play = "不讓分"
        spread = o.get("spread")
        if sp == "soccer":
            if pf >= 0.70 and spread is not None and abs(spread) >= 1:
                play = "讓分"
        elif pf >= 0.72 and spread is not None:
            m = ft.get("avg_margin5")
            if m is not None and m >= abs(spread):
                play = "讓分"
                reasons.append(f"{fn} 近 5 場平均贏 {m} 分，高於讓分 {abs(spread)}")
            else:
                cautions.append("賠率偏低，不讓分獲利有限")
        res.update(pick=fav, pick_label=fn, play=play, confidence=conf,
                   summary=f"傾向 {fn}，建議玩法「{play}」，信心{conf}")
    th = lg.get("total_th", 99)
    h, a = g["home"], g["away"]
    if o.get("total") and h.get("avg_total5") and a.get("avg_total5"):
        exp = (h["avg_total5"] + a["avg_total5"]) / 2
        diff = exp - o["total"]
        lean = "無明顯方向"
        if diff >= th:
            lean = "大"
            reasons.append(f"兩隊近 5 場平均總分 {exp:.1f}，高於盤口 {o['total']}，可考慮大分")
        elif diff <= -th:
            lean = "小"
            reasons.append(f"兩隊近 5 場平均總分 {exp:.1f}，低於盤口 {o['total']}，可考慮小分")
        res["totals"] = dict(lean=lean, line=o["total"], expected=round(exp, 1))
    # 台灣運彩盤口：用國際盤的公平勝率去算台灣賠率的期望值
    tw = g.get("tw") or {}
    ml = tw.get("ml") or {}
    if res["pick"] and ml.get(res["pick"]) and o.get(f"prob_{res['pick']}"):
        dec_tw = ml[res["pick"]]
        p = o[f"prob_{res['pick']}"]
        ev = p * dec_tw - 1
        res["tw_ev"] = round(ev * 100)
        res["tw_dec"] = dec_tw
        if ev >= 0.02:
            reasons.append(f"台灣運彩不讓分賠率 {dec_tw}，以國際盤勝率算期望值 +{ev * 100:.0f}%，划算")
        elif ev <= -0.10:
            cautions.append(f"台灣運彩不讓分賠率 {dec_tw} 偏低，期望值 {ev * 100:.0f}%，考慮改玩讓分或縮小注碼")
        else:
            reasons.append(f"台灣運彩不讓分賠率 {dec_tw}，期望值 {ev * 100:+.0f}%，普通")
    res["reasons"] = reasons
    res["cautions"] = cautions
    return res


# ---------- 整合 ----------
def finalize_team(team, lg, injuries_map):
    inj = injuries_map.get(team["id"], []) if team.get("id") else []
    order = {"maybe": 0, "out": 1, "ir": 2, "other": 3}
    inj = sorted(inj, key=lambda x: order.get(x["severity"], 3))
    team["injuries"] = inj
    team["out_count"] = sum(1 for x in inj if x["severity"] in ("out", "ir"))
    team["maybe_count"] = sum(1 for x in inj if x["severity"] == "maybe")
    leader_ids = {l["id"] for l in team.get("leaders") or []}
    leader_names = {l["player"] for l in team.get("leaders") or []}
    key = [x["player"] for x in inj if x["severity"] in ("out", "ir") and ((x["id"] and x["id"] in leader_ids) or x["player"] in leader_names)]
    team["key_out"] = key
    team["key_maybe"] = [x["player"] for x in inj if x["severity"] == "maybe" and ((x["id"] and x["id"] in leader_ids) or x["player"] in leader_names)]


def aggregate(games):
    """傷兵快報與火燙榜。同一隊在窗口內打兩場（背靠背）只列一次，掛在較早的那場。"""
    feed, hot = [], []
    seen_inj, seen_hot = set(), set()
    for g in sorted(games, key=lambda x: x["start"]):
        if g["status"]["state"] == "post":
            continue
        for side in ("home", "away"):
            t = g[side]
            for x in t.get("injuries") or []:
                key = (t.get("id"), x.get("id") or x.get("player"))
                if x["severity"] in ("out", "maybe", "ir") and key not in seen_inj:
                    seen_inj.add(key)
                    feed.append(dict(
                        game_id=g["id"], league=g["league"], league_zh=g["league_zh"], sport=g["sport"],
                        team_id=t["id"], team=team_label(t), team_logo=t.get("logo"),
                        player=x["player"], headshot=x.get("headshot"), position=x.get("position"),
                        status=x["status"], status_zh=x["status_zh"], severity=x["severity"],
                        detail=x.get("detail"), return_date=x.get("return_date"),
                        key=(x["player"] in t.get("key_out", []) or x["player"] in t.get("key_maybe", [])),
                        time_tw=g["time_tw"], day_tw=g["day_tw"],
                    ))
            for p in t.get("hot") or []:
                key = (t.get("id"), p.get("id") or p.get("player"))
                if p["label"] != "平穩" and key not in seen_hot:
                    seen_hot.add(key)
                    hot.append(dict(game_id=g["id"], league=g["league"], league_zh=g["league_zh"], sport=g["sport"],
                                    team=team_label(t), team_logo=t.get("logo"), **p))
    # 排序：主力優先；再來是「今天到底上不上」最有用的存疑/逐日觀察，然後缺陣，長期傷兵名單放最後
    sev_order = {"maybe": 0, "out": 1, "ir": 2}
    feed.sort(key=lambda x: (0 if x["key"] else 1, sev_order.get(x["severity"], 3), x["day_tw"], x["time_tw"]))
    # 每個聯盟最多 30 筆，避免 NFL 那種一隊二十幾個傷兵把版面塞滿
    per = {}
    trimmed = []
    for x in feed:
        per[x["league"]] = per.get(x["league"], 0) + 1
        if per[x["league"]] <= 30:
            trimmed.append(x)
    hot.sort(key=lambda x: -abs(x["delta_pct"]))
    return trimmed, hot


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "data", "today.json"))
    ap.add_argument("--leagues", default="", help="逗號分隔，只抓這些聯盟（測試用）")
    ap.add_argument("--debug-dir", default="", help="把每種 ESPN 回應的第一筆原始資料寫到這個資料夾")
    args = ap.parse_args()

    now = datetime.now(timezone.utc)
    now_tw = now.astimezone(TW)
    dates = [now.date() - timedelta(days=1), now.date(), now.date() + timedelta(days=1)]
    lo = now - timedelta(hours=WINDOW_BACK_H)
    hi = now + timedelta(hours=WINDOW_AHEAD_H)
    only = {s.strip() for s in args.leagues.split(",") if s.strip()}

    all_games = []
    standings = {}
    for lg in LEAGUES:
        if only and lg["key"] not in only:
            continue
        try:
            process_league(lg, dates, lo, hi, all_games, standings)
        except Exception as e:  # noqa: BLE001
            import traceback
            ERRORS.append(f"{lg['key']} 整個聯盟失敗: {e!r}")
            log(f"[{lg['key']}] 失敗：{e!r}")
            traceback.print_exc()

    all_games.sort(key=lambda g: (g["start"], g["league"]))
    # 台灣運彩盤口（抓不到就全部沒有，不影響其他）
    tw_count = 0
    try:
        from taiwan_odds import attach_taiwan_odds
        tw_count = attach_taiwan_odds(all_games, ERRORS, log)
    except Exception as e:  # noqa: BLE001
        ERRORS.append(f"台灣盤口失敗: {e!r}")
        log(f"台灣盤口失敗：{e!r}")
    log(f"台灣盤口：配對到 {tw_count} 場")
    for g in all_games:
        lg = LEAGUES_BY_KEY[g["league"]]
        try:
            g["suggestion"] = build_suggestion(g, lg)
        except Exception as e:  # noqa: BLE001
            ERRORS.append(f"{lg['key']} 建議 {g['id']} 失敗: {e!r}")
            g["suggestion"] = dict(pick=None, pick_label=None, play="觀望", confidence="低",
                                   summary="資料不完整，先觀望", reasons=[], cautions=[], totals=None)
    feed, hot = aggregate(all_games)
    out = dict(
        generated_at=now_tw.isoformat(timespec="minutes"),
        generated_ts=int(now.timestamp()),
        today_tw=now_tw.strftime("%Y-%m-%d"),
        sports=SPORTS,
        leagues={k: LEAGUE_ZH.get(k, k) for k in LEAGUES_BY_KEY},
        games=all_games,
        injury_feed=feed[:80],
        hot_players=hot[:40],
        standings=standings,
        source="ESPN 公開資料；盤口為美國莊家（ESPN BET），非台灣運彩",
        errors=ERRORS[:50],
        calls=CALLS["n"],
    )
    for g in all_games:
        for side in ("home", "away"):
            g[side].pop("leaders", None)
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, separators=(",", ":"))
    log(f"寫入 {args.out}：{len(all_games)} 場、傷兵 {len(feed)} 筆、火燙/冰冷 {len(hot)} 人、錯誤 {len(ERRORS)}")
    for e in ERRORS[:20]:
        log("  ! " + e)
    if args.debug_dir:
        os.makedirs(args.debug_dir, exist_ok=True)
        for name, obj in RAW.items():
            with open(os.path.join(args.debug_dir, f"{name}.json"), "w", encoding="utf-8") as f:
                json.dump(obj, f, ensure_ascii=False, indent=1)
        log(f"原始資料樣本 {len(RAW)} 份寫到 {args.debug_dir}")


def process_league(lg, dates, lo, hi, all_games, standings):
    if True:
        t0 = time.time()
        games = fetch_games(lg, dates)
        keep = []
        for g in games:
            st = parse_dt(g["start"])
            if g["status"]["state"] == "in" or (lo <= st <= hi):
                keep.append(g)
        keep.sort(key=lambda g: g["start"])
        log(f"[{lg['key']}] 賽事 {len(games)} 場，窗口內 {len(keep)} 場")
        if not keep:
            return
        injuries_map = fetch_injuries(lg) if lg.get("injuries") else {}
        news = fetch_news(lg) if not lg.get("tennis") else []
        seen_teams = {}
        for g in keep:
            for side in ("home", "away"):
                t = g[side]
                finalize_team(t, lg, injuries_map)
                if g.get("tennis"):
                    continue
                if t["id"] in seen_teams:
                    cached = seen_teams[t["id"]]
                    for k in ("form", "form_note", "streak", "last10", "last10_w", "avg_for5", "avg_against5", "avg_total5", "avg_margin5", "hot"):
                        t[k] = cached.get(k)
                    continue
                try:
                    fetch_form(lg, t)
                except Exception as e:  # noqa: BLE001
                    import traceback
                    tb = traceback.extract_tb(e.__traceback__)[-1]
                    ERRORS.append(f"{lg['key']} 球隊 {t.get('abbr')} 近況失敗: {e!r} @ {tb.name}:{tb.lineno} {tb.line}")
                if lg.get("gamelog") and g["status"]["state"] != "post":
                    hot = []
                    for ath in (t.get("leaders") or [])[:3]:
                        try:
                            pf = fetch_player_form(lg, ath)
                        except Exception as e:  # noqa: BLE001
                            ERRORS.append(f"{lg['key']} 球員 {ath.get('player')} 近況失敗: {e!r}")
                            pf = None
                        if pf:
                            hot.append(pf)
                    t["hot"] = hot
                seen_teams[t["id"]] = t
            ids = {g["home"]["id"], g["away"]["id"]}
            picked = []
            for a in news:
                if a["teams"] & ids:
                    picked.append(dict(title=a["title"], desc=a["desc"], url=a["url"], image=a["image"],
                                       published=a["published"],
                                       team="home" if g["home"]["id"] in a["teams"] else "away"))
                if len(picked) >= 4:
                    break
            g["news"] = picked
        all_games.extend(keep)
        if not lg.get("tennis"):
            try:
                s = fetch_standings(lg)
            except Exception as e:  # noqa: BLE001
                ERRORS.append(f"{lg['key']} 排名失敗: {e!r}")
                s = None
            if s:
                standings[lg["key"]] = dict(label=LEAGUE_ZH.get(lg["key"]), sport=lg["sport"], **s)
        log(f"[{lg['key']}] 完成，{time.time() - t0:.1f}s，累計呼叫 {CALLS['n']}")


if __name__ == "__main__":
    main()

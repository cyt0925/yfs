# -*- coding: utf-8 -*-
"""台灣運彩盤口：從玩運彩（playsport.cc）的盤口頁抓今明兩天的運彩盤，配對到 ESPN 的賽事上。
抓不到就什麼都不做；這份資料是加分項，不能影響主流程。

輸出掛在每場 g["tw"]：
  {"source": "玩運彩", "url": ..., "updated": iso,
   "spread": {"line": "雷霆 -4.5", "home": 1.75, "away": 1.75},
   "ml": {"home": 1.45, "away": 2.55, "draw": None},
   "total": {"line": 224.5, "over": 1.75, "under": 1.75}}
"""
import re
import time
import html as htmlmod
from datetime import datetime, timedelta, timezone

import requests

TW = timezone(timedelta(hours=8))
BASE = "https://www.playsport.cc"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,*/*;q=0.8",
    "Accept-Language": "zh-TW,zh;q=0.9,en;q=0.8",
    "Referer": "https://www.playsport.cc/",
}

# ESPN 聯盟 -> 玩運彩 allianceid（探路後填入；不確定的先不填）
ALLIANCE = {
    "nba": 3,
    # "nhl": ?, "nfl": ?, "eng.1": ?, ...
}

# 玩運彩的中文隊名跟我們的對照不一樣時，在這裡補
ALIASES = {
    "塞爾蒂克": "塞爾提克", "賽爾提克": "塞爾提克", "小牛": "獨行俠", "76人": "七六人",
    "灰狼": "灰狼", "拓荒者": "拓荒者", "快艇": "快艇", "國王": "國王",
}


def norm(name):
    if not name:
        return ""
    s = re.sub(r"\s+", "", str(name))
    s = s.replace("隊", "")
    return ALIASES.get(s, s)


def fetch_html(url, tries=2):
    for i in range(tries):
        try:
            r = requests.get(url, headers=HEADERS, timeout=25)
            if r.status_code == 200:
                time.sleep(0.6)
                return r.text
            if r.status_code in (403, 404):
                return None
        except Exception:  # noqa: BLE001
            pass
        time.sleep(2 * (i + 1))
    return None


def parse_scale_page(text):
    """解析 /predict/scale 頁面。回傳 [{"home":..., "away":..., "time":"07:00", "spread":{}, "ml":{}, "total":{}}]
    這裡的解析規則在探路確認版面後才會定案；目前先用保守的通用方法。"""
    games = []
    # TODO: 探路後依實際 HTML 結構改寫
    return games


def attach_taiwan_odds(games, errors, log):
    """把台灣盤口掛到 games 上，回傳配對到的場數。"""
    if not ALLIANCE:
        return 0
    now = datetime.now(TW)
    days = [now.strftime("%Y%m%d"), (now + timedelta(days=1)).strftime("%Y%m%d")]
    by_league = {}
    for g in games:
        if g.get("tennis") or g["status"]["state"] == "post":
            continue
        by_league.setdefault(g["league"], []).append(g)
    matched = 0
    for lg, aid in ALLIANCE.items():
        if lg not in by_league:
            continue
        rows = []
        for d in days:
            url = f"{BASE}/predict/scale?allianceid={aid}&gametime={d}"
            text = fetch_html(url)
            if not text:
                errors.append(f"台灣盤口 {lg} {d} 抓不到")
                continue
            try:
                for r in parse_scale_page(text):
                    r["day"] = d
                    r["url"] = url
                    rows.append(r)
            except Exception as e:  # noqa: BLE001
                errors.append(f"台灣盤口 {lg} {d} 解析失敗: {e!r}")
        if not rows:
            continue
        index = {}
        for r in rows:
            index[(r["day"], norm(r["home"]), norm(r["away"]))] = r
        for g in by_league[lg]:
            key = (g["day_tw"].replace("-", ""), norm(g["home"].get("name_zh")), norm(g["away"].get("name_zh")))
            r = index.get(key)
            if not r:
                continue
            g["tw"] = dict(source="玩運彩", url=r["url"], updated=now.isoformat(timespec="minutes"),
                           spread=r.get("spread"), ml=r.get("ml"), total=r.get("total"))
            matched += 1
        log(f"[{lg}] 台灣盤口 {len(rows)} 筆，配對 {sum(1 for g in by_league[lg] if g.get('tw'))} 場")
    return matched

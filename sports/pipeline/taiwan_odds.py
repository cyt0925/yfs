# -*- coding: utf-8 -*-
"""台灣運彩盤口：從玩運彩（playsport.cc）抓今明兩天的運彩盤，配對到 ESPN 的賽事上。
抓不到就什麼都不做；這份資料是加分項，不能影響主流程。

資料來源兩個：
  1. /predict/scale?allianceid=X&gametime=YYYYMMDD  盤口表（讓分、不讓分、大小）＋玩運彩 gameid
  2. https://ls.playsport.cc/ls_json.php?alliance=X&gamedate=YYYYMMDD  即時比分 JSON，有 gameid 對應的
     英文縮寫（official_id 像 NBA_20261014_NY@TOR），拿來跟 ESPN 的縮寫配對

輸出掛在每場 g["tw"]：
  {"source": "玩運彩", "url": ..., "updated": iso, "no": 325,
   "spread": {"line": "暴龍 -10.5", "line_home": -10.5, "home": 1.73, "away": 1.7},
   "ml": {"home": 1.16, "away": 3.2, "draw": None},
   "total": {"line": 216.5, "over": 1.7, "under": 1.73}}
"""
import re
import time
import json
from datetime import datetime, timedelta, timezone

import requests
from bs4 import BeautifulSoup

TW = timezone(timedelta(hours=8))
BASE = "https://www.playsport.cc"
LS = "https://ls.playsport.cc/ls_json.php"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/json;q=0.9,*/*;q=0.8",
    "Accept-Language": "zh-TW,zh;q=0.9,en;q=0.8",
    "Referer": "https://www.playsport.cc/",
}

# ESPN 運動/聯盟 -> 玩運彩 allianceid
ALLIANCE_BY_LEAGUE = {"nba": 3, "nhl": 91, "nfl": 93}
ALLIANCE_BY_SPORT = {"soccer": 4}

# 玩運彩中文隊名 -> 我們的中文隊名（兩邊不一樣時補在這裡）
ALIASES = {
    3: {"塞爾蒂克": "塞爾提克", "賽爾提克": "塞爾提克", "小牛": "獨行俠", "76人": "七六人"},
    91: {"猛瑪象": "猛獁象", "佛羅里": "美洲豹", "巨鴨": "鴨子", "海怪": "海妖", "荒野": "野人", "達拉斯": "星辰",
         "首都": "首都人", "加拿大": "加拿大人", "騎士": "金騎士", "黑鷹": "黑鷹"},
    4: {"阿森納": "兵工廠", "拜仁": "拜仁慕尼黑", "巴黎": "巴黎聖日耳曼", "巴黎聖日爾曼": "巴黎聖日耳曼",
        "馬競": "馬德里競技", "皇馬": "皇家馬德里", "國米": "國際米蘭", "尤文": "尤文圖斯", "多特": "多特蒙德",
        "比亞雷亞爾": "比利亞雷亞爾", "維京足球會": "維京", "PSV恩荷芬": "PSV", "RB萊比錫": "萊比錫",
        "薩巴赫馬薩齊爾": "薩巴赫", "斯拉維亞": "布拉格斯拉維亞", "葡萄牙體育": "里斯本競技", "士砵亭": "里斯本競技",
        "博德格林特": "博德閃耀", "頓涅茨克礦工": "頓內次克礦工", "飛燕諾": "費耶諾德", "波爾圖": "波圖", "阿積士": "阿賈克斯",
        "賓菲加": "本菲卡", "卡拉巴克": "卡拉巴赫", "聯合聖吉羅斯": "聖吉羅斯聯", "紐卡斯爾": "紐卡索", "紐卡素": "紐卡索",
        "加拉塔沙雷": "加拉塔薩雷", "費倫巴治": "費內巴切", "奧林比亞高斯": "奧林匹亞科斯", "LASK林茨": "林茨",
        "比利亞雷爾": "比利亞雷亞爾", "皇家馬德里": "皇家馬德里", "馬德里體育會": "馬德里競技", "紐卡索": "紐卡索", "愛華頓": "艾佛頓",
        "熱刺": "熱刺", "車路士": "切爾西", "阿士東維拉": "阿斯頓維拉", "曼徹斯特聯": "曼聯", "曼徹斯特城": "曼城"},
}

DEBUG = {}


def norm(name, aid=None):
    if not name:
        return ""
    s = re.sub(r"\s+", "", str(name))
    s = re.sub(r"[()（）\-–·．.]", "", s)
    for suf in ("隊", "足球俱樂部", "體育會", "足球會"):
        if s.endswith(suf) and len(s) > len(suf) + 1:
            s = s[: -len(suf)]
    return ALIASES.get(aid, {}).get(s, s)


def name_score(a, b, aid=None):
    """隊名相似度：2 完全相同、1 一邊包含另一邊、0 不像。"""
    a, b = norm(a, aid), norm(b, aid)
    if not a or not b:
        return 0
    if a == b:
        return 2
    if len(a) >= 2 and len(b) >= 2 and (a in b or b in a):
        return 1
    return 0


def names_match(a, b, aid=None):
    return name_score(a, b, aid) > 0


def fetch(url, tries=2, as_json=False):
    for i in range(tries):
        try:
            r = requests.get(url, headers=HEADERS, timeout=25)
            if r.status_code == 200:
                time.sleep(0.6)
                return r.json() if as_json else r.text
            if r.status_code in (403, 404):
                return None
        except Exception:  # noqa: BLE001
            pass
        time.sleep(2 * (i + 1))
    return None


def _num(s):
    if s is None:
        return None
    s = str(s).strip().replace(",", "").replace("+", "")
    try:
        return float(s)
    except ValueError:
        return None


def _cell(td):
    """一格運彩盤：回傳 (方向文字, 線, 賠率)。例：('客', '+10.5', '1.7') / ('主', '', '1.16') / ('大', '216.5', '1.7')"""
    if td is None:
        return None, None, None
    side = td.select_one(".team-side")
    wrap = td.select_one(".data-wrap")
    side_t = side.get_text(strip=True) if side else ""
    line_t, odds_t = "", ""
    if wrap:
        st = wrap.find("strong")
        sp = wrap.find("span")
        line_t = st.get_text(strip=True) if st else ""
        odds_t = sp.get_text(strip=True).lstrip(",").strip() if sp else ""
    return side_t, line_t, odds_t


def parse_scale_page(text):
    """解析 /predict/scale 頁。每場比賽兩列（客、主），足球多一列和局。
    不靠欄位 class 判斷玩法，而是看每格的方向文字：客/主＋有線＝讓分、客/主沒線＝不讓分、和＝和局、大/小＝大小。"""
    soup = BeautifulSoup(text, "html.parser")
    table = soup.select_one("table.predictgame-table")
    if not table:
        return []
    games = []
    cur = None
    for tr in table.select("tr"):
        info = tr.select_one("td.td-gameinfo")
        team_td = tr.select_one("td.td-teaminfo")
        if info is not None:
            h3 = info.find("h3")
            h4 = info.find("h4")
            cur = dict(gameid=tr.get("gameid"), no=h3.get_text(strip=True) if h3 else None,
                       time=(h4.get_text(strip=True) if h4 else ""), rows=[])
            games.append(cur)
        if cur is None:
            continue
        name, tid = "", None
        if team_td is not None:
            a = team_td.find("a")
            name = (a or team_td).get_text(strip=True)
            if a and a.get("href"):
                m = re.search(r"teamid=(\d+)", a["href"])
                tid = m.group(1) if m else None
        cells = []
        for td in tr.find_all("td"):
            cls = " ".join(td.get("class") or [])
            if "td-bank-bet" not in cls and "td-universal-bet" not in cls:
                continue
            side, line, odds = _cell(td)
            if not side and not line and not odds:
                continue
            cells.append(dict(intl=("td-universal" in cls), side=side, line=line, odds=odds))
        if not cells and not name:
            continue
        cur["rows"].append(dict(name=name, teamid=tid, cells=cells))
    out = []
    for g in games:
        rows = [r for r in g["rows"]]
        named = [r for r in rows if r["name"]]
        if len(named) < 2:
            continue
        away, home = named[0], named[1]
        spread = dict(line=None, line_home=None, home=None, away=None)
        ml = dict(home=None, away=None, draw=None)
        total = dict(line=None, over=None, under=None)
        intl = dict(spread_home=None, total=None)
        for r in rows:
            is_home = r is home
            is_away = r is away
            for c in r["cells"]:
                side, line, odds = c["side"], _num(c["line"]) if c["line"] else None, _num(c["odds"])
                if c["intl"]:
                    if side in ("主", "客") and line is not None:
                        intl["spread_home"] = line if is_home else (-line if is_away and intl["spread_home"] is None else intl["spread_home"])
                    elif side in ("大", "小") and line is not None:
                        intl["total"] = line
                    continue
                if side in ("和", "和局"):
                    ml["draw"] = odds
                elif side in ("主", "客") and c["line"]:
                    if is_home or side == "主":
                        spread["line_home"] = line
                        spread["home"] = odds
                        spread["line"] = f"{home['name']} {c['line']}"
                    else:
                        spread["away"] = odds
                        if spread["line_home"] is None and line is not None:
                            spread["line_home"] = -line
                        if not spread["line"]:
                            spread["line"] = f"{away['name']} {c['line']}"
                elif side in ("主", "客"):
                    if is_home or side == "主":
                        ml["home"] = odds
                    else:
                        ml["away"] = odds
                elif side == "大":
                    total["line"] = line
                    total["over"] = odds
                elif side == "小":
                    total["line"] = line if line is not None else total["line"]
                    total["under"] = odds
        if spread["home"] is None and spread["away"] is None:
            spread = None
        if ml["home"] is None and ml["away"] is None:
            ml = None
        if total["over"] is None and total["under"] is None:
            total = None
        if intl["spread_home"] is None and intl["total"] is None:
            intl = None
        m = re.match(r"(AM|PM)\s*(\d{1,2}):(\d{2})", g["time"])
        hhmm = None
        if m:
            h = int(m.group(2)) % 12 + (12 if m.group(1) == "PM" else 0)
            hhmm = f"{h:02d}:{m.group(3)}"
        out.append(dict(gameid=g["gameid"], no=g["no"], time=hhmm, away=away["name"], home=home["name"],
                        away_teamid=away["teamid"], home_teamid=home["teamid"],
                        spread=spread, ml=ml, total=total, intl=intl,
                        sides=[(r["name"], [(c["side"], c["line"], c["odds"]) for c in r["cells"]]) for r in rows]))
    return out


def fetch_livescore_ids(aid, day):
    d = fetch(f"{LS}?alliance={aid}&gamedate={day}", as_json=True)
    out = {}
    if not isinstance(d, dict):
        return out
    for k, v in d.items():
        if not isinstance(v, dict) or not v.get("gameid"):
            continue
        h, a = v.get("horiginname"), v.get("aoriginname")
        m = re.search(r"_([A-Za-z]+)@([A-Za-z]+)", str(v.get("official_id") or ""))
        if m:
            a, h = a or m.group(1), h or m.group(2)
        out[str(v["gameid"])] = dict(official=v.get("official_id"), h=h, a=a,
                                     hname=v.get("hname"), aname=v.get("aname"), dateon=v.get("dateon"))
    return out


def _abbr_set(t):
    s = {(t.get("abbr") or "").upper()}
    return {x for x in s if x}


def _minutes(t):
    try:
        h, m = str(t).split(":")
        return int(h) * 60 + int(m)
    except Exception:  # noqa: BLE001
        return None


def match_game(g, rows, aid=None):
    """在同一天、同一聯盟的玩運彩列表中找出對應 ESPN 賽事。
    1. 英文縮寫完全相同（NBA、NHL 的比分 JSON 有）
    2. 開賽時間相同（容許 20 分鐘）＋ 兩邊隊名相似度計分，取最高分
    """
    ha, aa = (g["home"].get("abbr") or "").upper(), (g["away"].get("abbr") or "").upper()
    hz, az = g["home"].get("name_zh") or g["home"].get("name"), g["away"].get("name_zh") or g["away"].get("name")
    for r in rows:
        ls = r.get("ls") or {}
        if ha and aa and ls.get("h") and ls.get("a") and ls["h"].upper() == ha and ls["a"].upper() == aa:
            return r, "abbr"
    gt = _minutes(g.get("time_tw"))
    best, best_score = None, 0
    for r in rows:
        rt = _minutes(r.get("time"))
        if gt is not None and rt is not None and abs(gt - rt) > 20:
            continue
        ls = r.get("ls") or {}
        sh = max(name_score(r["home"], hz, aid), name_score(ls.get("hname"), hz, aid))
        sa = max(name_score(r["away"], az, aid), name_score(ls.get("aname"), az, aid))
        if sh == 0 or sa == 0:
            continue
        score = sh + sa + (1 if gt is not None and rt is not None else 0)
        if score > best_score:
            best, best_score = r, score
    if best:
        return best, f"name{best_score}"
    return None, None


def attach_taiwan_odds(games, errors, log):
    """把台灣盤口掛到 games 上，回傳配對到的場數。"""
    now = datetime.now(TW)
    days = [(now + timedelta(days=i)).strftime("%Y%m%d") for i in range(3)]
    groups = {}
    for g in games:
        if g["status"]["state"] == "post":
            continue
        aid = ALLIANCE_BY_LEAGUE.get(g["league"]) or ALLIANCE_BY_SPORT.get(g["sport"])
        if not aid:
            continue
        groups.setdefault(aid, []).append(g)
    matched = 0
    for aid, lst in groups.items():
        rows_by_day = {}
        for d in days:
            url = f"{BASE}/predict/scale?allianceid={aid}&gametime={d}"
            text = fetch(url)
            if not text:
                errors.append(f"台灣盤口 alliance {aid} {d} 抓不到")
                continue
            try:
                rows = parse_scale_page(text)
            except Exception as e:  # noqa: BLE001
                errors.append(f"台灣盤口 alliance {aid} {d} 解析失敗: {e!r}")
                continue
            ids = fetch_livescore_ids(aid, d)
            for r in rows:
                r["url"] = url
                r["day"] = d
                r["ls"] = ids.get(str(r.get("gameid")))
            rows_by_day[d] = rows
            if rows and f"sample_{aid}" not in DEBUG:
                DEBUG[f"sample_{aid}"] = dict(rows[0])
                DEBUG[f"sides_{aid}"] = [r.get("sides") for r in rows[:3]]
        total_rows = sum(len(v) for v in rows_by_day.values())
        if not total_rows:
            log(f"[alliance {aid}] 玩運彩沒有盤口資料")
            continue
        hit = 0
        unmatched = []
        for g in lst:
            day = g["day_tw"].replace("-", "")
            rows = rows_by_day.get(day) or []
            r, how = match_game(g, rows, aid)
            if not r:
                gt = _minutes(g["time_tw"])
                near = [f"{x['away']}@{x['home']}" for x in rows if gt is not None and _minutes(x.get("time")) is not None and abs(_minutes(x["time"]) - gt) <= 20]
                unmatched.append(f"{g['league']} {g['away'].get('name_zh') or g['away']['name']}@{g['home'].get('name_zh') or g['home']['name']} {g['time_tw']}" + (f"（同時段：{'、'.join(near[:4])}）" if near else ""))
                continue
            g["tw"] = dict(source="玩運彩", url=r["url"], updated=now.isoformat(timespec="minutes"), no=r.get("no"),
                           matched_by=how, spread=r.get("spread"), ml=r.get("ml"), total=r.get("total"),
                           tw_names=f"{r['away']} @ {r['home']}")
            hit += 1
        matched += hit
        log(f"[alliance {aid}] 玩運彩 {total_rows} 筆，ESPN {len(lst)} 場，配對 {hit} 場")
        if unmatched:
            log(f"[alliance {aid}] 沒配到：{'；'.join(unmatched[:12])}")
            names = sorted({f"{r['away']}@{r['home']}" for rows in rows_by_day.values() for r in rows})
            log(f"[alliance {aid}] 玩運彩那邊的隊名：{'；'.join(names[:40])}")
    if DEBUG:
        for k, v in DEBUG.items():
            log(f"玩運彩樣本 {k}: " + json.dumps(v, ensure_ascii=False)[:900])
    return matched

# -*- coding: utf-8 -*-
"""球隊、聯盟、傷兵狀態的中文對照。找不到的就顯示英文原名。"""
import re
import unicodedata

LEAGUE_ZH = {
    "nba": "NBA",
    "nhl": "NHL 冰球",
    "nfl": "NFL 美式足球",
    "eng.1": "英超",
    "esp.1": "西甲",
    "ita.1": "義甲",
    "ger.1": "德甲",
    "fra.1": "法甲",
    "uefa.champions": "歐冠",
    "uefa.europa": "歐聯",
    "uefa.europa.conf": "歐協聯",
    "atp": "ATP 男網",
    "wta": "WTA 女網",
    # 以下只用在足球「近況」的賽事標籤
    "ned.1": "荷甲", "por.1": "葡超", "tur.1": "土超", "bel.1": "比甲", "sco.1": "蘇超",
    "nor.1": "挪超", "aut.1": "奧甲", "sui.1": "瑞士超", "gre.1": "希臘超", "den.1": "丹超",
    "cze.1": "捷甲", "ukr.1": "烏超", "swe.1": "瑞典超", "pol.1": "波甲", "cro.1": "克甲",
    "srb.1": "塞超", "hun.1": "匈甲", "aze.1": "亞塞拜然超", "svk.1": "斯洛伐克超",
    "cyp.1": "塞浦路斯甲", "kaz.1": "哈薩克超", "isr.1": "以超",
}

# 依 ESPN 的球隊縮寫
NBA = {
    "ATL": "老鷹", "BOS": "塞爾提克", "BKN": "籃網", "CHA": "黃蜂", "CHI": "公牛",
    "CLE": "騎士", "DAL": "獨行俠", "DEN": "金塊", "DET": "活塞", "GS": "勇士",
    "HOU": "火箭", "IND": "溜馬", "LAC": "快艇", "LAL": "湖人", "MEM": "灰熊",
    "MIA": "熱火", "MIL": "公鹿", "MIN": "灰狼", "NO": "鵜鶘", "NY": "尼克",
    "OKC": "雷霆", "ORL": "魔術", "PHI": "七六人", "PHX": "太陽", "POR": "拓荒者",
    "SAC": "國王", "SA": "馬刺", "TOR": "暴龍", "UTAH": "爵士", "WSH": "巫師",
    "GSW": "勇士", "NOP": "鵜鶘", "NYK": "尼克", "SAS": "馬刺", "UTA": "爵士", "WAS": "巫師",
}

NHL = {
    "ANA": "鴨子", "BOS": "棕熊", "BUF": "軍刀", "CGY": "火焰", "CAR": "颶風",
    "CHI": "黑鷹", "COL": "雪崩", "CBJ": "藍衣", "DAL": "星辰", "DET": "紅翼",
    "EDM": "油人", "FLA": "美洲豹", "LA": "國王", "LAK": "國王", "MIN": "野人",
    "MTL": "加拿大人", "NSH": "掠奪者", "NJ": "魔鬼", "NJD": "魔鬼", "NYI": "島人",
    "NYR": "遊騎兵", "OTT": "參議員", "PHI": "飛人", "PIT": "企鵝", "SJ": "鯊魚",
    "SJS": "鯊魚", "SEA": "海妖", "STL": "藍調", "TB": "閃電", "TBL": "閃電",
    "TOR": "楓葉", "UTAH": "猛獁象", "UTA": "猛獁象", "VAN": "加人", "VGK": "金騎士",
    "VGS": "金騎士", "WSH": "首都人", "WAS": "首都人", "WPG": "噴射機",
}

NFL = {
    "ARI": "紅雀", "ATL": "獵鷹", "BAL": "烏鴉", "BUF": "比爾", "CAR": "黑豹",
    "CHI": "熊", "CIN": "孟加拉虎", "CLE": "布朗", "DAL": "牛仔", "DEN": "野馬",
    "DET": "雄獅", "GB": "包裝工", "HOU": "德州人", "IND": "小馬", "JAX": "美洲虎",
    "KC": "酋長", "LV": "突襲者", "LAC": "閃電", "LAR": "公羊", "MIA": "海豚",
    "MIN": "維京人", "NE": "愛國者", "NO": "聖徒", "NYG": "巨人", "NYJ": "噴射機",
    "PHI": "老鷹", "PIT": "鋼人", "SF": "49人", "SEA": "海鷹", "TB": "海盜",
    "TEN": "泰坦", "WSH": "指揮官", "WAS": "指揮官",
}

# 足球依英文隊名（會先做正規化再比對）
SOCCER = {
    # 英超
    "arsenal": "兵工廠", "aston villa": "阿斯頓維拉", "bournemouth": "伯恩茅斯",
    "brentford": "布倫特福德", "brighton hove albion": "布萊頓", "brighton": "布萊頓",
    "burnley": "伯恩利", "chelsea": "切爾西", "crystal palace": "水晶宮", "everton": "艾佛頓",
    "fulham": "富勒姆", "leeds united": "里茲聯", "liverpool": "利物浦",
    "manchester city": "曼城", "manchester united": "曼聯", "newcastle united": "紐卡索",
    "nottingham forest": "諾丁漢森林", "sunderland": "桑德蘭", "tottenham hotspur": "熱刺",
    "west ham united": "西漢姆聯", "wolverhampton wanderers": "狼隊", "wolves": "狼隊",
    "leicester city": "萊斯特城", "southampton": "南安普敦", "ipswich town": "伊普斯維奇",
    "sheffield united": "雪菲爾聯", "luton town": "盧頓", "norwich city": "諾里奇",
    "watford": "沃特福德", "middlesbrough": "米德斯堡", "coventry city": "考文垂",
    "west bromwich albion": "西布朗",
    # 西甲
    "real madrid": "皇家馬德里", "barcelona": "巴塞隆納", "atletico madrid": "馬德里競技",
    "athletic club": "畢爾包競技", "athletic bilbao": "畢爾包競技", "real sociedad": "皇家社會",
    "real betis": "皇家貝蒂斯", "villarreal": "比利亞雷亞爾", "sevilla": "塞維亞",
    "valencia": "瓦倫西亞", "girona": "赫羅納", "celta vigo": "維戈塞爾塔", "osasuna": "奧薩蘇納",
    "getafe": "赫塔費", "rayo vallecano": "巴列卡諾", "mallorca": "馬約卡", "alaves": "阿拉維斯",
    "espanyol": "西班牙人", "levante": "萊萬特", "elche": "埃爾切", "real oviedo": "奧維耶多",
    "las palmas": "拉斯帕爾馬斯", "leganes": "萊加內斯", "real valladolid": "瓦拉多利德",
    # 義甲
    "internazionale": "國際米蘭", "inter milan": "國際米蘭", "inter": "國際米蘭",
    "ac milan": "AC 米蘭", "milan": "AC 米蘭", "juventus": "尤文圖斯", "napoli": "拿坡里",
    "roma": "羅馬", "as roma": "羅馬", "lazio": "拉齊奧", "atalanta": "亞特蘭大",
    "fiorentina": "佛羅倫斯", "bologna": "波隆那", "torino": "杜林", "udinese": "烏迪內斯",
    "genoa": "熱那亞", "cagliari": "卡利亞里", "parma": "帕爾馬", "lecce": "萊切",
    "como": "科莫", "hellas verona": "維羅納", "verona": "維羅納", "sassuolo": "薩索洛",
    "pisa": "比薩", "cremonese": "克雷莫納", "monza": "蒙扎", "empoli": "恩波利",
    "venezia": "威尼斯",
    # 德甲
    "bayern munich": "拜仁慕尼黑", "bayern munchen": "拜仁慕尼黑", "borussia dortmund": "多特蒙德",
    "bayer leverkusen": "勒沃庫森", "rb leipzig": "萊比錫", "eintracht frankfurt": "法蘭克福",
    "vfb stuttgart": "斯圖加特", "stuttgart": "斯圖加特", "freiburg": "弗萊堡",
    "borussia monchengladbach": "門興", "wolfsburg": "沃夫斯堡", "mainz": "美因茲",
    "werder bremen": "不萊梅", "augsburg": "奧格斯堡", "hoffenheim": "霍芬海姆",
    "union berlin": "柏林聯合", "heidenheim": "海登海姆", "st pauli": "聖保利",
    "hamburg": "漢堡", "hamburger sv": "漢堡", "koln": "科隆", "holstein kiel": "基爾",
    "bochum": "波鴻",
    # 法甲
    "paris saint germain": "巴黎聖日耳曼", "psg": "巴黎聖日耳曼", "marseille": "馬賽",
    "lyon": "里昂", "monaco": "摩納哥", "lille": "里爾", "nice": "尼斯", "lens": "朗斯",
    "rennes": "雷恩", "strasbourg": "史特拉斯堡", "toulouse": "土魯斯", "nantes": "南特",
    "brest": "布雷斯特", "auxerre": "歐塞爾", "le havre": "勒阿弗爾", "angers": "昂熱",
    "lorient": "洛里昂", "metz": "梅斯", "paris": "巴黎 FC", "montpellier": "蒙彼利埃",
    "reims": "蘭斯", "saint etienne": "聖埃蒂安",
    # 歐洲其他
    "ajax": "阿賈克斯", "ajax amsterdam": "阿賈克斯", "psv eindhoven": "PSV", "psv": "PSV",
    "feyenoord": "費耶諾德", "benfica": "本菲卡", "porto": "波圖", "sporting": "里斯本競技",
    "sporting lisbon": "里斯本競技", "celtic": "塞爾提克", "rangers": "流浪者",
    "galatasaray": "加拉塔薩雷", "fenerbahce": "費內巴切", "club brugge": "布魯日",
    "red bull salzburg": "薩爾茲堡", "salzburg": "薩爾茲堡", "shakhtar donetsk": "頓內次克礦工",
    "dinamo zagreb": "薩格勒布迪納摩", "olympiacos": "奧林匹亞科斯", "bodo glimt": "博德閃耀",
    "slavia prague": "布拉格斯拉維亞", "slavia praha": "布拉格斯拉維亞", "copenhagen": "哥本哈根",
    "fc copenhagen": "哥本哈根", "qarabag": "卡拉巴赫", "kairat": "凱拉特", "pafos": "帕福斯",
    "union saint gilloise": "聖吉羅斯聯", "sturm graz": "格拉茨", "young boys": "年輕男孩",
    "sparta prague": "布拉格斯巴達", "sparta praha": "布拉格斯巴達", "red star belgrade": "貝爾格勒紅星",
    "crvena zvezda": "貝爾格勒紅星",
    "brugge": "布魯日", "athletic": "畢爾包競技", "sabah": "薩巴赫", "viking": "維京",
    "lask linz": "林茨", "lask": "林茨", "aek athens": "AEK 雅典", "slovan bratislava": "布拉提斯拉瓦",
    "bod glimt": "博德閃耀", "ferencvaros": "費倫茨瓦羅斯", "kairat almaty": "凱拉特",
    "panathinaikos": "帕納辛奈科斯", "paok": "PAOK", "besiktas": "貝西克塔斯", "braga": "布拉加",
    "lille osc": "里爾", "nice ogc": "尼斯", "ogc nice": "尼斯", "rc lens": "朗斯",
    "stade rennais": "雷恩", "stade brestois": "布雷斯特", "rb salzburg": "薩爾茲堡",
    "fc midtjylland": "中日德蘭", "midtjylland": "中日德蘭", "malmo": "馬爾默", "genk": "亨克",
    "anderlecht": "安德萊赫特", "dynamo kyiv": "基輔迪納摩", "zenit": "聖彼得堡澤尼特",
    "nottingham": "諾丁漢森林", "brighton and hove albion": "布萊頓",
}

STATUS_ZH = {
    "Out": "缺陣",
    "Doubtful": "可能缺陣",
    "Questionable": "存疑",
    "Day-To-Day": "逐日觀察",
    "Probable": "可能上場",
    "Injured Reserve": "傷兵名單",
    "Injured Reserve - Designated to Return": "傷兵名單",
    "Suspension": "禁賽",
    "Suspended": "禁賽",
    "Out For Season": "賽季報銷",
    "Available": "可上場",
    "Active": "可上場",
    "Game Time Decision": "賽前決定",
}

# 哪些狀態視為「很可能不會上場」
OUT_LIKE = {"Out", "Doubtful", "Injured Reserve", "Injured Reserve - Designated to Return",
            "Suspension", "Suspended", "Out For Season"}
MAYBE_LIKE = {"Questionable", "Day-To-Day", "Game Time Decision", "Probable"}

POS_ZH = {
    "PG": "控球後衛", "SG": "得分後衛", "SF": "小前鋒", "PF": "大前鋒", "C": "中鋒",
    "G": "後衛", "F": "前鋒", "F-C": "前鋒/中鋒", "G-F": "後衛/前鋒",
    "LW": "左翼", "RW": "右翼", "D": "防守", "W": "側翼",
    "QB": "四分衛", "RB": "跑衛", "WR": "外接手", "TE": "邊鋒", "K": "踢球員",
    "LB": "線衛", "CB": "角衛", "S": "安全衛", "DE": "防守端鋒", "DT": "防守截鋒",
    "OT": "進攻截鋒", "OG": "進攻護鋒", "P": "棄踢員",
    "GK": "門將", "M": "中場", "FW": "前鋒",
}

_STRIP_WORDS = {"fc", "afc", "cf", "sc", "ac", "club", "de", "futbol", "football", "calcio", "ssc",
                "ss", "us", "sv", "tsg", "vfl", "fsv", "bsc", "1", "1899", "1909", "04", "05"}


def normalize(name):
    if not name:
        return ""
    s = unicodedata.normalize("NFKD", name)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = s.lower().replace("&", " ").replace("-", " ").replace(".", " ").replace("'", "")
    s = re.sub(r"[^a-z0-9 ]+", " ", s)
    words = [w for w in s.split() if w not in _STRIP_WORDS]
    return " ".join(words)


def zh_name(league_key, abbr, name, short=None):
    """回傳中文隊名；找不到就回傳 None。"""
    abbr = (abbr or "").upper()
    if league_key == "nba":
        return NBA.get(abbr)
    if league_key == "nhl":
        return NHL.get(abbr)
    if league_key == "nfl":
        return NFL.get(abbr)
    if league_key in ("atp", "wta"):
        return None
    for cand in (name, short):
        n = normalize(cand)
        if n in SOCCER:
            return SOCCER[n]
    # 再寬鬆一點：隊名中含有字典 key 的
    n = normalize(name)
    for key, zh in SOCCER.items():
        if len(key) >= 4 and (n == key or n.startswith(key + " ") or n.endswith(" " + key)):
            return zh
    return None

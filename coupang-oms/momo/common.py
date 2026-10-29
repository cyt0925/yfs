"""MOMO 共用：商品主檔自動化的工具拿來用（登入者、資料庫、歷程），再加 MOMO 自己的對照與倉別主檔。"""
import base64  # noqa: F401
import collections  # noqa: F401
import io  # noqa: F401
import json  # noqa: F401
import re  # noqa: F401

from flask import current_app, jsonify, render_template, request, send_file  # noqa: F401

from db import get_conn  # noqa: F401
from master.common import _guard_ready, _log, _operator, _row, _rows, now  # noqa: F401
from shopee.common import UNIT_CODES  # noqa: F401 — 單位代碼清單跟需求書第八章同一套

from . import momo_bp

SYSTEM = "MOMO 訂單管理系統"
CHANNEL = "富邦媒體科技(momo)"                    # 需求六：直送採購單 A1
LINES = ("紙品", "潔品", "瑪氏", "寶僑")
MARS = "瑪氏"
SHIP_METHODS = ("原廠直出", "竹運出貨")          # 需求一：預設原廠直出，使用者可改
DIRECT, ZHUYUN = SHIP_METHODS


@momo_bp.before_request
def _ready():
    return _guard_ready()


def split_code(s):
    """「M10279581_RMN」→（M10279581, R）；「M10201240-1_RMN」→（M10201240-1, R）；沒有底線→（原字, ''）。"""
    s = str(s or "").strip()
    if "_" not in s:
        return s, ""
    base, suf = s.rsplit("_", 1)
    return base.strip(), suf.strip()[:1].upper()


def line_of(base):
    """看料號分類（MOMO 下載檔裡紙品、潔品、瑪氏會混在一起，例如 10/23 瑪氏那份夾了一筆五月花）：
    M 開頭＝瑪氏；7 碼數字 0 開頭＝紙品、1 開頭＝潔品（紙潔價格本的料號就是這樣分的）；8 碼以上數字＝寶僑（國條）。認不出來回空字串。"""
    b = str(base or "").strip()
    if b[:1].upper() == "M":
        return MARS
    if b.isdigit() and len(b) == 7:
        return "紙品" if b[0] == "0" else "潔品" if b[0] == "1" else ""
    if b.isdigit() and len(b) >= 8:
        return "寶僑"
    return ""


def conv_line(line):
    """查價格本用的線別：紙品、潔品都在「紙潔價格本」（shp_conv 線別紙潔，跟蝦皮特選共用）。"""
    return "紙潔" if line in ("紙品", "潔品") else line


# ── 倉別主檔（需求第九章）────────────────────────────────────────────────
# 地址收件單位都是 momo進貨組。駐倉廠服的人名不寫進程式（同事要的話在「資料與設定」補），電話照需求書。
# aliases：MOMO 檔、業務表上看過的別名（DC1／DC1(自)、CS3／中衛三CS3、北衛二／北衛二(K)）。只認列出來的，不做模糊比對。
# mars_name：瑪氏採購單 C6 顯示的名稱。需求七原本寫「新竹物流幼獅倉-DC1(自)」，Alice 2026-10-29 更正：前面不加幼獅倉，
# DC1(自)→momo進貨組-DC1(自)、DC6→momo進貨組-DC6，地址用 DC 本身的；其他倉用倉名（SDC 在瑪氏採購單實際寫「MOMO購物網統倉(台南SDC)」，待問）。
_EVERY = "每日 08:00～12:30、13:30～16:00，全年無休；請於 16:00 前完成下貨點收作業。"
_EVERY9 = "每日 09:00～12:30、13:30～16:00，全年無休；請於 16:00 前完成下貨點收作業。"
_TEL = "02-6600-7600"
WH_DEFAULTS = {
    "中衛四": {"address": "台中市豐原區豐勢路一段577巷76號", "hours": _EVERY9, "phone": _TEL},
    "大園倉": {"address": "桃園市大園區航科路161號", "hours": "星期一至六 08:00～12:00、13:00～16:00；星期日、國定假日及颱風天不收貨。", "phone": _TEL},
    "SDC": {"address": "台南市新市區大營里大營1-8號", "hours": _EVERY, "phone": _TEL, "ship_to": "17647744"},
    "明治許厝港倉": {"address": "桃園市大園區北港里10鄰許厝港路529號",
                "hours": "星期一至五 08:00～12:00、13:00～16:00；星期六、日及國定假日 10:00～12:00、13:00～16:00；颱風天休息。請於 16:00 前完成下貨點收作業。",
                "phone": _TEL},
    "FS2": {"address": "桃園市楊梅區梅獅路395巷59號（A棟）", "hours": "星期一至六 09:00～12:30、13:30～16:00；請於 16:00 前完成下貨點收作業。", "phone": _TEL},
    "DC11": {"address": "桃園市楊梅區梅獅路395巷19號（B倉；359號入口處）", "hours": _EVERY, "phone": _TEL},
    "DC10": {"address": "桃園市大園區中山南路二段207-1號", "hours": _EVERY, "phone": _TEL},
    "DC7": {"address": "台南市永康區仁愛街391號", "hours": _EVERY, "phone": _TEL, "ship_to": "17647736"},
    "DC4": {"address": "桃園市蘆竹區蘆竹街177號", "hours": _EVERY, "phone": _TEL},
    "DC1(自)": {"address": "桃園市大園區高鐵站前西路三段77號", "hours": _EVERY, "phone": _TEL, "ship_to": "17647733",
               "aliases": ["DC1"], "mars_name": "momo進貨組-DC1(自)"},
    "中衛三": {"address": "台中市東區自由路三段30號3樓（請至自由一街後方進貨碼頭）", "hours": _EVERY9, "phone": _TEL, "aliases": ["CS3", "中衛三CS3"]},
    "DC6": {"address": "桃園市楊梅區高獅路589號", "hours": _EVERY, "phone": _TEL, "ship_to": "17617038", "mars_name": "momo進貨組-DC6"},
    "新竹物流幼獅倉": {"address": "桃園市楊梅區梅獅路359-1號（梅獅路395巷路底）",
                "hours": "星期一至五 09:00～12:30、13:30～15:00；星期六、日、國定假日及颱風天不收貨。請於 15:00 前完成下貨點收作業。", "phone": _TEL},
    "統一速達鳳仁倉": {"address": "高雄市岡山區岡燕路485巷103號",
                "hours": "星期一至五 10:00～12:30、13:30～15:30；例假日、國定假日及颱風天休息。請於 15:30 前完成下貨點收作業。",
                "phone": "07-6225525、0980467332"},
    "統一速達梧棲倉": {"address": "台中市梧棲區臨港路三段586號",
                "hours": "星期一至五 09:00～12:00、13:00～16:00；星期六、日、國定假日及颱風天不收貨。請於 16:00 前完成下貨點收作業。", "phone": _TEL},
    "高衛三": {"address": "屏東市華安路2號", "hours": _EVERY9, "phone": _TEL},
    "宅配通東元倉": {"address": "桃園市觀音區中山路一段1568-1號",
                "hours": "星期一至六 09:00～11:40、13:00～16:00；星期日、國定假日及颱風天休息。請於 16:00 前完成下貨點收作業。", "phone": "02-6600-7600 分機 2"},
    "立益平鎮倉": {"address": "桃園市平鎮區雙福路一段65號C棟",
               "hours": "星期一至五 08:00～11:50、13:00～16:00；例假日、國定假日及颱風天休息。請於 16:00 前完成下貨點收作業。", "phone": _TEL},
    "北衛二(K)": {"address": "新北市三重區興德路129號1樓", "hours": _EVERY9, "phone": _TEL, "aliases": ["北衛二"]},
}
WH_FIELDS = ("address", "hours", "phone", "contact", "ship_to", "mars_name")
WH_KEY = "momo_wh"
RECEIVER = "momo進貨組"


def wh_settings(conn):
    """{倉名: {address, hours, phone, contact, ship_to, mars_name, aliases, edited}}；資料與設定存過的蓋預設。"""
    row = conn.execute("SELECT value FROM mst_meta WHERE key = ?", (WH_KEY,)).fetchone()
    saved = {}
    if row:
        try:
            saved = json.loads(row["value"] or "{}")
        except ValueError:
            saved = {}
    out = {}
    for name, d in WH_DEFAULTS.items():
        v = {k: str(d.get(k) or "") for k in WH_FIELDS}
        v["mars_name"] = v["mars_name"] or name
        v.update({k: str(x) for k, x in (saved.get(name) or {}).items() if k in WH_FIELDS})
        v["aliases"] = list(d.get("aliases") or [])
        v["edited"] = bool(saved.get(name))
        out[name] = v
    return out


def wh_canon(name, whs=None):
    """MOMO 寫的倉別 → 主檔的倉名；別名照表認，對不到回空字串（不猜）。"""
    n = str(name or "").strip()
    src = whs or {k: {"aliases": v.get("aliases") or []} for k, v in WH_DEFAULTS.items()}
    if n in src:
        return n
    for k, v in src.items():
        if n in (v.get("aliases") or []):
            return k
    return ""


def wh_note(w):
    """直送採購單 C22：地址＋收件單位、商品入庫時間、駐倉廠服電話（需求六）。"""
    tel = w.get("phone") or ""
    if w.get("contact"):
        tel = f"{tel}，{w['contact']}" if tel else w["contact"]
    return "\n".join(x for x in (f"地址：{w.get('address')}，{RECEIVER}" if w.get("address") else "",
                                  f"商品入庫時間：{w.get('hours')}" if w.get("hours") else "",
                                  f"駐倉廠服電話：{tel}" if tel else "") if x)


__all__ = [
    "base64", "collections", "io", "json", "re", "current_app", "jsonify", "render_template", "request", "send_file",
    "get_conn", "_guard_ready", "_log", "_operator", "_row", "_rows", "now", "momo_bp",
    "SYSTEM", "CHANNEL", "LINES", "MARS", "SHIP_METHODS", "DIRECT", "ZHUYUN", "UNIT_CODES",
    "split_code", "line_of", "conv_line", "WH_DEFAULTS", "WH_FIELDS", "WH_KEY", "RECEIVER", "wh_settings", "wh_canon", "wh_note",
]

"""蝦皮特選共用：商品主檔自動化的工具拿來用（登入者、資料庫、歷程），再加蝦皮特選自己的對照。"""
import base64  # noqa: F401
import collections  # noqa: F401
import io  # noqa: F401
import json  # noqa: F401
import re  # noqa: F401

from flask import current_app, jsonify, render_template, request, send_file  # noqa: F401

from db import get_conn  # noqa: F401
from master.common import _guard_ready, _log, _operator, _row, _rows, now  # noqa: F401

from . import shopee_bp

SYSTEM = "蝦皮特選訂單管理系統"
FULFILS = ("竹運出貨", "竹運採購進貨", "供應商直送")      # 需求第四章：匯入時人工選
LINES = ("寶僑", "瑪氏", "紙潔")

# 單位代碼：蝦皮料號底線後面第一個字、竹運庫存效期表客戶品號最後一個字、寶僑價格本「料號3」都是這一套
# （PG價格本整合.xlsx「庫存資料更新」AI:AJ 那張對照表）
UNIT_CODES = {"A": "令", "B": "個", "C": "包", "D": "十個", "E": "台", "F": "外箱", "G": "式", "H": "張", "I": "才",
              "J": "打", "K": "批", "L": "捲", "M": "支", "N": "桶", "O": "片", "P": "瓶", "Q": "盒", "R": "箱",
              "S": "節", "T": "組", "U": "罐", "V": "袋"}

# 蝦皮倉別代碼 → 畫面上的名字（Shopee 供應商進倉規範 V20260901 第 6 頁）
WAREHOUSES = {"TWA": "觀音", "TWX": "安南三", "TWT": "安南一", "TWG": "高鐵南", "TWH": "楊梅", "TWW": "威獅", "TWK": "威獅"}

# 處理狀態（需求第十二章挑第一段用得到的）
ST_IMPORTED = "已匯入"
ST_DIFF = "匯入差異待確認"


@shopee_bp.before_request
def _ready():
    return _guard_ready()


def split_supplier_sku(s):
    """「4987176232878_CNN」→（4987176232878, C）；「M60002536-1_CMN」→（M60002536-1, C）；沒有底線→（原字, ''）。"""
    s = str(s or "").strip()
    if "_" not in s:
        return s, ""
    base, suf = s.rsplit("_", 1)
    suf = suf.strip()
    return base.strip(), suf[:1].upper() if suf else ""


def line_of(base):
    """看料號長相分線別（Jerry 2026-10-07 同意）：M 開頭＝瑪氏；7 碼數字＝紙潔（永豐料號）；8 碼以上數字＝寶僑（國條）。
    認不出來回空字串，由呼叫的地方用同一個檔其他品項的線別補。"""
    b = str(base or "").strip()
    if b[:1].upper() == "M":
        return "瑪氏"
    if b.isdigit() and len(b) == 7:
        return "紙潔"
    if b.isdigit() and len(b) >= 8:
        return "寶僑"
    return ""


def wh_name(code):
    return WAREHOUSES.get(code or "", code or "")


__all__ = [
    "base64", "collections", "io", "json", "re", "current_app", "jsonify", "render_template", "request", "send_file",
    "get_conn", "_guard_ready", "_log", "_operator", "_row", "_rows", "now",
    "shopee_bp", "SYSTEM", "FULFILS", "LINES", "UNIT_CODES", "WAREHOUSES", "ST_IMPORTED", "ST_DIFF",
    "split_supplier_sku", "line_of", "wh_name",
]

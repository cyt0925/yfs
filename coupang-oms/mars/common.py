"""瑪氏出貨共用：商品主檔自動化的工具拿來用（登入者、資料庫、訂單查詢），再加瑪氏自己的對照。"""
import collections  # noqa: F401
import io  # noqa: F401
import json  # noqa: F401
import re  # noqa: F401

import openpyxl  # noqa: F401
from flask import current_app, jsonify, render_template, request, send_file  # noqa: F401

from db import get_conn  # noqa: F401
from normalize import norm_decimal, norm_int, norm_key, norm_text  # noqa: F401
from master.common import (  # noqa: F401
    _guard_ready, _group_of, _line_groups, _log, _operator, _row, _rows, _same, now,
)

from . import mars_bp

LINE = "瑪氏"                      # ② 訂單明細裡線別歸到這個顯示名的才算瑪氏
CAT_CODE = {"chocolate": "CHO", "gum": "GUM", "petcare": "PET"}
# 採購單 B9 那格要寫的品類全名（Alice 文件：Chocolate巧克力／Gum糖果／Petcare寵物；「糖／巧」寫反了，Jerry 確認過）
CAT_FULL = {"CHO": "Chocolate巧克力", "GUM": "Gum糖果", "PET": "Petcare寵物"}
UNITS = ("箱", "盒", "包")
# 要不要依單位（箱／盒／包）再拆一層：Alice 文件有列，但 EIP 與採購單一律用箱，還在等她確認。
# 改成 False，拆單鍵的單位那段就固定寫「箱」，檔名也跟著變。
SPLIT_BY_UNIT = True
EIP_PO_RE = re.compile(r"^PO\d{9}$")   # EIP 採購單號長這樣：PO202609004


@mars_bp.before_request
def _ready():
    return _guard_ready()


def cat_code(category):
    return CAT_CODE.get(norm_text(category).lower(), "")


def label_text(label):
    return "需貼中標" if label == "V" else "不貼中標"


__all__ = [
    "collections", "io", "json", "re", "openpyxl", "current_app", "jsonify", "render_template", "request", "send_file",
    "get_conn", "norm_decimal", "norm_int", "norm_key", "norm_text",
    "_guard_ready", "_group_of", "_line_groups", "_log", "_operator", "_row", "_rows", "_same", "now",
    "mars_bp", "LINE", "CAT_CODE", "CAT_FULL", "UNITS", "SPLIT_BY_UNIT", "EIP_PO_RE", "cat_code", "label_text",
]

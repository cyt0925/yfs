"""共用的東西：Blueprint、常數、小工具、訂單查詢與篩選。其他模組一律 `from .common import *`。"""
import collections
import datetime as _dt
import io
import json
import os
import re
import openpyxl
from flask import (Blueprint, current_app, jsonify, render_template, request,
                   send_file, session)
import db
from db import get_conn, now
from importer import ImportError_, parse_workbook
from normalize import norm_date, norm_decimal, norm_int, norm_key, norm_text


master_bp = Blueprint("master", __name__)


UNCLASSIFIED = "未分類"


LINE_GROUPS_DEFAULT = {
    "rules": [{"prefix": "CPG", "group": "紙潔"}],   # 原始線別以這個開頭 → 畫面上叫這個名字
    "map": {},                                       # 原始線別 → 顯示名（明確指定，優先於 rules）
}


COUPANG_FIELDS = {
    "line": "線別", "barcode": "國條", "yf_sku": "永豐料號", "brand": "品牌",
    "product_name": "品名", "warehouse": "倉別", "order_type": "訂單類型",
    "unit": "單位", "unit_price": "下單單價", "qty_coupang": "下單數量",
    "qty_file_ship": "整合表出貨數量", "box_size_file": "整合表箱入數",
    "delivery_date_file": "整合表交貨日",
}


@master_bp.before_request
def _guard_ready():
    """資料表初始化失敗時（見 db.init_db），這個模組整個停用、講清楚原因，
    不要讓使用者看到一堆零散的資料庫錯誤。"""
    if not getattr(db, "MASTER_READY", False):
        msg = f"商品主檔自動化目前無法使用：{getattr(db, 'MASTER_ERROR', '')}。訂單管理不受影響。"
        if request.path.startswith("/api/"):
            return jsonify({"error": msg}), 503
        return f"<h2 style='font-family:sans-serif;padding:40px'>{msg}</h2>", 503


def _operator():
    return session.get("user", "")


def _is_admin():
    """管理員名單由訂單管理系統維護（設定 → 帳號），這裡直接借用，不另開一套。
    app.py 啟動時就 import 了本模組，所以這裡不能在檔頭 import app（會繞圈），
    等到真的被呼叫時再拿即可。"""
    import app as _app
    return _app.is_admin()


def _rows(cur):
    return [dict(r) for r in cur.fetchall()]


def _row(cur):
    r = cur.fetchone()
    return dict(r) if r is not None else None


def _log(conn, line, po, sku, barcode, field, label, old, new, operator, source, note=""):
    conn.execute(
        """INSERT INTO mst_logs
           (line, po_number, sku_id, barcode, field, field_label, old_value,
            new_value, operator, source, note, changed_at)
           VALUES (?,?,?,?,?,?,?,?,?,?,?,?)""",
        (line or "", po or "", sku or "", barcode or "", field, label,
         "" if old is None else str(old), "" if new is None else str(new),
         operator, source, note, now()))


def _this_month():
    return _dt.date.today().strftime("%Y-%m")


def _valid_month(text):
    return bool(re.fullmatch(r"\d{4}-\d{2}", text or ""))


def _cases(qty, box):
    """出貨數量 ÷ 箱入數。箱入數缺或 0 回 None（畫面標紅），不硬塞 0——
    0 是「這天沒出貨」，跟「算不出來」是兩回事。"""
    if qty is None or not box:
        return None
    return round(qty / box, 4)


def _same(a, b):
    if a is None or b is None:
        return a is None and b is None
    return str(a) == str(b)


def _split_lines(text):
    return [x for x in (text or "").split(",") if x]


def _short(v):
    """變動說明用的短寫：日期 2026-09-04 → 9/4、None → 空、數字去掉多餘的 .0。"""
    if v is None or v == "":
        return "空"
    t = str(v)
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", t):
        return f"{int(t[5:7])}/{int(t[8:10])}"
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return t


def _md(date_str):
    try:
        d = _dt.date.fromisoformat(date_str)
        return f"{d.month}/{d.day}交貨"
    except (ValueError, TypeError):
        return date_str or ""


def _line_groups():
    """線別的顯示規則寫死，不做設定面板（做過，同事看不懂，拿掉了）。
    要改規則直接改 LINE_GROUPS_DEFAULT。"""
    return LINE_GROUPS_DEFAULT


def _group_of(raw, cfg=None):
    """檔案裡的原始線別 → 畫面上的線別名。CPG-潔品／CPG-紙品 → 紙潔；空白 → 未分類；
    其他（寶僑、瑪氏…）照原名。"""
    raw = norm_text(raw)
    if not raw:
        return UNCLASSIFIED
    cfg = cfg or _line_groups()
    if raw in cfg["map"]:
        return cfg["map"][raw] or raw
    for rule in cfg["rules"]:
        if rule.get("prefix") and raw.upper().startswith(str(rule["prefix"]).upper()):
            return rule.get("group") or raw
    return raw


def _raw_lines_seen(conn):
    seen = set()
    for r in conn.execute("SELECT DISTINCT line FROM mst_orders").fetchall():
        seen.add(r["line"] or "")
    for r in conn.execute("SELECT lines_seen FROM mst_products WHERE lines_seen != ''").fetchall():
        seen.update(_split_lines(r["lines_seen"]))
    return sorted(seen)


def _products_map(conn):
    return {p["barcode"]: p for p in _rows(conn.execute("SELECT * FROM mst_products"))}


def _effective_box(order, product):
    if product and product.get("box_size"):
        return product["box_size"], "master"
    if order.get("box_size_file"):
        return order["box_size_file"], "file"
    return None, "none"


_ORDER_SELECT = """SELECT o.*, p.box_size AS box_size_master, p.note AS note_master,
                          p.product_name AS name_master, p.category AS category,
                          p.cost_price AS cost_price, p.yf_sku AS yf_sku_master,
                          p.brand AS brand_master, p.id AS product_id
                   FROM mst_orders o
                   LEFT JOIN mst_products p ON p.barcode = o.barcode"""


def _decorate(o, cfg=None):
    product = {"box_size": o.get("box_size_master")} if o.get("box_size_master") else None
    box, source = _effective_box(o, product)
    o["box_size"] = box
    o["box_source"] = source
    o["cases"] = _cases(o.get("qty_ship"), box)
    o["cases_file"] = _cases(o.get("qty_file_ship"), box)
    o["month"] = (o.get("delivery_date") or "")[:7]
    o["in_master"] = bool(o.get("product_id"))
    o["line_group"] = _group_of(o.get("line"), cfg)
    op_note = o.get("remarks") if o.get("remarks_overridden") else ""
    o["export_note"] = "；".join(x for x in (o.get("note_master") or "", op_note or "") if x)
    return o


def _month_orders(conn, month, cfg=None):
    """某月（含沒排日期的）全部訂單，帶主檔欄位。month 可以是一個月或一串月份。"""
    months = [month] if isinstance(month, str) else list(month or [])
    where = "1=1"
    params = []
    if months:
        likes = " OR ".join(["o.delivery_date LIKE ?"] * len(months))
        where = f"({likes} OR o.delivery_date = '' OR o.delivery_date IS NULL)"
        params = [m + "%" for m in months]
    sql = f"{_ORDER_SELECT} WHERE {where} ORDER BY o.delivery_date, o.po_number, o.sku_id"
    return [_decorate(o, cfg) for o in _rows(conn.execute(sql, params))]


def _month_span(month_from, month_to):
    """2026-07 ～ 2026-09 → ['2026-07', '2026-08', '2026-09']。起訖顛倒就自動對調。"""
    a, b = sorted([month_from, month_to])
    y, m = int(a[:4]), int(a[5:7])
    out = []
    while f"{y:04d}-{m:02d}" <= b and len(out) < 36:
        out.append(f"{y:04d}-{m:02d}")
        m += 1
        if m > 12:
            y, m = y + 1, 1
    return out


def _read_filters(args):
    split = lambda k: [x for x in (args.get(k) or "").split(",") if x]  # noqa: E731
    return {"lines": split("lines"), "dates": split("dates"), "pos": split("pos"),
            "brands": split("brands"), "warehouses": split("warehouses"),
            "q": norm_text(args.get("q")), "edited": args.get("edited") == "1",
            "missing": args.get("missing") == "1",
            "batch": norm_int(args.get("batch")), "batch_scope": (args.get("batch_scope") or "all")}


def _keep(o, f):
    if f["lines"] and o["line_group"] not in f["lines"]:
        return False
    if f["dates"] and (o["delivery_date"] or "") not in f["dates"]:
        return False
    if f["pos"] and o["po_number"] not in f["pos"]:
        return False
    if f["brands"] and o["brand"] not in f["brands"]:
        return False
    if f["warehouses"] and (o["warehouse"] or "") not in f["warehouses"]:
        return False
    if f["edited"] and not (o["qty_ship_overridden"] or o["delivery_date_overridden"] or o["remarks_overridden"]):
        return False
    if f["missing"] and o["cases"] is not None:
        return False
    if f.get("batch"):
        # 篩「某一次匯入」：new = 這批帶進來的；all = 這批帶進來或最後被這批改到的
        b = f["batch"]
        if f.get("batch_scope") == "all":
            if o.get("first_batch_id") != b and o.get("last_batch_id") != b:
                return False
        elif o.get("first_batch_id") != b:
            return False
    if f["q"]:
        hay = " ".join(str(o.get(k) or "") for k in
                       ("po_number", "sku_id", "barcode", "yf_sku", "brand", "product_name", "remarks", "line"))
        if f["q"].lower() not in hay.lower():
            return False
    return True


def _xlsx_response(wb, fname):
    out = io.BytesIO(); wb.save(out); out.seek(0)
    return send_file(out, as_attachment=True, download_name=fname,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


__all__ = [
    "master_bp",
    "UNCLASSIFIED",
    "LINE_GROUPS_DEFAULT",
    "COUPANG_FIELDS",
    "_guard_ready",
    "_operator",
    "_is_admin",
    "_rows",
    "_row",
    "_log",
    "_this_month",
    "_valid_month",
    "_cases",
    "_same",
    "_split_lines",
    "_short",
    "_md",
    "_line_groups",
    "_group_of",
    "_raw_lines_seen",
    "_products_map",
    "_effective_box",
    "_ORDER_SELECT",
    "_decorate",
    "_month_orders",
    "_month_span",
    "_read_filters",
    "_keep",
    "_xlsx_response",
    "collections",
    "_dt",
    "io",
    "json",
    "os",
    "re",
    "openpyxl",
    "Blueprint",
    "current_app",
    "jsonify",
    "render_template",
    "request",
    "send_file",
    "session",
    "db",
    "get_conn",
    "now",
    "ImportError_",
    "parse_workbook",
    "norm_date",
    "norm_decimal",
    "norm_int",
    "norm_key",
    "norm_text",
]

"""商品主檔自動化：取代「訂單彙總表 → 專案報價檔 → 總表」那條人工流程。

OP 原本的做法（見 Word 需求文件）：
  1. 從凱特系統下載「訂單彙總表」（就是 OMS 吃的那份整合表）。
  2. 貼進「專案報價檔」，依交貨日一個分頁一個分頁放，不同 PO 用黃底隔開；
     永豐料號／品類／品牌／箱入數／單價 用 VLOOKUP 從總表撈。
  3. 訂單數量、交貨日一改，報價檔就要跟著手改。
  4. 改完重做樞紐：國條 × 出貨數量(箱)，算出每個交貨日的出貨總箱數。
  5. 再 VLOOKUP 回填「總表」對應日期欄位（要貼成值）。

這個模組的三步：
  ① 商品主檔（mst_products）：全公司一份、鍵是國條。從總表匯入品類、品牌、
     永豐料號、箱入數、COGS 單價、Note——就是同事原本 VLOOKUP 總表那幾欄。
     線別不是商品的屬性（總表沒有線別欄），lines_seen 是匯入訂單時自動學來的
     「這個國條出現過哪些線別」。
  ② 訂單明細（mst_orders）：鍵是 (PO, SKU)。線別是每一列自己帶的原始值，同一張
     PO 可以跨線別（CPG-潔品／CPG-紙品）；改交貨日整張 PO 一起搬、不分線別。
     人改過的欄位（出貨數量、交貨日、備註）立旗標，再匯入不覆蓋。
  ③ 總表：不存快照，每次現算。箱數 = 出貨數量 ÷ 箱入數，依交貨日 + 國條加總。

線別顯示名：檔案裡的原始線別（CPG-潔品、CPG-紙品）照實保留在每一列，畫面上
篩選、總表、匯出用顯示名（CPG-* → 紙潔、空白 → 未分類、其他照原名）。規則寫死在
LINE_GROUPS_DEFAULT，不做設定面板。
"""
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


# ---------------------------------------------------------------- 小工具

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


def _md(date_str):
    try:
        d = _dt.date.fromisoformat(date_str)
        return f"{d.month}/{d.day}交貨"
    except (ValueError, TypeError):
        return date_str or ""


# ---------------------------------------------------------------- 線別顯示名

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


# ---------------------------------------------------------------- 頁面

@master_bp.route("/master")
def master_page():
    logo = ("logo_master.png"
            if os.path.exists(os.path.join(os.path.dirname(__file__), "static", "logo_master.png"))
            else "logo.png")
    return render_template("master.html", logo_file=logo, logged_in_user=_operator(),
                           is_admin=_is_admin(),
                           build_version=current_app.config.get("BUILD_VERSION", ""))


@master_bp.route("/api/master/lines")
def api_lines():
    cfg = _line_groups()
    conn = get_conn()
    try:
        raws = _raw_lines_seen(conn)
        months = [r["m"] for r in conn.execute(
            "SELECT DISTINCT substr(delivery_date, 1, 7) AS m FROM mst_orders "
            "WHERE delivery_date != '' ORDER BY m DESC").fetchall()]
    finally:
        conn.close()
    groups = sorted({_group_of(r, cfg) for r in raws})
    return jsonify({"groups": groups, "months": months, "this_month": _this_month()})


# ---------------------------------------------------------------- 主檔查詢（共用）

def _products_map(conn):
    return {p["barcode"]: p for p in _rows(conn.execute("SELECT * FROM mst_products"))}


def _effective_box(order, product):
    if product and product.get("box_size"):
        return product["box_size"], "master"
    if order.get("box_size_file"):
        return order["box_size_file"], "file"
    return None, "none"


# ---------------------------------------------------------------- 匯入（兩階段）

def _diff_import(conn, rows):
    products = _products_map(conn)
    new, updated, identical, missing = [], [], [], {}
    for r in rows:
        if r["barcode"] and r["barcode"] not in products:
            missing.setdefault(r["barcode"], {
                "barcode": r["barcode"], "sku_id": r["sku_id"], "yf_sku": r["yf_sku"],
                "brand": r["brand"], "product_name": r["product_name"],
                "box_size": r["box_size"], "line": r["line"] or ""})
        existing = _row(conn.execute(
            "SELECT * FROM mst_orders WHERE po_number = ? AND sku_id = ?",
            (r["po_number"], r["sku_id"])))
        if existing is None:
            new.append(r); continue
        changes = []
        mapping = {
            "line": r["line"] or "", "barcode": r["barcode"], "yf_sku": r["yf_sku"],
            "brand": r["brand"], "product_name": r["product_name"], "warehouse": r["warehouse"],
            "order_type": r["order_type"], "unit": r["unit"], "unit_price": r["unit_price"],
            "qty_coupang": r["qty_coupang"], "qty_file_ship": r["qty_file_ship"],
            "box_size_file": r["box_size"], "delivery_date_file": r["delivery_date"],
        }
        for field, value in mapping.items():
            if not _same(existing.get(field), value):
                changes.append({"field": field, "label": COUPANG_FIELDS[field],
                                "old": existing.get(field), "new": value})
        if existing["missing_in_file"]:
            changes.append({"field": "missing_in_file", "label": "品項重新出現",
                            "old": "檔案已無此品項", "new": "恢復"})
        if changes:
            updated.append({
                "existing_id": existing["id"], "changes": changes,
                "po_number": r["po_number"], "sku_id": r["sku_id"], "barcode": r["barcode"],
                "product_name": r["product_name"], "line": r["line"] or "",
                "qty_ship_overridden": existing["qty_ship_overridden"],
                "delivery_date_overridden": existing["delivery_date_overridden"],
            })
        else:
            identical.append(r)

    # 酷澎把品項數量下修到 0 時，匯出的表那一列直接消失。這張 PO 有在檔案裡、
    # 底下某個品項卻沒帶到 → 判定被拿掉：不刪列、出貨歸 0、標記。
    file_keys = {(r["po_number"], r["sku_id"]) for r in rows}
    removed = []
    for po in sorted({r["po_number"] for r in rows}):
        for ex in _rows(conn.execute("SELECT * FROM mst_orders WHERE po_number = ?", (po,))):
            if (po, ex["sku_id"]) in file_keys or ex["missing_in_file"]:
                continue
            removed.append({"id": ex["id"], "po_number": po, "sku_id": ex["sku_id"],
                            "barcode": ex["barcode"], "product_name": ex["product_name"],
                            "line": ex["line"], "qty_coupang": ex["qty_coupang"],
                            "qty_ship": ex["qty_ship"], "delivery_date": ex["delivery_date"]})
    return new, updated, identical, list(missing.values()), removed


def _parse_uploads(uploads):
    """把多份上傳檔各自 parse_workbook 再合併。回傳 (rows, warnings)；任一份壞掉
    就回傳 ((錯誤訊息, 400), None)，錯誤訊息點名是哪一份壞的，整批都不收。"""
    merged, warnings, seen = [], [], {}
    many = len(uploads) > 1
    for up in uploads:
        try:
            rows, warns = parse_workbook(io.BytesIO(up.read()), up.filename)
        except ImportError_ as exc:
            msg = f"「{up.filename}」：{exc}" if many else str(exc)
            if many:
                msg += "　（有檔案解析失敗，這次全部都沒匯入；修好再一起丟，或分開丟。）"
            return (msg, 400), None
        for w in warns:
            warnings.append(f"「{up.filename}」{w}" if many else w)
        dup = 0
        for r in rows:
            key = (r["po_number"], r["sku_id"])
            if key in seen:
                dup += 1
                merged = [x for x in merged if (x["po_number"], x["sku_id"]) != key]
            seen[key] = up.filename
            r["source_file"] = up.filename
            merged.append(r)
        if dup:
            warnings.append(f"「{up.filename}」有 {dup} 筆 PO／SKU 跟前面的檔案重複，以這份為準。")
    return merged, warnings


@master_bp.route("/api/master/import/preview", methods=["POST"])
def api_import_preview():
    """一次可以丟好幾份訂單彙總表（例如 7 月一份、8 月一份）。每份各自解析後
    合成一批預覽；跨檔案撞到同一個 (PO, SKU) 時以後面那份為準並提醒，跟單一
    檔案內重複的處理方式一致。"""
    uploads = [u for u in request.files.getlist("file") if u is not None and u.filename]
    if not uploads:
        return jsonify({"error": "沒有收到檔案。"}), 400
    rows, warnings = _parse_uploads(uploads)
    if isinstance(rows, tuple):          # (error message, status)
        return jsonify({"error": rows[0]}), rows[1]
    filename = "、".join(u.filename for u in uploads)
    cfg = _line_groups()
    conn = get_conn()
    try:
        new, updated, identical, missing, removed = _diff_import(conn, rows)
        payload = {"rows": rows, "missing_products": missing, "removed": removed,
                   "filename": filename}
        cur = conn.execute(
            """INSERT INTO mst_import_batches
               (filename, operator, rows_total, rows_new, rows_updated, rows_identical,
                committed, payload_json, created_at) VALUES (?,?,?,?,?,?,0,?,?)""",
            (filename, _operator(), len(rows), len(new), len(updated), len(identical),
             json.dumps(payload, ensure_ascii=False), now()))
        conn.commit()
        batch_id = cur.lastrowid
    finally:
        conn.close()

    by_raw = collections.Counter(r["line"] or "" for r in rows)
    by_group = collections.Counter(_group_of(r["line"], cfg) for r in rows)
    dates = sorted({r["delivery_date"] for r in rows if r["delivery_date"]})
    warnings = list(warnings)
    if by_raw.get(""):
        warnings.append(f"有 {by_raw['']} 筆沒有「線別」，歸到「{UNCLASSIFIED}」，畫面上會標紅；"
                        "請在訂單明細裡確認這幾筆。")
    return jsonify({
        "batch_id": batch_id, "filename": filename, "files": [u.filename for u in uploads],
        "rows_total": len(rows),
        "new_count": len(new), "updated_count": len(updated), "identical_count": len(identical),
        "lines_raw": dict(by_raw), "lines": dict(by_group), "dates": dates,
        "po_count": len({r["po_number"] for r in rows}), "warnings": warnings,
        "new_preview": new[:300], "updated": updated[:300],
        "missing_products": missing, "removed": removed, "removed_count": len(removed),
    })


@master_bp.route("/api/master/import/commit", methods=["POST"])
def api_import_commit():
    payload = request.get_json(silent=True) or {}
    batch_id = norm_int(payload.get("batch_id"))
    add_missing = payload.get("add_missing_products", True)
    operator = _operator()
    if batch_id is None:
        return jsonify({"error": "缺少 batch_id。"}), 400
    conn = get_conn()
    try:
        batch = _row(conn.execute("SELECT * FROM mst_import_batches WHERE id = ?", (batch_id,)))
        if batch is None:
            return jsonify({"error": "找不到這次的預覽，請重新上傳。"}), 404
        if batch["committed"]:
            return jsonify({"error": "這批已經匯入過了，不能重複確認。"}), 409
        data = json.loads(batch["payload_json"] or "{}")
        rows = data.get("rows", [])
        fname = data.get("filename", "")
        stamp = now()
        inserted = updated_n = identical_n = products_added = removed_n = 0

        if add_missing:
            for m in data.get("missing_products", []):
                if _row(conn.execute("SELECT id FROM mst_products WHERE barcode = ?", (m["barcode"],))):
                    continue
                conn.execute(
                    """INSERT INTO mst_products
                       (barcode, sku_id, yf_sku, brand, product_name, box_size, note,
                        auto_created, lines_seen, updated_by, updated_at)
                       VALUES (?,?,?,?,?,?,'',1,?,?,?)""",
                    (m["barcode"], m.get("sku_id") or "", m.get("yf_sku") or "", m.get("brand") or "",
                     m.get("product_name") or "", m.get("box_size"), m.get("line") or "", operator, stamp))
                _log(conn, m.get("line"), "", m.get("sku_id"), m["barcode"], "product_new", "新增主檔",
                     "", m.get("box_size"), operator, "import", fname)
                products_added += 1

        for rm in data.get("removed", []):
            ex = _row(conn.execute("SELECT * FROM mst_orders WHERE id = ?", (rm["id"],)))
            if ex is None or ex["missing_in_file"]:
                continue
            conn.execute(
                """UPDATE mst_orders SET qty_ship = 0, missing_in_file = 1, last_seen_at = ?,
                   updated_at = ?, last_batch_id = ?, version = version + 1 WHERE id = ?""",
                (stamp, stamp, batch_id, ex["id"]))
            _log(conn, ex["line"], ex["po_number"], ex["sku_id"], ex["barcode"], "qty_ship", "出貨數量",
                 ex["qty_ship"], 0, operator, "import", "這次的整合表裡這張 PO 已沒有這個品項，出貨數量歸 0")
            removed_n += 1

        # 主檔學線別：這個國條出現在哪些原始線別
        seen_by_bc = collections.defaultdict(set)
        for r in rows:
            if r["barcode"]:
                seen_by_bc[r["barcode"]].add(r["line"] or "")

        for r in rows:
            line = r["line"] or ""
            existing = _row(conn.execute(
                "SELECT * FROM mst_orders WHERE po_number = ? AND sku_id = ?", (r["po_number"], r["sku_id"])))
            ship = r["qty_file_ship"] if r["qty_file_ship"] is not None else r["qty_coupang"]
            if existing is None:
                conn.execute(
                    """INSERT INTO mst_orders
                       (po_number, sku_id, line, barcode, yf_sku, brand, product_name, warehouse,
                        order_type, unit, unit_price, qty_coupang, qty_file_ship, qty_ship,
                        box_size_file, delivery_date_file, delivery_date, remarks_file, remarks,
                        source_file, first_seen_at, last_seen_at, updated_at, version,
                        first_batch_id, last_batch_id)
                       VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,1,?,?)""",
                    (r["po_number"], r["sku_id"], line, r["barcode"], r["yf_sku"], r["brand"],
                     r["product_name"], r["warehouse"], r["order_type"], r["unit"], r["unit_price"],
                     r["qty_coupang"], r["qty_file_ship"], ship, r["box_size"], r["delivery_date"],
                     r["delivery_date"], r["remarks_file"], "", r["source_file"], stamp, stamp, stamp,
                     batch_id, batch_id))
                inserted += 1
                continue
            sets, vals, changed = [], [], False
            mapping = {
                "line": line, "barcode": r["barcode"], "yf_sku": r["yf_sku"], "brand": r["brand"],
                "product_name": r["product_name"], "warehouse": r["warehouse"],
                "order_type": r["order_type"], "unit": r["unit"], "unit_price": r["unit_price"],
                "qty_coupang": r["qty_coupang"], "qty_file_ship": r["qty_file_ship"],
                "box_size_file": r["box_size"], "delivery_date_file": r["delivery_date"],
                "remarks_file": r["remarks_file"],
            }
            for field, value in mapping.items():
                if not _same(existing.get(field), value):
                    sets.append(f"{field} = ?"); vals.append(value); changed = True
                    if field in COUPANG_FIELDS:
                        _log(conn, line, r["po_number"], r["sku_id"], r["barcode"], field,
                             COUPANG_FIELDS[field], existing.get(field), value, operator, "import", fname)
            if not existing["qty_ship_overridden"] and not _same(existing["qty_ship"], ship):
                sets.append("qty_ship = ?"); vals.append(ship); changed = True
                _log(conn, line, r["po_number"], r["sku_id"], r["barcode"], "qty_ship", "出貨數量",
                     existing["qty_ship"], ship, operator, "import", "隨整合表更新")
            if not existing["delivery_date_overridden"] and not _same(existing["delivery_date"], r["delivery_date"]):
                sets.append("delivery_date = ?"); vals.append(r["delivery_date"]); changed = True
                _log(conn, line, r["po_number"], r["sku_id"], r["barcode"], "delivery_date", "交貨日",
                     existing["delivery_date"], r["delivery_date"], operator, "import", "隨整合表更新")
            if existing["missing_in_file"]:
                sets.append("missing_in_file = 0"); changed = True
                if not existing["qty_ship_overridden"] and "qty_ship = ?" not in sets:
                    sets.append("qty_ship = ?"); vals.append(ship)
                _log(conn, line, r["po_number"], r["sku_id"], r["barcode"], "missing_in_file",
                     "品項重新出現", 1, 0, operator, "import", "整合表裡又有這個品項了")
            sets.append("last_seen_at = ?"); vals.append(stamp)
            sets.append("source_file = ?"); vals.append(r["source_file"])
            if changed:
                sets.append("updated_at = ?"); vals.append(stamp)
                sets.append("last_batch_id = ?"); vals.append(batch_id)
                sets.append("version = version + 1"); updated_n += 1
            else:
                identical_n += 1
            conn.execute(f"UPDATE mst_orders SET {', '.join(sets)} WHERE id = ?", vals + [existing["id"]])

        for bc, lines in seen_by_bc.items():
            p = _row(conn.execute("SELECT id, lines_seen FROM mst_products WHERE barcode = ?", (bc,)))
            if p is None:
                continue
            merged = sorted(set(_split_lines(p["lines_seen"])) | {l for l in lines if l})
            if ",".join(merged) != (p["lines_seen"] or ""):
                conn.execute("UPDATE mst_products SET lines_seen = ? WHERE id = ?", (",".join(merged), p["id"]))

        conn.execute(
            """UPDATE mst_import_batches SET committed = 1, committed_at = ?,
               rows_new = ?, rows_updated = ?, rows_identical = ? WHERE id = ?""",
            (stamp, inserted, updated_n, identical_n, batch_id))
        conn.commit()
    except Exception:
        conn.rollback(); raise
    finally:
        conn.close()
    return jsonify({"ok": True, "inserted": inserted, "updated": updated_n, "identical": identical_n,
                    "products_added": products_added, "removed": removed_n})


# ---------------------------------------------------------------- 訂單明細

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
            "batch": norm_int(args.get("batch")), "batch_scope": (args.get("batch_scope") or "new")}


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


@master_bp.route("/api/master/orders")
def api_orders():
    month = norm_text(request.args.get("month")) or _this_month()
    if not _valid_month(month):
        return jsonify({"error": "月份格式要像 2026-09。"}), 400
    filters = _read_filters(request.args)
    cfg = _line_groups()
    conn = get_conn()
    try:
        rows = _month_orders(conn, month, cfg)
    finally:
        conn.close()

    # 篩選面在「月份」範圔上算、線別除外也一樣，勾了什麼其他選項不會消失
    f_lines = collections.Counter(); f_dates = collections.OrderedDict()
    f_pos = collections.OrderedDict(); f_brands = collections.Counter(); f_wh = collections.Counter()
    for o in rows:
        f_lines[o["line_group"]] += 1
        f_wh[o["warehouse"] or ""] += 1
        if o["brand"]:
            f_brands[o["brand"]] += 1
        d = o["delivery_date"] or ""
        fd = f_dates.setdefault(d, {"date": d, "cases": 0.0, "rows": 0, "pos": set(), "missing_box": 0})
        fd["rows"] += 1; fd["pos"].add(o["po_number"])
        if o["cases"] is None:
            fd["missing_box"] += 1
        else:
            fd["cases"] += o["cases"]
        fp = f_pos.setdefault(o["po_number"], {"po_number": o["po_number"], "date": d, "rows": 0,
                                                "cases": 0.0, "warehouse": o["warehouse"],
                                                "lines": set(), "date_overridden": 0})
        fp["rows"] += 1; fp["cases"] += o["cases"] or 0; fp["lines"].add(o["line_group"])
        fp["date_overridden"] = fp["date_overridden"] or o["delivery_date_overridden"]
    for fd in f_dates.values():
        fd["po_count"] = len(fd.pop("pos")); fd["cases"] = round(fd["cases"], 2)
    for fp in f_pos.values():
        fp["cases"] = round(fp["cases"], 2); fp["lines"] = sorted(fp["lines"])

    shown = [o for o in rows if _keep(o, filters)]
    return jsonify({
        "month": month, "rows": shown, "count": len(shown),
        "total_cases": round(sum(o["cases"] or 0 for o in shown), 2),
        "facets": {
            "lines": [{"line": l, "rows": n} for l, n in sorted(f_lines.items(), key=lambda x: (x[0] == UNCLASSIFIED, x[0]))],
            "dates": list(f_dates.values()), "pos": list(f_pos.values()),
            "brands": [{"brand": b, "rows": n} for b, n in sorted(f_brands.items())],
            "warehouses": [{"warehouse": w, "rows": n} for w, n in sorted(f_wh.items())],
        },
    })


def _apply_item_edit(conn, o, payload, operator):
    """單一品項的編輯（出貨數量、備註、恢復整合表數字）。回傳 (sets, vals, changed)。"""
    sets, vals, changed = [], [], 0
    if "qty_ship" in payload and payload.get("qty_ship") not in (None, ""):
        new = norm_int(payload.get("qty_ship"))
        if new is None or new < 0:
            raise ValueError("出貨數量要是 0 或正整數。")
        if not _same(o["qty_ship"], new):
            sets += ["qty_ship = ?", "qty_ship_overridden = 1"]; vals.append(new)
            _log(conn, o["line"], o["po_number"], o["sku_id"], o["barcode"], "qty_ship", "出貨數量",
                 o["qty_ship"], new, operator, "manual")
            changed += 1
    if "remarks" in payload:
        new = norm_text(payload.get("remarks"))
        if not _same(o["remarks"], new):
            sets += ["remarks = ?", "remarks_overridden = ?"]; vals += [new, 1 if new else 0]
            _log(conn, o["line"], o["po_number"], o["sku_id"], o["barcode"], "remarks", "備註",
                 o["remarks"], new, operator, "manual")
            changed += 1
    if payload.get("reset_qty_ship"):
        file_ship = o["qty_file_ship"] if o["qty_file_ship"] is not None else o["qty_coupang"]
        sets += ["qty_ship = ?", "qty_ship_overridden = 0"]; vals.append(file_ship)
        _log(conn, o["line"], o["po_number"], o["sku_id"], o["barcode"], "qty_ship", "出貨數量",
             o["qty_ship"], file_ship, operator, "manual", "恢復為整合表數字")
        changed += 1
    return sets, vals, changed


@master_bp.route("/api/master/orders/<int:order_id>", methods=["PUT"])
def api_update_order(order_id):
    payload = request.get_json(silent=True) or {}
    operator = _operator()
    client_version = norm_int(payload.get("version"))
    if client_version is None:
        return jsonify({"error": "缺少版本資訊，請重新整理後再試。"}), 400
    conn = get_conn()
    try:
        o = _row(conn.execute("SELECT * FROM mst_orders WHERE id = ?", (order_id,)))
        if o is None:
            return jsonify({"error": "找不到這筆。"}), 404
        if o["version"] != client_version:
            return jsonify({"error": "conflict",
                            "message": f"這筆剛被另一次儲存修改過（{o['updated_at']}），請重新載入後再編輯。"}), 409
        try:
            sets, vals, changed = _apply_item_edit(conn, o, payload, operator)
        except ValueError as exc:
            conn.rollback(); return jsonify({"error": str(exc)}), 400
        if not changed:
            conn.rollback()
            return jsonify({"ok": True, "changed": 0})
        sets += ["updated_at = ?", "version = version + 1"]; vals.append(now())
        cur = conn.execute(f"UPDATE mst_orders SET {', '.join(sets)} WHERE id = ? AND version = ?",
                           vals + [order_id, client_version])
        if cur.rowcount == 0:
            conn.rollback()
            return jsonify({"error": "conflict", "message": "儲存瞬間有其他人也改了這筆，請重新載入。"}), 409
        conn.commit()
        fresh = _decorate(_row(conn.execute(f"{_ORDER_SELECT} WHERE o.id = ?", (order_id,))))
        return jsonify({"ok": True, "changed": changed, "row": fresh})
    finally:
        conn.close()


@master_bp.route("/api/master/orders/<int:order_id>", methods=["DELETE"])
def api_delete_order(order_id):
    conn = get_conn()
    try:
        o = _row(conn.execute("SELECT * FROM mst_orders WHERE id = ?", (order_id,)))
        if o is None:
            return jsonify({"error": "找不到這筆。"}), 404
        conn.execute("DELETE FROM mst_orders WHERE id = ?", (order_id,))
        _log(conn, o["line"], o["po_number"], o["sku_id"], o["barcode"], "delete", "刪除品項",
             f"出貨 {o['qty_ship']} / 交貨 {o['delivery_date']}", "", _operator(), "manual")
        conn.commit()
        return jsonify({"ok": True})
    finally:
        conn.close()


# ---------------------------------------------------------------- PO：改期（可批次）、明細視窗、整張儲存

def _move_po_dates(conn, pos, new_date, reset, expected, operator, note):
    """整張 PO（跨線別的全部品項）搬到新日期。expected = {po: 畫面上看到的舊日期}，
    對不上就是有人先改過 → 丟 conflict。回傳搬了幾個品項。"""
    moved = 0
    for po in pos:
        rows = _rows(conn.execute("SELECT * FROM mst_orders WHERE po_number = ?", (po,)))
        if not rows:
            raise LookupError(f"找不到 PO {po}。")
        if expected and po in expected and expected[po] is not None:
            current = {r["delivery_date"] or "" for r in rows}
            if current != {expected[po] or ""}:
                raise PermissionError(f"PO {po} 的交貨日剛被別人改過，請重新載入後再改。")
        stamp = now()
        for r in rows:
            target = r["delivery_date_file"] if reset else new_date
            flag = 0 if reset else 1
            if _same(r["delivery_date"], target) and r["delivery_date_overridden"] == flag:
                continue
            conn.execute(
                """UPDATE mst_orders SET delivery_date = ?, delivery_date_overridden = ?,
                   updated_at = ?, version = version + 1 WHERE id = ?""", (target, flag, stamp, r["id"]))
            _log(conn, r["line"], po, r["sku_id"], r["barcode"], "delivery_date", "交貨日",
                 r["delivery_date"], target, operator, "manual", note)
            moved += 1
    return moved


@master_bp.route("/api/master/pos/date", methods=["PUT"])
def api_update_po_date():
    payload = request.get_json(silent=True) or {}
    operator = _operator()
    pos = [norm_key(p) for p in (payload.get("po_numbers") or ([payload.get("po_number")] if payload.get("po_number") else [])) if norm_key(p)]
    reset = bool(payload.get("reset"))
    new_date = None if reset else norm_date(payload.get("delivery_date"))
    expected = payload.get("expected") or ({pos[0]: payload.get("expected_date")} if len(pos) == 1 and "expected_date" in payload else {})
    if not pos:
        return jsonify({"error": "缺少 PO 單號。"}), 400
    if not reset and not new_date:
        return jsonify({"error": "交貨日格式不對，請用 2026-09-17 這種寫法。"}), 400
    conn = get_conn()
    try:
        try:
            moved = _move_po_dates(conn, pos, new_date, reset, expected, operator,
                                   "恢復為整合表日期" if reset else ("批次改期" if len(pos) > 1 else "整張 PO 改期"))
        except LookupError as exc:
            conn.rollback(); return jsonify({"error": str(exc)}), 404
        except PermissionError as exc:
            conn.rollback(); return jsonify({"error": "conflict", "message": str(exc)}), 409
        conn.commit()
        return jsonify({"ok": True, "moved": moved, "po_count": len(pos),
                        "delivery_date": new_date if not reset else None})
    finally:
        conn.close()


@master_bp.route("/api/master/pos/<po>")
def api_po_detail(po):
    po = norm_key(po)
    cfg = _line_groups()
    conn = get_conn()
    try:
        rows = [_decorate(o, cfg) for o in _rows(conn.execute(
            f"{_ORDER_SELECT} WHERE o.po_number = ? ORDER BY o.line, o.sku_id", (po,)))]
        if not rows:
            return jsonify({"error": "找不到這張 PO。"}), 404
        logs = _rows(conn.execute(
            "SELECT * FROM mst_logs WHERE po_number = ? ORDER BY changed_at DESC, id DESC LIMIT 200", (po,)))
    finally:
        conn.close()
    dates = sorted({r["delivery_date"] or "" for r in rows})
    return jsonify({
        "po_number": po, "rows": rows, "logs": logs,
        "delivery_date": dates[0] if len(dates) == 1 else "",
        "dates": dates, "delivery_date_file": rows[0]["delivery_date_file"],
        "date_overridden": any(r["delivery_date_overridden"] for r in rows),
        "lines": sorted({r["line_group"] for r in rows}),
        "lines_raw": sorted({r["line"] or "" for r in rows}),
        "warehouse": rows[0]["warehouse"],
        "total_cases": round(sum(r["cases"] or 0 for r in rows), 2),
    })


@master_bp.route("/api/master/pos/<po>/save", methods=["PUT"])
def api_po_save(po):
    """PO 視窗一次儲存：交貨日（整張）＋各品項的出貨數量／備註。任一筆版本對不上
    整批退回，不會存一半。"""
    po = norm_key(po)
    payload = request.get_json(silent=True) or {}
    operator = _operator()
    items = payload.get("items") or []
    conn = get_conn()
    try:
        changed_total = 0
        if payload.get("delivery_date") or payload.get("reset_date"):
            new_date = norm_date(payload.get("delivery_date")) if not payload.get("reset_date") else None
            if not payload.get("reset_date") and not new_date:
                return jsonify({"error": "交貨日格式不對。"}), 400
            try:
                changed_total += _move_po_dates(conn, [po], new_date, bool(payload.get("reset_date")),
                                                {po: payload.get("expected_date")}, operator,
                                                "PO 視窗改期" if not payload.get("reset_date") else "恢復為整合表日期")
            except LookupError as exc:
                conn.rollback(); return jsonify({"error": str(exc)}), 404
            except PermissionError as exc:
                conn.rollback(); return jsonify({"error": "conflict", "message": str(exc)}), 409
        for it in items:
            oid = norm_int(it.get("id")); ver = norm_int(it.get("version"))
            if oid is None or ver is None:
                conn.rollback(); return jsonify({"error": "品項缺少 id 或版本。"}), 400
            o = _row(conn.execute("SELECT * FROM mst_orders WHERE id = ? AND po_number = ?", (oid, po)))
            if o is None:
                conn.rollback(); return jsonify({"error": f"找不到品項 {oid}。"}), 404
            # 改期那步已經把版本 +1，所以這裡用「版本 == 原本 或 原本+1（且是剛剛改期造成）」都不安全；
            # 改成比對 updated_at 之前的版本：改期只動 delivery_date，品項欄位的衝突仍用畫面帶來的版本判。
            if o["version"] not in (ver, ver + 1):
                conn.rollback()
                return jsonify({"error": "conflict",
                                "message": f"品項 {o['sku_id']} 剛被別人改過，請重新載入後再存。"}), 409
            try:
                sets, vals, changed = _apply_item_edit(conn, o, it, operator)
            except ValueError as exc:
                conn.rollback(); return jsonify({"error": str(exc)}), 400
            if changed:
                sets += ["updated_at = ?", "version = version + 1"]; vals.append(now())
                conn.execute(f"UPDATE mst_orders SET {', '.join(sets)} WHERE id = ?", vals + [oid])
                changed_total += changed
        conn.commit()
        return jsonify({"ok": True, "changed": changed_total})
    finally:
        conn.close()


# ---------------------------------------------------------------- 商品主檔（全公司一份）

@master_bp.route("/api/master/products")
def api_products():
    line = norm_text(request.args.get("line"))
    q = norm_text(request.args.get("q")).lower()
    cfg = _line_groups()
    conn = get_conn()
    try:
        rows = _rows(conn.execute(
            """SELECT p.*, (SELECT COUNT(*) FROM mst_orders o WHERE o.barcode = p.barcode) AS order_rows
               FROM mst_products p ORDER BY p.brand, p.product_name, p.barcode"""))
        orphans = _rows(conn.execute(
            """SELECT o.barcode, MAX(o.sku_id) AS sku_id, MAX(o.yf_sku) AS yf_sku, MAX(o.brand) AS brand,
                      MAX(o.product_name) AS product_name, MAX(o.box_size_file) AS box_size_file,
                      MAX(o.line) AS line, COUNT(*) AS order_rows
               FROM mst_orders o LEFT JOIN mst_products p ON p.barcode = o.barcode
               WHERE p.id IS NULL AND o.barcode != '' GROUP BY o.barcode ORDER BY o.barcode"""))
    finally:
        conn.close()
    for p in rows:
        raws = _split_lines(p["lines_seen"])
        p["line_groups"] = sorted({_group_of(r, cfg) for r in raws})
        p["lines_raw"] = raws
    for o in orphans:
        o["line_group"] = _group_of(o["line"], cfg)
    if line:
        rows = [p for p in rows if line in p["line_groups"]]
        orphans = [o for o in orphans if o["line_group"] == line]
    if q:
        rows = [p for p in rows if q in " ".join(str(p.get(k) or "") for k in
                ("barcode", "sku_id", "yf_sku", "brand", "product_name", "note", "category", "lines_seen", "unit", "master_line")).lower()]
    return jsonify({"products": rows, "orphans": orphans, "total": len(rows)})


def _upsert_product(conn, barcode, fields, box, cost, operator, source, fname="", only_filled=False):
    """新增或更新主檔一筆。only_filled=True 時（總表匯入）只更新有值的欄位。"""
    existing = _row(conn.execute("SELECT * FROM mst_products WHERE barcode = ?", (barcode,)))
    stamp = now()
    if cost is not None:
        cost = round(float(cost), 4)   # Excel 存的 137.0000000000 讀出來會是 137.00000000000003
    fields = dict(fields)
    for k in _PRODUCT_TEXT_FIELDS:
        fields.setdefault(k, "")
    fields.setdefault("shelf_days", None)
    master_line = fields.get("master_line") or ""
    if existing is None:
        conn.execute(
            """INSERT INTO mst_products
               (barcode, sku_id, yf_sku, brand, product_name, category, pgcode, cost_price,
                box_size, note, auto_created, lines_seen, master_line, unit, shelf_days, active,
                date_format, updated_by, updated_at)
               VALUES (?,?,?,?,?,?,?,?,?,?,0,?,?,?,?,?,?,?,?)""",
            (barcode, fields["sku_id"], fields["yf_sku"], fields["brand"], fields["product_name"],
             fields["category"], fields["pgcode"], cost, box, fields["note"], master_line, master_line,
             fields["unit"], fields["shelf_days"], fields["active"] or "Y", fields["date_format"],
             operator, stamp))
        _log(conn, master_line, "", fields["sku_id"], barcode, "product_new", "新增主檔", "", box, operator, source, fname)
        return "added"
    sets, params, changed = [], [], False
    if master_line and master_line not in _split_lines(existing.get("lines_seen")):
        merged = sorted(set(_split_lines(existing.get("lines_seen"))) | {master_line})
        sets.append("lines_seen = ?"); params.append(",".join(merged)); changed = True
    for k, v in fields.items():
        if only_filled and (v is None or v == ""):
            continue
        if not _same(existing[k], v):
            sets.append(f"{k} = ?"); params.append(v); changed = True
            if k == "note":
                _log(conn, "", "", existing["sku_id"], barcode, "note", "Note", existing["note"], v,
                     operator, source, fname)
    if (cost is not None or not only_filled) and not _same(existing["cost_price"], cost):
        sets.append("cost_price = ?"); params.append(cost); changed = True
    if (box or not only_filled) and not _same(existing["box_size"], box):
        sets.append("box_size = ?"); params.append(box); changed = True
        _log(conn, "", "", existing["sku_id"], barcode, "box_size", "箱入數", existing["box_size"], box,
             operator, source, fname)
    if existing["auto_created"]:
        sets.append("auto_created = 0"); changed = True
    if changed:
        sets += ["updated_by = ?", "updated_at = ?"]; params += [operator, stamp]
        conn.execute(f"UPDATE mst_products SET {', '.join(sets)} WHERE id = ?", params + [existing["id"]])
        return "updated"
    return "unchanged"


@master_bp.route("/api/master/products", methods=["POST"])
def api_save_product():
    payload = request.get_json(silent=True) or {}
    barcode = norm_key(payload.get("barcode"))
    if not barcode:
        return jsonify({"error": "國條必填。"}), 400
    box = norm_int(payload.get("box_size"))
    if payload.get("box_size") not in (None, "") and (box is None or box <= 0):
        return jsonify({"error": "箱入數要是正整數。"}), 400
    cost = None if payload.get("cost_price") in (None, "") else norm_decimal(payload.get("cost_price"))
    if payload.get("cost_price") not in (None, "") and cost is None:
        return jsonify({"error": "單價(含稅)要是數字。"}), 400
    fields = {k: (norm_key(payload.get(k)) if k in ("sku_id", "yf_sku") else norm_text(payload.get(k)))
              for k in _PRODUCT_TEXT_FIELDS}
    fields["active"] = (fields["active"] or "Y").upper()[:1]
    fields["shelf_days"] = norm_int(payload.get("shelf_days"))
    conn = get_conn()
    try:
        _upsert_product(conn, barcode, fields, box, cost, _operator(), "manual")
        conn.commit()
        fresh = _row(conn.execute("SELECT * FROM mst_products WHERE barcode = ?", (barcode,)))
        return jsonify({"ok": True, "product": fresh})
    finally:
        conn.close()


@master_bp.route("/api/master/products/<int:product_id>", methods=["DELETE"])
def api_delete_product(product_id):
    conn = get_conn()
    try:
        p = _row(conn.execute("SELECT * FROM mst_products WHERE id = ?", (product_id,)))
        if p is None:
            return jsonify({"error": "找不到這筆主檔。"}), 404
        conn.execute("DELETE FROM mst_products WHERE id = ?", (product_id,))
        _log(conn, "", "", p["sku_id"], p["barcode"], "product_delete", "刪除主檔", p["box_size"], "",
             _operator(), "manual")
        conn.commit()
        return jsonify({"ok": True})
    finally:
        conn.close()


# 主檔匯入認兩種檔：
#   ・酷澎主檔（一個線別一份）：SKU ID／條碼(國條)／永豐料號／線別／品牌／品名／單位／箱入數／
#     業務報價單價(含稅)／總效期天數／啟用(Y/N)／日期格式／報價備註
#   ・寶僑總表 Sheet1：skuid／Category／Brand／Pgcode／Barcode／永豐料號／SKU Name／COGS／Note／箱入數
# 都是認標題不認欄位位置。業務報價單價(含稅) = COGS，同一個欄位；報價備註 = Note，同一個欄位。
_HEADER_ALIASES = {
    "barcode": ["barcode", "國條", "條碼", "條碼(國條)", "國際條碼", "ean"],
    "category": ["category", "品類", "類別"],
    "pgcode": ["pgcode", "pg code", "pg_code"],
    "cost_price": ["cogs (pcs/ w. tax)", "cogs", "cogs(pcs)", "業務報價單價(含稅)", "業務報價單價", "報價單價",
                   "單價(含稅)", "成本(含稅)", "cost"],
    "sku_id": ["skuid", "sku id", "sku_id", "skuno.", "sku編號"],
    "yf_sku": ["永豐料號", "料號"],
    "brand": ["brand", "品牌"],
    "product_name": ["sku name", "品名", "product name", "名稱"],
    "box_size": ["箱入數", "箱入", "case pack", "pcs/cs", "轉換率"],
    "note": ["note", "報價備註", "備註", "說明"],
    "master_line": ["線別"],
    "unit": ["單位"],
    "shelf_days": ["總效期天數", "效期天數", "效期"],
    "active": ["啟用(y/n)", "啟用"],
    "date_format": ["日期格式"],
}
_PRODUCT_TEXT_FIELDS = ("sku_id", "yf_sku", "brand", "product_name", "category", "pgcode", "note",
                        "master_line", "unit", "active", "date_format")


def _find_columns(ws, want):
    best = (None, {}, [])
    for idx, row in enumerate(ws.iter_rows(min_row=1, max_row=15, values_only=True), start=1):
        headers = [norm_text(c) for c in row]
        lowered = [h.lower().replace(" ", "") for h in headers]
        mapping = {}
        for field, aliases in _HEADER_ALIASES.items():
            for a in aliases:
                key = a.lower().replace(" ", "")
                if key in lowered:
                    mapping[field] = lowered.index(key); break
        if all(f in mapping for f in want) and len(mapping) > len(best[1]):
            best = (idx, mapping, headers)
    return best


@master_bp.route("/api/master/products/import", methods=["POST"])
def api_import_products():
    """匯入主檔：認「酷澎主檔」（一個線別一份，有線別欄）和「寶僑總表 Sheet1」兩種格式，
    看標題自動分辨（見 _HEADER_ALIASES）。只更新有值的欄位，不會把既有資料洗成空白。
    酷澎主檔的線別會同時記進 lines_seen，篩選、總表立刻分得出線別，不用等訂單來學。"""
    upload = request.files.get("file")
    if upload is None or not upload.filename:
        return jsonify({"error": "沒有收到檔案。"}), 400
    try:
        wb = openpyxl.load_workbook(io.BytesIO(upload.read()), data_only=True, read_only=True)
    except Exception as exc:  # noqa: BLE001
        return jsonify({"error": f"無法開啟 Excel：{exc}"}), 400
    chosen = None
    for ws in wb.worksheets:
        hdr_idx, mapping, _ = _find_columns(ws, want=("barcode", "box_size"))
        if hdr_idx:
            chosen = (ws, hdr_idx, mapping); break
    if chosen is None:
        return jsonify({"error": "找不到同時有「國條／Barcode」和「箱入數」標題的工作表。"}), 400
    ws, hdr_idx, mapping = chosen

    def get(row, field):
        i = mapping.get(field)
        return row[i] if i is not None and i < len(row) else None

    operator = _operator()
    conn = get_conn()
    try:
        counts = collections.Counter()
        for row in ws.iter_rows(min_row=hdr_idx + 1, values_only=True):
            barcode = norm_key(get(row, "barcode"))
            if not barcode:
                continue
            fields = {k: (norm_key(get(row, k)) if k in ("sku_id", "yf_sku") else norm_text(get(row, k)))
                      for k in _PRODUCT_TEXT_FIELDS}
            fields["active"] = fields["active"].upper()[:1] if fields["active"] else ""
            fields["shelf_days"] = norm_int(get(row, "shelf_days"))
            counts[_upsert_product(conn, barcode, fields, norm_int(get(row, "box_size")),
                                   norm_decimal(get(row, "cost_price")), operator, "import",
                                   upload.filename, only_filled=True)] += 1
        conn.commit()
    finally:
        conn.close()
    return jsonify({"ok": True, "sheet": ws.title, "added": counts["added"], "updated": counts["updated"],
                    "unchanged": counts["unchanged"],
                    "columns_found": sorted(mapping.keys())})


# ---------------------------------------------------------------- 月配額（畫面先收起，API 留著）

@master_bp.route("/api/master/quota", methods=["PUT"])
def api_set_quota():
    payload = request.get_json(silent=True) or {}
    line = norm_text(payload.get("line")); barcode = norm_key(payload.get("barcode"))
    month = norm_text(payload.get("month"))
    if not line or not barcode or not _valid_month(month):
        return jsonify({"error": "線別、國條、月份（2026-09）都要有。"}), 400
    raw = payload.get("qty_cases")
    qty = None if raw in (None, "") else norm_decimal(raw)
    if raw not in (None, "") and qty is None:
        return jsonify({"error": "配額要是數字。"}), 400
    operator = _operator()
    conn = get_conn()
    try:
        existing = _row(conn.execute(
            "SELECT * FROM mst_quotas WHERE line = ? AND barcode = ? AND month = ?", (line, barcode, month)))
        stamp = now(); old = existing["qty_cases"] if existing else None
        if existing is None:
            conn.execute("""INSERT INTO mst_quotas (line, barcode, month, qty_cases, note, updated_by, updated_at)
                            VALUES (?,?,?,?,?,?,?)""", (line, barcode, month, qty, norm_text(payload.get("note")), operator, stamp))
        else:
            conn.execute("UPDATE mst_quotas SET qty_cases = ?, updated_by = ?, updated_at = ? WHERE id = ?",
                         (qty, operator, stamp, existing["id"]))
        if not _same(old, qty):
            _log(conn, line, "", "", barcode, "quota", f"{month} 配額", old, qty, operator, "manual")
        conn.commit()
        return jsonify({"ok": True, "qty_cases": qty})
    finally:
        conn.close()


# ---------------------------------------------------------------- 總表（現算）

def _build_summary(conn, group, month, cfg):
    products = _products_map(conn)
    orders = [o for o in _month_orders(conn, month, cfg)
              if o["delivery_date"] and o["line_group"] == group]
    quotas = {q["barcode"]: q for q in _rows(conn.execute(
        "SELECT * FROM mst_quotas WHERE line = ? AND month = ?", (group, month)))}

    by_bc = collections.OrderedDict(); dates = set()
    for o in orders:
        bc = o["barcode"] or f"(無國條) {o['sku_id']}"
        e = by_bc.setdefault(bc, {"by_date": collections.defaultdict(float), "month_total": 0.0,
                                  "missing_box": 0, "rows": 0, "name": o["product_name"],
                                  "brand": o["brand"], "sku_id": o["sku_id"], "yf_sku": o["yf_sku"],
                                  "box_size": o["box_size"], "lines": set()})
        e["rows"] += 1; e["lines"].add(o["line"] or ""); dates.add(o["delivery_date"])
        if o["cases"] is None:
            e["missing_box"] += 1
        else:
            e["by_date"][o["delivery_date"]] += o["cases"]; e["month_total"] += o["cases"]
    dates = sorted(dates)

    # 畫面上的總表只列「這個月有交貨」的商品；主檔裡有但沒出貨的不列（照 Chloe 的說法，
    # 總表只吃箱數）。寶僑要一模一樣的匇出走底稿那條路（見 api_export_template）。
    barcodes = list(by_bc)
    rows = []
    for bc in barcodes:
        p = products.get(bc); e = by_bc.get(bc); quota = quotas.get(bc)
        q = quota["qty_cases"] if quota else None
        total = round(e["month_total"], 2) if e else 0.0
        rows.append({
            "barcode": bc, "category": p["category"] if p else "",
            "sku_id": (p or e or {}).get("sku_id", ""), "yf_sku": (p or e or {}).get("yf_sku", ""),
            "brand": (p["brand"] if p and p["brand"] else (e["brand"] if e else "")),
            "product_name": (p["product_name"] if p and p["product_name"] else (e["name"] if e else "")),
            "box_size": p["box_size"] if p and p["box_size"] else (e["box_size"] if e else None),
            "in_master": p is not None, "note": p["note"] if p else "",
            "lines_raw": sorted(e["lines"]) if e else _split_lines(p["lines_seen"]) if p else [],
            "by_date": {d: round(v, 2) for d, v in e["by_date"].items()} if e else {},
            "month_total": total, "missing_box": e["missing_box"] if e else 0,
            "order_rows": e["rows"] if e else 0,
            "quota": q, "remaining": (round(q - total, 2) if q is not None else None),
        })
    rows.sort(key=lambda r: (r["brand"], r["product_name"], r["barcode"]))
    totals_by_date = {d: round(sum(r["by_date"].get(d, 0) for r in rows), 2) for d in dates}
    return {"line": group, "month": month, "dates": dates, "rows": rows, "totals_by_date": totals_by_date,
            "month_total": round(sum(r["month_total"] for r in rows), 2),
            "missing_box_rows": sum(r["missing_box"] for r in rows),
            "not_in_master": sum(1 for r in rows if not r["in_master"])}


@master_bp.route("/api/master/summary")
def api_summary():
    group = norm_text(request.args.get("line")); month = norm_text(request.args.get("month")) or _this_month()
    if not group:
        return jsonify({"error": "請選線別。"}), 400
    if not _valid_month(month):
        return jsonify({"error": "月份格式要像 2026-09。"}), 400
    cfg = _line_groups()
    conn = get_conn()
    try:
        return jsonify(_build_summary(conn, group, month, cfg))
    finally:
        conn.close()


# ---------------------------------------------------------------- 匯出

def _xlsx_response(wb, fname):
    out = io.BytesIO(); wb.save(out); out.seek(0)
    return send_file(out, as_attachment=True, download_name=fname,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ---------------------------------------------------------------- 總表底稿（寶僑總表一模一樣）

_DATE_HDR = re.compile(r"^\s*(\d{1,2})/(\d{0,2})交貨")   # 9/2交貨、7/9交貨_1、9/18交貨\n竹運出、9/交貨（空欄）


def _template_row(conn, line):
    return _row(conn.execute("SELECT * FROM mst_templates WHERE line = ?", (line,)))


def _template_meta(tpl, conn=None):
    """給畫面看的：檔名、誰、何時、底稿裡有哪幾個月的日期欄、各月還剩幾個空欄。"""
    import base64
    meta = {"exists": True, "line": tpl["line"], "filename": tpl["filename"], "sheet": tpl["sheet"],
            "uploaded_by": tpl["uploaded_by"], "uploaded_at": tpl["uploaded_at"], "months": {}}
    try:
        wb = openpyxl.load_workbook(io.BytesIO(base64.b64decode(tpl["content_b64"])), read_only=True, data_only=True)
        ws = wb[tpl["sheet"]]
        for row in ws.iter_rows(min_row=tpl["header_row"], max_row=tpl["header_row"], values_only=True):
            for h in row:
                m = _DATE_HDR.match(norm_text(h))
                if m:
                    e = meta["months"].setdefault(int(m.group(1)), {"dates": 0, "spare": 0})
                    e["dates" if m.group(2) else "spare"] += 1
    except Exception as exc:  # noqa: BLE001
        meta["error"] = str(exc)
    return meta


@master_bp.route("/api/master/template")
def api_template_get():
    line = norm_text(request.args.get("line"))
    conn = get_conn()
    try:
        tpl = _template_row(conn, line)
    finally:
        conn.close()
    return jsonify(_template_meta(tpl) if tpl else {"exists": False, "line": line})


@master_bp.route("/api/master/template", methods=["POST"])
def api_template_upload():
    """上傳某線別的總表底稿。要能找到「Barcode／國條」標題的工作表才收。"""
    import base64
    line = norm_text(request.form.get("line") or request.args.get("line"))
    upload = request.files.get("file")
    if not line:
        return jsonify({"error": "請先選線別。"}), 400
    if upload is None or not upload.filename:
        return jsonify({"error": "沒有收到檔案。"}), 400
    raw = upload.read()
    try:
        wb = openpyxl.load_workbook(io.BytesIO(raw), read_only=True, data_only=True)
    except Exception as exc:  # noqa: BLE001
        return jsonify({"error": f"無法開啟 Excel：{exc}"}), 400
    chosen = None
    for ws in wb.worksheets:
        hdr_idx, mapping, _ = _find_columns(ws, want=("barcode",))
        if hdr_idx and any(_DATE_HDR.match(norm_text(h)) for h in next(ws.iter_rows(min_row=hdr_idx, max_row=hdr_idx, values_only=True))):
            chosen = (ws.title, hdr_idx); break
    if chosen is None:
        return jsonify({"error": "找不到同時有「Barcode／國條」標題和「M/D交貨」日期欄的工作表，這不像總表。"}), 400
    conn = get_conn()
    try:
        conn.execute("DELETE FROM mst_templates WHERE line = ?", (line,))
        conn.execute(
            """INSERT INTO mst_templates (line, filename, sheet, header_row, content_b64, uploaded_by, uploaded_at)
               VALUES (?,?,?,?,?,?,?)""",
            (line, upload.filename, chosen[0], chosen[1], base64.b64encode(raw).decode("ascii"), _operator(), now()))
        _log(conn, line, "", "", "", "template", "總表底稿", "", upload.filename, _operator(), "manual",
             f"工作表「{chosen[0]}」")
        conn.commit()
        tpl = _template_row(conn, line)
    finally:
        conn.close()
    return jsonify({"ok": True, **_template_meta(tpl)})


@master_bp.route("/api/master/template", methods=["DELETE"])
def api_template_delete():
    line = norm_text(request.args.get("line"))
    conn = get_conn()
    try:
        tpl = _template_row(conn, line)
        if tpl is None:
            return jsonify({"error": "這個線別沒有底稿。"}), 404
        conn.execute("DELETE FROM mst_templates WHERE line = ?", (line,))
        _log(conn, line, "", "", "", "template", "總表底稿", tpl["filename"], "", _operator(), "manual", "移除底稿")
        conn.commit()
    finally:
        conn.close()
    return jsonify({"ok": True})


def _fill_template(tpl, summary, month, operator):
    """把這個月每個交貨日的箱數填進底稿。只動：這個月的日期欄（含把「9/交貨」空欄補上日期）
    和底稿裡對得到的商品列；其他列、其他月份、業務的公式、順序全部不碰。
    對不到的商品、沒地方填的日期，寫在最後一個分頁「系統填入說明」。"""
    import base64
    wb = openpyxl.load_workbook(io.BytesIO(base64.b64decode(tpl["content_b64"])))   # 保留公式
    ws = wb[tpl["sheet"]]
    hdr = tpl["header_row"]
    mm = int(month[5:7])
    headers = {c: norm_text(ws.cell(row=hdr, column=c).value) for c in range(1, ws.max_column + 1)}
    lowered = {c: h.lower().replace(" ", "") for c, h in headers.items()}
    def col_of(field):
        for a in _HEADER_ALIASES[field]:
            key = a.lower().replace(" ", "")
            for c, h in lowered.items():
                if h == key:
                    return c
        return None
    bc_col, sku_col = col_of("barcode"), col_of("sku_id")
    date_cols, spare = {}, []
    for c, h in headers.items():
        m = _DATE_HDR.match(h)
        if m and int(m.group(1)) == mm:
            if m.group(2):
                date_cols[int(m.group(2))] = c
            else:
                spare.append(c)
    by_bc, by_sku = {}, {}
    for r in range(hdr + 1, ws.max_row + 1):
        bc = norm_key(ws.cell(row=r, column=bc_col).value) if bc_col else ""
        sku = norm_key(ws.cell(row=r, column=sku_col).value) if sku_col else ""
        if bc and bc not in by_bc:
            by_bc[bc] = r
        if sku and sku not in by_sku:
            by_sku[sku] = r
    # 每個交貨日找欄：有同日期的欄就用；沒有就拿一個「M/交貨」空欄補上日期；都沒有就記下來
    missing_dates, renamed = [], []
    for d in summary["dates"]:
        day = int(d[8:10])
        if day in date_cols:
            continue
        if spare:
            c = spare.pop(0); date_cols[day] = c
            ws.cell(row=hdr, column=c).value = f"{mm}/{day}交貨"; renamed.append(f"{mm}/{day}")
        else:
            missing_dates.append(d)
    block = sorted(date_cols.values()) + spare
    matched, unmatched, filled = 0, [], 0
    for r0 in summary["rows"]:
        r = by_bc.get(r0["barcode"]) or by_sku.get(r0["sku_id"])
        if r is None:
            unmatched.append(r0); continue
        matched += 1
        for c in block:                      # 這個月的欄先清空再填，沒出貨的日子留白
            ws.cell(row=r, column=c).value = None
        for d, v in r0["by_date"].items():
            c = date_cols.get(int(d[8:10]))
            if c is not None:
                ws.cell(row=r, column=c).value = v; filled += 1
    # 說明分頁
    from openpyxl.styles import Font
    info = wb.create_sheet("系統填入說明")
    bold = Font(bold=True)
    info.append(["商品主檔自動化 填入紀錄"]); info["A1"].font = bold
    info.append(["填入時間", now(), "操作者", operator])
    info.append(["月份", month, "線別", summary["line"]])
    info.append(["底稿", tpl["filename"], "工作表", tpl["sheet"]])
    info.append(["對到的商品", matched, "填入格數", filled])
    if renamed:
        info.append(["補上日期的空欄", "、".join(renamed)])
    info.append([])
    if missing_dates:
        info.append(["⚠ 這些交貨日在底稿裡沒有欄位可填（連空欄都用完了），請業務在總表加欄後重匯："]); info.cell(row=info.max_row, column=1).font = bold
        for d in missing_dates:
            tot = summary["totals_by_date"].get(d)
            info.append([d, f"{tot} 箱" if tot is not None else ""])
        info.append([])
    if unmatched:
        info.append(["酷澎有下單、但總表裡沒有的商品（沒混進總表，列在這裡給你看）："]); info.cell(row=info.max_row, column=1).font = bold
        info.append(["國條", "SKU ID", "品名", "品牌", "箱入數"] + [_md(d) for d in summary["dates"]] + ["月加總"])
        for c in info[info.max_row]:
            c.font = bold
        for r0 in unmatched:
            info.append([r0["barcode"], r0["sku_id"], r0["product_name"], r0["brand"], r0["box_size"]]
                        + [r0["by_date"].get(d) for d in summary["dates"]] + [r0["month_total"]])
    info.column_dimensions["A"].width = 22; info.column_dimensions["B"].width = 18; info.column_dimensions["C"].width = 40
    return wb, {"matched": matched, "filled": filled, "unmatched": len(unmatched),
                "missing_dates": missing_dates, "renamed": renamed}


@master_bp.route("/api/master/export")
def api_export():
    """總表格式：一列一個國條，往右是各交貨日箱數、月加總。"""
    group = norm_text(request.args.get("line")); month = norm_text(request.args.get("month")) or _this_month()
    if not group or not _valid_month(month):
        return jsonify({"error": "請選線別與月份。"}), 400
    cfg = _line_groups()
    conn = get_conn()
    try:
        s = _build_summary(conn, group, month, cfg)
        orders = [o for o in _month_orders(conn, month, cfg) if o["line_group"] == group]
        tpl = _template_row(conn, group)
    finally:
        conn.close()
    if tpl is not None:
        # 有底稿就照底稿填（寶僑要跟他們的總表一模一樣，含商品順序），檔名沿用底稿的
        wb, _rep = _fill_template(tpl, s, month, _operator())
        return _xlsx_response(wb, tpl["filename"] or f"{group}_總表_{month}.xlsx")
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "總表"
    mm = int(month[5:7])
    head = ["國條", "SKU ID", "永豐料號", "品類", "品牌", "品名", "箱入數"] + [_md(d) for d in s["dates"]] + \
           [f"{mm}月TTL下單總箱數", "備註"]
    ws.append(head)
    for r in s["rows"]:
        ws.append([r["barcode"], r["sku_id"], r["yf_sku"], r["category"], r["brand"], r["product_name"], r["box_size"]]
                  + [r["by_date"].get(d) for d in s["dates"]] + [r["month_total"], r["note"]])
    ws.append([])
    ws.append(["合計", "", "", "", "", "", ""] + [s["totals_by_date"].get(d) for d in s["dates"]] + [s["month_total"], ""])
    bold = Font(bold=True); fill = PatternFill("solid", fgColor="DBEAFE")
    for c in ws[1]:
        c.font = bold; c.fill = fill; c.alignment = Alignment(horizontal="center", wrap_text=True)
    for c in ws[ws.max_row]:
        c.font = bold
    ws.freeze_panes = "H2"
    for i in range(1, len(head) + 1):
        ws.column_dimensions[get_column_letter(i)].width = 40 if i == 6 else 14
    for row in ws.iter_rows(min_row=2, min_col=1, max_col=3):
        for c in row:
            c.number_format = "@"
    ws2 = wb.create_sheet("每日合計"); ws2.append(["交貨日", "出貨總箱數", "PO 數", "品項數"])
    per_day = collections.OrderedDict()
    for o in orders:
        if not o["delivery_date"]:
            continue
        d = per_day.setdefault(o["delivery_date"], {"cases": 0.0, "pos": set(), "rows": 0})
        d["cases"] += o["cases"] or 0; d["pos"].add(o["po_number"]); d["rows"] += 1
    for dte, d in per_day.items():
        ws2.append([dte, round(d["cases"], 2), len(d["pos"]), d["rows"]])
    for c in ws2[1]:
        c.font = bold; c.fill = fill
    return _xlsx_response(wb, f"{group}_總表_{month}.xlsx")


@master_bp.route("/api/master/export/daily")
def api_export_daily():
    """專案報價檔：一個交貨日一個分頁（0904交貨、0903交貨…新的在前），A～R 每欄照
    原本公式的來源填。匯出的就是畫面上篩出來的（線別／日期／PO／倉別／品牌／關鍵字），
    什麼都沒篩才是整月；帶 month_to 可以一次匯好幾個月。"""
    month = norm_text(request.args.get("month")) or _this_month()
    month_to = norm_text(request.args.get("month_to")) or month
    if not _valid_month(month) or not _valid_month(month_to):
        return jsonify({"error": "月份格式要像 2026-09。"}), 400
    months = _month_span(month, month_to)       # 可以跨月：7 月～9 月就是三個月的交貨日全部進來
    filters = _read_filters(request.args)
    cfg = _line_groups()
    conn = get_conn()
    try:
        orders = [o for o in _month_orders(conn, months, cfg) if _keep(o, filters)]
    finally:
        conn.close()
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    head = ["SKU ID", "永豐料號", "國條", "品類", "品牌", "品名", "下單數量(酷澎單位)", "出貨數量", "箱入數",
            "出貨數量(箱)", "單價(含稅)", "酷澎下單價(含稅)", "箱單價(含稅)", "總計(含稅)", "備註",
            "驗收完成請打勾", "簽單完成請打勾", "交貨日"]
    yellow = PatternFill("solid", fgColor="FFFF00"); head_fill = PatternFill("solid", fgColor="F8CBAD"); bold = Font(bold=True)
    by_date = collections.OrderedDict()
    for o in orders:
        by_date.setdefault(o["delivery_date"] or "", collections.OrderedDict()).setdefault(o["po_number"], []).append(o)
    wb = openpyxl.Workbook(); wb.remove(wb.active)
    for date in sorted(by_date.keys(), key=lambda d: d or "0000", reverse=True):
        pos = by_date[date]
        if date:
            d = _dt.date.fromisoformat(date); title = f"{d.month:02d}{d.day:02d}交貨"; label = f"{d.month}/{d.day}交貨"
        else:
            title, label = "未排日期", ""
        ws = wb.create_sheet(title[:31]); ws.append(head)
        for c in ws[1]:
            c.font = bold; c.fill = head_fill
        first = True
        for po, rows in pos.items():
            if not first:
                ws.append([None] * len(head))
                for c in ws[ws.max_row]:
                    c.fill = yellow
            first = False
            for i, o in enumerate(rows):
                box = o["box_size"]; cp = o["unit_price"]
                box_price = (cp * box) if (cp is not None and box) else None
                total = (box_price * o["cases"]) if (box_price is not None and o["cases"] is not None) else None
                po_cell = (f"{po}_{label}({o['warehouse']})" if o["warehouse"] else f"{po}_{label}") if i == 0 else None
                ws.append([o["sku_id"], o.get("yf_sku_master") or o["yf_sku"], o["barcode"], o.get("category") or "",
                           o.get("brand_master") or o["brand"], o["product_name"], o["qty_coupang"], o["qty_ship"],
                           box, o["cases"], o.get("cost_price"), cp, box_price, total, o["export_note"], "", "", po_cell])
        ws.freeze_panes = "A2"
        for i, w in enumerate([16, 15, 15, 8, 14, 44, 10, 9, 8, 11, 10, 12, 11, 12, 18, 8, 8, 30], start=1):
            ws.column_dimensions[get_column_letter(i)].width = w
        for row in ws.iter_rows(min_row=2, min_col=1, max_col=3):
            for c in row:
                c.number_format = "@"
        for row in ws.iter_rows(min_row=2, min_col=18, max_col=18):
            for c in row:
                c.alignment = Alignment(horizontal="left")
    if not wb.sheetnames:
        ws = wb.create_sheet("無資料"); ws.append(["目前的篩選條件下沒有任何訂單"])
    # 檔名照同事原本的檔就叫「專案報價檔」，分頁是 0904交貨、0903交貨…新的在前
    return _xlsx_response(wb, "專案報價檔.xlsx")


# ---------------------------------------------------------------- 歷程

@master_bp.route("/api/master/reset", methods=["POST"])
def api_reset():
    """清除資料，跟訂單管理系統的「清空訂單資料」同一套規矩：只有管理員、
    要照著打「清空資料」四個字、清之前先自動備份（PostgreSQL 模式由 Supabase
    每日備份頂著，見 db.backup_db）。

    分兩個範圍讓人勾：訂單明細（連匯入預覽的暫存一起）、商品主檔（連配額）。
    最常見是「試用完把測試訂單清掉，主檔留著繼續用」，所以主檔預設不勾。
    修改歷程預設一起清；勾「保留」時會多寫一筆系統紀錄，之後查得到這件事。"""
    if not _is_admin():
        return jsonify({"error": "只有管理員可以清除資料。"}), 403
    payload = request.get_json(silent=True) or {}
    if norm_text(payload.get("confirm")) != "清空資料":
        return jsonify({"error": "請照著輸入「清空資料」四個字再確認。"}), 400
    clear_orders = bool(payload.get("orders", True))
    clear_products = bool(payload.get("products", False))
    keep_logs = bool(payload.get("keep_logs", False))
    if not clear_orders and not clear_products:
        return jsonify({"error": "至少要勾一個要清的範圍。"}), 400

    backup = db.backup_db("mst_reset")
    conn = get_conn()
    try:
        counts = {}
        for key, table in (("orders", "mst_orders"), ("products", "mst_products"),
                           ("quotas", "mst_quotas"), ("logs", "mst_logs")):
            counts[key] = conn.execute(f"SELECT COUNT(*) AS n FROM {table}").fetchone()["n"]
        parts = []
        if clear_orders:
            conn.execute("DELETE FROM mst_orders")
            conn.execute("DELETE FROM mst_import_batches")
            parts.append(f"訂單明細 {counts['orders']} 筆")
        if clear_products:
            conn.execute("DELETE FROM mst_quotas")
            conn.execute("DELETE FROM mst_products")
            parts.append(f"商品主檔 {counts['products']} 筆")
        if keep_logs:
            _log(conn, "", "", "", "", "reset", "清除資料", "、".join(parts), "已清空",
                 _operator(), "system", "管理員手動清除，修改歷程保留")
        else:
            conn.execute("DELETE FROM mst_logs")
        conn.commit()
    finally:
        conn.close()

    tail = "，修改歷程保留" if keep_logs else "，修改歷程一併清除"
    note = f"（清空前已自動備份：{os.path.basename(backup)}）" if backup else ""
    return jsonify({"ok": True, "message": f"已清除 {'、'.join(parts)}{tail}。{note}"})


@master_bp.route("/api/master/imports")
def api_imports():
    """匯入歷程：每一次「確認匯入」一列，附上「這批目前還算新增的有幾筆、落在哪幾個月」，
    讓 OP 分得出同一天匯兩次時，第二次多出來的是哪些（拿去只匯那批的專案報價檔）。"""
    limit = min(norm_int(request.args.get("limit")) or 60, 300)
    conn = get_conn()
    try:
        batches = _rows(conn.execute(
            """SELECT id, filename, operator, rows_total, rows_new, rows_updated, rows_identical, created_at, committed_at
               FROM mst_import_batches WHERE committed = 1 ORDER BY committed_at DESC, id DESC LIMIT ?""", (limit,)))
        for b in batches:
            months = _rows(conn.execute(
                """SELECT substr(delivery_date, 1, 7) AS m, COUNT(*) AS n FROM mst_orders
                   WHERE first_batch_id = ? GROUP BY substr(delivery_date, 1, 7) ORDER BY m""", (b["id"],)))
            b["new_now"] = sum(x["n"] for x in months)
            b["months"] = [x["m"] for x in months if x["m"]]
            b["changed_now"] = _row(conn.execute(
                "SELECT COUNT(*) AS n FROM mst_orders WHERE last_batch_id = ? AND first_batch_id != ?",
                (b["id"], b["id"])))["n"]
    finally:
        conn.close()
    return jsonify({"batches": batches})


@master_bp.route("/api/master/logs")
def api_logs():
    po = norm_key(request.args.get("po")); barcode = norm_key(request.args.get("barcode"))
    q = norm_text(request.args.get("q")); limit = min(norm_int(request.args.get("limit")) or 200, 1000)
    where, params = ["1=1"], []
    if po:
        where.append("po_number = ?"); params.append(po)
    if barcode:
        where.append("barcode = ?"); params.append(barcode)
    if q:
        where.append("(po_number LIKE ? OR barcode LIKE ? OR operator LIKE ? OR field_label LIKE ? "
                     "OR old_value LIKE ? OR new_value LIKE ? OR note LIKE ?)")
        params += [f"%{q}%"] * 7
    conn = get_conn()
    try:
        rows = _rows(conn.execute(
            f"SELECT * FROM mst_logs WHERE {' AND '.join(where)} ORDER BY changed_at DESC, id DESC LIMIT ?",
            params + [limit]))
        return jsonify({"logs": rows})
    finally:
        conn.close()

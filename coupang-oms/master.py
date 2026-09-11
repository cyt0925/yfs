"""業績總表自動化：取代「訂單彙總表 → 專案報價檔 → 總表」那條人工流程。

OP 原本的做法（見 Word 需求文件）：
  1. 從凱特系統下載「訂單彙總表」（就是 OMS 吃的那份整合表）。
  2. 貼進「專案報價檔」，依交貨日一個分頁一個分頁放，不同 PO 用黃底隔開。
  3. 訂單數量、交貨日一改，報價檔就要跟著手改。
  4. 改完重做樞紐：國條 × 出貨數量(箱)，算出每個交貨日的出貨總箱數。
  5. 再 VLOOKUP 回填「總表」對應日期欄位（要貼成值）。
  每次異動都要重跑 4、5，一個月做幾十次。

這個模組把 2～5 全部變成「上傳一次、畫面上改、總表自己算」：
  - 訂單明細（mst_orders）：依線別／交貨日／PO 分組，出貨數量、交貨日、
    備註都能直接改；人改過的欄位之後再匯入不會被蓋掉（_overridden 旗標，
    跟 OMS 的 qty_ship_overridden 同一套做法）。
  - 商品主檔（mst_products）：國條 → 箱入數，OP 自己維護。這是「箱數」的
    唯一來源；主檔沒有的國條退回用整合表自帶的箱入數，並在畫面上標出來。
  - 月配額（mst_quotas）：總表 O 欄「Supply_CS」，一個線別一個國條一個月
    一個數字，手填或從總表檔匯入。
  - 總表：每次都現算，不存快照。每日箱數 = Σ(出貨數量 ÷ 箱入數) 依
    交貨日 + 國條加總（= 樞紐），月加總 = 總表 BD 欄，剩餘可供貨量 =
    配額 − 月加總（= 總表 BE 欄）。OP 改任何一筆，總表下一次載入就是
    最新的，不用重新上傳、不用重拉樞紐。

跟酷澎訂單管理系統的關係：吃同一份整合表，但表完全分開（都以 mst_ 開頭），
互不影響；共用登入、資料庫連線、整合表解析（importer.parse_workbook）。
線別不寫死——整合表裡「線別」欄是什麼就分成什麼，寶僑、紙潔、瑪氏各自
一套主檔與配額。
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

# 沒有任何資料時畫面上先給的線別選項；真正的線別一律以整合表「線別」欄
# 為準，出現什麼就多什麼，不用改程式。
DEFAULT_LINES = ["寶僑", "紙潔", "瑪氏"]
UNCLASSIFIED_LINE = "未分類"

# 這些欄位是酷澎（整合表）擁有的：只有匯入能改，任何變動都記歷程。
COUPANG_FIELDS = {
    "barcode": "國條", "yf_sku": "永豐料號", "brand": "品牌",
    "product_name": "品名", "warehouse": "倉別", "order_type": "訂單類型",
    "unit": "單位", "unit_price": "下單單價", "qty_coupang": "下單數量",
    "qty_file_ship": "整合表出貨數量", "box_size_file": "整合表箱入數",
    "delivery_date_file": "整合表交貨日",
}

# 人可以改的欄位：改過就立旗標，匯入不再覆蓋。
HUMAN_FIELDS = {
    "qty_ship": ("qty_ship_overridden", "出貨數量"),
    "delivery_date": ("delivery_date_overridden", "交貨日"),
    "remarks": ("remarks_overridden", "備註"),
}


# ---------------------------------------------------------------- 小工具

def _operator():
    return session.get("user", "")


def _rows(cur):
    return [dict(r) for r in cur.fetchall()]


def _row(cur):
    r = cur.fetchone()
    return dict(r) if r is not None else None


def _log(conn, line, po, sku, barcode, field, label, old, new, operator,
         source, note=""):
    conn.execute(
        """INSERT INTO mst_logs
           (line, po_number, sku_id, barcode, field, field_label, old_value,
            new_value, operator, source, note, changed_at)
           VALUES (?,?,?,?,?,?,?,?,?,?,?,?)""",
        (line, po, sku, barcode, field, label,
         "" if old is None else str(old), "" if new is None else str(new),
         operator, source, note, now()))


def _month_of(date_str):
    return (date_str or "")[:7]


def _this_month():
    return _dt.date.today().strftime("%Y-%m")


def _valid_month(text):
    return bool(re.fullmatch(r"\d{4}-\d{2}", text or ""))


def _cases(qty, box):
    """出貨數量 ÷ 箱入數。箱入數缺或為 0 就算不出來，回 None 讓畫面標紅，
    不要硬塞 0——0 會被當成「這天沒出貨」，跟「算不出來」完全是兩回事。"""
    if qty is None or not box:
        return None
    return round(qty / box, 4)


def _same(a, b):
    if a is None or b is None:
        return a is None and b is None
    return str(a) == str(b)


def _lines(conn):
    found = set(DEFAULT_LINES)
    for table in ("mst_orders", "mst_products", "mst_quotas"):
        for r in conn.execute(f"SELECT DISTINCT line FROM {table}").fetchall():
            if r["line"]:
                found.add(r["line"])
    return sorted(found, key=lambda x: (x not in DEFAULT_LINES,
                                        DEFAULT_LINES.index(x) if x in DEFAULT_LINES else 0, x))


def _products_map(conn, line):
    return {p["barcode"]: p for p in _rows(conn.execute(
        "SELECT * FROM mst_products WHERE line = ?", (line,)))}


def _effective_box(order, product):
    """箱入數以主檔為準；主檔沒有（或沒填）才退回整合表自帶的值。"""
    if product and product.get("box_size"):
        return product["box_size"], "master"
    if order.get("box_size_file"):
        return order["box_size_file"], "file"
    return None, "none"


# ---------------------------------------------------------------- 頁面

@master_bp.route("/master")
def master_page():
    logo = ("logo_master.png"
            if os.path.exists(os.path.join(os.path.dirname(__file__),
                                           "static", "logo_master.png"))
            else "logo.png")
    return render_template("master.html", logo_file=logo,
                           logged_in_user=_operator(),
                           build_version=current_app.config.get("BUILD_VERSION", ""))


@master_bp.route("/api/master/lines")
def api_lines():
    conn = get_conn()
    try:
        return jsonify({"lines": _lines(conn), "this_month": _this_month()})
    finally:
        conn.close()


# ---------------------------------------------------------------- 匯入（兩階段）

def _diff_import(conn, rows):
    """把整合表列跟 mst_orders 比對，分成 new／updated／identical，
    並找出主檔裡沒有的國條。"""
    new, updated, identical, missing = [], [], [], {}
    prod_cache = {}
    for r in rows:
        line = r["line"] or UNCLASSIFIED_LINE
        r["line"] = line
        if line not in prod_cache:
            prod_cache[line] = _products_map(conn, line)
        products = prod_cache[line]
        if r["barcode"] and r["barcode"] not in products:
            key = (line, r["barcode"])
            if key not in missing:
                missing[key] = {
                    "line": line, "barcode": r["barcode"], "sku_id": r["sku_id"],
                    "yf_sku": r["yf_sku"], "brand": r["brand"],
                    "product_name": r["product_name"], "box_size": r["box_size"],
                }
        existing = _row(conn.execute(
            "SELECT * FROM mst_orders WHERE line = ? AND po_number = ? AND sku_id = ?",
            (line, r["po_number"], r["sku_id"])))
        if existing is None:
            new.append(r)
            continue
        changes = []
        mapping = {
            "barcode": r["barcode"], "yf_sku": r["yf_sku"], "brand": r["brand"],
            "product_name": r["product_name"], "warehouse": r["warehouse"],
            "order_type": r["order_type"], "unit": r["unit"],
            "unit_price": r["unit_price"], "qty_coupang": r["qty_coupang"],
            "qty_file_ship": r["qty_file_ship"], "box_size_file": r["box_size"],
            "delivery_date_file": r["delivery_date"],
        }
        for field, value in mapping.items():
            if not _same(existing.get(field), value):
                changes.append({"field": field, "label": COUPANG_FIELDS[field],
                                "old": existing.get(field), "new": value})
        if existing["missing_in_file"]:
            # 之前被酷澎拿掉、這次又出現：就算欄位都一樣也要讓 OP 在預覽看到
            changes.append({"field": "missing_in_file", "label": "品項重新出現",
                            "old": "檔案已無此品項", "new": "恢復"})
        if changes:
            updated.append({
                "row": r, "existing_id": existing["id"], "changes": changes,
                "po_number": r["po_number"], "sku_id": r["sku_id"],
                "barcode": r["barcode"], "product_name": r["product_name"],
                "line": line,
                "qty_ship_overridden": existing["qty_ship_overridden"],
                "delivery_date_overridden": existing["delivery_date_overridden"],
            })
        else:
            identical.append(r)
    # 酷澎把品項數量下修到 0 時，匯出的表裡那一列會直接消失，不是留一列 0。
    # 「資料庫有、檔案沒有」平常代表上傳的只是片段、不能動；但如果這張 PO 本身
    # 有出現在檔案裡、底下某個品項卻沒帶到，那就是被酷澎拿掉了——不刪列（報價
    # 檔的做法是留著寫 0、備註打單缺貨），出貨數量歸 0，標記讓 OP 看得到。
    file_keys = {(r["line"], r["po_number"], r["sku_id"]) for r in rows}
    file_pos = {(r["line"], r["po_number"]) for r in rows}
    removed = []
    for line, po in sorted(file_pos):
        for ex in _rows(conn.execute(
                "SELECT * FROM mst_orders WHERE line = ? AND po_number = ?", (line, po))):
            if (line, po, ex["sku_id"]) in file_keys or ex["missing_in_file"]:
                continue
            removed.append({"id": ex["id"], "line": line, "po_number": po,
                            "sku_id": ex["sku_id"], "barcode": ex["barcode"],
                            "product_name": ex["product_name"],
                            "qty_coupang": ex["qty_coupang"], "qty_ship": ex["qty_ship"],
                            "delivery_date": ex["delivery_date"]})
    return new, updated, identical, list(missing.values()), removed


@master_bp.route("/api/master/import/preview", methods=["POST"])
def api_import_preview():
    upload = request.files.get("file")
    if upload is None or not upload.filename:
        return jsonify({"error": "沒有收到檔案。"}), 400
    try:
        rows, warnings = parse_workbook(io.BytesIO(upload.read()), upload.filename)
    except ImportError_ as exc:
        return jsonify({"error": str(exc)}), 400

    conn = get_conn()
    try:
        new, updated, identical, missing, removed = _diff_import(conn, rows)
        payload = {"rows": rows, "missing_products": missing, "removed": removed,
                   "filename": upload.filename}
        cur = conn.execute(
            """INSERT INTO mst_import_batches
               (filename, operator, rows_total, rows_new, rows_updated,
                rows_identical, committed, payload_json, created_at)
               VALUES (?,?,?,?,?,?,0,?,?)""",
            (upload.filename, _operator(), len(rows), len(new), len(updated),
             len(identical), json.dumps(payload, ensure_ascii=False), now()))
        conn.commit()
        batch_id = cur.lastrowid
    finally:
        conn.close()

    by_line = collections.Counter(r["line"] for r in rows)
    dates = sorted({r["delivery_date"] for r in rows if r["delivery_date"]})
    if by_line.get(UNCLASSIFIED_LINE):
        warnings = list(warnings) + [
            f"有 {by_line[UNCLASSIFIED_LINE]} 筆沒有「線別」，先歸到「{UNCLASSIFIED_LINE}」；"
            "要看它們請在上方線別選單切到「未分類」。"]
    return jsonify({
        "batch_id": batch_id, "filename": upload.filename,
        "rows_total": len(rows), "new_count": len(new),
        "updated_count": len(updated), "identical_count": len(identical),
        "lines": dict(by_line), "dates": dates, "warnings": warnings,
        "new_preview": new[:300],
        "updated": [{k: v for k, v in u.items() if k != "row"} for u in updated[:300]],
        "missing_products": missing,
        "removed": removed, "removed_count": len(removed),
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
        batch = _row(conn.execute(
            "SELECT * FROM mst_import_batches WHERE id = ?", (batch_id,)))
        if batch is None:
            return jsonify({"error": "找不到這次的預覽，請重新上傳。"}), 404
        if batch["committed"]:
            return jsonify({"error": "這批已經匯入過了，不能重複確認。"}), 409
        data = json.loads(batch["payload_json"] or "{}")
        rows = data.get("rows", [])
        stamp = now()
        inserted = updated_n = identical_n = 0
        products_added = 0

        # 先補主檔（用整合表自帶的箱入數當起始值，備註寫明來源讓 OP 知道要核）
        if add_missing:
            for m in data.get("missing_products", []):
                exists = _row(conn.execute(
                    "SELECT id FROM mst_products WHERE line = ? AND barcode = ?",
                    (m["line"], m["barcode"])))
                if exists:
                    continue
                conn.execute(
                    """INSERT INTO mst_products
                       (line, barcode, sku_id, yf_sku, brand, product_name,
                        box_size, note, updated_by, updated_at)
                       VALUES (?,?,?,?,?,?,?,?,?,?)""",
                    (m["line"], m["barcode"], m.get("sku_id") or "",
                     m.get("yf_sku") or "", m.get("brand") or "",
                     m.get("product_name") or "", m.get("box_size"),
                     "由匯入自動建立，箱入數請核對", operator, stamp))
                _log(conn, m["line"], "", m.get("sku_id") or "", m["barcode"],
                     "product_new", "新增主檔", "", m.get("box_size"),
                     operator, "import", data.get("filename", ""))
                products_added += 1

        removed_n = 0
        for rm in data.get("removed", []):
            ex = _row(conn.execute("SELECT * FROM mst_orders WHERE id = ?", (rm["id"],)))
            if ex is None or ex["missing_in_file"]:
                continue
            conn.execute(
                """UPDATE mst_orders SET qty_ship = 0, missing_in_file = 1, last_seen_at = ?,
                   updated_at = ?, version = version + 1 WHERE id = ?""",
                (stamp, stamp, ex["id"]))
            _log(conn, ex["line"], ex["po_number"], ex["sku_id"], ex["barcode"],
                 "qty_ship", "出貨數量", ex["qty_ship"], 0, operator, "import",
                 "這次的整合表裡這張 PO 已沒有這個品項，出貨數量歸 0")
            removed_n += 1

        for r in rows:
            line = r["line"] or UNCLASSIFIED_LINE
            existing = _row(conn.execute(
                "SELECT * FROM mst_orders WHERE line = ? AND po_number = ? AND sku_id = ?",
                (line, r["po_number"], r["sku_id"])))
            ship = r["qty_file_ship"] if r["qty_file_ship"] is not None else r["qty_coupang"]
            if existing is None:
                conn.execute(
                    """INSERT INTO mst_orders
                       (line, po_number, sku_id, barcode, yf_sku, brand, product_name,
                        warehouse, order_type, unit, unit_price, qty_coupang,
                        qty_file_ship, qty_ship, box_size_file, delivery_date_file,
                        delivery_date, remarks_file, remarks, source_file,
                        first_seen_at, last_seen_at, updated_at, version)
                       VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,1)""",
                    (line, r["po_number"], r["sku_id"], r["barcode"], r["yf_sku"],
                     r["brand"], r["product_name"], r["warehouse"], r["order_type"],
                     r["unit"], r["unit_price"], r["qty_coupang"], r["qty_file_ship"],
                     ship, r["box_size"], r["delivery_date"], r["delivery_date"],
                     r["remarks_file"], r["remarks_file"], r["source_file"],
                     stamp, stamp, stamp))
                inserted += 1
                continue

            sets, vals, changed = [], [], False
            mapping = {
                "barcode": r["barcode"], "yf_sku": r["yf_sku"], "brand": r["brand"],
                "product_name": r["product_name"], "warehouse": r["warehouse"],
                "order_type": r["order_type"], "unit": r["unit"],
                "unit_price": r["unit_price"], "qty_coupang": r["qty_coupang"],
                "qty_file_ship": r["qty_file_ship"], "box_size_file": r["box_size"],
                "delivery_date_file": r["delivery_date"],
            }
            for field, value in mapping.items():
                if not _same(existing.get(field), value):
                    sets.append(f"{field} = ?"); vals.append(value); changed = True
                    _log(conn, line, r["po_number"], r["sku_id"], r["barcode"],
                         field, COUPANG_FIELDS[field], existing.get(field), value,
                         operator, "import", data.get("filename", ""))
            # 人沒動過的欄位才跟著整合表走
            if not existing["qty_ship_overridden"] and not _same(existing["qty_ship"], ship):
                sets.append("qty_ship = ?"); vals.append(ship); changed = True
                _log(conn, line, r["po_number"], r["sku_id"], r["barcode"],
                     "qty_ship", "出貨數量", existing["qty_ship"], ship,
                     operator, "import", "隨整合表更新")
            if (not existing["delivery_date_overridden"]
                    and not _same(existing["delivery_date"], r["delivery_date"])):
                sets.append("delivery_date = ?"); vals.append(r["delivery_date"]); changed = True
                _log(conn, line, r["po_number"], r["sku_id"], r["barcode"],
                     "delivery_date", "交貨日", existing["delivery_date"],
                     r["delivery_date"], operator, "import", "隨整合表更新")
            if (not existing["remarks_overridden"]
                    and not _same(existing["remarks"], r["remarks_file"])):
                sets.append("remarks = ?"); vals.append(r["remarks_file"]); changed = True
            if existing["missing_in_file"]:
                # 之前被酷澎拿掉、現在又出現了：解除標記，出貨數量照一般規則重新同步
                sets.append("missing_in_file = 0"); changed = True
                if not existing["qty_ship_overridden"] and "qty_ship = ?" not in sets:
                    sets.append("qty_ship = ?"); vals.append(ship)
                _log(conn, line, r["po_number"], r["sku_id"], r["barcode"],
                     "missing_in_file", "品項重新出現", 1, 0, operator, "import",
                     "整合表裡又有這個品項了")
            sets.append("last_seen_at = ?"); vals.append(stamp)
            sets.append("source_file = ?"); vals.append(r["source_file"])
            if changed:
                sets.append("updated_at = ?"); vals.append(stamp)
                sets.append("version = version + 1")
                updated_n += 1
            else:
                identical_n += 1
            conn.execute(f"UPDATE mst_orders SET {', '.join(sets)} WHERE id = ?",
                         vals + [existing["id"]])

        conn.execute(
            """UPDATE mst_import_batches SET committed = 1, committed_at = ?,
               rows_new = ?, rows_updated = ?, rows_identical = ? WHERE id = ?""",
            (stamp, inserted, updated_n, identical_n, batch_id))
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()
    return jsonify({"ok": True, "inserted": inserted, "updated": updated_n,
                    "identical": identical_n, "products_added": products_added,
                    "removed": removed_n})


# ---------------------------------------------------------------- 訂單明細

def _order_query(line, month, extra_where="", extra_params=()):
    where = ["1=1"]
    params = []
    if line:
        where.append("o.line = ?"); params.append(line)
    if month:
        # 沒排日期的單永遠帶出來，不然改期改到空白就找不到它了
        where.append("(o.delivery_date LIKE ? OR o.delivery_date = '' OR o.delivery_date IS NULL)")
        params.append(month + "%")
    if extra_where:
        where.append(extra_where); params.extend(extra_params)
    sql = f"""SELECT o.*, p.box_size AS box_size_master, p.note AS note_master,
                     p.product_name AS name_master
              FROM mst_orders o
              LEFT JOIN mst_products p ON p.line = o.line AND p.barcode = o.barcode
              WHERE {' AND '.join(where)}
              ORDER BY o.delivery_date, o.po_number, o.sku_id"""
    return sql, params


def _decorate(o):
    product = {"box_size": o.get("box_size_master")} if o.get("box_size_master") else None
    box, source = _effective_box(o, product)
    o["box_size"] = box
    o["box_source"] = source
    o["cases"] = _cases(o.get("qty_ship"), box)
    o["cases_file"] = _cases(o.get("qty_file_ship"), box)
    o["month"] = _month_of(o.get("delivery_date"))
    return o


@master_bp.route("/api/master/orders")
def api_orders():
    line = norm_text(request.args.get("line"))
    month = norm_text(request.args.get("month"))
    if month and not _valid_month(month):
        return jsonify({"error": "月份格式要像 2026-09。"}), 400
    dates = [d for d in (request.args.get("dates") or "").split(",") if d]
    pos = [p for p in (request.args.get("pos") or "").split(",") if p]
    brands = [b for b in (request.args.get("brands") or "").split(",") if b]
    warehouses = [w for w in (request.args.get("warehouses") or "").split(",") if w]
    q = norm_text(request.args.get("q"))
    conn = get_conn()
    try:
        sql, params = _order_query(line, month)
        rows = [_decorate(o) for o in _rows(conn.execute(sql, params))]
    finally:
        conn.close()

    # 篩選面（勾選用）先在「線別 + 月份」的範圍上算，不受其他勾選影響，
    # 這樣 OP 勾了一個日期，其他日期的選項不會消失。
    facet_dates = collections.OrderedDict()
    facet_pos = collections.OrderedDict()
    facet_brands = collections.Counter()
    facet_wh = collections.Counter()
    for o in rows:
        facet_wh[o["warehouse"] or ""] += 1
        d = o["delivery_date"] or ""
        fd = facet_dates.setdefault(d, {"date": d, "cases": 0.0, "rows": 0,
                                        "pos": set(), "missing_box": 0})
        fd["rows"] += 1; fd["pos"].add(o["po_number"])
        if o["cases"] is None:
            fd["missing_box"] += 1
        else:
            fd["cases"] += o["cases"]
        fp = facet_pos.setdefault(o["po_number"], {"po_number": o["po_number"],
                                                    "date": d, "rows": 0, "cases": 0.0,
                                                    "warehouse": o["warehouse"],
                                                    "date_overridden": 0})
        fp["rows"] += 1
        fp["cases"] += o["cases"] or 0
        fp["date_overridden"] = fp["date_overridden"] or o["delivery_date_overridden"]
        if o["brand"]:
            facet_brands[o["brand"]] += 1
    for fd in facet_dates.values():
        fd["po_count"] = len(fd.pop("pos"))
        fd["cases"] = round(fd["cases"], 2)
    for fp in facet_pos.values():
        fp["cases"] = round(fp["cases"], 2)

    def keep(o):
        if dates and (o["delivery_date"] or "") not in dates:
            return False
        if pos and o["po_number"] not in pos:
            return False
        if brands and o["brand"] not in brands:
            return False
        if warehouses and (o["warehouse"] or "") not in warehouses:
            return False
        if q:
            hay = " ".join(str(o.get(k) or "") for k in
                           ("po_number", "sku_id", "barcode", "yf_sku", "brand",
                            "product_name", "remarks"))
            if q.lower() not in hay.lower():
                return False
        return True

    shown = [o for o in rows if keep(o)]
    total_cases = round(sum(o["cases"] or 0 for o in shown), 2)
    return jsonify({
        "rows": shown, "count": len(shown), "total_cases": total_cases,
        "facets": {
            "dates": list(facet_dates.values()),
            "pos": list(facet_pos.values()),
            "brands": [{"brand": b, "rows": n} for b, n in sorted(facet_brands.items())],
            "warehouses": [{"warehouse": w, "rows": n} for w, n in sorted(facet_wh.items())],
        },
    })


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
                            "message": f"這筆剛被另一次儲存修改過（{o['updated_at']}），"
                                       "你看到的不是最新版本，請重新載入後再編輯。"}), 409
        sets, vals, changed = [], [], 0
        if "qty_ship" in payload:
            new = norm_int(payload.get("qty_ship"))
            if new is None:
                return jsonify({"error": "出貨數量要是整數。"}), 400
            if not _same(o["qty_ship"], new):
                sets += ["qty_ship = ?", "qty_ship_overridden = 1"]; vals.append(new)
                _log(conn, o["line"], o["po_number"], o["sku_id"], o["barcode"],
                     "qty_ship", "出貨數量", o["qty_ship"], new, operator, "manual")
                changed += 1
        if "remarks" in payload:
            new = norm_text(payload.get("remarks"))
            if not _same(o["remarks"], new):
                sets += ["remarks = ?", "remarks_overridden = 1"]; vals.append(new)
                _log(conn, o["line"], o["po_number"], o["sku_id"], o["barcode"],
                     "remarks", "備註", o["remarks"], new, operator, "manual")
                changed += 1
        if payload.get("reset_qty_ship"):
            file_ship = o["qty_file_ship"] if o["qty_file_ship"] is not None else o["qty_coupang"]
            sets += ["qty_ship = ?", "qty_ship_overridden = 0"]; vals.append(file_ship)
            _log(conn, o["line"], o["po_number"], o["sku_id"], o["barcode"],
                 "qty_ship", "出貨數量", o["qty_ship"], file_ship, operator, "manual",
                 "恢復為整合表數字")
            changed += 1
        if not changed:
            return jsonify({"ok": True, "changed": 0, "row": _decorate(o)})
        sets += ["updated_at = ?", "version = version + 1"]; vals.append(now())
        cur = conn.execute(
            f"UPDATE mst_orders SET {', '.join(sets)} WHERE id = ? AND version = ?",
            vals + [order_id, client_version])
        if cur.rowcount == 0:
            conn.rollback()
            return jsonify({"error": "conflict",
                            "message": "儲存瞬間有其他人也改了這筆，請重新載入。"}), 409
        conn.commit()
        sql, params = _order_query(None, None, "o.id = ?", (order_id,))
        fresh = _decorate(_row(conn.execute(sql, params)))
        return jsonify({"ok": True, "changed": changed, "row": fresh})
    finally:
        conn.close()


@master_bp.route("/api/master/pos/date", methods=["PUT"])
def api_update_po_date():
    """整張 PO 改交貨日：那張單底下所有品項一起搬到新日期。
    帶 expected_date（畫面上看到的舊日期）當防互蓋：資料庫裡已經不是那個
    日期，代表有人先改過，退回請重新載入。"""
    payload = request.get_json(silent=True) or {}
    operator = _operator()
    line = norm_text(payload.get("line"))
    po = norm_key(payload.get("po_number"))
    expected = norm_text(payload.get("expected_date"))
    reset = bool(payload.get("reset"))
    new_date = norm_date(payload.get("delivery_date")) if not reset else None
    if not line or not po:
        return jsonify({"error": "缺少線別或 PO 單號。"}), 400
    if not reset and not new_date:
        return jsonify({"error": "交貨日格式不對，請用 2026-09-17 這種寫法。"}), 400
    conn = get_conn()
    try:
        rows = _rows(conn.execute(
            "SELECT * FROM mst_orders WHERE line = ? AND po_number = ?", (line, po)))
        if not rows:
            return jsonify({"error": "找不到這張 PO。"}), 404
        current = {r["delivery_date"] or "" for r in rows}
        if expected is not None and payload.get("expected_date") is not None:
            if current != {expected}:
                return jsonify({"error": "conflict",
                                "message": "這張單的交貨日剛被別人改過，請重新載入後再改。"}), 409
        stamp = now()
        moved = 0
        for r in rows:
            target = r["delivery_date_file"] if reset else new_date
            flag = 0 if reset else 1
            if _same(r["delivery_date"], target) and r["delivery_date_overridden"] == flag:
                continue
            conn.execute(
                """UPDATE mst_orders SET delivery_date = ?, delivery_date_overridden = ?,
                   updated_at = ?, version = version + 1 WHERE id = ?""",
                (target, flag, stamp, r["id"]))
            _log(conn, line, po, r["sku_id"], r["barcode"], "delivery_date", "交貨日",
                 r["delivery_date"], target, operator, "manual",
                 "恢復為整合表日期" if reset else "整張 PO 改期")
            moved += 1
        conn.commit()
        return jsonify({"ok": True, "moved": moved, "delivery_date": target})
    finally:
        conn.close()


@master_bp.route("/api/master/orders/<int:order_id>", methods=["DELETE"])
def api_delete_order(order_id):
    """刪除單一品項。留給「這筆根本不該在這裡」的情況，會記歷程。"""
    conn = get_conn()
    try:
        o = _row(conn.execute("SELECT * FROM mst_orders WHERE id = ?", (order_id,)))
        if o is None:
            return jsonify({"error": "找不到這筆。"}), 404
        conn.execute("DELETE FROM mst_orders WHERE id = ?", (order_id,))
        _log(conn, o["line"], o["po_number"], o["sku_id"], o["barcode"], "delete",
             "刪除品項", f"出貨 {o['qty_ship']} / 交貨 {o['delivery_date']}", "",
             _operator(), "manual")
        conn.commit()
        return jsonify({"ok": True})
    finally:
        conn.close()


# ---------------------------------------------------------------- 商品主檔

@master_bp.route("/api/master/products")
def api_products():
    line = norm_text(request.args.get("line"))
    conn = get_conn()
    try:
        where, params = "", ()
        if line:
            where, params = "WHERE p.line = ?", (line,)
        rows = _rows(conn.execute(
            f"""SELECT p.*, (SELECT COUNT(*) FROM mst_orders o
                             WHERE o.line = p.line AND o.barcode = p.barcode) AS order_rows
                FROM mst_products p {where}
                ORDER BY p.brand, p.product_name, p.barcode""", params))
        # 訂單裡有、主檔沒有的國條也列出來（讓 OP 一眼看到要補什麼）
        orphan_sql = """SELECT o.line, o.barcode, MAX(o.sku_id) AS sku_id,
                               MAX(o.yf_sku) AS yf_sku, MAX(o.brand) AS brand,
                               MAX(o.product_name) AS product_name,
                               MAX(o.box_size_file) AS box_size_file, COUNT(*) AS order_rows
                        FROM mst_orders o
                        LEFT JOIN mst_products p ON p.line = o.line AND p.barcode = o.barcode
                        WHERE p.id IS NULL AND o.barcode != ''"""
        oparams = ()
        if line:
            orphan_sql += " AND o.line = ?"; oparams = (line,)
        orphan_sql += " GROUP BY o.line, o.barcode ORDER BY o.barcode"
        orphans = _rows(conn.execute(orphan_sql, oparams))
        return jsonify({"products": rows, "orphans": orphans})
    finally:
        conn.close()


@master_bp.route("/api/master/products", methods=["POST"])
def api_save_product():
    payload = request.get_json(silent=True) or {}
    line = norm_text(payload.get("line"))
    barcode = norm_key(payload.get("barcode"))
    if not line or not barcode:
        return jsonify({"error": "線別與國條必填。"}), 400
    box = norm_int(payload.get("box_size"))
    if payload.get("box_size") not in (None, "") and (box is None or box <= 0):
        return jsonify({"error": "箱入數要是正整數。"}), 400
    fields = {
        "sku_id": norm_key(payload.get("sku_id")),
        "yf_sku": norm_key(payload.get("yf_sku")),
        "brand": norm_text(payload.get("brand")),
        "product_name": norm_text(payload.get("product_name")),
        "note": norm_text(payload.get("note")),
    }
    operator = _operator()
    conn = get_conn()
    try:
        existing = _row(conn.execute(
            "SELECT * FROM mst_products WHERE line = ? AND barcode = ?", (line, barcode)))
        stamp = now()
        if existing is None:
            conn.execute(
                """INSERT INTO mst_products
                   (line, barcode, sku_id, yf_sku, brand, product_name, box_size, note,
                    updated_by, updated_at) VALUES (?,?,?,?,?,?,?,?,?,?)""",
                (line, barcode, fields["sku_id"], fields["yf_sku"], fields["brand"],
                 fields["product_name"], box, fields["note"], operator, stamp))
            _log(conn, line, "", fields["sku_id"], barcode, "product_new", "新增主檔",
                 "", box, operator, "manual")
        else:
            if not _same(existing["box_size"], box):
                _log(conn, line, "", existing["sku_id"], barcode, "box_size", "箱入數",
                     existing["box_size"], box, operator, "manual")
            conn.execute(
                """UPDATE mst_products SET sku_id = ?, yf_sku = ?, brand = ?,
                   product_name = ?, box_size = ?, note = ?, updated_by = ?, updated_at = ?
                   WHERE id = ?""",
                (fields["sku_id"], fields["yf_sku"], fields["brand"],
                 fields["product_name"], box, fields["note"], operator, stamp,
                 existing["id"]))
        conn.commit()
        fresh = _row(conn.execute(
            "SELECT * FROM mst_products WHERE line = ? AND barcode = ?", (line, barcode)))
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
        _log(conn, p["line"], "", p["sku_id"], p["barcode"], "product_delete", "刪除主檔",
             p["box_size"], "", _operator(), "manual")
        conn.commit()
        return jsonify({"ok": True})
    finally:
        conn.close()


_HEADER_ALIASES = {
    "barcode": ["barcode", "國條", "條碼", "條碼(國條)", "國際條碼", "ean"],
    "sku_id": ["skuid", "sku id", "sku_id", "skuno.", "sku編號"],
    "yf_sku": ["永豐料號", "料號"],
    "brand": ["brand", "品牌"],
    "product_name": ["sku name", "品名", "product name", "名稱"],
    "box_size": ["箱入數", "箱入", "case pack", "pcs/cs", "轉換率"],
    "note": ["note", "備註", "說明"],
    "line": ["線別", "line"],
}


def _find_columns(ws, want):
    """在前 15 列裡找標題列，用標題文字對欄位（不寫死欄位字母）。
    回傳 (header_row_index, {field: col_index}, headers_list)。"""
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
    """從總表（或任何有「國條 + 箱入數」標題的 Excel）匯入主檔。
    只更新有值的欄位：檔案裡箱入數空白不會把主檔既有的箱入數洗掉。"""
    upload = request.files.get("file")
    line = norm_text(request.form.get("line"))
    if upload is None or not upload.filename:
        return jsonify({"error": "沒有收到檔案。"}), 400
    if not line:
        return jsonify({"error": "請先選線別。"}), 400
    try:
        wb = openpyxl.load_workbook(io.BytesIO(upload.read()), data_only=True, read_only=True)
    except Exception as exc:  # noqa: BLE001
        return jsonify({"error": f"無法開啟 Excel：{exc}"}), 400
    operator = _operator()
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

    conn = get_conn()
    try:
        stamp = now(); added = updated = skipped = 0
        for row in ws.iter_rows(min_row=hdr_idx + 1, values_only=True):
            barcode = norm_key(get(row, "barcode"))
            if not barcode:
                continue
            box = norm_int(get(row, "box_size"))
            vals = {
                "sku_id": norm_key(get(row, "sku_id")), "yf_sku": norm_key(get(row, "yf_sku")),
                "brand": norm_text(get(row, "brand")),
                "product_name": norm_text(get(row, "product_name")),
                "note": norm_text(get(row, "note")),
            }
            existing = _row(conn.execute(
                "SELECT * FROM mst_products WHERE line = ? AND barcode = ?", (line, barcode)))
            if existing is None:
                conn.execute(
                    """INSERT INTO mst_products
                       (line, barcode, sku_id, yf_sku, brand, product_name, box_size, note,
                        updated_by, updated_at) VALUES (?,?,?,?,?,?,?,?,?,?)""",
                    (line, barcode, vals["sku_id"], vals["yf_sku"], vals["brand"],
                     vals["product_name"], box, vals["note"], operator, stamp))
                added += 1
            else:
                sets, params, changed = [], [], False
                for k, v in vals.items():
                    if v and not _same(existing[k], v):
                        sets.append(f"{k} = ?"); params.append(v); changed = True
                if box and not _same(existing["box_size"], box):
                    sets.append("box_size = ?"); params.append(box); changed = True
                    _log(conn, line, "", existing["sku_id"], barcode, "box_size", "箱入數",
                         existing["box_size"], box, operator, "import", upload.filename)
                if changed:
                    sets += ["updated_by = ?", "updated_at = ?"]; params += [operator, stamp]
                    conn.execute(f"UPDATE mst_products SET {', '.join(sets)} WHERE id = ?",
                                 params + [existing["id"]])
                    updated += 1
                else:
                    skipped += 1
        conn.commit()
    finally:
        conn.close()
    return jsonify({"ok": True, "sheet": ws.title, "added": added, "updated": updated,
                    "unchanged": skipped})


# ---------------------------------------------------------------- 月配額

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
    note = norm_text(payload.get("note"))
    operator = _operator()
    conn = get_conn()
    try:
        existing = _row(conn.execute(
            "SELECT * FROM mst_quotas WHERE line = ? AND barcode = ? AND month = ?",
            (line, barcode, month)))
        stamp = now()
        old = existing["qty_cases"] if existing else None
        if existing is None:
            conn.execute(
                """INSERT INTO mst_quotas (line, barcode, month, qty_cases, note, updated_by, updated_at)
                   VALUES (?,?,?,?,?,?,?)""", (line, barcode, month, qty, note, operator, stamp))
        else:
            conn.execute(
                """UPDATE mst_quotas SET qty_cases = ?, note = ?, updated_by = ?, updated_at = ?
                   WHERE id = ?""", (qty, note, operator, stamp, existing["id"]))
        if not _same(old, qty):
            _log(conn, line, "", "", barcode, "quota", f"{month} 配額", old, qty, operator, "manual")
        conn.commit()
        return jsonify({"ok": True, "qty_cases": qty})
    finally:
        conn.close()


@master_bp.route("/api/master/quota/import", methods=["POST"])
def api_import_quota():
    """從總表匯入某個月的配額。第一次呼叫不帶 column 會回傳可選的欄位標題
    （例如 Supply_CS (Sep)），OP 選了再呼叫一次真的寫入。也接受直接貼文字：
    一行一筆「國條 數量」。"""
    line = norm_text(request.form.get("line")); month = norm_text(request.form.get("month"))
    column = norm_text(request.form.get("column"))
    text = request.form.get("text") or ""
    if not line or not _valid_month(month):
        return jsonify({"error": "線別與月份（2026-09）都要有。"}), 400
    operator = _operator()
    pairs = []
    if text.strip():
        for ln in text.splitlines():
            parts = re.split(r"[\s,;\t]+", ln.strip())
            if len(parts) >= 2:
                b = norm_key(parts[0]); q = norm_decimal(parts[1])
                if b and q is not None:
                    pairs.append((b, q))
    else:
        upload = request.files.get("file")
        if upload is None or not upload.filename:
            return jsonify({"error": "沒有收到檔案，也沒有貼上文字。"}), 400
        try:
            wb = openpyxl.load_workbook(io.BytesIO(upload.read()), data_only=True, read_only=True)
        except Exception as exc:  # noqa: BLE001
            return jsonify({"error": f"無法開啟 Excel：{exc}"}), 400
        chosen = None
        for ws in wb.worksheets:
            hdr_idx, mapping, headers = _find_columns(ws, want=("barcode",))
            if hdr_idx:
                chosen = (ws, hdr_idx, mapping, headers); break
        if chosen is None:
            return jsonify({"error": "找不到有「國條／Barcode」標題的工作表。"}), 400
        ws, hdr_idx, mapping, headers = chosen
        numeric_headers = [h for h in headers if h and h.lower() not in ("barcode", "國條")]
        if not column:
            return jsonify({"need_column": True, "sheet": ws.title,
                            "columns": numeric_headers,
                            "suggested": [h for h in numeric_headers if "supply" in h.lower()]})
        if column not in headers:
            return jsonify({"error": f"檔案裡沒有「{column}」這個欄位。"}), 400
        ci = headers.index(column); bi = mapping["barcode"]
        for row in ws.iter_rows(min_row=hdr_idx + 1, values_only=True):
            b = norm_key(row[bi]) if bi < len(row) else ""
            q = norm_decimal(row[ci]) if ci < len(row) else None
            if b and q is not None:
                pairs.append((b, q))
    if not pairs:
        return jsonify({"error": "沒有讀到任何「國條 + 數字」。"}), 400
    conn = get_conn()
    try:
        stamp = now(); written = 0
        for b, q in pairs:
            existing = _row(conn.execute(
                "SELECT * FROM mst_quotas WHERE line = ? AND barcode = ? AND month = ?",
                (line, b, month)))
            if existing is None:
                conn.execute(
                    """INSERT INTO mst_quotas (line, barcode, month, qty_cases, note, updated_by, updated_at)
                       VALUES (?,?,?,?,?,?,?)""", (line, b, month, q, "", operator, stamp))
            elif not _same(existing["qty_cases"], q):
                conn.execute("UPDATE mst_quotas SET qty_cases = ?, updated_by = ?, updated_at = ? WHERE id = ?",
                             (q, operator, stamp, existing["id"]))
                _log(conn, line, "", "", b, "quota", f"{month} 配額", existing["qty_cases"], q,
                     operator, "import", column or "貼上")
            else:
                continue
            written += 1
        conn.commit()
    finally:
        conn.close()
    return jsonify({"ok": True, "written": written, "read": len(pairs)})


# ---------------------------------------------------------------- 總表（現算）

def _build_summary(conn, line, month):
    products = _products_map(conn, line)
    orders = [_decorate(o) for o in _rows(conn.execute(
        """SELECT o.*, p.box_size AS box_size_master
           FROM mst_orders o LEFT JOIN mst_products p ON p.line = o.line AND p.barcode = o.barcode
           WHERE o.line = ? AND o.delivery_date LIKE ?""", (line, month + "%")))]
    quotas = {q["barcode"]: q for q in _rows(conn.execute(
        "SELECT * FROM mst_quotas WHERE line = ? AND month = ?", (line, month)))}

    by_bc = collections.OrderedDict()
    dates = set()
    for o in orders:
        bc = o["barcode"] or f"(無國條) {o['sku_id']}"
        entry = by_bc.setdefault(bc, {"by_date": collections.defaultdict(float),
                                      "month_total": 0.0, "missing_box": 0,
                                      "rows": 0, "name": o["product_name"],
                                      "brand": o["brand"], "sku_id": o["sku_id"],
                                      "yf_sku": o["yf_sku"], "box_size": o["box_size"]})
        entry["rows"] += 1
        dates.add(o["delivery_date"])
        if o["cases"] is None:
            entry["missing_box"] += 1
        else:
            entry["by_date"][o["delivery_date"]] += o["cases"]
            entry["month_total"] += o["cases"]
    dates = sorted(dates)

    rows = []
    barcodes = list(products.keys()) + [b for b in by_bc if b not in products]
    for bc in barcodes:
        p = products.get(bc); e = by_bc.get(bc)
        quota = quotas.get(bc)
        q = quota["qty_cases"] if quota else None
        total = round(e["month_total"], 2) if e else 0.0
        rows.append({
            "barcode": bc,
            "sku_id": (p or e or {}).get("sku_id", ""),
            "yf_sku": (p or e or {}).get("yf_sku", ""),
            "brand": (p["brand"] if p else e["brand"]) if (p or e) else "",
            "product_name": (p["product_name"] if p and p["product_name"] else (e["name"] if e else "")),
            "box_size": p["box_size"] if p and p["box_size"] else (e["box_size"] if e else None),
            "in_master": p is not None,
            "note": p["note"] if p else "",
            "by_date": {d: round(v, 2) for d, v in e["by_date"].items()} if e else {},
            "month_total": total,
            "missing_box": e["missing_box"] if e else 0,
            "order_rows": e["rows"] if e else 0,
            "quota": q,
            "quota_note": quota["note"] if quota else "",
            "remaining": (round(q - total, 2) if q is not None else None),
        })
    totals_by_date = {d: round(sum(r["by_date"].get(d, 0) for r in rows), 2) for d in dates}
    return {
        "line": line, "month": month, "dates": dates, "rows": rows,
        "totals_by_date": totals_by_date,
        "month_total": round(sum(r["month_total"] for r in rows), 2),
        "quota_total": round(sum(r["quota"] or 0 for r in rows), 2),
        "missing_box_rows": sum(r["missing_box"] for r in rows),
        "not_in_master": sum(1 for r in rows if not r["in_master"]),
    }


@master_bp.route("/api/master/summary")
def api_summary():
    line = norm_text(request.args.get("line")); month = norm_text(request.args.get("month")) or _this_month()
    if not line:
        return jsonify({"error": "請選線別。"}), 400
    if not _valid_month(month):
        return jsonify({"error": "月份格式要像 2026-09。"}), 400
    conn = get_conn()
    try:
        return jsonify(_build_summary(conn, line, month))
    finally:
        conn.close()


def _md(date_str):
    """2026-09-17 → 9/17交貨（總表日期欄的寫法）"""
    try:
        d = _dt.date.fromisoformat(date_str)
        return f"{d.month}/{d.day}交貨"
    except ValueError:
        return date_str


@master_bp.route("/api/master/export")
def api_export():
    line = norm_text(request.args.get("line")); month = norm_text(request.args.get("month")) or _this_month()
    if not line or not _valid_month(month):
        return jsonify({"error": "請選線別與月份。"}), 400
    conn = get_conn()
    try:
        s = _build_summary(conn, line, month)
        sql, params = _order_query(line, month)
        orders = [_decorate(o) for o in _rows(conn.execute(sql, params))]
    finally:
        conn.close()

    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    wb = openpyxl.Workbook()
    ws = wb.active; ws.title = "總表"
    mm = int(month[5:7])
    # 配額／剩餘可供貨量兩欄先不放（畫面上也先收起來），要開回來時把 r["quota"]、
    # r["remaining"] 加回這裡與 head 即可，後端 _build_summary 一直有算。
    head = ["國條", "SKU ID", "永豐料號", "品牌", "品名", "箱入數"] + \
           [_md(d) for d in s["dates"]] + \
           [f"{mm}月TTL下單總箱數", "備註"]
    ws.append(head)
    for r in s["rows"]:
        ws.append([r["barcode"], r["sku_id"], r["yf_sku"], r["brand"], r["product_name"],
                   r["box_size"]] +
                  [r["by_date"].get(d) for d in s["dates"]] +
                  [r["month_total"], r["note"]])
    ws.append([])
    ws.append(["合計", "", "", "", "", ""] + [s["totals_by_date"].get(d) for d in s["dates"]] +
              [s["month_total"], ""])
    bold = Font(bold=True); fill = PatternFill("solid", fgColor="DBEAFE")
    for c in ws[1]:
        c.font = bold; c.fill = fill; c.alignment = Alignment(horizontal="center", wrap_text=True)
    for c in ws[ws.max_row]:
        c.font = bold
    ws.freeze_panes = "G2"
    for i, h in enumerate(head, start=1):
        ws.column_dimensions[get_column_letter(i)].width = 14 if i not in (5,) else 40
    for row in ws.iter_rows(min_row=2, min_col=1, max_col=3):
        for c in row:
            c.number_format = "@"

    ws2 = wb.create_sheet("每日合計")
    ws2.append(["交貨日", "出貨總箱數", "PO 數", "品項數"])
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

    ws3 = wb.create_sheet("訂單明細")
    ws3.append(["交貨日", "PO單號", "倉別", "SKU ID", "國條", "永豐料號", "品牌", "品名",
                "下單數量", "出貨數量", "單位", "箱入數", "出貨數量(箱)", "下單單價",
                "備註", "出貨數量人工調整", "交貨日人工調整"])
    for o in orders:
        ws3.append([o["delivery_date"], o["po_number"], o["warehouse"], o["sku_id"], o["barcode"],
                    o["yf_sku"], o["brand"], o["product_name"], o["qty_coupang"], o["qty_ship"],
                    o["unit"], o["box_size"], o["cases"], o["unit_price"], o["remarks"],
                    "是" if o["qty_ship_overridden"] else "", "是" if o["delivery_date_overridden"] else ""])
    for c in ws3[1]:
        c.font = bold; c.fill = fill
    ws3.freeze_panes = "A2"
    for row in ws3.iter_rows(min_row=2, min_col=2, max_col=6):
        for c in row:
            c.number_format = "@"

    out = io.BytesIO(); wb.save(out); out.seek(0)
    fname = f"{line}_總表_{month}.xlsx"
    return send_file(out, as_attachment=True, download_name=fname,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")



@master_bp.route("/api/master/export/daily")
def api_export_daily():
    """匯出成同事熟悉的「專案報價檔」長相：一個交貨日一個分頁（分頁名像
    0724交貨），欄位順序照報價檔，同一張 PO 的品項放一起，PO 之間用一列黃色
    空白列隔開，PO 單號寫在 R 欄（報價檔就是這樣放的，標題雖然叫「交貨日」
    但裡面一直是 PO 單號＋日期＋倉別）。

    價格欄：報價檔的「單價(含稅)」原本是 VLOOKUP 主檔的業務報價，這個模組
    沒有存業務報價，所以「單價(含稅)」跟「酷澎下單價(含稅)」都填整合表的
    酷澎下單單價；箱單價 = 單價 × 箱入數，總計 = 箱單價 × 出貨箱數。"""
    line = norm_text(request.args.get("line")); month = norm_text(request.args.get("month")) or _this_month()
    if not line or not _valid_month(month):
        return jsonify({"error": "請選線別與月份。"}), 400
    conn = get_conn()
    try:
        sql, params = _order_query(line, month)
        orders = [_decorate(o) for o in _rows(conn.execute(sql, params))]
    finally:
        conn.close()

    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    head = ["SKU ID", "永豐料號", "國條", "品類", "品牌", "品名", "下單數量(酷澎單位)",
            "出貨數量", "箱入數", "出貨數量(箱)", "單價(含稅)", "酷澎下單價(含稅)",
            "箱單價(含稅)", "總計(含稅)", "備註", "驗收完成請打勾", "簽單完成請打勾", "交貨日"]
    yellow = PatternFill("solid", fgColor="FFFF00")
    head_fill = PatternFill("solid", fgColor="F8CBAD")
    bold = Font(bold=True)

    by_date = collections.OrderedDict()
    for o in orders:
        by_date.setdefault(o["delivery_date"] or "", collections.OrderedDict()) \
               .setdefault(o["po_number"], []).append(o)

    wb = openpyxl.Workbook(); wb.remove(wb.active)
    # 報價檔的分頁順序是最新日期在最前面
    for date in sorted(by_date.keys(), key=lambda d: d or "0000", reverse=True):
        pos = by_date[date]
        if date:
            d = _dt.date.fromisoformat(date)
            title = f"{d.month:02d}{d.day:02d}交貨"
            label_date = f"{d.month}/{d.day}交貨"
        else:
            title, label_date = "未排日期", ""
        ws = wb.create_sheet(title[:31])
        ws.append(head)
        for c in ws[1]:
            c.font = bold; c.fill = head_fill
        first_group = True
        for po, rows in pos.items():
            if not first_group:
                ws.append([None] * len(head))
                for c in ws[ws.max_row]:
                    c.fill = yellow
            first_group = False
            for i, o in enumerate(rows):
                box = o["box_size"]; price = o["unit_price"]
                box_price = (price * box) if (price is not None and box) else None
                total = (box_price * o["cases"]) if (box_price is not None and o["cases"] is not None) else None
                po_cell = (f"{po}_{label_date}({o['warehouse']})" if o["warehouse"] else f"{po}_{label_date}") if i == 0 else None
                ws.append([o["sku_id"], o["yf_sku"], o["barcode"], "", o["brand"], o["product_name"],
                           o["qty_coupang"], o["qty_ship"], box, o["cases"], price, price,
                           box_price, total, o["remarks"], "", "", po_cell])
        ws.freeze_panes = "A2"
        widths = [16, 15, 15, 8, 14, 44, 10, 9, 8, 11, 10, 12, 11, 12, 18, 8, 8, 30]
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w
        for row in ws.iter_rows(min_row=2, min_col=1, max_col=3):
            for c in row:
                c.number_format = "@"
        for row in ws.iter_rows(min_row=2, min_col=18, max_col=18):
            for c in row:
                c.alignment = Alignment(horizontal="left")

    if not wb.sheetnames:
        ws = wb.create_sheet("無資料"); ws.append(["這個月沒有任何訂單"])

    out = io.BytesIO(); wb.save(out); out.seek(0)
    fname = f"{line}_專案報價檔格式_{month}.xlsx"
    return send_file(out, as_attachment=True, download_name=fname,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ---------------------------------------------------------------- 歷程

@master_bp.route("/api/master/logs")
def api_logs():
    line = norm_text(request.args.get("line")); po = norm_key(request.args.get("po"))
    barcode = norm_key(request.args.get("barcode"))
    limit = min(norm_int(request.args.get("limit")) or 200, 1000)
    where, params = ["1=1"], []
    if line:
        where.append("line = ?"); params.append(line)
    if po:
        where.append("po_number = ?"); params.append(po)
    if barcode:
        where.append("barcode = ?"); params.append(barcode)
    conn = get_conn()
    try:
        rows = _rows(conn.execute(
            f"SELECT * FROM mst_logs WHERE {' AND '.join(where)} ORDER BY changed_at DESC, id DESC LIMIT ?",
            params + [limit]))
        return jsonify({"logs": rows})
    finally:
        conn.close()

"""酷澎訂單管理系統 — 本地端中繼站。

啟動：python app.py  然後開 http://127.0.0.1:5000
"""

import io
import json
import os
import datetime as _dt

from flask import Flask, jsonify, render_template, request, send_file

import db
from db import get_conn, init_db, now
from importer import (
    COUPANG_FIELDS, CRITICAL_FIELDS, ImportError_, diff_rows, parse_workbook,
)
from normalize import norm_date, norm_int, norm_key, norm_text, norm_warehouse

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# 每次改版就手動往上加一號。畫面右下角跟啟動視窗都會印出這個號碼，
# 用來確認「現在看到的畫面」跟「最新給的檔案」是不是同一份——
# 之前吃過虧：舊的黑視窗沒關乾淨，背景還留著一個沒更新到的伺服器
# 在跑，怎麼換檔案畫面都不會變，肉眼完全看不出來是這個原因。
BUILD_VERSION = "2026-08-19.7"

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 64 * 1024 * 1024
# 模板一律即時讀檔，不要用啟動當下快取的舊版本（debug=False 時 Flask
# 預設不會自動重讀模板檔，只改 index.html 卻要重開程式才生效，容易
# 造成「明明換了檔案，畫面卻沒變」的誤判）。
app.config["TEMPLATES_AUTO_RELOAD"] = True

# 單一 SKU 自己的欄位，改一個品項不影響同張單的其他品項。
# 匯入永遠不會覆蓋這些欄位。
EDITABLE_FIELDS = {
    "qty_ship":       "出貨數量",
    "remarks":        "備註",
    "receiving_note": "驗收註記",
}

# 整張 PO 共用的欄位，改了就是整張單一起改（OP 拋 ERP、交倉庫都是整張
# 單一起行動，不會拆開）。存在 po_headers，匯入一律不覆蓋。
PO_EDITABLE_FIELDS = {
    "po_status":        "PO狀態",
    "receiving_status": "驗收狀態",
    "filed_date":       "建檔日",
}

# 整張 PO 共用、但屬於酷澎來源的欄位：OP 可以改（跟酷澎談好調整），
# 改的時候整張單一起改，但它們存在 orders 上、且會參與匯入比對。
PO_COUPANG_FIELDS = {
    "delivery_date": "交期",
    "warehouse":     "倉別",
}

# 酷澎來源欄位中，允許 OP「補空白」的欄位。
# 只有原本是空白時才能填，已經有值的一律不給改——避免有人手滑把瑪氏
# 改成寶僑。上游主檔之後補建了，匯入仍會以主檔為準覆蓋回來並記歷程。
FILLABLE_FIELDS = {
    "line":  "線別",
    "brand": "品牌",
}

INSERT_FIELDS = [
    "po_number", "sku_id", "order_type", "parent_po", "line", "brand",
    "product_name", "barcode", "yf_sku", "warehouse", "address",
    "delivery_date", "qty_coupang", "unit", "box_size", "unit_price",
    "expiry_note", "seq_no",
]


# ---------------------------------------------------------------- 設定檔

def load_json(name, default):
    # 設定檔住在資料資料夾，不在程式資料夾——更新版本時不會被覆蓋
    path = os.path.join(db.DATA_DIR, name)
    try:
        with open(path, encoding="utf-8") as fh:
            return json.load(fh)
    except (OSError, ValueError):
        return default


def get_config():
    return load_json("config.json", {
        "operators": ["OP"],
        "po_statuses": ["已建立", "已回覆", "處理中", "已完成", "修改中", "已取消"],
        "receiving_statuses": ["未驗收", "完成", "異常", "重啟"],
        "order_types": ["一般", "NS", "補單", "拆單"],
    })


def get_profiles():
    return load_json("export_profiles.json", {"profiles": {}}).get("profiles", {})


# ---------------------------------------------------------------- 歷程

def log_change(conn, order, field, label, old, new, operator, source, note=""):
    conn.execute(
        """INSERT INTO edit_logs
           (order_id, po_number, sku_id, field, field_label, old_value,
            new_value, operator, source, note, changed_at)
           VALUES (?,?,?,?,?,?,?,?,?,?,?)""",
        (order["id"], order["po_number"], order["sku_id"], field, label,
         "" if old is None else str(old), "" if new is None else str(new),
         operator, source, note, now()),
    )


# ---------------------------------------------------------------- 頁面

@app.route("/")
def index():
    return render_template("index.html", build_version=BUILD_VERSION)


@app.route("/api/config")
def api_config():
    cfg = get_config()
    profiles = get_profiles()
    conn = get_conn()
    try:
        def distinct(column):
            rows = conn.execute(
                f"SELECT DISTINCT {column} AS v FROM order_rows "
                f"WHERE {column} IS NOT NULL AND {column} != '' ORDER BY v"
            ).fetchall()
            return [r["v"] for r in rows]

        return jsonify({
            "operators": cfg.get("operators", ["OP"]),
            "po_statuses": cfg.get("po_statuses", []),
            "receiving_statuses": cfg.get("receiving_statuses", []),
            "order_types": cfg.get("order_types", []),
            "brands": distinct("brand"),
            "lines": distinct("line"),
            "warehouses": distinct("warehouse"),
            "export_profiles": [
                {"key": k, "label": v.get("label", k), "note": v.get("note", ""),
                 "mark_pulled_default": bool(v.get("mark_pulled_default", False)),
                 "column_count": len(v.get("columns", []))}
                for k, v in profiles.items()
            ],
        })
    finally:
        conn.close()


# ---------------------------------------------------------------- 查詢

def build_filter(args):
    where, params = [], []

    def add(clause, *values):
        where.append(clause)
        params.extend(values)

    po = norm_text(args.get("po_number"))
    if po:
        add("po_number LIKE ?", f"%{po}%")

    keyword = norm_text(args.get("keyword"))
    if keyword:
        add("(product_name LIKE ? OR yf_sku LIKE ? OR sku_id LIKE ? OR barcode LIKE ?)",
            *[f"%{keyword}%"] * 4)

    for column, key in (("brand", "brand"), ("line", "line"),
                        ("warehouse", "warehouse"), ("po_status", "po_status"),
                        ("receiving_status", "receiving_status"),
                        ("order_type", "order_type")):
        value = norm_text(args.get(key))
        if value:
            add(f"{column} = ?", value)

    date_from = norm_date(args.get("date_from"))
    if date_from:
        add("delivery_date >= ?", date_from)
    date_to = norm_date(args.get("date_to"))
    if date_to:
        add("delivery_date <= ?", date_to)

    pulled = args.get("is_pulled")
    if pulled in ("0", "1"):
        add("is_pulled = ?", int(pulled))

    if args.get("needs_review") == "1":
        add("needs_review = 1")

    if args.get("alert_level"):
        add("alert_level = ?", norm_text(args.get("alert_level")))

    clause = (" WHERE " + " AND ".join(where)) if where else ""
    return clause, params


@app.route("/api/orders")
def api_orders():
    clause, params = build_filter(request.args)
    page = max(1, norm_int(request.args.get("page")) or 1)
    size = min(500, max(10, norm_int(request.args.get("page_size")) or 100))

    conn = get_conn()
    try:
        total = conn.execute(
            f"SELECT COUNT(*) AS c FROM order_rows{clause}", params
        ).fetchone()["c"]

        summary = conn.execute(
            f"""SELECT COALESCE(SUM(qty_ship),0) AS qty,
                       SUM(CASE WHEN needs_review=1 THEN 1 ELSE 0 END) AS review,
                       SUM(CASE WHEN alert_level='changed_after_pull' THEN 1 ELSE 0 END) AS after_pull,
                       SUM(CASE WHEN is_pulled=1 THEN 1 ELSE 0 END) AS pulled
                FROM order_rows{clause}""",
            params,
        ).fetchone()

        rows = conn.execute(
            f"""SELECT * FROM order_rows{clause}
                ORDER BY needs_review DESC,
                         CASE alert_level WHEN 'changed_after_pull' THEN 0
                                          WHEN 'changed' THEN 1 ELSE 2 END,
                         delivery_date, po_number, seq_no, sku_id
                LIMIT ? OFFSET ?""",
            params + [size, (page - 1) * size],
        ).fetchall()

        return jsonify({
            "total": total,
            "page": page,
            "page_size": size,
            "summary": {
                "qty_ship": summary["qty"] or 0,
                "needs_review": summary["review"] or 0,
                "changed_after_pull": summary["after_pull"] or 0,
                "pulled": summary["pulled"] or 0,
            },
            "rows": [dict(r) for r in rows],
        })
    finally:
        conn.close()


@app.route("/api/orders/<int:order_id>/logs")
def api_order_logs(order_id):
    conn = get_conn()
    try:
        rows = conn.execute(
            "SELECT * FROM edit_logs WHERE order_id = ? ORDER BY changed_at DESC, id DESC",
            (order_id,),
        ).fetchall()
        return jsonify([dict(r) for r in rows])
    finally:
        conn.close()


def build_log_filter(args):
    """歷程查詢條件。日期用 changed_at 前綴比對，避免時分秒干擾。"""
    where, params = [], []

    def add(clause, *values):
        where.append(clause)
        params.extend(values)

    if norm_text(args.get("po_number")):
        add("po_number LIKE ?", f"%{norm_text(args.get('po_number'))}%")
    if norm_text(args.get("field")):
        add("field = ?", norm_text(args.get("field")))
    if norm_text(args.get("operator")):
        add("operator = ?", norm_text(args.get("operator")))
    if norm_text(args.get("source")):
        add("source = ?", norm_text(args.get("source")))
    if norm_text(args.get("keyword")):
        kw = f"%{norm_text(args.get('keyword'))}%"
        add("(old_value LIKE ? OR new_value LIKE ? OR note LIKE ? OR sku_id LIKE ?)",
            kw, kw, kw, kw)
    if norm_date(args.get("date_from")):
        add("changed_at >= ?", norm_date(args.get("date_from")) + " 00:00:00")
    if norm_date(args.get("date_to")):
        add("changed_at <= ?", norm_date(args.get("date_to")) + " 23:59:59")

    return ((" WHERE " + " AND ".join(where)) if where else ""), params


@app.route("/api/logs")
def api_logs():
    """全域歷程檢視：誰在什麼時候改了什麼，可依單號／欄位／人員／期間篩。"""
    clause, params = build_log_filter(request.args)
    page = max(1, norm_int(request.args.get("page")) or 1)
    size = min(500, max(10, norm_int(request.args.get("page_size")) or 100))

    conn = get_conn()
    try:
        total = conn.execute(
            f"SELECT COUNT(*) AS c FROM edit_logs{clause}", params).fetchone()["c"]
        rows = conn.execute(
            f"""SELECT * FROM edit_logs{clause}
                ORDER BY changed_at DESC, id DESC LIMIT ? OFFSET ?""",
            params + [size, (page - 1) * size]).fetchall()

        # 篩選面板用的選項，只列真的出現過的值
        fields = conn.execute(
            "SELECT DISTINCT field, field_label FROM edit_logs ORDER BY field_label"
        ).fetchall()
        operators = conn.execute(
            "SELECT DISTINCT operator FROM edit_logs WHERE operator != '' ORDER BY operator"
        ).fetchall()

        return jsonify({
            "total": total, "page": page, "page_size": size,
            "rows": [dict(r) for r in rows],
            "fields": [{"field": f["field"], "label": f["field_label"]} for f in fields],
            "operators": [o["operator"] for o in operators],
        })
    finally:
        conn.close()


@app.route("/api/logs/export", methods=["POST"])
def api_logs_export():
    """把目前篩選出來的歷程匯成 Excel。

    這張表最終是拿去跟酷澎對帳、釐清倉庫出錯責任用的，所以要能整份帶走，
    不能只留在畫面上一筆一筆看。
    """
    payload = request.get_json(silent=True) or {}
    operator = norm_text(payload.get("operator"))
    if not operator:
        return jsonify({"error": "請先選擇操作人員。"}), 400

    clause, params = build_log_filter(payload.get("filters") or {})
    conn = get_conn()
    try:
        rows = conn.execute(
            f"SELECT * FROM edit_logs{clause} ORDER BY changed_at DESC, id DESC "
            f"LIMIT 50000", params).fetchall()
    finally:
        conn.close()

    if not rows:
        return jsonify({"error": "目前的篩選條件沒有任何歷程可以匯出。"}), 400

    import openpyxl
    from openpyxl.styles import Alignment, Font, PatternFill

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "修改歷程"
    headers = ["時間", "PO 單號", "SKU ID", "層級", "欄位", "改前", "改後",
               "操作人員", "來源", "備註"]
    ws.append(headers)
    fill = PatternFill("solid", fgColor="161D29")
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

    source_zh = {"import": "匯入", "manual": "手動", "system": "系統"}
    for r in rows:
        ws.append([
            r["changed_at"], r["po_number"], r["sku_id"],
            "整張單" if not r["sku_id"] else "單一品項",
            r["field_label"], r["old_value"], r["new_value"],
            r["operator"], source_zh.get(r["source"], r["source"]), r["note"],
        ])

    for idx, width in enumerate([19, 18, 18, 10, 12, 22, 22, 10, 8, 30], start=1):
        letter = ws.cell(row=1, column=idx).column_letter
        ws.column_dimensions[letter].width = width
        if idx in (2, 3):                      # 長數字一律當文字，避免科學記號
            for cell in ws[letter][1:]:
                cell.number_format = "@"
    ws.freeze_panes = "A2"

    stamp = _dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    stream = io.BytesIO()
    wb.save(stream)
    stream.seek(0)
    return send_file(
        stream, as_attachment=True, download_name=f"修改歷程_{stamp}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ---------------------------------------------------------------- 線上編輯

@app.route("/api/orders/<int:order_id>", methods=["PUT"])
def api_update_order(order_id):
    payload = request.get_json(silent=True) or {}
    operator = norm_text(payload.get("operator"))
    if not operator:
        return jsonify({"error": "請先在右上角選擇操作人員，才能編輯訂單。"}), 400

    client_version = norm_int(payload.get("version"))
    if client_version is None:
        return jsonify({"error": "缺少版本資訊，請重新整理後再試。"}), 400

    conn = get_conn()
    try:
        order = conn.execute("SELECT * FROM order_rows WHERE id = ?", (order_id,)).fetchone()
        if order is None:
            return jsonify({"error": "找不到這筆訂單。"}), 404
        order = dict(order)

        # 樂觀鎖：三個人共用一份資料，last-write-wins 會無聲蓋掉同事的修改
        if order["version"] != client_version:
            return jsonify({
                "error": "conflict",
                "message": (
                    f"這筆訂單剛剛被「{order['updated_at']}」的另一次儲存修改過，"
                    "你看到的已經不是最新版本。請重新載入後再編輯。"
                ),
                "current": order,
            }), 409

        if order["is_pulled"] and not payload.get("force_edit"):
            return jsonify({
                "error": "locked",
                "message": "這筆訂單已拉單並鎖定。若確實需要修改，請先解除拉單鎖定。",
            }), 423

        changes = []
        for field, label in EDITABLE_FIELDS.items():
            if field not in payload:
                continue
            raw = payload[field]
            new_val = norm_int(raw) if field == "qty_ship" else norm_text(raw)

            old_val = order[field]
            if _same(old_val, new_val):
                continue
            changes.append((field, label, old_val, new_val))

        # 補空白：只有原本沒有值的欄位才收，已經有值的直接忽略
        for field, label in FILLABLE_FIELDS.items():
            if field not in payload:
                continue
            new_val = norm_text(payload[field])
            if not new_val:
                continue
            if norm_text(order[field]):
                return jsonify({
                    "error": "readonly",
                    "message": (
                        f"「{label}」已經有值（{order[field]}），不開放修改。"
                        "這個欄位以酷澎整合表為準，需要更正請回頭修主檔後重新上傳。"
                    ),
                }), 400
            changes.append((field, label, order[field], new_val))

        if not changes:
            return jsonify({"ok": True, "changed": 0, "order": order})

        sets = ", ".join(f"{f} = ?" for f, _, _, _ in changes)
        values = [v for _, _, _, v in changes]
        cur = conn.execute(
            f"""UPDATE orders SET {sets}, updated_at = ?, version = version + 1
                WHERE id = ? AND version = ?""",
            values + [now(), order_id, client_version],
        )
        if cur.rowcount == 0:
            conn.rollback()
            return jsonify({"error": "conflict",
                            "message": "儲存瞬間有其他人也改了這筆，請重新載入。"}), 409

        for field, label, old_val, new_val in changes:
            log_change(conn, order, field, label, old_val, new_val, operator, "manual")
            if field == "qty_ship":
                conn.execute("UPDATE orders SET qty_ship_overridden = 1 WHERE id = ?",
                             (order_id,))
        conn.commit()

        updated = dict(conn.execute("SELECT * FROM order_rows WHERE id = ?",
                                    (order_id,)).fetchone())
        return jsonify({"ok": True, "changed": len(changes), "order": updated})
    finally:
        conn.close()


# ---------------------------------------------------------------- PO 層級

def log_po_change(conn, po_number, field, label, old, new, operator,
                  source="manual", note=""):
    """整張 PO 的變更只記一筆，sku_id 留空代表『這是整張單的事』。"""
    conn.execute(
        """INSERT INTO edit_logs
           (order_id, po_number, sku_id, field, field_label, old_value,
            new_value, operator, source, note, changed_at)
           VALUES (0,?,'',?,?,?,?,?,?,?,?)""",
        (po_number, field, label,
         "" if old is None else str(old), "" if new is None else str(new),
         operator, source, note, now()),
    )


def po_summary_sql(clause):
    """把 SKU 收攏成一列一張 PO。

    交期／倉別在同一張 PO 底下本來就一致（實際資料驗證過），用 MIN 取值
    即可；品牌則可能混多個，用逗號串起來，這樣首頁篩品牌時混單也篩得到。

    篩選條件只用來決定「哪幾張 PO 入選」，統計值一律涵蓋該 PO 底下的
    全部品項。否則搜「whiskas」時，那張含 5 個品牌的單會只顯示 whiskas、
    品項數與數量也只算到符合的那幾項，資訊是錯的。
    """
    return f"""
        SELECT po_number,
               MIN(order_type)      AS order_type,
               MIN(parent_po)       AS parent_po,
               MIN(po_status)       AS po_status,
               MIN(receiving_status) AS receiving_status,
               MAX(is_pulled)       AS is_pulled,
               MIN(pulled_at)       AS pulled_at,
               MIN(pulled_by)       AS pulled_by,
               MIN(filed_date)      AS filed_date,
               MIN(delivery_date)   AS delivery_date,
               MIN(warehouse)       AS warehouse,
               GROUP_CONCAT(DISTINCT line)  AS lines_csv,
               GROUP_CONCAT(DISTINCT brand) AS brands_csv,
               COUNT(*)             AS sku_count,
               COALESCE(SUM(qty_coupang),0) AS qty_coupang,
               COALESCE(SUM(qty_ship),0)    AS qty_ship,
               SUM(CASE WHEN needs_review=1 THEN 1 ELSE 0 END) AS review_count,
               SUM(CASE WHEN alert_level='changed_after_pull' THEN 1 ELSE 0 END)
                                            AS after_pull_count,
               MIN(po_version)      AS po_version
        FROM order_rows
        WHERE po_number IN (SELECT po_number FROM order_rows{clause})
        GROUP BY po_number
    """


def _csv_clean(value):
    """GROUP_CONCAT 出來的字串去掉空值、排序後用逗號串好。"""
    if not value:
        return ""
    parts = sorted({p.strip() for p in value.split(",") if p.strip()})
    return ", ".join(parts)


@app.route("/api/pos")
def api_pos():
    """首頁：一列一張 PO。"""
    clause, params = build_filter(request.args)
    page = max(1, norm_int(request.args.get("page")) or 1)
    size = min(500, max(10, norm_int(request.args.get("page_size")) or 50))

    conn = get_conn()
    try:
        base = po_summary_sql(clause)
        total = conn.execute(
            f"SELECT COUNT(*) AS c FROM ({base})", params).fetchone()["c"]

        summary = conn.execute(
            f"""SELECT COALESCE(SUM(qty_ship),0) AS qty,
                       SUM(CASE WHEN review_count>0 THEN 1 ELSE 0 END) AS review,
                       SUM(CASE WHEN after_pull_count>0 THEN 1 ELSE 0 END) AS after_pull,
                       SUM(CASE WHEN is_pulled=1 THEN 1 ELSE 0 END) AS pulled
                FROM ({base})""", params).fetchone()

        rows = conn.execute(
            f"""SELECT * FROM ({base})
                ORDER BY after_pull_count DESC, review_count DESC,
                         delivery_date, po_number
                LIMIT ? OFFSET ?""",
            params + [size, (page - 1) * size]).fetchall()

        out = []
        for r in rows:
            item = dict(r)
            item["brands"] = _csv_clean(item.pop("brands_csv"))
            item["lines"] = _csv_clean(item.pop("lines_csv"))
            out.append(item)

        return jsonify({
            "total": total, "page": page, "page_size": size,
            "summary": {
                "qty_ship": summary["qty"] or 0,
                "needs_review": summary["review"] or 0,
                "changed_after_pull": summary["after_pull"] or 0,
                "pulled": summary["pulled"] or 0,
            },
            "rows": out,
        })
    finally:
        conn.close()


@app.route("/api/pos/<po_number>")
def api_po_detail(po_number):
    """點開一張 PO：表頭 + 這張單所有 SKU + 整張單的歷程。"""
    po_number = norm_key(po_number)
    conn = get_conn()
    try:
        header = conn.execute(
            "SELECT * FROM po_headers WHERE po_number = ?", (po_number,)).fetchone()
        if header is None:
            return jsonify({"error": "找不到這張 PO。"}), 404

        skus = conn.execute(
            """SELECT * FROM order_rows WHERE po_number = ?
               ORDER BY seq_no, sku_id""", (po_number,)).fetchall()
        logs = conn.execute(
            """SELECT * FROM edit_logs WHERE po_number = ?
               ORDER BY changed_at DESC, id DESC LIMIT 300""", (po_number,)).fetchall()

        skus = [dict(s) for s in skus]
        return jsonify({
            "header": dict(header),
            "brands": _csv_clean(",".join(s["brand"] or "" for s in skus)),
            "lines": _csv_clean(",".join(s["line"] or "" for s in skus)),
            "delivery_date": skus[0]["delivery_date"] if skus else "",
            "warehouse": skus[0]["warehouse"] if skus else "",
            "skus": skus,
            "logs": [dict(l) for l in logs],
        })
    finally:
        conn.close()


@app.route("/api/pos/<po_number>", methods=["PUT"])
def api_update_po(po_number):
    """改整張 PO：狀態類欄位寫 po_headers，交期／倉別套用到全部 SKU。"""
    po_number = norm_key(po_number)
    payload = request.get_json(silent=True) or {}
    operator = norm_text(payload.get("operator"))
    if not operator:
        return jsonify({"error": "請先在右上角選擇操作人員，才能編輯訂單。"}), 400

    client_version = norm_int(payload.get("po_version"))
    if client_version is None:
        return jsonify({"error": "缺少版本資訊，請重新整理後再試。"}), 400

    conn = get_conn()
    try:
        header = conn.execute(
            "SELECT * FROM po_headers WHERE po_number = ?", (po_number,)).fetchone()
        if header is None:
            return jsonify({"error": "找不到這張 PO。"}), 404
        header = dict(header)

        if header["version"] != client_version:
            return jsonify({
                "error": "conflict",
                "message": (f"這張單剛剛被另一次儲存修改過（{header['updated_at']}），"
                            "你看到的不是最新版本，請重新載入後再編輯。"),
            }), 409

        if header["is_pulled"] and not payload.get("force_edit"):
            return jsonify({
                "error": "locked",
                "message": "這張單已拉單並鎖定。若確實需要修改，請先解除拉單鎖定。",
            }), 423

        stamp = now()
        head_changes = []
        for field, label in PO_EDITABLE_FIELDS.items():
            if field not in payload:
                continue
            new_val = (norm_date(payload[field]) if field == "filed_date"
                       else norm_text(payload[field]))
            if _same(header[field], new_val):
                continue
            head_changes.append((field, label, header[field], new_val))

        # 交期／倉別存在每一列 SKU 上（要參與匯入比對），但 OP 是整張單一起改
        row_changes = []
        first = conn.execute(
            "SELECT delivery_date, warehouse FROM orders WHERE po_number = ? LIMIT 1",
            (po_number,)).fetchone()
        if first is not None:
            for field, label in PO_COUPANG_FIELDS.items():
                if field not in payload:
                    continue
                new_val = (norm_date(payload[field]) if field == "delivery_date"
                           else norm_warehouse(payload[field]))
                if _same(first[field], new_val):
                    continue
                row_changes.append((field, label, first[field], new_val))

        if not head_changes and not row_changes:
            return jsonify({"ok": True, "changed": 0})

        if head_changes:
            sets = ", ".join(f"{f} = ?" for f, _, _, _ in head_changes)
            cur = conn.execute(
                f"""UPDATE po_headers SET {sets}, updated_at = ?, version = version + 1
                    WHERE po_number = ? AND version = ?""",
                [v for _, _, _, v in head_changes] + [stamp, po_number, client_version])
            if cur.rowcount == 0:
                conn.rollback()
                return jsonify({"error": "conflict",
                                "message": "儲存瞬間有其他人也改了這張單，請重新載入。"}), 409

        for field, label, old_val, new_val in row_changes:
            # 一併標記「這個欄位人工調整過」，之後匯入就不再覆蓋它
            conn.execute(
                f"""UPDATE orders SET {field} = ?, {field}_overridden = 1,
                        updated_at = ?, version = version + 1
                    WHERE po_number = ?""", (new_val, stamp, po_number))

        for field, label, old_val, new_val in head_changes + row_changes:
            log_po_change(conn, po_number, field, label, old_val, new_val, operator)
        conn.commit()
        return jsonify({"ok": True, "changed": len(head_changes) + len(row_changes)})
    finally:
        conn.close()


@app.route("/api/pos/review", methods=["POST"])
def api_clear_review():
    """OP 確認過異動後把整張單的警示燈關掉，每張單各記一筆歷程。"""
    payload = request.get_json(silent=True) or {}
    operator = norm_text(payload.get("operator"))
    pos = [norm_key(p) for p in payload.get("po_numbers", []) if norm_key(p)]
    if not operator:
        return jsonify({"error": "請先選擇操作人員。"}), 400
    if not pos:
        return jsonify({"error": "沒有選取任何訂單。"}), 400

    conn = get_conn()
    try:
        placeholders = ",".join("?" * len(pos))
        rows = conn.execute(
            f"""SELECT po_number, COUNT(*) AS n FROM orders
                WHERE po_number IN ({placeholders}) AND needs_review = 1
                GROUP BY po_number""", pos).fetchall()
        for row in rows:
            log_po_change(conn, row["po_number"], "needs_review", "異動確認",
                          f"{row['n']} 項待確認", "已確認", operator)
        conn.execute(
            f"""UPDATE orders SET needs_review = 0, alert_level = '',
                   review_reason = '', updated_at = ?, version = version + 1
                WHERE po_number IN ({placeholders}) AND needs_review = 1""",
            [now()] + pos)
        conn.commit()
        return jsonify({"ok": True, "count": len(rows)})
    finally:
        conn.close()


@app.route("/api/pos/pull", methods=["POST"])
def api_set_pulled():
    """手動調整整張單的拉單狀態。

    拉單代表「已經正式拋 ERP、交給倉庫」，是對外承諾，所以不管是標記
    還是解除，都必須填原因並寫進歷程——事後倉庫拿到的是哪一版、為什麼
    重出，要查得到。
    """
    payload = request.get_json(silent=True) or {}
    operator = norm_text(payload.get("operator"))
    reason = norm_text(payload.get("reason"))
    pos = [norm_key(p) for p in payload.get("po_numbers", []) if norm_key(p)]
    target = 1 if payload.get("pulled") else 0
    if not operator:
        return jsonify({"error": "請先選擇操作人員。"}), 400
    if not reason:
        return jsonify({"error": "調整拉單狀態必須填寫原因。"}), 400
    if not pos:
        return jsonify({"error": "沒有選取任何訂單。"}), 400

    conn = get_conn()
    try:
        placeholders = ",".join("?" * len(pos))
        rows = conn.execute(
            f"""SELECT po_number, is_pulled FROM po_headers
                WHERE po_number IN ({placeholders}) AND is_pulled != ?""",
            pos + [target]).fetchall()
        stamp = now()
        for row in rows:
            log_po_change(conn, row["po_number"], "is_pulled", "拉單狀態",
                          "已拉單" if row["is_pulled"] else "未拉單",
                          "已拉單" if target else "未拉單",
                          operator, "manual", reason)
        if target:
            conn.execute(
                f"""UPDATE po_headers SET is_pulled = 1, pulled_at = ?, pulled_by = ?,
                       updated_at = ?, version = version + 1
                    WHERE po_number IN ({placeholders}) AND is_pulled = 0""",
                [stamp, operator, stamp] + pos)
        else:
            conn.execute(
                f"""UPDATE po_headers SET is_pulled = 0, pulled_at = '', pulled_by = '',
                       pulled_batch_id = NULL, updated_at = ?, version = version + 1
                    WHERE po_number IN ({placeholders}) AND is_pulled = 1""",
                [stamp] + pos)
        conn.commit()
        return jsonify({"ok": True, "count": len(rows)})
    finally:
        conn.close()


@app.route("/api/pos/status", methods=["POST"])
def api_batch_status():
    """批次改整批 PO 的狀態，每張單各記一筆歷程（不是只記一筆批次）。"""
    payload = request.get_json(silent=True) or {}
    operator = norm_text(payload.get("operator"))
    pos = [norm_key(p) for p in payload.get("po_numbers", []) if norm_key(p)]
    field = norm_text(payload.get("field"))
    value = norm_text(payload.get("value"))
    if not operator:
        return jsonify({"error": "請先選擇操作人員。"}), 400
    if not pos:
        return jsonify({"error": "沒有選取任何訂單。"}), 400
    if field not in ("po_status", "receiving_status"):
        return jsonify({"error": "只能批次修改 PO 狀態或驗收狀態。"}), 400
    if not value:
        return jsonify({"error": "請選擇要改成什麼狀態。"}), 400

    label = PO_EDITABLE_FIELDS[field]
    conn = get_conn()
    try:
        placeholders = ",".join("?" * len(pos))
        rows = conn.execute(
            f"""SELECT po_number, {field} AS old FROM po_headers
                WHERE po_number IN ({placeholders}) AND {field} != ?""",
            pos + [value]).fetchall()
        for row in rows:
            log_po_change(conn, row["po_number"], field, label,
                          row["old"], value, operator)
        conn.execute(
            f"""UPDATE po_headers SET {field} = ?, updated_at = ?, version = version + 1
                WHERE po_number IN ({placeholders})""",
            [value, now()] + pos)
        conn.commit()
        return jsonify({"ok": True, "count": len(rows)})
    finally:
        conn.close()


# ---------------------------------------------------------------- 匯入

@app.route("/api/import/preview", methods=["POST"])
def api_import_preview():
    """第一階段：只解析與比對，不寫入 orders。OP 看完報告才決定。"""
    operator = norm_text(request.form.get("operator"))
    if not operator:
        return jsonify({"error": "請先在右上角選擇操作人員，才能上傳。"}), 400
    upload = request.files.get("file")
    if upload is None or not upload.filename:
        return jsonify({"error": "沒有收到檔案。"}), 400

    try:
        data = io.BytesIO(upload.read())
        rows, warnings = parse_workbook(data, upload.filename)
    except ImportError_ as exc:
        return jsonify({"error": str(exc)}), 400

    conn = get_conn()
    try:
        result = diff_rows(conn, rows)
        preview = {
            "filename": upload.filename,
            "operator": operator,
            "warnings": warnings,
            "new": result["new"],
            "updated": result["updated"],
            "identical_count": len(result["identical"]),
            "identical_keys": [[r["po_number"], r["sku_id"]]
                               for r in result["identical"]],
        }
        cur = conn.execute(
            """INSERT INTO import_batches
               (filename, operator, rows_total, rows_new, rows_updated,
                rows_identical, rows_error, committed, preview_json, created_at)
               VALUES (?,?,?,?,?,?,?,0,?,?)""",
            (upload.filename, operator, len(rows), len(result["new"]),
             len(result["updated"]), len(result["identical"]), len(warnings),
             json.dumps(preview, ensure_ascii=False), now()),
        )
        conn.commit()
        batch_id = cur.lastrowid
    finally:
        conn.close()

    after_pull = [u for u in result["updated"] if u["after_pull"]]
    return jsonify({
        "batch_id": batch_id,
        "filename": upload.filename,
        "rows_total": len(rows),
        "new_count": len(result["new"]),
        "updated_count": len(result["updated"]),
        "identical_count": len(result["identical"]),
        "after_pull_count": len(after_pull),
        "warnings": warnings,
        "new_preview": result["new"][:200],
        "updated": result["updated"][:200],
    })


@app.route("/api/import/commit", methods=["POST"])
def api_import_commit():
    """第二階段：全有全無寫入。中間任何一列出錯就整批 rollback。"""
    payload = request.get_json(silent=True) or {}
    batch_id = norm_int(payload.get("batch_id"))
    operator = norm_text(payload.get("operator"))
    if not batch_id or not operator:
        return jsonify({"error": "缺少批次或操作人員資訊。"}), 400

    conn = get_conn()
    try:
        batch = conn.execute("SELECT * FROM import_batches WHERE id = ?",
                             (batch_id,)).fetchone()
        if batch is None:
            return jsonify({"error": "找不到這個匯入批次，請重新上傳。"}), 404
        if batch["committed"]:
            return jsonify({"error": "這個批次已經匯入過了，請重新上傳檔案。"}), 400

        preview = json.loads(batch["preview_json"])
    finally:
        conn.close()

    db.backup_db("import")   # 寫入前先留一份身家

    conn = get_conn()
    try:
        conn.execute("BEGIN IMMEDIATE")
        stamp = now()
        inserted = updated = 0

        today = _dt.date.today().isoformat()
        for row in preview["new"]:
            # PO 表頭只在第一次見到這張單時建立。之後同一張單再上傳，
            # 這裡什麼都不做——狀態與建檔日一律以系統為準，不被 Excel 覆蓋。
            conn.execute(
                """INSERT OR IGNORE INTO po_headers
                   (po_number, po_status, receiving_status, is_pulled,
                    filed_date, created_at, updated_at)
                   VALUES (?, '已建立', '未驗收', 0, ?, ?, ?)""",
                (row["po_number"], today, stamp, stamp))

            columns = ", ".join(INSERT_FIELDS)
            marks = ", ".join("?" * len(INSERT_FIELDS))
            values = [row.get(f) for f in INSERT_FIELDS]
            cur = conn.execute(
                f"""INSERT INTO orders ({columns}, qty_ship,
                        source_file, first_seen_at, last_seen_at,
                        created_at, updated_at)
                    VALUES ({marks}, ?, ?, ?, ?, ?, ?)""",
                values + [row.get("qty_coupang"),
                          preview["filename"], stamp, stamp, stamp, stamp],
            )
            order = {"id": cur.lastrowid, "po_number": row["po_number"],
                     "sku_id": row["sku_id"]}
            log_change(conn, order, "_created", "新增訂單", "",
                       f"{row['po_number']} / {row['sku_id']}", operator,
                       "import", preview["filename"])
            inserted += 1

        for item in preview["updated"]:
            row = item["row"]
            order_id = item["order_id"]
            current = conn.execute("SELECT * FROM order_rows WHERE id = ?",
                                   (order_id,)).fetchone()
            if current is None:
                continue
            current = dict(current)

            sets, values = [], []
            critical_hit = False
            unsynced = []
            for change in item["changes"]:
                field = change["field"]
                # OP 已經人工調整過的交期／倉別不覆蓋——那是跟酷澎談好的
                # 結果，酷澎後台只是還沒更新。但差異照樣記歷程、照樣亮燈，
                # 不會把「兩邊還沒同步」這件事藏起來。
                if current.get(f"{field}_overridden"):
                    unsynced.append(f"{change['label']}（酷澎仍為 {change['new']}）")
                    log_change(conn, current, field, change["label"],
                               change["new"], change["old"], operator, "import",
                               "酷澎檔案仍是舊值，保留人工調整結果不覆蓋")
                    critical_hit = critical_hit or field in CRITICAL_FIELDS
                    continue
                sets.append(f"{field} = ?")
                values.append(row.get(field))
                log_change(conn, current, field, change["label"],
                           change["old"], change["new"], operator, "import",
                           preview["filename"])
                if field in CRITICAL_FIELDS:
                    critical_hit = True

            # 出貨數量若 OP 從未手動改過，就跟著酷澎的下單數量走；
            # 一旦手動調整過，匯入永遠不再碰它。
            qty_changed = any(c["field"] == "qty_coupang" for c in item["changes"])
            if qty_changed and not current["qty_ship_overridden"]:
                sets.append("qty_ship = ?")
                values.append(row.get("qty_coupang"))
                log_change(conn, current, "qty_ship", "出貨數量",
                           current["qty_ship"], row.get("qty_coupang"),
                           operator, "import", "隨下單數量連動")

            reason = "、".join(
                f"{c['label']} {c['old']}→{c['new']}" for c in item["changes"]
                if not current.get(f"{c['field']}_overridden")
            )
            if unsynced:
                reason = ("【與酷澎尚未同步】" + "、".join(unsynced)
                          + ("；" + reason if reason else ""))
            if current["is_pulled"] and critical_hit:
                alert = "changed_after_pull"
                reason = "【已拉單後遭變更】" + reason
            else:
                alert = "changed"

            sets += ["needs_review = 1", "alert_level = ?", "review_reason = ?",
                     "last_seen_at = ?", "updated_at = ?", "source_file = ?",
                     "version = version + 1"]
            values += [alert, reason[:900], stamp, stamp, preview["filename"]]

            conn.execute(f"UPDATE orders SET {', '.join(sets)} WHERE id = ?",
                         values + [order_id])
            updated += 1

        # 內容完全相同的列只更新「最後出現時間」，不動任何業務欄位、
        # 不寫歷程、不改 version —— 這就是「靜默去重」。
        conn.executemany(
            "UPDATE orders SET last_seen_at = ? WHERE po_number = ? AND sku_id = ?",
            [(stamp, po, sku) for po, sku in preview.get("identical_keys", [])],
        )

        conn.execute(
            "UPDATE import_batches SET committed = 1, committed_at = ? WHERE id = ?",
            (stamp, batch_id),
        )
        conn.commit()
    except Exception as exc:  # noqa: BLE001
        conn.rollback()
        return jsonify({"error": f"匯入失敗，資料已全部還原，沒有寫入任何一筆：{exc}"}), 500
    finally:
        conn.close()

    return jsonify({
        "ok": True,
        "inserted": inserted,
        "updated": updated,
        "identical": preview["identical_count"],
    })


# ---------------------------------------------------------------- 匯出

def _format_cell(value, fmt):
    if value is None:
        return ""
    if fmt == "int":
        return norm_int(value)
    if fmt == "decimal":
        try:
            return float(value)
        except (TypeError, ValueError):
            return ""
    if fmt == "date":
        return norm_date(value)
    return value


@app.route("/api/export", methods=["POST"])
def api_export():
    """依畫面上的篩選條件匯出。

    是否標記為「已拉單」由 OP 明確決定：試算用的匯出不該把單鎖掉。
    """
    payload = request.get_json(silent=True) or {}
    operator = norm_text(payload.get("operator"))
    profile_key = norm_text(payload.get("profile")) or "warehouse"
    mark_pulled = bool(payload.get("mark_pulled"))
    filters = payload.get("filters") or {}
    selected_pos = [norm_key(x) for x in payload.get("po_numbers", []) if norm_key(x)]

    if not operator:
        return jsonify({"error": "請先選擇操作人員。"}), 400

    profiles = get_profiles()
    profile = profiles.get(profile_key)
    if not profile:
        return jsonify({"error": f"找不到匯出格式「{profile_key}」，請檢查 export_profiles.json。"}), 400

    conn = get_conn()
    try:
        if selected_pos:
            placeholders = ",".join("?" * len(selected_pos))
            rows = conn.execute(
                f"SELECT * FROM order_rows WHERE po_number IN ({placeholders}) "
                f"ORDER BY delivery_date, po_number, seq_no, sku_id",
                selected_pos,
            ).fetchall()
            filter_desc = {"po_numbers": selected_pos}
        else:
            clause, params = build_filter(filters)
            rows = conn.execute(
                f"SELECT * FROM order_rows{clause} "
                f"ORDER BY delivery_date, po_number, seq_no, sku_id", params
            ).fetchall()
            filter_desc = filters

        if not rows:
            return jsonify({"error": "目前的篩選條件沒有任何資料可以匯出。"}), 400

        import openpyxl
        from openpyxl.styles import Alignment, Font, PatternFill

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = profile.get("label", profile_key)[:28]

        columns = profile.get("columns", [])
        headers = [c.get("header", c.get("field", "")) for c in columns]
        ws.append(headers)
        head_fill = PatternFill("solid", fgColor="1F3A5F")
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = head_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")

        for row in rows:
            record = dict(row)
            line = []
            for col in columns:
                if "const" in col:
                    line.append(col["const"])
                    continue
                value = record.get(col.get("field"))
                line.append(_format_cell(value, col.get("format")))
            ws.append(line)

        # 長 PO 單號/料號一定要當文字，否則 Excel 會轉成科學記號
        for idx, col in enumerate(columns, start=1):
            letter = ws.cell(row=1, column=idx).column_letter
            if col.get("format") == "text":
                for cell in ws[letter][1:]:
                    cell.number_format = "@"
            widths = {"text": 18, "date": 12, "int": 10, "decimal": 12}
            ws.column_dimensions[letter].width = widths.get(col.get("format"), 16)
        ws.freeze_panes = "A2"

        stamp = _dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"酷澎出貨表_{profile_key}_{stamp}.xlsx"

        cur = conn.execute(
            """INSERT INTO export_batches
               (operator, profile, filename, row_count, mark_pulled,
                filter_json, created_at)
               VALUES (?,?,?,?,?,?,?)""",
            (operator, profile_key, filename, len(rows), int(mark_pulled),
             json.dumps(filter_desc, ensure_ascii=False), now()),
        )
        batch_id = cur.lastrowid
        conn.executemany(
            """INSERT INTO export_batch_items
               (batch_id, order_id, po_number, sku_id, qty_ship)
               VALUES (?,?,?,?,?)""",
            [(batch_id, r["id"], r["po_number"], r["sku_id"], r["qty_ship"])
             for r in rows],
        )

        if mark_pulled:
            # 拉單是整張 PO 的事：匯出裡只要有這張單的任何一個品項，
            # 整張單就算交出去了。
            stamp_now = now()
            pos = sorted({r["po_number"] for r in rows if not r["is_pulled"]})
            note = f"匯出批次 #{batch_id}／{profile.get('label', profile_key)}"
            for po in pos:
                log_po_change(conn, po, "is_pulled", "拉單狀態", "未拉單",
                              "已拉單", operator, "system", note)
            if pos:
                placeholders = ",".join("?" * len(pos))
                conn.execute(
                    f"""UPDATE po_headers SET is_pulled = 1, pulled_at = ?,
                            pulled_by = ?, pulled_batch_id = ?, updated_at = ?,
                            version = version + 1
                        WHERE po_number IN ({placeholders})""",
                    [stamp_now, operator, batch_id, stamp_now] + pos,
                )
        conn.commit()
    finally:
        conn.close()

    stream = io.BytesIO()
    wb.save(stream)
    stream.seek(0)
    return send_file(
        stream, as_attachment=True, download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/api/export/batches")
def api_export_batches():
    conn = get_conn()
    try:
        rows = conn.execute(
            "SELECT * FROM export_batches ORDER BY id DESC LIMIT 50"
        ).fetchall()
        return jsonify([dict(r) for r in rows])
    finally:
        conn.close()


def _same(a, b):
    if a is None and b is None:
        return True
    if (a is None) != (b is None):
        return str(a or "") == str(b or "")
    return str(a).strip() == str(b).strip()


def lan_ip():
    """抓這台電腦在辦公室網路上的位址，好告訴同事要連哪裡。"""
    import socket
    sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        # 不會真的送出封包，只是讓作業系統挑出對外那張網卡
        sock.connect(("8.8.8.8", 80))
        return sock.getsockname()[0]
    except OSError:
        try:
            return socket.gethostbyname(socket.gethostname())
        except OSError:
            return ""
    finally:
        sock.close()


if __name__ == "__main__":
    import socket as _socket

    # 5000 埠被占用最常見的原因：舊視窗沒關乾淨、背景還留著一個沒更新到
    # 的伺服器在跑。這種狀況畫面通常還打得開（連到舊的那個），怎麼換
    # 檔案都不會生效，卻完全看不出原因——所以啟動前先檢查一次，寧可
    # 明確擋下來，也不要讓兩個版本同時活著。
    probe = _socket.socket(_socket.AF_INET, _socket.SOCK_STREAM)
    already_running = probe.connect_ex(("127.0.0.1", 5000)) == 0
    probe.close()
    if already_running:
        print("=" * 62)
        print(" ⚠ 無法啟動：5000 這個埠已經有別的程式在用了")
        print()
        print(" 最常見的原因：還有一個舊的酷澎訂單系統視窗沒關乾淨，")
        print(" 背景還留著一份沒更新到最新版的伺服器在跑。")
        print()
        print(" 解法：")
        print(" 1. 找找看工作列或工作管理員裡，是不是還有另一個黑視窗")
        print("    （或另一個 python.exe），把它整個關掉")
        print(" 2. 打開工作管理員（Ctrl+Shift+Esc），結束所有 python.exe")
        print(" 3. 再重新雙擊 START 一次")
        print("=" * 62)
        input(" 按 Enter 關閉這個視窗...")
        raise SystemExit(1)

    moved = db.ensure_data_dir()
    init_db()
    ip = lan_ip()

    print("=" * 62)
    print(" 酷澎訂單管理系統 已啟動")
    print(f" 版本號： {BUILD_VERSION}　（畫面右下角也會顯示這個號碼，")
    print("           兩邊對得起來，才代表你看到的是最新版）")
    print()
    print("  你自己用：      http://127.0.0.1:5000")
    if ip and ip != "127.0.0.1":
        print(f"  給同事的網址：  http://{ip}:5000")
        print("                 （同一個辦公室網路才連得到，")
        print("                   而且這台電腦要開著、程式要跑著）")
    else:
        print("  ※ 抓不到辦公室網路位址，同事可能連不進來。")
    print()
    print(f" 資料與設定放在： {db.DATA_DIR}")
    print(" （更新程式時，這個資料夾不要動，裡面是你的訂單和設定）")
    if moved:
        print()
        print(f" ※ 已把 {'、'.join(moved)} 從舊位置搬進資料資料夾")
    print("=" * 62)

    # 綁 0.0.0.0 才收得到區網連線；threaded 讓多人同時操作不用排隊
    app.run(host="0.0.0.0", port=5000, debug=False, threaded=True)

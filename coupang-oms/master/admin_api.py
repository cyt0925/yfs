"""清除資料、匯入歷程（各次匯入新增／有變）、修改歷程查詢。"""
from .common import *  # noqa: F401,F403 — 共用工具、Flask、db、openpyxl 都從這裡來


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
            conn.execute("DELETE FROM mst_mars_splits")      # 瑪氏拆單是從訂單拆出來的，訂單清了一起清
            parts.append(f"訂單明細 {counts['orders']} 筆")
        if clear_products:
            conn.execute("DELETE FROM mst_quotas")
            conn.execute("DELETE FROM mst_month_stats")
            conn.execute("DELETE FROM mst_brand_targets")
            conn.execute("DELETE FROM mst_source_uploads")
            conn.execute("DELETE FROM mst_field_src")
            conn.execute("DELETE FROM mst_mars_products")
            conn.execute("DELETE FROM mst_mars_uploads")
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
    """匯入歷程：每一次「確認匯入」一列，附上「這批目前算新增／有變／消失的有幾筆、落在哪幾個月」。

    新增＝品項第一次出現就是這批；有變＝舊品項最後一次被這批改到（數量、日期…）；
    消失＝這批的檔案裡已經沒有這個品項（列保留、出貨歸 0）。三個數字互斥，
    畫面上排在一起看就對得起來。整批用兩句 GROUP BY 算完，不是一批一批查——
    匯個一百次歷程視窗也不會拖慢。"""
    limit = min(norm_int(request.args.get("limit")) or 60, 300)
    conn = get_conn()
    try:
        batches = _rows(conn.execute(
            """SELECT id, filename, operator, rows_total, rows_new, rows_updated, rows_identical, created_at, committed_at
               FROM mst_import_batches WHERE committed = 1 ORDER BY committed_at DESC, id DESC LIMIT ?""", (limit,)))
        if not batches:
            return jsonify({"batches": []})
        ids = [b["id"] for b in batches]
        marks = ",".join("?" * len(ids))
        new_rows = _rows(conn.execute(
            f"""SELECT first_batch_id AS b, substr(delivery_date, 1, 7) AS m, COUNT(*) AS n, COUNT(DISTINCT po_number) AS pos
                FROM mst_orders WHERE first_batch_id IN ({marks}) GROUP BY first_batch_id, substr(delivery_date, 1, 7)""", ids))
        upd_rows = _rows(conn.execute(
            f"""SELECT last_batch_id AS b, substr(delivery_date, 1, 7) AS m, missing_in_file AS gone,
                       COUNT(*) AS n, COUNT(DISTINCT po_number) AS pos
                FROM mst_orders WHERE last_batch_id IN ({marks}) AND first_batch_id != last_batch_id
                GROUP BY last_batch_id, substr(delivery_date, 1, 7), missing_in_file""", ids))
        # PO 張數（OP 腦子裡的單位是 PO，不是品項）：一批裡「新增」的 PO＝那批第一次出現的 PO；
        # 「有變」「消失」的 PO 照品項的最後一批算。跨月的 PO 只算一次，所以另外用 DISTINCT 算，不從月份加總。
        po_new = {x["b"]: x["n"] for x in _rows(conn.execute(
            f"SELECT first_batch_id AS b, COUNT(DISTINCT po_number) AS n FROM mst_orders WHERE first_batch_id IN ({marks}) GROUP BY first_batch_id", ids))}
        po_upd = {}
        for x in _rows(conn.execute(
                f"""SELECT last_batch_id AS b, missing_in_file AS gone, COUNT(DISTINCT po_number) AS n FROM mst_orders
                    WHERE last_batch_id IN ({marks}) AND first_batch_id != last_batch_id GROUP BY last_batch_id, missing_in_file""", ids)):
            po_upd[(x["b"], 1 if x["gone"] else 0)] = x["n"]
        per = {b["id"]: {} for b in batches}

        def slot(bid, m):
            return per[bid].setdefault(m or "", {"m": m or "", "new": 0, "changed": 0, "removed": 0, "pos": 0})
        for x in new_rows:
            sl = slot(x["b"], x["m"]); sl["new"] += x["n"]; sl["pos"] += x["pos"]
        for x in upd_rows:
            sl = slot(x["b"], x["m"]); sl["removed" if x["gone"] else "changed"] += x["n"]; sl["pos"] += x["pos"]
        for b in batches:
            mc = per[b["id"]]
            b["month_counts"] = [mc[k] for k in sorted(mc)]
            b["new_now"] = sum(x["new"] for x in b["month_counts"])
            b["changed_now"] = sum(x["changed"] for x in b["month_counts"])
            b["removed_now"] = sum(x["removed"] for x in b["month_counts"])
            b["months"] = [x["m"] for x in b["month_counts"] if x["m"]]
            b["new_pos"] = po_new.get(b["id"], 0)
            b["changed_pos"] = po_upd.get((b["id"], 0), 0)
            b["removed_pos"] = po_upd.get((b["id"], 1), 0)
            b["label"] = f"{(b['committed_at'] or '')[5:16]} {b['filename']}：新增 {b['new_pos']} 張 PO（{b['new_now']} 品項）、有變 {b['changed_pos']}、消失 {b['removed_pos']}"
    finally:
        conn.close()
    return jsonify({"batches": batches})


@master_bp.route("/api/master/imports/pos")
def api_imports_pos():
    """每一批動到的 PO 單號（抽屜搜尋用：貼一個單號，找出它在哪幾批出現過）。一句 SQL 撈完。"""
    conn = get_conn()
    try:
        rows = _rows(conn.execute(
            """SELECT first_batch_id AS b, po_number FROM mst_orders WHERE first_batch_id IS NOT NULL
               UNION SELECT last_batch_id AS b, po_number FROM mst_orders WHERE last_batch_id IS NOT NULL"""))
    finally:
        conn.close()
    out = {}
    for r in rows:
        out.setdefault(str(r["b"]), []).append(r["po_number"])
    return jsonify({"pos": out})


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


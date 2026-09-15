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
            changed = _rows(conn.execute(
                """SELECT substr(delivery_date, 1, 7) AS m, COUNT(*) AS n FROM mst_orders
                   WHERE last_batch_id = ? AND first_batch_id != ? GROUP BY substr(delivery_date, 1, 7) ORDER BY m""",
                (b["id"], b["id"])))
            b["changed_now"] = sum(x["n"] for x in changed)
            # 這次匯入的新增／有變各落在哪幾個月（畫面上做成可以點的月份標籤，不再自動跳月份）
            pos = _rows(conn.execute(
                """SELECT substr(delivery_date, 1, 7) AS m, COUNT(DISTINCT po_number) AS n FROM mst_orders
                   WHERE first_batch_id = ? OR last_batch_id = ? GROUP BY substr(delivery_date, 1, 7)""",
                (b["id"], b["id"])))
            mc = {}
            def slot(m):
                return mc.setdefault(m or "", {"m": m or "", "new": 0, "changed": 0, "pos": 0})
            for x in months:
                slot(x["m"])["new"] += x["n"]
            for x in changed:
                slot(x["m"])["changed"] += x["n"]
            for x in pos:
                slot(x["m"])["pos"] += x["n"]
            b["month_counts"] = [mc[k] for k in sorted(mc)]
            b["label"] = f"{(b['committed_at'] or '')[5:16]} {b['filename']}：新增 {b['new_now']}、有變 {b['changed_now']}"
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


"""右側面板：點一張單，把這張單的全部放在一起（Jerry 2026-10-13 改版第二次）。

清單不被蓋住，一邊看清單一邊改（NN/g Data Tables：看一筆、改一筆用不擋畫面的側邊面板）。
這裡把原本散在各頁籤的資料湊成一份：訂單總覽那一列（待辦、通知狀態）、採購單的品項計畫（轉換率、除不盡、EIP）、
竹運出貨的出貨量與效期批次、缺貨紀錄、這張單的 EIP 檔、差異、歷程。

竹運出貨的出貨量要跟清單用同一組到貨日／線別算（庫存是照到貨日由早到晚分的），所以畫面會把篩選條件一起送來。
"""
from .common import *  # noqa: F401,F403
from .orders import BUY_FULFILS, _group_ids, overview
from .purchase import build_plan
from .ship import build_ship
from .short import build_short


@shopee_bp.route("/api/shopee/order")
def api_side():
    no = str(request.args.get("no") or "").strip()
    if not no:
        return jsonify({"error": "缺少訂單編號。"}), 400
    a = request.args
    d1, d2, line = a.get("from", ""), a.get("to", ""), a.get("line", "")
    conn = get_conn()
    try:
        g = next((x for x in overview(conn, {"q": no}) if x["order_no"] == no), None)
        if g is None:
            return jsonify({"error": f"找不到訂單 {no}，可能已被移除，請重新整理頁面。"}), 404
        ids = _group_ids(conn, no)
        fulfils = {i["fulfil"] for i in g["items"]}
        buy = None
        for f in BUY_FULFILS:
            if f in fulfils:
                buy = next(iter(build_plan(conn, f, order_nos=[no])), None) or buy
                if buy:
                    buy["fulfil"] = f
        ship = None
        if "竹運出貨" in fulfils and a.get("ship") != "0":       # ship=0：畫面剛從 /api/shopee/ship/plan 拿到同一組篩選的計畫，不用再算一次（搜尋太慢，2026-10-20）
            groups, ctx = build_ship(conn, d1, d2, line)
            ship = next((x for x in groups if x["order_no"] == no), None)
            if ship is None:                              # 不在目前篩選的範圍裡：只算這張單（不跟別張搶庫存）
                groups, ctx = build_ship(conn, g["expected_date"], g["expected_date"], "")
                ship = next((x for x in groups if x["order_no"] == no), None)
                if ship is not None:
                    ship["alone"] = True
        short = next(iter(build_short(conn, order_nos=[no])), None)
        eips, diffs, logs = [], [], []
        if ids:
            marks = ",".join("?" * len(ids))
            eips = _rows(conn.execute(
                f"SELECT DISTINCT e.id, e.line, e.fulfil, e.unit_group, e.filename, e.eip_po, e.updated_by, e.updated_at, e.created_at "
                f"FROM shp_eip e JOIN shp_eip_items i ON i.eip_id = e.id WHERE i.order_id IN ({marks}) ORDER BY e.id", ids))
            diffs = _rows(conn.execute(
                f"SELECT d.*, o.supplier_sku_id, o.sku_name FROM shp_diffs d JOIN shp_orders o ON o.id = d.order_id "
                f"WHERE d.order_id IN ({marks}) ORDER BY d.id DESC", ids))
        keys = sorted({no, g.get("pr_id") or "", g.get("po_id") or ""} - {""})
        marks = ",".join("?" * len(keys))
        logs = _rows(conn.execute(
            f"SELECT field, field_label, old_value, new_value, operator, note, reason, changed_at FROM mst_logs "
            f"WHERE po_number IN ({marks}) AND field LIKE ? ORDER BY id DESC LIMIT 60", keys + ["shopee_%"]))
    finally:
        conn.close()
    for e in eips:
        e["mars"] = e["line"] == "瑪氏"
    for d in diffs:
        d["confirmed"] = bool(d["confirmed_at"])
    return jsonify({"order": g, "buy": buy, "ship": ship, "short": short, "eips": eips, "diffs": diffs, "logs": logs})


__all__ = ["api_side"]

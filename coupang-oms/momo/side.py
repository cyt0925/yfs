"""右側面板：點一張入庫單，這張單的全部在這裡看、在這裡改（照蝦皮特選，清單不被蓋住）。
品項與下採計畫（箱入數、除不盡、改數量）、這張單的 EIP 檔（回填單號、下載採購單）、③ 缺貨與效期、MOMO 改過的差異、歷程。"""
from .common import *  # noqa: F401,F403
from .check import check_orders
from .orders import _ids, overview
from .purchase import build_plan, eips_of


@momo_bp.route("/api/momo/order")
def api_side():
    no = str(request.args.get("no") or "").strip()
    if not no:
        return jsonify({"error": "缺少入庫單號。"}), 400
    conn = get_conn()
    try:
        g = next((x for x in overview(conn, {"q": no}) if x["inbound_no"] == no), None)
        if g is None:
            return jsonify({"error": f"找不到入庫單 {no}，請重新整理頁面。"}), 404
        plan = next(iter(build_plan(conn, nos=[no])), None)
        ids = _ids(conn, no)
        marks = ",".join("?" * len(ids))
        eip_ids = [r["eip_id"] for r in _rows(conn.execute(f"SELECT DISTINCT eip_id FROM momo_eip_items WHERE order_id IN ({marks})", ids))]
        eips = eips_of(conn, f"id IN ({','.join('?' * len(eip_ids))})", eip_ids) if eip_ids else []
        diffs = _rows(conn.execute(f"SELECT d.*, o.factory_code, o.product_name FROM momo_diffs d JOIN momo_orders o ON o.id = d.order_id "
                                   f"WHERE d.order_id IN ({marks}) ORDER BY d.id DESC", ids))
        logs = _rows(conn.execute("SELECT field, field_label, old_value, new_value, operator, note, reason, changed_at FROM mst_logs "
                                  "WHERE po_number = ? AND field LIKE ? ORDER BY id DESC LIMIT 60", (no, "momo_%")))
        whs = wh_settings(conn)
        check = check_orders(conn, [no]).get(no)
    finally:
        conn.close()
    for d in diffs:
        d["confirmed"] = bool(d["confirmed_at"])
    w = whs.get(wh_canon(g["warehouse"], whs))
    return jsonify({"order": g, "plan": plan, "eips": eips, "diffs": diffs, "logs": logs, "check": check,
                    "warehouse": {"name": wh_canon(g["warehouse"], whs), "note": wh_note(w)} if w else None})


__all__ = ["api_side"]

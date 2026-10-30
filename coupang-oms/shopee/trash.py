"""刪除、還原蝦皮訂單（Alice 2026-10-30：贈品料號那種不要的單要能刪；Jerry：大家都能刪、要填原因、要能還原）。

不動 shp_orders，只在 shp_deleted 記一筆；清單、產 EIP、拋單檔、缺貨通知、勇信核對都用 not_deleted() 跳過。
之後再上傳蝦皮檔，這張單也不會跑回來（PR、PO 都記，PR 後來變成 PO 也對得到）。已產出 EIP 或拋單檔的要先取消產出才能刪。
"""
from .common import *  # noqa: F401,F403
from .orders import _group_ids, overview
from .purchase import po_status_map
from .ship import ship_status_map

DELETE_REASONS = ("贈品不用出", "重複的單", "蝦皮取消")


@shopee_bp.route("/api/shopee/order/delete", methods=["POST"])
def api_order_delete():
    p = request.get_json(silent=True) or {}
    no, reason = str(p.get("order_no") or "").strip(), str(p.get("reason") or "").strip()
    if not no:
        return jsonify({"error": "缺少訂單編號。"}), 400
    if not reason:
        return jsonify({"error": "請填刪除原因。"}), 400
    op = _operator()
    conn = get_conn()
    try:
        g = next((x for x in overview(conn, {"q": no, "deleted": "1"}) if x["order_no"] == no), None)
        if g is None:
            return jsonify({"error": f"找不到訂單 {no}，請重新整理頁面。"}), 404
        if g["trash"]:
            return jsonify({"error": f"{no} 已經刪除了。"}), 400
        ids = _group_ids(conn, no)
        if po_status_map(conn, ids):
            return jsonify({"error": f"{no} 已產出 EIP 採購單，請先取消產出 EIP 採購單再刪除。"}), 400
        if ship_status_map(conn, ids):
            return jsonify({"error": f"{no} 已產出拋單檔，請先取消拋單檔產出再刪除。"}), 400
        conn.execute("INSERT INTO shp_deleted (order_no, pr_id, po_id, reason, deleted_by, deleted_at) VALUES (?,?,?,?,?,?)",
                     (no, g["pr_id"] or "", g["po_id"] or "", reason, op, now()))
        _log(conn, g["line"], no, "", "", "shopee_delete", "刪除訂單", "", f"{g['item_count']} 品、{g['qty']:,} 件", op, "manual", reason=reason)
        conn.commit()
    finally:
        conn.close()
    return jsonify({"ok": True, "message": f"已刪除 {no}。按錯可以在「全部訂單」點「顯示已刪除」還原。"})


@shopee_bp.route("/api/shopee/order/restore", methods=["POST"])
def api_order_restore():
    p = request.get_json(silent=True) or {}
    no = str(p.get("order_no") or "").strip()
    op = _operator()
    conn = get_conn()
    try:
        g = next((x for x in overview(conn, {"q": no, "deleted": "1"}) if x["order_no"] == no), None)
        if g is None or not g["trash"]:
            return jsonify({"error": f"{no} 沒有被刪除，請重新整理頁面。"}), 400
        t = g["trash"]
        conn.execute("DELETE FROM shp_deleted WHERE order_no = ?", (t["order_no"],))
        _log(conn, g["line"], no, "", "", "shopee_restore", "還原訂單", f"{t['deleted_by']} 刪除：{t['reason']}", "還原", op, "manual")
        conn.commit()
    finally:
        conn.close()
    return jsonify({"ok": True, "message": f"已還原 {no}。"})


__all__ = ["api_order_delete", "api_order_restore", "DELETE_REASONS"]

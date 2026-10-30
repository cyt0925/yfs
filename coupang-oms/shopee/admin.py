"""清除資料（Jerry 2026-10-14：「歷程記錄要可以刪除 資料要可以清除」）。

照商品主檔自動化的「清除資料」那套規矩：只有管理員、要打「清空資料」四個字、清之前先備份（PostgreSQL 模式由 Supabase 每日備份頂著）。
兩個範圍讓人勾：訂單資料（預設勾）、參考資料（價格本、庫存資料，預設不勾）；蝦皮特選的歷程預設一起刪，可以選保留。
蝦皮倉資料、拋單檔收件資料、效期標準是設定，不清；在蝦皮特選補登的瑪氏商品是共用的商品總表，也不清。
"""
import os

import db
from master.common import _is_admin

from .common import *  # noqa: F401,F403

# 訂單資料：子表先刪
ORDER_TABLES = ("shp_deleted", "shp_notify", "shp_short", "shp_dispatch_items", "shp_dispatch", "shp_ship", "shp_eip_items", "shp_eip",
                "shp_qty_adj", "shp_diffs", "shp_order_versions", "shp_orders", "shp_uploads")
REF_TABLES = ("shp_conv", "shp_stock")


@shopee_bp.route("/api/shopee/reset", methods=["POST"])
def api_shopee_reset():
    if not _is_admin():
        return jsonify({"error": "只有管理員可以清除資料。"}), 403
    p = request.get_json(silent=True) or {}
    if str(p.get("confirm") or "").strip() != "清空資料":
        return jsonify({"error": "請輸入「清空資料」以確認。"}), 400
    orders, refs, logs = bool(p.get("orders", True)), bool(p.get("refs", False)), bool(p.get("logs", True))
    if not orders and not refs:
        return jsonify({"error": "請至少勾選一項清除範圍。"}), 400
    backup = db.backup_db("shp_reset")
    conn = get_conn()
    try:
        n_orders = conn.execute("SELECT COUNT(DISTINCT CASE WHEN po_id != '' THEN po_id ELSE pr_id END) AS n FROM shp_orders").fetchone()["n"]
        parts = []
        if orders:
            for t in ORDER_TABLES:
                conn.execute(f"DELETE FROM {t}")
            parts.append(f"訂單資料 {n_orders} 張")
        if refs:
            for t in REF_TABLES:
                conn.execute(f"DELETE FROM {t}")
            parts.append("參考資料（價格本、庫存資料）")
        if logs:
            n_logs = conn.execute("DELETE FROM mst_logs WHERE field LIKE ?", ("shopee_%",)).rowcount
        else:
            n_logs = 0
            _log(conn, "", "", "", "", "shopee_reset", "清除資料", "、".join(parts), "已清空", _operator(), "system", "管理員手動清除，歷程保留")
        conn.commit()
    finally:
        conn.close()
    tail = f"，歷程紀錄 {n_logs} 筆一併刪除" if logs else "，歷程紀錄保留"
    note = f"（清除前已自動備份：{os.path.basename(backup)}）" if backup else ""
    return jsonify({"ok": True, "message": f"已清除{'、'.join(parts)}{tail}。{note}"})


__all__ = ["api_shopee_reset"]

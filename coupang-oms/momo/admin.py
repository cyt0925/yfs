"""清除資料（Jerry 2026-10-29「MOMO 訂單管理系統一樣要有清除資料的功能」，照蝦皮特選那套）。

只有管理員、要打「清空資料」四個字、清之前先備份（PostgreSQL 模式由 Supabase 每日備份頂著）。
清的是 MOMO 的訂單資料：下載檔與各版本、差異、改過的數量、EIP 與回填單號、勇信核對、人工判斷、效期檔紀錄、入庫通知單、嘜頭、上傳紀錄。
倉別主檔、業務專用章是設定，不清；價格本、瑪氏商品總表跟蝦皮特選共用，也不清。MOMO 的歷程預設一起刪，可以選保留。
"""
import os

import db
from master.common import _is_admin

from .common import *  # noqa: F401,F403

# 子表先刪
ORDER_TABLES = ("momo_mark", "momo_notice", "momo_expfill", "momo_check", "momo_yx", "momo_eip_items", "momo_eip", "momo_qty_adj",
                "momo_diffs", "momo_order_versions", "momo_orders", "momo_upload_files", "momo_uploads")


@momo_bp.route("/api/momo/reset", methods=["POST"])
def api_momo_reset():
    if not _is_admin():
        return jsonify({"error": "只有管理員可以清除資料。"}), 403
    p = request.get_json(silent=True) or {}
    if str(p.get("confirm") or "").strip() != "清空資料":
        return jsonify({"error": "請輸入「清空資料」以確認。"}), 400
    logs = bool(p.get("logs", True))
    backup = db.backup_db("momo_reset")
    conn = get_conn()
    try:
        n = conn.execute("SELECT COUNT(DISTINCT inbound_no) AS n FROM momo_orders").fetchone()["n"] or 0
        for t in ORDER_TABLES:
            conn.execute(f"DELETE FROM {t}")
        if logs:
            n_logs = conn.execute("DELETE FROM mst_logs WHERE field LIKE ? AND field NOT IN (?, ?)", ("momo_%", "momo_wh", "momo_stamp")).rowcount
        else:
            n_logs = 0
            _log(conn, "", "", "", "", "momo_reset", "清除資料", f"訂單資料 {n} 張", "已清空", _operator(), "system", "管理員手動清除，歷程保留")
        conn.commit()
    finally:
        conn.close()
    tail = f"，歷程紀錄 {n_logs} 筆一併刪除" if logs else "，歷程紀錄保留"
    note = f"（清除前已自動備份：{os.path.basename(backup)}）" if backup else ""
    return jsonify({"ok": True, "message": f"已清除訂單資料 {n} 張入庫單{tail}。倉別主檔、印章、價格本沒動。{note}"})


__all__ = ["api_momo_reset"]

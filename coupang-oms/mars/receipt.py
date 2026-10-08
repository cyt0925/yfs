"""⑥ 驗收單缺貨劃線：把簽好名的酷澎驗收單 PDF 丟進來，照系統出貨數量劃掉缺貨品項，再交給勇信。

以前是人在 PDF 上手改（Jerry 2026-10-08 給的 signed_PO_DELIVERY_INVOICE_13000000558341／569338）：
整筆不出整列畫一條粗線、部分下修把原出貨數量塗黑、旁邊寫新數量，兩聯都改、合計不改。
- 數量照系統裡的出貨數量（mst_orders.qty_ship；⑤ 勇信缺貨按確認後就是實際要出的），Jerry 確認。
- 先只做瑪氏：那張 PO 在 mst_orders 有線別「瑪氏」的列才劃，其他照原檔還回去。
- 一開始做在酷澎訂單管理的「驗收單批次簽名」裡，Jerry 說要放瑪氏出貨頁勇信缺貨旁邊，就搬過來（簽名視窗恢復原樣）。
讀表格、畫線的程式在 pdfsign.py（find_item_rows／shortage_plan／draw_marks／mark_shortage）。
"""
import datetime as _dt
import zipfile

import pdfsign

from .common import *  # noqa: F401,F403

MAX_PDF_BYTES = 25 * 1024 * 1024


def ship_qty_lookup(conn):
    """mark_shortage 用的查表：PO 單號 → {SKU: 系統出貨數量}；不是瑪氏的單回 None（照原檔）。"""
    def lookup(po):
        rows = _rows(conn.execute("SELECT sku_id, qty_ship, line FROM mst_orders WHERE po_number = ?", (po,)))
        if not any((r["line"] or "") == LINE for r in rows):
            return None
        return {str(r["sku_id"]): r["qty_ship"] for r in rows if r["qty_ship"] is not None}
    return lookup


def note_of(info):
    items = (info or {}).get("items") or []
    strike = sum(1 for i in items if i["kind"] == "strike")
    return "、".join(x for x in (f"整列劃掉 {strike} 品" if strike else "", f"改數量 {len(items) - strike} 品" if len(items) > strike else "") if x)


def _process(files, conn):
    """每份：(檔名, 新的 bytes, 結果)。壞檔、不是 PDF 的照原檔、結果寫原因。"""
    lookup = ship_qty_lookup(conn)
    out = []
    for f in files:
        raw = f.read()
        info = {"po_number": "", "applies": False, "items": [], "missing": [], "rows": 0, "error": ""}
        if not f.filename.lower().endswith(".pdf"):
            info["error"] = "不是 PDF 檔"
        elif len(raw) > MAX_PDF_BYTES:
            info["error"] = "檔案太大（超過 25 MB）"
        else:
            try:
                raw, got = pdfsign.mark_shortage(raw, lookup)
                info.update(got)
            except pdfsign.SignError as exc:
                info["error"] = str(exc)
        info["note"] = note_of(info)
        out.append((f.filename, raw, info))
    return out


@mars_bp.route("/api/mars/receipts/preview", methods=["POST"])
def api_receipts_preview():
    """先看：哪幾份是瑪氏、哪幾品會整列劃掉或改數量、哪些 SKU 系統裡找不到。不存任何東西。"""
    files = [f for f in request.files.getlist("file") if f and f.filename]
    if not files:
        return jsonify({"error": "沒有收到檔案。"}), 400
    conn = get_conn()
    try:
        done = _process(files, conn)
    finally:
        conn.close()
    return jsonify({"files": [{"filename": n, **i} for n, _, i in done]})


@mars_bp.route("/api/mars/receipts/mark", methods=["POST"])
def api_receipts_mark():
    """劃好線下載：一份直接回 PDF，多份打包 zip；檔名照原檔。有劃到的記一筆歷程。"""
    files = [f for f in request.files.getlist("file") if f and f.filename]
    if not files:
        return jsonify({"error": "沒有收到檔案。"}), 400
    operator = _operator()
    conn = get_conn()
    try:
        done = _process(files, conn)
        for name, _, info in done:
            if info["note"]:
                _log(conn, LINE, info["po_number"], "", "", "mars_receipt", "驗收單缺貨劃線", "", info["note"], operator, "manual", name)
        conn.commit()
    finally:
        conn.close()
    marked = sum(1 for _, _, i in done if i["note"])
    if len(done) == 1:
        name, data, _ = done[0]
        resp = send_file(io.BytesIO(data), as_attachment=True, download_name=name, mimetype="application/pdf")
    else:
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for name, data, _ in done:
                zf.writestr(name, data)
        buf.seek(0)
        resp = send_file(buf, as_attachment=True, download_name=f"驗收單缺貨劃線_{_dt.date.today():%Y%m%d}.zip", mimetype="application/zip")
    resp.headers["X-Mars-Receipt-Files"] = str(len(done))
    resp.headers["X-Mars-Receipt-Marked"] = str(marked)
    return resp


__all__ = ["ship_qty_lookup", "note_of", "api_receipts_preview", "api_receipts_mark"]

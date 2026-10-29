"""資料與設定：倉別主檔（需求第九章）、上傳入口認檔、上傳區旁邊的資料更新時間。"""
from mars.products import parse_master
from shopee.purchase import conv_status, parse_conv

from .common import *  # noqa: F401,F403
from .parse import looks_like

WH_LABELS = {"address": "地址", "hours": "商品入庫時間", "phone": "駐倉廠服電話", "contact": "聯絡人", "ship_to": "Ship-to", "mars_name": "瑪氏採購單倉名"}


@momo_bp.route("/api/momo/warehouses")
def api_warehouses():
    conn = get_conn()
    try:
        whs = wh_settings(conn)
    finally:
        conn.close()
    return jsonify({"warehouses": [{"name": k, **v, "note": wh_note(v)} for k, v in whs.items()], "labels": WH_LABELS})


@momo_bp.route("/api/momo/warehouses/<path:name>", methods=["PUT"])
def api_warehouse_save(name):
    if name not in WH_DEFAULTS:
        return jsonify({"error": f"倉別主檔沒有「{name}」。"}), 404
    payload = request.get_json(silent=True) or {}
    operator = _operator()
    conn = get_conn()
    try:
        row = conn.execute("SELECT value FROM mst_meta WHERE key = ?", (WH_KEY,)).fetchone()
        saved = json.loads(row["value"] or "{}") if row else {}
        before = wh_settings(conn)[name]
        cur, changed = dict(saved.get(name) or {}), []
        for k in WH_FIELDS:
            if k in payload:
                v = str(payload.get(k) or "").strip()
                if v != before[k]:
                    changed.append((k, before[k], v))
                cur[k] = v
        if not changed:
            return jsonify({"ok": True, "changed": 0})
        saved[name] = cur
        value = json.dumps(saved, ensure_ascii=False)
        if row:
            conn.execute("UPDATE mst_meta SET value = ? WHERE key = ?", (value, WH_KEY))
        else:
            conn.execute("INSERT INTO mst_meta (key, value) VALUES (?, ?)", (WH_KEY, value))
        for k, old, new in changed:
            _log(conn, "", "", "", "", "momo_wh", f"MOMO 倉別 {name} {WH_LABELS[k]}", old, new, operator, "manual")
        conn.commit()
        return jsonify({"ok": True, "changed": len(changed)})
    finally:
        conn.close()


def detect(data, filename):
    """orders（MOMO 下載檔）／mars（瑪氏商品總表）／conv（價格本）／yx（PDF，勇信配送明細表）／notice（PDF，MOMO 入庫通知單）；認不出來回空字串。蝦皮的檔不收，免得傳錯頁。"""
    if data[:5] == b"%PDF-":
        from .notice import is_notice
        return "notice" if is_notice(data) else "yx"     # MOMO 入庫通知單，不然當勇信配送明細表（讀不到收貨單號，/api/momo/yx 會講）
    if looks_like(data, filename):
        return "orders"
    try:
        if parse_master(io.BytesIO(data))[0]:
            return "mars"
    except Exception:  # noqa: BLE001
        pass
    try:
        if parse_conv(data, filename)[0]:
            return "conv"
    except Exception:  # noqa: BLE001
        pass
    return ""


@momo_bp.route("/api/momo/kinds", methods=["POST"])
def api_kinds():
    """上傳入口：先認每個檔是什麼。MOMO 下載檔畫面再一起送去預覽；價格本、瑪氏商品總表送去原本那支（跟蝦皮特選、瑪氏出貨共用）。"""
    from .check import is_expfill
    out = []
    conn = get_conn()
    try:
        for f in request.files.getlist("file"):
            if f and f.filename:
                data = f.read()
                kind = detect(data, f.filename)
                if kind == "orders" and is_expfill(conn, data, f.filename):
                    kind = "expfill"                      # MOMO 下載檔，裡面的單都已經核過勇信表：多半是效期匯入檔，畫面讓人選
                out.append({"name": f.filename, "kind": kind})
    finally:
        conn.close()
    return jsonify({"files": out})


@momo_bp.route("/api/momo/fresh")
def api_fresh():
    """上傳區右邊：MOMO 下載檔最後一次匯入、價格本、瑪氏商品總表什麼時候更新的（帶到幾點幾分）。"""
    conn = get_conn()
    try:
        last = _row(conn.execute("SELECT filenames, committed_at FROM momo_uploads WHERE status = 'ok' ORDER BY id DESC LIMIT 1"))
        st = conv_status(conn)
    finally:
        conn.close()
    return jsonify({"orders": last, "conv": st})


__all__ = ["detect", "api_kinds", "api_fresh", "api_warehouses", "api_warehouse_save"]

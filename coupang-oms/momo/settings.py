"""資料與設定：倉別主檔（需求第九章）、上傳入口認檔、上傳區旁邊的資料更新時間。"""
from mars.products import parse_master
from shopee.purchase import conv_status, parse_conv

from .common import *  # noqa: F401,F403
from .parse import looks_like

WH_LABELS = {"address": "地址", "hours": "商品入庫時間", "phone": "駐倉廠服電話", "contact": "聯絡人", "ship_to": "Ship-to", "mars_name": "瑪氏採購單倉名",
             "aliases": "也認（別名）"}


def _missing(conn, whs):
    """MOMO 檔出現過、主檔對不到的倉別：[{name, orders}]。新增只能從這裡挑，倉名才會跟 MOMO 檔一字不差（Jerry 2026-10-29）。"""
    rows = _rows(conn.execute("SELECT warehouse, COUNT(DISTINCT inbound_no) AS n FROM momo_orders WHERE warehouse != '' GROUP BY warehouse"))
    return sorted(({"name": r["warehouse"], "orders": int(r["n"])} for r in rows if not wh_canon(r["warehouse"], whs)), key=lambda x: (-x["orders"], x["name"]))


def _split_aliases(v):
    return [a for a in dict.fromkeys(x.strip() for x in re.split(r"[、,，;；\n]+", str(v or ""))) if a]


@momo_bp.route("/api/momo/warehouses")
def api_warehouses():
    conn = get_conn()
    try:
        whs = wh_settings(conn)
        missing = _missing(conn, whs)
    finally:
        conn.close()
    return jsonify({"warehouses": [{"name": k, **v, "note": wh_note(v)} for k, v in whs.items()], "labels": WH_LABELS, "missing": missing})


@momo_bp.route("/api/momo/warehouses", methods=["POST"])
def api_warehouse_add():
    """主檔沒有的倉：加成新的倉（name），或當成某個倉的別名（alias_of）。只收 MOMO 檔真的出現過的倉名。"""
    p = request.get_json(silent=True) or {}
    name, alias_of = str(p.get("name") or "").strip(), str(p.get("alias_of") or "").strip()
    operator = _operator()
    conn = get_conn()
    try:
        whs = wh_settings(conn)
        if name not in {m["name"] for m in _missing(conn, whs)}:
            return jsonify({"error": f"「{name}」不是 MOMO 檔裡出現過、主檔還沒有的倉，請重新整理頁面。"}), 400
        saved, existed = _saved_wh(conn)
        if alias_of:
            if alias_of not in whs:
                return jsonify({"error": f"倉別主檔沒有「{alias_of}」。"}), 404
            cur = dict(saved.get(alias_of) or {})
            cur["aliases"] = whs[alias_of]["aliases"] + [name]
            saved[alias_of] = cur
            _log(conn, "", "", "", "", "momo_wh", f"MOMO 倉別 {alias_of} 也認（別名）", "、".join(whs[alias_of]["aliases"]), "、".join(cur["aliases"]), operator, "manual")
        else:
            saved[name] = {"custom": True}
            _log(conn, "", "", "", "", "momo_wh", "MOMO 倉別主檔", "", f"新增 {name}", operator, "manual")
        save_wh(conn, saved, existed)
        conn.commit()
    finally:
        conn.close()
    return jsonify({"ok": True})


@momo_bp.route("/api/momo/warehouses/<path:name>", methods=["PUT"])
def api_warehouse_save(name):
    payload = request.get_json(silent=True) or {}
    operator = _operator()
    conn = get_conn()
    try:
        whs = wh_settings(conn)
        if name not in whs:
            return jsonify({"error": f"倉別主檔沒有「{name}」。"}), 404
        saved, existed = _saved_wh(conn)
        before = whs[name]
        cur, changed = dict(saved.get(name) or {}), []
        for k in WH_FIELDS:
            if k in payload:
                v = str(payload.get(k) or "").strip()
                if v != before[k]:
                    changed.append((k, before[k], v))
                cur[k] = v
        if "aliases" in payload:
            al = _split_aliases(payload.get("aliases"))
            taken = {a: k for k, w in whs.items() if k != name for a in [k] + w["aliases"]}
            clash = [f"「{a}」已經是 {taken[a]}" for a in al if a in taken]
            if clash or name in al:
                return jsonify({"error": "別名不能重複：" + "、".join(clash or [f"「{name}」就是這個倉的名字"])}), 400
            if al != before["aliases"]:
                changed.append(("aliases", "、".join(before["aliases"]), "、".join(al)))
                cur["aliases"] = al
        if not changed:
            return jsonify({"ok": True, "changed": 0})
        saved[name] = cur
        save_wh(conn, saved, existed)
        for k, old, new in changed:
            _log(conn, "", "", "", "", "momo_wh", f"MOMO 倉別 {name} {WH_LABELS[k]}", old, new, operator, "manual")
        conn.commit()
        return jsonify({"ok": True, "changed": len(changed)})
    finally:
        conn.close()


def detect(data, filename):
    """orders（MOMO 下載檔）／mars（瑪氏商品總表）／conv（價格本）／yx（PDF，勇信配送明細表）／notice（PDF，MOMO 入庫通知單）；認不出來回空字串。蝦皮的檔不收，免得傳錯頁。"""
    if data[:5] == b"%PDF-":
        from .notice import is_mark, is_notice
        return "notice" if is_notice(data) or is_mark(data) else "yx"   # MOMO 入庫通知單、嘜頭（同一支收），不然當勇信配送明細表
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


__all__ = ["detect", "api_kinds", "api_fresh", "api_warehouses", "api_warehouse_add", "api_warehouse_save"]

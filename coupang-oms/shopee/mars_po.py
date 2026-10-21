"""蝦皮特選瑪氏直送：瑪氏制式採購單（需求第十一章）＋蝦皮倉資料。

EIP 採購單在 purchase.py 產（一份＝一個倉、一個到貨日、一個品類、一個單位、中標一樣，需求 11.3）；
回填 EIP 採購單號後，這裡照那份存下來的內容（shp_eip.rows_json）產瑪氏採購單，一張 EIP 對一份瑪氏採購單（Jerry 2026-10-12「照建議」）。
套酷澎瑪氏同一份 V2 範本、同一支 mars.po.po_file，只換蝦皮的值：
  C3 下單日＝到貨日往前 2 個工作天（跳過週末、國定假日、瑪氏出貨頁「採購單設定」裡加的假日）    F3 配送日＝到貨日
  C5 EIP 採購單號      C6／C7／C8／F7／F8 蝦皮倉資料（下面 WH_DEFAULTS，可在畫面改）      B9 品類全名
  F6 固定「1/2效期以上」（需求 11.1；酷澎那邊巧克力、口香糖是 3/5，蝦皮照需求書）
  F5 作業備註一行一條：蝦皮特選嘜頭(入庫通知單)+採購單／需貼中盒標（中標 V）／貼板嘜（箱單位）／指定品需加工貼小白標（有採購單箱備註）
  第 12 列起：A 永豐料號（蝦皮 Supplier SKU）、D 單位需求＝蝦皮數量、E 採購數量(箱)＝蝦皮數量÷箱入數，其他照商品總表
檔名：PO202610101永豐Mars採購單_1013_觀音_盒_PET_需貼中標.xlsx（前面 EIP 採購單號，加到貨日和倉別；Jerry 2026-10-12）。
"""
import zipfile

from mars.po import load_settings as mars_settings
from mars.po import po_file

from .common import *  # noqa: F401,F403

MARS_LINE = "瑪氏"
WH_KEY = "shopee_wh"
WH_FIELDS = ("name", "address", "phone", "ship_to", "contact")
# 觀音、安南三照需求書的倉別表；其他倉的地址、電話照蝦皮供應商進倉規範 V20260901 第 6 頁，ship-to 不知道先空著
WH_DEFAULTS = {
    "TWA": {"name": "永豐商店蝦皮特選-觀音倉(TWA)", "address": "桃園市觀音區寶倉街108號 (C棟1樓 3, 4, 5號碼頭)", "phone": "03-2600500#5924",
            "ship_to": "17617040", "contact": "進貨組"},
    "TWX": {"name": "永豐商店蝦皮特選-安南三倉(TWX)", "address": "臺南市安定區安新二路688號 (01～12碼頭)", "phone": "0983-655-699",
            "ship_to": "17647732", "contact": "曾小姐"},
    "TWT": {"name": "永豐商店蝦皮特選-安南一倉(TWT)", "address": "台南市安南區新吉一路118號", "phone": "06-600-1450#5871", "ship_to": "", "contact": "安南一進貨組"},
    "TWG": {"name": "永豐商店蝦皮特選-高鐵南倉(TWG)", "address": "桃園市中壢區高鐵南路四段888號", "phone": "0966-233-806", "ship_to": "", "contact": "高鐵南進貨組"},
    "TWH": {"name": "永豐商店蝦皮特選-楊梅倉(TWH)", "address": "桃園市楊梅區和平路576號 (物流共和國-楊梅-2A 08-13碼頭)", "phone": "03-262-6560#3151",
            "ship_to": "", "contact": ""},
    "TWW": {"name": "永豐商店蝦皮特選-威獅倉(TWW)", "address": "桃園市楊梅區梅獅路二段666號", "phone": "0975-213-830", "ship_to": "", "contact": "威獅進貨組"},
    "TWK": {"name": "永豐商店蝦皮特選-威獅倉(TWK)", "address": "桃園市楊梅區梅獅路二段666號 (5F)", "phone": "0975-213-830", "ship_to": "", "contact": "威獅進貨組"},
}
F5_FIXED = "蝦皮特選嘜頭(入庫通知單)+採購單"
SHELF_REQ = "1/2效期以上"
LEAD_DAYS = 2


# ── 蝦皮倉資料 ───────────────────────────────────────────────────────────────
def wh_settings(conn):
    """{倉代碼: {name, address, phone, ship_to, contact, edited}}；存過的蓋預設。"""
    row = conn.execute("SELECT value FROM mst_meta WHERE key = ?", (WH_KEY,)).fetchone()
    saved = {}
    if row:
        try:
            saved = json.loads(row["value"] or "{}")
        except ValueError:
            saved = {}
    out = {}
    for code, d in WH_DEFAULTS.items():
        v = dict(d)
        v.update({k: str(x) for k, x in (saved.get(code) or {}).items() if k in WH_FIELDS})
        v["edited"] = bool(saved.get(code))
        out[code] = v
    return out


@shopee_bp.route("/api/shopee/warehouses")
def api_shopee_warehouses():
    conn = get_conn()
    try:
        whs = wh_settings(conn)
    finally:
        conn.close()
    return jsonify({"warehouses": [{"code": c, "short": wh_name(c), **v} for c, v in whs.items()]})


@shopee_bp.route("/api/shopee/warehouses/<code>", methods=["PUT"])
def api_shopee_warehouse_save(code):
    code = (code or "").strip().upper()
    if code not in WH_DEFAULTS:
        return jsonify({"error": f"沒有 {code} 這個蝦皮倉。"}), 404
    payload = request.get_json(silent=True) or {}
    operator = _operator()
    conn = get_conn()
    try:
        row = conn.execute("SELECT value FROM mst_meta WHERE key = ?", (WH_KEY,)).fetchone()
        saved = json.loads(row["value"] or "{}") if row else {}
        before = wh_settings(conn)[code]
        cur = dict(saved.get(code) or {})
        changed = []
        for k in WH_FIELDS:
            if k in payload:
                v = str(payload.get(k) or "").strip()
                if v != before[k]:
                    changed.append((k, before[k], v))
                cur[k] = v
        if not changed:
            return jsonify({"ok": True, "changed": 0})
        saved[code] = cur
        value = json.dumps(saved, ensure_ascii=False)
        if row:
            conn.execute("UPDATE mst_meta SET value = ? WHERE key = ?", (value, WH_KEY))
        else:
            conn.execute("INSERT INTO mst_meta (key, value) VALUES (?, ?)", (WH_KEY, value))
        labels = {"name": "入倉倉別", "address": "地址", "phone": "電話", "ship_to": "Ship-to", "contact": "送貨聯絡人"}
        for k, old, new in changed:
            _log(conn, MARS_LINE, "", "", "", "shopee_wh", f"蝦皮倉資料 {code} {labels[k]}", old, new, operator, "manual", wh_name(code))
        conn.commit()
        return jsonify({"ok": True, "changed": len(changed)})
    finally:
        conn.close()


# ── 瑪氏採購單 ───────────────────────────────────────────────────────────────
def f5_text(unit, label, rows):
    lines = [F5_FIXED]
    if label == "V":
        lines.append("需貼中盒標")
    if unit == "箱":
        lines.append("貼板嘜")
    if any((r.get("po_case_note") or "").strip() for r in rows):
        lines.append("指定品需加工貼小白標")
    return "\n".join(lines)


def po_name(e, rows):
    r0 = rows[0] if rows else {}
    lab = "需貼中標" if r0.get("label") == "V" else "不貼中標"
    d = e["expected_date"] or ""
    return f"{e['eip_po'] or ''}永豐Mars採購單_{d[5:7]}{d[8:10]}_{wh_name(e['warehouse'])}_{r0.get('unit', '')}_{r0.get('category') or '品類不明'}_{lab}.xlsx"


def po_bytes(conn, e, whs=None, base=None):
    """一張瑪氏 EIP（shp_eip 那列）→ 瑪氏採購單 xlsx。"""
    rows = json.loads(e["rows_json"] or "[]")
    r0 = rows[0] if rows else {}
    unit, cat, lab = r0.get("unit", ""), r0.get("category", ""), r0.get("label", "")
    whs = whs or wh_settings(conn)
    w = whs.get(e["warehouse"]) or {"name": wh_name(e["warehouse"]), "address": "", "phone": "", "ship_to": "", "contact": ""}
    settings = dict(base or mars_settings(conn))
    settings.update(lead_days=LEAD_DAYS, shelf_req=SHELF_REQ, shelf_req_cat={}, label_line="", white_line="",
                    special_text=f5_text(unit, lab, rows))
    s = {"delivery_date": e["expected_date"], "eip_po": e["eip_po"], "category": cat, "label": lab, "unit": unit, "slot_time": ""}
    items = [{"yf_sku": r["code"], "unit": r["unit"], "mars_code": r.get("mars_code", ""), "mars_name": r.get("mars_name", ""),
              "qty_ship": r["qty"], "cases": r.get("cases", 0), "price": r.get("price"), "pcs_per_case": r.get("pcs_per_case"),
              "inner_per_case": r.get("inner_per_case"), "note": r.get("note", ""), "po_case_note": r.get("po_case_note", "")} for r in rows]
    wh = {"name": w["name"], "address": w["address"], "phone": w["phone"], "ship_to": w["ship_to"], "contact": w["contact"], "special_note": ""}
    return po_name(e, rows), po_file(s, items, wh, settings)


@shopee_bp.route("/api/shopee/eips/<int:eid>/mars_po")
def api_shopee_mars_po(eid):
    conn = get_conn()
    try:
        e = _row(conn.execute("SELECT * FROM shp_eip WHERE id = ?", (eid,)))
        if e is None or e["line"] != MARS_LINE:
            return jsonify({"error": "找不到這張瑪氏 EIP 採購單，請重新整理頁面。"}), 404
        if not e["eip_po"]:
            return jsonify({"error": "請先回填 EIP 採購單號，才能產出瑪氏採購單。"}), 400
        name, data = po_bytes(conn, e)
    finally:
        conn.close()
    return send_file(io.BytesIO(data), as_attachment=True, download_name=name,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@shopee_bp.route("/api/shopee/mars_po.zip")
def api_shopee_mars_po_zip():
    """到貨日區間內、已回填 EIP 採購單號的瑪氏採購單一次下載。"""
    where, params = ["line = ?", "eip_po != ''"], [MARS_LINE]
    if request.args.get("from"):
        where.append("expected_date >= ?"); params.append(request.args["from"])
    if request.args.get("to"):
        where.append("expected_date <= ?"); params.append(request.args["to"])
    usql, up = upload_cond(request.args.get("up_from", ""), request.args.get("up_to", ""), "o.")
    if usql:
        where.append(f"id IN (SELECT i.eip_id FROM shp_eip_items i JOIN shp_orders o ON o.id = i.order_id WHERE {usql})"); params += up
    conn = get_conn()
    try:
        eips = _rows(conn.execute(f"SELECT * FROM shp_eip WHERE {' AND '.join(where)} ORDER BY expected_date, warehouse, id", params))
        if not eips:
            return jsonify({"error": "這段到貨日沒有已回填 EIP 採購單號的瑪氏採購單。"}), 400
        whs, base = wh_settings(conn), mars_settings(conn)
        files = [po_bytes(conn, e, whs, base) for e in eips]
    finally:
        conn.close()
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in files:
            zf.writestr(name, data)
    buf.seek(0)
    ds = sorted({e["expected_date"] for e in eips})
    md = lambda d: f"{d[5:7]}{d[8:10]}"  # noqa: E731
    zname = f"蝦皮特選瑪氏採購單_{md(ds[0])}" + (f"-{md(ds[-1])}" if ds[-1] != ds[0] else "") + ".zip"
    resp = send_file(buf, as_attachment=True, download_name=zname, mimetype="application/zip")
    resp.headers["X-Shopee-Files"] = str(len(files))
    return resp


__all__ = ["wh_settings", "f5_text", "po_name", "po_bytes", "api_shopee_warehouses", "api_shopee_warehouse_save",
           "api_shopee_mars_po", "api_shopee_mars_po_zip"]

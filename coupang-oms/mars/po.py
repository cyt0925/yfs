"""③ 瑪氏採購單：把已回填 EIP 採購單號的拆單表，套進瑪氏的訂貨通知單範本（永豐Mars採購單(範本)_V2）。

範本存在 purchase_templates/mars_po_v2.xlsx，格子位置照 Alice 的規則（她的規則寫的是蝦皮版，位置跟 V2 一樣）：
  C3 下單日＝配送日往前推 N 個工作天（預設 2，扣週末＋設定裡的假日）    F3 配送日＝到貨日
  C5 永豐PO單號＝EIP 採購單號                                          F5 特殊需求（見 special_text）
  C6 入倉倉別、C7 地址、C8 電話、F7 ship-to、F8 聯絡人 ＝ 採購單設定裡那個倉的資料   F6 效期要求
  B9 品類全名（Chocolate巧克力／Gum糖果／Petcare寵物）
  第 12 列起一列一個下採料號：A 永豐料號  B 瑪氏貨號  C 品名  D 單位需求(for嘜頭)  E 採購數量(箱)  F 系統價格
  G 小計（範本公式）  H 中盒需貼標 V  I 指定效期（空）  J 每箱產品數  K 每箱中盒數  L 備註  M 需加工貼小白單
  E10／F10 的加總公式範本只到第 28 列，這裡改成蓋到最後一列；超過 52 列就往下補格式。

幾個決定（細節在 docs/瑪氏出貨_設計筆記.md）：
- D「單位需求(for嘜頭)」直接填出貨數量（酷澎下單的單位數）。Alice 的公式（盒＝每箱中盒數×箱數）在整合表箱入數跟
  商品總表不一樣時會算錯（M10403852：36 vs 酷澎要的 432）；嘜頭要寫的就是酷澎下單的數量。
- 同一個下採料號（同單位）合成一列，箱數、出貨數量加總；組出商品在備註寫「訂單料號 M…」。
- 一定要先填 EIP 採購單號（C5）跟約倉時間（特殊需求那行）才產得出來，缺哪個講清楚。
- 各倉的地址、電話、ship-to、聯絡人放 mst_mars_warehouses，畫面「採購單設定」裡填；地址沒填就用該倉訂單上的地址。
- 檔名：永豐Mars採購單_單位_品類_中標_倉_酷澎PO.xlsx（Alice 的規則；單位那段照拆單的單位，才不會盒跟包撞名）。
"""
import copy
import datetime as _dt
import os
import zipfile

from .common import *  # noqa: F401,F403

TEMPLATE = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "purchase_templates", "mars_po_v2.xlsx")
SETTINGS_KEY = "mars_po_settings"
FIRST_ROW, LAST_TEMPLATE_ROW = 12, 52     # 範本第 12～52 列有格式與 G 欄公式
TIME_RE = re.compile(r"\d{1,2}:\d{2}")

DEFAULT_SETTINGS = {
    "lead_days": 2,                       # 下單日＝配送日往前幾個工作天
    "shelf_req": "1/2效期以上",           # F6
    "contact_default": "酷澎",            # F8 沒填時
    "label_line": "需貼中盒標",           # 這份要貼中標時，特殊需求第一行
    "white_line": "指定品需加工貼小白標",  # 有品項有採購單箱備註時，特殊需求最後一行
    # 特殊需求固定文字，一行一條；{約倉時間} 換成那份填的約倉時間，{約倉開始} 換成約倉時間裡第一個 HH:MM
    "special_text": "箱嘜，需當面對點數量\n酷澎嘜頭+驗收單\n進倉時間{約倉時間}\n*請在{約倉開始}前抵達，以免被算遲到，謝謝",
    "holidays": [],                       # 扣掉的假日，YYYY-MM-DD
}
TEXT_KEYS = ("shelf_req", "contact_default", "label_line", "white_line", "special_text")


# ── 設定 ─────────────────────────────────────────────────────────────────────
def load_settings(conn):
    row = conn.execute("SELECT value FROM mst_meta WHERE key = ?", (SETTINGS_KEY,)).fetchone()
    s = dict(DEFAULT_SETTINGS)
    if row:
        try:
            s.update({k: v for k, v in json.loads(row["value"] or "{}").items() if k in DEFAULT_SETTINGS})
        except ValueError:
            pass
    s["holidays"] = sorted(set(s.get("holidays") or []))
    return s


def parse_holiday(text):
    """接受 2026-09-25、2026/9/25、9/25（今年）。回 ISO 字串，看不懂回 None。"""
    t = norm_text(text).replace("／", "/").replace("－", "-").replace("．", ".").replace(".", "/")
    if not t:
        return None
    m = re.fullmatch(r"(\d{4})[-/](\d{1,2})[-/](\d{1,2})", t)
    if m:
        y, mo, d = map(int, m.groups())
    else:
        m = re.fullmatch(r"(\d{1,2})[-/](\d{1,2})", t)
        if not m:
            return None
        y, mo, d = _dt.date.today().year, int(m.group(1)), int(m.group(2))
    try:
        return _dt.date(y, mo, d).isoformat()
    except ValueError:
        return None


def save_settings(conn, payload, operator):
    cur = load_settings(conn)
    new = dict(cur)
    if "lead_days" in payload:
        n = norm_int(payload.get("lead_days"))
        if n is None or n < 0 or n > 10:
            return None, "下單日提前幾個工作天要填 0～10 的整數。"
        new["lead_days"] = n
    for k in TEXT_KEYS:
        if k in payload:
            new[k] = str(payload.get(k) or "").replace("\r", "").strip()[:600]
    if "holidays" in payload:
        raw = payload.get("holidays")
        parts = raw if isinstance(raw, list) else re.split(r"[\n,，、;；\s]+", str(raw or ""))
        days, bad = [], []
        for p in parts:
            if not norm_text(p):
                continue
            d = parse_holiday(p)
            (days if d else bad).append(d or p)
        if bad:
            return None, f"這幾個假日看不懂日期：{'、'.join(bad[:5])}。請用 2026-09-25 或 9/25 這種寫法。"
        new["holidays"] = sorted(set(days))
    if new != cur:
        conn.execute("DELETE FROM mst_meta WHERE key = ?", (SETTINGS_KEY,))
        conn.execute("INSERT INTO mst_meta (key, value) VALUES (?, ?)", (SETTINGS_KEY, json.dumps(new, ensure_ascii=False)))
        changed = [k for k in new if new[k] != cur[k]]
        _log(conn, LINE, "", "", "", "mars_po_settings", "瑪氏採購單設定", "", "、".join(changed), operator, "manual")
    return new, ""


def order_date(delivery, lead_days, holidays):
    """配送日往前推 lead_days 個工作天（跳過週六日與假日）。"""
    d = delivery
    left = lead_days
    while left > 0:
        d -= _dt.timedelta(days=1)
        if d.weekday() < 5 and d.isoformat() not in holidays:
            left -= 1
    return d


# ── 倉庫資料 ─────────────────────────────────────────────────────────────────
def _codes_in_orders(conn):
    """瑪氏訂單出現過的倉，以及各倉最常見的地址（整合表有地址欄）。"""
    cfg = _line_groups()
    seen = {}
    for o in _rows(conn.execute("SELECT line, warehouse, address FROM mst_orders WHERE warehouse != ''")):
        if _group_of(o["line"], cfg) != LINE:
            continue
        cnt = seen.setdefault(o["warehouse"], collections.Counter())
        if o["address"]:
            cnt[o["address"]] += 1
    for r in _rows(conn.execute("SELECT DISTINCT warehouse FROM mst_mars_splits WHERE warehouse != ''")):
        seen.setdefault(r["warehouse"], collections.Counter())
    return {code: (cnt.most_common(1)[0][0] if cnt else "") for code, cnt in seen.items()}


def warehouse_rows(conn, settings=None):
    settings = settings or load_settings(conn)
    saved = {r["code"]: r for r in _rows(conn.execute("SELECT * FROM mst_mars_warehouses ORDER BY code"))}
    from_orders = _codes_in_orders(conn)
    out = []
    for code in sorted(set(saved) | set(from_orders)):
        r = saved.get(code) or {}
        row = {"code": code, "saved": bool(r),
               "name": r.get("name") or f"永豐商店酷澎-{code}",
               "address": r.get("address") or from_orders.get(code, ""),
               "address_from_orders": not r.get("address") and bool(from_orders.get(code)),
               "phone": r.get("phone") or "", "ship_to": r.get("ship_to") or "",
               "contact": r.get("contact") or settings["contact_default"],
               "updated_by": r.get("updated_by") or "", "updated_at": r.get("updated_at") or ""}
        row["missing"] = [lbl for k, lbl in (("address", "地址"), ("phone", "電話"), ("ship_to", "ship-to")) if not row[k]]
        out.append(row)
    return out


def _empty_wh(code, settings):
    return {"code": code, "name": f"永豐商店酷澎-{code}", "address": "", "phone": "", "ship_to": "",
            "contact": settings["contact_default"], "missing": ["地址", "電話", "ship-to"]}


def warehouse_info(conn, code, settings=None):
    settings = settings or load_settings(conn)
    for r in warehouse_rows(conn, settings):
        if r["code"] == code:
            return r
    return _empty_wh(code, settings)


def save_warehouse(conn, code, payload, operator):
    code = norm_text(code)
    if not code:
        return "倉別不能是空的。"
    vals = {k: str(payload.get(k) or "").strip()[:200] for k in ("name", "address", "phone", "ship_to", "contact") if k in payload}
    old = _row(conn.execute("SELECT * FROM mst_mars_warehouses WHERE code = ?", (code,)))
    stamp = now()
    if old is None:
        cols = {"name": "", "address": "", "phone": "", "ship_to": "", "contact": ""}; cols.update(vals)
        conn.execute("INSERT INTO mst_mars_warehouses (code, name, address, phone, ship_to, contact, updated_by, updated_at) VALUES (?,?,?,?,?,?,?,?)",
                     (code, cols["name"], cols["address"], cols["phone"], cols["ship_to"], cols["contact"], operator, stamp))
        changed = [k for k, v in vals.items() if v]
    else:
        changed = [k for k, v in vals.items() if v != (old[k] or "")]
        if changed:
            sets = ", ".join(f"{k} = ?" for k in changed)
            conn.execute(f"UPDATE mst_mars_warehouses SET {sets}, updated_by = ?, updated_at = ? WHERE code = ?",
                         [vals[k] for k in changed] + [operator, stamp, code])
    if changed:
        _log(conn, LINE, "", "", "", "mars_warehouse", f"瑪氏採購單倉庫資料 {code}", "", "、".join(changed), operator, "manual")
    return ""


# ── 一份採購單 ───────────────────────────────────────────────────────────────
def special_text(settings, s, items):
    slot = (s.get("slot_time") or "").strip()
    m = TIME_RE.search(slot)
    lines = []
    if s.get("label") == "V" and settings["label_line"]:
        lines.append(settings["label_line"])
    for ln in settings["special_text"].split("\n"):
        ln = ln.replace("{約倉時間}", slot).replace("{約倉開始}", m.group(0) if m else slot).strip()
        if ln:
            lines.append(ln)
    if settings["white_line"] and any((i.get("po_case_note") or "").strip() for i in items):
        lines.append(settings["white_line"])
    return "\n".join(lines)


def po_rows(items):
    """同一個下採料號（同單位）合成一列；箱數、出貨數量加總。"""
    rows = collections.OrderedDict()
    for it in items:
        key = (it.get("purchase_code") or it.get("yf_sku") or "", it.get("unit") or "")
        r = rows.get(key)
        if r is None:
            r = rows[key] = {"code": key[0], "mars_code": it.get("mars_code") or "", "name": it.get("mars_name") or it.get("product_name") or "",
                             "qty": 0, "cases": 0, "price": it.get("price"), "pcs_per_case": it.get("pcs_per_case"),
                             "inner_per_case": it.get("inner_per_case"), "note": it.get("note") or "",
                             "po_case_note": it.get("po_case_note") or "", "order_codes": []}
        r["qty"] += it.get("qty_ship") or 0
        r["cases"] += it.get("cases") or 0
        for k, src in (("note", "note"), ("po_case_note", "po_case_note"), ("price", "price")):   # 同料號的商品總表欄一樣；缺的補上
            if not r[k] and it.get(src):
                r[k] = it[src]
        if it.get("yf_sku") and it["yf_sku"] != key[0] and it["yf_sku"] not in r["order_codes"]:
            r["order_codes"].append(it["yf_sku"])
    out = []
    for r in rows.values():
        if r["order_codes"]:
            extra = "訂單料號 " + "、".join(r["order_codes"])
            r["note"] = f"{r['note']}；{extra}" if r["note"] else extra
        r["cases"] = int(round(r["cases"])) if abs(r["cases"] - round(r["cases"])) < 1e-6 else r["cases"]
        if isinstance(r["qty"], float) and r["qty"].is_integer():
            r["qty"] = int(r["qty"])
        out.append(r)
    return out


def po_missing(s, items, wh, settings=None):
    """還差什麼才能產採購單；空 list＝可以。"""
    miss = []
    if not s.get("eip_po"):
        miss.append("還沒填 EIP 採購單號")
    if not (s.get("slot_time") or "").strip():
        miss.append("還沒填約倉時間")
    if s.get("category") not in CAT_FULL:
        miss.append("分不出品類")
    if not items:
        miss.append("沒有品項")
    if wh.get("missing"):
        miss.append(f"{s.get('warehouse') or '倉'} 的{'、'.join(wh['missing'])}還沒填（採購單設定）")
    return miss


def po_filename(s):
    unit = s["unit"] if SPLIT_BY_UNIT else "箱"
    return f"永豐Mars採購單_{unit}_{s['category'] or '品類不明'}_{label_text(s['label'])}_{s['warehouse']}_{s['po_number']}.xlsx"


def _copy_style(src, dst):
    dst.font = copy.copy(src.font); dst.border = copy.copy(src.border); dst.fill = copy.copy(src.fill)
    dst.number_format = src.number_format; dst.alignment = copy.copy(src.alignment); dst.protection = copy.copy(src.protection)


def po_file(s, items, wh, settings):
    wb = openpyxl.load_workbook(TEMPLATE)
    ws = wb.active
    deliver = _dt.date.fromisoformat(s["delivery_date"])
    ws["C3"] = _dt.datetime.combine(order_date(deliver, settings["lead_days"], set(settings["holidays"])), _dt.time())
    ws["F3"] = _dt.datetime.combine(deliver, _dt.time())
    ws["C5"] = s["eip_po"]
    ws["F5"] = special_text(settings, s, items)
    ws["C6"] = wh["name"]
    ws["F6"] = settings["shelf_req"]
    ws["C7"] = wh["address"]
    ws["F7"] = int(wh["ship_to"]) if str(wh["ship_to"]).isdigit() else wh["ship_to"]
    ws["C8"] = wh["phone"]
    ws["F8"] = wh["contact"]
    ws["B9"] = CAT_FULL.get(s["category"], s["category"])
    rows = po_rows(items)
    last = FIRST_ROW + max(len(rows), 1) - 1
    for r in range(LAST_TEMPLATE_ROW + 1, last + 1):          # 超過範本格式的列：照最後一列補格式與公式
        for col in range(1, 14):
            _copy_style(ws.cell(LAST_TEMPLATE_ROW, col), ws.cell(r, col))
        ws.cell(r, 7).value = f"=IFERROR(E{r}*F{r},0)"
    for i, row in enumerate(rows):
        r = FIRST_ROW + i
        ws.cell(r, 1).value = row["code"]
        ws.cell(r, 2).value = row["mars_code"]
        ws.cell(r, 3).value = row["name"]
        ws.cell(r, 4).value = row["qty"]
        ws.cell(r, 5).value = row["cases"]
        ws.cell(r, 6).value = row["price"]
        ws.cell(r, 8).value = "V" if s["label"] == "V" else None
        ws.cell(r, 10).value = row["pcs_per_case"]
        ws.cell(r, 11).value = row["inner_per_case"]
        ws.cell(r, 12).value = row["note"] or None
        ws.cell(r, 13).value = row["po_case_note"] or None
    ws["E10"] = f"=SUM(E{FIRST_ROW}:E{last})"
    ws["F10"] = f"=SUM(G{FIRST_ROW}:G{last})"
    ws.title = "採購單"
    buf = io.BytesIO(); wb.save(buf)
    return buf.getvalue()


def _items_of(s):
    return json.loads(s["items_json"] or "[]")


# ── 路由 ─────────────────────────────────────────────────────────────────────
@mars_bp.route("/api/mars/po/settings")
def api_po_settings():
    conn = get_conn()
    try:
        st = load_settings(conn)
        return jsonify({"settings": st, "warehouses": warehouse_rows(conn, st), "defaults": DEFAULT_SETTINGS,
                        "template": os.path.basename(TEMPLATE)})
    finally:
        conn.close()


@mars_bp.route("/api/mars/po/settings", methods=["PUT"])
def api_po_settings_save():
    payload = request.get_json(silent=True) or {}
    conn = get_conn()
    try:
        new, err = save_settings(conn, payload, _operator())
        if err:
            return jsonify({"error": err}), 400
        conn.commit()
        return jsonify({"ok": True, "settings": new})
    finally:
        conn.close()


@mars_bp.route("/api/mars/warehouses/<code>", methods=["PUT"])
def api_warehouse_save(code):
    payload = request.get_json(silent=True) or {}
    conn = get_conn()
    try:
        err = save_warehouse(conn, code, payload, _operator())
        if err:
            return jsonify({"error": err}), 400
        conn.commit()
        return jsonify({"ok": True, "warehouses": warehouse_rows(conn)})
    finally:
        conn.close()


@mars_bp.route("/api/mars/splits/<int:split_id>/po")
def api_split_po(split_id):
    """一份拆單表 → 一張瑪氏採購單。差什麼就 400 講清楚。"""
    conn = get_conn()
    try:
        s = _row(conn.execute("SELECT * FROM mst_mars_splits WHERE id = ?", (split_id,)))
        if s is None:
            return jsonify({"error": "找不到這份拆單表，可能已經重拆過了。"}), 404
        st = load_settings(conn)
        wh = warehouse_info(conn, s["warehouse"], st)
    finally:
        conn.close()
    items = _items_of(s)
    miss = po_missing(s, items, wh, st)
    if miss:
        return jsonify({"error": "這份還產不了瑪氏採購單：" + "；".join(miss), "missing": miss}), 400
    return send_file(io.BytesIO(po_file(s, items, wh, st)), as_attachment=True, download_name=po_filename(s),
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@mars_bp.route("/api/mars/po/zip", methods=["POST"])
def api_po_zip():
    """這段到貨日裡、產得出來的採購單全部打包；產不了的列在 zip 裡的說明檔跟回應標頭。"""
    payload = request.get_json(silent=True) or {}
    d1, d2 = norm_text(payload.get("from")), norm_text(payload.get("to"))
    try:
        _dt.date.fromisoformat(d1); _dt.date.fromisoformat(d2)
    except ValueError:
        return jsonify({"error": "請選到貨日的起訖。"}), 400
    d1, d2 = sorted([d1, d2])
    conn = get_conn()
    try:
        st = load_settings(conn)
        whs = {w["code"]: w for w in warehouse_rows(conn, st)}
        saved = _rows(conn.execute("SELECT * FROM mst_mars_splits WHERE delivery_date >= ? AND delivery_date <= ? ORDER BY delivery_date, po_number, filename", (d1, d2)))
    finally:
        conn.close()
    if not saved:
        return jsonify({"error": "這段期間還沒有產出過拆單表，先按「產出拆單表」。"}), 400
    ok, skipped = [], []
    for s in saved:
        items = _items_of(s)
        wh = whs.get(s["warehouse"]) or _empty_wh(s["warehouse"], st)
        miss = po_missing(s, items, wh, st)
        if miss:
            skipped.append(f"{s['filename']}：{'；'.join(miss)}")
        else:
            ok.append((s, items, wh))
    if not ok:
        return jsonify({"error": "這段期間沒有一份產得出採購單。", "details": skipped[:30]}), 400
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for s, items, wh in ok:
            zf.writestr(po_filename(s), po_file(s, items, wh, st))
        if skipped:
            zf.writestr("沒產出的.txt", "這幾份還差東西，沒有產出採購單：\n\n" + "\n".join(skipped))
    buf.seek(0)
    name = f"瑪氏採購單_{d1.replace('-', '')}" + ("" if d1 == d2 else f"-{d2.replace('-', '')}") + ".zip"
    resp = send_file(buf, as_attachment=True, download_name=name, mimetype="application/zip")
    resp.headers["X-Mars-Po-Count"] = str(len(ok))
    resp.headers["X-Mars-Po-Skipped"] = str(len(skipped))
    return resp


__all__ = ["DEFAULT_SETTINGS", "load_settings", "save_settings", "parse_holiday", "order_date", "warehouse_rows", "warehouse_info",
           "save_warehouse", "special_text", "po_rows", "po_missing", "po_filename", "po_file",
           "api_po_settings", "api_po_settings_save", "api_warehouse_save", "api_split_po", "api_po_zip"]

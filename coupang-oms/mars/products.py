"""瑪氏商品總表（Alice 的「自動化_Mars整合商品資料」）：上傳就整份覆蓋。

一個永豐料號分箱／盒／包三列（L 單位類別），三列的資料幾乎一樣，只有箱入數、料號、採購單箱備註會不同，
所以鍵是 (永豐料號, 單位)。表頭用文字認，不靠欄位字母；分頁不看名稱，找到第一個有「永豐料號」「單位類別」的分頁。
"""
from .common import *  # noqa: F401,F403

# 表頭（去空白、去換行、小寫）→ 欄位
HEADERS = {
    "瑪氏貨號": "mars_code", "category": "category", "brand": "brand", "永豐料號": "yf_sku",
    "品名(永豐建檔)": "name", "系統價格(未稅)": "price", "每中盒包數": "per_inner", "每箱中盒數": "inner_per_case",
    "每箱產品數(最小單位)": "pcs_per_case", "料號": "item_code", "料號2": "item_code2", "單位類別": "unit",
    "單位代碼": "unit_code", "箱入數": "box_qty", "中盒貼標": "inner_label", "效期天": "shelf_days",
    "備註": "note", "採購單箱備註": "po_case_note",
}
REQUIRED = ("yf_sku", "unit", "category", "inner_label")
NUMERIC = ("price", "per_inner", "inner_per_case", "pcs_per_case", "box_qty")


def _h(v):
    return norm_text(v).lower().replace(" ", "").replace("\n", "").replace("（", "(").replace("）", ")")


def _find_sheet(wb):
    for ws in wb.worksheets:
        for idx, row in enumerate(ws.iter_rows(min_row=1, max_row=10, values_only=True), start=1):
            cols = {}
            for i, v in enumerate(row):
                f = HEADERS.get(_h(v))
                if f and f not in cols:
                    cols[f] = i
            if all(k in cols for k in REQUIRED):
                return ws, idx, cols
    return None, None, None


def parse_master(fileobj):
    """回傳 (rows, warnings)。rows 已正規化；有問題的列跳過並寫進 warnings。"""
    try:
        wb = openpyxl.load_workbook(fileobj, data_only=True, read_only=True)
    except Exception as exc:  # noqa: BLE001
        raise ValueError("無法開啟此 Excel 檔，請確認檔案格式。") from exc
    ws, hdr, cols = _find_sheet(wb)
    if ws is None:
        raise ValueError("找不到瑪氏商品總表：要有「永豐料號」「單位類別」「Category」「中盒貼標」這幾個欄位。")
    rows, warnings, seen = [], [], {}
    for n, raw in enumerate(ws.iter_rows(min_row=hdr + 1, values_only=True), start=hdr + 1):
        def get(f):
            i = cols.get(f)
            return raw[i] if i is not None and i < len(raw) else None
        yf = norm_key(get("yf_sku"))
        if not yf:
            continue
        unit = norm_text(get("unit"))
        if unit not in UNITS:
            warnings.append(f"第 {n} 列 {yf}：單位類別是「{unit or '空白'}」，不是箱／盒／包，已略過。"); continue
        rec = {"yf_sku": yf, "unit": unit}
        for f in ("mars_code", "item_code", "item_code2"):
            rec[f] = norm_key(get(f))
        for f in ("category", "brand", "name", "unit_code", "note", "po_case_note"):
            rec[f] = norm_text(get(f))
        for f in NUMERIC:
            v = get(f)
            d = norm_decimal(v) if v not in (None, "") else None
            rec[f] = float(d) if d is not None else None
        rec["inner_label"] = "V" if norm_text(get("inner_label")).upper() == "V" else ""
        rec["shelf_days"] = norm_int(get("shelf_days"))
        if not cat_code(rec["category"]):
            warnings.append(f"第 {n} 列 {yf}：Category 是「{rec['category'] or '空白'}」，不是 Chocolate／Gum／Petcare，拆單時會標示提醒。")
        key = (yf, unit)
        if key in seen:
            warnings.append(f"第 {n} 列 {yf}（{unit}）與第 {seen[key]} 列重複，以後者為準。")
            rows = [r for r in rows if (r["yf_sku"], r["unit"]) != key]
        seen[key] = n
        rows.append(rec)
    if not rows:
        raise ValueError("這份商品總表裡沒有讀到任何商品。")
    return rows, warnings


COLS = ("yf_sku", "unit", "mars_code", "category", "brand", "name", "price", "per_inner", "inner_per_case",
        "pcs_per_case", "item_code", "item_code2", "unit_code", "box_qty", "inner_label", "shelf_days", "note", "po_case_note")


def load_products(conn):
    """{(永豐料號, 單位): 商品}、{永豐料號: {單位: 商品}}"""
    rows = _rows(conn.execute("SELECT * FROM mst_mars_products"))
    by_key = {(r["yf_sku"], r["unit"]): r for r in rows}
    by_code = collections.defaultdict(dict)
    for r in rows:
        by_code[r["yf_sku"]][r["unit"]] = r
    return by_key, by_code


@mars_bp.route("/api/mars/products/import", methods=["POST"])
def api_import_mars_products():
    up = request.files.get("file")
    if up is None or not up.filename:
        return jsonify({"error": "沒有收到檔案。"}), 400
    try:
        rows, warnings = parse_master(io.BytesIO(up.read()))
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400
    operator = _operator()
    conn = get_conn()
    try:
        before = {r["yf_sku"] for r in _rows(conn.execute("SELECT DISTINCT yf_sku FROM mst_mars_products"))}
        # 在蝦皮特選補登、Excel 裡還沒有的品項留著（Jerry 2026-10-14）；Excel 有了就以 Excel 為準
        new_keys = {(r["yf_sku"], r["unit"]) for r in rows}
        keep = [m for m in _rows(conn.execute("SELECT * FROM mst_mars_products WHERE COALESCE(manual_by, '') != ''"))
                if (m["yf_sku"], m["unit"]) not in new_keys]
        conn.execute("DELETE FROM mst_mars_products")
        for r in rows:
            conn.execute(f"INSERT INTO mst_mars_products ({', '.join(COLS)}) VALUES ({', '.join('?' * len(COLS))})",
                         [r[c] for c in COLS])
        kcols = COLS + ("manual_by", "manual_at")
        for m in keep:
            conn.execute(f"INSERT INTO mst_mars_products ({', '.join(kcols)}) VALUES ({', '.join('?' * len(kcols))})",
                         [m[c] for c in kcols])
        codes = {r["yf_sku"] for r in rows} | {m["yf_sku"] for m in keep}
        conn.execute("INSERT INTO mst_mars_uploads (filename, operator, uploaded_at, rows_total, codes) VALUES (?,?,?,?,?)",
                     (up.filename, operator, now(), len(rows), len(codes)))
        added, removed = sorted(codes - before), sorted(before - codes)
        _log(conn, LINE, "", "", "", "mars_products", "瑪氏商品總表", f"{len(before)} 個料號", f"{len(codes)} 個料號",
             operator, "import", f"{up.filename}；新增 {len(added)}、拿掉 {len(removed)}"
             + (f"；保留系統補登 {len(keep)} 品" if keep else ""))
        conn.commit()
    finally:
        conn.close()
    by_cat = collections.Counter(cat_code(r["category"]) or "?" for r in rows if r["unit"] == "箱")
    return jsonify({"ok": True, "filename": up.filename, "rows": len(rows), "codes": len(codes),
                    "by_category": dict(by_cat), "added": len(added), "removed": len(removed),
                    "added_examples": added[:5], "removed_examples": removed[:5], "warnings": warnings[:50],
                    "kept_manual": [f"{m['yf_sku']}（{m['unit']}）" for m in keep]})


@mars_bp.route("/api/mars/products")
def api_search_mars_products():
    """查商品總表（Jerry 2026-10-13：「瑪氏沒有搜尋功能喔」——蝦皮特選說商品總表找不到某料號，要能自己查）。
    空白隔開多個關鍵字，全部符合才列；備註有換號紀錄，所以舊料號也查得到。"""
    words = [w for w in str(request.args.get("q") or "").split() if w][:5]
    if not words:
        return jsonify({"rows": [], "total": 0})
    cols = ("yf_sku", "unit", "mars_code", "name", "item_code", "item_code2", "category", "brand", "note")
    where, params = [], []
    for w in words:            # LIKE 樣式一律綁參數（PostgreSQL 字串裡不能有裸的 %）
        where.append("(" + " OR ".join(f"LOWER({c}) LIKE ?" for c in cols) + ")")   # 大小寫都找得到（Postgres 的 LIKE 分大小寫）
        params += [f"%{w.lower()}%"] * len(cols)
    conn = get_conn()
    try:
        total = conn.execute(f"SELECT COUNT(*) AS n FROM mst_mars_products WHERE {' AND '.join(where)}", params).fetchone()["n"]
        rows = _rows(conn.execute(
            f"SELECT yf_sku, unit, mars_code, name, category, box_qty, inner_label, shelf_days, note, manual_by, manual_at FROM mst_mars_products "
            f"WHERE {' AND '.join(where)} ORDER BY yf_sku, unit LIMIT 100", params))
    finally:
        conn.close()
    return jsonify({"rows": rows, "total": total})


@mars_bp.route("/api/mars/status")
def api_mars_status():
    conn = get_conn()
    try:
        last = _row(conn.execute("SELECT * FROM mst_mars_uploads ORDER BY id DESC LIMIT 1"))
        n = conn.execute("SELECT COUNT(DISTINCT yf_sku) AS n FROM mst_mars_products").fetchone()["n"]
        manual = [f"{r['yf_sku']}（{r['unit']}）" for r in _rows(conn.execute(
            "SELECT yf_sku, unit FROM mst_mars_products WHERE COALESCE(manual_by, '') != '' ORDER BY yf_sku, unit"))]
        dates = [r["d"] for r in _rows(conn.execute(
            "SELECT DISTINCT delivery_date AS d, line FROM mst_orders WHERE delivery_date != '' ORDER BY delivery_date"))
            if _group_of(r["line"], _line_groups()) == LINE]
        # 最近一次有瑪氏單的匯入（不管是從這頁還是 ② 傳的）
        last_imp = None
        cfg = _line_groups()
        for b in _rows(conn.execute("SELECT id, filename, operator, committed_at, payload_json FROM mst_import_batches WHERE committed = 1 ORDER BY id DESC LIMIT 20")):
            try:
                rows_b = json.loads(b["payload_json"] or "{}").get("rows", [])
            except ValueError:
                rows_b = []
            if any(_group_of(r.get("line"), cfg) == LINE for r in rows_b):
                last_imp = {k: b[k] for k in ("id", "filename", "operator", "committed_at")}; break
    finally:
        conn.close()
    return jsonify({"last_upload": last, "codes": n, "manual": manual, "dates": sorted(set(dates)), "last_order_import": last_imp,
                    "split_by_unit": SPLIT_BY_UNIT})


@mars_bp.route("/mars")
def mars_page():
    return render_template("mars.html", logged_in_user=_operator(),
                           build_version=current_app.config.get("BUILD_VERSION", ""))


__all__ = ["parse_master", "load_products", "api_import_mars_products", "api_mars_status", "mars_page"]

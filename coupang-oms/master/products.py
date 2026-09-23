"""① 商品主檔：清單、手動存、刪除、從酷澎主檔／寶僑總表匯入（含把總表存成匯出樣式）。"""
from .common import *  # noqa: F401,F403 — 共用工具、Flask、db、openpyxl 都從這裡來
from . import sources


@master_bp.route("/api/master/products")
def api_products():
    line = norm_text(request.args.get("line"))
    q = norm_text(request.args.get("q")).lower()
    cfg = _line_groups()
    conn = get_conn()
    try:
        rows = _rows(conn.execute(
            """SELECT p.*, (SELECT COUNT(*) FROM mst_orders o WHERE o.barcode = p.barcode) AS order_rows
               FROM mst_products p ORDER BY p.brand, p.product_name, p.barcode"""))
        orphans = _rows(conn.execute(
            """SELECT o.barcode, MAX(o.sku_id) AS sku_id, MAX(o.yf_sku) AS yf_sku, MAX(o.brand) AS brand,
                      MAX(o.product_name) AS product_name, MAX(o.box_size_file) AS box_size_file,
                      MAX(o.line) AS line, COUNT(*) AS order_rows
               FROM mst_orders o LEFT JOIN mst_products p ON p.barcode = o.barcode
               WHERE p.id IS NULL AND o.barcode != '' GROUP BY o.barcode ORDER BY o.barcode"""))
    finally:
        conn.close()
    for p in rows:
        raws = _split_lines(p["lines_seen"])
        p["line_groups"] = sorted({_group_of(r, cfg) for r in raws})
        p["lines_raw"] = raws
    for o in orphans:
        o["line_group"] = _group_of(o["line"], cfg)
    if line:
        rows = [p for p in rows if line in p["line_groups"]]
        orphans = [o for o in orphans if o["line_group"] == line]
    if q:
        rows = [p for p in rows if q in " ".join(str(p.get(k) or "") for k in
                ("barcode", "sku_id", "yf_sku", "brand", "product_name", "note", "category", "lines_seen", "unit", "master_line")).lower()]
    return jsonify({"products": rows, "orphans": orphans, "total": len(rows)})


def _upsert_product(conn, barcode, fields, box, cost, operator, source, fname="", only_filled=False, guard=None):
    """新增或更新主檔一筆。only_filled=True 時（總表匯入）只更新有值的欄位。
    guard（SheetGuard）有給 = 這是重匯總表：系統上改過的格子要先問它能不能蓋。
    人在系統上改（source=manual）的欄位記進 mst_field_src；酷澎主檔匯入蓋過的欄位從 mst_field_src 拿掉。"""
    existing = _row(conn.execute("SELECT * FROM mst_products WHERE barcode = ?", (barcode,)))
    stamp = now()
    if cost is not None:
        cost = round(float(cost), 4)   # Excel 存的 137.0000000000 讀出來會是 137.00000000000003
    fields = dict(fields)
    for k in _PRODUCT_TEXT_FIELDS:
        fields.setdefault(k, "")
    fields.setdefault("shelf_days", None)
    master_line = fields.get("master_line") or ""

    def track(k):
        if source == "manual":
            _mark_src(conn, "product", barcode, k, "manual", operator)
        elif guard is None:
            _clear_src(conn, "product", barcode, k)

    if existing is None:
        conn.execute(
            """INSERT INTO mst_products
               (barcode, sku_id, yf_sku, brand, product_name, category, pgcode, cost_price,
                box_size, note, auto_created, lines_seen, master_line, unit, shelf_days, active,
                date_format, updated_by, updated_at)
               VALUES (?,?,?,?,?,?,?,?,?,?,0,?,?,?,?,?,?,?,?)""",
            (barcode, fields["sku_id"], fields["yf_sku"], fields["brand"], fields["product_name"],
             fields["category"], fields["pgcode"], cost, box, fields["note"], master_line, master_line,
             fields["unit"], fields["shelf_days"], fields["active"] or "Y", fields["date_format"],
             operator, stamp))
        _log(conn, master_line, "", fields["sku_id"], barcode, "product_new", "新增主檔", "", box, operator, source, fname)
        if source == "manual":
            for k in ("sku_id", "yf_sku", "brand", "product_name", "category", "pgcode", "note"):
                if fields[k]:
                    track(k)
            if cost is not None:
                track("cost_price")
            if box:
                track("box_size")
        return "added"
    sets, params, changed = [], [], False
    if master_line and master_line not in _split_lines(existing.get("lines_seen")):
        merged = sorted(set(_split_lines(existing.get("lines_seen"))) | {master_line})
        sets.append("lines_seen = ?"); params.append(",".join(merged)); changed = True
    for k, v in fields.items():
        if only_filled and (v is None or v == ""):
            continue
        if not _same(existing[k], v):
            if guard is not None and not guard.allow("product", barcode, k, existing[k], v):
                continue
            sets.append(f"{k} = ?"); params.append(v); changed = True; track(k)
            if k == "note":
                _log(conn, "", "", existing["sku_id"], barcode, "note", "Note", existing["note"], v,
                     operator, source, fname)
    if (cost is not None or not only_filled) and not _same(existing["cost_price"], cost) \
            and (guard is None or guard.allow("product", barcode, "cost_price", existing["cost_price"], cost)):
        sets.append("cost_price = ?"); params.append(cost); changed = True; track("cost_price")
    if (box or not only_filled) and not _same(existing["box_size"], box) \
            and (guard is None or guard.allow("product", barcode, "box_size", existing["box_size"], box)):
        sets.append("box_size = ?"); params.append(box); changed = True; track("box_size")
        _log(conn, "", "", existing["sku_id"], barcode, "box_size", "箱入數", existing["box_size"], box,
             operator, source, fname)
    if existing["auto_created"]:
        sets.append("auto_created = 0"); changed = True
    if changed:
        sets += ["updated_by = ?", "updated_at = ?"]; params += [operator, stamp]
        conn.execute(f"UPDATE mst_products SET {', '.join(sets)} WHERE id = ?", params + [existing["id"]])
        return "updated"
    return "unchanged"


@master_bp.route("/api/master/products", methods=["POST"])
def api_save_product():
    payload = request.get_json(silent=True) or {}
    barcode = norm_key(payload.get("barcode"))
    if not barcode:
        return jsonify({"error": "國條必填。"}), 400
    box = norm_int(payload.get("box_size"))
    if payload.get("box_size") not in (None, "") and (box is None or box <= 0):
        return jsonify({"error": "箱入數要是正整數。"}), 400
    cost = None if payload.get("cost_price") in (None, "") else norm_decimal(payload.get("cost_price"))
    if payload.get("cost_price") not in (None, "") and cost is None:
        return jsonify({"error": "單價(含稅)要是數字。"}), 400
    fields = {k: (norm_key(payload.get(k)) if k in ("sku_id", "yf_sku") else norm_text(payload.get(k)))
              for k in _PRODUCT_TEXT_FIELDS}
    fields["active"] = (fields["active"] or "Y").upper()[:1]
    fields["shelf_days"] = norm_int(payload.get("shelf_days"))
    prices = {}
    for k in ("giv", "niv"):
        if k in payload:
            v = payload.get(k)
            prices[k] = None if v in (None, "") else norm_decimal(v)
            if v not in (None, "") and prices[k] is None:
                return jsonify({"error": f"{k.upper()} 要是數字。"}), 400
    operator = _operator()
    conn = get_conn()
    try:
        _upsert_product(conn, barcode, fields, box, cost, operator, "manual")
        for k, v in prices.items():
            sources.set_product_value(conn, barcode, k, v, operator)
        conn.commit()
        fresh = _row(conn.execute("SELECT * FROM mst_products WHERE barcode = ?", (barcode,)))
        return jsonify({"ok": True, "product": fresh})
    finally:
        conn.close()


@master_bp.route("/api/master/products/<int:product_id>", methods=["DELETE"])
def api_delete_product(product_id):
    conn = get_conn()
    try:
        p = _row(conn.execute("SELECT * FROM mst_products WHERE id = ?", (product_id,)))
        if p is None:
            return jsonify({"error": "找不到這筆主檔。"}), 404
        conn.execute("DELETE FROM mst_products WHERE id = ?", (product_id,))
        _log(conn, "", "", p["sku_id"], p["barcode"], "product_delete", "刪除主檔", p["box_size"], "",
             _operator(), "manual")
        conn.commit()
        return jsonify({"ok": True})
    finally:
        conn.close()


# 主檔匯入認兩種檔：
#   ・酷澎主檔（一個線別一份）：SKU ID／條碼(國條)／永豐料號／線別／品牌／品名／單位／箱入數／
#     業務報價單價(含稅)／總效期天數／啟用(Y/N)／日期格式／報價備註
#   ・寶僑總表 Sheet1：skuid／Category／Brand／Pgcode／Barcode／永豐料號／SKU Name／COGS／Note／箱入數
# 都是認標題不認欄位位置。業務報價單價(含稅) = COGS，同一個欄位；報價備註 = Note，同一個欄位。
_HEADER_ALIASES = {
    "barcode": ["barcode", "國條", "條碼", "條碼(國條)", "國際條碼", "ean"],
    "category": ["category", "品類", "類別"],
    "pgcode": ["pgcode", "pg code", "pg_code"],
    "cost_price": ["cogs (pcs/ w. tax)", "cogs", "cogs(pcs)", "業務報價單價(含稅)", "業務報價單價", "報價單價",
                   "單價(含稅)", "成本(含稅)", "cost"],
    "sku_id": ["skuid", "sku id", "sku_id", "skuno.", "sku編號"],
    "yf_sku": ["永豐料號", "料號"],
    "brand": ["brand", "品牌"],
    "product_name": ["sku name", "品名", "product name", "名稱"],
    "box_size": ["箱入數", "箱入", "case pack", "pcs/cs", "轉換率"],
    "note": ["note", "報價備註", "備註", "說明"],
    "master_line": ["線別"],
    "unit": ["單位"],
    "shelf_days": ["總效期天數", "效期天數", "效期"],
    "active": ["啟用(y/n)", "啟用"],
    "date_format": ["日期格式"],
}


_PRODUCT_TEXT_FIELDS = ("sku_id", "yf_sku", "brand", "product_name", "category", "pgcode", "note",
                        "master_line", "unit", "active", "date_format")


def _find_columns(ws, want):
    best = (None, {}, [])
    for idx, row in enumerate(ws.iter_rows(min_row=1, max_row=15, values_only=True), start=1):
        headers = [norm_text(c) for c in row]
        lowered = [h.lower().replace(" ", "") for h in headers]
        mapping = {}
        for field, aliases in _HEADER_ALIASES.items():
            for a in aliases:
                key = a.lower().replace(" ", "")
                if key in lowered:
                    mapping[field] = lowered.index(key); break
        if all(f in mapping for f in want) and len(mapping) > len(best[1]):
            best = (idx, mapping, headers)
    return best


@master_bp.route("/api/master/products/import", methods=["POST"])
def api_import_products():
    """匯入主檔：認「酷澎主檔」（一個線別一份，有線別欄）和「寶僑總表 Sheet1」兩種格式，
    看標題自動分辨（見 _HEADER_ALIASES）。只更新有值的欄位，不會把既有資料洗成空白。
    酷澎主檔的線別會同時記進 lines_seen，篩選、總表立刻分得出線別，不用等訂單來學。"""
    upload = request.files.get("file")
    if upload is None or not upload.filename:
        return jsonify({"error": "沒有收到檔案。"}), 400
    try:
        raw = upload.read()
        wb = openpyxl.load_workbook(io.BytesIO(raw), data_only=True, read_only=True)
    except Exception as exc:  # noqa: BLE001
        return jsonify({"error": f"無法開啟 Excel：{exc}"}), 400
    # 先看是不是三種外部來源檔（supply 表／Coupang Master／庫存銷售表），是的話走 sources.py，
    # 只更新它們負責的欄位，不碰品名、品牌那些。
    kind, sws, sidx, shs, _raw = sources.find_source_sheet(wb)
    if kind:
        conn = get_conn()
        try:
            fn = {"supply": sources.import_supply, "master_price": sources.import_master_price, "stock": sources.import_stock}[kind]
            result = fn(conn, sws, sidx, shs, _operator(), upload.filename)
            conn.commit()
        finally:
            conn.close()
        result.update({"ok": True, "added": 0, "unchanged": 0, "template_line": "", "columns_found": []})
        return jsonify(result)
    chosen = None
    for ws in wb.worksheets:
        hdr_idx, mapping, _ = _find_columns(ws, want=("barcode", "box_size"))
        if hdr_idx:
            chosen = (ws, hdr_idx, mapping); break
    if chosen is None:
        return jsonify({"error": "找不到同時有「國條／Barcode」和「箱入數」標題的工作表，也不是 supply 表／Coupang Master／庫存銷售表。"}), 400
    ws, hdr_idx, mapping = chosen

    def get(row, field):
        i = mapping.get(field)
        return row[i] if i is not None and i < len(row) else None

    operator = _operator()
    header_cells = next(ws.iter_rows(min_row=hdr_idx, max_row=hdr_idx, values_only=True))
    is_sheet = any(_DATE_HDR.match(norm_text(h)) for h in header_cells)   # 有「M/D交貨」日期欄 → 這是業務的總表
    mode = norm_text(request.form.get("on_conflict"))
    if mode not in ("", "keep", "overwrite"):
        mode = ""
    seen_barcodes = []
    conn = get_conn()
    try:
        guard = SheetGuard(conn, mode) if is_sheet else None
        counts = collections.Counter()
        for row in ws.iter_rows(min_row=hdr_idx + 1, values_only=True):
            barcode = norm_key(get(row, "barcode"))
            if not barcode:
                continue
            seen_barcodes.append(barcode)
            fields = {k: (norm_key(get(row, k)) if k in ("sku_id", "yf_sku") else norm_text(get(row, k)))
                      for k in _PRODUCT_TEXT_FIELDS}
            fields["active"] = fields["active"].upper()[:1] if fields["active"] else ""
            fields["shelf_days"] = norm_int(get(row, "shelf_days"))
            counts[_upsert_product(conn, barcode, fields, norm_int(get(row, "box_size")),
                                   norm_decimal(get(row, "cost_price")), operator, "import",
                                   upload.filename, only_filled=True, guard=guard)] += 1
        template_line = ""; extras = {}
        if is_sheet and seen_barcodes:
            # 業務的總表就是「匯出總表要長的樣子」，整份記起來當底稿，之後匯出只填箱數。
            # 總表沒有線別欄，線別看這些商品在主檔屬於誰（多數決）；都不知道就當寶僑——
            # 目前只有寶僑有這種總表。
            template_line = _guess_line(conn, seen_barcodes) or "寶僑"
            _save_template(conn, template_line, upload.filename, ws.title, hdr_idx, raw, operator)
            # 這份是寶僑的總表，裡面的商品就算寶僑的：lines_seen 補上，沒出過貨、只在總表裡的商品
            # 看板與匯出才會把它們算進寶僑（不然 Supply／庫存有值卻填不回總表）
            for i in range(0, len(seen_barcodes), 400):
                chunk = seen_barcodes[i:i + 400]
                for pr in _rows(conn.execute(f"SELECT id, lines_seen FROM mst_products WHERE barcode IN ({','.join('?' * len(chunk))})", chunk)):
                    cur_lines = _split_lines(pr["lines_seen"])
                    if template_line not in cur_lines:
                        conn.execute("UPDATE mst_products SET lines_seen = ? WHERE id = ?", (",".join(sorted(set(cur_lines) | {template_line})), pr["id"]))
            # 總表 K／L 的 GIV／NIV、M～R 的 Supply／demand 也一起帶進來（read_only 的 ws 要重開一次才能再讀）
            ws2 = openpyxl.load_workbook(io.BytesIO(raw), data_only=True, read_only=True)[ws.title]
            extras = sources.absorb_sheet_extras(conn, ws2, hdr_idx, list(header_cells), operator, upload.filename, template_line, guard=guard)
        if guard is not None and guard.conflicts and mode == "":
            # 總表要蓋掉系統上改過／來源檔更新過的數字：這次什麼都不寫，把清單交給畫面問人
            items = _describe_conflicts(conn, guard.conflicts)
            conn.rollback()
            return jsonify({"ok": False, "needs_confirm": True, "filename": upload.filename, "count": len(items), "conflicts": items[:300]})
        conn.commit()
    finally:
        conn.close()
    kept = len(guard.conflicts) if guard is not None and mode == "keep" else 0
    overwritten = len(guard.conflicts) if guard is not None and mode == "overwrite" else 0
    return jsonify({"ok": True, "kind": "sheet" if is_sheet else "coupang_master", "sheet": ws.title,
                    "added": counts["added"], "updated": counts["updated"],
                    "unchanged": counts["unchanged"], "template_line": template_line,
                    "columns_found": sorted(mapping.keys()), "extras": extras, "kept": kept, "overwritten": overwritten})


def _describe_conflicts(conn, conflicts):
    """衝突清單補上品名，給畫面列表用。"""
    names = {r["barcode"]: r["product_name"] for r in _rows(conn.execute("SELECT barcode, product_name FROM mst_products"))}
    out = []
    for c in conflicts:
        parts = c["rkey"].split("|")
        if c["kind"] == "brand":
            who, where = parts[2], f"品牌目標（{_quarter_label(parts[1])}）"
        else:
            who, where = f"{names.get(parts[0], '')} {parts[0]}".strip(), (f"{int(parts[1][5:7])} 月" if c["kind"] == "month" else "")
        out.append({"item": who, "field": (c["field_label"] + (f"（{where}）" if where and c["kind"] == "month" else "")) if c["kind"] != "brand" else f"{c['field_label']}（{_quarter_label(parts[1])}）",
                    "current": c["current"], "incoming": c["incoming"],
                    "from": c["src_label"] + (f"（{c['operator']}）" if c["src"] == "manual" and c["operator"] else ""), "at": (c["at"] or "")[:16]})
    return out


_DATE_HDR = re.compile(r"^\s*(\d{1,2})/(\d{0,2})交貨")   # 9/2交貨、7/9交貨_1、9/18交貨\n竹運出、9/交貨（空欄）


def _template_row(conn, line):
    return _row(conn.execute("SELECT * FROM mst_templates WHERE line = ?", (line,)))


def _guess_line(conn, barcodes):
    """一批國條多數屬於哪個線別：先看主檔給的線別，再看訂單學到的。"""
    votes = collections.Counter()
    for i in range(0, len(barcodes), 400):
        chunk = barcodes[i:i + 400]
        rows = _rows(conn.execute(
            f"SELECT master_line, lines_seen FROM mst_products WHERE barcode IN ({','.join('?' * len(chunk))})", chunk))
        for r in rows:
            if r["master_line"]:
                votes[r["master_line"]] += 1
            else:
                for l in _split_lines(r["lines_seen"]):
                    votes[l] += 1
    if not votes:
        return ""
    return _group_of(votes.most_common(1)[0][0])


def _save_template(conn, line, filename, sheet, header_row, raw, operator):
    import base64
    old = _template_row(conn, line)
    conn.execute("DELETE FROM mst_templates WHERE line = ?", (line,))
    conn.execute(
        """INSERT INTO mst_templates (line, filename, sheet, header_row, content_b64, uploaded_by, uploaded_at)
           VALUES (?,?,?,?,?,?,?)""",
        (line, filename, sheet, header_row, base64.b64encode(raw).decode("ascii"), operator, now()))
    _log(conn, line, "", "", "", "template", "總表樣式", old["filename"] if old else "", filename, operator,
         "import", f"匯出總表照這份的樣子（工作表「{sheet}」）")


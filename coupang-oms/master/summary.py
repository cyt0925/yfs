"""③ 總表：現算每日箱數、月配額、匯出總表（有底稿就照底稿填、缺的日期自動插欄）。"""
from .common import *  # noqa: F401,F403 — 共用工具、Flask、db、openpyxl 都從這裡來
from .products import _DATE_HDR, _HEADER_ALIASES, _template_row


@master_bp.route("/api/master/quota", methods=["PUT"])
def api_set_quota():
    payload = request.get_json(silent=True) or {}
    line = norm_text(payload.get("line")); barcode = norm_key(payload.get("barcode"))
    month = norm_text(payload.get("month"))
    if not line or not barcode or not _valid_month(month):
        return jsonify({"error": "線別、國條、月份（2026-09）都要有。"}), 400
    raw = payload.get("qty_cases")
    qty = None if raw in (None, "") else norm_decimal(raw)
    if raw not in (None, "") and qty is None:
        return jsonify({"error": "配額要是數字。"}), 400
    operator = _operator()
    conn = get_conn()
    try:
        existing = _row(conn.execute(
            "SELECT * FROM mst_quotas WHERE line = ? AND barcode = ? AND month = ?", (line, barcode, month)))
        stamp = now(); old = existing["qty_cases"] if existing else None
        if existing is None:
            conn.execute("""INSERT INTO mst_quotas (line, barcode, month, qty_cases, note, updated_by, updated_at)
                            VALUES (?,?,?,?,?,?,?)""", (line, barcode, month, qty, norm_text(payload.get("note")), operator, stamp))
        else:
            conn.execute("UPDATE mst_quotas SET qty_cases = ?, updated_by = ?, updated_at = ? WHERE id = ?",
                         (qty, operator, stamp, existing["id"]))
        if not _same(old, qty):
            _log(conn, line, "", "", barcode, "quota", f"{month} 配額", old, qty, operator, "manual")
        conn.commit()
        return jsonify({"ok": True, "qty_cases": qty})
    finally:
        conn.close()


def _build_summary(conn, group, month, cfg):
    products = _products_map(conn)
    orders = [o for o in _month_orders(conn, month, cfg)
              if o["delivery_date"] and o["line_group"] == group]
    quotas = {q["barcode"]: q for q in _rows(conn.execute(
        "SELECT * FROM mst_quotas WHERE line = ? AND month = ?", (group, month)))}

    by_bc = collections.OrderedDict(); dates = set()
    for o in orders:
        bc = o["barcode"] or f"(無國條) {o['sku_id']}"
        e = by_bc.setdefault(bc, {"by_date": collections.defaultdict(float), "month_total": 0.0,
                                  "missing_box": 0, "rows": 0, "name": o["product_name"],
                                  "brand": o["brand"], "sku_id": o["sku_id"], "yf_sku": o["yf_sku"],
                                  "box_size": o["box_size"], "lines": set()})
        e["rows"] += 1; e["lines"].add(o["line"] or ""); dates.add(o["delivery_date"])
        if o["cases"] is None:
            e["missing_box"] += 1
        else:
            e["by_date"][o["delivery_date"]] += o["cases"]; e["month_total"] += o["cases"]
    dates = sorted(dates)

    # 畫面上的總表只列「這個月有交貨」的商品；主檔裡有但沒出貨的不列（照 Chloe 的說法，
    # 總表只吃箱數）。寶僑要一模一樣的匯出走底稿那條路（見 _fill_template）。
    barcodes = list(by_bc)
    rows = []
    for bc in barcodes:
        p = products.get(bc); e = by_bc.get(bc); quota = quotas.get(bc)
        q = quota["qty_cases"] if quota else None
        total = round(e["month_total"], 2) if e else 0.0
        rows.append({
            "barcode": bc, "category": p["category"] if p else "",
            "sku_id": (p or e or {}).get("sku_id", ""), "yf_sku": (p or e or {}).get("yf_sku", ""),
            "brand": (p["brand"] if p and p["brand"] else (e["brand"] if e else "")),
            "product_name": (p["product_name"] if p and p["product_name"] else (e["name"] if e else "")),
            "box_size": p["box_size"] if p and p["box_size"] else (e["box_size"] if e else None),
            "in_master": p is not None, "note": p["note"] if p else "",
            "lines_raw": sorted(e["lines"]) if e else _split_lines(p["lines_seen"]) if p else [],
            "by_date": {d: round(v, 2) for d, v in e["by_date"].items()} if e else {},
            "month_total": total, "missing_box": e["missing_box"] if e else 0,
            "order_rows": e["rows"] if e else 0,
            "quota": q, "remaining": (round(q - total, 2) if q is not None else None),
        })
    rows.sort(key=lambda r: (r["brand"], r["product_name"], r["barcode"]))
    totals_by_date = {d: round(sum(r["by_date"].get(d, 0) for r in rows), 2) for d in dates}
    return {"line": group, "month": month, "dates": dates, "rows": rows, "totals_by_date": totals_by_date,
            "month_total": round(sum(r["month_total"] for r in rows), 2),
            "missing_box_rows": sum(r["missing_box"] for r in rows),
            "not_in_master": sum(1 for r in rows if not r["in_master"])}


@master_bp.route("/api/master/summary")
def api_summary():
    group = norm_text(request.args.get("line")); month = norm_text(request.args.get("month")) or _this_month()
    if not group:
        return jsonify({"error": "請選線別。"}), 400
    if not _valid_month(month):
        return jsonify({"error": "月份格式要像 2026-09。"}), 400
    cfg = _line_groups()
    conn = get_conn()
    try:
        return jsonify(_build_summary(conn, group, month, cfg))
    finally:
        conn.close()


def _template_meta(tpl, conn=None):
    """給畫面看的：檔名、誰、何時、底稿裡有哪幾個月的日期欄、各月還剩幾個空欄。"""
    import base64
    meta = {"exists": True, "line": tpl["line"], "filename": tpl["filename"], "sheet": tpl["sheet"],
            "uploaded_by": tpl["uploaded_by"], "uploaded_at": tpl["uploaded_at"], "months": {}}
    try:
        wb = openpyxl.load_workbook(io.BytesIO(base64.b64decode(tpl["content_b64"])), read_only=True, data_only=True)
        ws = wb[tpl["sheet"]]
        for row in ws.iter_rows(min_row=tpl["header_row"], max_row=tpl["header_row"], values_only=True):
            for h in row:
                m = _DATE_HDR.match(norm_text(h))
                if m:
                    e = meta["months"].setdefault(int(m.group(1)), {"dates": 0, "spare": 0})
                    e["dates" if m.group(2) else "spare"] += 1
    except Exception as exc:  # noqa: BLE001
        meta["error"] = str(exc)
    return meta


@master_bp.route("/api/master/template")
def api_template_get():
    line = norm_text(request.args.get("line"))
    conn = get_conn()
    try:
        tpl = _template_row(conn, line)
    finally:
        conn.close()
    return jsonify(_template_meta(tpl) if tpl else {"exists": False, "line": line})


_REF_PART = re.compile(r"^(\$?)([A-Za-z]{1,3})(\$?)(\d*)$")


def _shift_ref_text(ref, at):
    """一個儲存格參照（A1、$B$2、AO2:BC2、$F:$AL）遇到「在第 at 欄前面插一欄」要怎麼改。
    有 '!' 的是別的工作表或外部檔，不動。"""
    if "!" in ref:
        return ref
    parts = ref.split(":")
    out = []
    for part in parts:
        m = _REF_PART.match(part)
        if not m:
            return ref
        d1, letters, d2, row = m.groups()
        from openpyxl.utils import column_index_from_string, get_column_letter
        idx = column_index_from_string(letters.upper())
        if idx >= at:
            letters = get_column_letter(idx + 1)
        out.append(f"{d1}{letters}{d2}{row}")
    return ":".join(out)


def _shift_formula(formula, at):
    from openpyxl.formula.tokenizer import Tokenizer, Token
    try:
        tok = Tokenizer(formula)
    except Exception:  # noqa: BLE001
        return formula
    pieces = []
    for t in tok.items:
        if t.type == Token.OPERAND and t.subtype == Token.RANGE:
            pieces.append(_shift_ref_text(t.value, at))
        else:
            pieces.append(t.value)
    return "=" + "".join(pieces)


def _insert_column(ws, at, hdr, title, style_from):
    """在第 at 欄前面插一欄，模仿 Excel 插欄：右邊的值搬過去、所有公式的參照跟著位移、
    跨過插入點的範圍自動變寬、合併儲存格與欄寬跟著移。標題與樣式抄隔壁那個日期欄。"""
    from copy import copy
    from openpyxl.utils import get_column_letter
    max_col = ws.max_column
    if style_from is not None and style_from >= at:
        style_from += 1            # 樣板欄在插入點右邊，插完會往右移一格
    ws.insert_cols(at)
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith("="):
                c.value = _shift_formula(c.value, at)
    merged = [str(r) for r in ws.merged_cells.ranges]
    for r in merged:
        ws.unmerge_cells(r)
    from openpyxl.utils.cell import range_boundaries
    for r in merged:
        c1, r1, c2, r2 = range_boundaries(r)
        if c1 >= at:
            c1 += 1; c2 += 1
        elif c2 >= at:
            c2 += 1
        ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
    dims = ws.column_dimensions
    for col in range(max_col, at - 1, -1):
        src = dims.get(get_column_letter(col))
        if src is not None and src.width:
            dims[get_column_letter(col + 1)].width = src.width
    if style_from:
        src_w = dims.get(get_column_letter(style_from))
        if src_w is not None and src_w.width:
            dims[get_column_letter(at)].width = src_w.width
        for r in range(1, ws.max_row + 1):
            ws.cell(row=r, column=at)._style = copy(ws.cell(row=r, column=style_from)._style)
    ws.cell(row=hdr, column=at).value = title


def _sheet_layout(ws, hdr, mm):
    """讀底稿標題列，找出這個月用得到的欄：Barcode／skuid 欄、每個日期欄（day → col）、
    「M/交貨」空欄、這個月的 TTL 加總欄。插欄之後要重新呼叫一次，因為欄位都位移了。"""
    headers = {c: norm_text(ws.cell(row=hdr, column=c).value) for c in range(1, ws.max_column + 1)}
    lowered = {c: h.lower().replace(" ", "") for c, h in headers.items()}
    def col_of(field):
        for a in _HEADER_ALIASES[field]:
            key = a.lower().replace(" ", "")
            for c, h in lowered.items():
                if h == key:
                    return c
        return None
    date_cols, spare = {}, []
    for c, h in headers.items():
        m = _DATE_HDR.match(h)
        if m and int(m.group(1)) == mm:
            if m.group(2):
                date_cols[int(m.group(2))] = c
            else:
                spare.append(c)
    # TTL 加總欄用「位置」認：這個月日期欄（含空欄）最右邊那欄之後、第一個「N月TTL」標題就是它，
    # 不管標題寫幾月。Chloe 的總表 9 月區塊的 TTL 標題複製貼上寫成「12月TTL」，照標題找會找不到，
    # 新插的日期欄就不會被加進 SUM（2026-09-18 抓到）。真的找不到才退回照標題月份找。
    ttl_re = re.compile(r"^\s*\d{1,2}月TTL")
    block_cols = list(date_cols.values()) + spare
    ttl_col = None
    if block_cols:
        for c in range(max(block_cols) + 1, ws.max_column + 1):
            h = headers[c]
            if ttl_re.match(h):
                ttl_col = c; break
            if _DATE_HDR.match(h):      # 撞到別的月份的日期欄還沒看到 TTL，這個月沒有 TTL 欄
                break
    if ttl_col is None:
        ttl_col = next((c for c, h in headers.items() if re.match(rf"^\s*{mm}月TTL", h)), None)
    return {"headers": headers, "bc_col": col_of("barcode"), "sku_col": col_of("sku_id"),
            "date_cols": date_cols, "spare": spare, "ttl_col": ttl_col,
            "month_cols": [c for c, h in headers.items() if _DATE_HDR.match(h) or ttl_re.match(h)]}


def _ensure_date_columns(ws, hdr, mm, dates):
    """每個交貨日都要有欄：有同日期的欄就用；沒有就照 Chloe 說的自動插一欄，插在日期順序該在的
    位置（右邊公式跟著位移，見 _insert_column）；「M/交貨」空欄是業務的，不動。
    底稿完全沒有這個月的區塊時，接在最後一個月份區塊後面補「日期欄＋TTL 加總欄」。回傳 (layout, 新增的日期)。"""
    lay = _sheet_layout(ws, hdr, mm); added = []
    for d in dates:
        day = int(d[8:10])
        if day in lay["date_cols"]:
            continue
        later = sorted(c for dd, c in lay["date_cols"].items() if dd > day)
        if later:
            at = later[0]                                    # 插在下一個日期前面
        elif lay["spare"]:
            at = lay["spare"][0]                             # 最後一個日期之後、空欄之前
        elif lay["ttl_col"]:
            at = lay["ttl_col"]                              # 空欄也沒有 → 加總欄前面
        elif lay["date_cols"]:
            at = max(lay["date_cols"].values()) + 1
        else:
            at = (max(lay["month_cols"]) + 1) if lay["month_cols"] else ws.max_column + 1
            _insert_column(ws, at, hdr, f"{mm}月TTL下單總箱數", max(lay["month_cols"]) if lay["month_cols"] else None)
            lay = _sheet_layout(ws, hdr, mm)
        neighbor = min(lay["date_cols"].values(), key=lambda c: abs(c - at)) if lay["date_cols"] else None
        _insert_column(ws, at, hdr, f"{mm}/{day}交貨", neighbor)
        added.append(f"{mm}/{day}")
        lay = _sheet_layout(ws, hdr, mm)
    return lay, added


def _row_index(ws, hdr, bc_col, sku_col):
    """底稿每一列的 Barcode／skuid → 列號（第一次出現的為準）。"""
    by_bc, by_sku = {}, {}
    for r in range(hdr + 1, ws.max_row + 1):
        bc = norm_key(ws.cell(row=r, column=bc_col).value) if bc_col else ""
        sku = norm_key(ws.cell(row=r, column=sku_col).value) if sku_col else ""
        if bc and bc not in by_bc:
            by_bc[bc] = r
        if sku and sku not in by_sku:
            by_sku[sku] = r
    return by_bc, by_sku


def _rewrite_month_total(ws, hdr, lay, block):
    """這個月的 TTL 加總欄一律重寫成涵蓋整個區塊（含新插的欄）；插在區塊尾端時舊公式會漏掉新欄。"""
    if not (lay["ttl_col"] and block):
        return
    from openpyxl.utils import get_column_letter as _L
    lo, hi = _L(min(block)), _L(max(block))
    for r in range(hdr + 1, ws.max_row + 1):
        v = ws.cell(row=r, column=lay["ttl_col"]).value
        if (isinstance(v, str) and v.upper().startswith("=SUM(")) or (v is None and ws.cell(row=r, column=lay["bc_col"] or 1).value):
            ws.cell(row=r, column=lay["ttl_col"]).value = f"=SUM({lo}{r}:{hi}{r})"


def _write_cases(ws, lay, block, by_bc, by_sku, summary):
    """把每個商品各交貨日的箱數填進對應格；對得到的列先把這個月的欄清空再填（沒出貨的日子留白）。
    對不到的商品回傳出來，不混進總表。"""
    matched, unmatched, filled = 0, [], 0
    for r0 in summary["rows"]:
        r = by_bc.get(r0["barcode"]) or by_sku.get(r0["sku_id"])
        if r is None:
            unmatched.append(r0); continue
        matched += 1
        for c in block:
            ws.cell(row=r, column=c).value = None
        for d, v in r0["by_date"].items():
            c = lay["date_cols"].get(int(d[8:10]))
            if c is not None:
                ws.cell(row=r, column=c).value = v; filled += 1
    return matched, unmatched, filled


def _write_report_sheet(wb, tpl, summary, month, operator, matched, filled, added, unmatched):
    """最後一個分頁「系統填入說明」：這次填了什麼、自動新增了哪些日期欄、哪些商品總表沒有。"""
    from openpyxl.styles import Font
    info = wb.create_sheet("系統填入說明")
    bold = Font(bold=True)
    info.append(["商品主檔自動化 填入紀錄"]); info["A1"].font = bold
    info.append(["填入時間", now(), "操作者", operator])
    info.append(["月份", month, "線別", summary["line"]])
    info.append(["底稿", tpl["filename"], "工作表", tpl["sheet"]])
    info.append(["對到的商品", matched, "填入格數", filled])
    if added:
        info.append(["自動新增的日期欄", "、".join(added)])
    info.append([])
    if unmatched:
        info.append(["酷澎有下單、但總表裡沒有的商品（沒混進總表，列在這裡給你看）："]); info.cell(row=info.max_row, column=1).font = bold
        info.append(["國條", "SKU ID", "品名", "品牌", "箱入數"] + [_md(d) for d in summary["dates"]] + ["月加總"])
        for c in info[info.max_row]:
            c.font = bold
        for r0 in unmatched:
            info.append([r0["barcode"], r0["sku_id"], r0["product_name"], r0["brand"], r0["box_size"]]
                        + [r0["by_date"].get(d) for d in summary["dates"]] + [r0["month_total"]])
    info.column_dimensions["A"].width = 22; info.column_dimensions["B"].width = 18; info.column_dimensions["C"].width = 40


def _fill_template(tpl, summary, month, operator):
    """把這個月每個交貨日的箱數填進底稿：只動這個月的日期欄和對得到的商品列，其他列、其他月份、
    業務的公式、順序全部不碰。步驟：確保日期欄都在（缺的插欄）→ 對商品列 → 重寫 TTL → 填箱數 → 說明分頁。"""
    import base64
    wb = openpyxl.load_workbook(io.BytesIO(base64.b64decode(tpl["content_b64"])))   # 保留公式
    ws = wb[tpl["sheet"]]
    hdr = tpl["header_row"]; mm = int(month[5:7])
    lay, added = _ensure_date_columns(ws, hdr, mm, summary["dates"])
    by_bc, by_sku = _row_index(ws, hdr, lay["bc_col"], lay["sku_col"])
    block = sorted(lay["date_cols"].values()) + lay["spare"]
    _rewrite_month_total(ws, hdr, lay, block)
    matched, unmatched, filled = _write_cases(ws, lay, block, by_bc, by_sku, summary)
    _write_report_sheet(wb, tpl, summary, month, operator, matched, filled, added, unmatched)
    return wb, {"matched": matched, "filled": filled, "unmatched": len(unmatched),
                "missing_dates": [], "renamed": added}


@master_bp.route("/api/master/export")
def api_export():
    """總表格式：一列一個國條，往右是各交貨日箱數、月加總。"""
    group = norm_text(request.args.get("line")); month = norm_text(request.args.get("month")) or _this_month()
    if not group or not _valid_month(month):
        return jsonify({"error": "請選線別與月份。"}), 400
    cfg = _line_groups()
    conn = get_conn()
    try:
        s = _build_summary(conn, group, month, cfg)
        orders = [o for o in _month_orders(conn, month, cfg) if o["line_group"] == group]
        tpl = _template_row(conn, group)
    finally:
        conn.close()
    if tpl is not None:
        # 有底稿就照底稿填（寶僑要跟他們的總表一模一樣，含商品順序），檔名沿用底稿的
        wb, _rep = _fill_template(tpl, s, month, _operator())
        # 檔名不能跟底稿一樣：Chloe 開著底稿再開匯出檔，Excel 會說「無法同時開啟兩個相同名稱的活頁簿」
        # （2026-09-18）。改成「9月_總表.xlsx」。
        return _xlsx_response(wb, f"{int(month[5:7])}月_總表.xlsx")
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "總表"
    mm = int(month[5:7])
    head = ["國條", "SKU ID", "永豐料號", "品類", "品牌", "品名", "箱入數"] + [_md(d) for d in s["dates"]] + \
           [f"{mm}月TTL下單總箱數", "備註"]
    ws.append(head)
    for r in s["rows"]:
        ws.append([r["barcode"], r["sku_id"], r["yf_sku"], r["category"], r["brand"], r["product_name"], r["box_size"]]
                  + [r["by_date"].get(d) for d in s["dates"]] + [r["month_total"], r["note"]])
    ws.append([])
    ws.append(["合計", "", "", "", "", "", ""] + [s["totals_by_date"].get(d) for d in s["dates"]] + [s["month_total"], ""])
    bold = Font(bold=True); fill = PatternFill("solid", fgColor="DBEAFE")
    for c in ws[1]:
        c.font = bold; c.fill = fill; c.alignment = Alignment(horizontal="center", wrap_text=True)
    for c in ws[ws.max_row]:
        c.font = bold
    ws.freeze_panes = "H2"
    for i in range(1, len(head) + 1):
        ws.column_dimensions[get_column_letter(i)].width = 40 if i == 6 else 14
    for row in ws.iter_rows(min_row=2, min_col=1, max_col=3):
        for c in row:
            c.number_format = "@"
    ws2 = wb.create_sheet("每日合計"); ws2.append(["交貨日", "出貨總箱數", "PO 數", "品項數"])
    per_day = collections.OrderedDict()
    for o in orders:
        if not o["delivery_date"]:
            continue
        d = per_day.setdefault(o["delivery_date"], {"cases": 0.0, "pos": set(), "rows": 0})
        d["cases"] += o["cases"] or 0; d["pos"].add(o["po_number"]); d["rows"] += 1
    for dte, d in per_day.items():
        ws2.append([dte, round(d["cases"], 2), len(d["pos"]), d["rows"]])
    for c in ws2[1]:
        c.font = bold; c.fill = fill
    return _xlsx_response(wb, f"{int(month[5:7])}月_總表.xlsx")


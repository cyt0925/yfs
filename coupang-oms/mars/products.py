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


def _code_canon(conn, conflicts=None):
    """任何出現過的瑪氏貨號（含備註裡的舊貨號）→ 商品總表現在用的貨號。

    同一個號碼對到兩個料號時（Alice 2026-10-16 價格本實際有 4 組）：備註有「停產」的讓給還在賣的（3 組）；
    都還在賣的（10214470：M10214470 自己的貨號、也是 M10290449 備註裡換掉的舊號）記到 conflicts，比對勇信表時
    看那張採購單訂的是哪個料號（resolve_code）；都沒訂就先照「把這個號碼換掉的那個」（M10290449，較新；Jerry 2026-10-16）。
    conflicts = {號碼: {"targets": {目前貨號: [料號, ...]}, "default": 目前貨號}}。
    同一個貨號分兩個料號（M10281733、M10281733-1 有包單位）對到的是同一個貨號，不算衝突。"""
    cands = {}      # 號碼 → {目前貨號: {"dead": 都停產?, "skus": [...], "former": 這個號碼是它換掉的舊號?}}
    for r in _rows(conn.execute("SELECT DISTINCT yf_sku, mars_code, note FROM mst_mars_products WHERE mars_code != '' ORDER BY yf_sku")):
        dead = "停產" in (r["note"] or "")
        for code in [r["mars_code"]] + re.findall(r"\d{8}", r["note"] or ""):
            t = cands.setdefault(code, {}).setdefault(r["mars_code"], {"dead": True, "skus": [], "former": code != r["mars_code"]})
            t["dead"] = t["dead"] and dead
            if r["yf_sku"] not in t["skus"]:
                t["skus"].append(r["yf_sku"])
    canon = {}
    for code, targets in cands.items():
        alive = {k: v for k, v in targets.items() if not v["dead"]} or targets
        best = sorted(alive, key=lambda k: (not alive[k]["former"], alive[k]["skus"][0]))[0]
        canon[code] = best
        if len(alive) > 1 and conflicts is not None:
            conflicts[code] = {"targets": {k: list(v["skus"]) for k, v in alive.items()}, "default": best}
    return canon


def sku_codes(conn):
    """永豐料號 → 商品總表現在的瑪氏貨號。"""
    return {r["yf_sku"]: r["mars_code"] for r in _rows(conn.execute("SELECT DISTINCT yf_sku, mars_code FROM mst_mars_products WHERE mars_code != ''"))}


def item_code(yf_sku, snap_code, skus, canon, conflicts):
    """拆單／EIP 品項的貨號身分：照料號查商品總表現在的貨號（拆單時存的可能是舊號）；
    那個號碼本身有衝突就直接用它——品項的料號已經很明確，不用猜。"""
    code = skus.get(yf_sku or "") or (snap_code or "")
    return code if code in conflicts else canon.get(code, code)


def resolve_code(code, conflicts, canon, present):
    """勇信表上的號碼 → 這張採購單上要算給哪個貨號。回 (貨號, 要不要人工確認)。
    有衝突的號碼看 present（這張單品項的貨號）：只有一個在單上就是它；兩個都在才要人確認（先給預設）；都不在用預設。"""
    if code not in conflicts:
        return canon.get(code, code), False
    hit = [t for t in conflicts[code]["targets"] if t in present]
    if len(hit) == 1:
        return hit[0], False
    return conflicts[code]["default"], len(hit) > 1


def conflict_skus(cf):
    return sorted(s for v in cf["targets"].values() for s in v)


def conflict_warnings(codes, conflicts):
    """上傳商品總表時列出：這些號碼對到兩個料號，比對時怎麼處理。"""
    out = []
    for code in sorted(c for c in codes if c in conflicts):
        cf = conflicts[code]
        out.append(f"瑪氏貨號 {code} 對到 {len(conflict_skus(cf))} 個料號（{'、'.join(conflict_skus(cf))}）：比對勇信表時看採購單訂的是哪個；都沒訂的話先照 {cf['targets'][cf['default']][0]}")
    return out


def merge_conflicts(eip, d, conflicts, canon, present, warnings, extra_keys=()):
    """一張採購單：PDF 上有衝突的號碼算給哪個貨號，把 qty／names（和 extra_keys 指的其他字典）併過去；兩個都有訂的記提醒。"""
    for code in [k for k in d["qty"] if k in conflicts]:
        target, ask = resolve_code(code, conflicts, canon, present)
        if ask:
            warnings.append(f"{eip}：勇信表上的瑪氏貨號 {code} 對到 {'、'.join(conflict_skus(conflicts[code]))}，這張單兩個都有訂，"
                            f"先算給 {conflicts[code]['targets'][target][0]}，請人工確認")
        if target != code:
            d["qty"][target] += d["qty"].pop(code)
            d["names"].setdefault(target, d["names"].pop(code, ""))
            for k in extra_keys:
                d[k][target] = d[k].get(target, []) + d[k].pop(code, [])


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
        conflicts = {}
        _code_canon(conn, conflicts)
        code_conflicts = conflict_warnings(conflicts, conflicts)
    finally:
        conn.close()
    by_cat = collections.Counter(cat_code(r["category"]) or "?" for r in rows if r["unit"] == "箱")
    return jsonify({"ok": True, "filename": up.filename, "rows": len(rows), "codes": len(codes), "code_conflicts": code_conflicts,
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


_NOTE_CHAIN = re.compile(r"^(?:\d{8}|[>›]|轉貨號|轉|停產|\s)+")     # 備註開頭的換號鏈：10247136 轉貨號 ／ 10214470>10290449>60037559 ／ 10230369 轉>停產


@mars_bp.route("/api/mars/products/list")
def api_list_mars_products():
    """商品總表檢視（Jerry 2026-10-16 示意圖定案）：一個料號一列（箱盒包合起來），看得到現在的貨號、換過的舊號、停產、
    對到兩個料號、蝦皮特選補登；只看不能改（要改去 Alice 的價格本；Excel 沒有的用蝦皮特選「補商品資料」）。
    filter：all／changed（有換號）／dead（停產）／conflict（對到兩個料號）／manual（補登）；q 空白隔開多個關鍵字，舊號也查得到。"""
    words = [w.lower() for w in str(request.args.get("q") or "").split() if w][:5]
    flt = request.args.get("filter") or "all"
    try:
        page = max(1, int(request.args.get("page") or 1)); per = min(100, max(5, int(request.args.get("per") or 20)))
    except ValueError:
        page, per = 1, 20
    conn = get_conn()
    try:
        conflicts = {}
        _code_canon(conn, conflicts)
        rows = _rows(conn.execute("SELECT yf_sku, unit, mars_code, name, shelf_days, note, manual_by, manual_at, item_code, item_code2, category, brand "
                                  "FROM mst_mars_products ORDER BY yf_sku, unit"))
    finally:
        conn.close()
    by = collections.OrderedDict()
    for r in rows:
        by.setdefault(r["yf_sku"], []).append(r)
    cur_of, claims = {}, {}        # 現在貨號 → [料號]；備註裡的舊號 → [料號]（誰說這個號碼是它換掉的）
    for yf, rs in by.items():
        if rs[0]["mars_code"]:
            cur_of.setdefault(rs[0]["mars_code"], []).append(yf)
        for c in dict.fromkeys(re.findall(r"\d{8}", rs[0]["note"] or "")):      # 同一個號碼在備註出現兩次只算一次
            if c != rs[0]["mars_code"]:
                claims.setdefault(c, []).append(yf)
    items = []
    for yf, rs in by.items():
        r0 = rs[0]; note = r0["note"] or ""; code = r0["mars_code"] or ""
        olds = list(dict.fromkeys(c for c in re.findall(r"\d{8}", note) if c != code))
        cfs = []
        for c in [code] + olds:
            cf = conflicts.get(c)
            if cf and c not in [x["code"] for x in cfs]:
                cfs.append({"code": c, "others": [y for y in conflict_skus(cf) if y != yf], "default_sku": cf["targets"][cf["default"]][0],
                            "is_default": yf in cf["targets"][cf["default"]]})
        items.append({"yf_sku": yf, "name": r0["name"], "mars_code": code, "old_codes": olds, "note_text": _NOTE_CHAIN.sub("", note).strip(),   # 鏈之後的字照原樣留（外箱條碼有更新、11/7 換(…)）
                      "shelf_days": r0["shelf_days"], "units": "".join(sorted({r["unit"] for r in rs}, key=lambda u: "箱盒包".find(u))),
                      "dead": "停產" in note, "manual_by": r0["manual_by"] or "", "manual_at": r0["manual_at"] or "",
                      "from": sorted({y for c in olds for y in cur_of.get(c, []) if y != yf}),          # 舊號是別人現在的貨號 → 從它轉來
                      "to": sorted(y for y in claims.get(code, []) if y != yf) if code else [],          # 自己的貨號被別人寫成舊號 → 已轉給它
                      "same_code": sorted(y for y in cur_of.get(code, []) if y != yf) if code else [], "conflicts": cfs,
                      "_text": " ".join(str(v or "") for r in rs for v in (r["yf_sku"], r["mars_code"], r["name"], r["note"], r["item_code"], r["item_code2"], r["category"], r["brand"])).lower()})
    counts = {"all": len(items), "changed": sum(1 for i in items if i["old_codes"]), "dead": sum(1 for i in items if i["dead"]),
              "conflict": len(conflicts), "manual": sum(1 for i in items if i["manual_by"])}
    keep = {"all": lambda i: True, "changed": lambda i: i["old_codes"], "dead": lambda i: i["dead"], "conflict": lambda i: i["conflicts"],
            "manual": lambda i: i["manual_by"]}.get(flt, lambda i: True)
    sel = [i for i in items if keep(i) and all(w in i["_text"] for w in words)]
    pages = max(1, -(-len(sel) // per)); page = min(page, pages)
    out = [{k: v for k, v in i.items() if k != "_text"} for i in sel[(page - 1) * per: page * per]]
    return jsonify({"rows": out, "total": len(sel), "page": page, "pages": pages, "per": per, "counts": counts})


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


__all__ = ["parse_master", "load_products", "api_import_mars_products", "api_search_mars_products", "api_list_mars_products", "api_mars_status", "mars_page"]

"""MOMO 第一批：換成箱數 → EIP 採購單；回填 EIP 單號；瑪氏採購單、直送採購單（需求第四、六、七章）。

- 跟原廠下採一律用「箱」（需求四.2、七）：數量÷箱入數。箱入數：紙品、潔品、寶僑查「價格本」（料號＋單位代碼，shp_conv），
  瑪氏查「瑪氏商品總表」（永豐料號＋單位，mst_mars_products）；兩份都跟蝦皮特選、瑪氏出貨共用，更新一次全部換新。
- 除不盡（瑪氏口香糖一箱 30 包、MOMO 開 96 包）：瑪氏一律捨去（96 包→3 箱＝90 包；Becky 8435 的採購單就是這樣，Jerry 2026-10-29「平常是怎樣那就是怎樣」），
  少的那些到 ④ 入庫通知單畫掉寫實際數；不到一箱的擋下來要人決定。紙品、潔品還沒看過實際做法，先照蝦皮特選擋下來要人改數量。
  改成 0＝這品不下採（贈品、價格本查不到又不用下採的），不用箱入數。這是 2026-10-28 先用的做法，問題清單第 2 題 Alice 回了再改。
- 分檔：瑪氏依倉、到貨日、來源單位（箱／盒包）、品類、中標分（需求一、七）；紙品、潔品、寶僑一天一倉一份（還沒問，先跟蝦皮特選直送一樣）。
  同一個料號同一個單位合成一列。
- 竹運出貨的入庫單不用下採（第二階段接蝦皮那套竹運判斷）。
"""
import datetime as _dt
import os
import zipfile

import xlrd
from xlutils.copy import copy as xl_copy

from mars.common import CAT_FULL, cat_code
from mars.po import load_settings as mars_settings
from mars.po import po_file
from shopee.direct_po import direct_xlsx

from .common import *  # noqa: F401,F403

TPL_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "purchase_templates")
TEMPLATES = {MARS: ("mars_template.xls", 60), "寶僑": ("pg_template.xls", 249), "紙品": ("paper_template.xls", 45), "潔品": ("paper_template.xls", 45)}
EIP_PO_RE = re.compile(r"^PO\d{9}$")
F5_FIXED = "momo嘜頭+入倉單"         # 需求七 F5
SHELF_REQ = "1/2效期以上"            # 需求七 F6
LEAD_DAYS = 2                        # 需求七 C3：交付日期往前推 2 個工作日


# ── 查價格本 ─────────────────────────────────────────────────────────────────
def _conv_map(conn, codes):
    codes = sorted(c for c in codes if c)
    if not codes:
        return {}
    marks = ",".join("?" * len(codes))
    return {r["code3"]: r for r in _rows(conn.execute(f"SELECT * FROM shp_conv WHERE code3 IN ({marks})", codes))}


def _mars_map(conn, items):
    codes = sorted({i["base_code"] for i in items if i["line"] == MARS and i["base_code"]})
    if not codes:
        return {}
    marks = ",".join("?" * len(codes))
    return {(r["yf_sku"], r["unit"]): r for r in _rows(conn.execute(f"SELECT * FROM mst_mars_products WHERE yf_sku IN ({marks})", codes))}


def conv_for_item(conn, it):
    if it["line"] == MARS:
        return _row(conn.execute("SELECT * FROM mst_mars_products WHERE yf_sku = ? AND unit = ?", (it["base_code"], it["unit_name"])))
    return _row(conn.execute("SELECT * FROM shp_conv WHERE code3 = ?", ((it["base_code"] or "") + (it["unit_code"] or ""),)))


def box_of(conv):
    try:
        v = float(conv["box_qty"]) if conv and conv["box_qty"] not in (None, "") else 0
    except (TypeError, ValueError):
        return None
    return int(v) if v > 0 and v.is_integer() else (v if v > 0 else None)


def _adj_map(conn, ids):
    if not ids:
        return {}
    marks = ",".join("?" * len(ids))
    return {r["order_id"]: r for r in _rows(conn.execute(f"SELECT * FROM momo_qty_adj WHERE order_id IN ({marks})", list(ids)))}


def eip_map(conn, ids):
    """品項 → 掛在哪一張 EIP（最新那張）。"""
    if not ids:
        return {}
    marks = ",".join("?" * len(ids))
    out = {}
    for r in _rows(conn.execute(f"SELECT i.order_id, e.id, e.eip_po, e.filename FROM momo_eip_items i JOIN momo_eip e ON e.id = i.eip_id "
                                f"WHERE i.order_id IN ({marks}) ORDER BY e.id", list(ids))):
        out[r["order_id"]] = {"eip_id": r["id"], "eip_po": r["eip_po"] or "", "filename": r["filename"]}
    return out


def skipped_ids(conn, items):
    """改成 0 的品項（不下採）：不算待產 EIP。"""
    adj = _adj_map(conn, [i["id"] for i in items])
    return {i["id"] for i in items if i["id"] in adj and adj[i["id"]]["qty_after"] == 0 and adj[i["id"]]["qty_before"] == i["qty"]}


def _name(it):
    spec = (it.get("spec") or "").strip()
    return (it.get("product_name") or "") + (f" {spec}" if spec and spec != "無" else "")


def item_plan(it, conv, adj, eip, other_units=()):
    line = it["line"]
    qty = it["qty"] or 0
    use = adj["qty_after"] if adj and adj["qty_before"] == qty else qty      # MOMO 數量又變了，之前改的就不算
    code3 = (it["base_code"] or "") + (it["unit_code"] or "")
    box = box_of(conv)
    p = {"order_id": it["id"], "factory_code": it["factory_code"], "code": it["base_code"], "code3": code3, "name": _name(it),
         "unit": it["unit_name"], "qty": qty, "use_qty": use, "conv": box, "adj": adj if adj and adj["qty_before"] == qty else None,
         "problems": [], "blocking": False, "eip": eip, "line": line, "unit_group": "", "out_qty": None, "skip": use == 0}
    if it["diff_pending"]:
        p["problems"].append("MOMO 改過這張單，請先點這張單確認差異")
        p["blocking"] = True
    if eip and eip["eip_po"]:
        p["problems"].append(f"已回填 EIP 採購單號 {eip['eip_po']}；要重新產出請先清除單號")
        p["blocking"] = True
    if p["skip"]:
        return p
    if not line:
        p["problems"].append(f"認不出是紙品、潔品還是瑪氏（料號 {it['base_code']}）")
        p["blocking"] = True
        return p
    if conv is None:
        p["problems"].append((f"瑪氏商品總表有 {it['base_code']} 的{'、'.join(other_units)}，沒有{p['unit']}" if other_units
                              else f"瑪氏商品總表找不到 {it['base_code']}（{p['unit']}）") if line == MARS
                             else f"{'寶僑' if line == '寶僑' else '紙潔'}價格本找不到 {code3}")
        p["problems"].append("贈品或不用下採的，請按「改數量」改成 0")
        p["blocking"] = True
        return p
    if not box:
        p["problems"].append(f"{'瑪氏商品總表' if line == MARS else '價格本'} {code3} 沒有箱入數")
        p["blocking"] = True
        return p
    rem = use % box
    if rem and line == MARS and not p["adj"] and use - rem > 0:   # 瑪氏：除不盡一律捨去（Becky 的做法）；人改過數量就照人改的
        use -= rem
        p["use_qty"], p["rounded"] = use, True
        p["note"] = f"{qty:,} {p['unit']} ÷ 箱入數 {box} 除不盡，捨去成 {use // box} 箱（{use:,} {p['unit']}），少的在入庫通知單畫掉"
        rem = 0
    if rem:
        p["problems"].append(f"{use:,} {p['unit']} ÷ 箱入數 {box} 除不盡（{use / box:.1f} 箱），請修改數量" + ("；不到一箱，改成 0 就是不下採" if use < box else ""))
        p["blocking"] = True
    p["out_qty"] = use // box if not rem else round(use / box, 2)
    if line == MARS:
        cat = cat_code(conv["category"])
        lab = "V" if (conv["inner_label"] or "").strip().upper() == "V" else ""
        uclass = "箱" if p["unit"] == "箱" else "盒包"           # 來源單位類別是拆單條件，不是採購單位（需求四.1）
        p["unit_group"] = f"{uclass}_{cat or '品類不明'}_{'需貼中標' if lab else '不貼中標'}"
        p["mars"] = {"category": cat, "label": lab, "unit_class": uclass, "mars_code": conv["mars_code"] or "", "mars_name": conv["name"] or "",
                     "price": conv["price"], "pcs_per_case": conv["pcs_per_case"], "inner_per_case": conv["inner_per_case"],
                     "note": conv["note"] or "", "po_case_note": conv["po_case_note"] or ""}
    return p


def build_plan(conn, nos=None, args=None):
    """入庫單 → 每個品項的下採計畫。只看原廠直出、沒被移除的。"""
    from .orders import _filter
    where, params = _filter(args or {})
    where += " AND removed = 0 AND ship_method = ?"
    params = params + [DIRECT]
    items = _rows(conn.execute(f"SELECT * FROM momo_orders WHERE {where} ORDER BY expected_date, warehouse, inbound_no, id", params))
    if nos is not None:
        want = set(nos)
        items = [i for i in items if i["inbound_no"] in want]
    ids = [i["id"] for i in items]
    convs = _conv_map(conn, {(i["base_code"] or "") + (i["unit_code"] or "") for i in items if i["line"] != MARS})
    mars = _mars_map(conn, items)
    adjs, eips = _adj_map(conn, ids), eip_map(conn, ids)
    groups = collections.OrderedDict()
    for it in items:
        g = groups.setdefault(it["inbound_no"], {"inbound_no": it["inbound_no"], "warehouse": it["warehouse"], "expected_date": it["expected_date"],
                                                 "items": []})
        conv = mars.get((it["base_code"], it["unit_name"])) if it["line"] == MARS else convs.get((it["base_code"] or "") + (it["unit_code"] or ""))
        others = sorted(u for (c, u) in mars if c == it["base_code"] and u != it["unit_name"]) if it["line"] == MARS and conv is None else ()
        g["items"].append(item_plan(it, conv, adjs.get(it["id"]), eips.get(it["id"]), others))
    out = []
    for g in groups.values():
        its = g["items"]
        todo = [p for p in its if not p["skip"]]
        g.update(blocking=sum(1 for p in its if p["blocking"]), lines=sorted({p["line"] or "認不出" for p in its}),
                 qty=sum(p["qty"] for p in its), boxes=None if any(p["out_qty"] is None for p in todo) else round(sum(p["out_qty"] for p in todo), 2),
                 po_status="已回填" if todo and all(p["eip"] and p["eip"]["eip_po"] for p in todo) else "已產出" if any(p["eip"] for p in todo)
                 else "未產出" if todo else "不需產出")
        out.append(g)
    return out


@momo_bp.route("/api/momo/purchase/plan")
def api_purchase_plan():
    conn = get_conn()
    try:
        orders = build_plan(conn, args=request.args)
    finally:
        conn.close()
    return jsonify({"orders": orders})


@momo_bp.route("/api/momo/purchase/adjust", methods=["POST"])
def api_purchase_adjust():
    """改下採數量：要整除箱入數、要寫原因；改成 0＝不下採（不用箱入數）；清空＝取消修改。"""
    p = request.get_json(silent=True) or {}
    try:
        oid = int(p.get("order_id"))
    except (TypeError, ValueError):
        return jsonify({"error": "缺少品項。"}), 400
    reason, raw = str(p.get("reason") or "").strip(), p.get("qty")
    operator = _operator()
    conn = get_conn()
    try:
        it = _row(conn.execute("SELECT * FROM momo_orders WHERE id = ?", (oid,)))
        if it is None:
            return jsonify({"error": "找不到這個品項，請重新整理頁面。"}), 404
        old = _row(conn.execute("SELECT * FROM momo_qty_adj WHERE order_id = ?", (oid,)))
        if raw in (None, ""):
            if old:
                conn.execute("DELETE FROM momo_qty_adj WHERE order_id = ?", (oid,))
                _log(conn, it["line"], it["inbound_no"], it["product_no"], "", "momo_qty_adj", "下採數量", old["qty_after"],
                     f"取消修改（回到 {it['qty']}）", operator, "manual", it["factory_code"])
                conn.commit()
            return jsonify({"ok": True, "cleared": True})
        try:
            qty = int(str(raw).replace(",", ""))
        except ValueError:
            return jsonify({"error": "修改後數量請填整數。"}), 400
        if qty < 0:
            return jsonify({"error": "數量不能是負數。"}), 400
        box = box_of(conv_for_item(conn, it))
        if qty:
            if not box:
                return jsonify({"error": "查不到這個品項的箱入數，只能改成 0（不下採）；要下採請先更新價格本或瑪氏商品總表。"}), 400
            if qty % box:
                return jsonify({"error": f"修改後數量 {qty} 仍除不盡箱入數 {box}，請改成 {box} 的倍數（例如 {it['qty'] // box * box} 或 {(it['qty'] // box + 1) * box}）。"}), 400
        if not reason:
            return jsonify({"error": "請填寫修改原因。"}), 400
        stamp = now()
        if old:
            conn.execute("UPDATE momo_qty_adj SET qty_before = ?, qty_after = ?, conv = ?, reason = ?, operator = ?, updated_at = ? WHERE order_id = ?",
                         (it["qty"], qty, box, reason, operator, stamp, oid))
        else:
            conn.execute("INSERT INTO momo_qty_adj (order_id, qty_before, qty_after, conv, reason, operator, updated_at) VALUES (?,?,?,?,?,?,?)",
                         (oid, it["qty"], qty, box, reason, operator, stamp))
        _log(conn, it["line"], it["inbound_no"], it["product_no"], "", "momo_qty_adj", "下採數量", it["qty"], qty if qty else "0（不下採）", operator,
             "manual", f"{it['factory_code']}；箱入數 {box or '—'}；{reason}", reason)
        conn.commit()
        return jsonify({"ok": True, "qty": qty})
    finally:
        conn.close()


# ── 產 EIP ──────────────────────────────────────────────────────────────────
def _md(d):
    return f"{d[5:7]}{d[8:10]}" if d and len(d) >= 10 else "日期不明"


def _merge_rows(plans):
    rows = collections.OrderedDict()
    for p in plans:
        k = (p["code"], p["unit"])
        r = rows.get(k)
        if r is None:
            r = rows[k] = {"code": p["code"], "name": p["name"], "unit": "箱", "qty": 0, "in_unit": p["unit"], "in_qty": 0, "conv": p["conv"],
                           "remark": p["remark"], "parts": []}
            if p.get("mars"):
                r.update(p["mars"], cases=0)
        r["qty"] += p["out_qty"]
        r["in_qty"] += p["use_qty"]
        r["parts"].append({"order_id": p["order_id"], "use_qty": p["use_qty"]})
        if "cases" in r:
            r["cases"] += p["out_qty"]
    return list(rows.values())


def eip_xls(rows, line):
    """套公司 EIP 範本（產品領用單）：紙品、潔品用紙潔那份（MOMO 範例就是這份），瑪氏用瑪氏那份，寶僑用寶僑那份。"""
    name, wipe = TEMPLATES.get(line, TEMPLATES["紙品"])
    rb = xlrd.open_workbook(os.path.join(TPL_DIR, name), formatting_info=True)
    wb = xl_copy(rb)
    ws = wb.get_sheet(0)
    for r in range(1, max(wipe, len(rows) + 1)):
        for c in range(10):
            ws.write(r, c, "")
    for i, row in enumerate(rows, start=1):
        ws.write(i, 0, i)
        ws.write(i, 1, str(row["code"]))
        ws.write(i, 2, row["name"])
        ws.write(i, 3, row["unit"])
        ws.write(i, 6, row["qty"])
        ws.write(i, 8, row["remark"])
    from purchase import _write_qty_header_formula       # 標題「採購數量398」帶總數（紙潔範本本來就有，瑪氏也加；Jerry 2026-10-28 蝦皮要跟酷澎一樣）
    _write_qty_header_formula(wb, ws, rows)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def eip_filename(line, date, wh, unit_group):
    return f"產品採購表上傳-MOMO{line}_{_md(date)}到{wh}" + (f"_{unit_group}" if unit_group else "") + ".xls"


def _remark(line, date, wh):
    if line == MARS:
        return f"MARS入倉{(date or '').replace('-', '')}台灣瑪氏送momo-{wh}"
    return f"momo-{wh} {int(date[5:7])}/{int(date[8:10])}到貨" if len(date or "") >= 10 else f"momo-{wh}"


@momo_bp.route("/api/momo/purchase/generate", methods=["POST"])
def api_purchase_generate():
    p = request.get_json(silent=True) or {}
    nos = [str(x).strip() for x in (p.get("inbound_nos") or []) if str(x).strip()]
    if not nos:
        return jsonify({"error": "請先勾選要產出 EIP 採購單的入庫單。"}), 400
    operator, stamp = _operator(), now()
    conn = get_conn()
    try:
        orders = build_plan(conn, nos=nos)
        if not orders:
            return jsonify({"error": "勾選的入庫單沒有原廠直出的品項，請重新整理頁面。"}), 400
        bad = [f"{g['inbound_no']} {x['factory_code']}：{'；'.join(x['problems'])}" for g in orders for x in g["items"] if x["blocking"]]
        if bad:
            return jsonify({"error": f"有 {len(bad)} 個品項需要先處理，無法產出。", "details": bad[:30]}), 400
        files = collections.OrderedDict()
        for g in orders:
            for x in g["items"]:
                if x["skip"]:
                    continue
                x["remark"] = _remark(x["line"], g["expected_date"], g["warehouse"])
                files.setdefault((x["line"], g["expected_date"], g["warehouse"], x["unit_group"]), []).append(x)
        if not files:
            return jsonify({"error": "勾選的入庫單品項都改成 0（不下採），沒有要產出的 EIP。"}), 400
        ids = [x["order_id"] for g in orders for x in g["items"] if not x["skip"]]
        marks = ",".join("?" * len(ids))
        old = _rows(conn.execute(f"SELECT i.id, i.eip_id FROM momo_eip_items i JOIN momo_eip e ON e.id = i.eip_id "
                                 f"WHERE i.order_id IN ({marks}) AND COALESCE(e.eip_po, '') = ''", ids))
        for x in old:                                      # 之前產過、還沒回填的：以這次為準
            conn.execute("DELETE FROM momo_eip_items WHERE id = ?", (x["id"],))
        for eid in {x["eip_id"] for x in old}:
            if conn.execute("SELECT COUNT(*) AS n FROM momo_eip_items WHERE eip_id = ?", (eid,)).fetchone()["n"] == 0:
                conn.execute("DELETE FROM momo_eip WHERE id = ?", (eid,))
        buf, names = io.BytesIO(), []
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for (line, date, wh, ug), ps in files.items():
                rows = _merge_rows(ps)
                name = eip_filename(line, date, wh, ug)
                zf.writestr(name, eip_xls(rows, line))
                names.append(name)
                cur = conn.execute("INSERT INTO momo_eip (batch, line, warehouse, expected_date, unit_group, filename, rows_json, eip_po, created_by, "
                                   "created_at, updated_by, updated_at) VALUES (?,?,?,?,?,?,?,?,?,?,?,?)",
                                   (stamp, line, wh, date, ug, name, json.dumps(rows, ensure_ascii=False), "", operator, stamp, operator, stamp))
                for x in ps:
                    conn.execute("INSERT INTO momo_eip_items (eip_id, order_id) VALUES (?,?)", (cur.lastrowid, x["order_id"]))
        for g in orders:
            _log(conn, "、".join(g["lines"]), g["inbound_no"], "", "", "momo_eip", "產出 EIP 採購單", "", f"{len(g['items'])} 品項", operator, "manual",
                 "、".join(names))
        conn.commit()
    finally:
        conn.close()
    ds = sorted({k[1] for k in files})
    buf.seek(0)
    resp = send_file(buf, as_attachment=True, download_name=f"MOMO_EIP採購單_{_md(ds[0])}" + (f"-{_md(ds[-1])}" if ds[-1] != ds[0] else "") + ".zip",
                     mimetype="application/zip")
    resp.headers["X-Momo-Files"] = str(len(names))
    return resp


# ── EIP 檔：清單、回填單號、取消產出、重新下載 ─────────────────────────────
def _eip_out(e, links):
    rows = json.loads(e["rows_json"] or "[]")
    return {k: e[k] for k in ("id", "line", "warehouse", "expected_date", "unit_group", "filename", "eip_po", "created_by", "created_at", "updated_by",
                              "updated_at")} | {"rows": len(rows), "boxes": round(sum(r["qty"] for r in rows), 2), "mars": e["line"] == MARS,
                                                "direct": e["line"] != MARS, "orders": sorted(set(links.get(e["id"], [])))}


def eips_of(conn, where="1=1", params=()):
    eips = _rows(conn.execute(f"SELECT * FROM momo_eip WHERE {where} ORDER BY expected_date DESC, warehouse, line, unit_group, id", list(params)))
    links = collections.defaultdict(list)
    if eips:
        marks = ",".join("?" * len(eips))
        for r in _rows(conn.execute(f"SELECT i.eip_id, o.inbound_no FROM momo_eip_items i JOIN momo_orders o ON o.id = i.order_id "
                                    f"WHERE i.eip_id IN ({marks})", [e["id"] for e in eips])):
            links[r["eip_id"]].append(r["inbound_no"])
    return [_eip_out(e, links) for e in eips]


@momo_bp.route("/api/momo/eips")
def api_eips():
    where, params = ["1=1"], []
    if request.args.get("from"):
        where.append("expected_date >= ?"); params.append(request.args["from"])
    if request.args.get("to"):
        where.append("expected_date <= ?"); params.append(request.args["to"])
    if request.args.get("open") == "1":
        where.append("COALESCE(eip_po, '') = ''")
    conn = get_conn()
    try:
        return jsonify({"eips": eips_of(conn, " AND ".join(where), params)})
    finally:
        conn.close()


@momo_bp.route("/api/momo/eips/<int:eid>", methods=["PUT"])
def api_eip_update(eid):
    eip = str((request.get_json(silent=True) or {}).get("eip_po") or "").strip().upper().replace(" ", "")
    if eip and not EIP_PO_RE.match(eip):
        return jsonify({"error": f"EIP 採購單號格式為 PO 加 9 碼數字（例如 PO202610003），目前填的是「{eip}」。"}), 400
    operator = _operator()
    conn = get_conn()
    try:
        e = _row(conn.execute("SELECT * FROM momo_eip WHERE id = ?", (eid,)))
        if e is None:
            return jsonify({"error": "找不到這張 EIP 採購單，可能已重新產出，請重新整理頁面。"}), 404
        if eip:                                            # EIP 單號全公司不重複：蝦皮特選用過的也擋
            dup = _row(conn.execute("SELECT filename FROM momo_eip WHERE eip_po = ? AND id != ?", (eip, eid))) \
                or _row(conn.execute("SELECT filename FROM shp_eip WHERE eip_po = ?", (eip,)))
            if dup:
                return jsonify({"error": f"{eip} 已用於「{dup['filename']}」，請確認號碼是否正確。"}), 400
        if eip != (e["eip_po"] or ""):
            conn.execute("UPDATE momo_eip SET eip_po = ?, updated_by = ?, updated_at = ? WHERE id = ?", (eip, operator, now(), eid))
            _log(conn, e["line"], "", "", "", "momo_eip_po", "EIP 採購單號", e["eip_po"], eip, operator, "manual", e["filename"])
            conn.commit()
        return jsonify({"ok": True, "eip_po": eip})
    finally:
        conn.close()


@momo_bp.route("/api/momo/eips/<int:eid>", methods=["DELETE"])
def api_eip_cancel(eid):
    """取消產出：這份檔刪掉，裡面的入庫單退回待產 EIP 採購單。已回填單號的要先清掉單號。"""
    operator = _operator()
    conn = get_conn()
    try:
        e = _row(conn.execute("SELECT * FROM momo_eip WHERE id = ?", (eid,)))
        if e is None:
            return jsonify({"error": "找不到這張 EIP 採購單，可能已重新產出，請重新整理頁面。"}), 404
        if e["eip_po"]:
            return jsonify({"error": f"這份已回填 EIP 採購單號 {e['eip_po']}，請先清掉單號再取消產出。"}), 400
        nos = sorted({r["inbound_no"] for r in _rows(conn.execute(
            "SELECT o.inbound_no FROM momo_eip_items i JOIN momo_orders o ON o.id = i.order_id WHERE i.eip_id = ?", (eid,)))})
        conn.execute("DELETE FROM momo_eip_items WHERE eip_id = ?", (eid,))
        conn.execute("DELETE FROM momo_eip WHERE id = ?", (eid,))
        for no in nos:
            _log(conn, e["line"], no, "", "", "momo_eip", "取消產出 EIP 採購單", e["filename"], "未產出", operator, "manual")
        conn.commit()
        return jsonify({"ok": True, "orders": nos})
    finally:
        conn.close()


def _get_eip(conn, eid):
    return _row(conn.execute("SELECT * FROM momo_eip WHERE id = ?", (eid,)))


@momo_bp.route("/api/momo/eips/<int:eid>/file")
def api_eip_file(eid):
    conn = get_conn()
    try:
        e = _get_eip(conn, eid)
    finally:
        conn.close()
    if e is None:
        return jsonify({"error": "找不到這張 EIP 採購單。"}), 404
    name = e["filename"] if not e["eip_po"] else e["filename"].replace(".xls", f"({e['eip_po']}).xls")
    return send_file(io.BytesIO(eip_xls(json.loads(e["rows_json"] or "[]"), e["line"])), as_attachment=True, download_name=name,
                     mimetype="application/vnd.ms-excel")


# ── 瑪氏採購單（需求七）────────────────────────────────────────────────────
def f5_text(uclass, label, rows):
    lines = [F5_FIXED]
    if label == "V":
        lines.append("需貼中盒標")
    if uclass == "箱":
        lines.append("貼板嘜")
    if any((r.get("po_case_note") or "").strip() for r in rows):
        lines.append("指定品需加工貼小白標")
    return "\n".join(lines)


def mars_po_name(e, rows):
    r0 = rows[0] if rows else {}
    lab = "需貼中標" if r0.get("label") == "V" else "不貼中標"
    tail = "(盒包)" if r0.get("unit_class") == "盒包" else ""        # 箱單、盒包單其他片段一樣時不撞名（問題清單第 12 題，先這樣做）
    return f"{e['eip_po'] or ''}永豐Mars採購單_箱_{r0.get('category') or '品類不明'}_{lab}_{e['warehouse']}_{_md(e['expected_date'])}{tail}.xlsx"


def mars_po_bytes(conn, e, whs=None, base=None):
    rows = json.loads(e["rows_json"] or "[]")
    r0 = rows[0] if rows else {}
    whs = whs or wh_settings(conn)
    w = whs.get(wh_canon(e["warehouse"], whs)) or {}
    settings = dict(base or mars_settings(conn))
    settings.update(lead_days=LEAD_DAYS, shelf_req=SHELF_REQ, shelf_req_cat={}, label_line="", white_line="",
                    special_text=f5_text(r0.get("unit_class"), r0.get("label"), rows))
    s = {"delivery_date": e["expected_date"], "eip_po": e["eip_po"], "category": r0.get("category", ""), "label": r0.get("label", ""),
         "unit": "箱" if r0.get("unit_class") == "箱" else r0.get("in_unit", ""), "slot_time": ""}
    items = [{"yf_sku": r["code"], "unit": r["in_unit"], "mars_code": r.get("mars_code", ""), "mars_name": r.get("mars_name", ""),
              "qty_ship": r["in_qty"], "cases": r.get("cases", r["qty"]), "price": r.get("price"), "pcs_per_case": r.get("pcs_per_case"),
              "inner_per_case": r.get("inner_per_case"), "note": r.get("note", ""), "po_case_note": r.get("po_case_note", "")} for r in rows]
    wh = {"name": w.get("mars_name") or e["warehouse"], "address": w.get("address", ""), "phone": w.get("phone", ""),
          "ship_to": w.get("ship_to", ""), "contact": w.get("contact", ""), "special_note": ""}
    return mars_po_name(e, rows), po_file(s, items, wh, settings)


# ── 直送採購單（需求六）────────────────────────────────────────────────────
def direct_rows(conn, e):
    rows = json.loads(e["rows_json"] or "[]")
    ids = [x["order_id"] for r in rows for x in r.get("parts") or []]
    orders = {}
    if ids:
        marks = ",".join("?" * len(ids))
        orders = {o["id"]: o for o in _rows(conn.execute(f"SELECT * FROM momo_orders WHERE id IN ({marks})", ids))}
    out = []
    for r in rows:
        for x in r.get("parts") or []:
            o = orders.get(x["order_id"]) or {}
            out.append({"po_id": o.get("inbound_no", ""), "sku": o.get("factory_code") or r["code"], "name": _name(o) if o else r["name"],
                        "qty": x["use_qty"], "unit": r["in_unit"], "box": r.get("conv") or 1})
    return out


def direct_name(e):
    return f"{e['eip_po'] or ''}MOMO直送採購單_{e['line']}_{_md(e['expected_date'])}_{e['warehouse']}.xlsx"


def direct_bytes(conn, e, whs=None):
    whs = whs or wh_settings(conn)
    w = whs.get(wh_canon(e["warehouse"], whs)) or {}
    data = direct_xlsx(e, direct_rows(conn, e), wh_note(w), today=_dt.date.fromisoformat(now()[:10]), channel=CHANNEL,
                       a22=f"momo-{e['warehouse']}", notice="")
    return direct_name(e), data


@momo_bp.route("/api/momo/eips/<int:eid>/mars_po")
def api_mars_po(eid):
    conn = get_conn()
    try:
        e = _get_eip(conn, eid)
        if e is None or e["line"] != MARS:
            return jsonify({"error": "找不到這張瑪氏 EIP 採購單，請重新整理頁面。"}), 404
        if not e["eip_po"]:
            return jsonify({"error": "請先回填 EIP 採購單號，才能產出瑪氏採購單。"}), 400
        name, data = mars_po_bytes(conn, e)
    finally:
        conn.close()
    return send_file(io.BytesIO(data), as_attachment=True, download_name=name,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@momo_bp.route("/api/momo/eips/<int:eid>/direct_po")
def api_direct_po(eid):
    conn = get_conn()
    try:
        e = _get_eip(conn, eid)
        if e is None or e["line"] == MARS:
            return jsonify({"error": "找不到這張 EIP 採購單，請重新整理頁面。"}), 404
        if not e["eip_po"]:
            return jsonify({"error": "請先回填 EIP 採購單號，才能產出直送採購單。"}), 400
        name, data = direct_bytes(conn, e)
    finally:
        conn.close()
    return send_file(io.BytesIO(data), as_attachment=True, download_name=name,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@momo_bp.route("/api/momo/po.zip")
def api_po_zip():
    """到貨日範圍內、已回填 EIP 單號的採購單一次下載：kind=mars 瑪氏採購單、direct 直送採購單。"""
    kind = request.args.get("kind", "")
    where, params = ["eip_po != ''", "line = ?" if kind == "mars" else "line != ?"], [MARS]
    for k, op in (("from", ">="), ("to", "<=")):
        if request.args.get(k):
            where.append(f"expected_date {op} ?"); params.append(request.args[k])
    conn = get_conn()
    try:
        eips = _rows(conn.execute(f"SELECT * FROM momo_eip WHERE {' AND '.join(where)} ORDER BY expected_date, warehouse, id", params))
        if not eips:
            return jsonify({"error": f"這段到貨日沒有已回填 EIP 採購單號的{'瑪氏' if kind == 'mars' else '直送'}採購單。"}), 400
        whs = wh_settings(conn)
        base = mars_settings(conn)
        files = [mars_po_bytes(conn, e, whs, base) if kind == "mars" else direct_bytes(conn, e, whs) for e in eips]
    finally:
        conn.close()
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in files:
            zf.writestr(name, data)
    buf.seek(0)
    ds = sorted({e["expected_date"] for e in eips})
    zname = f"MOMO{'瑪氏' if kind == 'mars' else '直送'}採購單_{_md(ds[0])}" + (f"-{_md(ds[-1])}" if ds[-1] != ds[0] else "") + ".zip"
    resp = send_file(buf, as_attachment=True, download_name=zname, mimetype="application/zip")
    resp.headers["X-Momo-Files"] = str(len(files))
    return resp


__all__ = ["build_plan", "eip_map", "skipped_ids", "eips_of", "eip_xls", "mars_po_bytes", "direct_bytes", "CAT_FULL"]

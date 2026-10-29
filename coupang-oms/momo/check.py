"""③ 待核缺貨與效期（需求四.5、五；Jerry 2026-10-28 看過示意圖「OK 照這樣做」）。只有瑪氏；紙品、潔品的效期照舊人工填 MOMO 後台。

- 勇信配送明細表（PDF，跟酷澎瑪氏、蝦皮特選同一種，讀檔用 mars.shortage.parse_yx_pdf）：收貨單號前面那段是 EIP 採購單號，
  對回 MOMO 的瑪氏 EIP（回填過單號的）。同一份 EIP 再傳就換掉上一次。不是 MOMO 的單號（蝦皮、酷澎的）略過、列出來。
- 一品一品比：這品下了幾箱（EIP 存的數量÷箱入數）、勇信出了幾箱。同一份 EIP 有好幾張入庫單訂同一個貨號，照入庫單順序分，最後一張拿剩下的。
  勇信 0 箱 → 不出貨（先勾好，人可以改）；一樣多 → 出貨、寫效期；少了（部分缺貨）、多了 → 標黃要人決定（問題清單：部分缺貨怎麼處理還沒問）。
- 效期：兩批效期取數量多的那批（需求五「瑪氏雙效期」）；一樣多 → 要人選。效期要晚於 MOMO 的「入庫最低有效期限」才寫（需求五「晚於」，
  剛好同一天算不夠，問題清單第 14 題）；不夠、或勇信沒印效期 → 要人看。
- 改成 0 不下採的品項（贈品）→ 不出貨（劃除範例裡贈品是畫掉的）。
- 效期匯入檔：MOMO 後台下載的那份（就是 UnsendList，「商品有效日期」U 欄空著），照入庫單號＋商品編號＋單品編號（倉別也要對）一列一列找，
  不照列的順序（需求五）；出貨的寫效期，不出貨的留白。還有要人決定的品項就不產，列出來。
"""
import collections
import datetime as _dt

import openpyxl
import xlrd
from xlutils.copy import copy as xl_copy

from mars.products import _code_canon, item_code, resolve_code, sku_codes
from mars.shortage import parse_yx_pdf

from .common import *  # noqa: F401,F403
from .orders import okey
from .parse import BAD, ParseError, _sheets, _text

S_WAIT, S_OK, S_NONE, S_PART, S_OVER, S_SKIP = "wait", "ok", "none", "partial", "over", "skip"


def _num(x):
    return int(x) if isinstance(x, float) and x.is_integer() else x


# ── 勇信配送明細表 ───────────────────────────────────────────────────────────
@momo_bp.route("/api/momo/yx", methods=["POST"])
def api_yx():
    files = [f for f in request.files.getlist("file") if f and f.filename]
    if not files:
        return jsonify({"error": "請上傳勇信配送明細表（PDF）。"}), 400
    pages, errs = [], []
    for f in files:
        try:
            ps = parse_yx_pdf(f.read(), f.filename)
        except ValueError as exc:
            errs.append(str(exc)); continue
        for p in ps:
            p["filename"] = f.filename
        pages += ps
    if not pages:
        return jsonify({"error": "讀不到勇信配送明細表。", "details": errs}), 400
    operator, stamp = _operator(), now()
    conn = get_conn()
    try:
        conflicts = {}
        canon = _code_canon(conn, conflicts)
        by_eip = collections.OrderedDict()
        for p in pages:
            if p.get("eip_po"):
                by_eip.setdefault(p["eip_po"], []).append(p)
        matched, unknown, warnings = [], [], [w for p in pages for w in p.get("warnings") or [] if "讀不到倉別" not in w]
        for eip_po, ps in by_eip.items():
            eips = _rows(conn.execute("SELECT * FROM momo_eip WHERE eip_po = ? AND line = ?", (eip_po, MARS)))
            if not eips:
                unknown.append(eip_po); continue
            skus = sku_codes(conn)
            present = {item_code(r.get("code"), r.get("mars_code"), skus, canon, conflicts) for e in eips for r in json.loads(e["rows_json"] or "[]")}
            data = {}
            for p in ps:
                for it in p.get("items") or []:
                    code, _ask = resolve_code(it["code"], conflicts, canon, present)
                    d = data.setdefault(code, {"cases": 0, "name": it.get("name") or "", "batches": []})
                    d["cases"] += it.get("qty") or 0
                    d["batches"] += [dict(b) for b in it.get("batches") or []]
            label = "、".join(f"{p['filename']} 第 {p['page']} 頁" for p in ps)
            for e in eips:
                conn.execute("DELETE FROM momo_yx WHERE eip_id = ?", (e["id"],))
                conn.execute("INSERT INTO momo_yx (eip_id, eip_po, filename, pages, data_json, warnings_json, uploaded_by, uploaded_at) VALUES (?,?,?,?,?,?,?,?)",
                             (e["id"], eip_po, ps[0]["filename"], label, json.dumps(data, ensure_ascii=False), "[]", operator, stamp))
                matched.append(e["id"])
            for r in _rows(conn.execute(f"SELECT DISTINCT o.inbound_no FROM momo_eip_items i JOIN momo_orders o ON o.id = i.order_id "
                                        f"WHERE i.eip_id IN ({','.join('?' * len(eips))})", [e["id"] for e in eips])):
                _log(conn, MARS, r["inbound_no"], "", "", "momo_yx", "勇信配送明細表", "", f"{eip_po}：{sum(d['cases'] for d in data.values())} 箱",
                     operator, "import", label)
        conn.commit()
        nos = sorted({r["inbound_no"] for r in _rows(conn.execute(
            f"SELECT DISTINCT o.inbound_no FROM momo_eip_items i JOIN momo_orders o ON o.id = i.order_id WHERE i.eip_id IN ({','.join('?' * len(matched))})",
            matched))}) if matched else []
        res = check_orders(conn, nos)
    finally:
        conn.close()
    return jsonify({"ok": True, "eip_pos": [k for k in by_eip if k not in unknown], "unknown": unknown, "warnings": warnings + errs,
                    "orders": [{"inbound_no": no, "warehouse": res[no]["warehouse"], **res[no]["summary"]} for no in nos if no in res]})


# ── 核對結果（一張入庫單一品一列）────────────────────────────────────────────
def _choose(batches):
    """兩批以上效期：取數量多的那批；一樣多回 None（要人選）。同一個效期的批次先加總。"""
    by = collections.Counter()
    for b in batches:
        if b.get("expiry"):
            by[b["expiry"]] += b.get("qty") or 0
    if not by:
        return "", False
    top = by.most_common()
    if len(top) > 1 and top[0][1] == top[1][1]:
        return "", True
    return top[0][0], False


def check_orders(conn, nos=None):
    """{入庫單號: {warehouse, expected_date, items:[...], summary:{...}}}；只算瑪氏、已回填 EIP 單號的入庫單。"""
    eips = _rows(conn.execute("SELECT * FROM momo_eip WHERE line = ? AND eip_po != '' ORDER BY id", (MARS,)))
    if not eips:
        return {}
    eids = [e["id"] for e in eips]
    marks = ",".join("?" * len(eids))
    links = collections.defaultdict(list)
    for r in _rows(conn.execute(f"SELECT i.eip_id, o.* FROM momo_eip_items i JOIN momo_orders o ON o.id = i.order_id "
                                f"WHERE i.eip_id IN ({marks}) AND o.removed = 0 ORDER BY o.inbound_no, o.id", eids)):
        links[r["eip_id"]].append(r)
    want = set(nos) if nos is not None else None
    yx = {r["eip_id"]: r for r in _rows(conn.execute(f"SELECT * FROM momo_yx WHERE eip_id IN ({marks})", eids))}
    conflicts = {}
    canon, skus = _code_canon(conn, conflicts), sku_codes(conn)
    out = {}
    for e in eips:
        orders = links.get(e["id"], [])
        if want is not None and not any(o["inbound_no"] in want for o in orders):
            continue
        rows = json.loads(e["rows_json"] or "[]")
        part_of = {}                                      # order_id → (EIP 那一列, 這張單下的數量)
        for r in rows:
            for x in r.get("parts") or []:
                part_of[x["order_id"]] = (r, x["use_qty"])
        y = yx.get(e["id"])
        ydata = json.loads(y["data_json"] or "{}") if y else {}
        left = collections.Counter({c: d["cases"] for c, d in ydata.items()})
        last = {}
        items = []
        for o in orders:
            r, use = part_of.get(o["id"], ({}, 0))
            box = r.get("conv") or 0
            code = item_code(r.get("code") or o["base_code"], r.get("mars_code"), skus, canon, conflicts)
            it = {"order_id": o["id"], "inbound_no": o["inbound_no"], "warehouse": o["warehouse"], "factory_code": o["factory_code"],
                  "product_no": o["product_no"], "item_no": o["item_no"], "name": o["product_name"], "qty": o["qty"], "unit": o["unit_name"],
                  "mars_code": code, "ordered": _num(round(use / box, 4)) if box else 0, "conv": box, "min_expiry": o["min_expiry"] or "",
                  "eip_po": e["eip_po"], "shipped": None, "batches": (ydata.get(code) or {}).get("batches", []), "yx_at": y["uploaded_at"] if y else ""}
            last[code] = len(items)
            items.append(it)
        for i, it in enumerate(items):
            if not y:
                continue
            c = it["mars_code"]
            got = left.get(c, 0) if last.get(c) == i else min(left.get(c, 0), it["ordered"])
            left[c] -= got
            it["shipped"] = _num(got)
        for it in items:
            out.setdefault(it["inbound_no"], {"items": []})["items"].append(it)
    if want is not None:
        out = {k: v for k, v in out.items() if k in want}
    if not out:
        return {}
    # 同一張入庫單沒下採（改成 0）的品項也要列：入庫通知單上一樣要畫掉
    nos_all = sorted(out)
    marks = ",".join("?" * len(nos_all))
    linked = {it["order_id"] for v in out.values() for it in v["items"]}
    for o in _rows(conn.execute(f"SELECT * FROM momo_orders WHERE inbound_no IN ({marks}) AND removed = 0 ORDER BY id", nos_all)):
        if o["id"] not in linked and o["ship_method"] == DIRECT:
            out[o["inbound_no"]]["items"].append({"order_id": o["id"], "inbound_no": o["inbound_no"], "warehouse": o["warehouse"],
                                                   "factory_code": o["factory_code"], "product_no": o["product_no"], "item_no": o["item_no"],
                                                   "name": o["product_name"], "qty": o["qty"], "unit": o["unit_name"], "mars_code": "",
                                                   "ordered": 0, "conv": 0, "min_expiry": o["min_expiry"] or "", "eip_po": "", "shipped": None,
                                                   "batches": [], "yx_at": "", "skip": True})
    ids = [it["order_id"] for v in out.values() for it in v["items"]]
    m2 = ",".join("?" * len(ids))
    over = {r["order_id"]: r for r in _rows(conn.execute(f"SELECT * FROM momo_check WHERE order_id IN ({m2})", ids))}
    heads = {r["inbound_no"]: r for r in _rows(conn.execute(
        f"SELECT inbound_no, MIN(warehouse) AS warehouse, MIN(expected_date) AS expected_date FROM momo_orders WHERE inbound_no IN ({marks}) GROUP BY inbound_no", nos_all))}
    marks_ = {r["inbound_no"]: r for r in _rows(conn.execute(
        f"SELECT inbound_no, filename, uploaded_by, uploaded_at, drawn_by, drawn_at FROM momo_mark WHERE inbound_no IN ({marks})", nos_all))}
    notices = {r["inbound_no"]: r for r in _rows(conn.execute(
        f"SELECT inbound_no, filename, uploaded_by, uploaded_at, drawn_by, drawn_at FROM momo_notice WHERE inbound_no IN ({marks})", nos_all))}
    fills = collections.defaultdict(str)
    for f in _rows(conn.execute("SELECT inbound_nos, created_at FROM momo_expfill ORDER BY id")):
        for no in (f["inbound_nos"] or "").split(","):
            if no in out:
                fills[no] = max(fills[no], f["created_at"])
    for no, v in out.items():
        for it in v["items"]:
            judge(it, over.get(it["order_id"]))
            # 實際出貨數（MOMO 的單位）：勇信出的箱數 × 箱入數；不出貨＝0。跟 MOMO 訂的不一樣就是「短出」，入庫通知單要畫掉寫實際數（Becky 2026-10-29）
            it["actual"] = _num((it["shipped"] or 0) * (it["conv"] or 0)) if it["decision"] == "ship" else (0 if it["decision"] == "noship" else None)
        its = v["items"]
        yx_at = max((it["yx_at"] for it in its), default="")
        filled = fills.get(no, "")
        ship = [it for it in its if it["decision"] == "ship"]
        v.update(warehouse=heads[no]["warehouse"], expected_date=heads[no]["expected_date"],
                 summary={"items": len(its), "wait": sum(1 for it in its if it["decision"] == "wait"),
                          "pending": sum(1 for it in its if it["decision"] == "pending"), "ship": len(ship),
                          "noship": sum(1 for it in its if it["decision"] == "noship"),
                          "noship_bought": sum(1 for it in its if it["decision"] == "noship" and not it.get("skip")),
                          "skip": sum(1 for it in its if it.get("skip")),
                          "ask_qty": sum(1 for it in its if it["decision"] == "pending" and it["status"] in (S_PART, S_OVER)),
                          "short": sum(1 for it in its if it["decision"] == "ship" and it["actual"] is not None and it["actual"] != (it["qty"] or 0)),
                          "yx_at": yx_at if all(it["yx_at"] or it.get("skip") for it in its) else "",
                          "filled_at": filled if filled and filled >= yx_at else ""})
        s = v["summary"]
        s["done"] = not s["wait"] and not s["pending"] and (not s["ship"] or bool(s["filled_at"]))
        nt = notices.get(no)
        s["notice"] = {k: nt[k] for k in ("filename", "uploaded_by", "uploaded_at", "drawn_by", "drawn_at")} if nt else None
        mk = marks_.get(no)
        s["mark"] = {k: mk[k] for k in ("filename", "uploaded_by", "uploaded_at", "drawn_by", "drawn_at")} if mk else None
        s["notice_need"] = s["done"] and (s["noship"] > 0 or s["short"] > 0)          # ④：有不出貨或短出的品項才要畫
        s["notice_done"] = bool(nt and nt["drawn_at"] and nt["drawn_at"] >= yx_at)
    return out


def judge(it, ov):
    """一品的結果：decision＝ship（寫效期）／noship（畫掉）／pending（要人決定）／wait（勇信表還沒來）。"""
    it["override"] = ({"ship": ov["ship"], "expiry": ov["expiry"] or "", "reason": ov["reason"], "operator": ov["operator"], "at": ov["updated_at"]}
                      if ov else None)
    if it.get("skip"):
        it.update(status=S_SKIP, decision="noship", expiry="", why="不下採（數量改成 0）")
        if ov and ov["ship"] == 1:
            it.update(decision="pending", why="不下採的品項不能改成出貨")
        return it
    if it["shipped"] is None:
        it.update(status=S_WAIT, decision="wait", expiry="", why="勇信配送明細表還沒來")
        return it
    shipped, ordered = it["shipped"], it["ordered"]
    it["status"] = S_OK if shipped == ordered else S_NONE if shipped == 0 else S_OVER if shipped > ordered else S_PART
    exp, tie = _choose(it["batches"])
    if ov and ov["expiry"]:
        exp, tie = ov["expiry"], False
    it["expiry"], it["tie"] = exp, tie
    if ov and ov["ship"] == 0:
        it.update(decision="noship", why="人工改成不出貨")
        return it
    forced = bool(ov and ov["ship"] == 1)
    if it["status"] == S_NONE and not forced:
        it.update(decision="noship", why=f"勇信沒出（訂 {ordered} 箱）")
        return it
    if it["status"] in (S_PART, S_OVER) and not forced:
        it.update(decision="pending", why=f"勇信出 {shipped} 箱，訂 {ordered} 箱，請決定出不出貨")
        return it
    if tie:
        it.update(decision="pending", why="兩批效期數量一樣，請選一個")
    elif not exp:
        it.update(decision="pending", why="勇信表上沒有效期，請填")
    elif it["min_expiry"] and exp.replace("/", "-") <= it["min_expiry"].replace("/", "-"):
        it.update(decision="pending", why=f"效期 {exp} 沒有晚於 MOMO 最低效期 {it['min_expiry']}")
    else:
        it.update(decision="ship", why="")
    return it


@momo_bp.route("/api/momo/check")
def api_check():
    no = str(request.args.get("no") or "").strip()
    conn = get_conn()
    try:
        res = check_orders(conn, [no] if no else None)
    finally:
        conn.close()
    return jsonify({"orders": [{"inbound_no": k, **v} for k, v in sorted(res.items(), key=lambda kv: (kv[1]["expected_date"], kv[0]))]})


@momo_bp.route("/api/momo/check", methods=["POST"])
def api_check_set():
    """人改一品：ship＝1 出貨、0 不出貨、空＝回到系統判斷；expiry＝用哪個效期（YYYY-MM-DD）。"""
    p = request.get_json(silent=True) or {}
    try:
        oid = int(p.get("order_id"))
    except (TypeError, ValueError):
        return jsonify({"error": "缺少品項。"}), 400
    ship = p.get("ship")
    ship = None if ship in (None, "") else (1 if str(ship) in ("1", "true", "True") else 0)
    exp = str(p.get("expiry") or "").strip().replace("/", "-")
    if exp:
        try:
            exp = _dt.date.fromisoformat(exp).isoformat()
        except ValueError:
            return jsonify({"error": "效期請填日期（例如 2028-02-08）。"}), 400
    reason = str(p.get("reason") or "").strip()
    operator = _operator()
    conn = get_conn()
    try:
        o = _row(conn.execute("SELECT * FROM momo_orders WHERE id = ?", (oid,)))
        if o is None:
            return jsonify({"error": "找不到這個品項，請重新整理頁面。"}), 404
        old = _row(conn.execute("SELECT * FROM momo_check WHERE order_id = ?", (oid,)))
        if ship is None and not exp:
            if old:
                conn.execute("DELETE FROM momo_check WHERE order_id = ?", (oid,))
                _log(conn, MARS, o["inbound_no"], o["product_no"], "", "momo_check", "缺貨與效期", _desc(old), "回到系統判斷", operator, "manual", o["factory_code"])
                conn.commit()
            return jsonify({"ok": True, "cleared": True})
        stamp = now()
        if old:
            conn.execute("UPDATE momo_check SET ship = ?, expiry = ?, reason = ?, operator = ?, updated_at = ? WHERE order_id = ?",
                         (ship, exp, reason, operator, stamp, oid))
        else:
            conn.execute("INSERT INTO momo_check (order_id, ship, expiry, reason, operator, updated_at) VALUES (?,?,?,?,?,?)",
                         (oid, ship, exp, reason, operator, stamp))
        _log(conn, MARS, o["inbound_no"], o["product_no"], "", "momo_check", "缺貨與效期", _desc(old) if old else "系統判斷",
             _desc({"ship": ship, "expiry": exp}), operator, "manual", o["factory_code"], reason)
        conn.commit()
        return jsonify({"ok": True})
    finally:
        conn.close()


def _desc(x):
    parts = [{1: "出貨", 0: "不出貨"}.get(x.get("ship"), "")]
    if x.get("expiry"):
        parts.append(f"效期 {x['expiry']}")
    return "、".join(p for p in parts if p) or "—"


# ── 效期匯入檔 ───────────────────────────────────────────────────────────────
def _exp_col(rows):
    """表頭列號、各欄位置（入庫單號、倉別、商品編號、單品編號、商品有效日期、採購數量）。"""
    need = ("入庫單號", "商品編號", "單品編號", "商品有效日期")
    for i, row in enumerate(rows[:5]):
        heads = [_text(v).replace("\n", "") for v in row]
        if all(h in heads for h in need):
            return i, {h: heads.index(h) for h in heads if h in ("入庫單號", "倉別", "商品編號", "單品編號", "商品有效日期", "採購數量")}
    return None, None


def plan_fill(conn, data, filename):
    """讀 MOMO 效期匯入檔 → (要寫的 [(列號, 效期)], 這份的入庫單號, 問題清單)。"""
    for title, rows in _sheets(data, filename):
        hi, col = _exp_col(rows)
        if hi is not None:
            break
    else:
        raise ParseError(f"{filename}：找不到「入庫單號、商品編號、單品編號、商品有效日期」欄位，請上傳 MOMO 後台下載的效期匯入檔。")
    cur_no, cur_wh, keyed = "", "", []
    for r in range(hi + 1, len(rows)):
        row = rows[r]
        get = lambda k: _text(row[col[k]]) if k in col and col[k] < len(row) else ""  # noqa: E731
        if get("入庫單號"):
            cur_no, cur_wh = get("入庫單號"), get("倉別")
        prod = get("商品編號")
        if not prod or prod in BAD:
            continue
        keyed.append((r, okey({"inbound_no": cur_no, "product_no": prod, "item_no": get("單品編號")}), cur_no, cur_wh))
    nos = sorted({k[2] for k in keyed})
    res = check_orders(conn, nos)
    by_okey = {}
    for no, v in res.items():
        for it in v["items"]:
            by_okey[okey(it)] = it
    writes, problems = [], []
    missing = [no for no in nos if no not in res]
    if missing:
        problems.append(f"{'、'.join(missing)} 不在「待核缺貨與效期」（不是瑪氏、還沒回填 EIP 單號，或系統裡沒有這張單）")
    for r, k, no, wh in keyed:
        it = by_okey.get(k)
        if it is None:
            if no in res:
                problems.append(f"第 {r + 1} 列 {k.replace('|', ' ')}：系統裡這張入庫單沒有這個品項")
            continue
        if wh and it["warehouse"] and wh != it["warehouse"]:
            problems.append(f"第 {r + 1} 列 {no}：倉別 {wh} 跟系統的 {it['warehouse']} 不一樣")
            continue
        if it["decision"] in ("wait", "pending"):
            problems.append(f"第 {r + 1} 列 {no} {it['factory_code']}：{it['why']}")
        elif it["decision"] == "ship":
            writes.append((r, it["expiry"]))
    return title, col["商品有效日期"], writes, nos, problems


def write_fill(data, filename, sheet, ucol, writes):
    """把效期寫進原檔（.xls 用 xlutils 保留格式，.xlsx 用 openpyxl），日期照 V 欄的寫法 YYYY/MM/DD。"""
    if data[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" or filename.lower().endswith(".xls"):
        rb = xlrd.open_workbook(file_contents=data, formatting_info=True)
        wb = xl_copy(rb)
        ws = wb.get_sheet(rb.sheet_names().index(sheet))
        for r, exp in writes:
            ws.write(r, ucol, exp.replace("-", "/"))
        buf = io.BytesIO(); wb.save(buf)
        return buf.getvalue()
    wb = openpyxl.load_workbook(io.BytesIO(data))
    ws = wb[sheet]
    for r, exp in writes:
        ws.cell(row=r + 1, column=ucol + 1, value=exp.replace("-", "/"))
    buf = io.BytesIO(); wb.save(buf)
    return buf.getvalue()


@momo_bp.route("/api/momo/expfill", methods=["POST"])
def api_expfill():
    f = request.files.get("file")
    if f is None or not f.filename:
        return jsonify({"error": "請上傳 MOMO 後台下載的效期匯入檔。"}), 400
    data = f.read()
    operator = _operator()
    conn = get_conn()
    try:
        try:
            sheet, ucol, writes, nos, problems = plan_fill(conn, data, f.filename)
        except ParseError as exc:
            return jsonify({"error": str(exc)}), 400
        if problems:
            return jsonify({"error": f"有 {len(problems)} 個地方要先處理，還不能填效期。", "details": problems[:30]}), 400
        out = write_fill(data, f.filename, sheet, ucol, writes)
        conn.execute("INSERT INTO momo_expfill (filename, inbound_nos, filled, created_by, created_at) VALUES (?,?,?,?,?)",
                     (f.filename, ",".join(nos), len(writes), operator, now()))
        for no in nos:
            _log(conn, MARS, no, "", "", "momo_expfill", "填效期匯入檔", "", f"寫入 {len(writes)} 筆", operator, "manual", f.filename)
        conn.commit()
    finally:
        conn.close()
    base, dot, ext = f.filename.rpartition(".")
    resp = send_file(io.BytesIO(out), as_attachment=True, download_name=f"{base}_已填效期.{ext}" if dot else f"{f.filename}_已填效期",
                     mimetype="application/vnd.ms-excel")
    resp.headers["X-Momo-Filled"] = str(len(writes))
    return resp


def is_expfill(conn, data, filename):
    """上傳入口認檔：MOMO 下載檔，裡面每張入庫單都已經在「待核缺貨與效期」、而且勇信表來了 → 多半是效期匯入檔（畫面再讓人選）。"""
    try:
        _sheet, _u, _w, nos, _p = plan_fill(conn, data, filename)
    except Exception:  # noqa: BLE001
        return False
    if not nos:
        return False
    res = check_orders(conn, nos)
    return all(no in res and res[no]["summary"]["yx_at"] for no in nos)


__all__ = ["check_orders", "judge", "is_expfill", "api_yx", "api_check", "api_check_set", "api_expfill"]

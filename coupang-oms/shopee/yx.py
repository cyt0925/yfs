"""蝦皮特選瑪氏缺貨：勇信配送明細表核對（需求 8.1）。

勇信的配送明細表 PDF 跟酷澎瑪氏出貨用的是同一種（Jerry 2026-10-12：「永信效期表是長一樣的」），
讀檔直接用 mars.shortage.parse_yx_pdf，比對的想法也照酷澎那套：

- 每頁「收貨單號：PO202610101-…」前面那段是 EIP 採購單號 → 找到蝦皮特選那份瑪氏 EIP（採購單產出頁回填的號碼）。
  同一個號碼回填在好幾份上，就合起來比。
- 產品編號＝58880＋瑪氏貨號；舊貨號靠商品總表備註的換號紀錄對回現在的貨號（mars.shortage._code_canon）。
  同一頁同一貨號分兩批效期會列兩行，加總。
- 配送數量「34C」是箱數，跟我們訂的箱數比（訂購箱數＝蝦皮數量÷商品總表箱入數，有改過數量的用改過的）。
  相等＝不缺、少＝部分缺貨、0＝全缺、多＝數量異常（不動，要人看）。
- 範圍：PDF 的指送日期＋表上那幾份 EIP 的倉。那天那個倉已回填 EIP 單號、表上卻沒有的，算整份沒出
  （主管 2026-10-08 講酷澎的：整份缺的勇信不會印在表上，不能等人記得去勾）。人可以取消勾選「這次不核」的那份。
  還沒回填 EIP 單號的對不到，不算缺，列出來要人先補。
- 同一份 EIP 裡同一個貨號有好幾張蝦皮單（同一天同倉合在一列），勇信出的箱數依訂單先後分，最後一張拿剩下的（多出來的才看得到）。
- 效期（Jerry 2026-10-13「照文件改」：需求 8.1「核對庫存及效期」、5.3「不符合效期比例的不得計入可出貨量」、5.4 缺貨原因）：
  勇信每批印「箱數、批號、效期」，一批一批比蝦皮允收效期（到貨日＋商品總表效期天×竹運出貨頁的效期標準）。
  效期不夠的那幾箱不算可出貨，原因「效期不足」；勇信也少出的話兩個原因一起標。商品總表沒有效期天的沒辦法算，只提醒、不扣。

確認後寫進缺貨紀錄 shp_short（瑪氏只有一次，來源「勇信核對」），之後在缺貨維護與通知勾選訂單產 LINE 通知，
跟寶僑、紙潔同一套。還沒通知前可以重傳 PDF 重跑，舊的那次會被換掉；通知過就鎖住不再改。
"""
import datetime as _dt

from mars.shortage import _code_canon, parse_yx_pdf

from .common import *  # noqa: F401,F403
from .orders import order_no
from .purchase import _adj_map
from .ship import ratio_for, ship_settings

MARS = "瑪氏"
SOURCE = "勇信核對"
R_STOCK, R_EXP = "庫存不足", "效期不足"


def _num(x):
    return int(x) if isinstance(x, float) and x.is_integer() else x


def _need(shelf, expected_date, st):
    """蝦皮允收的最早效期（到貨日＋取整數(效期天×比例)，跟竹運出貨同一套）。沒有效期天或到貨日回 None。"""
    if not shelf or not expected_date:
        return None
    try:
        arrive = _dt.date.fromisoformat(expected_date)
    except ValueError:
        return None
    ratio, text = ratio_for(shelf, st)
    return {"date": (arrive + _dt.timedelta(days=int(shelf * ratio))).isoformat(), "rule": f"效期天 {shelf}×{text}"}


def compare(conn, pages, skip_eips=()):
    skip = {int(x) for x in skip_eips if str(x).strip().isdigit()}
    canon = _code_canon(conn)
    st = ship_settings(conn)
    by_eip = collections.OrderedDict()
    for p in pages:
        if not p.get("eip_po"):
            continue
        d = by_eip.setdefault(p["eip_po"], {"qty": collections.Counter(), "pages": [], "names": {}, "batches": collections.defaultdict(list)})
        d["pages"].append(p["page"])
        for it in p.get("items") or []:
            c = canon.get(it["code"], it["code"])
            d["qty"][c] += it.get("qty") or 0
            d["names"].setdefault(c, it.get("name") or "")
            d["batches"][c] += [dict(b) for b in it.get("batches") or []] or (
                [{"qty": it.get("qty") or 0, "expiry": (it.get("expiry") or "").split("、")[0]}] if it.get("qty") else [])
    dates = sorted({p["ship_date"] for p in pages if p.get("ship_date")})

    eips = _rows(conn.execute("SELECT * FROM shp_eip WHERE line = ? ORDER BY expected_date, warehouse, id", (MARS,)))
    groups = collections.OrderedDict()
    for e in eips:
        if e["eip_po"]:
            groups.setdefault(e["eip_po"], []).append(e)
    unknown = [{"eip_po": k, "pages": v["pages"], "cases": sum(v["qty"].values())} for k, v in by_eip.items() if k not in groups]
    found_ids = {e["id"] for k in by_eip for e in groups.get(k, [])}
    whs = sorted({e["warehouse"] for e in eips if e["id"] in found_ids})
    in_scope = [e for e in eips if e["id"] in found_ids or (e["expected_date"] in dates and e["warehouse"] in whs)]

    # 每份 EIP 底下的蝦皮品項
    ids = [e["id"] for e in in_scope]
    links = collections.defaultdict(list)
    if ids:
        marks = ",".join("?" * len(ids))
        for r in _rows(conn.execute(f"SELECT i.eip_id, o.* FROM shp_eip_items i JOIN shp_orders o ON o.id = i.order_id "
                                    f"WHERE i.eip_id IN ({marks}) AND o.removed = 0 ORDER BY o.id", ids)):
            links[r["eip_id"]].append(r)
    oids = [o["id"] for v in links.values() for o in v]
    adj = _adj_map(conn, oids)
    rounds = collections.defaultdict(list)
    if oids:
        marks = ",".join("?" * len(oids))
        for r in _rows(conn.execute(f"SELECT * FROM shp_short WHERE order_id IN ({marks}) ORDER BY round", oids)):
            rounds[r["order_id"]].append(r)
    shelf = {}
    codes = sorted({o["base_code"] for v in links.values() for o in v})
    if codes:
        marks = ",".join("?" * len(codes))
        shelf = {(r["yf_sku"], r["unit"]): r["shelf_days"] for r in _rows(conn.execute(
            f"SELECT yf_sku, unit, shelf_days FROM mst_mars_products WHERE yf_sku IN ({marks})", codes)) if r["shelf_days"]}

    def item_rows(e):
        snap = {(r.get("code"), r.get("unit")): r for r in json.loads(e["rows_json"] or "[]")}
        out = []
        for o in links.get(e["id"], []):
            unit = o["unit_name"] or UNIT_CODES.get(o["unit_code"] or "", "")
            s = snap.get((o["base_code"], unit)) or {}
            box = s.get("conv") or 0
            a = adj.get(o["id"])
            use = a["qty_after"] if a and a["qty_before"] == o["qty"] else (o["qty"] or 0)
            code = canon.get(s.get("mars_code") or "", s.get("mars_code") or "")
            rs = rounds.get(o["id"], [])
            out.append({"order_id": o["id"], "order_no": order_no(o), "pr_id": o["pr_id"] or "", "po_id": o["po_id"] or "",
                        "supplier_sku_id": o["supplier_sku_id"], "shopee_sku_id": o["shopee_sku_id"], "yf_sku": o["base_code"],
                        "mars_code": code, "name": o["sku_name"], "unit": unit, "box": box, "qty": o["qty"] or 0, "use_qty": use,
                        "ordered": _num(round(use / box, 4)) if box else 0, "shipped": 0, "good": 0, "short": 0, "status": "", "final": None,
                        "reasons": [], "batches": [], "need": "", "need_rule": "", "warn": "" if box else "EIP 上沒有箱入數，無法換算",
                        "locked": any(r["notify_id"] for r in rs),
                        "pending": next((r for r in rs if not r["notify_id"]), None), "_shelf": shelf.get((o["base_code"], unit))})
        return out

    def finish(row, shipped, good, info):
        """shipped＝分到的勇信箱數；good＝其中效期合格的箱數；info＝這個貨號的批次與允收效期（表上沒有就 None）。"""
        row.pop("_shelf", None)
        ordered = row["ordered"]
        judged = bool(info and info["need"])
        cap = min(shipped, ordered)
        usable = _num(min(good, cap)) if judged else cap
        row.update(shipped=shipped, good=usable)
        if info:
            row.update(batches=info["batches"], need=info["need"]["date"] if judged else "", need_rule=info["need"]["rule"] if judged else "")
            if not judged and info["batches"] and not row["warn"]:
                row["warn"] = "瑪氏商品總表沒有效期天，無法判斷效期（不扣）"
        if shipped > ordered:
            row["status"] = "over"
        else:
            row["status"] = "ok" if usable == ordered else "none" if usable == 0 else "partial"
            row["reasons"] = ([R_STOCK] if shipped < ordered else []) + ([R_EXP] if usable < cap else [])
            row["short"] = _num(round(ordered - usable, 4))
        if row["status"] in ("partial", "none") and row["box"]:
            row["final"] = min(row["qty"], int(round(usable * row["box"])))
        pend = row.pop("pending")
        row["was"] = {"after": pend["qty_after"], "note": pend["note"]} if pend else None
        if row["locked"]:
            row["action"] = "已通知，不再修改"
        elif row["status"] == "over":
            row["action"] = "數量異常，不修改"
        elif row["final"] is not None and row["final"] < row["qty"]:
            row["action"] = "刪單" if row["final"] == 0 else f"下修至 {row['final']}"
        else:
            row["action"] = "取消前次缺貨" if pend else "無缺貨"
        return row

    out, no_eip = [], []
    for e in in_scope:
        base = {"eip_id": e["id"], "eip_po": e["eip_po"] or "", "filename": e["filename"], "warehouse": e["warehouse"],
                "warehouse_name": wh_name(e["warehouse"]), "expected_date": e["expected_date"], "unit_group": e["unit_group"],
                "skipped": e["id"] in skip, "in_pdf": e["id"] in found_ids, "pages": [], "items": [], "extra": []}
        if not e["eip_po"]:
            no_eip.append(base)
            continue
        out.append(base)
    # PDF 上有的號碼：同一個號碼的幾份合起來分箱數；表上沒有的那份全部算 0
    for eip_po, d in by_eip.items():
        group = [b for b in out if b["eip_po"] == eip_po]
        if not group:
            continue
        rows = [(b, item_rows(next(e for e in in_scope if e["id"] == b["eip_id"]))) for b in group]
        last = {}
        for gi, (_b, its) in enumerate(rows):
            for ii, r in enumerate(its):
                last[r["mars_code"]] = (gi, ii)
        # 每個貨號：每批效期夠不夠（同一份 EIP 到貨日一樣，效期天看第一個訂這個貨號的品項）
        info = {}
        for gi, (b, its) in enumerate(rows):
            e = next(x for x in in_scope if x["id"] == b["eip_id"])
            for r in its:
                c = r["mars_code"]
                if c in info:
                    continue
                need = _need(r["_shelf"], e["expected_date"], st)
                bs = [{**x, "ok": (need is None or not x["expiry"] or x["expiry"] >= need["date"])} for x in d["batches"].get(c, [])]
                info[c] = {"need": need, "batches": bs, "good": sum(x["qty"] for x in bs if x["ok"])}
        left = collections.Counter(d["qty"])
        left_good = collections.Counter({c: v["good"] for c, v in info.items()})
        for gi, (b, its) in enumerate(rows):
            b["pages"] = d["pages"]
            for ii, r in enumerate(its):
                c = r["mars_code"]
                is_last = last.get(c) == (gi, ii)
                got = left.get(c, 0) if is_last else min(left.get(c, 0), r["ordered"])
                good = left_good.get(c, 0) if is_last else min(left_good.get(c, 0), r["ordered"])
                left[c] -= got
                left_good[c] -= good
                b["items"].append(finish(r, _num(got), _num(good), info.get(c)))
        seen = {r["mars_code"] for _b, its in rows for r in its}
        group[0]["extra"] = [{"mars_code": c, "name": d["names"].get(c, ""), "cases": q} for c, q in d["qty"].items() if c not in seen and q]
    for b in out:
        if not b["in_pdf"]:
            e = next(x for x in in_scope if x["id"] == b["eip_id"])
            b["items"] = [finish(r, 0, 0, None) for r in item_rows(e)]
    todo = [r for b in out if not b["skipped"] for r in b["items"]]
    summary = {"pages": sum(1 + len(p.get("cont_pages") or []) for p in pages), "eip_pos": len(by_eip), "files": len(out),
               "absent": sum(1 for b in out if not b["in_pdf"] and not b["skipped"]), "skipped": sum(1 for b in out if b["skipped"]),
               "items": len(todo), "ok": sum(1 for r in todo if r["status"] == "ok"), "partial": sum(1 for r in todo if r["status"] == "partial"),
               "none": sum(1 for r in todo if r["status"] == "none"), "over": sum(1 for r in todo if r["status"] == "over"),
               "short_cases": _num(round(sum(r["short"] for r in todo if r["status"] in ("partial", "none")), 4)),
               "changes": sum(1 for r in todo if not r["locked"] and r["action"].startswith(("下修", "刪單", "取消"))),
               "locked": sum(1 for r in todo if r["locked"]), "warns": sum(1 for r in todo if r["warn"]),
               "expiry_short": sum(1 for r in todo if R_EXP in r["reasons"])}
    page_warnings = [w for p in pages for w in (p.get("warnings") or []) if "讀不到倉別" not in w]   # 倉別本來就看系統裡 EIP 的
    return {"eips": out, "unknown": unknown, "no_eip": no_eip, "page_warnings": page_warnings, "dates": dates,
            "warehouses": whs, "summary": summary, "skip_eips": sorted(skip)}


def apply(conn, result, operator):
    """把比對結果寫進缺貨紀錄。回傳寫了幾筆（新增或更新）、取消幾筆。"""
    stamp, written, cleared = now(), 0, 0
    for b in result["eips"]:
        if b["skipped"]:
            continue
        for r in b["items"]:
            if r["locked"] or r["status"] == "over":
                continue
            old = _row(conn.execute("SELECT * FROM shp_short WHERE order_id = ? AND round = 1 AND (notify_id IS NULL OR notify_id = 0)",
                                    (r["order_id"],)))
            need = r["final"] is not None and r["final"] < r["qty"]
            if not need:
                if old:
                    conn.execute("DELETE FROM shp_short WHERE id = ?", (old["id"],))
                    _log(conn, MARS, r["order_no"], r["shopee_sku_id"], "", "shopee_short", "第 1 次下修", f"可出貨 {old['qty_after']}",
                         "取消（勇信重新核對無缺貨）", operator, "manual", r["supplier_sku_id"])
                    cleared += 1
                continue
            note = (f"勇信 {b['eip_po']} 出 {r['shipped']} 箱，訂 {r['ordered']} 箱" if b["in_pdf"]
                    else f"勇信配送明細表上沒有 {b['eip_po']}，視為整份未出貨")
            if R_EXP in r["reasons"]:
                note += f"；效期早於 {r['need']} 的 {_num(min(r['shipped'], r['ordered']) - r['good'])} 箱不算（{r['need_rule']}）"
            reasons = "、".join(r["reasons"]) or R_STOCK
            if old and (old["qty_before"], old["qty_after"], old["note"], old["reasons"]) == (r["qty"], r["final"], note, reasons):
                continue
            if old:
                conn.execute("UPDATE shp_short SET source = ?, qty_before = ?, qty_after = ?, reasons = ?, note = ?, operator = ?, updated_at = ? WHERE id = ?",
                             (SOURCE, r["qty"], r["final"], reasons, note, operator, stamp, old["id"]))
            else:
                conn.execute("INSERT INTO shp_short (order_id, round, source, qty_before, qty_after, reasons, note, operator, updated_at) VALUES (?,?,?,?,?,?,?,?,?)",
                             (r["order_id"], 1, SOURCE, r["qty"], r["final"], reasons, note, operator, stamp))
            _log(conn, MARS, r["order_no"], r["shopee_sku_id"], "", "shopee_short", "第 1 次下修", f"可出貨 {r['qty']}",
                 f"可出貨 {r['final']}（缺 {r['qty'] - r['final']}）", operator, "manual", f"{r['supplier_sku_id']}；{SOURCE}；{note}", reasons)
            written += 1
    s = result["summary"]
    _log(conn, MARS, "", "", "", "shopee_yx", "勇信核對缺貨", "", f"{s['files']} 份 EIP、寫入 {written} 筆、取消 {cleared} 筆", operator, "manual",
         "、".join(result.get("warehouses") or []) + " " + "、".join(result.get("dates") or []))
    return written, cleared


def _pages_from_request():
    """上傳 PDF（第一次）或送回上次讀好的頁面（改勾選、確認時）。回傳 (pages, skip, files, 錯誤)。"""
    files = [f for f in request.files.getlist("file") if f and f.filename]
    if files:
        pages, names = [], []
        for f in files:
            try:
                ps = parse_yx_pdf(f.read(), f.filename)
            except ValueError as exc:
                return None, None, None, str(exc)
            pages += ps
            names.append({"filename": f.filename, "pages": sum(1 + len(p["cont_pages"]) for p in ps)})
        skip = [x for x in re.split(r"[,\s]+", request.form.get("skip_eips", "")) if x]
        return pages, skip, names, ""
    p = request.get_json(silent=True) or {}
    pages = p.get("pages")
    if not isinstance(pages, list) or not pages:
        return None, None, None, "請上傳勇信配送明細表（PDF）。"
    return pages, p.get("skip_eips") or [], p.get("files") or [], ""


@shopee_bp.route("/api/shopee/yx/compare", methods=["POST"])
def api_yx_compare():
    pages, skip, names, err = _pages_from_request()
    if err:
        return jsonify({"error": err}), 400
    conn = get_conn()
    try:
        if conn.execute("SELECT COUNT(*) AS n FROM shp_eip WHERE line = ? AND eip_po != ''", (MARS,)).fetchone()["n"] == 0:
            return jsonify({"error": "目前沒有已回填 EIP 採購單號的瑪氏 EIP 採購單，無法比對。請先至「採購單產出」回填。"}), 400
        result = compare(conn, pages, skip)
    finally:
        conn.close()
    return jsonify({**result, "pages": pages, "files": names})


@shopee_bp.route("/api/shopee/yx/apply", methods=["POST"])
def api_yx_apply():
    pages, skip, _names, err = _pages_from_request()
    if err:
        return jsonify({"error": err}), 400
    operator = _operator()
    conn = get_conn()
    try:
        result = compare(conn, pages, skip)                 # 確認時用伺服器當下的資料重算，不信畫面送回來的結果
        written, cleared = apply(conn, result, operator)
        conn.commit()
    finally:
        conn.close()
    return jsonify({"ok": True, "written": written, "cleared": cleared, "summary": result["summary"]})


__all__ = ["compare", "apply", "api_yx_compare", "api_yx_apply"]

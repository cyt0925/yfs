"""⑤ 勇信缺貨：讀勇信物流的「配送明細表(預排鮮度)」PDF，對回已填 EIP 採購單號的拆單表，逐料號比箱數，
確認後把系統裡的訂單出貨數量改成實際要出的，並產酷澎的「下修上傳檔」給小真上傳。

規則（Alice 的 html 第三節 ＋ Jerry 2026-10-01）：
- 第一層：PDF 每頁右上「收貨單號：PO202609065-酷澎-T」，前面那段就是 EIP 採購單號 → 找到那份拆單表。
- 第二層：產品編號＝58880＋瑪氏貨號；同一頁同商品分兩批效期會列兩行，加總。舊貨號靠商品總表備註的換號紀錄
  （10254053>60019810>60023881）對回現在的貨號。
- 第三層：配送數量「34C」的 C 是箱；比我們訂的箱數。相等＝無缺貨、少＝部分缺貨、沒出現或 0＝全數缺貨、
  多＝數量異常（不進下修檔，要人看）。
- 整張酷澎 PO 全缺（那張 PO 的每一份都在表裡、每個品項都是 0）：下修檔每個品項都列，第一支變更數量 1、其他 0，
  評論「{酷澎PO}此單不出，如有不便，請見諒。」；系統裡一律改 0（畫面提醒）。
- 下修檔變更數量＝實際要出的數量＝勇信箱數×箱入數（酷澎的單位），請求原因固定「4. 供應商庫存不足」，只列有缺的品項。
  套酷澎給的範本 purchase_templates/coupang_downgrade.xlsx，第 1 列那串代碼不能動。
- 範圍（主管 2026-10-08）：PDF 的指送日期＋倉別，系統裡那天那個倉的每一張酷澎 PO 都是「這次要出的訂單」，預設全部核對，
  **核不到就是缺**（整張缺的單勇信不會印在表上）。人可以取消勾選這次不出的 PO（skip_pos）。別的倉、別的日子不判缺貨。
  還沒填 EIP 單號的那份對不到 PDF，不當成缺，列出來要人先補。那張 PO 的每一份都是 0（而且沒有沒填號碼的）才算整張不出。
- 確認後：系統訂單出貨數量改成實際出貨量（記歷程，原因「缺貨」）、已送 EIP 的拆單快照一起改，之後 EMMA 檔、採購單都跟著變。
"""
import datetime as _dt
import os

import pymupdf

from master.orders_api import _apply_item_edit

from .common import *  # noqa: F401,F403
from .products import load_products

TEMPLATE = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "purchase_templates", "coupang_downgrade.xlsx")
REASON_TEXT = "4. 供應商庫存不足"
PO_RE = re.compile(r"收貨單號[：:]\s*(PO\d{9})")
CODE_RE = re.compile(r"^5888(\d{8,9})\s*$")
QTY_RE = re.compile(r"^(\d+)\s*C\s*$")
EXP_RE = re.compile(r"^(\d{4})/(\d{2})/(\d{2})\s*$")
TOTAL_RE = re.compile(r"合計箱數[：:]\s*(\d+)")
WH_RE = re.compile(r"永豐商店\s*-?\s*酷澎\s*-\s*([A-Za-z0-9]+)")      # TAO3 寫「永豐商店酷澎-TAO3」，TAO5 寫「永豐商店-酷澎-TAO5」
SHIP_RE = re.compile(r"指送日期[：:]\s*(\d{4})/(\d{2})/(\d{2})")


def _mars_code(code_tail):
    """5888060019810 → 60019810、5888010266398 → 10266398（58880＋8 碼）。"""
    return code_tail[1:] if len(code_tail) == 9 and code_tail.startswith("0") else code_tail


def _read_items(lines, items):
    """把一頁的文字行接到 items 後面。cur 從上一頁最後一個品項接著讀（品項跨頁時，續頁開頭可能還是上一個品項的箱數、效期）。"""
    cur = items[-1] if items else None
    for ln in lines:
        cm = CODE_RE.match(ln)
        if cm:
            cur = {"code": _mars_code(cm.group(1)), "raw_code": ln, "qty": None, "expiry": "", "name": ""}
            items.append(cur); continue
        if cur is None:
            continue
        qm = QTY_RE.match(ln)
        if qm:
            cur["qty"] = (cur["qty"] or 0) + int(qm.group(1)); continue      # 同商品分兩批效期：同一格裡會有兩個 5C／4C，加總
        em = EXP_RE.match(ln)
        if em:
            d = f"{em.group(1)}-{em.group(2)}-{em.group(3)}"
            cur["expiry"] = f"{cur['expiry']}、{d}" if cur["expiry"] else d; continue
        if ln and not ln.isdigit() and not cur["name"] and not ln.startswith("備註") and cur["qty"] is None:
            cur["name"] = re.sub(r"\s*\d+:\d+:\d+\s*$", "", ln)
    return items


def parse_yx_pdf(data, filename=""):
    """一頁一張配送單 → [{eip_po, warehouse, ship_date, items:[{code, qty, expiry, name}], total, page, warnings}]
    一張配送單品項多時會印成兩頁（「1 / 2 頁」）：第二頁沒有收貨單號，品項和「合計箱數」接回上一頁那張（TAO5 第 9～10 頁）。"""
    try:
        doc = pymupdf.open(stream=data, filetype="pdf")
    except Exception as exc:  # noqa: BLE001
        raise ValueError(f"{filename}：無法開啟此 PDF 檔。") from exc
    pages = []
    for pno, page in enumerate(doc, start=1):
        text = page.get_text()
        lines = [ln.strip() for ln in text.splitlines()]
        m = PO_RE.search(text)
        tot = TOTAL_RE.search(text)
        if not m:
            prev = pages[-1] if pages else None
            if prev and prev["eip_po"] and prev["total"] is None and prev["_last"] == pno - 1:
                _read_items(lines, prev["items"])                      # 續頁：接回上一張
                prev["total"] = int(tot.group(1)) if tot else None
                prev["cont_pages"].append(pno); prev["_last"] = pno
                continue
            if "勇信" in text or "配送明細表" in text or "配 送 明 細 表" in text:
                pages.append({"page": pno, "eip_po": "", "items": [], "total": None, "warehouse": "", "ship_date": "",
                              "warnings": [f"第 {pno} 頁找不到收貨單號"], "cont_pages": [], "_last": pno})
            continue
        wh = WH_RE.search(text); sd = SHIP_RE.search(text)
        pages.append({"page": pno, "eip_po": m.group(1), "warehouse": wh.group(1).upper() if wh else "",
                      "ship_date": f"{sd.group(1)}-{sd.group(2)}-{sd.group(3)}" if sd else "",
                      "items": _read_items(lines, []), "total": int(tot.group(1)) if tot else None, "warnings": [],
                      "cont_pages": [], "_last": pno})
    for p in pages:
        if not p["eip_po"]:
            continue
        label = f"第 {p['page']}" + (f"～{p['cont_pages'][-1]}" if p["cont_pages"] else "") + " 頁"
        p["warnings"] += [f"{label} {it['raw_code']} 未讀取到配送數量" for it in p["items"] if it["qty"] is None]
        got = sum(it["qty"] or 0 for it in p["items"])
        if p["total"] is None:
            p["warnings"].append(f"{label}找不到合計箱數，請人工核對明細是否完整")
        elif got != p["total"]:
            p["warnings"].append(f"{label}明細加總 {got} 箱與合計箱數 {p['total']} 箱不符，請人工核對")
        if not p["warehouse"]:
            p["warnings"].append(f"{label}讀不到倉別，改用系統中 {p['eip_po']} 的倉別")
    for p in pages:
        p.pop("_last", None)
    if not pages:
        raise ValueError(f"{filename}：此 PDF 不是勇信配送明細表（找不到收貨單號）；掃描圖檔的 PDF 無法讀取。")
    return pages


def _code_canon(conn):
    """任何出現過的瑪氏貨號（含備註裡的舊貨號）→ 商品總表現在用的貨號。"""
    canon = {}
    for r in _rows(conn.execute("SELECT DISTINCT mars_code, note FROM mst_mars_products WHERE mars_code != ''")):
        canon.setdefault(r["mars_code"], r["mars_code"])
        for old in re.findall(r"\d{8}", r["note"] or ""):
            canon.setdefault(old, r["mars_code"])
    return canon


def _orig_qty(v):
    """歷程的舊值是文字（"32"、"32.0"），轉回數字；空的就當沒有。"""
    try:
        return float(v) if v not in (None, "") else None
    except (TypeError, ValueError):
        return None


def _item_rows(s, items, canon, orig, take):
    """一份拆單的品項 → 比對列。take(貨號, 第幾個品項, 訂購箱數) 回傳勇信出了幾箱。"""
    rows = []
    for ii, it in enumerate(items):
        c = canon.get(it.get("mars_code") or "", it.get("mars_code") or "")
        box = it.get("box_file") or 0
        qty0 = _orig_qty(orig.get((s["po_number"], it["sku_id"])))
        ordered = round(qty0 / box, 4) if qty0 is not None and box else (it.get("cases") or 0)
        shipped = take(c, ii, ordered)
        if shipped == ordered:
            status = "ok"
        elif shipped == 0:
            status = "none"
        elif shipped < ordered:
            status = "partial"
        else:
            status = "over"
        rows.append({"sku_id": it["sku_id"], "yf_sku": it.get("yf_sku"), "purchase_code": it.get("purchase_code"), "mars_code": c,
                     "product_name": it.get("product_name"), "unit": it.get("unit"), "box_file": box,
                     "qty_ship": qty0 if qty0 is not None else it.get("qty_ship"), "ordered": ordered, "shipped": shipped,
                     "short": max(ordered - shipped, 0), "status": status, "new_qty": int(round(shipped * box)) if box else None})
    return rows


def _entry(s, eip, pages, rows, not_in_pdf=False):
    return {"split_id": s["id"], "eip_po": eip, "filename": s["filename"], "warehouse": s["warehouse"],
            "delivery_date": s["delivery_date"], "category": s["category"], "unit": s["unit"], "label": s["label"],
            "pages": pages, "items": rows, "extra": [], "not_in_pdf": not_in_pdf}


def compare(conn, pages, skip_pos=()):
    """把 PDF 幾頁對回拆單表。回傳給畫面看的結果；apply 再把這份結果送回來執行。

    範圍（主管 2026-10-08）：PDF 的指送日期＋倉別，系統裡那天那個倉的每一張酷澎 PO 都是「這次要出的訂單」，預設全部核對，
    **核不到就是缺**——整張缺的單勇信根本不會印在這份表上（早上就說不出的那批下午的表也不會有），不能等人記得去勾。
    人可以取消勾選「這次不出」的 PO（skip_pos），那張就整張不核、不改、不進下修檔。
    還沒填 EIP 單號的那份對不到 PDF，不能當成缺（Jerry 2026-10-08 TAO3：號碼沒填齊，其實有出），列出來要人先補。"""
    canon = _code_canon(conn)
    skip = {str(x) for x in skip_pos}
    # 重跑也要看得到：已經因勇信缺貨改成 0 的品項，訂購量用改之前的（看歷程裡第一次「缺貨／勇信出…」那筆的舊值），
    # 不然按過確認再跑一次，全部變「沒缺」、下修檔就少了那些（Jerry 2026-10-08 第二次跑拿到「沒有缺貨品項」）
    orig = {}
    for r in _rows(conn.execute("SELECT po_number, sku_id, old_value FROM mst_logs WHERE field = 'qty_ship' AND reason = '缺貨' AND note LIKE ? ORDER BY id", ("勇信出%",))):
        orig.setdefault((r["po_number"], r["sku_id"]), r["old_value"])
    by_eip = {}
    for p in pages:
        if not p["eip_po"]:
            continue
        d = by_eip.setdefault(p["eip_po"], {"qty": collections.Counter(), "pages": [], "names": {}})
        d["pages"].append(p["page"])
        for it in p["items"]:
            c = canon.get(it["code"], it["code"])
            d["qty"][c] += it["qty"] or 0
            d["names"].setdefault(c, it["name"])
    all_splits = _rows(conn.execute("SELECT * FROM mst_mars_splits ORDER BY id"))
    # 一個 EIP 採購單號可能對到好幾份拆單（盒、包同一張 EIP 採購單，Shanin 2026-10-07）
    eip_splits = collections.OrderedDict()
    for x in all_splits:
        if x["eip_po"]:
            eip_splits.setdefault(x["eip_po"], []).append(x)
    unknown = [{"eip_po": k, "pages": v["pages"], "cases": sum(v["qty"].values())} for k, v in by_eip.items() if k not in eip_splits]
    # PDF 讀不到倉別時用系統裡那個 EIP 單號的倉別（Jerry 2026-10-08 TAO5）
    whs = {p["warehouse"] for p in pages if p["warehouse"]} | {eip_splits[p["eip_po"]][0]["warehouse"] for p in pages if not p["warehouse"] and p["eip_po"] in eip_splits}
    dates = {p["ship_date"] for p in pages if p["ship_date"]}

    # PDF 上有的號碼：同一個號碼的幾份合起來比，勇信出的箱數照順序分給各份的同一個貨號，最後一個拿剩下的（多出來的才看得到超量）
    found = {}           # 拆單 id → 比對結果
    for eip, d in by_eip.items():
        group = eip_splits.get(eip)
        if not group:
            continue
        loaded = [(x, json.loads(x["items_json"] or "[]")) for x in group]
        last_of = {}
        for gi, (_, its) in enumerate(loaded):
            for ii, it in enumerate(its):
                last_of[canon.get(it.get("mars_code") or "", it.get("mars_code") or "")] = (gi, ii)
        left = collections.Counter(d["qty"]); seen = set(); entries = []
        for gi, (x, items) in enumerate(loaded):
            def take(c, ii, ordered, gi=gi):
                seen.add(c)
                got = left.get(c, 0) if last_of.get(c) == (gi, ii) else min(left.get(c, 0), ordered)
                left[c] -= got
                return got
            e = _entry(x, eip, d["pages"], _item_rows(x, items, canon, orig, take))
            entries.append(e); found[x["id"]] = e
        # 勇信有出、但這個號碼的拆單都沒訂的貨號，掛在第一份
        entries[0]["extra"] = [{"mars_code": c, "name": d["names"].get(c, ""), "cases": q} for c, q in d["qty"].items() if c not in seen and q]

    # 這次要核的訂單：PDF 那天那個倉的每一張酷澎 PO，加上 PDF 上號碼對到的
    in_scope = lambda x: x["warehouse"] in whs and x["delivery_date"] in dates   # noqa: E731
    po_splits = collections.OrderedDict()
    for x in all_splits:
        if in_scope(x) or x["id"] in found:
            po_splits.setdefault(x["po_number"], []).append(x)
    pos, scope, no_eip_all = [], [], []
    for po, sp in po_splits.items():
        in_pdf = [x for x in sp if x["id"] in found]
        no_eip = [x for x in sp if x["id"] not in found and not x["eip_po"]]
        absent = [x for x in sp if x["id"] not in found and x["eip_po"]]
        scope.append({"po_number": po, "warehouse": sp[0]["warehouse"], "delivery_date": sp[0]["delivery_date"], "files": len(sp),
                      "cases": round(sum(x["cases_total"] or 0 for x in sp), 2), "in_pdf": len(in_pdf), "not_in_pdf": len(absent),
                      "no_eip": len(no_eip), "eip_pos": sorted({x["eip_po"] for x in sp if x["eip_po"]}), "skipped": po in skip})
        if po in skip:
            continue
        no_eip_all += [{"po_number": po, "filename": x["filename"]} for x in no_eip]
        entries = [found[x["id"]] for x in in_pdf]
        for x in absent:                 # 有號碼、勇信表上沒有 → 核不到就是缺（0 箱）
            entries.append(_entry(x, x["eip_po"], [], _item_rows(x, json.loads(x["items_json"] or "[]"), canon, orig, lambda c, ii, o: 0), True))
        if not entries:
            continue
        every_zero = all(r["shipped"] == 0 for y in entries for r in y["items"]) and any(r["ordered"] for y in entries for r in y["items"])
        elsewhere = [x for x in all_splits if x["po_number"] == po and x["id"] not in {y["id"] for y in sp}]   # 同一張 PO 別天、別倉的份
        full = every_zero and not no_eip and not elsewhere
        pos.append({"po_number": po, "warehouse": sp[0]["warehouse"], "delivery_date": sp[0]["delivery_date"], "full": full,
                    "full_note": "整張 PO 全部沒出" if full else ("所有品項出貨皆為 0，但此 PO 尚有拆單還沒填 EIP 單號，暫以部分缺貨處理" if every_zero and no_eip
                                                              else ("所有品項出貨皆為 0，但此 PO 另有別天或別倉的拆單，暫以部分缺貨處理" if every_zero and elsewhere else "")),
                    "missing_splits": [x["filename"] + "（還沒填 EIP 單號）" for x in no_eip], "splits": entries})
    pos.sort(key=lambda p: (p["delivery_date"], p["po_number"]))
    scope.sort(key=lambda p: (p["delivery_date"], p["warehouse"], p["po_number"]))
    rows_all = [r for p in pos for y in p["splits"] for r in y["items"]]
    summary = {"pages": sum(1 + len(p.get("cont_pages") or []) for p in pages), "eip_pos": len(by_eip), "pos": len(pos), "items": len(rows_all),
               "ok": sum(1 for r in rows_all if r["status"] == "ok"), "partial": sum(1 for r in rows_all if r["status"] == "partial"),
               "none": sum(1 for r in rows_all if r["status"] == "none"), "over": sum(1 for r in rows_all if r["status"] == "over"),
               "full_pos": sum(1 for p in pos if p["full"]), "short_cases": sum(r["short"] for r in rows_all),
               "scope_pos": len(scope), "skipped_pos": sum(1 for p in scope if p["skipped"]),
               "absent_files": sum(1 for p in pos for y in p["splits"] if y["not_in_pdf"]), "no_eip_files": len(no_eip_all)}
    return {"pos": pos, "unknown": unknown, "scope": scope, "skip_pos": sorted(skip), "no_eip": no_eip_all,
            "page_warnings": [w for p in pages for w in p["warnings"]],
            "warehouses": sorted(whs), "ship_dates": sorted(dates), "summary": summary}


def downgrade_rows(result):
    """下修檔的列：有缺的品項；全缺的 PO 每個品項都列，第一支 1 其他 0、評論寫此單不出。數量異常不列。"""
    out = []
    for p in result["pos"]:
        if p["full"]:
            first = True
            for y in p["splits"]:
                for r in y["items"]:
                    out.append({"sku_id": r["sku_id"], "po_number": p["po_number"], "qty": 1 if first else 0, "reason": REASON_TEXT,
                                "comment": f"{p['po_number']}此單不出，如有不便，請見諒。"})
                    first = False
        else:
            for y in p["splits"]:
                for r in y["items"]:
                    if r["status"] in ("partial", "none") and r["new_qty"] is not None:
                        out.append({"sku_id": r["sku_id"], "po_number": p["po_number"], "qty": r["new_qty"], "reason": REASON_TEXT, "comment": ""})
    return out


def downgrade_file(rows):
    wb = openpyxl.load_workbook(TEMPLATE)
    ws = wb["issueReport"]
    if ws.max_row >= 3:
        ws.delete_rows(3, ws.max_row - 2)                      # 範本第 3 列是範例，拿掉；第 1、2 列不能動
    for i, r in enumerate(rows, start=3):
        vals = [r["sku_id"], r["po_number"], str(r["qty"]), r["reason"], "", r["comment"]]
        for col, v in enumerate(vals, start=1):
            c = ws.cell(i, col); c.value = v if v != "" else None; c.number_format = "@"
    buf = io.BytesIO(); wb.save(buf)
    return buf.getvalue()


def downgrade_filename(result):
    ds = result.get("ship_dates") or []
    d = "、".join(f"{int(x[5:7]):02d}{int(x[8:10]):02d}" for x in ds) if ds else _dt.date.today().strftime("%m%d")
    whs = "、".join(result.get("warehouses") or []) or "倉"
    return f"酷澎下修_{d}交貨-{whs}.xlsx"


def apply_changes(conn, result, operator):
    """把比對結果套進系統：訂單出貨數量改成實際出貨量（全缺 → 0），拆單快照跟著改。回傳 (改了幾筆, 說明)。"""
    changed, notes = 0, []
    for p in result["pos"]:
        for y in p["splits"]:
            s = _row(conn.execute("SELECT * FROM mst_mars_splits WHERE id = ?", (y["split_id"],)))
            if s is None:
                continue
            snap = json.loads(s["items_json"] or "[]")
            touched = False
            for r in y["items"]:
                if r["status"] == "over":
                    continue                                    # 出得比訂得多：不動，要人看
                new_qty = 0 if p["full"] else r["new_qty"]
                if new_qty is None or r["status"] == "ok" and not p["full"]:
                    continue
                o = _row(conn.execute("SELECT * FROM mst_orders WHERE po_number = ? AND sku_id = ?", (p["po_number"], r["sku_id"])))
                if o is None:
                    notes.append(f"{p['po_number']} {r['sku_id']}：系統裡找不到這筆訂單，沒改"); continue
                if not _same(o["qty_ship"], new_qty):
                    note = f"勇信出 {r['shipped']} 箱，訂 {r['ordered']} 箱" + ("；整張 PO 不出" if p["full"] else "")
                    sets, vals, n = _apply_item_edit(conn, o, {"qty_ship": new_qty}, operator, reason=("缺貨", note))
                    if n:
                        sets += ["updated_at = ?", "version = version + 1"]; vals.append(now())
                        conn.execute(f"UPDATE mst_orders SET {', '.join(sets)} WHERE id = ?", vals + [o["id"]])
                        changed += 1
                for it in snap:
                    if it["sku_id"] == r["sku_id"] and not _same(it.get("qty_ship"), new_qty):
                        it["qty_ship"] = new_qty
                        it["cases"] = round(new_qty / it["box_file"], 4) if it.get("box_file") else it.get("cases")
                        touched = True
            if touched:
                conn.execute("UPDATE mst_mars_splits SET items_json = ?, cases_total = ?, updated_by = ?, updated_at = ? WHERE id = ?",
                             (json.dumps(snap, ensure_ascii=False), round(sum(i.get("cases") or 0 for i in snap), 4), operator, now(), s["id"]))
    _log(conn, LINE, "", "", "", "mars_shortage", "勇信缺貨比對", "",
         f"{result['summary']['pos']} 張 PO、改 {changed} 筆、全缺 {result['summary']['full_pos']} 張", operator, "manual",
         "、".join(result.get("warehouses") or []) + " " + "、".join(result.get("ship_dates") or []))
    return changed, notes


# ── 路由 ─────────────────────────────────────────────────────────────────────
@mars_bp.route("/api/mars/shortage/compare", methods=["POST"])
def api_shortage_compare():
    files = [f for f in request.files.getlist("file") if f and f.filename]
    if not files:
        return jsonify({"error": "沒有收到檔案。"}), 400
    pages, names = [], []
    for f in files:
        try:
            ps = parse_yx_pdf(f.read(), f.filename)
        except ValueError as exc:
            return jsonify({"error": str(exc)}), 400
        pages += ps; names.append({"filename": f.filename, "pages": sum(1 + len(p["cont_pages"]) for p in ps)})
    conn = get_conn()
    try:
        if conn.execute("SELECT COUNT(*) AS n FROM mst_mars_splits WHERE eip_po != ''").fetchone()["n"] == 0:
            return jsonify({"error": "目前沒有拆單表填入 EIP 採購單號，無法比對。"}), 400
        skip_pos = [x.strip() for x in re.split(r"[,\s]+", request.form.get("skip_pos", "")) if x.strip()]
        result = compare(conn, pages, skip_pos)
    finally:
        conn.close()
    result["files"] = names
    result["downgrade_rows"] = len(downgrade_rows(result))
    return jsonify(result)


@mars_bp.route("/api/mars/shortage/apply", methods=["POST"])
def api_shortage_apply():
    """確認：改系統數量 ＋ 回傳下修檔。只改訂單數量不產檔用 only_apply；只要檔不改系統用 only_file。"""
    payload = request.get_json(silent=True) or {}
    result = payload.get("result") or {}
    if not isinstance(result.get("pos"), list):
        return jsonify({"error": "沒有比對結果，請重新上傳勇信配送明細表。"}), 400
    operator = _operator()
    rows = downgrade_rows(result)
    changed, notes = 0, []
    if not payload.get("only_file"):
        conn = get_conn()
        try:
            changed, notes = apply_changes(conn, result, operator)
            conn.commit()
        finally:
            conn.close()
    if payload.get("only_apply"):
        return jsonify({"ok": True, "changed": changed, "notes": notes})
    if not rows:
        return jsonify({"ok": True, "changed": changed, "notes": notes, "error": "沒有缺貨品項，不需酷澎下修檔。"}), 200
    resp = send_file(io.BytesIO(downgrade_file(rows)), as_attachment=True, download_name=downgrade_filename(result),
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    resp.headers["X-Mars-Changed"] = str(changed)
    resp.headers["X-Mars-Rows"] = str(len(rows))
    return resp


__all__ = ["parse_yx_pdf", "compare", "downgrade_rows", "downgrade_file", "downgrade_filename", "apply_changes",
           "api_shortage_compare", "api_shortage_apply"]

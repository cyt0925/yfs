"""④ 改單統計：每個月幾張 PO、幾張被改過、改了幾次、為什麼。主管要拿這個講「改單都是成本」。

名詞怎麼算（同事問起就照這裡講）：
- PO 歸哪個月：看它目前的交貨日（多個日期取最早那天）。沒排日期的另外一列「未排日期」。
- 被改過：這張 PO 在變動紀錄（mst_logs）裡有至少一次「改單事件」。
- 改單事件：同一張 PO、同一個來源、同一時間的變動算一次。因為整張 PO 改期會一個品項記一筆，
  一次動作記 20 筆不能算 20 次。我們改的看到「秒」（一次儲存同一秒寫完，兩次儲存分得開）；
  酷澎改的看到「分鐘」（匯入一批寫紀錄會跨秒）。
  - 酷澎改的（source=import）：整合表出貨數量／整合表交貨日變了、或品項從檔案消失。
    不看 qty_ship「隨整合表更新」那筆——它跟 qty_file_ship 是同一件事，算兩次會重複。
  - 我們改的（source=manual）：人手改出貨數量／交貨日。「恢復為整合表…」是把人改的撤掉，不算。
- 改了什麼：一次改單可能同時改數量和交期，所以「改數量」「改交期」「品項被拿掉」三類各自算次數，加起來會
  大於改單次數，這是正常的。改數量再分下修／上修（每個品項比新舊值；一次裡有下有上就算「有上有下」）。
- 原因：只有我們改的才有（缺貨／沒車／酷澎要求／其他）；2026-09-22 以前的紀錄沒有原因，列「未填」。
- 明細：每一次改單一列（時間、PO、誰、來源、改了什麼、原因），畫面上點任何數字都能看到它背後是哪幾次；
  Excel 也有「明細」分頁，主管想自己做樞紐分析就用那一頁。
- 測試：原因記成「測試」的改單（同事在正式站試功能時勾的）預設不算，統計頁「含測試」勾了才算；
  排除掉的次數另外回 test_events，畫面上寫一句「另有 N 次測試沒算」讓人知道有這回事。
- 估計工時：每次改單大約幾分鐘由主管給（畫面上填），系統只是乘出來，不是系統決定的。
資料從系統開始用（2026-09 中）才有，之前 Excel 時代的不在裡面。

2026-10-19 改版（Ivan 要追「原本下多少、最後出多少、為什麼改」，Jerry 看過示意圖）：
- 兩個系統合起來算：同一份訂單彙總表會匯進酷澎訂單管理系統（edit_logs，小真）和商品主檔自動化（mst_logs，Alice、妮可），
  兩邊都記「酷澎改了什麼」。每張 PO 每個品項、每個欄位（下單數量、訂單彙總表出貨數量、到貨日、倉別、品項移除）把兩邊看過的值
  照時間排成一條，值真的變了才算一次；誰先匯入就用誰看到的時間。所以同一個改動不會算兩次，只匯進其中一邊的也不會漏。
- 我方修改兩邊都算（各自是不同的動作）：商品主檔改出貨數量／到貨日、酷澎訂單管理改出貨數量／到貨日／倉別。
- 原因：每一次改單都要有原因，酷澎修改也要（很多是我們請酷澎改的）。原因只用大項（REASONS，草稿等 Alice、Ivan 定）；
  改的時候已經選了的（商品主檔、瑪氏）照舊對應過來；沒有的列「待補」，在改單統計頁勾好幾筆一起補，存在 mst_change_reasons。
- 原始下單→目前出貨用金額（數量 × 下單單價）：一張單有箱、盒、包，數量直接加起來沒有意思。
  原始＝兩邊最早看到的下單數量；目前＝兩邊最近更新的那筆出貨數量（品項移除算 0）。
- 少出拆兩段（Jerry 2026-10-19「這怎算的」→「拆成兩段」）：
  進系統前就少了＝第一次匯入時訂單彙總表的出貨數量就比下單數量少（在 Kate 的系統先改的，沒有改單紀錄、沒地方補原因）；
  進系統後才改＝之後的變動（有改單紀錄、有原因）。第一次的出貨數量＝兩邊出貨數量最早那筆紀錄改之前的值，沒紀錄就是現在的值。
- 酷澎拆單（原訂單欄有值、只有酷澎訂單管理有存）：原單被拿掉的品項如果在拆出來的新單裡，新單那項接回原單算，
  不算少出也不算新的下單（不然原單少出、新單又多一筆下單，少出會算太多）。兩張單都要在統計期間內才接。
"""
from .common import *  # noqa: F401,F403 — 共用工具、Flask、db、openpyxl 都從這裡來

UNDATED = "未排日期"
NO_REASON = "未填"
# 原因大項（Jerry 2026-10-19 草稿，等 Alice、Ivan 定案；定案後商品主檔改單視窗的四個也換成這份）
REASONS = ["酷澎倉限制", "原廠／產銷缺貨", "運能（沒車）", "酷澎要求", "我方作業", "其他"]
PENDING = "待補"
OLD_REASON = {"缺貨": "原廠／產銷缺貨", "沒車": "運能（沒車）", "酷澎要求": "酷澎要求", "其他": "其他"}
SYS_LABEL = {"c": "酷澎訂單管理", "m": "商品主檔"}
_RESTORE_PREFIXES = ("恢復為整合表", "恢復為訂單彙總表")   # 舊紀錄寫「整合表」，2026-10-07 起寫「訂單彙總表」


def _stats_range():
    month = request.args.get("month") or _this_month()
    month_to = request.args.get("month_to") or month
    if not _valid_month(month) or not _valid_month(month_to):
        return None, None, None
    months = _month_span(month, month_to)
    return months[0], months[-1], months


def _po_month(dates):
    real = sorted(d for d in dates if d)
    return real[0][:7] if real else UNDATED


KIND_LABEL = {"qty": "改數量", "date": "改到貨日", "wh": "改倉別", "gone": "品項移除"}
DIR_LABEL = {"down": "下修", "up": "上修", "mixed": "上修與下修", "": ""}
CANON_LABEL = {"order_qty": "下單數量", "file_ship": "訂單彙總表出貨數量", "ship": "出貨數量"}


def _num(v):
    try:
        return float(str(v).strip())
    except (TypeError, ValueError):
        return None


def _nv(canon, v):
    """比對用的值：數字 100／100.0 一樣、日期只看到日、倉別大寫。"""
    if v is None:
        return ""
    t = str(v).strip()
    if canon in ("order_qty", "file_ship", "ship"):
        n = _num(t)
        return "" if n is None else (str(int(n)) if n == int(n) else str(n))
    if canon == "date":
        return t[:10]
    if canon == "wh":
        return t.upper()
    return t


_KEPT = ("酷澎檔案仍是舊值", "整合表出貨數量已變")    # 人工調整過、匯入不覆蓋：檔案的新值記在 old_value


def _canon_import(sys, l):
    """一筆匯入紀錄 → (欄位, 檔案上的新值, 改之前的值)；不是要算的欄位回 None。sys：c＝酷澎訂單管理、m＝商品主檔。"""
    f, note = l["field"], l.get("note") or ""
    kept = note.startswith(_KEPT)
    old = "" if kept else l["old_value"]
    if sys == "m":
        if f == "qty_coupang":
            return "order_qty", l["new_value"], old
        if f == "qty_file_ship":
            return "file_ship", l["new_value"], old
        if f == "delivery_date_file":
            return "date", l["new_value"], old
        if f == "warehouse":
            return "wh", l["new_value"], old
        if f == "qty_ship" and "已沒有這個品項" in note:
            return "gone", "gone", l["old_value"]
        if f == "missing_in_file":
            return "gone", "back", ""
        return None
    if f == "qty_coupang":
        return "order_qty", l["new_value"], old
    if f == "delivery_date":
        return "date", (l["old_value"] if kept else l["new_value"]), old
    if f == "warehouse":
        return "wh", (l["old_value"] if kept else l["new_value"]), old
    if f == "qty_ship":
        if note.startswith("隨整合表出貨數量連動"):
            return "file_ship", l["new_value"], old
        if note.startswith("整合表出貨數量已變"):
            return "file_ship", l["old_value"], ""
        if note.startswith("酷澎後台已移除"):
            return "gone", "gone", l["old_value"]
        return None                                         # 隨下單數量連動：下單數量那筆已經算了
    if f == "removed_from_coupang":
        return "gone", "back", ""
    return None


def _new_event(key, po, source, sys, when, operator):
    return {"key": key, "po": po, "source": source, "sys": sys, "when": when, "operator": operator, "reason": "", "note": "",
            "kinds": set(), "qty_dir": set(), "_qty": collections.defaultdict(list), "_date": [], "_wh": [], "_gone": 0, "_back": 0,
            "_delta": {}, "delta": 0.0}


def _add_qty(e, canon, sku, old, new, price):
    e["kinds"].add("qty"); e["_qty"][canon].append((old, new))
    a, b = _num(old), _num(new)
    if a is not None and b is not None and a != b:
        e["qty_dir"].add("down" if b < a else "up")
        # 同一個品項同一次：下單數量優先，其次訂單彙總表出貨數量、出貨數量（不重複算）
        rank = {"order_qty": 0, "file_ship": 1, "ship": 2}[canon]
        cur = e["_delta"].get(sku)
        if cur is None or rank < cur[0]:
            e["_delta"][sku] = (rank, (b - a) * (price or 0))


def _finish_events(ev):
    for e in ev.values():
        e["qty_dir"] = ("mixed" if len(e["qty_dir"]) > 1 else next(iter(e["qty_dir"]), ""))
        parts = []
        for canon in ("order_qty", "file_ship", "ship"):
            pairs = e["_qty"].get(canon) or []
            if pairs:
                ex = "、".join(f"{_short(a)}→{_short(b)}" for a, b in pairs[:3]) + ("…" if len(pairs) > 3 else "")
                parts.append(f"{CANON_LABEL[canon]} {len(pairs)} 項{'（' + DIR_LABEL[e['qty_dir']] + '）' if e['qty_dir'] else ''}：{ex}")
        if e["_date"]:
            pairs = sorted({(a, b) for a, b in e["_date"]})
            parts.append(f"到貨日 {'、'.join(f'{_short(a)}→{_short(b)}' for a, b in pairs[:2])}（{len(e['_date'])} 項）")
        if e["_wh"]:
            pairs = sorted({(a or "—", b) for a, b in e["_wh"]})
            parts.append(f"倉別 {'、'.join(f'{a}→{b}' for a, b in pairs[:2])}（{len(e['_wh'])} 項）")
        if e["_gone"]:
            parts.append(f"品項移除 {e['_gone']} 項")
        if e["_back"]:
            parts.append(f"品項重新出現 {e['_back']} 項")
        e["summary"] = "；".join(parts)
        e["kind_labels"] = "、".join(KIND_LABEL[k] for k in ("qty", "date", "wh", "gone") if k in e["kinds"])
        e["delta"] = round(e["delta"] + sum(v for _r, v in e["_delta"].values()), 2)
        for k in ("_qty", "_date", "_wh", "_gone", "_back", "_delta"):
            e.pop(k)
    return ev


def _import_events(logs, prices):
    """兩個系統的匯入紀錄 → 酷澎修改的改單事件（同一個改動只算一次，見檔頭）。另外回原始下單數量 {(po, sku): 值}。"""
    obs = collections.defaultdict(list)
    for sys, l in logs:
        c = _canon_import(sys, l)
        if c:
            canon, val, old = c
            obs[(l["po_number"], l["sku_id"], canon)].append((l["changed_at"] or "", sys, _nv(canon, val), _nv(canon, old), l))
    ev, first_qty = {}, {}
    for (po, sku, canon), pts in obs.items():
        pts.sort(key=lambda p: (p[0], p[1]))
        prev = None
        for t, sys, v, old, l in pts:
            if prev is not None and v == prev:
                continue                                   # 另一個系統後來才看到同一個值：不重複算
            before = prev if prev is not None else old
            if canon == "order_qty" and (po, sku) not in first_qty and before != "":
                first_qty[(po, sku)] = _num(before)
            key = f"import|{sys}|{po}|{t[:16]}"
            e = ev.setdefault(key, _new_event(key, po, "import", sys, t, l["operator"]))
            price = prices.get((po, sku))
            if canon in ("order_qty", "file_ship"):
                _add_qty(e, canon, sku, before, v, price)
            elif canon == "date":
                e["kinds"].add("date"); e["_date"].append((before, v))
            elif canon == "wh":
                e["kinds"].add("wh"); e["_wh"].append((before, v))
            elif v == "gone":
                e["kinds"].add("gone"); e["_gone"] += 1
                q = _num(old)
                if q and (po, sku) not in e["_delta"]:
                    e["delta"] -= q * (price or 0)
            else:
                e["_back"] += 1
            prev = v
    return ev, first_qty


def _manual_events(mlogs, clogs, prices):
    """我方修改：商品主檔（有原因）、酷澎訂單管理（沒有原因，要補）。人改的時間看到秒。"""
    ev = {}
    for sys, l in [("m", x) for x in mlogs] + [("c", x) for x in clogs]:
        stamp = l["changed_at"] or ""
        reason = (l.get("reason") or "") if sys == "m" else ""
        key = f"manual|{sys}|{l['po_number']}|{stamp[:19]}" + (f"|{reason}" if reason else "")
        e = ev.setdefault(key, _new_event(key, l["po_number"], "manual", sys, stamp, l["operator"]))
        if not e["reason"]:
            e["reason"] = reason
        note = l.get("note") or ""
        if note and note not in e["note"]:
            e["note"] = f"{e['note']}；{note}" if e["note"] else note
        f = l["field"]
        if f == "qty_ship":
            _add_qty(e, "ship", l["sku_id"], _nv("ship", l["old_value"]), _nv("ship", l["new_value"]), prices.get((l["po_number"], l["sku_id"])))
        elif f == "delivery_date":
            e["kinds"].add("date"); e["_date"].append((_nv("date", l["old_value"]), _nv("date", l["new_value"])))
        elif f == "warehouse":
            e["kinds"].add("wh"); e["_wh"].append((_nv("wh", l["old_value"]), _nv("wh", l["new_value"])))
    return ev


def _event_row(e, p):
    """給 API／Excel 用的一列。"""
    return {"key": e["key"], "when": e["when"], "po_number": e["po"], "month": p["month"], "lines": "／".join(sorted(p["lines"])),
            "source": e["source"], "source_label": "酷澎修改" if e["source"] == "import" else "我方修改",
            "system": SYS_LABEL[e["sys"]], "operator": e["operator"], "kinds": sorted(e["kinds"]), "kind_labels": e["kind_labels"],
            "qty_dir": e["qty_dir"], "summary": e["summary"], "delta": e["delta"],
            "reason": e["final_reason"], "reason_from": e["reason_from"], "pending": e["final_reason"] == PENDING, "note": e["note"]}


def _include_test():
    return request.args.get("include_test") in ("1", "true")


def _item_rows(conn):
    """兩個系統的品項（同一個 PO＋SKU 兩邊都有就兩筆）。"""
    out = []
    for r in _rows(conn.execute("SELECT po_number, sku_id, line, delivery_date, qty_coupang, qty_ship, qty_file_ship, unit_price, "
                                "missing_in_file AS gone, first_seen_at, updated_at FROM mst_orders")):
        r["sys"] = "m"; r["parent_po"] = ""; out.append(r)
    for r in _rows(conn.execute("SELECT po_number, sku_id, line, delivery_date, qty_coupang, qty_ship, unit_price, parent_po, "
                                "removed_from_coupang AS gone, first_seen_at, updated_at FROM orders")):
        r["sys"] = "c"; r["qty_file_ship"] = None; out.append(r)
    return out


def _first_ship(conn, keep):
    """每個品項第一次看到的出貨數量：兩個系統出貨數量最早那筆紀錄「改之前」的值（沒紀錄的品項不在裡面＝從來沒變過）。"""
    logs = []
    for sys_, sql in (("m", "SELECT po_number, sku_id, field, old_value, new_value, note, changed_at FROM mst_logs "
                            "WHERE po_number != '' AND field IN ('qty_file_ship', 'qty_ship')"),
                      ("c", "SELECT po_number, sku_id, field, old_value, new_value, note, changed_at FROM edit_logs "
                            "WHERE po_number != '' AND field = 'qty_ship'")):
        logs += [(l["changed_at"] or "", sys_, l) for l in _rows(conn.execute(sql)) if l["po_number"] in keep]
    out = {}
    for t, _s, l in sorted(logs, key=lambda x: (x[0], x[1])):
        k = (l["po_number"], l["sku_id"])
        if k in out:
            continue
        before = l["new_value"] if (l["note"] or "").startswith(_KEPT) else l["old_value"]   # 人工調整過的那種：現在的值記在 new_value
        n = _num(before)
        if n is not None:
            out[k] = n
    return out


def _collect(conn, month_from, month_to, line_group, include_test=False):
    """範圍內的 PO（keep）與它們的改單事件（events）。統計頁和明細 API 共用。
    回傳的第三個值是被排除的「測試」次數（include_test=True 時為 0）。"""
    cfg = _line_groups()
    items = _item_rows(conn)
    pos, by_item, prices = {}, collections.defaultdict(list), {}
    for r in items:
        p = pos.setdefault(r["po_number"], {"dates": set(), "lines": set(), "skus": set()})
        p["dates"].add(r["delivery_date"] or ""); p["lines"].add(_group_of(r["line"], cfg)); p["skus"].add(r["sku_id"])
        by_item[(r["po_number"], r["sku_id"])].append(r)
        if r["unit_price"] is not None:
            prices.setdefault((r["po_number"], r["sku_id"]), _num(r["unit_price"]) or 0)
    span = set(_month_span(month_from, month_to))
    keep = {}
    for po, p in pos.items():
        m = _po_month(p["dates"])
        if m != UNDATED and m not in span:
            continue
        if line_group and line_group not in p["lines"]:
            continue
        p["month"] = m; p["items"] = len(p["skus"])
        keep[po] = p
    if not keep:
        return keep, {}, 0
    restore = tuple(f"{x}%" for x in _RESTORE_PREFIXES)
    mlogs = _rows(conn.execute(
        """SELECT po_number, sku_id, source, field, old_value, new_value, operator, note, reason, changed_at FROM mst_logs
           WHERE po_number != '' AND (
                 (source = 'import' AND field IN ('qty_coupang', 'qty_file_ship', 'delivery_date_file', 'warehouse', 'missing_in_file'))
              OR (source = 'import' AND field = 'qty_ship' AND note LIKE ?)
              OR (source = 'manual' AND field IN ('qty_ship', 'delivery_date') AND note NOT LIKE ? AND note NOT LIKE ?))""",
        ("%已沒有這個品項%",) + restore))
    clogs = _rows(conn.execute(
        """SELECT po_number, sku_id, source, field, old_value, new_value, operator, note, changed_at FROM edit_logs
           WHERE po_number != '' AND (
                 (source = 'import' AND field IN ('qty_coupang', 'delivery_date', 'warehouse', 'qty_ship', 'removed_from_coupang'))
              OR (source = 'manual' AND field IN ('qty_ship', 'delivery_date', 'warehouse') AND note NOT LIKE ?))""", ("恢復%",)))
    mlogs = [l for l in mlogs if l["po_number"] in keep]; clogs = [l for l in clogs if l["po_number"] in keep]
    imp, first_qty = _import_events([("m", l) for l in mlogs if l["source"] == "import"] + [("c", l) for l in clogs if l["source"] == "import"], prices)
    man = _manual_events([l for l in mlogs if l["source"] == "manual"], [l for l in clogs if l["source"] == "manual"], prices)
    events = _finish_events({**imp, **man})
    stored = {r["event_key"]: r for r in _rows(conn.execute("SELECT event_key, reason, operator, updated_at FROM mst_change_reasons"))}
    for e in events.values():
        s_ = stored.get(e["key"])
        if s_:
            e["final_reason"], e["reason_from"] = s_["reason"], f"{s_['operator']} {(s_['updated_at'] or '')[:10]} 補"
        elif e["reason"] == TEST_REASON:
            e["final_reason"], e["reason_from"] = TEST_REASON, "改的時候已選"
        elif e["reason"]:
            e["final_reason"], e["reason_from"] = OLD_REASON.get(e["reason"], e["reason"]), "改的時候已選"
        else:
            e["final_reason"], e["reason_from"] = PENDING, ""
    # 原始下單、第一次看到的出貨、目前出貨（數量）：每個品項看兩邊
    first_ship = _first_ship(conn, keep)
    amt = {}
    for po, p in keep.items():
        for sku in p["skus"]:
            rows = by_item[(po, sku)]
            orig = first_qty.get((po, sku))
            if orig is None:
                first = min(rows, key=lambda r: r["first_seen_at"] or "9")
                orig = _num(first["qty_coupang"]) or 0
            last = max(rows, key=lambda r: r["updated_at"] or "")
            cur = 0 if last["gone"] else _num(last["qty_ship"] if last["qty_ship"] is not None else (last["qty_file_ship"] if last["qty_file_ship"] is not None else last["qty_coupang"])) or 0
            fs = first_ship.get((po, sku))
            amt[(po, sku)] = {"price": prices.get((po, sku)) or 0, "orig": orig, "first": cur if fs is None else fs, "cur": cur,
                              "gone": bool(last["gone"]), "parent": next((r["parent_po"] for r in rows if r.get("parent_po")), ""), "moved": 0.0}
    # 酷澎拆單：原單拿掉的品項在新單裡 → 新單那項接回原單算（見檔頭）
    for (po, sku), a in amt.items():
        par = amt.get((a["parent"], sku)) if a["parent"] and a["parent"] != po else None
        if par is None or not par["gone"]:
            continue
        par["cur"] = a["cur"]; par["moved"] = a["orig"] * a["price"]
        a["orig"] = a["first"] = a["cur"] = 0
    for po, p in keep.items():
        p["orig_amount"] = p["cur_amount"] = p["pre_amount"] = p["post_amount"] = p["moved_amount"] = 0.0
        for sku in p["skus"]:
            a = amt[(po, sku)]
            p["orig_amount"] += a["orig"] * a["price"]; p["cur_amount"] += a["cur"] * a["price"]
            p["pre_amount"] += (a["first"] - a["orig"]) * a["price"]; p["post_amount"] += (a["cur"] - a["first"]) * a["price"]
            p["moved_amount"] += a["moved"]
    if include_test:
        return keep, events, 0
    tests = [k for k, e in events.items() if e["final_reason"] == TEST_REASON]
    for k in tests:
        events.pop(k)
    return keep, events, len(tests)


def _build_stats(conn, month_from, month_to, line_group, include_test=False):
    keep, events, n_test = _collect(conn, month_from, month_to, line_group, include_test)
    reasons = REASONS + ([TEST_REASON] if include_test else []) + [PENDING]
    if not keep:
        return {"months": [], "total": _finish_month(_empty_month("合計"), reasons), "reasons": reasons, "reason_choices": REASONS,
                "top_pos": [], "month_from": month_from, "month_to": month_to, "test_events": 0}

    per_po = {po: {"events": 0, "coupang": 0, "manual": 0, "reasons": collections.Counter(), "reason_amount": collections.Counter(),
                   "kinds": collections.Counter()} for po in keep}
    for e in events.values():
        s = per_po[e["po"]]
        s["events"] += 1
        side = "c" if e["source"] == "import" else "m"
        s["coupang" if side == "c" else "manual"] += 1
        s["reasons"][e["final_reason"]] += 1
        s["reason_amount"][e["final_reason"]] += e["delta"]
        for k in e["kinds"]:
            s["kinds"][f"{k}_{side}"] += 1
        if "qty" in e["kinds"] and e["qty_dir"]:
            s["kinds"][f"qty_{e['qty_dir']}"] += 1

    months = {}
    order = _month_span(month_from, month_to) + [UNDATED]

    def add(m, s, p):
        m["pos"] += 1; m["items"] += p["items"]; m["_skus"].update(p["skus"])
        for k in ("orig_amount", "cur_amount", "pre_amount", "post_amount", "moved_amount"):
            m[k] += p[k]
        if s["events"]:
            m["changed_pos"] += 1
        m["events"] += s["events"]; m["coupang"] += s["coupang"]; m["manual"] += s["manual"]
        for k, v in s["reasons"].items():
            m["reasons"][k] = m["reasons"].get(k, 0) + v
        for k, v in s["reason_amount"].items():
            m["reason_amount"][k] = m["reason_amount"].get(k, 0) + v
        for k, v in s["kinds"].items():
            m["kinds"][k] = m["kinds"].get(k, 0) + v
    total = _empty_month("合計")
    for po, p in keep.items():
        add(months.setdefault(p["month"], _empty_month(p["month"])), per_po[po], p)
        add(total, per_po[po], p)
    out_months = [_finish_month(months[m], reasons) for m in order if m in months]
    total = _finish_month(total, reasons)

    top = sorted(((po, keep[po], per_po[po]) for po in keep if per_po[po]["events"]),
                 key=lambda x: (-x[2]["events"], x[0]))[:15]
    top_pos = [{
        "po_number": po, "month": p["month"], "lines": "／".join(sorted(p["lines"])), "items": p["items"],
        "events": s["events"], "coupang": s["coupang"], "manual": s["manual"],
        "reasons": "、".join(f"{k}{v}" if v > 1 else k for k, v in s["reasons"].most_common()),
    } for po, p, s in top]
    return {"months": out_months, "total": total, "reasons": reasons, "reason_choices": REASONS,
            "top_pos": top_pos, "month_from": month_from, "month_to": month_to, "test_events": n_test}


def _event_rows(keep, events):
    rows = [_event_row(e, keep[e["po"]]) for e in events.values()]
    rows.sort(key=lambda r: (r["when"], r["po_number"]), reverse=True)
    return rows


KIND_KEYS = ["qty_c", "qty_m", "qty_down", "qty_up", "qty_mixed", "date_c", "date_m", "wh_c", "wh_m", "gone_c"]


def _empty_month(label):
    return {"month": label, "pos": 0, "items": 0, "_skus": set(), "changed_pos": 0,
            "events": 0, "coupang": 0, "manual": 0, "reasons": {}, "reason_amount": {}, "kinds": {},
            "orig_amount": 0.0, "cur_amount": 0.0, "pre_amount": 0.0, "post_amount": 0.0, "moved_amount": 0.0}


def _finish_month(m, reasons=None):
    m["skus"] = len(m.pop("_skus"))
    m["changed_pct"] = round(m["changed_pos"] * 100 / m["pos"]) if m["pos"] else 0
    m["per_changed_po"] = round(m["events"] / m["changed_pos"], 1) if m["changed_pos"] else 0
    keys = reasons or REASONS + [PENDING]
    m["reasons"] = {k: m["reasons"].get(k, 0) for k in keys}
    m["reason_amount"] = {k: round(m["reason_amount"].get(k, 0), 2) for k in keys}
    m["pending"] = m["reasons"].get(PENDING, 0)
    m["kinds"] = {k: m["kinds"].get(k, 0) for k in KIND_KEYS}
    for k in ("orig_amount", "cur_amount", "pre_amount", "post_amount", "moved_amount"):
        m[k] = round(m[k], 2)
    m["short_amount"] = round(m["cur_amount"] - m["orig_amount"], 2)
    m["short_pct"] = round((m["orig_amount"] - m["cur_amount"]) * 100 / m["orig_amount"], 1) if m["orig_amount"] else 0
    return m


@master_bp.route("/api/master/stats")
def api_stats():
    month_from, month_to, months = _stats_range()
    if not months:
        return jsonify({"error": "月份格式錯誤，請使用 2026-09 格式。"}), 400
    line = norm_text(request.args.get("line"))
    conn = get_conn()
    try:
        data = _build_stats(conn, month_from, month_to, line, _include_test())
    finally:
        conn.close()
    data["line"] = line; data["include_test"] = _include_test()
    return jsonify(data)


@master_bp.route("/api/master/stats/events")
def api_stats_events():
    """一次改單一列，畫面上點數字時撈的。篩選在前端做（一段期間的事件頂多幾百幾千筆）。"""
    month_from, month_to, months = _stats_range()
    if not months:
        return jsonify({"error": "月份格式錯誤，請使用 2026-09 格式。"}), 400
    line = norm_text(request.args.get("line"))
    conn = get_conn()
    try:
        keep, events, n_test = _collect(conn, month_from, month_to, line, _include_test())
    finally:
        conn.close()
    return jsonify({"events": _event_rows(keep, events), "month_from": month_from, "month_to": month_to, "test_events": n_test})


@master_bp.route("/api/master/stats/reasons", methods=["POST"])
def api_stats_reasons():
    """補原因：勾好幾次改單一起設同一個原因（Jerry 2026-10-19）。已經有原因的也可以改。"""
    payload = request.get_json(silent=True) or {}
    keys = [str(k) for k in (payload.get("keys") or []) if str(k).startswith(("import|", "manual|"))][:2000]
    reason = norm_text(payload.get("reason"))
    if not keys:
        return jsonify({"error": "沒有勾選任何改單。"}), 400
    if reason not in REASONS:
        return jsonify({"error": f"原因要從清單裡選：{'、'.join(REASONS)}。"}), 400
    operator, stamp = _operator(), now()
    conn = get_conn()
    try:
        for k in keys:
            conn.execute("""INSERT INTO mst_change_reasons (event_key, reason, operator, updated_at) VALUES (?, ?, ?, ?)
                            ON CONFLICT (event_key) DO UPDATE SET reason = EXCLUDED.reason, operator = EXCLUDED.operator,
                            updated_at = EXCLUDED.updated_at""", (k, reason, operator, stamp))
        _log(conn, "", "", "", "", "change_reason", "改單原因", "", f"{len(keys)} 次改單 → {reason}", operator, "manual", "")
        conn.commit()
    finally:
        conn.close()
    return jsonify({"ok": True, "count": len(keys), "reason": reason})


@master_bp.route("/api/master/stats/export")
def api_stats_export():
    month_from, month_to, months = _stats_range()
    if not months:
        return jsonify({"error": "月份格式錯誤，請使用 2026-09 格式。"}), 400
    line = norm_text(request.args.get("line"))
    minutes = norm_int(request.args.get("minutes")) or 0
    conn = get_conn()
    try:
        d = _build_stats(conn, month_from, month_to, line, _include_test())
        keep, events, n_test = _collect(conn, month_from, month_to, line, _include_test())
    finally:
        conn.close()
    from openpyxl.styles import Alignment, Font, PatternFill
    bold = Font(bold=True); head = PatternFill("solid", fgColor="DBEAFE"); tot = PatternFill("solid", fgColor="F1F5F9")
    wb = openpyxl.Workbook()

    ws = wb.active; ws.title = "每月統計"
    hdr = ["月份", "PO 張數", "品項數", "不同 SKU 數", "曾改單的 PO", "曾改單比例", "改單次數",
           "其中酷澎修改", "其中我方修改", "曾改單的 PO 平均改單次數", "待補原因", "原始下單金額", "目前出貨金額", "少出金額", "少出比例",
           "其中進系統前就少了", "其中進系統後才改", "拆單接回原單的金額"]
    if minutes:
        hdr.append(f"估計工時（每次 {minutes} 分鐘）")
    ws.append(hdr)
    for m in d["months"] + [d["total"]]:
        row = [m["month"], m["pos"], m["items"], m["skus"], m["changed_pos"], f"{m['changed_pct']}%",
               m["events"], m["coupang"], m["manual"], m["per_changed_po"], m["pending"],
               m["orig_amount"], m["cur_amount"], m["short_amount"], f"{m['short_pct']}%",
               m["pre_amount"], m["post_amount"], m["moved_amount"]]
        if minutes:
            row.append(round(m["events"] * minutes / 60, 1))
        ws.append(row)
    _style(ws, hdr, bold, head, tot, total_row=True)

    ws2 = wb.create_sheet("改單原因")
    hdr2 = ["月份", "改單次數"] + d["reasons"] + [f"{r} 金額差" for r in d["reasons"]]
    ws2.append(hdr2)
    for m in d["months"] + [d["total"]]:
        ws2.append([m["month"], m["events"]] + [m["reasons"].get(r, 0) for r in d["reasons"]] + [m["reason_amount"].get(r, 0) for r in d["reasons"]])
    _style(ws2, hdr2, bold, head, tot, total_row=True)

    wsk = wb.create_sheet("改了什麼")
    hdrk = ["月份", "改數量（酷澎）", "改數量（我方）", "其中下修", "其中上修", "其中上修與下修",
            "改到貨日（酷澎）", "改到貨日（我方）", "改倉別（酷澎）", "改倉別（我方）", "品項移除（酷澎）"]
    wsk.append(hdrk)
    for m in d["months"] + [d["total"]]:
        k = m["kinds"]
        wsk.append([m["month"]] + [k[x] for x in KIND_KEYS])
    _style(wsk, hdrk, bold, head, tot, total_row=True)

    ws3 = wb.create_sheet("改最多次的 PO")
    hdr3 = ["PO 單號", "月份", "線別", "品項數", "改單次數", "酷澎修改", "我方修改", "原因"]
    ws3.append(hdr3)
    for t in d["top_pos"]:
        ws3.append([t["po_number"], t["month"], t["lines"], t["items"], t["events"], t["coupang"], t["manual"], t["reasons"]])
    _style(ws3, hdr3, bold, head, tot, total_row=False)

    wsd = wb.create_sheet("明細")
    hdrd = ["時間", "月份", "PO 單號", "線別", "操作人員", "來源", "記在哪個系統", "改單類型", "修改內容", "金額差", "原因", "原因怎麼來的", "說明"]
    wsd.append(hdrd)
    for r in _event_rows(keep, events):
        wsd.append([r["when"], r["month"], r["po_number"], r["lines"], r["operator"], r["source_label"], r["system"],
                    r["kind_labels"], r["summary"], r["delta"], r["reason"], r["reason_from"], r["note"]])
    _style(wsd, hdrd, bold, head, tot, total_row=False)
    wsd.column_dimensions["I"].width = 60; wsd.column_dimensions["A"].width = 20

    ws4 = wb.create_sheet("怎麼算的")
    for line_text in [
        f"範圍：{month_from} ～ {month_to}" + (f"，線別 {line}" if line else "，全部線別")
        + ("，含勾了「這是測試」的改單" if _include_test() else f"，不含勾了「這是測試」的改單（這段期間有 {n_test} 次）"),
        "PO 歸屬月份：依目前的到貨日（多個日期取最早那天）。未排日期的另列「未排日期」。",
        "曾改單：這張 PO 在修改歷程裡至少有一次改單事件。",
        "改單事件：同一張 PO、同一來源、同一時間的變動算一次（整張 PO 改期會一個品項記一筆，一次動作不能算很多次）。",
        "兩個系統合起來算：訂單彙總表會匯進酷澎訂單管理系統和商品主檔自動化，兩邊都會記酷澎改了什麼；同一個改動只算一次，用先匯入的那邊的時間。",
        "酷澎修改：訂單彙總表上的下單數量、出貨數量、到貨日、倉別與上次不同，或品項從檔案中消失。",
        "我方修改：人工修改出貨數量、到貨日、倉別（兩個系統都算）；恢復成訂單彙總表數字／日期不算。",
        "改單類型：分改數量、改到貨日、改倉別、品項移除，一次改單可能同時屬於多類，各類加總會大於改單次數。改數量再分下修／上修。",
        f"原因：每一次改單都要有，酷澎修改也要。大項：{'、'.join(REASONS)}（草稿）。改的時候已選的照舊對應（缺貨→原廠／產銷缺貨、沒車→運能（沒車））；沒有的列「待補」，在改單統計頁補。",
        "金額：數量 × 下單單價（一張單有箱、盒、包，數量直接相加沒有意思）。原始下單＝兩個系統最早看到的下單數量；目前出貨＝最近更新那邊的出貨數量，品項移除算 0。金額差是那次改單讓出貨金額變多少。",
        "少出拆兩段：進系統前就少了＝第一次匯入時出貨數量就比下單數量少（在 Kate 系統先改的，沒有改單紀錄）；進系統後才改＝之後的變動（有改單紀錄、有原因）。",
        "酷澎拆單：原單拿掉的品項在拆出來的新單裡（原訂單欄），新單那項接回原單算，不算少出、也不算新的下單。",
        "明細分頁：一次改單一列，要自己做樞紐分析就用這一頁。",
        "估計工時：改單次數 × 每次幾分鐘 ÷ 60。每次幾分鐘是人填的，不是系統算的。",
        "資料從系統開始用（2026 年 9 月中）才有，之前 Excel 時代的紀錄不在裡面。",
    ]:
        ws4.append([line_text])
    ws4.column_dimensions["A"].width = 110
    for r in ws4.iter_rows():
        r[0].alignment = Alignment(wrap_text=True, vertical="top")

    label = month_from if month_from == month_to else f"{month_from}～{month_to}"
    return _xlsx_response(wb, f"改單統計_{label}.xlsx")


def _style(ws, hdr, bold, head, tot, total_row):
    from openpyxl.utils import get_column_letter
    for c in ws[1]:
        c.font = bold; c.fill = head
    if total_row and ws.max_row > 1:
        for c in ws[ws.max_row]:
            c.font = bold; c.fill = tot
    for i, h in enumerate(hdr, start=1):
        ws.column_dimensions[get_column_letter(i)].width = max(12, min(40, len(str(h)) * 2 + 4))
    ws.freeze_panes = "A2"


__all__ = ["api_stats", "api_stats_events", "api_stats_reasons", "api_stats_export", "_build_stats", "_collect", "_event_rows", "REASONS", "PENDING"]

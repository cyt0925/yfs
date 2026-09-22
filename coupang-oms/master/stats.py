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
"""
from .common import *  # noqa: F401,F403 — 共用工具、Flask、db、openpyxl 都從這裡來

UNDATED = "未排日期"
NO_REASON = "未填"
_RESTORE_PREFIX = "恢復為整合表"


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


KIND_LABEL = {"qty": "改數量", "date": "改交期", "gone": "品項被拿掉"}
DIR_LABEL = {"down": "下修", "up": "上修", "mixed": "有上有下", "": ""}


def _num(v):
    try:
        return float(str(v).strip())
    except (TypeError, ValueError):
        return None


def _events(logs):
    """把一筆一筆的紀錄併成「改單事件」：(po, source, 時間) 一個事件。人改的時間看到秒、
    匯入的看到分鐘（見檔頭）。每個事件帶：誰、什麼時候、改了哪幾類（qty／date／gone）、
    數量是下修還是上修、原因、說明，還有一句人看得懂的「改了什麼」。"""
    ev = {}
    for l in logs:
        stamp = l["changed_at"] or ""
        # 人改的多帶原因進 key：同一秒兩次儲存、原因不同就分成兩次（測試裡會發生，實務上不會少算）
        key = (l["po_number"], l["source"], stamp[:19], l.get("reason") or "") if l["source"] == "manual" \
            else (l["po_number"], l["source"], stamp[:16], "")
        e = ev.setdefault(key, {"po": l["po_number"], "source": l["source"], "when": stamp, "operator": l["operator"],
                                "reason": "", "note": "", "kinds": set(), "qty_dir": set(), "_qty": [], "_date": [], "_gone": 0})
        if stamp and stamp < e["when"]:
            e["when"] = stamp
        if l["source"] == "manual":
            if not e["reason"]:
                e["reason"] = l.get("reason") or ""
            note = l.get("note") or ""
            if note and note not in e["note"]:
                e["note"] = f"{e['note']}；{note}" if e["note"] else note
        f = l["field"]
        if f == "qty_ship" and l["source"] == "import":          # 品項從檔案消失
            e["kinds"].add("gone"); e["_gone"] += 1
        elif f in ("qty_ship", "qty_file_ship"):
            e["kinds"].add("qty"); e["_qty"].append((l["old_value"], l["new_value"]))
            a, b = _num(l["old_value"]), _num(l["new_value"])
            if a is not None and b is not None and a != b:
                e["qty_dir"].add("down" if b < a else "up")
        elif f in ("delivery_date", "delivery_date_file"):
            e["kinds"].add("date"); e["_date"].append((l["old_value"], l["new_value"]))
    for e in ev.values():
        e["qty_dir"] = ("mixed" if len(e["qty_dir"]) > 1 else next(iter(e["qty_dir"]), ""))
        parts = []
        if e["_qty"]:
            ex = "、".join(f"{_short(a)}→{_short(b)}" for a, b in e["_qty"][:3]) + ("…" if len(e["_qty"]) > 3 else "")
            parts.append(f"出貨數量 {len(e['_qty'])} 項{'（' + DIR_LABEL[e['qty_dir']] + '）' if e['qty_dir'] else ''}：{ex}")
        if e["_date"]:
            pairs = {(a, b) for a, b in e["_date"]}
            ex = "、".join(f"{_short(a)}→{_short(b)}" for a, b in sorted(pairs)[:2])
            parts.append(f"交貨日 {ex}（{len(e['_date'])} 項）")
        if e["_gone"]:
            parts.append(f"品項被拿掉 {e['_gone']} 項")
        e["summary"] = "；".join(parts)
        e["kind_labels"] = "、".join(KIND_LABEL[k] for k in ("qty", "date", "gone") if k in e["kinds"])
        for k in ("_qty", "_date", "_gone"):
            e.pop(k)
    return ev


def _event_row(e, p):
    """給 API／Excel 用的一列。"""
    return {"when": e["when"], "po_number": e["po"], "month": p["month"], "lines": "／".join(sorted(p["lines"])),
            "source": e["source"], "source_label": "酷澎" if e["source"] == "import" else "我們",
            "operator": e["operator"], "kinds": sorted(e["kinds"]), "kind_labels": e["kind_labels"],
            "qty_dir": e["qty_dir"], "summary": e["summary"],
            "reason": (e["reason"] or NO_REASON) if e["source"] == "manual" else "", "note": e["note"]}


def _include_test():
    return request.args.get("include_test") in ("1", "true")


def _collect(conn, month_from, month_to, line_group, include_test=False):
    """範圍內的 PO（keep）與它們的改單事件（events）。統計頁和明細 API 共用。
    回傳的第三個值是被排除的「測試」次數（include_test=True 時為 0）。"""
    cfg = _line_groups()
    rows = _rows(conn.execute("SELECT po_number, sku_id, line, delivery_date FROM mst_orders"))
    pos = {}
    for r in rows:
        p = pos.setdefault(r["po_number"], {"dates": set(), "lines": set(), "items": 0, "skus": set()})
        p["dates"].add(r["delivery_date"] or ""); p["lines"].add(_group_of(r["line"], cfg))
        p["items"] += 1; p["skus"].add(r["sku_id"])
    span = set(_month_span(month_from, month_to))
    keep = {}
    for po, p in pos.items():
        m = _po_month(p["dates"])
        if m != UNDATED and m not in span:
            continue
        if line_group and line_group not in p["lines"]:
            continue
        p["month"] = m
        keep[po] = p
    if not keep:
        return keep, {}, 0
    logs = _rows(conn.execute(
        """SELECT po_number, source, field, field_label, old_value, new_value, operator, note, reason, changed_at
           FROM mst_logs
           WHERE po_number != '' AND (
                 (source = 'import' AND field IN ('qty_file_ship', 'delivery_date_file'))
              OR (source = 'import' AND field = 'qty_ship' AND note LIKE ?)
              OR (source = 'manual' AND field IN ('qty_ship', 'delivery_date') AND note NOT LIKE ?))""",
        ("%已沒有這個品項%", f"{_RESTORE_PREFIX}%")))
    events = _events([l for l in logs if l["po_number"] in keep])
    if include_test:
        return keep, events, 0
    tests = [k for k, e in events.items() if e["source"] == "manual" and e["reason"] == TEST_REASON]
    for k in tests:
        events.pop(k)
    return keep, events, len(tests)


def _build_stats(conn, month_from, month_to, line_group, include_test=False):
    keep, events, n_test = _collect(conn, month_from, month_to, line_group, include_test)
    reasons = CHANGE_REASONS + ([TEST_REASON] if include_test else []) + [NO_REASON]
    if not keep:
        return {"months": [], "total": _finish_month(_empty_month("合計"), reasons), "reasons": reasons,
                "top_pos": [], "month_from": month_from, "month_to": month_to, "test_events": 0}

    per_po = {po: {"events": 0, "coupang": 0, "manual": 0, "reasons": collections.Counter(), "kinds": collections.Counter()}
              for po in keep}
    for e in events.values():
        s = per_po[e["po"]]
        s["events"] += 1
        side = "c" if e["source"] == "import" else "m"
        if side == "c":
            s["coupang"] += 1
        else:
            s["manual"] += 1
            s["reasons"][e["reason"] or NO_REASON] += 1
        for k in e["kinds"]:
            s["kinds"][f"{k}_{side}"] += 1
        if "qty" in e["kinds"] and e["qty_dir"]:
            s["kinds"][f"qty_{e['qty_dir']}"] += 1

    months = {}
    order = _month_span(month_from, month_to) + [UNDATED]
    for po, p in keep.items():
        m = months.setdefault(p["month"], _empty_month(p["month"]))
        m["pos"] += 1; m["items"] += p["items"]; m["_skus"].update(p["skus"])
        s = per_po[po]
        if s["events"]:
            m["changed_pos"] += 1
        m["events"] += s["events"]; m["coupang"] += s["coupang"]; m["manual"] += s["manual"]
        for k, v in s["reasons"].items():
            m["reasons"][k] = m["reasons"].get(k, 0) + v
        for k, v in s["kinds"].items():
            m["kinds"][k] = m["kinds"].get(k, 0) + v
    total = _empty_month("合計")
    for m in months.values():
        total["pos"] += m["pos"]; total["items"] += m["items"]; total["_skus"].update(m["_skus"])
        total["changed_pos"] += m["changed_pos"]; total["events"] += m["events"]
        total["coupang"] += m["coupang"]; total["manual"] += m["manual"]
        for k, v in m["reasons"].items():
            total["reasons"][k] = total["reasons"].get(k, 0) + v
        for k, v in m["kinds"].items():
            total["kinds"][k] = total["kinds"].get(k, 0) + v
    out_months = [_finish_month(months[m], reasons) for m in order if m in months]   # 合計要先加完才能收尾（收尾會把 SKU 集合換成數字）
    total = _finish_month(total, reasons)

    top = sorted(((po, keep[po], per_po[po]) for po in keep if per_po[po]["events"]),
                 key=lambda x: (-x[2]["events"], x[0]))[:15]
    top_pos = [{
        "po_number": po, "month": p["month"], "lines": "／".join(sorted(p["lines"])), "items": p["items"],
        "events": s["events"], "coupang": s["coupang"], "manual": s["manual"],
        "reasons": "、".join(f"{k}{v}" if v > 1 else k for k, v in s["reasons"].most_common()),
    } for po, p, s in top]
    return {"months": out_months, "total": total, "reasons": reasons,
            "top_pos": top_pos, "month_from": month_from, "month_to": month_to, "test_events": n_test}


def _event_rows(keep, events):
    rows = [_event_row(e, keep[e["po"]]) for e in events.values()]
    rows.sort(key=lambda r: (r["when"], r["po_number"]), reverse=True)
    return rows


KIND_KEYS = ["qty_c", "qty_m", "qty_down", "qty_up", "qty_mixed", "date_c", "date_m", "gone_c"]


def _empty_month(label):
    return {"month": label, "pos": 0, "items": 0, "_skus": set(), "changed_pos": 0,
            "events": 0, "coupang": 0, "manual": 0, "reasons": {}, "kinds": {}}


def _finish_month(m, reasons=None):
    m["skus"] = len(m.pop("_skus"))
    m["changed_pct"] = round(m["changed_pos"] * 100 / m["pos"]) if m["pos"] else 0
    m["per_changed_po"] = round(m["events"] / m["changed_pos"], 1) if m["changed_pos"] else 0
    m["reasons"] = {k: m["reasons"].get(k, 0) for k in (reasons or CHANGE_REASONS + [NO_REASON])}
    m["kinds"] = {k: m["kinds"].get(k, 0) for k in KIND_KEYS}
    return m


@master_bp.route("/api/master/stats")
def api_stats():
    month_from, month_to, months = _stats_range()
    if not months:
        return jsonify({"error": "月份格式不對，請用 2026-09 這種寫法。"}), 400
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
        return jsonify({"error": "月份格式不對，請用 2026-09 這種寫法。"}), 400
    line = norm_text(request.args.get("line"))
    conn = get_conn()
    try:
        keep, events, n_test = _collect(conn, month_from, month_to, line, _include_test())
    finally:
        conn.close()
    return jsonify({"events": _event_rows(keep, events), "month_from": month_from, "month_to": month_to, "test_events": n_test})


@master_bp.route("/api/master/stats/export")
def api_stats_export():
    month_from, month_to, months = _stats_range()
    if not months:
        return jsonify({"error": "月份格式不對，請用 2026-09 這種寫法。"}), 400
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
    hdr = ["月份", "PO 張數", "品項數", "不同 SKU 數", "被改過的 PO", "被改過比例", "改單次數",
           "其中酷澎改的", "其中我們改的", "被改過的 PO 平均改幾次"]
    if minutes:
        hdr.append(f"估計工時（每次 {minutes} 分鐘）")
    ws.append(hdr)
    for m in d["months"] + [d["total"]]:
        row = [m["month"], m["pos"], m["items"], m["skus"], m["changed_pos"], f"{m['changed_pct']}%",
               m["events"], m["coupang"], m["manual"], m["per_changed_po"]]
        if minutes:
            row.append(round(m["events"] * minutes / 60, 1))
        ws.append(row)
    _style(ws, hdr, bold, head, tot, total_row=True)

    ws2 = wb.create_sheet("改單原因")
    hdr2 = ["月份", "我們改的次數"] + d["reasons"]
    ws2.append(hdr2)
    for m in d["months"] + [d["total"]]:
        ws2.append([m["month"], m["manual"]] + [m["reasons"].get(r, 0) for r in d["reasons"]])
    _style(ws2, hdr2, bold, head, tot, total_row=True)

    wsk = wb.create_sheet("改了什麼")
    hdrk = ["月份", "改數量（酷澎）", "改數量（我們）", "其中下修", "其中上修", "其中有上有下",
            "改交期（酷澎）", "改交期（我們）", "品項被拿掉（酷澎）"]
    wsk.append(hdrk)
    for m in d["months"] + [d["total"]]:
        k = m["kinds"]
        wsk.append([m["month"]] + [k[x] for x in KIND_KEYS])
    _style(wsk, hdrk, bold, head, tot, total_row=True)

    ws3 = wb.create_sheet("改最多次的 PO")
    hdr3 = ["PO 單號", "月份", "線別", "品項數", "改單次數", "酷澎改的", "我們改的", "原因"]
    ws3.append(hdr3)
    for t in d["top_pos"]:
        ws3.append([t["po_number"], t["month"], t["lines"], t["items"], t["events"], t["coupang"], t["manual"], t["reasons"]])
    _style(ws3, hdr3, bold, head, tot, total_row=False)

    wsd = wb.create_sheet("明細")
    hdrd = ["時間", "月份", "PO 單號", "線別", "誰", "來源", "改了什麼", "內容", "原因", "說明"]
    wsd.append(hdrd)
    for r in _event_rows(keep, events):
        wsd.append([r["when"], r["month"], r["po_number"], r["lines"], r["operator"], r["source_label"],
                    r["kind_labels"], r["summary"], r["reason"], r["note"]])
    _style(wsd, hdrd, bold, head, tot, total_row=False)
    wsd.column_dimensions["H"].width = 60; wsd.column_dimensions["A"].width = 20

    ws4 = wb.create_sheet("怎麼算的")
    for line_text in [
        f"範圍：{month_from} ～ {month_to}" + (f"，線別 {line}" if line else "，全部線別")
        + ("，含勾了「這是測試」的改單" if _include_test() else f"，不含勾了「這是測試」的改單（這段期間有 {n_test} 次）"),
        "PO 歸哪個月：看它目前的交貨日（多個日期取最早那天）。沒排日期的另列「未排日期」。",
        "被改過：這張 PO 在變動紀錄裡至少有一次改單事件。",
        "改單事件：同一張 PO、同一來源、同一時間的變動算一次（整張 PO 改期會一個品項記一筆，一次動作不能算很多次）。",
        "酷澎改的：整合表上的出貨數量／交貨日跟上次不一樣，或品項從檔案裡消失。",
        "我們改的：人手改出貨數量／交貨日；恢復成整合表數字／日期不算。",
        "改了什麼：一次改單可能同時改數量和交期，三類各自算，加起來會大於改單次數。改數量再分下修／上修。",
        "原因：只有我們改的才有。缺貨／沒車／酷澎要求／其他；2026-09-22 之前的紀錄沒有原因，列「未填」。",
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


__all__ = ["api_stats", "api_stats_events", "api_stats_export", "_build_stats", "_collect", "_event_rows"]

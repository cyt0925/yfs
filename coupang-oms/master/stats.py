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
- 原因：只有我們改的才有（缺貨／沒車／酷澎要求／其他）；2026-09-22 以前的紀錄沒有原因，列「未填」。
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


def _events(logs):
    """把一筆一筆的紀錄併成「改單事件」：(po, source, 時間) 一個事件。人改的時間看到秒、
    匯入的看到分鐘（見檔頭）。回傳 {key: {"po", "source", "reason", "fields"}}。"""
    ev = {}
    for l in logs:
        stamp = l["changed_at"] or ""
        # 人改的多帶原因進 key：同一秒兩次儲存、原因不同就分成兩次（測試裡會發生，實務上不會少算）
        key = (l["po_number"], l["source"], stamp[:19], l.get("reason") or "") if l["source"] == "manual" \
            else (l["po_number"], l["source"], stamp[:16], "")
        e = ev.setdefault(key, {"po": l["po_number"], "source": l["source"], "reason": "", "fields": set()})
        e["fields"].add(l["field_label"])
        if l["source"] == "manual" and not e["reason"]:
            e["reason"] = l.get("reason") or ""
    return ev


def _build_stats(conn, month_from, month_to, line_group):
    cfg = _line_groups()
    rows = _rows(conn.execute("SELECT po_number, sku_id, line, delivery_date FROM mst_orders"))
    # 每張 PO：歸哪個月、線別、品項、SKU
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
        return {"months": [], "total": _finish_month(_empty_month("合計")), "reasons": CHANGE_REASONS + [NO_REASON],
                "top_pos": [], "month_from": month_from, "month_to": month_to}

    # 變動紀錄：只抓這些 PO、只抓算「改單」的欄位
    logs = _rows(conn.execute(
        """SELECT po_number, source, field, field_label, note, reason, changed_at FROM mst_logs
           WHERE po_number != '' AND (
                 (source = 'import' AND field IN ('qty_file_ship', 'delivery_date_file'))
              OR (source = 'import' AND field = 'qty_ship' AND note LIKE ?)
              OR (source = 'manual' AND field IN ('qty_ship', 'delivery_date') AND note NOT LIKE ?))""",
        ("%已沒有這個品項%", f"{_RESTORE_PREFIX}%")))
    logs = [l for l in logs if l["po_number"] in keep]
    events = _events(logs)

    per_po = {po: {"events": 0, "coupang": 0, "manual": 0, "reasons": collections.Counter()} for po in keep}
    for e in events.values():
        s = per_po[e["po"]]
        s["events"] += 1
        if e["source"] == "import":
            s["coupang"] += 1
        else:
            s["manual"] += 1
            s["reasons"][e["reason"] or NO_REASON] += 1

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
    total = _empty_month("合計")
    for m in months.values():
        total["pos"] += m["pos"]; total["items"] += m["items"]; total["_skus"].update(m["_skus"])
        total["changed_pos"] += m["changed_pos"]; total["events"] += m["events"]
        total["coupang"] += m["coupang"]; total["manual"] += m["manual"]
        for k, v in m["reasons"].items():
            total["reasons"][k] = total["reasons"].get(k, 0) + v
    out_months = [_finish_month(months[m]) for m in order if m in months]   # 合計要先加完才能收尾（收尾會把 SKU 集合換成數字）
    total = _finish_month(total)

    top = sorted(((po, keep[po], per_po[po]) for po in keep if per_po[po]["events"]),
                 key=lambda x: (-x[2]["events"], x[0]))[:15]
    top_pos = [{
        "po_number": po, "month": p["month"], "lines": "／".join(sorted(p["lines"])), "items": p["items"],
        "events": s["events"], "coupang": s["coupang"], "manual": s["manual"],
        "reasons": "、".join(f"{k}{v}" if v > 1 else k for k, v in s["reasons"].most_common()),
    } for po, p, s in top]
    return {"months": out_months, "total": total, "reasons": CHANGE_REASONS + [NO_REASON],
            "top_pos": top_pos, "month_from": month_from, "month_to": month_to}


def _empty_month(label):
    return {"month": label, "pos": 0, "items": 0, "_skus": set(), "changed_pos": 0,
            "events": 0, "coupang": 0, "manual": 0, "reasons": {}}


def _finish_month(m):
    m["skus"] = len(m.pop("_skus"))
    m["changed_pct"] = round(m["changed_pos"] * 100 / m["pos"]) if m["pos"] else 0
    m["per_changed_po"] = round(m["events"] / m["changed_pos"], 1) if m["changed_pos"] else 0
    m["reasons"] = {k: m["reasons"].get(k, 0) for k in CHANGE_REASONS + [NO_REASON]}
    return m


@master_bp.route("/api/master/stats")
def api_stats():
    month_from, month_to, months = _stats_range()
    if not months:
        return jsonify({"error": "月份格式不對，請用 2026-09 這種寫法。"}), 400
    line = norm_text(request.args.get("line"))
    conn = get_conn()
    try:
        data = _build_stats(conn, month_from, month_to, line)
    finally:
        conn.close()
    data["line"] = line
    return jsonify(data)


@master_bp.route("/api/master/stats/export")
def api_stats_export():
    month_from, month_to, months = _stats_range()
    if not months:
        return jsonify({"error": "月份格式不對，請用 2026-09 這種寫法。"}), 400
    line = norm_text(request.args.get("line"))
    minutes = norm_int(request.args.get("minutes")) or 0
    conn = get_conn()
    try:
        d = _build_stats(conn, month_from, month_to, line)
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

    ws3 = wb.create_sheet("改最多次的 PO")
    hdr3 = ["PO 單號", "月份", "線別", "品項數", "改單次數", "酷澎改的", "我們改的", "原因"]
    ws3.append(hdr3)
    for t in d["top_pos"]:
        ws3.append([t["po_number"], t["month"], t["lines"], t["items"], t["events"], t["coupang"], t["manual"], t["reasons"]])
    _style(ws3, hdr3, bold, head, tot, total_row=False)

    ws4 = wb.create_sheet("怎麼算的")
    for line_text in [
        f"範圍：{month_from} ～ {month_to}" + (f"，線別 {line}" if line else "，全部線別"),
        "PO 歸哪個月：看它目前的交貨日（多個日期取最早那天）。沒排日期的另列「未排日期」。",
        "被改過：這張 PO 在變動紀錄裡至少有一次改單事件。",
        "改單事件：同一張 PO、同一來源、同一時間的變動算一次（整張 PO 改期會一個品項記一筆，一次動作不能算很多次）。",
        "酷澎改的：整合表上的出貨數量／交貨日跟上次不一樣，或品項從檔案裡消失。",
        "我們改的：人手改出貨數量／交貨日；恢復成整合表數字／日期不算。",
        "原因：只有我們改的才有。缺貨／沒車／酷澎要求／其他；2026-09-22 之前的紀錄沒有原因，列「未填」。",
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


__all__ = ["api_stats", "api_stats_export", "_build_stats"]

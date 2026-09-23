"""③ 總表看板：把寶僑總表那一百欄，用「一個月一頁、商品層＋品牌層」的方式呈現在介面上。

每一格的來源與公式（對照 docs/寶僑總表_欄位與公式對照.md，欄位字母是總表的）：
  商品層（一列一個商品）
    COGS／GIV／NIV               主檔（總表 H／K／L；GIV、NIV 由 Coupang Master 更新）
    Supply／demand               mst_month_stats（總表 M／P，由 supply 表更新）
    本月下單 TTL                 訂單明細現算：Σ 出貨數量 ÷ 箱入數（總表 AF）
    剩餘可供貨                   Supply − TTL（總表 BE；BF 目前等於 BE）
    剩餘庫存／庫存天數／到月底   sources.stock_by_barcode（庫存銷售表 AK／AM／AO／AP 公式）
    Supply GIV                   Supply × GIV（總表 BK）
    下單 GIV                     TTL × GIV（總表 BN）
    下單 COGS 含稅               TTL × COGS × 箱入數（總表 BO）
    下單永豐成本未稅             NIV × 1.02 × TTL（總表 BP）
  品牌層（一列一個品牌，總表 BY～CU）
    各金額 = 該品牌商品加總（SUMIF）；目標／REBATE目標 人填（mst_brand_targets）；
    達成 = 下單 GIV ÷ 目標；diff = 目標 − 下單 GIV；REBATE DIFF = REBATE目標 − COGS 含稅

列哪些商品：這個線別、這個月「有下單」或「supply 表有給 Supply／demand」或「有庫存資料」的商品。
主檔裡有但三者皆無的不列（跟 Excel 一樣，沒事的商品不佔位）。
警示（flags）：over（下單 > Supply）、no_over（Note 是「不可超打」且下單 > Supply）、
low_stock（庫存天數 < LOW_STOCK_DAYS）、no_box（算不出箱數）、no_price（缺 GIV）。
"""
from .common import *  # noqa: F401,F403 — 共用工具、Flask、db、openpyxl 都從這裡來
from .summary import _build_summary
from .sources import stock_by_barcode

LOW_STOCK_DAYS = 14
YF_COST_FACTOR = 1.02        # 總表 BP：永豐成本未稅 = NIV × 1.02 × 下單箱數


def _mul(a, b, c=1.0):
    if a is None or b is None or c is None:
        return None
    return round(float(a) * float(b) * float(c), 2)


def _line_products(conn, group, cfg):
    """這個線別的主檔商品（master_line 或 lines_seen 歸到這個顯示名）。"""
    out = {}
    for p in _rows(conn.execute("SELECT * FROM mst_products")):
        groups = {_group_of(l, cfg) for l in _split_lines(p["lines_seen"])}
        if p["master_line"]:
            groups.add(_group_of(p["master_line"], cfg))
        if group in groups:
            out[p["barcode"]] = p
    return out


def build_board(conn, group, month, cfg):
    summary = _build_summary(conn, group, month, cfg)          # 訂單那一段：每日箱數、TTL
    sum_by_bc = {r["barcode"]: r for r in summary["rows"]}
    products = _line_products(conn, group, cfg)
    stats = {r["barcode"]: r for r in _rows(conn.execute(
        "SELECT * FROM mst_month_stats WHERE month = ?", (month,)))}
    targets = {t["brand"]: t for t in _rows(conn.execute(
        "SELECT * FROM mst_brand_targets WHERE line = ? AND month = ?", (group, month)))}

    barcodes = set(sum_by_bc) | {bc for bc in stats if bc in products}
    stock = stock_by_barcode(conn, month, sorted(barcodes)) if barcodes else {}
    barcodes |= {bc for bc, st in stock.items() if st["has_stock_data"] and bc in products}

    rows = []
    for bc in barcodes:
        p = products.get(bc) or {}
        s = sum_by_bc.get(bc) or {}
        m = stats.get(bc) or {}
        st = stock.get(bc) or {}
        ttl = s.get("month_total") or 0.0
        supply, demand = m.get("supply_cs"), m.get("demand_cs")
        giv, niv, cost, box = p.get("giv"), p.get("niv"), p.get("cost_price"), p.get("box_size") or s.get("box_size")
        note = p.get("note") or s.get("note") or ""
        remaining_supply = round(supply - ttl, 2) if supply is not None else None
        flags = []
        if supply is not None and ttl > supply:
            flags.append("no_over" if "不可超打" in note else "over")
        if st.get("stock_days") is not None and st["stock_days"] < LOW_STOCK_DAYS:
            flags.append("low_stock")
        if s.get("missing_box"):
            flags.append("no_box")
        if ttl and giv is None:
            flags.append("no_price")
        rows.append({
            "barcode": bc, "sku_id": p.get("sku_id") or s.get("sku_id", ""), "yf_sku": p.get("yf_sku") or s.get("yf_sku", ""),
            "brand": p.get("brand") or s.get("brand", ""), "category": p.get("category") or s.get("category", ""),
            "product_name": p.get("product_name") or s.get("product_name", ""), "note": note,
            "in_master": bool(p), "box_size": box, "cost": cost, "giv": giv, "niv": niv,
            "supply": supply, "demand": demand, "ttl": round(ttl, 2), "remaining_supply": remaining_supply,
            "by_date": s.get("by_date") or {}, "missing_box": s.get("missing_box", 0), "order_rows": s.get("order_rows", 0),
            "stock_remaining": st.get("remaining_cs") if st.get("has_stock_data") else None,
            "stock_days": st.get("stock_days"), "stock_remaining_eom": st.get("remaining_eom"),
            "stock_days_eom": st.get("stock_days_eom"), "sold": st.get("sold_cs"), "sold_days": st.get("sold_days"),
            "stock_giv": _mul(st.get("remaining_cs") if st.get("has_stock_data") else None, giv),
            "supply_giv": _mul(supply, giv), "ttl_giv": _mul(ttl, giv),
            "cogs_tax": _mul(ttl, cost, box), "yf_cost": _mul(ttl, niv, YF_COST_FACTOR),
            "flags": flags,
        })
    rows.sort(key=lambda r: (r["brand"] or "", r["product_name"] or "", r["barcode"]))

    # 品牌層
    brands = collections.OrderedDict()
    for r in rows:
        b = brands.setdefault(r["brand"] or "（無品牌）", {"brand": r["brand"] or "（無品牌）", "category": r["category"],
                                                          "products": 0, "supply": 0.0, "ttl": 0.0, "supply_giv": 0.0,
                                                          "ttl_giv": 0.0, "cogs_tax": 0.0, "yf_cost": 0.0,
                                                          "stock_remaining": 0.0, "stock_giv": 0.0, "flags": 0})
        b["products"] += 1; b["supply"] += r["supply"] or 0; b["ttl"] += r["ttl"]
        for k in ("supply_giv", "ttl_giv", "cogs_tax", "yf_cost", "stock_remaining", "stock_giv"):
            b[k] += r[k] or 0
        b["flags"] += len(r["flags"])
    brand_rows = []
    for b in brands.values():
        t = targets.get(b["brand"]) or {}
        tg, rb = t.get("target_giv"), t.get("rebate_target")
        for k in ("supply", "ttl", "supply_giv", "ttl_giv", "cogs_tax", "yf_cost", "stock_remaining", "stock_giv"):
            b[k] = round(b[k], 2)
        b.update({"target_giv": tg, "rebate_target": rb,
                  "target_pct": round(b["ttl_giv"] / tg * 100) if tg else None,
                  "target_diff": round(tg - b["ttl_giv"], 2) if tg is not None else None,
                  "rebate_diff": round(rb - b["cogs_tax"], 2) if rb is not None else None})
        brand_rows.append(b)
    brand_rows.sort(key=lambda b: -b["ttl_giv"])

    totals = {k: round(sum(r[k] or 0 for r in rows), 2) for k in
              ("supply", "ttl", "supply_giv", "ttl_giv", "cogs_tax", "yf_cost", "stock_giv")}
    totals["stock_remaining"] = round(sum(r["stock_remaining"] or 0 for r in rows), 2)
    totals["products"] = len(rows)
    totals["with_orders"] = sum(1 for r in rows if r["ttl"])
    totals["alerts"] = sum(1 for r in rows if r["flags"])
    totals["over"] = sum(1 for r in rows if "over" in r["flags"] or "no_over" in r["flags"])
    totals["low_stock"] = sum(1 for r in rows if "low_stock" in r["flags"])
    totals["no_box"] = sum(1 for r in rows if "no_box" in r["flags"])
    totals["has_supply"] = any(r["supply"] is not None for r in rows)
    totals["has_stock"] = any(r["stock_remaining"] is not None for r in rows)
    totals["target_giv"] = round(sum(b["target_giv"] or 0 for b in brand_rows), 2) or None
    return {"line": group, "month": month, "dates": summary["dates"], "totals_by_date": summary["totals_by_date"],
            "rows": rows, "brands": brand_rows, "totals": totals, "low_stock_days": LOW_STOCK_DAYS}


@master_bp.route("/api/master/board")
def api_board():
    group = norm_text(request.args.get("line")); month = norm_text(request.args.get("month")) or _this_month()
    if not group:
        return jsonify({"error": "請選線別。"}), 400
    if not _valid_month(month):
        return jsonify({"error": "月份格式要像 2026-09。"}), 400
    cfg = _line_groups()
    conn = get_conn()
    try:
        return jsonify(build_board(conn, group, month, cfg))
    finally:
        conn.close()


@master_bp.route("/api/master/brand_targets", methods=["PUT"])
def api_set_brand_target():
    """品牌目標（下單 GIV 目標、REBATE 目標）：人填。空字串 = 清掉。"""
    payload = request.get_json(silent=True) or {}
    group, month, brand = norm_text(payload.get("line")), norm_text(payload.get("month")), norm_text(payload.get("brand"))
    if not group or not _valid_month(month) or not brand:
        return jsonify({"error": "缺線別、月份或品牌。"}), 400
    vals = {}
    for k in ("target_giv", "rebate_target"):
        if k in payload:
            v = payload.get(k)
            if v in (None, ""):
                vals[k] = None
            else:
                n = norm_decimal(v)
                if n is None:
                    return jsonify({"error": "目標要是數字。"}), 400
                vals[k] = float(n)
    if not vals:
        return jsonify({"error": "沒有要改的欄位。"}), 400
    operator = _operator()
    conn = get_conn()
    try:
        cur = _row(conn.execute("SELECT * FROM mst_brand_targets WHERE line = ? AND month = ? AND brand = ?", (group, month, brand)))
        if cur is None:
            conn.execute("INSERT INTO mst_brand_targets (line, month, brand, target_giv, rebate_target, updated_by, updated_at) VALUES (?,?,?,?,?,?,?)",
                         (group, month, brand, vals.get("target_giv"), vals.get("rebate_target"), operator, now()))
        else:
            sets = ", ".join(f"{k} = ?" for k in vals)
            conn.execute(f"UPDATE mst_brand_targets SET {sets}, updated_by = ?, updated_at = ? WHERE id = ?",
                         list(vals.values()) + [operator, now(), cur["id"]])
        for k, v in vals.items():
            _log(conn, group, "", "", "", f"brand_{k}", "品牌目標" if k == "target_giv" else "REBATE目標",
                 (cur or {}).get(k), v, operator, "manual", f"{brand} {month}")
        conn.commit()
        fresh = _row(conn.execute("SELECT * FROM mst_brand_targets WHERE line = ? AND month = ? AND brand = ?", (group, month, brand)))
        return jsonify({"ok": True, "target": fresh})
    finally:
        conn.close()


__all__ = ["build_board", "api_board", "api_set_brand_target", "LOW_STOCK_DAYS"]

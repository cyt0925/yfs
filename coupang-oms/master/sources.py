"""① 外部來源檔匯入：寶僑 supply 表、Coupang Master、庫存銷售表，加上總表 Sheet1 裡的 GIV／NIV／供需。

這些檔以前是業務在總表裡用 VLOOKUP 撈的（總表 M／P 欄指 supply 表，K／L 欄指 Master，
BH～BJ 欄指庫存銷售表）。系統接手後改成「檔案拖進同一個匯入口，看表頭認出是哪一種」，
把值寫進主檔（GIV／NIV）與 mst_month_stats（每商品每月的 supply／demand／下單／實銷）。
剩餘庫存、庫存天數不存，讀的時候現算：累計下單 − 累計實銷、剩餘 ÷ 日均實銷。

認檔規則（都看表頭文字，不看欄位位置、不看分頁名稱——分頁名帶月份每季會換）：
  supply       表頭有 SKU ID，且有「N月supply cs」／「N月demand cs」
  stock        表頭有「Y26 N月…下單總箱數」或「…實銷總箱數」（Y26 = 2026）
  master_price 表頭同時有 Barcode、GIV、NIV，而且沒有總表才有的 Supply_CS／M/D交貨 欄
  sheet        寶僑總表 Sheet1（走原本的主檔匯入），這裡只多吸 K／L／M～R 幾欄

年份：supply 表與總表的月份欄只寫「10月」「(Oct)」沒有年，取離今天最近的那一年（±6 個月內）。
來源格子是空白、#N/A、#REF 一律不覆蓋，只算進 skipped；來源沒有的商品不新建主檔，只回報。
"""
import datetime as _dt
from .common import *  # noqa: F401,F403 — 共用工具、Flask、db、openpyxl 都從這裡來

KIND_LABEL = {"supply": "寶僑 supply 表", "master_price": "Coupang Master", "stock": "庫存銷售表", "sheet": "寶僑總表"}
_MON_EN = {m: i for i, m in enumerate(["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"], start=1)}


def _h(v):
    """表頭正規化：小寫、去空白與換行、全形括號轉半形。"""
    t = norm_text(v).lower().replace(" ", "").replace("\n", "").replace("（", "(").replace("）", ")")
    return t


def _headers(ws, max_scan=15):
    """找表頭列：前 15 列裡「非空格最多」的那一列。回傳 (列號, [正規化表頭], [原表頭])。"""
    best = (None, [], [])
    for idx, row in enumerate(ws.iter_rows(min_row=1, max_row=max_scan, values_only=True), start=1):
        raw = [norm_text(c) for c in row]
        n = sum(1 for c in raw if c)
        if n > len([c for c in best[1] if c]):
            best = (idx, [_h(c) for c in raw], raw)
    return best


def _year_for(month, today=None):
    """沒寫年的月份 → 離今天最近的那一年（例如今天 2026-09，「12月」→ 2026，「2月」→ 2027，「7月」→ 2026）。"""
    today = today or _dt.date.today()
    best = None
    for y in (today.year - 1, today.year, today.year + 1):
        diff = abs((y - today.year) * 12 + (month - today.month))
        if best is None or diff < best[0]:
            best = (diff, y)
    return best[1]


def _num(v):
    """數字格：#N/A、#REF!、空白、文字都當「沒有值」→ None。"""
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    t = str(v).strip()
    if not t or t.startswith("#"):
        return None
    try:
        return float(t.replace(",", ""))
    except ValueError:
        return None


# ───────── 認檔 ─────────
_RE_SUPPLY = re.compile(r"^(\d{1,2})月(supply|demand)cs$")
_RE_STOCK = re.compile(r"^y(\d{2})(\d{1,2})月.*?(下單|實銷)總箱數")
_RE_STOCK_PART = re.compile(r"^y(\d{2})(\d{1,2})/(\d{1,2})-(\d{1,2})/(\d{1,2})實銷總箱數")
_RE_SHEET_MONTH = re.compile(r"^(supply_cs|demand_cs)\((\w{3})\)$")


def detect_kind(headers):
    """headers 已正規化。回傳 'supply'／'stock'／'master_price'／''。總表由主檔匯入自己認（有 M/D交貨），這裡不回 sheet。"""
    hs = set(headers)
    if any(_RE_SUPPLY.match(h) for h in headers) and ({"skuid", "sku id".replace(" ", "")} & hs):
        return "supply"
    if any(_RE_STOCK.match(h) or _RE_STOCK_PART.match(h) for h in headers):
        return "stock"
    if {"barcode", "giv", "niv"} <= hs and not any(_RE_SHEET_MONTH.match(h) for h in headers) \
            and not any(re.match(r"^\d{1,2}/\d{0,2}交貨", h) for h in headers):
        return "master_price"
    return ""


def find_source_sheet(wb):
    """整本翻一遍，回傳第一個認得出來的 (kind, ws, 表頭列號, headers, raw_headers)。"""
    for ws in wb.worksheets:
        idx, hs, raw = _headers(ws)
        if not idx:
            continue
        kind = detect_kind(hs)
        if kind:
            return kind, ws, idx, hs, raw
    return "", None, None, [], []


# ───────── 寫入 ─────────
def _barcode_map(conn):
    """主檔：barcode 集合、skuid → barcode。"""
    rows = _rows(conn.execute("SELECT barcode, sku_id, giv, niv FROM mst_products"))
    by_bc = {r["barcode"]: r for r in rows}
    by_sku = {r["sku_id"]: r["barcode"] for r in rows if r["sku_id"]}
    return by_bc, by_sku


def _upsert_month(conn, barcode, month, operator, **vals):
    """只寫有值的欄位；回傳是否有改到。"""
    vals = {k: v for k, v in vals.items() if v is not None}
    if not vals:
        return False
    cur = _row(conn.execute("SELECT * FROM mst_month_stats WHERE barcode = ? AND month = ?", (barcode, month)))
    stamp = now()
    if cur is None:
        cols = ["barcode", "month", "updated_by", "updated_at"] + list(vals)
        conn.execute(f"INSERT INTO mst_month_stats ({', '.join(cols)}) VALUES ({', '.join('?' * len(cols))})",
                     [barcode, month, operator, stamp] + list(vals.values()))
        return True
    changed = {k: v for k, v in vals.items() if not _same(cur.get(k), v)}
    if not changed:
        return False
    sets = ", ".join(f"{k} = ?" for k in changed) + ", updated_by = ?, updated_at = ?"
    conn.execute(f"UPDATE mst_month_stats SET {sets} WHERE id = ?", list(changed.values()) + [operator, stamp, cur["id"]])
    return True


def _set_price(conn, barcode, cur, giv, niv, operator, fname):
    """GIV／NIV：來源有數字才更新，有變才記歷程。回傳是否有改到。"""
    sets, vals, changed = [], [], False
    for field, label, new in (("giv", "GIV", giv), ("niv", "NIV", niv)):
        if new is None:
            continue
        new = round(new, 4)
        if not _same(cur.get(field), new):
            sets.append(f"{field} = ?"); vals.append(new); changed = True
            _log(conn, "", "", cur.get("sku_id") or "", barcode, field, label, cur.get(field), new, operator, "import", fname)
    if changed:
        conn.execute(f"UPDATE mst_products SET {', '.join(sets)}, updated_by = ?, updated_at = ? WHERE barcode = ?",
                     vals + [operator, now(), barcode])
    return changed


def _record_upload(conn, kind, fname, sheet, operator, total, matched, updated, skipped, months, note=""):
    conn.execute(
        """INSERT INTO mst_source_uploads (kind, filename, sheet, operator, uploaded_at, rows_total, rows_matched,
           rows_updated, rows_skipped, months, note) VALUES (?,?,?,?,?,?,?,?,?,?,?)""",
        (kind, fname, sheet, operator, now(), total, matched, updated, skipped, ",".join(sorted(months)), note))
    _log(conn, "", "", "", "", "source_upload", KIND_LABEL.get(kind, kind), "", f"{matched}/{total}", operator, "import",
         f"{fname}：對到 {matched} 個商品、更新 {updated}、來源沒值 {skipped}" + (f"；月份 {'、'.join(sorted(months))}" if months else ""))


def import_supply(conn, ws, hdr_idx, headers, operator, fname):
    """寶僑 supply 表：SKU ID（或 Barcode）→ 每月 supply_cs／demand_cs。"""
    by_bc, by_sku = _barcode_map(conn)
    i_sku = headers.index("skuid") if "skuid" in headers else None
    i_bc = headers.index("barcode") if "barcode" in headers else None
    month_cols = {}   # (月份字串, 'supply_cs'|'demand_cs') → 欄位 index
    for i, h in enumerate(headers):
        m = _RE_SUPPLY.match(h)
        if m:
            mon = int(m.group(1)); month_cols[(f"{_year_for(mon):04d}-{mon:02d}", m.group(2) + "_cs")] = i
    total = matched = updated = skipped = 0; missing = []
    for row in ws.iter_rows(min_row=hdr_idx + 1, values_only=True):
        sku = norm_key(row[i_sku]) if i_sku is not None and i_sku < len(row) else ""
        bc = norm_key(row[i_bc]) if i_bc is not None and i_bc < len(row) else ""
        if not sku and not bc:
            continue
        total += 1
        barcode = bc if bc in by_bc else by_sku.get(sku, "")
        if not barcode:
            missing.append(sku or bc); continue
        matched += 1
        per_month = collections.defaultdict(dict)
        for (month, field), i in month_cols.items():
            v = _num(row[i]) if i < len(row) else None
            if v is None:
                skipped += 1
            else:
                per_month[month][field] = v
        for month, vals in per_month.items():
            if _upsert_month(conn, barcode, month, operator, **vals):
                updated += 1
    months = {m for m, _ in month_cols}
    _record_upload(conn, "supply", fname, ws.title, operator, total, matched, updated, skipped, months)
    return {"kind": "supply", "label": KIND_LABEL["supply"], "sheet": ws.title, "rows": total, "matched": matched,
            "updated": updated, "skipped": skipped, "months": sorted(months), "not_in_master": len(missing),
            "not_in_master_examples": missing[:5]}


def import_master_price(conn, ws, hdr_idx, headers, operator, fname):
    """Coupang Master：Barcode → GIV／NIV。"""
    by_bc, _ = _barcode_map(conn)
    i_bc, i_giv, i_niv = headers.index("barcode"), headers.index("giv"), headers.index("niv")
    total = matched = updated = skipped = 0; missing = []
    for row in ws.iter_rows(min_row=hdr_idx + 1, values_only=True):
        bc = norm_key(row[i_bc]) if i_bc < len(row) else ""
        if not bc:
            continue
        total += 1
        cur = by_bc.get(bc)
        if cur is None:
            missing.append(bc); continue
        matched += 1
        giv = _num(row[i_giv]) if i_giv < len(row) else None
        niv = _num(row[i_niv]) if i_niv < len(row) else None
        skipped += (giv is None) + (niv is None)
        if _set_price(conn, bc, cur, giv, niv, operator, fname):
            updated += 1
    _record_upload(conn, "master_price", fname, ws.title, operator, total, matched, updated, skipped, set())
    return {"kind": "master_price", "label": KIND_LABEL["master_price"], "sheet": ws.title, "rows": total, "matched": matched,
            "updated": updated, "skipped": skipped, "months": [], "not_in_master": len(missing),
            "not_in_master_examples": missing[:5]}


def import_stock(conn, ws, hdr_idx, headers, operator, fname):
    """庫存銷售表：Bardcode／skuid → 每月 ordered_cs（下單）、sold_cs（實銷，含「9/1-9/20」這種部分月）。
    只吃這兩種原料，剩餘庫存、庫存天數系統自己算。"""
    by_bc, by_sku = _barcode_map(conn)
    i_bc = next((i for i, h in enumerate(headers) if h in ("bardcode", "barcode", "國條", "條碼")), None)
    i_sku = headers.index("skuid") if "skuid" in headers else None
    cols = {}   # index → (month, field, sold_days)
    for i, h in enumerate(headers):
        m = _RE_STOCK_PART.match(h)
        if m:
            # 「Y26 9/1-9/20實銷總箱數」：群組 2=月、3=起日、5=迄日 → 涵蓋 20 天
            yy, mon, d1, d2 = int(m.group(1)), int(m.group(2)), int(m.group(3)), int(m.group(5))
            cols[i] = (f"{2000 + yy:04d}-{mon:02d}", "sold_cs", d2 - d1 + 1); continue
        m = _RE_STOCK.match(h)
        if m:
            yy, mon = int(m.group(1)), int(m.group(2))
            cols[i] = (f"{2000 + yy:04d}-{mon:02d}", "ordered_cs" if m.group(3) == "下單" else "sold_cs", None)
    total = matched = updated = skipped = 0; missing = []
    for row in ws.iter_rows(min_row=hdr_idx + 1, values_only=True):
        bc = norm_key(row[i_bc]) if i_bc is not None and i_bc < len(row) else ""
        sku = norm_key(row[i_sku]) if i_sku is not None and i_sku < len(row) else ""
        if not bc and not sku:
            continue
        total += 1
        barcode = bc if bc in by_bc else by_sku.get(sku, "")
        if not barcode:
            missing.append(bc or sku); continue
        matched += 1
        per_month = collections.defaultdict(dict)
        for i, (month, field, days) in cols.items():
            v = _num(row[i]) if i < len(row) else None
            if v is None:
                skipped += 1; continue
            per_month[month][field] = v
            if field == "sold_cs":
                per_month[month]["sold_days"] = days or _days_in_month(month)
        for month, vals in per_month.items():
            if _upsert_month(conn, barcode, month, operator, **vals):
                updated += 1
    months = {m for m, _, _ in cols.values()}
    _record_upload(conn, "stock", fname, ws.title, operator, total, matched, updated, skipped, months)
    return {"kind": "stock", "label": KIND_LABEL["stock"], "sheet": ws.title, "rows": total, "matched": matched,
            "updated": updated, "skipped": skipped, "months": sorted(months), "not_in_master": len(missing),
            "not_in_master_examples": missing[:5]}


def absorb_sheet_extras(conn, ws, hdr_idx, raw_headers, operator, fname):
    """寶僑總表 Sheet1 多吸幾欄：GIV／NIV → 主檔；Supply_CS (Oct)／demand_CS (Oct) → 月統計。
    主檔匯入（products.py）認出是總表時呼叫。回傳摘要 dict。"""
    headers = [_h(c) for c in raw_headers]
    by_bc, _ = _barcode_map(conn)
    i_bc = headers.index("barcode") if "barcode" in headers else None
    if i_bc is None:
        return {}
    i_giv = headers.index("giv") if "giv" in headers else None
    i_niv = headers.index("niv") if "niv" in headers else None
    month_cols = {}
    for i, h in enumerate(headers):
        m = _RE_SHEET_MONTH.match(h)
        if m and m.group(2) in _MON_EN:
            mon = _MON_EN[m.group(2)]
            month_cols[(f"{_year_for(mon):04d}-{mon:02d}", m.group(1))] = i
    price_updated = month_updated = 0
    for row in ws.iter_rows(min_row=hdr_idx + 1, values_only=True):
        bc = norm_key(row[i_bc]) if i_bc < len(row) else ""
        if not bc:
            continue
        cur = by_bc.get(bc)
        if cur is None:   # 主檔匯入剛建的新商品，重撈一次
            cur = _row(conn.execute("SELECT barcode, sku_id, giv, niv FROM mst_products WHERE barcode = ?", (bc,)))
            if cur is None:
                continue
        giv = _num(row[i_giv]) if i_giv is not None and i_giv < len(row) else None
        niv = _num(row[i_niv]) if i_niv is not None and i_niv < len(row) else None
        if _set_price(conn, bc, cur, giv, niv, operator, fname):
            price_updated += 1
        per_month = collections.defaultdict(dict)
        for (month, field), i in month_cols.items():
            v = _num(row[i]) if i < len(row) else None
            if v is not None:
                per_month[month][field] = v
        for month, vals in per_month.items():
            if _upsert_month(conn, bc, month, operator, **vals):
                month_updated += 1
    months = {m for m, _ in month_cols}
    _record_upload(conn, "sheet", fname, ws.title, operator, 0, 0, price_updated + month_updated, 0, months,
                   note="總表順便帶進 GIV／NIV／供需")
    return {"price_updated": price_updated, "month_updated": month_updated, "months": sorted(months)}


def _days_in_month(month):
    y, m = int(month[:4]), int(month[5:7])
    nxt = _dt.date(y + (m == 12), 1 if m == 12 else m + 1, 1)
    return (nxt - _dt.date(y, m, 1)).days


# ───────── 讀 ─────────
@master_bp.route("/api/master/sources")
def api_sources():
    """每種來源檔最後一次上傳（畫面上「上次上傳」那一行）。"""
    conn = get_conn()
    try:
        rows = _rows(conn.execute("SELECT * FROM mst_source_uploads ORDER BY uploaded_at DESC, id DESC LIMIT 200"))
    finally:
        conn.close()
    latest = {}
    for r in rows:
        latest.setdefault(r["kind"], r)
    out = []
    for kind in ("supply", "master_price", "stock", "sheet"):
        r = latest.get(kind)
        out.append({"kind": kind, "label": KIND_LABEL[kind], "last": r})
    return jsonify({"sources": out, "recent": rows[:30]})


def stock_by_barcode(conn, month, barcodes=None):
    """到 month 為止每個商品的庫存現算（公式見 api_month_stats）。回傳 {barcode: {...}}。
    barcodes 給了就只算那些；None 算全部有月統計的商品。"""
    if barcodes is None:
        return _stock_chunk(conn, month, None)
    out_all = {}
    for i in range(0, len(barcodes), 400):          # IN (...) 一次最多 400 個，PostgreSQL／SQLite 都安全
        out_all.update(_stock_chunk(conn, month, barcodes[i:i + 400]))
    return out_all


def _stock_chunk(conn, month, barcodes):
    where, params = ["s.month <= ?"], [month]
    if barcodes:
        where.append(f"s.barcode IN ({','.join('?' * len(barcodes))})"); params += barcodes
    rows = _rows(conn.execute(
        f"SELECT s.* FROM mst_month_stats s WHERE {' AND '.join(where)} ORDER BY s.barcode, s.month", params))
    dim = _days_in_month(month)
    acc = {}
    for r in rows:
        a = acc.setdefault(r["barcode"], {"ordered": 0.0, "sold": 0.0, "cur": None})
        a["ordered"] += r["ordered_cs"] or 0; a["sold"] += r["sold_cs"] or 0
        if r["month"] == month:
            a["cur"] = r
    out = {}
    for bc, a in acc.items():
        cur = a["cur"] or {}
        remaining = round(a["ordered"] - a["sold"], 2)
        sold, days = cur.get("sold_cs"), cur.get("sold_days")
        daily = (sold / days) if sold and days else None
        month_sales = round(daily * dim, 2) if daily else None
        remaining_eom = round(remaining - daily * (dim - days), 2) if daily and days else None
        eom_days = None
        if month_sales:
            eom_days = round((a["ordered"] - ((a["sold"] - (sold or 0)) + month_sales)) / month_sales * dim, 1)
        out[bc] = {"supply_cs": cur.get("supply_cs"), "demand_cs": cur.get("demand_cs"),
                   "ordered_cs": cur.get("ordered_cs"), "sold_cs": sold, "sold_days": days,
                   "has_stock_data": bool(a["ordered"] or a["sold"]),
                   "remaining_cs": remaining, "stock_days": round(remaining / daily, 1) if daily else None,
                   "month_sales_proj": month_sales, "remaining_eom": remaining_eom, "stock_days_eom": eom_days}
    return out


@master_bp.route("/api/master/month_stats")
def api_month_stats():
    """某個月每個商品的供需與進銷，加上現算的庫存數字。公式照庫存銷售表原樣（欄位代號是那份表的）：
      剩餘庫存(箱)        AK = Σ下單(O~T) − Σ實銷(X~AC)            → remaining_cs
      日均實銷            AC ÷ 20（20 = 實銷欄涵蓋的天數）            → daily
      目前庫存天數        AM = AK ÷ 日均實銷                        → stock_days
      預計到月底銷售(箱)  AN = 日均實銷 × 30（30 = 該月天數）          → month_sales_proj
      預計到月底剩餘(箱)  AO = AK − 日均實銷 × 剩下的天數             → remaining_eom
      預計到月底庫存天數  AP = (Σ下單 − (Σ過去月實銷 + AN)) ÷ AN × 30 → stock_days_eom
      剩餘庫金            AL = AK × GIV                             → remaining_giv
    另外總表的：PG 最大剩餘可供貨量 BE = Supply − 該月下單；最低應打完 BF = 綜合CS − 該月下單（綜合CS 目前就等於 Supply）。"""
    month = request.args.get("month") or _this_month()
    if not _valid_month(month):
        return jsonify({"error": "月份格式不對，請用 2026-09 這種寫法。"}), 400
    barcode = norm_key(request.args.get("barcode"))
    conn = get_conn()
    try:
        where, params = ["s.month <= ?"], [month]
        if barcode:
            where.append("s.barcode = ?"); params.append(barcode)
        rows = _rows(conn.execute(
            f"""SELECT s.*, p.product_name, p.brand, p.category, p.giv, p.niv, p.box_size
                FROM mst_month_stats s LEFT JOIN mst_products p ON p.barcode = s.barcode
                WHERE {' AND '.join(where)} ORDER BY s.barcode, s.month""", params))
    finally:
        conn.close()
    acc = {}
    for r in rows:
        a = acc.setdefault(r["barcode"], {"ordered": 0.0, "sold": 0.0, "cur": None, "product_name": r["product_name"],
                                          "brand": r["brand"], "category": r["category"], "giv": r["giv"], "niv": r["niv"]})
        a["ordered"] += r["ordered_cs"] or 0; a["sold"] += r["sold_cs"] or 0
        if r["month"] == month:
            a["cur"] = r
    dim = _days_in_month(month)
    out = []
    for bc, a in acc.items():
        cur = a["cur"] or {}
        ordered_all, sold_all = a["ordered"], a["sold"]
        remaining = round(ordered_all - sold_all, 2)                      # AK
        sold, days = cur.get("sold_cs"), cur.get("sold_days")
        daily = (sold / days) if sold and days else None                   # AC ÷ 20
        month_sales = round(daily * dim, 2) if daily else None             # AN
        remaining_eom = round(remaining - daily * (dim - days), 2) if daily and days else None   # AO
        eom_days = None                                                    # AP
        if month_sales:
            eom_days = round((ordered_all - ((sold_all - (sold or 0)) + month_sales)) / month_sales * dim, 1)
        supply, ordered = cur.get("supply_cs"), cur.get("ordered_cs")
        out.append({"barcode": bc, "product_name": a["product_name"], "brand": a["brand"], "category": a["category"],
                    "giv": a["giv"], "niv": a["niv"], "month": month,
                    "supply_cs": supply, "demand_cs": cur.get("demand_cs"),
                    "ordered_cs": ordered, "sold_cs": sold, "sold_days": days,
                    "remaining_cs": remaining, "stock_days": round(remaining / daily, 1) if daily else None,
                    "month_sales_proj": month_sales, "remaining_eom": remaining_eom, "stock_days_eom": eom_days,
                    "remaining_giv": round(remaining * a["giv"], 2) if a["giv"] is not None else None,
                    "pg_remaining_supply": round(supply - (ordered or 0), 2) if supply is not None else None})
    out.sort(key=lambda x: (x["brand"] or "", x["product_name"] or ""))
    return jsonify({"month": month, "rows": out, "count": len(out)})


__all__ = ["KIND_LABEL", "detect_kind", "find_source_sheet", "import_supply", "import_master_price", "import_stock",
           "absorb_sheet_extras", "api_sources", "api_month_stats", "stock_by_barcode"]

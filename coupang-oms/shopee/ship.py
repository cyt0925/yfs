"""蝦皮特選竹運出貨（需求第五章）：EC 庫存＋竹運效期判斷 → 出貨量、缺貨 → 訂單系統拋單檔＋下修刪單清單。

判斷邏輯照 Alice 做的「核對竹運出貨效期及庫存 V5.xlsm」（Jerry 2026-10-12：「以這個為主，他們現在就是用這個在判斷庫存能不能出」），
她的邏輯整理 PDF 第 13 節「建議的正式業務定義」也一起照做：
- 三份資料：EC 即時庫存（訂單系統「全商品」，料號＋單位 → EC可用量）、竹運庫存效期表（只看「良品區」，客戶品號＋效期 → 可出貨量，
  同效期不同儲位加總；批號欄是空的，所以一個效期＝一批）、竹運商品主檔（竹運後台「產品基本資料」，客戶品號 → 商品年限(日)；
  也收 Alice 那個 .xlsm，只讀裡面的「商品主檔」頁）。
  客戶品號＝料號＋單位代碼＝蝦皮 Supplier SKU 的料號＋尾碼第一個字（4987176377708_CNN → 4987176377708C）。
- 總效期用商品主檔的商品年限，不用效期表的製造日期（Alice V3.0：「自動避開系統錯誤的製造日期」）。主檔沒有年限的擋下來要補，
  不當成 0 天（PDF 12.4 提到的毛病）。瑪氏主檔沒有的，退回瑪氏商品總表的效期天。
- 合格效期：效期 ≥ 基準日 ＋ 取整數(商品年限 × 比例)。基準日用蝦皮到貨日（需求 5.3；Alice 是手填的出貨日，蝦皮驗的是到倉那天）。
  比例預設照蝦皮供應商進倉規範（總保存天數 90 天內 2/3、365 天內 2/5、超過 1/3），也可以改成自訂一個比例（像 Alice 的 U1，常用 0.66）。
- EC 可用量和竹運合格量都要夠才全出；不夠的話系統建議＝訂購量、EC 剩餘、合格效期量取最小（需求要算可出貨量，Alice 只判可不可出）。
- 由上到下逐列扣庫存（到貨日早的先），同一份庫存不會分給兩張單（PDF 第 9 節）。已產拋單檔的不再分配。
- 多個效期都合格就由早到晚一起扣，不用人再一品一品勾（OP 經 Jerry 2026-10-20：「可以混合多效期出貨，不需要再經過每一次確認」）。
  原本預設「單一效期」（蝦皮同一張入庫單效期要一致、需求 5.3.1 不自動合併），一批不夠就只出量最多那批、等人勾別的效期合併，拆掉了。
  人還是可以點開那品改效期、改數量。
- 缺貨量＝訂購量－最終出貨量。原因：庫存不足（不看效期也不夠）、效期不足（有貨但效期不合格），可同時有。
- 出貨量大於 0 但少於 12 件提醒：蝦皮進倉規範「以低於 12 pcs 數量入倉的商品，將一律拒收」。
- 拋單檔照 Jerry 給的「指定到貨日20261007蝦皮(觀音)_PG.xlsx」：訂單單號＝入庫單號、姓名「蝦皮-觀音」、料號＝國條、數量＝出貨量；出貨量 0 不放。
"""
import datetime as _dt
import zipfile

import openpyxl
import xlrd
from openpyxl.styles import Font, PatternFill

from .common import *  # noqa: F401,F403
from .orders import order_no

FULFIL = "竹運出貨"
KINDS = {"ec": "EC 即時庫存", "lots": "竹運庫存效期表", "life": "竹運商品主檔"}
MIN_PCS = 12
SET_KEY = "shopee_ship"
DEFAULT_SET = {"rule": "shopee", "ratio": 0.66}      # 以前還有 mode（single／mixed），2026-10-20 一律混合效期，舊存的值不理
SHOPEE_TIERS = ((90, 2 / 3, "2/3"), (365, 2 / 5, "2/5"), (None, 1 / 3, "1/3"))
LINE_TAG = {"寶僑": "PG", "紙潔": "紙潔", "瑪氏": "瑪氏"}
DWH_KEY = "shopee_dispatch_wh"
DWH_FIELDS = ("name", "address", "phone")
# 觀音照拋單檔範例；其他倉照蝦皮供應商進倉規範第 6 頁
DWH_DEFAULTS = {
    "TWA": {"name": "蝦皮-觀音", "address": "桃園市觀音區寶倉街108號C棟1樓", "phone": "(03)262-6560"},
    "TWX": {"name": "蝦皮-安南三", "address": "臺南市安定區安新二路688號", "phone": "06-600-1450#5883"},
    "TWT": {"name": "蝦皮-安南一", "address": "台南市安南區新吉一路118號", "phone": "06-600-1450#5871"},
    "TWG": {"name": "蝦皮-高鐵南", "address": "桃園市中壢區高鐵南路四段888號", "phone": "0966-233-806"},
    "TWH": {"name": "蝦皮-楊梅", "address": "桃園市楊梅區和平路576號", "phone": "03-262-6560#3151"},
    "TWW": {"name": "蝦皮-威獅", "address": "桃園市楊梅區梅獅路二段666號", "phone": "0975-213-830"},
    "TWK": {"name": "蝦皮-威獅", "address": "桃園市楊梅區梅獅路二段666號5F", "phone": "0975-213-830"},
}


# ── 讀檔 ─────────────────────────────────────────────────────────────────────
def _h(v):
    return re.sub(r"\s+", "", str(v if v is not None else "")).replace("（", "(").replace("）", ")")


def _code(v):
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return re.sub(r"\s+", "", str(v if v is not None else "")).upper()


def _num(v):
    try:
        return float(str(v).replace(",", "")) if v not in (None, "") else 0.0
    except ValueError:
        return 0.0


def _date(v, datemode=0):
    """效期欄：20290502、2029/05/02、Excel 日期 → '2029-05-02'；看不懂回空字串。"""
    if isinstance(v, (_dt.datetime, _dt.date)):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, float) and 20000000 <= v <= 30000000:
        v = str(int(v))
    elif isinstance(v, float) and v > 0:
        try:
            return xlrd.xldate_as_datetime(v, datemode).strftime("%Y-%m-%d")
        except Exception:  # noqa: BLE001
            return ""
    s = re.sub(r"\D", "", str(v or ""))
    if len(s) == 8:
        try:
            return _dt.date(int(s[:4]), int(s[4:6]), int(s[6:])).isoformat()
        except ValueError:
            return ""
    return ""


def _sheets(data, filename):
    """[(工作表名, 列的 list, datemode)]。.xls 真的是舊格式用 xlrd，其他（含 .xlsm）用 openpyxl。"""
    if data[:2] != b"PK":
        try:
            wb = xlrd.open_workbook(file_contents=data)
        except Exception as exc:  # noqa: BLE001
            raise ValueError(f"{filename}：無法開啟此檔案，請確認是 Excel 檔。") from exc
        return [(sh.name, (sh.row_values(r) for r in range(sh.nrows)), wb.datemode) for sh in wb.sheets()]
    try:
        wb = openpyxl.load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    except Exception as exc:  # noqa: BLE001
        raise ValueError(f"{filename}：無法開啟此檔案，請確認是 Excel 檔。") from exc
    return [(ws.title, (list(r) for r in ws.iter_rows(values_only=True)), 0) for ws in wb.worksheets]


NEEDS = {"life": ("客戶品號", "商品年限(日)"), "lots": ("客戶品號", "效期", "可出貨量"), "ec": ("料號", "單位", "EC可用量")}


def parse_stock(data, filename):
    """認是哪一種、整理好。回 (kind, data, rows_total, file_date, notes)。
    同一個檔有好幾種（Alice 的 .xlsm 三頁都有）只讀商品主檔：她檔裡的庫存是舊的。"""
    found = {}
    for title, rows, dm in _sheets(data, filename):
        for i, row in enumerate(rows):
            heads = [_h(v) for v in (row or [])]
            hit = next((k for k in ("life", "lots", "ec") if all(n in heads for n in NEEDS[k])), None)
            if hit:
                found.setdefault(hit, (title, {h: j for j, h in enumerate(heads) if h}, rows, dm))
                break
            if i >= 5:
                break
        if "life" in found:
            break
    kind = next((k for k in ("life", "lots", "ec") if k in found), None)
    if not kind:
        raise ValueError(f"{filename}：認不出是哪一種檔。EC 即時庫存要有「料號、單位、EC可用量」，竹運庫存效期表要有「客戶品號、效期、可出貨量」，"
                         "竹運商品主檔要有「客戶品號、商品年限(日)」。")
    title, ix, rows, dm = found[kind]
    get = lambda row, k: row[ix[k]] if k in ix and ix[k] < len(row) else None  # noqa: E731
    notes, total, file_date = [], 0, ""
    if kind == "ec":
        out = collections.defaultdict(float)
        for row in rows:
            row = list(row or [])
            code, unit = _code(get(row, "料號")), str(get(row, "單位") or "").strip()
            if not code:
                continue
            total += 1
            file_date = file_date or str(get(row, "結餘日期") or "").strip()
            out[f"{code}|{unit}"] += _num(get(row, "EC可用量"))
        out = {k: (int(v) if float(v).is_integer() else v) for k, v in out.items() if v}
    elif kind == "lots":
        out, skipped, bad = {}, 0, 0
        for row in rows:
            row = list(row or [])
            code = _code(get(row, "客戶品號"))
            if not code:
                continue
            total += 1
            if "儲區" in ix and str(get(row, "儲區") or "").strip() != "良品區":
                skipped += 1
                continue
            exp = _date(get(row, "效期"), dm)
            if not exp:
                bad += 1
                continue
            lot = out.setdefault(code, {}).setdefault(exp, {"qty": 0, "stock": 0, "mfg": "", "locs": 0})
            lot["qty"] += int(_num(get(row, "可出貨量")))
            lot["stock"] += int(_num(get(row, "庫存量")))
            lot["locs"] += 1
            lot["mfg"] = lot["mfg"] or _date(get(row, "製造日期"), dm)
        if skipped:
            notes.append(f"非良品區 {skipped} 列不計")
        if bad:
            notes.append(f"效期看不懂 {bad} 列不計")
    else:
        out, bad = {}, 0
        for row in rows:
            row = list(row or [])
            code = _code(get(row, "客戶品號"))
            if not code:
                continue
            total += 1
            days = _num(get(row, "商品年限(日)"))
            if days > 0:
                out[code] = int(days)
            else:
                bad += 1
        if bad:
            notes.append(f"商品年限空白或不是數字 {bad} 列（這些品項核對時會要求補上）")
        if len(found) > 1:
            notes.append(f"檔案裡另有{'、'.join(KINDS[k] for k in found if k != 'life')}的分頁，只匯入商品主檔")
    return kind, out, total, file_date, notes


def stock_status(conn):
    st = {}
    for r in _rows(conn.execute("SELECT kind, filename, file_date, rows_total, uploaded_by, uploaded_at FROM shp_stock")):
        st[r["kind"]] = dict(r, label=KINDS.get(r["kind"], r["kind"]), today=(r["uploaded_at"] or "")[:10] == now()[:10])
    return st


def load_stock(conn):
    data = {"ec": {}, "lots": {}, "life": {}}
    for r in _rows(conn.execute("SELECT kind, data_json FROM shp_stock")):
        try:
            data[r["kind"]] = json.loads(r["data_json"] or "{}")
        except ValueError:
            pass
    return data


@shopee_bp.route("/api/shopee/stock")
def api_shopee_stock():
    conn = get_conn()
    try:
        return jsonify({"stock": stock_status(conn), "settings": ship_settings(conn)})
    finally:
        conn.close()


@shopee_bp.route("/api/shopee/stock/import", methods=["POST"])
def api_shopee_stock_import():
    f = request.files.get("file")
    if not f or not f.filename:
        return jsonify({"error": "沒有收到檔案。"}), 400
    try:
        kind, out, total, file_date, notes = parse_stock(f.read(), f.filename)
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400
    operator, stamp = _operator(), now()
    conn = get_conn()
    try:
        old = _row(conn.execute("SELECT filename, uploaded_at FROM shp_stock WHERE kind = ?", (kind,)))
        conn.execute("DELETE FROM shp_stock WHERE kind = ?", (kind,))
        conn.execute("INSERT INTO shp_stock (kind, filename, file_date, rows_total, data_json, uploaded_by, uploaded_at) VALUES (?,?,?,?,?,?,?)",
                     (kind, f.filename, file_date, total, json.dumps(out, ensure_ascii=False), operator, stamp))
        _log(conn, "", "", "", "", "shopee_stock", KINDS[kind], (old or {}).get("filename", ""), f"{f.filename}（{total} 列）", operator, "import",
             "；".join(notes))
        conn.commit()
        return jsonify({"ok": True, "kind": kind, "label": KINDS[kind], "rows": total, "items": len(out), "notes": notes, "stock": stock_status(conn)})
    finally:
        conn.close()


# ── 設定 ─────────────────────────────────────────────────────────────────────
def _meta(conn, key):
    row = conn.execute("SELECT value FROM mst_meta WHERE key = ?", (key,)).fetchone()
    try:
        return json.loads(row["value"] or "{}") if row else {}
    except ValueError:
        return {}


def _meta_save(conn, key, value):
    conn.execute("DELETE FROM mst_meta WHERE key = ?", (key,))
    conn.execute("INSERT INTO mst_meta (key, value) VALUES (?, ?)", (key, json.dumps(value, ensure_ascii=False)))


def ship_settings(conn):
    s = dict(DEFAULT_SET)
    s.update({k: v for k, v in _meta(conn, SET_KEY).items() if k in DEFAULT_SET})
    return s


@shopee_bp.route("/api/shopee/ship/settings", methods=["PUT"])
def api_ship_settings():
    p = request.get_json(silent=True) or {}
    conn = get_conn()
    try:
        s = ship_settings(conn)
        new = dict(s)
        if p.get("rule") in ("shopee", "custom"):
            new["rule"] = p["rule"]
        if "ratio" in p:
            try:
                r = float(str(p["ratio"]).strip().rstrip("%"))
            except ValueError:
                return jsonify({"error": "效期比例請填 0 到 1 之間的數字，例如 2/3 填 0.66。"}), 400
            if r > 1:
                r = r / 100
            if not 0 < r <= 1:
                return jsonify({"error": "效期比例請填 0 到 1 之間的數字，例如 2/3 填 0.66。"}), 400
            new["ratio"] = round(r, 4)
        if new != s:
            _meta_save(conn, SET_KEY, new)
            _log(conn, "", "", "", "", "shopee_ship_set", "竹運出貨設定", json.dumps(s, ensure_ascii=False), json.dumps(new, ensure_ascii=False),
                 _operator(), "manual")
            conn.commit()
        return jsonify({"ok": True, "settings": new})
    finally:
        conn.close()


def dispatch_wh(conn):
    saved = _meta(conn, DWH_KEY)
    out = {}
    for code, d in DWH_DEFAULTS.items():
        v = dict(d)
        v.update({k: str(x) for k, x in (saved.get(code) or {}).items() if k in DWH_FIELDS})
        out[code] = v
    return out


@shopee_bp.route("/api/shopee/dispatch_wh")
def api_dispatch_wh():
    conn = get_conn()
    try:
        return jsonify({"warehouses": [{"code": c, "short": wh_name(c), **v} for c, v in dispatch_wh(conn).items()]})
    finally:
        conn.close()


@shopee_bp.route("/api/shopee/dispatch_wh/<code>", methods=["PUT"])
def api_dispatch_wh_save(code):
    code = (code or "").strip().upper()
    if code not in DWH_DEFAULTS:
        return jsonify({"error": f"沒有 {code} 這個蝦皮倉。"}), 404
    p = request.get_json(silent=True) or {}
    conn = get_conn()
    try:
        saved = _meta(conn, DWH_KEY)
        before = dispatch_wh(conn)[code]
        cur = dict(saved.get(code) or {})
        changed = []
        for k in DWH_FIELDS:
            if k in p:
                v = str(p.get(k) or "").strip()
                if v != before[k]:
                    changed.append((k, before[k], v))
                cur[k] = v
        if changed:
            saved[code] = cur
            _meta_save(conn, DWH_KEY, saved)
            labels = {"name": "姓名", "address": "地址", "phone": "電話"}
            for k, old, new in changed:
                _log(conn, "", "", "", "", "shopee_dispatch_wh", f"拋單檔收件資料 {code} {labels[k]}", old, new, _operator(), "manual", wh_name(code))
            conn.commit()
        return jsonify({"ok": True, "changed": len(changed)})
    finally:
        conn.close()


# ── 判斷 ─────────────────────────────────────────────────────────────────────
def ratio_for(days, st):
    if st["rule"] == "custom":
        return float(st["ratio"]), f"自訂 {st['ratio']}"
    for limit, r, text in SHOPEE_TIERS:
        if limit is None or days <= limit:
            return r, text
    return 1 / 3, "1/3"


def _mars_life(conn, items):
    codes = sorted({i["base_code"] for i in items if i["line"] == "瑪氏" and i["base_code"]})
    if not codes:
        return {}
    marks = ",".join("?" * len(codes))
    return {r["yf_sku"] + "|" + r["unit"]: r["shelf_days"] for r in _rows(conn.execute(
        f"SELECT yf_sku, unit, shelf_days FROM mst_mars_products WHERE yf_sku IN ({marks})", codes)) if r["shelf_days"]}


def _dispatched(conn, ids):
    """品項 → 最新一次產拋單檔的快照。"""
    if not ids:
        return {}
    marks = ",".join("?" * len(ids))
    out = {}
    for r in _rows(conn.execute(f"SELECT i.*, d.filename, d.created_at, d.created_by FROM shp_dispatch_items i JOIN shp_dispatch d ON d.id = i.dispatch_id "
                                f"WHERE i.order_id IN ({marks}) ORDER BY i.id", list(ids))):
        out[r["order_id"]] = r
    return out


def _alloc(lots, need):
    """lots: [{exp, qty}] 已照效期早到晚排、只含合格且有量的。由早到晚扣到夠為止，回 [(exp, qty)]。"""
    out, left = [], need
    for l in lots:
        take = min(l["qty"], left)
        if take > 0:
            out.append((l["exp"], take))
            left -= take
        if left <= 0:
            break
    return out


def _cap(alloc, cap):
    out, left = [], cap
    for exp, q in alloc:
        take = min(q, left)
        if take > 0:
            out.append((exp, take))
            left -= take
    return out


def build_ship(conn, d1="", d2="", line="", order_nos=None, up1="", up2=""):
    """竹運出貨的單，照到貨日由早到晚逐列扣 EC 與竹運合格庫存。回 (groups, ctx)。"""
    st = ship_settings(conn)
    stock = load_stock(conn)
    status = stock_status(conn)
    sc, sp = ship_cond()
    where, params = [sc, "removed = 0"], list(sp)
    if d1:
        where.append("expected_date >= ?"); params.append(d1)
    if d2:
        where.append("expected_date <= ?"); params.append(d2)
    usql, up = upload_cond(up1, up2)
    if usql:
        where.append(usql); params += up
    if line:
        where.append("line = ?"); params.append(line)
    items = _rows(conn.execute(f"SELECT * FROM shp_orders WHERE {' AND '.join(where)} ORDER BY expected_date, warehouse, id", params))
    for it in items:
        it["_no"] = order_no(it)
    items.sort(key=lambda i: (i["expected_date"] or "", i["warehouse"] or "", i["_no"], i["id"]))
    ids = [i["id"] for i in items]
    marks = ",".join("?" * len(ids)) if ids else ""
    overrides = {r["order_id"]: r for r in _rows(conn.execute(f"SELECT * FROM shp_ship WHERE order_id IN ({marks})", ids))} if ids else {}
    done = _dispatched(conn, ids)
    mars_life = _mars_life(conn, items)
    ec_left = dict(stock["ec"])
    lots_left = {k: {e: v["qty"] for e, v in d.items()} for k, d in stock["lots"].items()}
    used = {"ec": collections.defaultdict(list), "lots": collections.defaultdict(list)}   # 誰先用掉了：{key: [(訂單編號, 量)]}，畫面要寫出來
    have = {k: k in status for k in KINDS}
    groups = collections.OrderedDict()
    for it in items:
        g = groups.setdefault(it["_no"], {"order_no": it["_no"], "pr_id": it["pr_id"], "po_id": it["po_id"], "line": it["line"], "warehouse": it["warehouse"],
                                          "warehouse_name": wh_name(it["warehouse"]), "expected_date": it["expected_date"], "items": [], "_rows": []})
        g["_rows"].append(it)
        g["items"].append(_ship_item(it, st, stock, ec_left, lots_left, overrides.get(it["id"]), done.get(it["id"]), mars_life, have, used))
    out = []
    for g in groups.values():
        its = g["items"]
        made = [p for p in its if p["dispatch"]]
        g.update(group_meta(g.pop("_rows")))
        g.update(qty=sum(p["qty"] for p in its), ship=sum(p["final"] for p in its), short=sum(p["short"] for p in its),
                 blocking=sum(1 for p in its if p["blocking"]), warnings=sum(1 for p in its if p["warnings"]),
                 deleted=sum(1 for p in its if p["qty"] and not p["final"]),
                 dispatch_status=("已產出" if len(made) == len(its) else "部分產出") if made else "未產出",
                 dispatch_at=max((p["dispatch"]["created_at"] for p in made), default=""))
        out.append(g)
    if order_nos is not None:
        want = set(order_nos)
        out = [g for g in out if g["order_no"] in want]
    return out, {"settings": st, "stock": status, "have": have}


def _ship_item(it, st, stock, ec_left, lots_left, ov, done, mars_life, have, used):
    base, unit = it["base_code"] or "", it["unit_name"] or UNIT_CODES.get(it["unit_code"] or "", "")
    code3 = base + (it["unit_code"] or "")
    qty = it["qty"] or 0
    arrive = _dt.date.fromisoformat(it["expected_date"]) if it["expected_date"] else None
    p = {"order_id": it["id"], "supplier_sku_id": it["supplier_sku_id"], "code": base, "code3": code3, "name": it["sku_name"], "unit": unit,
         "qty": qty, "inbound_ids": it["inbound_ids"] or "", "problems": [], "warnings": [], "blocking": False, "reasons": [], "adj": None,
         "dispatch": None, "lots": [], "alloc": [], "suggest": 0, "final": 0, "short": 0, "ec": None, "book": 0, "qualified": 0, "max_ship": 0,
         "life": None, "life_src": "", "ratio_text": "", "threshold": ""}
    if done:                                       # 已產拋單檔：用當時的快照，不再分配庫存
        p.update(dispatch={"filename": done["filename"], "created_at": done["created_at"], "created_by": done["created_by"]},
                 final=done["qty_ship"], short=max((done["qty_order"] or 0) - (done["qty_ship"] or 0), 0),
                 alloc=json.loads(done["lots_json"] or "[]"),
                 reasons=[r for r in (done["short_reasons"] or "").split("、") if r])
        if done["qty_order"] != qty:
            p["warnings"].append(f"產出拋單檔後蝦皮數量變了（{done['qty_order']} → {qty}），請取消產出後重新判斷")
        return p
    if it["diff_pending"]:
        p["problems"].append("這張單有差異待確認，請先到訂單總覽確認")
    if not have["ec"] or not have["lots"]:
        p["problems"].append("請先上傳 " + "、".join(KINDS[k] for k in ("ec", "lots") if not have[k]))
    life = stock["life"].get(code3)
    if life:
        p["life_src"] = "竹運商品主檔"
    elif it["line"] == "瑪氏" and mars_life.get(f"{base}|{unit}"):
        life, p["life_src"] = int(mars_life[f"{base}|{unit}"]), "瑪氏商品總表"
    p["life"] = life
    e = ec_left.get(f"{base}|{unit}", 0)
    p["ec"] = e
    left = lots_left.get(code3, {})
    p["book"] = sum(v for v in left.values() if v > 0)
    lot_info = stock["lots"].get(code3, {})
    # 畫面上要講清楚「檔案是多少、到貨日較早的幾張單先用了多少」（Jerry 2026-10-20：EC 可用 231，訂單系統卻是 315，以為是單位換算）
    for k, key, total in (("ec", f"{base}|{unit}", stock["ec"].get(f"{base}|{unit}", 0)), ("book", code3, sum(v["qty"] for v in lot_info.values()))):
        taken = used["ec" if k == "ec" else "lots"][key]
        p[f"{k}_file"] = total
        p[f"{k}_used"] = sum(q for _, q in taken)
        p[f"{k}_users"] = len({no for no, _ in taken})
    if not life and min(e, p["book"]) > 0:         # 沒貨就不用判斷效期，直接算庫存不足
        p["problems"].append(f"竹運商品主檔沒有 {code3} 的商品年限，無法判斷效期，請上傳最新的竹運商品主檔")
    if arrive and life:
        ratio, p["ratio_text"] = ratio_for(life, st)
        threshold = arrive + _dt.timedelta(days=int(life * ratio))
        p["threshold"] = threshold.isoformat()
    lots = []
    for exp in sorted(set(left) | set(lot_info)):
        rem = left.get(exp, 0)
        d = _dt.date.fromisoformat(exp)
        days = (d - arrive).days if arrive else None
        ok = bool(life and arrive and d >= threshold)
        lots.append({"exp": exp, "mfg": (lot_info.get(exp) or {}).get("mfg", ""), "qty": rem, "locs": (lot_info.get(exp) or {}).get("locs", 0),
                     "days": days, "pct": round(days / life, 3) if days is not None and life else None, "ok": ok})
    p["lots"] = lots
    qual = [l for l in lots if l["ok"] and l["qty"] > 0]
    p["qualified"] = sum(l["qty"] for l in qual)
    p["max_ship"] = max(0, min(qty, e, p["qualified"])) if life else 0
    need = max(0, min(qty, e))                       # EC 只夠這麼多，就照這個量挑效期
    suggest = _alloc(qual, need)
    p["suggest"] = sum(q for _, q in suggest)
    for i, l in enumerate(sorted(qual, key=lambda x: x["exp"])):      # 建議順位跟系統怎麼挑一致：由早到晚
        l["rank"] = i + 1
    alloc = suggest
    if ov and ov["qty_before"] == qty:
        sel = [(x["exp"], int(x["qty"])) for x in json.loads(ov["lots_json"] or "[]") if int(x["qty"]) > 0]
        p["adj"] = {"qty": ov["qty_final"], "reason": ov["reason"], "operator": ov["operator"], "updated_at": ov["updated_at"], "lots": sel}
        okq = {l["exp"]: l["qty"] for l in qual}
        if ov["qty_final"] > min(qty, e) or any(q > okq.get(exp, 0) for exp, q in sel):
            p["problems"].append("人工確認的出貨量已超過目前可用的庫存（庫存檔更新過，或前面的單先用掉了），請重新確認")
            alloc = _cap(sel, max(0, min(qty, e)))
        else:
            alloc = sel
    p["alloc"] = alloc
    p["final"] = sum(q for _, q in alloc)
    p["short"] = max(qty - p["final"], 0)
    if p["short"] and (life or min(e, p["book"]) <= 0):
        if life and p["qualified"] < qty and p["book"] > p["qualified"]:
            p["reasons"].append("效期不足")
        if min(e, p["book"]) < qty or p["final"] < min(qty, e, p["qualified"]) and not p["adj"]:
            p["reasons"].append("庫存不足")
        if not p["reasons"]:
            p["reasons"].append("庫存不足")
    if 0 < p["final"] < MIN_PCS:
        p["warnings"].append(f"出貨量 {p['final']} 少於 {MIN_PCS} 件，蝦皮進倉規範會拒收，建議改成 0（刪單）")
    if not p["inbound_ids"] and p["final"]:
        p["problems"].append("還沒有入庫單號，請先上傳蝦皮入庫單")
    elif "," in p["inbound_ids"] and p["final"]:
        p["problems"].append(f"這品有 {p['inbound_ids'].count(',') + 1} 張入庫單（{p['inbound_ids']}），拋單檔請手動拆")
    p["blocking"] = bool(p["problems"])
    for exp, q in alloc:                                              # 扣掉，後面的單看到的是剩下的
        left[exp] = left.get(exp, 0) - q
    lots_left[code3] = left
    ec_left[f"{base}|{unit}"] = e - p["final"]
    if p["final"]:
        used["ec"][f"{base}|{unit}"].append((it["_no"], p["final"]))
        used["lots"][code3].append((it["_no"], p["final"]))
    p["alloc"] = [{"exp": exp, "qty": q} for exp, q in alloc]
    return p


@shopee_bp.route("/api/shopee/ship/plan")
def api_ship_plan():
    conn = get_conn()
    try:
        orders, ctx = build_ship(conn, request.args.get("from", ""), request.args.get("to", ""), request.args.get("line", ""),
                                 up1=request.args.get("up_from", ""), up2=request.args.get("up_to", ""))
    finally:
        conn.close()
    return jsonify({"orders": orders, **ctx})


@shopee_bp.route("/api/shopee/ship/adjust", methods=["POST"])
def api_ship_adjust():
    """改出貨量／選效期（需求 5.5）。lots＝[{exp, qty}]，加總要等於出貨量；不能超過可出貨量；要寫原因。clear＝回到系統建議。"""
    p = request.get_json(silent=True) or {}
    try:
        oid = int(p.get("order_id"))
    except (TypeError, ValueError):
        return jsonify({"error": "缺少品項。"}), 400
    operator = _operator()
    conn = get_conn()
    try:
        it = _row(conn.execute("SELECT * FROM shp_orders WHERE id = ?", (oid,)))
        if it is None:
            return jsonify({"error": "找不到這個品項，請重新整理頁面。"}), 404
        no = order_no(it)
        old = _row(conn.execute("SELECT * FROM shp_ship WHERE order_id = ?", (oid,)))
        if p.get("clear"):
            if old:
                conn.execute("DELETE FROM shp_ship WHERE order_id = ?", (oid,))
                _log(conn, it["line"], no, it["shopee_sku_id"], "", "shopee_ship_qty", "竹運出貨量", old["qty_final"], "回到系統建議", operator, "manual",
                     it["supplier_sku_id"])
                conn.commit()
            return jsonify({"ok": True, "cleared": True})
        reason = str(p.get("reason") or "").strip()
        try:
            lots = [{"exp": str(x["exp"]), "qty": int(x["qty"])} for x in (p.get("lots") or []) if int(x.get("qty") or 0) > 0]
        except (TypeError, ValueError, KeyError):
            return jsonify({"error": "選用數量請填整數。"}), 400
        total = sum(x["qty"] for x in lots)
        if any(x["qty"] < 0 for x in lots):
            return jsonify({"error": "選用數量不能小於 0。"}), 400
        orders, _ctx = build_ship(conn, p.get("from", ""), p.get("to", ""), p.get("line", ""), [no])
        item = next((x for g in orders for x in g["items"] if x["order_id"] == oid), None)
        if item is None:
            return jsonify({"error": "這個品項不在目前的竹運出貨清單裡，請重新整理頁面。"}), 400
        if item["dispatch"]:
            return jsonify({"error": "這張單已產出拋單檔，要修改請先取消產出。"}), 400
        if not item["life"]:
            return jsonify({"error": "商品主檔沒有這個品項的商品年限，無法判斷效期，請先上傳竹運商品主檔。"}), 400
        avail = {l["exp"]: l["qty"] for l in item["lots"] if l["ok"]}
        for x in lots:
            if x["exp"] not in avail:
                return jsonify({"error": f"效期 {x['exp']} 不符合效期標準，不能選。"}), 400
            if x["qty"] > avail[x["exp"]]:
                return jsonify({"error": f"效期 {x['exp']} 只剩 {avail[x['exp']]}，不能選 {x['qty']}。"}), 400
        cap = min(item["qty"], item["ec"] or 0)
        if total > cap:
            return jsonify({"error": f"出貨量 {total} 超過可出貨量 {min(cap, item['qualified'])}（訂購量 {item['qty']}、EC 可用 {item['ec']}）。"}), 400
        # 原因統一在右側卡片「調整原因」那一區填（Jerry 2026-10-21）：效期視窗不再填原因。沒帶原因就沿用這品原本的，
        # 這品沒有就用這張單其他改過的品項最常用的原因（卡片那區填過，之後再改數量的自動帶）
        if not reason:
            reason = (old or {}).get("reason") or _order_reason(conn, it, oid)
        _write_ship(conn, it, no, item, lots, reason, operator, bool(old))
        conn.commit()
        return jsonify({"ok": True, "qty": total})
    finally:
        conn.close()


def _write_ship(conn, it, no, item, lots, reason, operator, exists):
    """存人工確認的出貨量（效期、數量、原因）＋記歷程。lots＝[{exp, qty}]。"""
    total = sum(x["qty"] for x in lots)
    stamp = now(); oid = it["id"]
    lj = json.dumps(lots, ensure_ascii=False)
    if exists:
        conn.execute("UPDATE shp_ship SET qty_before = ?, qty_final = ?, lots_json = ?, reason = ?, operator = ?, updated_at = ? WHERE order_id = ?",
                     (it["qty"], total, lj, reason, operator, stamp, oid))
    else:
        conn.execute("INSERT INTO shp_ship (order_id, qty_before, qty_final, lots_json, reason, operator, updated_at) VALUES (?,?,?,?,?,?,?)",
                     (oid, it["qty"], total, lj, reason, operator, stamp))
    _log(conn, it["line"], no, it["shopee_sku_id"], "", "shopee_ship_qty", "竹運出貨量", item["final"], total, operator, "manual",
         f"{it['supplier_sku_id']}；" + ("、".join(f"{x['exp']}×{x['qty']}" for x in lots) or "改成 0（刪單）") + (f"；{reason}" if reason else ""), reason)
    return total


def _order_reason(conn, it, oid):
    """這張單其他改過出貨量的品項最常用的調整原因（沒有就空白）。"""
    rows = _rows(conn.execute("SELECT s.reason FROM shp_ship s JOIN shp_orders o ON o.id = s.order_id "
                              "WHERE s.order_id != ? AND s.reason != '' AND o.po_id = ? AND o.pr_id = ?", (oid, it["po_id"], it["pr_id"])))
    cnt = collections.Counter(r["reason"] for r in rows)
    return cnt.most_common(1)[0][0] if cnt else ""


def _ids_of(p):
    out = []
    for x in (p.get("order_ids") or []):
        try:
            out.append(int(x))
        except (TypeError, ValueError):
            pass
    return out


@shopee_bp.route("/api/shopee/ship/zero", methods=["POST"])
def api_ship_zero():
    """改成 0（刪單）：勾到的品項全部出 0，原因一起填。右側卡片那張單「少於 12 件的 N 項改成 0」用（Jerry 2026-10-21：做在右側卡片）。已產拋單檔的跳過。"""
    p = request.get_json(silent=True) or {}
    ids = _ids_of(p)
    reason = str(p.get("reason") or "").strip()
    if not ids:
        return jsonify({"error": "請勾選品項。"}), 400
    operator = _operator(); n = 0
    conn = get_conn()
    try:
        marks = ",".join("?" * len(ids))
        its = _rows(conn.execute(f"SELECT * FROM shp_orders WHERE id IN ({marks})", ids))
        nos = sorted({order_no(i) for i in its})
        orders, _ctx = build_ship(conn, p.get("from", ""), p.get("to", ""), p.get("line", ""), nos) if nos else ([], None)
        items = {x["order_id"]: x for g in orders for x in g["items"]}
        for it in its:
            item = items.get(it["id"])
            if item is None or item["dispatch"]:
                continue
            exists = _row(conn.execute("SELECT 1 AS x FROM shp_ship WHERE order_id = ?", (it["id"],))) is not None
            _write_ship(conn, it, order_no(it), item, [], reason, operator, exists); n += 1
        conn.commit()
    finally:
        conn.close()
    return jsonify({"ok": True, "count": n})


@shopee_bp.route("/api/shopee/ship/reasons", methods=["POST"])
def api_ship_reasons():
    """一次填原因：好幾項改過出貨量的，同一個調整原因一起套（Alice 2026-10-21）。"""
    p = request.get_json(silent=True) or {}
    ids = _ids_of(p)
    reason = str(p.get("reason") or "").strip()
    if not ids:
        return jsonify({"error": "請勾選品項。"}), 400
    if not reason:
        return jsonify({"error": "請選一個原因，或自己打。"}), 400
    operator = _operator(); stamp = now(); n = 0
    conn = get_conn()
    try:
        marks = ",".join("?" * len(ids))
        rows = _rows(conn.execute(f"SELECT s.order_id, s.reason AS old_reason, o.line, o.po_id, o.pr_id, o.shopee_sku_id, o.supplier_sku_id "
                                  f"FROM shp_ship s JOIN shp_orders o ON o.id = s.order_id WHERE s.order_id IN ({marks})", ids))
        for r in rows:
            if (r["old_reason"] or "") == reason:
                continue
            conn.execute("UPDATE shp_ship SET reason = ?, operator = ?, updated_at = ? WHERE order_id = ?", (reason, operator, stamp, r["order_id"]))
            _log(conn, r["line"], r["po_id"] or r["pr_id"], r["shopee_sku_id"], "", "shopee_ship_reason", "竹運出貨調整原因", r["old_reason"] or "", reason,
                 operator, "manual", r["supplier_sku_id"], reason)
            n += 1
        conn.commit()
    finally:
        conn.close()
    return jsonify({"ok": True, "count": n})


# ── 產檔 ─────────────────────────────────────────────────────────────────────
DISPATCH_HEAD = ["訂單單號", "姓名", "地址", "電話", "手機", "出貨備註", "料號", "產品名稱", "單位", "數量"]


def dispatch_xlsx(rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "工作表1"
    ws.append(DISPATCH_HEAD)
    for r in rows:
        ws.append([r["inbound_id"], r["name"], r["address"], r["phone"], None, None, str(r["code"]), r["product"], r["unit"], r["qty"]])
    for row in ws.iter_rows(min_row=1, min_col=7, max_col=7):
        for c in row:
            c.number_format = "@"
    for col, w in zip("ABCDEFGHIJ", (20, 10, 30, 13, 8, 10, 16, 60, 6, 7)):
        ws.column_dimensions[col].width = w
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def short_xlsx(rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "下修刪單清單"
    head = ["處理", "到貨日", "倉別", "訂單編號", "入庫單號", "廠商料號", "品名", "單位", "原始訂購量", "最終出貨量", "下修數量", "缺貨原因", "調整原因"]
    ws.append(head)
    for c in ws[1]:
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", fgColor="FBE4DC")
    for r in rows:
        ws.append([r["action"], r["date"], r["wh"], r["order_no"], r["inbound_id"], r["sku"], r["name"], r["unit"], r["qty"], r["final"], r["short"],
                   r["reasons"], r["adj_reason"]])
    for col, w in zip("ABCDEFGHIJKLM", (6, 11, 8, 22, 20, 18, 50, 6, 10, 10, 9, 16, 24)):
        ws.column_dimensions[col].width = w
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _md(d):
    return f"{d[5:7]}{d[8:10]}" if d and len(d) >= 10 else "日期不明"


@shopee_bp.route("/api/shopee/ship/generate", methods=["POST"])
def api_ship_generate():
    """勾選的單 → zip（拋單檔一天一倉一線別一個＋下修刪單清單）。有要處理的就整批不產。"""
    p = request.get_json(silent=True) or {}
    nos = [str(x) for x in (p.get("order_nos") or []) if str(x).strip()]
    if not nos:
        return jsonify({"error": "請先勾選要產出拋單檔的訂單。"}), 400
    operator, stamp = _operator(), now()
    conn = get_conn()
    try:
        orders, ctx = build_ship(conn, p.get("from", ""), p.get("to", ""), p.get("line", ""), nos)
        if not orders:
            return jsonify({"error": "勾選的訂單已不在竹運出貨清單裡，請重新整理頁面。"}), 400
        orders = [g for g in orders if g["dispatch_status"] != "已產出"]
        if not orders:
            return jsonify({"error": "勾選的訂單都已產出拋單檔；要重新產出請先取消產出。"}), 400
        bad = [f"{g['order_no']} {x['supplier_sku_id']}：{'；'.join(x['problems'])}" for g in orders for x in g["items"] if x["blocking"] and not x["dispatch"]]
        if bad:
            return jsonify({"error": f"有 {len(bad)} 個品項需要先處理，無法產出。", "details": bad[:30]}), 400
        whs = dispatch_wh(conn)
        files, shorts = collections.OrderedDict(), []
        for g in orders:
            w = whs.get(g["warehouse"]) or {"name": f"蝦皮-{wh_name(g['warehouse'])}", "address": "", "phone": ""}
            for x in g["items"]:
                if x["dispatch"]:
                    continue
                key = (g["expected_date"] or "", g["warehouse"] or "", g["line"] or "")
                f = files.setdefault(key, {"rows": [], "items": []})
                f["items"].append((g, x))
                if x["final"]:
                    f["rows"].append({"inbound_id": x["inbound_ids"], "name": w["name"], "address": w["address"], "phone": w["phone"], "code": x["code"],
                                      "product": x["name"], "unit": x["unit"], "qty": x["final"]})
                if x["short"]:
                    shorts.append({"action": "刪單" if not x["final"] else "下修", "date": g["expected_date"], "wh": g["warehouse_name"], "order_no": g["order_no"],
                                   "inbound_id": x["inbound_ids"], "sku": x["supplier_sku_id"], "name": x["name"], "unit": x["unit"], "qty": x["qty"],
                                   "final": x["final"], "short": x["short"], "reasons": "、".join(x["reasons"]), "adj_reason": (x["adj"] or {}).get("reason", "")})
        buf = io.BytesIO()
        names = []
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for (date, wh, line), f in files.items():
                name = f"指定到貨日{date.replace('-', '')}蝦皮({wh_name(wh)})_{LINE_TAG.get(line, line)}.xlsx"
                if f["rows"]:
                    zf.writestr(name, dispatch_xlsx(f["rows"]))
                    names.append(name)
                cur = conn.execute("INSERT INTO shp_dispatch (batch, line, warehouse, expected_date, filename, rows_json, created_by, created_at) VALUES (?,?,?,?,?,?,?,?)",
                                   (stamp, line, wh, date, name if f["rows"] else "（全部刪單，無拋單檔）", json.dumps(f["rows"], ensure_ascii=False), operator, stamp))
                did = cur.lastrowid
                for g, x in f["items"]:
                    conn.execute("INSERT INTO shp_dispatch_items (dispatch_id, order_id, qty_order, qty_ship, lots_json, short_reasons) VALUES (?,?,?,?,?,?)",
                                 (did, x["order_id"], x["qty"], x["final"], json.dumps(x["alloc"], ensure_ascii=False), "、".join(x["reasons"])))
            dates = sorted({k[0] for k in files})
            tag = _md(dates[0]) + (f"-{_md(dates[-1])}" if dates and dates[-1] != dates[0] else "")
            if shorts:
                sname = f"蝦皮特選下修刪單清單_{tag}.xlsx"
                zf.writestr(sname, short_xlsx(shorts))
                names.append(sname)
        for g in orders:
            _log(conn, g["line"], g["order_no"], "", "", "shopee_ship", "產出竹運拋單檔", "", f"出貨 {g['ship']}／訂購 {g['qty']}，缺 {g['short']}", operator, "manual",
                 "、".join(n for n in names if n.startswith("指定到貨日")))
        conn.commit()
    finally:
        conn.close()
    buf.seek(0)
    resp = send_file(buf, as_attachment=True, download_name=f"蝦皮特選竹運出貨_{tag}.zip", mimetype="application/zip")
    resp.headers["X-Shopee-Files"] = str(len(names))
    resp.headers["X-Shopee-Short"] = str(len(shorts))
    return resp


@shopee_bp.route("/api/shopee/ship/undo", methods=["POST"])
def api_ship_undo():
    """取消產出：這張單的品項回到「未產出」，重新照目前庫存判斷。拋單檔已經匯進訂單系統的話要自己去那邊刪。"""
    p = request.get_json(silent=True) or {}
    no = str(p.get("order_no") or "").strip()
    conn = get_conn()
    try:
        sc, sp = ship_cond()
        items = [i for i in _rows(conn.execute(f"SELECT * FROM shp_orders WHERE {sc} AND removed = 0", sp)) if order_no(i) == no]
        if not items:
            return jsonify({"error": "找不到這張單。"}), 404
        ids = [i["id"] for i in items]
        marks = ",".join("?" * len(ids))
        dids = {r["dispatch_id"] for r in _rows(conn.execute(f"SELECT dispatch_id FROM shp_dispatch_items WHERE order_id IN ({marks})", ids))}
        if not dids:
            return jsonify({"ok": True, "removed": 0})
        conn.execute(f"DELETE FROM shp_dispatch_items WHERE order_id IN ({marks})", ids)
        for did in dids:
            if conn.execute("SELECT COUNT(*) AS n FROM shp_dispatch_items WHERE dispatch_id = ?", (did,)).fetchone()["n"] == 0:
                conn.execute("DELETE FROM shp_dispatch WHERE id = ?", (did,))
        _log(conn, items[0]["line"], no, "", "", "shopee_ship", "取消竹運拋單檔", "已產出", "未產出", _operator(), "manual")
        conn.commit()
        return jsonify({"ok": True, "removed": len(ids)})
    finally:
        conn.close()


def ship_status_map(conn, ids):
    """訂單總覽用：品項 → 拋單檔快照（出貨量、缺貨）。"""
    return _dispatched(conn, ids)


__all__ = ["parse_stock", "stock_status", "load_stock", "ship_settings", "dispatch_wh", "ratio_for", "build_ship", "dispatch_xlsx", "short_xlsx",
           "ship_status_map", "api_shopee_stock", "api_shopee_stock_import", "api_ship_settings", "api_dispatch_wh", "api_dispatch_wh_save",
           "api_ship_plan", "api_ship_adjust", "api_ship_generate", "api_ship_undo"]

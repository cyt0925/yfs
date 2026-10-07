"""讀蝦皮後台下載的兩種檔：採購單（PurchaseOrder_…xlsx）、入庫單（InboundOrder_…xlsx，工作表 ASN Detail）。

兩種檔都認表頭文字、不認欄位位置（蝦皮改欄位順序也不會讀錯）。是哪一種看表頭：
有「PR ID」「Shopee Requested Qty (unit)」是採購單；有「Inbound ID」「Expected Qty (Unit)」是入庫單。
"""
import datetime as _dt

from purchase import _load_workbook

from .common import UNIT_CODES, collections, line_of, split_supplier_sku


class ParseError(Exception):
    pass


PURCHASE_COLS = {   # 我們要的欄 → 蝦皮表頭
    "pr_id": "PR ID", "po_id": "PO ID", "warehouse": "Warehouse", "pr_created": "PR Creation Date",
    "po_created": "PO Creation Date", "expected_date": "Expected Delivery Time", "shopee_sku_id": "Shopee SKU ID",
    "supplier_sku_id": "Supplier SKU ID", "sku_name": "Shopee SKU Name", "qty": "Shopee Requested Qty (unit)",
    "ean": "EAN / UPC", "selling_type": "Selling Type", "inbound_id": "Inbound ID",
}
INBOUND_COLS = {
    "inbound_id": "Inbound ID", "shopee_sku_id": "Shopee SKU ID", "supplier_sku_id": "Supplier SKU ID",
    "sku_name": "SKU Name", "po_id": "PO ID", "expected_date": "Estimated purchase date",
    "qty": "Expected Qty (Unit)", "warehouse": "Warehouse", "selling_type": "Selling Type",
    "asn_status": "ASN Status", "purchase_type": "Purchase Type",
}
REQUIRED = {"purchase": ("pr_id", "shopee_sku_id", "supplier_sku_id", "qty", "expected_date", "warehouse"),
            "inbound": ("inbound_id", "po_id", "shopee_sku_id", "supplier_sku_id", "qty")}
KIND_TEXT = {"purchase": "採購單", "inbound": "入庫單"}


def _cell(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    s = str(v).strip()
    return "" if s.lower() in ("none", "/") else s


def _date(v):
    if isinstance(v, (_dt.datetime, _dt.date)):
        return v.strftime("%Y-%m-%d")
    s = _cell(v)[:10].replace("/", "-")
    try:
        return _dt.date.fromisoformat(s).isoformat()
    except ValueError:
        return ""


def _qty(v):
    s = _cell(v).replace(",", "")
    try:
        f = float(s)
    except ValueError:
        return None
    return int(f) if f.is_integer() and f >= 0 else None


def _find_table(wb):
    """回（工作表, 表頭那一列的編號, 表頭 list, kind）。表頭可能不在第一列，往下找幾列。"""
    for ws in wb.worksheets:
        for i, row in enumerate(ws.iter_rows(min_row=1, max_row=6, values_only=True), start=1):
            hdr = [_cell(v) for v in row]
            if "PR ID" in hdr and "Shopee Requested Qty (unit)" in hdr:
                return ws, i, hdr, "purchase"
            if "Inbound ID" in hdr and "Expected Qty (Unit)" in hdr:
                return ws, i, hdr, "inbound"
    return None, 0, [], ""


def parse_file(file_bytes, filename=""):
    """回 {"kind", "rows", "errors", "lines"}。rows 已經補上 base_code、unit_code、unit_name、line。
    同一張單（採購單看 PR＋蝦皮商品編號；入庫單看 PO＋蝦皮商品編號）出現兩列時合成一列：
    入庫單會這樣（同一品項兩個效期、分兩張入庫單），數量加總、入庫單號串起來。"""
    try:
        wb = _load_workbook(file_bytes)
    except Exception as exc:  # noqa: BLE001 — 原始英文錯誤不給畫面看
        raise ParseError("無法開啟檔案，請確認是蝦皮後台下載的 Excel 檔（.xlsx）。") from exc
    ws, hrow, hdr, kind = _find_table(wb)
    if not kind:
        raise ParseError("找不到蝦皮採購單或入庫單的表頭（需有「PR ID」或「Inbound ID」欄）。請確認是蝦皮後台下載的原始檔。")
    cols = PURCHASE_COLS if kind == "purchase" else INBOUND_COLS
    idx = {k: hdr.index(v) for k, v in cols.items() if v in hdr}
    missing = [cols[k] for k in REQUIRED[kind] if k not in idx]
    if missing:
        raise ParseError(f"{KIND_TEXT[kind]}缺少欄位：{'、'.join(missing)}。")

    rows, errors, merged = [], [], collections.OrderedDict()
    for n, raw in enumerate(ws.iter_rows(min_row=hrow + 1, values_only=True), start=hrow + 1):
        vals = list(raw)
        if not any(_cell(v) for v in vals):
            continue
        get = lambda k: vals[idx[k]] if k in idx and idx[k] < len(vals) else None  # noqa: E731
        r = {k: _cell(get(k)) for k in cols}
        r["expected_date"] = _date(get("expected_date"))
        r["qty"] = _qty(get("qty"))
        r["row"] = n
        miss = [cols[k] for k in REQUIRED[kind] if k != "qty" and not r.get(k)]
        if miss:
            errors.append(f"第 {n} 列缺少：{'、'.join(miss)}，已略過。")
            continue
        if r["qty"] is None:
            errors.append(f"第 {n} 列（{r['supplier_sku_id']}）數量「{_cell(get('qty'))}」不是正整數，已略過。")
            continue
        base, ucode = split_supplier_sku(r["supplier_sku_id"])
        r.update(base_code=base, unit_code=ucode, unit_name=UNIT_CODES.get(ucode, ""), line=line_of(base))
        key = (r["pr_id"] if kind == "purchase" else r["po_id"], r["shopee_sku_id"])
        if key in merged:
            m = merged[key]
            m["qty"] += r["qty"]
            if r.get("inbound_id") and r["inbound_id"] not in m.get("inbound_id", "").split(","):
                m["inbound_id"] = ",".join(x for x in (m.get("inbound_id"), r["inbound_id"]) if x)
            continue
        merged[key] = r
    rows = list(merged.values())

    # 認不出線別的（例如瑪氏的 DH56M_QMN），用同一個檔裡最多的那個線別補，並標出來
    counts = collections.Counter(r["line"] for r in rows if r["line"])
    major = counts.most_common(1)[0][0] if counts else ""
    for r in rows:
        if not r["line"]:
            r["line"] = major
            r["line_guessed"] = True
            errors.append(f"第 {r['row']} 列（{r['supplier_sku_id']}）看料號認不出線別，依同檔其他品項當成「{major or '未知'}」。")
    return {"kind": kind, "rows": rows, "errors": errors, "lines": dict(collections.Counter(r["line"] for r in rows))}


__all__ = ["ParseError", "parse_file", "KIND_TEXT"]

"""讀 MOMO 後台下載的未出貨清單（UnsendList，需求第二、三章）。

瑪氏業務同意直接用這份（2026-10-27 開會）；紙品貼在「紙箱購mo回單確認」、潔品依倉別貼在「momo實際開單數量」的，也是這份的內容，
所以系統只認這一種格式：表頭用文字認（入庫單號、採購數量、商品原廠編號…），不看欄位字母——紙品那份前面多一欄查詢公式也讀得到。
- 單據層的欄位（倉別、入庫單號、入庫預定日期、採購日、採購總量）只寫在每張單的第一列，往下帶（需求三）。
- 採購數量是明細數量；採購總量只拿來核對同一張單的明細加總（需求三），對不上列警告。
- 數量空白、0、錯誤值、不是數字 → 不算補貨明細；入庫單號、倉別、到貨日、商品原廠編號缺了 → 錯誤清單（檔名、工作表、列號、倉別、原因）。
- 單位代碼：商品原廠編號底線後第一個字照第八章那張表換，不在表上的列錯誤，不猜。
- 編號、料號、單號一律當文字，前面的 0 不能掉。
"""
import datetime as _dt
import io
import re

import openpyxl
import xlrd

from .common import UNIT_CODES, line_of, split_code

NEED = ("入庫單號", "採購數量", "商品原廠編號", "商品編號")
COLS = {"倉別": "warehouse", "入庫單號": "inbound_no", "入庫預定日期": "expected_date", "採購日": "purchase_date", "採購總量": "total",
        "momo採購備註": "momo_note", "商品編號": "product_no", "商品名稱": "product_name", "單品編號": "item_no", "單品詳細": "spec",
        "保存天數": "shelf_days", "入庫最低有效期限": "min_expiry", "採購數量": "qty", "商品原廠編號": "factory_code"}
DOC_LEVEL = ("warehouse", "inbound_no", "expected_date", "purchase_date", "total", "momo_note")
BAD = ("#N/A", "#REF!", "#VALUE!", "#DIV/0!", "#NAME?", "#NUM!", "#NULL!")


class ParseError(Exception):
    pass


def _text(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    if isinstance(v, (_dt.datetime, _dt.date)):
        return v.strftime("%Y-%m-%d")
    return str(v).replace("\t", "").strip()


def _date(v):
    s = _text(v)
    if not s or s in BAD:
        return ""
    m = re.match(r"^(\d{4})[/\-.](\d{1,2})[/\-.](\d{1,2})", s)
    if not m:
        return ""
    try:
        return _dt.date(int(m.group(1)), int(m.group(2)), int(m.group(3))).isoformat()
    except ValueError:
        return ""


def _num(v):
    s = _text(v).replace(",", "")
    if not s or s in BAD:
        return None
    try:
        n = float(s)
    except ValueError:
        return None
    return int(n) if n.is_integer() else n


def _sheets(data, filename):
    """[(工作表名, [[值…], …])]；.xls 用 xlrd，其他用 openpyxl（只讀值）。"""
    name = (filename or "").lower()
    if data[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" or name.endswith(".xls"):
        try:
            wb = xlrd.open_workbook(file_contents=data)
        except Exception as exc:  # noqa: BLE001
            raise ParseError(f"{filename}：無法開啟此檔案。") from exc
        out = []
        for s in wb.sheets():
            rows = []
            for r in range(s.nrows):
                row = []
                for c in range(s.ncols):
                    v = s.cell_value(r, c)
                    if s.cell_type(r, c) == xlrd.XL_CELL_DATE:
                        try:
                            v = xlrd.xldate_as_datetime(v, wb.datemode)
                        except Exception:  # noqa: BLE001
                            pass
                    row.append(v)
                rows.append(row)
            out.append((s.name, rows))
        return out
    try:
        wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True, read_only=True)
    except Exception as exc:  # noqa: BLE001
        raise ParseError(f"{filename}：無法開啟此檔案，請確認是 MOMO 下載的 Excel。") from exc
    return [(ws.title, [list(r) for r in ws.iter_rows(values_only=True)]) for ws in wb.worksheets]


def _header(rows):
    """前 5 列找表頭；回 (列號, {欄位: 第幾欄})，找不到回 (None, None)。"""
    for i, row in enumerate(rows[:5]):
        heads = [_text(v).replace("\n", "") for v in row]
        if all(h in heads for h in NEED):
            return i, {COLS[h]: j for j, h in enumerate(heads) if h in COLS}
    return None, None


def looks_like(data, filename):
    """這是不是 MOMO 下載檔（上傳入口認檔用）。"""
    try:
        return any(_header(rows)[0] is not None for _t, rows in _sheets(data, filename))
    except ParseError:
        return False


def parse_file(data, filename):
    """→ {"rows": [...], "errors": [...], "warnings": [...], "sheet": 工作表名}。一個檔只讀第一張有 MOMO 表頭的工作表。"""
    for title, rows in _sheets(data, filename):
        hi, idx = _header(rows)
        if hi is not None:
            break
    else:
        raise ParseError(f"{filename}：找不到 MOMO 下載檔的表頭（入庫單號、採購數量、商品原廠編號…）。請上傳 MOMO 後台下載的未出貨清單。")
    out, errors, warnings = [], [], []
    doc = {k: "" for k in DOC_LEVEL}
    sums = {}
    for r in range(hi + 1, len(rows)):
        row = rows[r]
        get = lambda k: row[idx[k]] if k in idx and idx[k] < len(row) else None  # noqa: E731
        if _text(get("inbound_no")):                              # 新的一張單：單據層欄位換掉
            doc = {"warehouse": _text(get("warehouse")), "inbound_no": _text(get("inbound_no")),
                   "expected_date": _date(get("expected_date")), "purchase_date": _date(get("purchase_date")),
                   "total": _num(get("total")), "momo_note": _text(get("momo_note")), "_raw_date": _text(get("expected_date"))}
        code, prod = _text(get("factory_code")), _text(get("product_no"))
        qty = _num(get("qty"))
        if not code and not prod and qty is None:
            continue                                               # 空白列
        where = {"file": filename, "sheet": title, "row": r + 1, "warehouse": doc.get("warehouse", "")}
        if not qty or qty <= 0:                                    # 空白、0、錯誤值、不是數字：不是補貨明細（需求三）
            continue
        why = []
        if not doc.get("inbound_no") or doc["inbound_no"] in BAD:
            why.append("入庫單號空白或錯誤")
        if not doc.get("warehouse") or doc["warehouse"] in BAD:
            why.append("倉別空白或錯誤")
        if not doc.get("expected_date"):
            why.append(f"入庫預定日期「{doc.get('_raw_date') or '空白'}」不是日期")
        if not code or code in BAD:
            why.append("商品原廠編號空白")
        base, ucode = split_code(code)
        if code and code not in BAD and not ucode:
            why.append(f"商品原廠編號「{code}」沒有底線，認不出單位")
        elif ucode and ucode not in UNIT_CODES:
            why.append(f"商品原廠編號「{code}」的單位代碼 {ucode} 不在單位代碼清單上")
        if not isinstance(qty, int):
            why.append(f"採購數量 {qty} 不是整數")
        if why:
            errors.append({**where, "reason": "；".join(why)})
            continue
        out.append({"inbound_no": doc["inbound_no"], "warehouse": doc["warehouse"], "expected_date": doc["expected_date"],
                    "purchase_date": doc["purchase_date"], "momo_note": doc["momo_note"],
                    "product_no": prod, "item_no": _text(get("item_no")), "product_name": _text(get("product_name")),
                    "spec": _text(get("spec")), "shelf_days": _text(get("shelf_days")), "min_expiry": _date(get("min_expiry")),
                    "qty": qty, "factory_code": code, "base_code": base, "unit_code": ucode, "unit_name": UNIT_CODES[ucode],
                    "line": line_of(base), "source_file": filename, "source_sheet": title, "source_row": r + 1})
        s = sums.setdefault(doc["inbound_no"], {"total": doc["total"], "sum": 0, "row": r + 1})
        s["sum"] += qty
    for no, s in sums.items():
        if s["total"] is not None and s["total"] != s["sum"]:
            warnings.append(f"{filename}：入庫單 {no} 採購總量 {s['total']}，明細加總 {s['sum']}，對不上，請確認檔案是否完整")
    return {"rows": out, "errors": errors, "warnings": warnings, "sheet": title}


__all__ = ["ParseError", "parse_file", "looks_like"]

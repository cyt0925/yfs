"""欄位正規化：匯入、比對、查詢全部走這裡，不允許有第二份實作。

只要交期 / 倉別 / SKU 參與識別或比對，任何格式差異都會製造出幽靈列，
所以正規化是安全機制而不是美觀問題。
"""

import datetime as _dt
import re
import unicodedata

# Excel 1900 日期系統的原點（含 Excel 的閏年 bug，序號 1 = 1900-01-01）
_EXCEL_EPOCH = _dt.date(1899, 12, 30)

_DATE_PATTERNS = (
    "%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d",
    "%Y-%m-%d %H:%M:%S", "%Y/%m/%d %H:%M:%S",
    "%Y%m%d", "%m/%d/%Y", "%d/%m/%Y",
)


def norm_text(value):
    """全形轉半形、去除換行與前後空白、收斂連續空白。None -> ''。"""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    text = str(value)
    text = unicodedata.normalize("NFKC", text)
    text = text.replace("　", " ")
    text = re.sub(r"[\r\n\t]+", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def norm_key(value):
    """識別用文字（PO 單號、SKU ID、料號）。

    Excel 會把 13 位數的 PO 單號讀成 float 並顯示成科學記號，
    這裡一律還原成不帶小數點的純文字，避免同一張單產生兩把鍵。
    """
    if value is None:
        return ""
    if isinstance(value, bool):
        return ""
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
        return repr(value)
    text = norm_text(value)
    # 科學記號還原：1.3E+13 -> 13000000000000
    if re.fullmatch(r"[0-9]+(\.[0-9]+)?[eE]\+?[0-9]+", text):
        try:
            return str(int(float(text)))
        except (ValueError, OverflowError):
            return text
    # 去掉 Excel 硬加上的 .0 尾巴
    if re.fullmatch(r"[0-9]+\.0+", text):
        return text.split(".")[0]
    return text


def norm_warehouse(value):
    """倉別代碼：TAO1 / TXRC8 …，統一大寫並移除空白。"""
    text = norm_text(value).upper().replace(" ", "")
    return text


def norm_date(value):
    """任何形式的日期 -> 'YYYY-MM-DD'。無法解析時回傳 ''（呼叫端負責記錄）。"""
    if value is None or value == "":
        return ""
    if isinstance(value, _dt.datetime):
        return value.date().isoformat()
    if isinstance(value, _dt.date):
        return value.isoformat()
    # Excel 日期序號
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        serial = int(value)
        if 1 <= serial <= 80000:
            return (_EXCEL_EPOCH + _dt.timedelta(days=serial)).isoformat()
        return ""
    text = norm_text(value)
    if not text:
        return ""
    text = text.split(" ")[0] if re.match(r"^\d{4}[-/.]\d{1,2}[-/.]\d{1,2} ", text) else text
    for pattern in _DATE_PATTERNS:
        try:
            return _dt.datetime.strptime(text, pattern).date().isoformat()
        except ValueError:
            continue
    # 純數字字串也可能是 Excel 序號
    if re.fullmatch(r"\d{1,5}", text):
        serial = int(text)
        if 1 <= serial <= 80000:
            return (_EXCEL_EPOCH + _dt.timedelta(days=serial)).isoformat()
    return ""


def norm_int(value):
    """數量欄位。無法解析回傳 None，絕不默默當成 0。"""
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, int):
        return value
    if isinstance(value, float):
        return int(round(value))
    text = norm_text(value).replace(",", "")
    if text in ("", "*", "-", "—"):
        return None
    try:
        return int(round(float(text)))
    except ValueError:
        return None


def norm_decimal(value):
    """單價等數值欄位。無法解析回傳 None。"""
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = norm_text(value).replace(",", "").replace("$", "")
    try:
        return float(text)
    except ValueError:
        return None


def norm_header(value):
    """表頭比對用：去掉所有空白與換行，全形轉半形。"""
    return re.sub(r"\s+", "", norm_text(value))

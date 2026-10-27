"""快到貨還沒完成（Jerry 2026-10-27 看過示意圖「OK 照這樣做」）。

到貨日在 3 個工作天內（跳過週六日、內建的台灣國定假日；選 B 工作天），或已經過了到貨日，
PO 狀態還是已建立／已回覆／處理中／修改中的 PO → 清單到貨日紅字＋「剩 N 天／今天／已過 N 天」、
酷澎頁上方紅色數字卡、首頁今日事項。不跳視窗（Jerry 沒選，照建議）。
國定假日用瑪氏出貨同一份（mars.po.BUILTIN_HOLIDAYS）。
"""
import datetime as _dt

DUE_STATUSES = ("已建立", "已回覆", "處理中", "修改中")
WORK_DAYS = 3


def _holidays():
    try:
        from mars.po import BUILTIN_HOLIDAYS
        return BUILTIN_HOLIDAYS
    except Exception:                       # 瑪氏模組讀不到也不要讓酷澎頁壞掉：只跳週末
        return set()


def _is_work(d, hol):
    return d.weekday() < 5 and d.isoformat() not in hol


def cutoff(today):
    """today（date）往後第 3 個工作天；到貨日 ≤ 這天就算快到貨。"""
    hol, d, left = _holidays(), today, WORK_DAYS
    while left > 0:
        d += _dt.timedelta(days=1)
        if _is_work(d, hol):
            left -= 1
    return d


def label(delivery, today):
    """到貨日旁邊的小字：剩 N 天（工作天）／今天／已過 N 天（日曆天）。"""
    try:
        d = _dt.date.fromisoformat((delivery or "")[:10])
    except ValueError:
        return ""
    if d < today:
        return f"已過 {(today - d).days} 天"
    if d == today:
        return "今天"
    hol, n, x = _holidays(), 0, today
    while x < d:
        x += _dt.timedelta(days=1)
        if _is_work(x, hol):
            n += 1
    return f"剩 {max(n, 1)} 天"


def where(today):
    """挑出快到貨還沒完成的品項（order_rows 一列一個品項）：清單篩選、上面那張卡、首頁今日事項都用這句，數字才會一樣。"""
    sql = ("po_status IN (" + ",".join("?" * len(DUE_STATUSES)) + ") "
           "AND delivery_date != '' AND delivery_date <= ?")
    return sql, list(DUE_STATUSES) + [cutoff(today).isoformat()]


def is_due(status, delivery, today):
    return status in DUE_STATUSES and bool(delivery) and delivery[:10] <= cutoff(today).isoformat()

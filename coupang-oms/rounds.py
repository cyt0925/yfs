"""PO 修改紀錄：一輪處理完、打一個 OK，算一筆修改（小真 2026-10-20 的 Excel；Jerry 定案）。

小真原本的做法：PO 視窗的備註把每個動作用「=>」接成一串，處理完在最後打「10/20-OK」，
備註裡有幾個 OK 就是這張 PO 修改過幾次。一輪裡面改交期、改倉、拆單、下修好幾個動作、跨好幾天，都算一次。
改成按鈕記（Jerry：靠打字算次數不穩、而且只有次數）：
- 打 OK 的時候選「這輪在哪個階段結束」「這輪做了什麼（可複選）」「主因（可複選，Jerry 2026-10-21；存成逗號隔開）」，說明選填。
  統計時一輪算進每個勾到的主因，主因加總會大於修改次數（跟「做了什麼」一樣）。
- 「做了什麼」系統先勾：看上一個 OK 之後這張單的修改歷程有沒有下修、改交期／倉別、拆單。階段照 PO 狀態先帶。
- 按下去：po_rounds 多一筆、修改歷程記一筆。記錯可以刪（軟刪、記歷程）。
- 佐證只有一份（Jerry 2026-10-20「備註跟修改紀錄理當只能由一個當佐證」）：修改紀錄才算，備註回到純文字只寫過程，
  系統不再把「=>10/20-OK」接到備註、大家也不用再打 OK。第一版有接，拆掉了。
- 防呆（Jerry：「他們一定得填」）：備註裡自己打了 OK 存檔會擋下來（不是拿備註當證據，是提醒去按鈕記）；PO 狀態改成已完成／已取消時，
  上一個 OK 之後有改動還沒記 OK 也擋下來；首頁今日事項列「改過但還沒記 OK 的 PO」。
- 從 SINCE 那天開始算，以前備註裡打的 OK 不回頭算（小真：從今天開始）。
原因清單照小真的檔：數量下修依階段三組，改交期／倉別、拆單三個階段共用。
"""
import re

import db

STAGES = [("1", "第一階段", "收單確認"), ("2", "第二階段", "已排單確認"), ("3", "第三階段", "出貨確認")]
STAGE_LABEL = {k: f"{a} {b}" for k, a, b in STAGES}
STAGE_SHORT = {k: a for k, a, _b in STAGES}
KINDS = [("qty", "數量下修"), ("date", "改交期／倉別"), ("split", "拆單")]
KIND_LABEL = dict(KINDS)
REASONS_QTY = {
    "1": ["未達出貨數", "商品配貨", "暫不供貨", "庫存不足", "金額錯誤", "條碼錯誤", "停產", "酷澎取消"],
    "2": ["產能不及", "商品配貨", "運能問題", "效期不足", "系統問題", "酷澎異動"],
    "3": ["缺貨", "運能問題", "效期問題", "酷澎異動", "其他"],
}
REASONS_DATE = ["備貨時間不及", "約倉問題", "車趟考量", "倉容量問題", "改不了倉（酷澎規範限制）"]
REASONS_SPLIT = ["單位不同", "車趟考量", "產能不及"]
STAGE_BY_STATUS = {"已建立": "1", "已回覆": "1", "處理中": "2", "已完成": "3"}   # 修改中、已取消三個階段都有：帶上一筆的
END_STATUSES = ("已完成", "已取消")
SINCE = "2026-10-20"                     # 這天以前的改動不算（小真：從今天開始）
CHANGE_FIELDS = ("qty_coupang", "qty_ship", "delivery_date", "warehouse")
OPTIONS = {"stages": [{"key": k, "label": STAGE_LABEL[k]} for k, _a, _b in STAGES],
           "kinds": [{"key": k, "label": v} for k, v in KINDS],
           "reasons": {"qty": REASONS_QTY, "date": REASONS_DATE, "split": REASONS_SPLIT}}
MSG_TYPED_OK = "備註裡打了 OK。修改次數以下方「修改紀錄」為準，請改按「這一輪處理完了（OK）」記這一輪，備註不用再打 OK。"
MSG_STATUS = "這張單在上一個 OK 之後還有改動，請先按「這一輪處理完了（OK）」記這一輪，再改 PO 狀態。"
_OK = re.compile(r"(?<![A-Za-z])OK(?![A-Za-z])", re.I)


def count_ok(text):
    return len(_OK.findall(text or ""))


def reason_options(stage, kinds):
    """[(組名, [原因])]，照小真的檔分組；主因從勾到的類別裡挑一個。"""
    out = []
    if "qty" in kinds:
        out.append((f"數量下修 · {STAGE_SHORT.get(stage, '')}", REASONS_QTY.get(stage, [])))
    if "date" in kinds:
        out.append(("改交期／倉別", REASONS_DATE))
    if "split" in kinds:
        out.append(("拆單", REASONS_SPLIT))
    return out


def _num(v):
    try:
        return float(str(v).replace(",", "").strip())
    except (TypeError, ValueError):
        return None


def _row(r):
    d = dict(r)
    d["reasons"] = [x for x in (d.get("reason") or "").split(",") if x]
    d["reason"] = "、".join(d["reasons"])
    d["kinds"] = [k for k in (d.get("kinds") or "").split(",") if k]
    d["kind_labels"] = [KIND_LABEL.get(k, k) for k in d["kinds"]]
    d["stage_label"] = STAGE_LABEL.get(d["stage"], d["stage"])
    return d


def rounds_for(conn, po):
    return [_row(r) for r in conn.execute(
        "SELECT id, po_number, ok_at, stage, kinds, reason, note, operator FROM po_rounds "
        "WHERE po_number = ? AND deleted_at = '' ORDER BY ok_at, id", (po,)).fetchall()]


def last_ok(conn, po):
    r = conn.execute("SELECT MAX(ok_at) AS t FROM po_rounds WHERE po_number = ? AND deleted_at = ''", (po,)).fetchone()
    return (r["t"] if r else "") or ""


def changes_since(conn, po, since):
    """上一個 OK 之後這張單改了什麼：{kinds: set, n: 筆數}。下修＝下單數量或出貨數量變少（品項被拿掉也是出貨歸 0）。"""
    marks = ",".join("?" * len(CHANGE_FIELDS))
    rows = conn.execute(
        f"SELECT field, old_value, new_value FROM edit_logs WHERE po_number = ? AND changed_at > ? AND source IN ('import', 'manual') "
        f"AND field IN ({marks})", [po, since] + list(CHANGE_FIELDS)).fetchall()
    kinds, n = set(), 0
    for r in rows:
        if r["field"] in ("delivery_date", "warehouse"):
            kinds.add("date"); n += 1
        else:
            a, b = _num(r["old_value"]), _num(r["new_value"])
            if a is not None and b is not None and b < a:
                kinds.add("qty"); n += 1
    kid = conn.execute("SELECT COUNT(*) AS n FROM orders WHERE parent_po = ? AND first_seen_at > ?", (po, since)).fetchone()["n"]
    if kid:
        kinds.add("split"); n += 1
    return {"kinds": kinds, "n": n}


def _since(conn, po):
    return max(last_ok(conn, po), f"{SINCE} 00:00:00")


def pending_for(conn, po):
    """上一個 OK 之後還有改動（防呆、今日事項用）。"""
    return changes_since(conn, po, _since(conn, po))["n"] > 0


def prefill(conn, po, status, rows=None):
    """PO 視窗打開時先帶好：階段照 PO 狀態（修改中、已取消帶上一筆的），做了什麼照修改歷程。"""
    rows = rounds_for(conn, po) if rows is None else rows
    since = _since(conn, po)
    ch = changes_since(conn, po, since)
    stage = STAGE_BY_STATUS.get(status) or (rows[-1]["stage"] if rows else "1")
    return {"stage": stage, "kinds": [k for k, _v in KINDS if k in ch["kinds"]], "changes": ch["n"],
            "since": since if since > f"{SINCE} 00:00:00" else ""}


def pending_pos(conn):
    """改過但還沒記 OK 的 PO（從 SINCE 起算）。"""
    last = {r["po"]: r["t"] for r in conn.execute(
        "SELECT po_number AS po, MAX(ok_at) AS t FROM po_rounds WHERE deleted_at = '' GROUP BY po_number").fetchall()}
    marks = ",".join("?" * len(CHANGE_FIELDS))
    cand = {r["po"] for r in conn.execute(
        f"SELECT DISTINCT po_number AS po FROM edit_logs WHERE changed_at >= ? AND source IN ('import', 'manual') AND field IN ({marks})",
        [f"{SINCE} 00:00:00"] + list(CHANGE_FIELDS)).fetchall()}
    cand |= {r["po"] for r in conn.execute("SELECT DISTINCT parent_po AS po FROM orders WHERE parent_po != '' AND first_seen_at >= ?",
                                          (f"{SINCE} 00:00:00",)).fetchall()}
    out = []
    for po in sorted(cand):
        since = max(last.get(po) or "", f"{SINCE} 00:00:00")
        if changes_since(conn, po, since)["n"]:
            out.append(po)
    return out


def _log(conn, po, field, label, old, new, operator, note=""):
    conn.execute(
        """INSERT INTO edit_logs (order_id, po_number, sku_id, field, field_label, old_value, new_value, operator, source, note, changed_at)
           VALUES (0,?,'',?,?,?,?,?,'manual',?,?)""", (po, field, label, old or "", new or "", operator, note, db.now()))


def add_round(conn, po, stage, kinds, reasons, note, operator):
    """記一輪：驗證 → 寫 po_rounds → 歷程。備註不動（修改紀錄才是佐證）。reasons 可以好幾個。回 (錯誤訊息, 新的 rounds)。"""
    if stage not in STAGE_LABEL:
        return "請選擇這一輪在哪個階段結束。", None
    kinds = [k for k, _v in KINDS if k in (kinds or [])]
    if not kinds:
        return "請勾選這一輪做了什麼（至少一項）。", None
    allowed = []
    for _g, lst in reason_options(stage, kinds):
        allowed += [r for r in lst if r not in allowed]
    picked = {str(r).strip() for r in (reasons or []) if str(r).strip()}
    if not picked or not picked <= set(allowed):
        return "請勾選主因（只能選勾到的類別裡的原因，至少一個）。", None
    reason = ",".join(r for r in allowed if r in picked)
    if conn.execute("SELECT 1 FROM po_headers WHERE po_number = ?", (po,)).fetchone() is None:
        return "找不到這張 PO。", None
    stamp = db.now()
    conn.execute("INSERT INTO po_rounds (po_number, ok_at, stage, kinds, reason, note, operator, deleted_at, deleted_by) VALUES (?,?,?,?,?,?,?,'','')",
                 (po, stamp, stage, ",".join(kinds), reason, note or "", operator))
    summary = f"{STAGE_SHORT[stage]} OK（{'、'.join(KIND_LABEL[k] for k in kinds)}；主因 {reason}）"
    _log(conn, po, "po_round", "修改紀錄", "", summary, operator, note or "")
    return "", rounds_for(conn, po)


def delete_round(conn, po, rid, operator):
    r = conn.execute("SELECT * FROM po_rounds WHERE id = ? AND po_number = ? AND deleted_at = ''", (rid, po)).fetchone()
    if r is None:
        return "找不到這筆修改紀錄。", None
    conn.execute("UPDATE po_rounds SET deleted_at = ?, deleted_by = ? WHERE id = ?", (db.now(), operator, rid))
    r = _row(r)
    _log(conn, po, "po_round", "修改紀錄", f"{STAGE_SHORT.get(r['stage'], '')} OK（{'、'.join(r['kind_labels'])}；主因 {r['reason']}）", "", operator, "刪除修改紀錄")
    return "", rounds_for(conn, po)

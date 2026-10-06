"""YFS 訂單系統入口（網址 /portal）。

Alice 2026-10-06：「各通路可以分開……最外層我們做一個訂單系統入口，從這邊可以點選到各通路後台，
也能將各通路訂單重點，採用圖表或是表格的方式呈現」。Jerry 定名「YFS 訂單系統」、配色酒紅（深紫 #171028 漸層到 #96002a）。

這一頁只讀、不寫：數字全部從既有的表現算，不新增任何資料表。
- 酷澎：訂單管理的 orders／po_headers（整合表整份匯進來的那套），加上瑪氏拆單、竹運匯出紀錄、改單統計。
- 蝦皮特選、PChome：還沒有資料，只顯示「建置中／規劃中」，**不放假數字**。
- 左側選單的酷澎工具是捷徑，跟酷澎系統裡的「線別工具」是同樣的工具（Jerry：「兩邊都有，捷徑的概念」）。
- 營運 SOP 檢索是另一個獨立網站（GitHub Pages），這裡只是開在框裡，不動它。

各數字怎麼算（同事問起就照這裡講）：
- 今日新進：po_headers.filed_date＝今天（系統第一次看到這張單的日子）。
- 待出貨：交貨日今天以後、品項沒被酷澎拿掉、PO 狀態不是「已完成／已取消」的 PO 張數。
- 本月出貨箱數：交貨日在這個月的品項，出貨數量（沒有就用下單數量）÷ 箱入數，不湊整；箱入數空白的不算。
- 今日事項（原名「今天要處理的事」）：只列數量大於 0 的。
  1. 酷澎改了單要確認：orders.needs_review＝1 的 PO。
  2. 瑪氏拆單還沒填 EIP 單號：交貨日今天以後、eip_po 空白的拆單。
  3. 寶僑竹運拋檔還沒產：今天以後最近一個有寶僑訂單的交貨日，每個倉查竹運匯出紀錄有沒有那天那倉。
  4. 過了交期還沒驗收：交貨日在過去 14 天、驗收狀態還是「未驗收」、沒取消的 PO。
- 改單：直接用改單統計（master/stats.py）這個月的合計，估計工時用改單統計頁填的「每次幾分鐘」（存在瀏覽器）。
"""
import datetime as _dt

from flask import Blueprint, current_app, jsonify, render_template, session

import db

portal_bp = Blueprint("portal", __name__)

SOP_URL = "https://cyt0925.github.io/yfs/"
WARROOM_URL = "https://yfs-warroom.vercel.app/"
PG_LINE = "寶僑"
DONE_STATUSES = ("已完成", "已取消")

CHANNELS = [
    # title／logo：首頁「各通路系統」卡片的標題與圖示（Jerry 2026-10-06 給的三個圖檔，縮成 160px 高放 static/）
    {"key": "coupang", "name": "酷澎", "title": "酷澎訂單管理系統", "logo": "logo_coupang.png", "live": True, "status": ""},
    {"key": "shopee", "name": "蝦皮特選", "title": "蝦皮訂單管理系統", "logo": "logo_shopee.png", "live": False, "status": "建置中"},
    {"key": "pchome", "name": "PChome", "title": "PChome 訂單管理系統", "logo": "logo_pchome.png", "live": False, "status": "規劃中"},
]


def _iso(d):
    return d.isoformat()


def _boxes_expr():
    # 箱數不湊整；箱入數空白或 0 的不算。乘 1.0 讓 SQLite／PostgreSQL 都走小數除法。
    return "CASE WHEN o.box_size > 0 THEN COALESCE(o.qty_ship, o.qty_coupang, 0) * 1.0 / o.box_size ELSE 0 END"


def _one(conn, sql, params=()):
    """單一個值。正式站的 PostgreSQL 回的是「欄名 → 值」的 dict，不能用 row[0]，所以一律取第一欄。"""
    row = conn.execute(sql, params).fetchone()
    if row is None:
        return None
    if isinstance(row, dict):
        return next(iter(row.values()), None)
    return row[0]


def coupang_summary(conn, today):
    """酷澎卡片與兩張圖的數字。today 是 date。"""
    t = _iso(today)
    month_start = _iso(today.replace(day=1))
    nxt = (today.replace(day=28) + _dt.timedelta(days=4)).replace(day=1)
    month_end = _iso(nxt - _dt.timedelta(days=1))
    done = list(DONE_STATUSES)

    new_today = _one(conn, "SELECT COUNT(*) FROM po_headers WHERE filed_date = ?", (t,)) or 0
    pending = _one(conn, f"""
        SELECT COUNT(DISTINCT o.po_number) FROM orders o LEFT JOIN po_headers h ON h.po_number = o.po_number
        WHERE o.delivery_date >= ? AND o.removed_from_coupang = 0
          AND COALESCE(h.po_status, '') NOT IN (?, ?)""", (t, *done)) or 0
    month_boxes = _one(conn, f"""
        SELECT SUM({_boxes_expr()}) FROM orders o LEFT JOIN po_headers h ON h.po_number = o.po_number
        WHERE o.delivery_date >= ? AND o.delivery_date <= ? AND o.removed_from_coupang = 0
          AND COALESCE(h.po_status, '') <> ?""", (month_start, month_end, "已取消")) or 0
    last_import = _one(conn, "SELECT MAX(committed_at) FROM import_batches WHERE committed = 1") or ""

    # 近 30 天每天新進 PO
    d0 = today - _dt.timedelta(days=29)
    rows = conn.execute("SELECT filed_date AS d, COUNT(*) AS n FROM po_headers WHERE filed_date >= ? AND filed_date <= ? GROUP BY filed_date",
                        (_iso(d0), t)).fetchall()
    by_day = {r["d"]: r["n"] for r in rows}
    daily = [{"date": _iso(d0 + _dt.timedelta(days=i)), "value": int(by_day.get(_iso(d0 + _dt.timedelta(days=i)), 0))} for i in range(30)]

    # 今天起 7 天，每天要出的箱數
    d7 = today + _dt.timedelta(days=6)
    rows = conn.execute(f"""
        SELECT o.delivery_date AS d, SUM({_boxes_expr()}) AS n FROM orders o LEFT JOIN po_headers h ON h.po_number = o.po_number
        WHERE o.delivery_date >= ? AND o.delivery_date <= ? AND o.removed_from_coupang = 0
          AND COALESCE(h.po_status, '') <> ?
        GROUP BY o.delivery_date""", (t, _iso(d7), "已取消")).fetchall()
    by_day = {r["d"]: float(r["n"] or 0) for r in rows}
    ship = [{"date": _iso(today + _dt.timedelta(days=i)), "value": round(by_day.get(_iso(today + _dt.timedelta(days=i)), 0), 1)} for i in range(7)]

    return {"new_today": int(new_today), "pending_pos": int(pending), "month_boxes": round(float(month_boxes)),
            "last_import": last_import, "daily_new": daily, "ship_7d": ship}


def tasks(conn, today):
    """今日事項，只回數量大於 0 的。"""
    t = _iso(today)
    out = []

    rows = conn.execute("SELECT DISTINCT po_number AS po FROM orders WHERE needs_review = 1 ORDER BY po_number").fetchall()
    if rows:
        pos = [r["po"] for r in rows]
        out.append({"key": "review", "level": "crit", "level_text": "要確認", "icon": "bi-pencil-square",
                    "title": f"酷澎改了 {len(pos)} 張單", "detail": "PO " + "、".join(pos[:2]) + ("…" if len(pos) > 2 else ""),
                    "url": "/", "count": len(pos)})

    # 商品主檔、瑪氏的表 init_db 一定會建；不包 try：PostgreSQL 一句失敗整個交易就不能再查，包了反而後面全錯
    n = _one(conn, "SELECT COUNT(*) FROM mst_mars_splits WHERE COALESCE(eip_po, '') = '' AND delivery_date >= ?", (t,)) or 0
    if n:
        out.append({"key": "mars_eip", "level": "warn", "level_text": "待辦", "icon": "bi-receipt",
                    "title": f"瑪氏 {n} 份拆單還沒填 EIP 單號", "detail": "填了才能產瑪氏採購單", "url": "/mars", "count": int(n)})

    nxt = _one(conn, "SELECT MIN(delivery_date) FROM orders WHERE line = ? AND delivery_date > ? AND removed_from_coupang = 0", (PG_LINE, t))
    if nxt and nxt <= _iso(today + _dt.timedelta(days=7)):
        whs = [r["wh"] for r in conn.execute(
            "SELECT DISTINCT warehouse AS wh FROM orders WHERE line = ? AND delivery_date = ? AND removed_from_coupang = 0 ORDER BY warehouse",
            (PG_LINE, nxt)).fetchall()]
        missing = []
        for wh in whs:
            done = _one(conn, "SELECT COUNT(*) FROM mst_logs WHERE field = ? AND new_value LIKE ?", ("zhuyun_export", f"{nxt} {wh}：%")) or 0
            if not done:
                missing.append(wh or "倉別空白")
        if missing:
            md = f"{int(nxt[5:7])}/{int(nxt[8:10])}"
            out.append({"key": "zhuyun", "level": "warn", "level_text": "要出貨", "icon": "bi-truck",
                        "title": f"寶僑 {md} 到貨還有 {len(missing)} 個倉沒產竹運拋檔", "detail": "、".join(missing),
                        "url": "/zhuyun", "count": len(missing)})

    d14 = _iso(today - _dt.timedelta(days=14))
    n = _one(conn, """
        SELECT COUNT(DISTINCT o.po_number) FROM orders o JOIN po_headers h ON h.po_number = o.po_number
        WHERE o.delivery_date >= ? AND o.delivery_date < ? AND o.removed_from_coupang = 0
          AND h.receiving_status = ? AND h.po_status <> ?""", (d14, t, "未驗收", "已取消")) or 0
    if n:
        out.append({"key": "receiving", "level": "info", "level_text": "待辦", "icon": "bi-clipboard-check",
                    "title": f"過了交期還沒驗收 {n} 張 PO", "detail": "近 14 天到貨的", "url": "/", "count": int(n)})
    return out


def change_stats(conn, today):
    """這個月的改單：直接用改單統計的合計（跟改單統計頁同一套算法）。"""
    from master.stats import _build_stats
    month = _iso(today)[:7]
    d = _build_stats(conn, month, month, "", False)
    tot = d["total"]
    reasons = [{"name": k, "count": v} for k, v in tot.get("reasons", {}).items() if v]
    return {"month": month, "events": tot["events"], "coupang": tot["coupang"], "manual": tot["manual"],
            "pos": tot["pos"], "changed_pos": tot["changed_pos"], "changed_pct": tot["changed_pct"], "reasons": reasons}


def recent_actions(conn, limit=6):
    """最近的動作：人手改的訂單（edit_logs 的 manual）＋商品主檔、瑪氏、竹運、採購表的紀錄（mst_logs）。"""
    rows = []
    cols = "changed_at, operator, field_label, old_value, new_value, po_number"
    pick = lambda r: {"at": r["changed_at"], "who": r["operator"], "what": r["field_label"],  # noqa: E731
                      "old": r["old_value"] or "", "new": r["new_value"] or "", "po": r["po_number"] or ""}
    for r in conn.execute(f"SELECT {cols} FROM edit_logs WHERE source = ? ORDER BY id DESC LIMIT ?", ("manual", limit)).fetchall():
        rows.append(pick(r))
    for r in conn.execute(f"SELECT {cols} FROM mst_logs ORDER BY id DESC LIMIT ?", (limit,)).fetchall():
        rows.append(pick(r))
    rows.sort(key=lambda x: x["at"] or "", reverse=True)
    return rows[:limit]


def summary(conn, today):
    channels = []
    for c in CHANNELS:
        item = dict(c)
        if c["key"] == "coupang":
            item.update(coupang_summary(conn, today))
        channels.append(item)
    return {"today": _iso(today), "now": db.now(), "channels": channels, "tasks": tasks(conn, today),
            "changes": change_stats(conn, today), "recent": recent_actions(conn)}


# ---------------------------------------------------------------- 路由

def _page(view):
    return render_template("portal.html", view=view, logged_in_user=session.get("user", ""),
                           build_version=current_app.config.get("BUILD_VERSION", ""),
                           sop_url=SOP_URL, warroom_url=WARROOM_URL, channels=CHANNELS)


@portal_bp.route("/portal")
def portal_page():
    return _page("home")


@portal_bp.route("/portal/sop")
def portal_sop():
    return _page("sop")


@portal_bp.route("/api/portal/summary")
def api_portal_summary():
    today = _dt.date.fromisoformat(db.today())
    conn = db.get_conn()
    try:
        data = summary(conn, today)
    finally:
        conn.close()
    return jsonify(data)

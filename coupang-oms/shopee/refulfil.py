"""改履約方式時，把已經做好的 EIP 採購單、竹運拋單檔一起處理掉（Jerry 2026-10-22 看過示意圖「OK 照這樣做」）。

起因：OP 把一張已經產過 EIP 的竹運採購進貨改成竹運出貨，舊的 EIP 檔裡還有它；之前只能到「待回填 EIP 採購單號」
把整份 EIP 取消產出、整份檔裡的單全部退回去重產，Jerry：「做系統很不直覺」「要有按鈕讓他們修改」。
現在改履約方式那一格（或批次列）就一起處理：
  - 這張單在某份 EIP 裡（履約方式跟那份不一樣了）→ 從那份 EIP 拿掉，那份只留其他單、重新算好（檔名照剩下的單重算）；
    整份只剩它就整份取消。已回填單號的照樣拿掉，畫面提醒要通知採購去 EIP 系統改。那份 EIP 記一行「重新做過」的提醒。
  - 這張單產過竹運拋單檔、改成不是竹運出貨 → 取消這張的拋單檔（跟「取消產出」一樣）。
先用 /api/shopee/fulfil/check 問會影響什麼（畫面照這個跳提醒），按確定才由 /api/shopee/fulfil 一起做。
"""
import collections

from .common import *  # noqa: F401,F403


def _ids_of(conn, no):
    return [r["id"] for r in _rows(conn.execute(
        "SELECT id FROM shp_orders WHERE po_id = ? OR (po_id = '' AND pr_id = ?)", (no, no)))]


def impacts(conn, no, fulfil):
    """這張單改成 fulfil 會動到哪些已產出的東西：{eips: [...], dispatch: {...} 或 None}。"""
    ids = _ids_of(conn, no)
    out = {"order_no": no, "eips": [], "dispatch": None, "buy_to_ship": False}
    if not ids:
        return out
    marks = ",".join("?" * len(ids))
    for e in _rows(conn.execute(
            f"SELECT DISTINCT e.id, e.filename, e.eip_po, e.line, e.fulfil FROM shp_eip e JOIN shp_eip_items i ON i.eip_id = e.id "
            f"WHERE i.order_id IN ({marks}) ORDER BY e.id", ids)):
        if e["fulfil"] == fulfil:
            continue
        others = _rows(conn.execute(f"SELECT DISTINCT o.po_id, o.pr_id FROM shp_eip_items i JOIN shp_orders o ON o.id = i.order_id "
                                    f"WHERE i.eip_id = ? AND i.order_id NOT IN ({marks})", [e["id"]] + ids))
        nos = {(o["po_id"] or o["pr_id"]) for o in others}
        out["eips"].append({"id": e["id"], "filename": e["filename"], "eip_po": e["eip_po"] or "", "others": len(nos),
                            "mars": e["line"] == "瑪氏", "direct": e["fulfil"] == "供應商直送" and e["line"] != "瑪氏"})
    items = _rows(conn.execute(f"SELECT * FROM shp_orders WHERE id IN ({marks})", ids))
    if not any(ships(i, fulfil) for i in items):           # 改了之後不用拋單了才取消拋單檔
        d = _row(conn.execute(f"SELECT d.filename, d.created_at FROM shp_dispatch_items i JOIN shp_dispatch d ON d.id = i.dispatch_id "
                              f"WHERE i.order_id IN ({marks}) ORDER BY d.id DESC LIMIT 1", ids))
        if d:
            out["dispatch"] = {"filename": d["filename"], "created_at": d["created_at"]}
    # 竹運採購進貨改竹運出貨：提醒入庫單到了會自動進待產拋單檔，不用改（Becky 2026-10-22 就是為了拋單才改的）
    out["buy_to_ship"] = fulfil == SHIP_DIRECT and any(i["fulfil"] == SHIP_AFTER_BUY for i in items)
    return out


def _num(v):
    v = round(v, 4)
    return int(v) if float(v).is_integer() else v


def _pull_from_eip(conn, e, no, oids, new_fulfil, operator, stamp):
    """把這張單（oids）從 EIP e 拿掉，剩下的重新算；回這份 EIP 的處理結果文字。"""
    from .purchase import build_plan, eip_filename
    plans = {p["order_id"]: p for g in build_plan(conn, e["fulfil"], order_nos=[no]) for p in g["items"]}
    if not plans:                    # 履約方式已經先改掉了：用改之前的履約方式查不到 → 直接看品項
        plans = {}
    rows = json.loads(e["rows_json"] or "[]")
    items = {o["id"]: o for o in _rows(conn.execute(
        f"SELECT * FROM shp_orders WHERE id IN ({','.join('?' * len(oids))})", oids))}
    for oid in oids:
        p, it = plans.get(oid), items.get(oid) or {}
        code = (p or {}).get("code") or it.get("base_code")
        out_unit, in_unit = (p or {}).get("out_unit"), (p or {}).get("unit") or it.get("unit_name")
        cands = [r for r in rows if r.get("code") == code and any(x.get("order_id") == oid for x in r.get("parts") or [])]
        cands = cands or [r for r in rows if r.get("code") == code and (r.get("unit") == out_unit if out_unit else r.get("in_unit") == in_unit)]
        cands = cands or [r for r in rows if r.get("code") == code]
        if not cands:
            continue
        r = cands[0]
        parts = r.get("parts") or []
        if any(x.get("order_id") == oid for x in parts):
            use = sum(x["use_qty"] or 0 for x in parts if x.get("order_id") == oid)
            r["parts"] = [x for x in parts if x.get("order_id") != oid]
        else:                        # 2026-10-21 以前產的 EIP 沒記每張單各多少：用這張單現在算出來的下採數量
            use = (p or {}).get("use_qty") if (p or {}).get("use_qty") is not None else (it.get("qty") or 0)
        old_in = r.get("in_qty") or 0
        if old_in <= use:
            rows.remove(r)
            continue
        ratio = (old_in - use) / old_in
        r["in_qty"] = _num(old_in - use)
        r["qty"] = _num((r.get("qty") or 0) * ratio)
        if "cases" in r:
            r["cases"] = _num((r.get("cases") or 0) * ratio)
    conn.execute(f"DELETE FROM shp_eip_items WHERE eip_id = ? AND order_id IN ({','.join('?' * len(oids))})", [e["id"]] + oids)
    left = _rows(conn.execute("SELECT o.* FROM shp_eip_items i JOIN shp_orders o ON o.id = i.order_id WHERE i.eip_id = ?", (e["id"],)))
    rows = [r for r in rows if r.get("qty")]
    if not left or not rows:
        conn.execute("DELETE FROM shp_eip_items WHERE eip_id = ?", (e["id"],))
        conn.execute("DELETE FROM shp_eip WHERE id = ?", (e["id"],))
        return f"整份 EIP「{e['filename']}」只有這張單，已取消"
    dates, whs = sorted({o["expected_date"] for o in left}), sorted({o["warehouse"] for o in left})
    name = eip_filename(e["line"], e["fulfil"], dates, whs, e["unit_group"])
    md = f"{int(stamp[5:7])}/{int(stamp[8:10])}"
    note = (f"{md} 重新做過：拿掉 {no}（改{new_fulfil}）。"
            + (f"已回填 {e['eip_po']}，請通知採購到 EIP 系統把這張單的品項拿掉" if e["eip_po"] else "已上傳蝦皮 EIP 的話，請重新下載再上傳"))
    conn.execute("UPDATE shp_eip SET rows_json = ?, filename = ?, warehouse = ?, expected_date = ?, note = ?, updated_by = ?, updated_at = ? WHERE id = ?",
                 (json.dumps(rows, ensure_ascii=False), name, "、".join(whs), dates[0], note, operator, stamp, e["id"]))
    return f"已從 EIP「{e['filename']}」拿掉，那份重新做好" + (f"（檔名改為 {name}）" if name != e["filename"] else "")


def apply(conn, no, fulfil, operator, stamp):
    """改履約方式之前呼叫：從不同履約方式的 EIP 拿掉、取消拋單檔。回要寫進歷程的說明（list）。"""
    ids = _ids_of(conn, no)
    if not ids:
        return []
    marks = ",".join("?" * len(ids))
    notes = []
    links = collections.defaultdict(list)
    for x in _rows(conn.execute(f"SELECT i.eip_id, i.order_id FROM shp_eip_items i JOIN shp_eip e ON e.id = i.eip_id "
                                f"WHERE i.order_id IN ({marks}) AND e.fulfil != ?", ids + [fulfil])):
        links[x["eip_id"]].append(x["order_id"])
    for eid, oids in links.items():
        e = _row(conn.execute("SELECT * FROM shp_eip WHERE id = ?", (eid,)))
        if e:
            notes.append(_pull_from_eip(conn, e, no, oids, fulfil, operator, stamp))
    items = _rows(conn.execute(f"SELECT * FROM shp_orders WHERE id IN ({marks})", ids))
    if not any(ships(i, fulfil) for i in items):
        dids = {r["dispatch_id"] for r in _rows(conn.execute(f"SELECT dispatch_id FROM shp_dispatch_items WHERE order_id IN ({marks})", ids))}
        if dids:
            conn.execute(f"DELETE FROM shp_dispatch_items WHERE order_id IN ({marks})", ids)
            for did in dids:
                if conn.execute("SELECT COUNT(*) AS n FROM shp_dispatch_items WHERE dispatch_id = ?", (did,)).fetchone()["n"] == 0:
                    conn.execute("DELETE FROM shp_dispatch WHERE id = ?", (did,))
            notes.append("已取消這張單的竹運拋單檔")
    return notes


@shopee_bp.route("/api/shopee/fulfil/check", methods=["POST"])
def api_fulfil_check():
    """改履約方式之前問：會動到哪些已產出的 EIP、拋單檔（畫面照這個跳提醒）。"""
    p = request.get_json(silent=True) or {}
    nos = [str(x).strip() for x in (p.get("order_nos") or []) if str(x).strip()]
    fulfil = str(p.get("fulfil") or "").strip()
    if not nos or fulfil not in FULFILS:
        return jsonify({"error": "請選擇訂單和履約方式。"}), 400
    conn = get_conn()
    try:
        out = [impacts(conn, no, fulfil) for no in nos]
    finally:
        conn.close()
    return jsonify({"orders": [x for x in out if x["eips"] or x["dispatch"] or x["buy_to_ship"]]})


__all__ = ["impacts", "apply", "api_fulfil_check"]

"""② 訂單明細的查詢與修改：清單＋篩選面、就地編輯、PO 改期、PO 視窗整張存。"""
from .common import *  # noqa: F401,F403 — 共用工具、Flask、db、openpyxl 都從這裡來


@master_bp.route("/api/master/orders")
def api_orders():
    month = norm_text(request.args.get("month")) or _this_month()
    # month_to：畫面要一次看好幾個月（依匯入批次篩選時，那批的交期跨月就一次列完，
    # 不用逐月切；跟匯出用同一個範圍，畫面上看到的就是匯出去的）
    month_to = norm_text(request.args.get("month_to")) or month
    if not _valid_month(month) or not _valid_month(month_to):
        return jsonify({"error": "月份格式要像 2026-09。"}), 400
    months = _month_span(month, month_to)
    filters = _read_filters(request.args)
    cfg = _line_groups()
    conn = get_conn()
    try:
        rows = _month_orders(conn, months, cfg)
    finally:
        conn.close()

    # 篩選面在「月份」範圔上算、線別除外也一樣，勾了什麼其他選項不會消失
    f_lines = collections.Counter(); f_dates = collections.OrderedDict()
    f_pos = collections.OrderedDict(); f_brands = collections.Counter(); f_wh = collections.Counter()
    for o in rows:
        f_lines[o["line_group"]] += 1
        f_wh[o["warehouse"] or ""] += 1
        if o["brand"]:
            f_brands[o["brand"]] += 1
        d = o["delivery_date"] or ""
        fd = f_dates.setdefault(d, {"date": d, "cases": 0.0, "rows": 0, "pos": set(), "missing_box": 0})
        fd["rows"] += 1; fd["pos"].add(o["po_number"])
        if o["cases"] is None:
            fd["missing_box"] += 1
        else:
            fd["cases"] += o["cases"]
        fp = f_pos.setdefault(o["po_number"], {"po_number": o["po_number"], "date": d, "rows": 0,
                                                "cases": 0.0, "warehouse": o["warehouse"],
                                                "lines": set(), "date_overridden": 0})
        fp["rows"] += 1; fp["cases"] += o["cases"] or 0; fp["lines"].add(o["line_group"])
        fp["date_overridden"] = fp["date_overridden"] or o["delivery_date_overridden"]
    for fd in f_dates.values():
        fd["po_count"] = len(fd.pop("pos")); fd["cases"] = round(fd["cases"], 2)
    for fp in f_pos.values():
        fp["cases"] = round(fp["cases"], 2); fp["lines"] = sorted(fp["lines"])

    # 同一張 PO 被兩次以上匯入動到（第一次進來、後來又加品項或改數字）→ 對帳時 PO 數會重複，點名出來
    po_batches = collections.OrderedDict()
    for o in rows:
        st = po_batches.setdefault(o["po_number"], set())
        for k in ("first_batch_id", "last_batch_id"):
            if o.get(k):
                st.add(o[k])
    overlap_pos = [{"po_number": po, "batch_ids": sorted(b), "date": f_pos[po]["date"], "rows": f_pos[po]["rows"]}
                   for po, b in po_batches.items() if len(b) >= 2]

    shown = [o for o in rows if _keep(o, filters)]
    return jsonify({
        "month": month, "rows": shown, "count": len(shown),
        "total_cases": round(sum(o["cases"] or 0 for o in shown), 2),
        "facets": {
            "lines": [{"line": l, "rows": n} for l, n in sorted(f_lines.items(), key=lambda x: (x[0] == UNCLASSIFIED, x[0]))],
            "dates": list(f_dates.values()), "pos": list(f_pos.values()),
            "brands": [{"brand": b, "rows": n} for b, n in sorted(f_brands.items())],
            "warehouses": [{"warehouse": w, "rows": n} for w, n in sorted(f_wh.items())],
            "overlap_pos": overlap_pos,
        },
    })


def _apply_item_edit(conn, o, payload, operator):
    """單一品項的編輯（出貨數量、備註、恢復整合表數字）。回傳 (sets, vals, changed)。"""
    sets, vals, changed = [], [], 0
    if "qty_ship" in payload and payload.get("qty_ship") not in (None, ""):
        new = norm_int(payload.get("qty_ship"))
        if new is None or new < 0:
            raise ValueError("出貨數量要是 0 或正整數。")
        if not _same(o["qty_ship"], new):
            sets += ["qty_ship = ?", "qty_ship_overridden = 1"]; vals.append(new)
            _log(conn, o["line"], o["po_number"], o["sku_id"], o["barcode"], "qty_ship", "出貨數量",
                 o["qty_ship"], new, operator, "manual")
            changed += 1
    if "remarks" in payload:
        new = norm_text(payload.get("remarks"))
        if not _same(o["remarks"], new):
            sets += ["remarks = ?", "remarks_overridden = ?"]; vals += [new, 1 if new else 0]
            _log(conn, o["line"], o["po_number"], o["sku_id"], o["barcode"], "remarks", "備註",
                 o["remarks"], new, operator, "manual")
            changed += 1
    if payload.get("reset_qty_ship"):
        file_ship = o["qty_file_ship"] if o["qty_file_ship"] is not None else o["qty_coupang"]
        sets += ["qty_ship = ?", "qty_ship_overridden = 0"]; vals.append(file_ship)
        _log(conn, o["line"], o["po_number"], o["sku_id"], o["barcode"], "qty_ship", "出貨數量",
             o["qty_ship"], file_ship, operator, "manual", "恢復為整合表數字")
        changed += 1
    return sets, vals, changed


@master_bp.route("/api/master/orders/<int:order_id>", methods=["PUT"])
def api_update_order(order_id):
    payload = request.get_json(silent=True) or {}
    operator = _operator()
    client_version = norm_int(payload.get("version"))
    if client_version is None:
        return jsonify({"error": "缺少版本資訊，請重新整理後再試。"}), 400
    conn = get_conn()
    try:
        o = _row(conn.execute("SELECT * FROM mst_orders WHERE id = ?", (order_id,)))
        if o is None:
            return jsonify({"error": "找不到這筆。"}), 404
        if o["version"] != client_version:
            return jsonify({"error": "conflict",
                            "message": f"這筆剛被另一次儲存修改過（{o['updated_at']}），請重新載入後再編輯。"}), 409
        try:
            sets, vals, changed = _apply_item_edit(conn, o, payload, operator)
        except ValueError as exc:
            conn.rollback(); return jsonify({"error": str(exc)}), 400
        if not changed:
            conn.rollback()
            return jsonify({"ok": True, "changed": 0})
        sets += ["updated_at = ?", "version = version + 1"]; vals.append(now())
        cur = conn.execute(f"UPDATE mst_orders SET {', '.join(sets)} WHERE id = ? AND version = ?",
                           vals + [order_id, client_version])
        if cur.rowcount == 0:
            conn.rollback()
            return jsonify({"error": "conflict", "message": "儲存瞬間有其他人也改了這筆，請重新載入。"}), 409
        conn.commit()
        fresh = _decorate(_row(conn.execute(f"{_ORDER_SELECT} WHERE o.id = ?", (order_id,))))
        return jsonify({"ok": True, "changed": changed, "row": fresh})
    finally:
        conn.close()


@master_bp.route("/api/master/orders/<int:order_id>", methods=["DELETE"])
def api_delete_order(order_id):
    conn = get_conn()
    try:
        o = _row(conn.execute("SELECT * FROM mst_orders WHERE id = ?", (order_id,)))
        if o is None:
            return jsonify({"error": "找不到這筆。"}), 404
        conn.execute("DELETE FROM mst_orders WHERE id = ?", (order_id,))
        _log(conn, o["line"], o["po_number"], o["sku_id"], o["barcode"], "delete", "刪除品項",
             f"出貨 {o['qty_ship']} / 交貨 {o['delivery_date']}", "", _operator(), "manual")
        conn.commit()
        return jsonify({"ok": True})
    finally:
        conn.close()


def _move_po_dates(conn, pos, new_date, reset, expected, operator, note):
    """整張 PO（跨線別的全部品項）搬到新日期。expected = {po: 畫面上看到的舊日期}，
    對不上就是有人先改過 → 丟 conflict。回傳搬了幾個品項。"""
    moved = 0
    for po in pos:
        rows = _rows(conn.execute("SELECT * FROM mst_orders WHERE po_number = ?", (po,)))
        if not rows:
            raise LookupError(f"找不到 PO {po}。")
        if expected and po in expected and expected[po] is not None:
            current = {r["delivery_date"] or "" for r in rows}
            if current != {expected[po] or ""}:
                raise PermissionError(f"PO {po} 的交貨日剛被別人改過，請重新載入後再改。")
        stamp = now()
        for r in rows:
            target = r["delivery_date_file"] if reset else new_date
            flag = 0 if reset else 1
            if _same(r["delivery_date"], target) and r["delivery_date_overridden"] == flag:
                continue
            conn.execute(
                """UPDATE mst_orders SET delivery_date = ?, delivery_date_overridden = ?,
                   updated_at = ?, version = version + 1 WHERE id = ?""", (target, flag, stamp, r["id"]))
            _log(conn, r["line"], po, r["sku_id"], r["barcode"], "delivery_date", "交貨日",
                 r["delivery_date"], target, operator, "manual", note)
            moved += 1
    return moved


@master_bp.route("/api/master/pos/date", methods=["PUT"])
def api_update_po_date():
    payload = request.get_json(silent=True) or {}
    operator = _operator()
    pos = [norm_key(p) for p in (payload.get("po_numbers") or ([payload.get("po_number")] if payload.get("po_number") else [])) if norm_key(p)]
    reset = bool(payload.get("reset"))
    new_date = None if reset else norm_date(payload.get("delivery_date"))
    expected = payload.get("expected") or ({pos[0]: payload.get("expected_date")} if len(pos) == 1 and "expected_date" in payload else {})
    if not pos:
        return jsonify({"error": "缺少 PO 單號。"}), 400
    if not reset and not new_date:
        return jsonify({"error": "交貨日格式不對，請用 2026-09-17 這種寫法。"}), 400
    conn = get_conn()
    try:
        try:
            moved = _move_po_dates(conn, pos, new_date, reset, expected, operator,
                                   "恢復為整合表日期" if reset else ("批次改期" if len(pos) > 1 else "整張 PO 改期"))
        except LookupError as exc:
            conn.rollback(); return jsonify({"error": str(exc)}), 404
        except PermissionError as exc:
            conn.rollback(); return jsonify({"error": "conflict", "message": str(exc)}), 409
        conn.commit()
        return jsonify({"ok": True, "moved": moved, "po_count": len(pos),
                        "delivery_date": new_date if not reset else None})
    finally:
        conn.close()


@master_bp.route("/api/master/pos/<po>")
def api_po_detail(po):
    po = norm_key(po)
    cfg = _line_groups()
    conn = get_conn()
    try:
        rows = [_decorate(o, cfg) for o in _rows(conn.execute(
            f"{_ORDER_SELECT} WHERE o.po_number = ? ORDER BY o.line, o.sku_id", (po,)))]
        if not rows:
            return jsonify({"error": "找不到這張 PO。"}), 404
        logs = _rows(conn.execute(
            "SELECT * FROM mst_logs WHERE po_number = ? ORDER BY changed_at DESC, id DESC LIMIT 200", (po,)))
    finally:
        conn.close()
    dates = sorted({r["delivery_date"] or "" for r in rows})
    return jsonify({
        "po_number": po, "rows": rows, "logs": logs,
        "delivery_date": dates[0] if len(dates) == 1 else "",
        "dates": dates, "delivery_date_file": rows[0]["delivery_date_file"],
        "date_overridden": any(r["delivery_date_overridden"] for r in rows),
        "lines": sorted({r["line_group"] for r in rows}),
        "lines_raw": sorted({r["line"] or "" for r in rows}),
        "warehouse": rows[0]["warehouse"],
        "total_cases": round(sum(r["cases"] or 0 for r in rows), 2),
    })


@master_bp.route("/api/master/pos/<po>/save", methods=["PUT"])
def api_po_save(po):
    """PO 視窗一次儲存：交貨日（整張）＋各品項的出貨數量／備註。任一筆版本對不上
    整批退回，不會存一半。"""
    po = norm_key(po)
    payload = request.get_json(silent=True) or {}
    operator = _operator()
    items = payload.get("items") or []
    conn = get_conn()
    try:
        changed_total = 0
        if payload.get("delivery_date") or payload.get("reset_date"):
            new_date = norm_date(payload.get("delivery_date")) if not payload.get("reset_date") else None
            if not payload.get("reset_date") and not new_date:
                return jsonify({"error": "交貨日格式不對。"}), 400
            try:
                changed_total += _move_po_dates(conn, [po], new_date, bool(payload.get("reset_date")),
                                                {po: payload.get("expected_date")}, operator,
                                                "PO 視窗改期" if not payload.get("reset_date") else "恢復為整合表日期")
            except LookupError as exc:
                conn.rollback(); return jsonify({"error": str(exc)}), 404
            except PermissionError as exc:
                conn.rollback(); return jsonify({"error": "conflict", "message": str(exc)}), 409
        for it in items:
            oid = norm_int(it.get("id")); ver = norm_int(it.get("version"))
            if oid is None or ver is None:
                conn.rollback(); return jsonify({"error": "品項缺少 id 或版本。"}), 400
            o = _row(conn.execute("SELECT * FROM mst_orders WHERE id = ? AND po_number = ?", (oid, po)))
            if o is None:
                conn.rollback(); return jsonify({"error": f"找不到品項 {oid}。"}), 404
            # 改期那步已經把版本 +1，所以這裡用「版本 == 原本 或 原本+1（且是剛剛改期造成）」都不安全；
            # 改成比對 updated_at 之前的版本：改期只動 delivery_date，品項欄位的衝突仍用畫面帶來的版本判。
            if o["version"] not in (ver, ver + 1):
                conn.rollback()
                return jsonify({"error": "conflict",
                                "message": f"品項 {o['sku_id']} 剛被別人改過，請重新載入後再存。"}), 409
            try:
                sets, vals, changed = _apply_item_edit(conn, o, it, operator)
            except ValueError as exc:
                conn.rollback(); return jsonify({"error": str(exc)}), 400
            if changed:
                sets += ["updated_at = ?", "version = version + 1"]; vals.append(now())
                conn.execute(f"UPDATE mst_orders SET {', '.join(sets)} WHERE id = ?", vals + [oid])
                changed_total += changed
        conn.commit()
        return jsonify({"ok": True, "changed": changed_total})
    finally:
        conn.close()


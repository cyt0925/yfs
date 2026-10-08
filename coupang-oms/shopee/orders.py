"""蝦皮特選第一段：上傳採購單／入庫單 → 預覽 → 選履約方式確認匯入；重傳比對差異；訂單總覽；改履約方式；上傳紀錄。

規則（需求 3.1.1、第四章，Jerry 2026-10-07「照建議」）：
- 同一張單：採購單用「PR 號碼＋蝦皮商品編號」認；先用 PR 傳、之後變 PO 再傳算同一張，PO 號碼補上去。
- 入庫單：用「PO 號碼＋蝦皮商品編號」找回採購單那筆，補上入庫單號；數量、到貨日、倉別不一樣就算差異。
  找不到採購單的（只拿到入庫單）就直接建一筆。
- 重傳：新的蓋舊的，舊的每一版都留著（shp_order_versions）。有差異 → 那張單「匯入差異待確認」，人確認後才算處理完；
  內容完全一樣只記上傳紀錄、不跳警示。同一張 PR 重傳時少掉的品項算「品項移除」（數量改 0、留紀錄）。
- 履約方式匯入時人選一次（同一個檔的新單都套用），之後可以整張改，要寫原因、留紀錄。
"""
from .common import (
    FULFILS, ST_DIFF, ST_IMPORTED, SYSTEM, _log, _operator, _row, _rows, base64, current_app, get_conn, io, json,
    jsonify, now, render_template, request, send_file, shopee_bp, wh_name,
)
from .parse import KIND_TEXT, ParseError, parse_file

# 重傳時要比的欄位：不一樣就算差異、要人確認
TRACKED = {"qty": "數量", "expected_date": "到貨日", "warehouse": "倉別", "supplier_sku_id": "廠商料號"}
# 只是補資料、不算差異（PR 變 PO、拿到入庫單號、品名小改）
INFO = ("po_id", "inbound_ids", "sku_name", "ean", "selling_type", "base_code", "unit_code", "unit_name", "line")


def _okey_pr(pr, sku):
    return f"PR:{pr}|{sku}"


def _okey_po(po, sku):
    return f"PO:{po}|{sku}"


def order_no(o):
    """畫面上的訂單編號：有 PO 用 PO，還沒有就用 PR。"""
    return o.get("po_id") or o.get("pr_id") or ""


def _find(conn, kind, r):
    if kind == "purchase":
        o = _row(conn.execute("SELECT * FROM shp_orders WHERE okey = ?", (_okey_pr(r["pr_id"], r["shopee_sku_id"]),)))
        if o is None and r.get("po_id"):        # 之前只用入庫單建過這張
            o = _row(conn.execute("SELECT * FROM shp_orders WHERE okey = ?", (_okey_po(r["po_id"], r["shopee_sku_id"]),)))
        return o
    o = _row(conn.execute("SELECT * FROM shp_orders WHERE po_id = ? AND shopee_sku_id = ? ORDER BY id LIMIT 1",
                          (r["po_id"], r["shopee_sku_id"])))
    if o is None:      # 採購單還在 PR 階段（沒有 PO 號碼）：同一個商品、同倉、同到貨日、還沒有 PO 的那一筆
        cands = _rows(conn.execute(
            "SELECT * FROM shp_orders WHERE po_id = '' AND shopee_sku_id = ? AND warehouse = ? AND expected_date = ?",
            (r["shopee_sku_id"], r["warehouse"], r["expected_date"])))
        o = cands[0] if len(cands) == 1 else None
    return o


def _new_fields(kind, r):
    """這一列要寫進訂單的欄位值。"""
    f = {k: r.get(k, "") for k in ("warehouse", "expected_date", "shopee_sku_id", "supplier_sku_id", "base_code",
                                    "unit_code", "unit_name", "line", "sku_name", "selling_type")}
    f["qty"] = r["qty"]
    f["po_id"] = r.get("po_id", "")
    if kind == "purchase":
        f["pr_id"] = r["pr_id"]
        f["ean"] = r.get("ean", "")
        if r.get("inbound_id"):
            f["inbound_ids"] = r["inbound_id"]
    else:
        f["inbound_ids"] = r.get("inbound_id", "")
        f["inbound_qty"] = r["qty"]
    return f


def plan(conn, kind, rows):
    """算出這個檔匯進來會怎樣，不寫資料庫。預覽跟確認匯入共用同一套，確認時重算一次（中間有人動過也不會錯）。"""
    acts, diffs = [], []
    for r in rows:
        o = _find(conn, kind, r)
        new = _new_fields(kind, r)
        if o is None:
            acts.append({"op": "new", "row": r, "fields": new})
            continue
        ch = []
        for k, label in TRACKED.items():
            if k == "warehouse" and not new.get(k):
                continue
            if str(new.get(k, "")) != str(o.get(k) if o.get(k) is not None else ""):
                ch.append({"field": k, "label": label, "old": o.get(k), "new": new.get(k)})
        if o.get("po_id") and new.get("po_id") and new["po_id"] != o["po_id"]:
            ch.append({"field": "po_id", "label": "PO 號碼", "old": o["po_id"], "new": new["po_id"]})
        if o.get("removed"):
            ch.append({"field": "removed", "label": "品項", "old": "已移除", "new": "重新出現"})
        info = {k: new[k] for k in INFO if k in new and new[k] and str(new[k]) != str(o.get(k) or "")}
        if kind == "inbound" and o.get("inbound_qty") != new.get("inbound_qty"):
            info["inbound_qty"] = new["inbound_qty"]
        if kind == "purchase" and not o.get("pr_id"):
            info["pr_id"] = new["pr_id"]
        if not ch and not info:
            acts.append({"op": "same", "order": o})
            continue
        acts.append({"op": "update", "order": o, "fields": new, "changes": ch, "info": info, "row": r})
        for c in ch:
            diffs.append({**c, "order_id": o["id"], "order_no": order_no(o), "supplier_sku_id": o["supplier_sku_id"],
                          "sku_name": o["sku_name"], "qty_delta": (new["qty"] - (o["qty"] or 0)) if c["field"] == "qty" else None})
    # 同一張 PR 重傳時少掉的品項（只看採購單；入庫單常常只傳一部分，不能拿來判斷）
    if kind == "purchase":
        seen = {(r["pr_id"], r["shopee_sku_id"]) for r in rows}
        for pr in sorted({r["pr_id"] for r in rows}):
            for o in _rows(conn.execute("SELECT * FROM shp_orders WHERE pr_id = ? AND removed = 0", (pr,))):
                if (pr, o["shopee_sku_id"]) not in seen:
                    ch = [{"field": "removed", "label": "品項", "old": f"數量 {o['qty']}", "new": "已移除（數量改 0）"}]
                    acts.append({"op": "remove", "order": o, "changes": ch})
                    diffs.append({**ch[0], "order_id": o["id"], "order_no": order_no(o), "supplier_sku_id": o["supplier_sku_id"],
                                  "sku_name": o["sku_name"], "qty_delta": -(o["qty"] or 0)})
    return acts, diffs


def _counts(acts):
    c = {"new": 0, "update": 0, "same": 0, "remove": 0}
    for a in acts:
        c[a["op"]] += 1
    c["changed"] = sum(1 for a in acts if a["op"] == "update" and a["changes"])
    return c


def _snapshot(conn, order_id, version, upload_id, kind):
    o = _row(conn.execute("SELECT * FROM shp_orders WHERE id = ?", (order_id,)))
    data = {k: o[k] for k in ("pr_id", "po_id", "inbound_ids", "warehouse", "expected_date", "shopee_sku_id", "supplier_sku_id",
                               "sku_name", "qty", "inbound_qty", "removed", "fulfil")}
    conn.execute("INSERT INTO shp_order_versions (order_id, version, upload_id, kind, data_json, created_at) VALUES (?,?,?,?,?,?)",
                 (order_id, version, upload_id, kind, json.dumps(data, ensure_ascii=False), now()))


def apply(conn, upload, acts, fulfil, operator):
    """把 plan 的結果寫進去。回 counts。"""
    stamp = now(); uid = upload["id"]; kind = upload["kind"]
    for a in acts:
        if a["op"] == "same":
            continue
        if a["op"] == "new":
            f = a["fields"]
            okey = _okey_pr(f["pr_id"], f["shopee_sku_id"]) if kind == "purchase" else _okey_po(f["po_id"], f["shopee_sku_id"])
            cur = conn.execute(
                """INSERT INTO shp_orders (okey, pr_id, po_id, inbound_ids, warehouse, expected_date, shopee_sku_id, supplier_sku_id,
                   base_code, unit_code, unit_name, line, sku_name, ean, selling_type, qty, qty_original, inbound_qty, fulfil, status,
                   diff_pending, version, first_upload_id, last_upload_id, created_at, updated_at)
                   VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,0,1,?,?,?,?)""",
                (okey, f.get("pr_id", ""), f.get("po_id", ""), f.get("inbound_ids", ""), f["warehouse"], f["expected_date"],
                 f["shopee_sku_id"], f["supplier_sku_id"], f["base_code"], f["unit_code"], f["unit_name"], f["line"], f["sku_name"],
                 f.get("ean", ""), f["selling_type"], f["qty"], f["qty"], f.get("inbound_qty"), fulfil, ST_IMPORTED,
                 uid, uid, stamp, stamp))
            _snapshot(conn, cur.lastrowid, 1, uid, kind)
            continue
        o = a["order"]
        sets, vals = [], []
        if a["op"] == "remove":
            sets += ["qty = ?", "removed = ?"]; vals += [0, 1]
        else:
            f = a["fields"]
            for c in a["changes"]:
                if c["field"] == "removed":
                    sets.append("removed = ?"); vals.append(0)
                elif c["field"] in f:
                    sets.append(f"{c['field']} = ?"); vals.append(f[c["field"]])
            for k, v in a["info"].items():
                sets.append(f"{k} = ?"); vals.append(v)
            if kind == "purchase" and o["okey"].startswith("PO:"):        # 之前只用入庫單建的，補上 PR 之後改用 PR 認
                sets.append("okey = ?"); vals.append(_okey_pr(f["pr_id"], f["shopee_sku_id"]))
        ch = a.get("changes") or []
        if ch:
            sets += ["diff_pending = ?", "status = ?"]; vals += [1, ST_DIFF]
            for c in ch:
                delta = None
                if c["field"] == "qty":
                    delta = (c["new"] or 0) - (c["old"] or 0)
                elif a["op"] == "remove":
                    delta = -(o["qty"] or 0)
                conn.execute("""INSERT INTO shp_diffs (order_id, upload_id, field, field_label, old_value, new_value, qty_delta,
                                uploaded_by, uploaded_at) VALUES (?,?,?,?,?,?,?,?,?)""",
                             (o["id"], uid, c["field"], c["label"], "" if c["old"] is None else str(c["old"]),
                              "" if c["new"] is None else str(c["new"]), delta, operator, stamp))
                _log(conn, o["line"], order_no(o), o["shopee_sku_id"], o["base_code"], f"shopee_{c['field']}",
                     f"蝦皮{c['label']}", c["old"], c["new"], operator, "import", upload["filename"])
        sets += ["version = ?", "last_upload_id = ?", "updated_at = ?"]; vals += [o["version"] + 1, uid, stamp]
        conn.execute(f"UPDATE shp_orders SET {', '.join(sets)} WHERE id = ?", vals + [o["id"]])
        _snapshot(conn, o["id"], o["version"] + 1, uid, kind)
    if kind == "purchase":       # PR 變成 PO：同一張 PR 的品項（包括這次被移除的）都補上 PO 號碼，總覽才會在同一張單底下
        for pr, po in sorted({(a["row"]["pr_id"], a["row"]["po_id"]) for a in acts if a.get("row") and a["row"].get("po_id")}):
            conn.execute("UPDATE shp_orders SET po_id = ? WHERE pr_id = ? AND po_id = ''", (po, pr))
    return _counts(acts)


# ── 路由 ─────────────────────────────────────────────────────────────────────

@shopee_bp.route("/shopee")
def shopee_page():
    return render_template("shopee.html", logged_in_user=_operator(), fulfils=FULFILS,
                           build_version=current_app.config.get("BUILD_VERSION", ""))


@shopee_bp.route("/api/shopee/import/preview", methods=["POST"])
def api_import_preview():
    f = request.files.get("file")
    if f is None or not f.filename:
        return jsonify({"error": "未收到檔案，請重新上傳。"}), 400
    data = f.read()
    try:
        parsed = parse_file(data, f.filename)
    except ParseError as exc:
        return jsonify({"error": str(exc)}), 400
    if not parsed["rows"]:
        return jsonify({"error": "檔案裡沒有可匯入的品項。", "details": parsed["errors"][:30]}), 400
    operator = _operator()
    conn = get_conn()
    try:
        acts, diffs = plan(conn, parsed["kind"], parsed["rows"])
        c = _counts(acts)
        cur = conn.execute(
            """INSERT INTO shp_uploads (kind, filename, content_b64, payload_json, rows_total, new_count, changed_count, same_count,
               removed_count, errors_json, status, uploaded_by, uploaded_at) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (parsed["kind"], f.filename, base64.b64encode(data).decode("ascii"), json.dumps(parsed["rows"], ensure_ascii=False),
             len(parsed["rows"]), c["new"], c["changed"], c["same"], c["remove"], json.dumps(parsed["errors"], ensure_ascii=False),
             "pending", operator, now()))
        conn.commit()
        uid = cur.lastrowid
    finally:
        conn.close()
    orders = sorted({(r.get("po_id") or r.get("pr_id")) for r in parsed["rows"]})
    return jsonify({"upload_id": uid, "kind": parsed["kind"], "kind_text": KIND_TEXT[parsed["kind"]], "filename": f.filename,
                    "rows_total": len(parsed["rows"]), "orders": orders, "lines": parsed["lines"],
                    "warehouses": sorted({r["warehouse"] for r in parsed["rows"] if r["warehouse"]}),
                    "dates": sorted({r["expected_date"] for r in parsed["rows"] if r["expected_date"]}),
                    "counts": c, "diffs": diffs[:500], "errors": parsed["errors"], "needs_fulfil": c["new"] > 0})


def _pending_upload(conn, uid):
    u = _row(conn.execute("SELECT * FROM shp_uploads WHERE id = ?", (uid,)))
    if u is None or u["status"] != "pending":
        return None
    return u


@shopee_bp.route("/api/shopee/import/commit", methods=["POST"])
def api_import_commit():
    p = request.get_json(silent=True) or {}
    fulfil = str(p.get("fulfil") or "").strip()
    operator = _operator()
    conn = get_conn()
    try:
        u = _pending_upload(conn, p.get("upload_id"))
        if u is None:
            return jsonify({"error": "找不到這次上傳的預覽，可能已經匯入或取消，請重新上傳。"}), 404
        rows = json.loads(u["payload_json"] or "[]")
        acts, _diffs = plan(conn, u["kind"], rows)
        if any(a["op"] == "new" for a in acts) and fulfil not in FULFILS:
            return jsonify({"error": f"有新訂單，請先選擇履約方式（{'、'.join(FULFILS)}）。"}), 400
        c = apply(conn, u, acts, fulfil, operator)
        conn.execute("""UPDATE shp_uploads SET status = 'ok', fulfil = ?, payload_json = '', new_count = ?, changed_count = ?,
                        same_count = ?, removed_count = ?, committed_at = ? WHERE id = ?""",
                     (fulfil if c["new"] else "", c["new"], c["changed"], c["same"], c["remove"], now(), u["id"]))
        lines = sorted({r["line"] for r in rows if r.get("line")})
        _log(conn, "、".join(lines), "", "", "", "shopee_import", f"匯入蝦皮{KIND_TEXT[u['kind']]}", "",
             f"新增 {c['new']}、有變 {c['changed']}、移除 {c['remove']}、無異動 {c['same']}", operator, "import",
             u["filename"] + (f"（履約方式：{fulfil}）" if c["new"] else ""))
        conn.commit()
    finally:
        conn.close()
    return jsonify({"ok": True, "counts": c})


@shopee_bp.route("/api/shopee/import/cancel", methods=["POST"])
def api_import_cancel():
    p = request.get_json(silent=True) or {}
    conn = get_conn()
    try:
        conn.execute("UPDATE shp_uploads SET status = 'canceled', payload_json = '' WHERE id = ? AND status = 'pending'", (p.get("upload_id"),))
        conn.commit()
    finally:
        conn.close()
    return jsonify({"ok": True})


def _order_filter(args):
    where, params = ["1=1"], []
    d1, d2 = str(args.get("from") or ""), str(args.get("to") or "")
    if d1:
        where.append("expected_date >= ?"); params.append(d1)
    if d2:
        where.append("expected_date <= ?"); params.append(d2)
    for k in ("line", "fulfil", "warehouse"):
        if args.get(k):
            where.append(f"{k} = ?"); params.append(args.get(k))
    if args.get("pending") in ("1", "true"):
        where.append("diff_pending = 1")
    q = str(args.get("q") or "").strip()
    if q:
        where.append("(pr_id LIKE ? OR po_id LIKE ? OR inbound_ids LIKE ? OR shopee_sku_id LIKE ? OR supplier_sku_id LIKE ? OR sku_name LIKE ?)")
        params += [f"%{q}%"] * 6
    return " AND ".join(where), params


@shopee_bp.route("/api/shopee/orders")
def api_orders():
    """訂單總覽：一張單（PO，還沒有 PO 就用 PR）一列，底下是品項。"""
    where, params = _order_filter(request.args)
    conn = get_conn()
    try:
        items = _rows(conn.execute(f"SELECT * FROM shp_orders WHERE {where} ORDER BY expected_date, warehouse, id", params))
        pending_all = conn.execute("SELECT COUNT(DISTINCT CASE WHEN po_id != '' THEN po_id ELSE pr_id END) AS n FROM shp_orders "
                                   "WHERE diff_pending = 1").fetchone()["n"]
        from .purchase import po_status_map            # 第二段：採購單產出狀態
        po_st = po_status_map(conn, [i["id"] for i in items if not i["removed"]])
    finally:
        conn.close()
    groups = {}
    for it in items:
        no = order_no(it)
        g = groups.get(no)
        if g is None:
            g = groups[no] = {"order_no": no, "pr_id": it["pr_id"], "po_id": it["po_id"], "line": it["line"],
                              "warehouse": it["warehouse"], "warehouse_name": wh_name(it["warehouse"]),
                              "expected_date": it["expected_date"], "fulfils": set(), "lines": set(), "items": []}
        g["items"].append(it)
        g["fulfils"].add(it["fulfil"] or "")
        g["lines"].add(it["line"] or "")
    out = []
    for g in groups.values():
        its = g["items"]
        alive = [i for i in its if not i["removed"]]
        pending = any(i["diff_pending"] for i in its)
        fulfil_txt = "、".join(sorted(x for x in g.pop("fulfils") if x)) or "未選"
        g.update(
            fulfil=fulfil_txt,
            line="、".join(sorted(x for x in g.pop("lines") if x)),
            item_count=len(alive), removed_count=len(its) - len(alive),
            qty_original=sum(i["qty_original"] or 0 for i in its), qty=sum(i["qty"] or 0 for i in alive),
            inbound_count=sum(1 for i in alive if i["inbound_ids"]),
            status=ST_DIFF if pending else ST_IMPORTED, diff_pending=pending,
            notify_status="未通知",
            po_status=("不需產出" if fulfil_txt == "竹運出貨" else
                       "已回填" if alive and all(po_st.get(i["id"]) == "已回填" for i in alive) else
                       "已產出" if any(i["id"] in po_st for i in alive) else "未產出"),
        )
        out.append(g)
    out.sort(key=lambda g: (not g["diff_pending"], g["expected_date"], g["warehouse"], g["order_no"]))
    return jsonify({"orders": out, "pending_orders": pending_all,
                    "summary": {"orders": len(out), "items": sum(g["item_count"] for g in out), "qty": sum(g["qty"] for g in out),
                                "pending": sum(1 for g in out if g["diff_pending"])}})


def _group_ids(conn, order_no_):
    return [r["id"] for r in _rows(conn.execute(
        "SELECT id FROM shp_orders WHERE po_id = ? OR (po_id = '' AND pr_id = ?)", (order_no_, order_no_)))]


@shopee_bp.route("/api/shopee/diffs")
def api_diffs():
    """某張單的差異（含已確認的），給差異視窗。"""
    no = str(request.args.get("order_no") or "").strip()
    conn = get_conn()
    try:
        ids = _group_ids(conn, no)
        if not ids:
            return jsonify({"diffs": []})
        marks = ",".join("?" * len(ids))
        rows = _rows(conn.execute(
            f"""SELECT d.*, o.supplier_sku_id, o.sku_name, o.pr_id, o.po_id FROM shp_diffs d JOIN shp_orders o ON o.id = d.order_id
                WHERE d.order_id IN ({marks}) ORDER BY d.id DESC""", ids))
    finally:
        conn.close()
    for r in rows:
        r["order_no"] = r["po_id"] or r["pr_id"]
        r["confirmed"] = bool(r["confirmed_at"])
    return jsonify({"diffs": rows})


@shopee_bp.route("/api/shopee/diffs/confirm", methods=["POST"])
def api_diffs_confirm():
    """確認一張單的差異（需求 3.1.1：操作人員確認差異後才能繼續後續作業）。"""
    p = request.get_json(silent=True) or {}
    nos = [str(x).strip() for x in (p.get("order_nos") or []) if str(x).strip()]
    if not nos:
        return jsonify({"error": "請選擇要確認的訂單。"}), 400
    operator = _operator(); stamp = now(); done = 0
    conn = get_conn()
    try:
        for no in nos:
            ids = _group_ids(conn, no)
            if not ids:
                continue
            marks = ",".join("?" * len(ids))
            n = conn.execute(f"SELECT COUNT(*) AS n FROM shp_diffs WHERE confirmed_at = '' AND order_id IN ({marks})", ids).fetchone()["n"]
            if not n:
                continue
            conn.execute(f"UPDATE shp_diffs SET confirmed_by = ?, confirmed_at = ? WHERE confirmed_at = '' AND order_id IN ({marks})",
                         [operator, stamp] + ids)
            conn.execute(f"UPDATE shp_orders SET diff_pending = 0, status = ?, updated_at = ? WHERE id IN ({marks})",
                         [ST_IMPORTED, stamp] + ids)
            line = _row(conn.execute("SELECT line FROM shp_orders WHERE id = ?", (ids[0],)))["line"]
            _log(conn, line, no, "", "", "shopee_diff_confirm", "確認蝦皮差異", "", f"{n} 項差異", operator, "manual")
            done += 1
        conn.commit()
    finally:
        conn.close()
    return jsonify({"ok": True, "confirmed": done})


@shopee_bp.route("/api/shopee/fulfil", methods=["POST"])
def api_fulfil():
    """整張單改履約方式（需求第四章：保留原履約方式、修改後、修改人員、時間、原因）。"""
    p = request.get_json(silent=True) or {}
    no = str(p.get("order_no") or "").strip()
    fulfil = str(p.get("fulfil") or "").strip()
    reason = str(p.get("reason") or "").strip()
    if fulfil not in FULFILS:
        return jsonify({"error": f"履約方式只能選：{'、'.join(FULFILS)}。"}), 400
    if not reason:
        return jsonify({"error": "請填寫修改原因。"}), 400
    operator = _operator()
    conn = get_conn()
    try:
        ids = _group_ids(conn, no)
        if not ids:
            return jsonify({"error": "找不到這張訂單，請重新整理頁面。"}), 404
        marks = ",".join("?" * len(ids))
        olds = sorted({r["fulfil"] or "未選" for r in _rows(conn.execute(f"SELECT fulfil FROM shp_orders WHERE id IN ({marks})", ids))})
        if olds == [fulfil]:
            return jsonify({"ok": True, "changed": False})
        conn.execute(f"UPDATE shp_orders SET fulfil = ?, updated_at = ? WHERE id IN ({marks})", [fulfil, now()] + ids)
        line = _row(conn.execute("SELECT line FROM shp_orders WHERE id = ?", (ids[0],)))["line"]
        _log(conn, line, no, "", "", "shopee_fulfil", "履約方式", "、".join(olds), fulfil, operator, "manual", reason=reason)
        conn.commit()
    finally:
        conn.close()
    return jsonify({"ok": True, "changed": True})


@shopee_bp.route("/api/shopee/versions/<int:order_id>")
def api_versions(order_id):
    conn = get_conn()
    try:
        rows = _rows(conn.execute("""SELECT v.version, v.kind, v.created_at, v.data_json, u.filename, u.uploaded_by
                                     FROM shp_order_versions v LEFT JOIN shp_uploads u ON u.id = v.upload_id
                                     WHERE v.order_id = ? ORDER BY v.version DESC""", (order_id,)))
    finally:
        conn.close()
    for r in rows:
        r["data"] = json.loads(r.pop("data_json") or "{}")
    return jsonify({"versions": rows})


@shopee_bp.route("/api/shopee/uploads")
def api_uploads():
    conn = get_conn()
    try:
        rows = _rows(conn.execute("""SELECT id, kind, filename, fulfil, rows_total, new_count, changed_count, same_count, removed_count,
                                     errors_json, status, uploaded_by, uploaded_at, committed_at FROM shp_uploads
                                     WHERE status = 'ok' ORDER BY id DESC LIMIT 200"""))
    finally:
        conn.close()
    for r in rows:
        r["kind_text"] = KIND_TEXT.get(r["kind"], r["kind"])
        r["errors"] = json.loads(r.pop("errors_json") or "[]")
    return jsonify({"uploads": rows})


@shopee_bp.route("/api/shopee/uploads/<int:uid>/file")
def api_upload_file(uid):
    conn = get_conn()
    try:
        u = _row(conn.execute("SELECT filename, content_b64 FROM shp_uploads WHERE id = ?", (uid,)))
    finally:
        conn.close()
    if u is None or not u["content_b64"]:
        return jsonify({"error": "找不到這次上傳的原始檔。"}), 404
    return send_file(io.BytesIO(base64.b64decode(u["content_b64"])), as_attachment=True, download_name=u["filename"] or f"upload_{uid}.xlsx")


def portal_summary(conn, today_iso):
    """YFS 訂單系統首頁的蝦皮特選卡片與今日事項用。"""
    def one(sql, params=()):
        r = conn.execute(sql, params).fetchone()
        return (next(iter(dict(r).values())) if r is not None else 0) or 0
    no = "CASE WHEN po_id != '' THEN po_id ELSE pr_id END"
    import datetime as _dt
    d0 = _dt.date.fromisoformat(today_iso) - _dt.timedelta(days=29)
    rows = conn.execute(f"""SELECT SUBSTR(created_at, 1, 10) AS d, COUNT(DISTINCT {no}) AS n FROM shp_orders
                            WHERE created_at >= ? GROUP BY SUBSTR(created_at, 1, 10)""", (d0.isoformat(),)).fetchall()
    by_day = {r["d"]: int(r["n"]) for r in rows}
    daily = [{"date": (d0 + _dt.timedelta(days=i)).isoformat(), "value": by_day.get((d0 + _dt.timedelta(days=i)).isoformat(), 0)} for i in range(30)]
    return {
        "daily_new": daily,
        "new_today": one(f"SELECT COUNT(DISTINCT {no}) AS n FROM shp_orders WHERE created_at >= ?", (today_iso,)),
        "pending_pos": one(f"SELECT COUNT(DISTINCT {no}) AS n FROM shp_orders WHERE expected_date >= ? AND removed = 0", (today_iso,)),
        "diff_pos": one(f"SELECT COUNT(DISTINCT {no}) AS n FROM shp_orders WHERE diff_pending = 1"),
        "last_import": one("SELECT MAX(committed_at) AS m FROM shp_uploads WHERE status = 'ok'") or "",
    }


__all__ = ["plan", "apply", "order_no", "portal_summary", "SYSTEM"]

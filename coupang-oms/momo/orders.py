"""MOMO 第一批：上傳 MOMO 下載檔（一次好幾個檔）→ 合在一起預覽 → 確認匯入；重傳比對差異；訂單總覽；出貨方式；上傳紀錄。

規則（照蝦皮特選的做法，Jerry 2026-10-28 看過示意圖「好 開始做」）：
- 同一筆：入庫單號＋商品編號＋單品編號。同一份或不同份檔裡重複出現，以後面讀到的為準，預覽講清楚。
- 重傳：內容一樣只記上傳紀錄；數量、到貨日、倉別、商品原廠編號不一樣 → 記差異、那張單「差異待確認」，確認後才能往下做。
  同一張入庫單重傳時少掉的品項算移除（數量改 0、留紀錄）；沒出現在這次檔案裡的入庫單不動（MOMO 下載檔常常只有一部分）。
- 每一版都留著（momo_order_versions）。
"""
from .common import *  # noqa: F401,F403
from .parse import ParseError, parse_file

TRACKED = {"qty": "數量", "expected_date": "到貨日", "warehouse": "倉別", "factory_code": "商品原廠編號"}
INFO = ("product_name", "spec", "shelf_days", "min_expiry", "momo_note", "purchase_date", "base_code", "unit_code", "unit_name", "line")
# 來源檔名、工作表、列號留第一次匯入的就好；同一份換個檔名再傳不算有變


def okey(r):
    return f"{r['inbound_no']}|{r['product_no']}|{r['item_no']}"


def plan(conn, rows):
    """這批匯進來會怎樣，不寫資料庫。預覽跟確認共用，確認時重算一次。"""
    acts, diffs = [], []
    seen = {}
    for r in rows:                                       # 同一筆出現兩次：後面的為準
        seen[okey(r)] = r
    for k, r in seen.items():
        o = _row(conn.execute("SELECT * FROM momo_orders WHERE okey = ?", (k,)))
        if o is None:
            acts.append({"op": "new", "row": r})
            continue
        ch = [{"field": f, "label": lab, "old": o[f], "new": r[f]} for f, lab in TRACKED.items() if str(r[f]) != str(o[f] if o[f] is not None else "")]
        if o["removed"]:
            ch.append({"field": "removed", "label": "品項", "old": "已移除", "new": "重新出現"})
        info = {f: r[f] for f in INFO if str(r.get(f, "")) != str(o.get(f) if o.get(f) is not None else "")}
        if not ch and not info:
            acts.append({"op": "same", "order": o})
            continue
        acts.append({"op": "update", "order": o, "row": r, "changes": ch, "info": info})
        for c in ch:
            diffs.append({**c, "order_id": o["id"], "inbound_no": o["inbound_no"], "factory_code": o["factory_code"], "product_name": o["product_name"]})
    nos = sorted({r["inbound_no"] for r in seen.values()})
    for no in nos:                                       # 同一張入庫單少掉的品項
        for o in _rows(conn.execute("SELECT * FROM momo_orders WHERE inbound_no = ? AND removed = 0", (no,))):
            if o["okey"] not in seen:
                ch = [{"field": "removed", "label": "品項", "old": f"數量 {o['qty']}", "new": "已移除（數量改 0）"}]
                acts.append({"op": "remove", "order": o, "changes": ch})
                diffs.append({**ch[0], "order_id": o["id"], "inbound_no": no, "factory_code": o["factory_code"], "product_name": o["product_name"]})
    return acts, diffs, len(rows) - len(seen)


def counts(acts):
    c = {"new": 0, "update": 0, "same": 0, "remove": 0}
    for a in acts:
        c[a["op"]] += 1
    c["changed"] = sum(1 for a in acts if a["op"] == "update" and a["changes"])
    return c


SNAP = ("inbound_no", "warehouse", "expected_date", "product_no", "item_no", "product_name", "factory_code", "qty", "removed", "ship_method")


def _snapshot(conn, order_id, version, upload_id):
    o = _row(conn.execute("SELECT * FROM momo_orders WHERE id = ?", (order_id,)))
    conn.execute("INSERT INTO momo_order_versions (order_id, version, upload_id, data_json, created_at) VALUES (?,?,?,?,?)",
                 (order_id, version, upload_id, json.dumps({k: o[k] for k in SNAP}, ensure_ascii=False), now()))


NEW_COLS = ("inbound_no", "warehouse", "expected_date", "purchase_date", "product_no", "item_no", "product_name", "spec", "shelf_days",
            "min_expiry", "factory_code", "base_code", "unit_code", "unit_name", "line", "qty", "momo_note", "source_file", "source_sheet",
            "source_row")


def apply(conn, upload, acts, operator):
    stamp, uid = now(), upload["id"]
    for a in acts:
        if a["op"] == "same":
            continue
        if a["op"] == "new":
            r = a["row"]
            cols = ("okey",) + NEW_COLS + ("qty_original", "ship_method", "first_upload_id", "last_upload_id", "created_at", "updated_at")
            vals = [okey(r)] + [r[c] for c in NEW_COLS] + [r["qty"], DIRECT, uid, uid, stamp, stamp]
            cur = conn.execute(f"INSERT INTO momo_orders ({', '.join(cols)}) VALUES ({', '.join('?' * len(cols))})", vals)
            _snapshot(conn, cur.lastrowid, 1, uid)
            continue
        o = a["order"]
        sets, vals = [], []
        if a["op"] == "remove":
            sets += ["qty = ?", "removed = ?"]; vals += [0, 1]
        else:
            r = a["row"]
            for c in a["changes"]:
                if c["field"] == "removed":
                    sets.append("removed = ?"); vals.append(0)
                else:
                    sets.append(f"{c['field']} = ?"); vals.append(r[c["field"]])
            for k, v in a["info"].items():
                sets.append(f"{k} = ?"); vals.append(v)
        ch = a.get("changes") or []
        if ch:
            sets.append("diff_pending = ?"); vals.append(1)
            for c in ch:
                conn.execute("""INSERT INTO momo_diffs (order_id, upload_id, field, field_label, old_value, new_value, uploaded_by, uploaded_at)
                                VALUES (?,?,?,?,?,?,?,?)""", (o["id"], uid, c["field"], c["label"], "" if c["old"] is None else str(c["old"]),
                                                              "" if c["new"] is None else str(c["new"]), operator, stamp))
                _log(conn, o["line"], o["inbound_no"], o["product_no"], "", f"momo_{c['field']}", f"MOMO{c['label']}", c["old"], c["new"],
                     operator, "import", upload["filenames"])
        sets += ["version = ?", "last_upload_id = ?", "updated_at = ?"]; vals += [o["version"] + 1, uid, stamp]
        conn.execute(f"UPDATE momo_orders SET {', '.join(sets)} WHERE id = ?", vals + [o["id"]])
        _snapshot(conn, o["id"], o["version"] + 1, uid)
    return counts(acts)


# ── 路由 ─────────────────────────────────────────────────────────────────────
@momo_bp.route("/momo")
def momo_page():
    from master.common import _is_admin
    return render_template("momo.html", logged_in_user=_operator(), is_admin=_is_admin(), ship_methods=SHIP_METHODS,
                           build_version=current_app.config.get("BUILD_VERSION", ""))


def _preview_out(uid, names, rows, errors, warnings, c, diffs, dup):
    by_line = collections.Counter(r["line"] or "認不出" for r in rows)
    return {"upload_id": uid, "filenames": names, "rows_total": len(rows), "orders": sorted({r["inbound_no"] for r in rows}),
            "lines": dict(by_line), "warehouses": sorted({r["warehouse"] for r in rows}), "dates": sorted({r["expected_date"] for r in rows}),
            "counts": c, "diffs": diffs[:300], "errors": errors[:200], "warnings": warnings[:50], "dup": dup}


@momo_bp.route("/api/momo/import/preview", methods=["POST"])
def api_import_preview():
    """一次好幾個檔（Jerry 2026-10-28「一次丟好幾個檔」）：每個檔各自讀，合在一起預覽。讀不了的檔列出來，其他照樣預覽。"""
    files = [f for f in request.files.getlist("file") if f and f.filename]
    if not files:
        return jsonify({"error": "未收到檔案，請重新上傳。"}), 400
    rows, errors, warnings, names, blobs = [], [], [], [], []
    for f in files:
        data = f.read()
        try:
            d = parse_file(data, f.filename)
        except ParseError as exc:
            errors.append({"file": f.filename, "sheet": "", "row": "", "warehouse": "", "reason": str(exc)})
            continue
        if not d["rows"] and not d["errors"]:
            warnings.append(f"{f.filename}：沒有採購數量大於 0 的品項")
        rows += d["rows"]; errors += d["errors"]; warnings += d["warnings"]
        names.append(f.filename); blobs.append((f.filename, data))
    if not rows:
        return jsonify({"error": "這幾個檔裡沒有可匯入的品項。", "details": [f"{e['file']} {e['row']}：{e['reason']}" for e in errors][:30]}), 400
    operator = _operator()
    conn = get_conn()
    try:
        acts, diffs, dup = plan(conn, rows)
        c = counts(acts)
        cur = conn.execute("""INSERT INTO momo_uploads (filenames, payload_json, rows_total, new_count, changed_count, same_count, removed_count,
                              errors_json, status, uploaded_by, uploaded_at) VALUES (?,?,?,?,?,?,?,?,?,?,?)""",
                           ("、".join(names), json.dumps(rows, ensure_ascii=False), len(rows), c["new"], c["changed"], c["same"], c["remove"],
                            json.dumps(errors, ensure_ascii=False), "pending", operator, now()))
        uid = cur.lastrowid
        for n, b in blobs:
            conn.execute("INSERT INTO momo_upload_files (upload_id, filename, content_b64) VALUES (?,?,?)", (uid, n, base64.b64encode(b).decode("ascii")))
        conn.commit()
    finally:
        conn.close()
    return jsonify(_preview_out(uid, names, rows, errors, warnings, c, diffs, dup))


@momo_bp.route("/api/momo/import/pending")
def api_import_pending():
    """今天上傳了、還沒按「確認匯入」的：重新整理頁面後把預覽放回來。"""
    conn = get_conn()
    try:
        ups = _rows(conn.execute("SELECT id, filenames, payload_json, errors_json FROM momo_uploads WHERE status = 'pending' AND uploaded_by = ? "
                                 "AND uploaded_at >= ? ORDER BY id", (_operator(), now()[:10])))
        out = []
        for u in ups:
            rows = json.loads(u["payload_json"] or "[]")
            if rows:
                acts, diffs, dup = plan(conn, rows)
                out.append(_preview_out(u["id"], u["filenames"].split("、"), rows, json.loads(u["errors_json"] or "[]"), [], counts(acts), diffs, dup))
    finally:
        conn.close()
    return jsonify({"pending": out})


@momo_bp.route("/api/momo/import/commit", methods=["POST"])
def api_import_commit():
    p = request.get_json(silent=True) or {}
    operator = _operator()
    conn = get_conn()
    try:
        u = _row(conn.execute("SELECT * FROM momo_uploads WHERE id = ?", (p.get("upload_id"),)))
        if u is None or u["status"] != "pending":
            return jsonify({"error": "找不到這次上傳的預覽，可能已經匯入或取消，請重新上傳。"}), 404
        rows = json.loads(u["payload_json"] or "[]")
        acts, _d, _dup = plan(conn, rows)
        c = apply(conn, u, acts, operator)
        conn.execute("""UPDATE momo_uploads SET status = 'ok', payload_json = '', new_count = ?, changed_count = ?, same_count = ?, removed_count = ?,
                        committed_at = ? WHERE id = ?""", (c["new"], c["changed"], c["same"], c["remove"], now(), u["id"]))
        lines = sorted({r["line"] for r in rows if r.get("line")})
        _log(conn, "、".join(lines), "", "", "", "momo_import", "匯入 MOMO 下載檔", "",
             f"新增 {c['new']}、有變 {c['changed']}、移除 {c['remove']}、無異動 {c['same']}", operator, "import", u["filenames"])
        conn.commit()
    finally:
        conn.close()
    return jsonify({"ok": True, "counts": c})


@momo_bp.route("/api/momo/import/cancel", methods=["POST"])
def api_import_cancel():
    p = request.get_json(silent=True) or {}
    conn = get_conn()
    try:
        conn.execute("UPDATE momo_uploads SET status = 'canceled', payload_json = '' WHERE id = ? AND status = 'pending'", (p.get("upload_id"),))
        conn.commit()
    finally:
        conn.close()
    return jsonify({"ok": True})


# ── 訂單總覽：一張入庫單一列 ─────────────────────────────────────────────
STAGES = (("diff", "差異待確認", "確認差異"), ("buy", "待產 EIP 採購單", "產出 EIP 採購單"), ("eip", "待回填 EIP 採購單號", "回填 EIP 採購單號"),
          ("yx", "待核缺貨與效期", "核對缺貨與效期"), ("notice", "待畫入庫通知單", "畫入庫通知單"))


def _filter(args):
    """搜尋、篩選（Jerry 2026-10-28「要記得做搜尋功能」）。線別、倉別可以複選（逗號分開）。"""
    where, params = ["1=1"], []
    for k, col in (("from", "expected_date >="), ("to", "expected_date <=")):
        if args.get(k):
            where.append(f"{col} ?"); params.append(str(args.get(k)))
    up1, up2 = str(args.get("up_from") or ""), str(args.get("up_to") or "")
    if up1 or up2:                                       # 上傳日＝那張入庫單第一次進系統那天
        hv, hp = [], []
        if up1:
            hv.append("MIN(created_at) >= ?"); hp.append(up1)
        if up2:
            hv.append("MIN(created_at) <= ?"); hp.append(up2 + " 23:59:59")
        where.append(f"inbound_no IN (SELECT inbound_no FROM momo_orders GROUP BY inbound_no HAVING {' AND '.join(hv)})"); params += hp
    for k in ("line", "warehouse", "ship_method"):
        vals = [v for v in str(args.get(k) or "").split(",") if v.strip()]
        if vals:
            where.append(f"inbound_no IN (SELECT inbound_no FROM momo_orders WHERE {k} IN ({','.join('?' * len(vals))}))"); params += vals
    q = str(args.get("q") or "").strip()
    if q:
        where.append("inbound_no IN (SELECT inbound_no FROM momo_orders WHERE inbound_no LIKE ? OR product_no LIKE ? OR factory_code LIKE ? "
                     "OR base_code LIKE ? OR product_name LIKE ?)")
        params += [f"%{q}%"] * 5
    return " AND ".join(where), params


def overview(conn, args):
    from .check import check_orders
    from .purchase import eip_map, skipped_ids
    where, params = _filter(args)
    items = _rows(conn.execute(f"SELECT * FROM momo_orders WHERE {where} ORDER BY expected_date, warehouse, inbound_no, id", params))
    eips = eip_map(conn, [i["id"] for i in items])
    skip = skipped_ids(conn, items)                       # 改成 0（不下採）的品項不算待產
    groups = collections.OrderedDict()
    for it in items:
        g = groups.setdefault(it["inbound_no"], {"inbound_no": it["inbound_no"], "warehouse": it["warehouse"], "expected_date": it["expected_date"],
                                                 "items": []})
        g["items"].append(it)
    whs = wh_settings(conn)
    checks = check_orders(conn, list(groups)) if groups else {}
    out = []
    for g in groups.values():
        its = g["items"]
        alive = [i for i in its if not i["removed"]]
        buy = [i for i in alive if i["ship_method"] == DIRECT and (i["qty"] or 0) > 0 and i["id"] not in skip]
        pending = any(i["diff_pending"] for i in its)
        g.update(lines=sorted({i["line"] or "認不出" for i in alive}), ship_method="、".join(sorted({i["ship_method"] for i in alive})) or DIRECT,
                 item_count=len(alive), removed_count=len(its) - len(alive), qty=sum(i["qty"] or 0 for i in alive),
                 upload_date=min((i["created_at"] or "" for i in its), default="")[:10], diff_pending=pending,
                 wh_known=bool(wh_canon(g["warehouse"], whs)),
                 eip_pos=sorted({eips[i["id"]]["eip_po"] for i in alive if i["id"] in eips and eips[i["id"]]["eip_po"]}),
                 po_status="不需產出" if not buy else "已回填" if all(eips.get(i["id"], {}).get("eip_po") for i in buy)
                 else "已產出" if any(i["id"] in eips for i in buy) else "未產出")
        todo = {"diff": pending, "buy": any(i["id"] not in eips for i in buy),
                "eip": any(i["id"] in eips and not eips[i["id"]]["eip_po"] for i in buy)}
        ck = checks.get(g["inbound_no"])
        g["check"] = ck["summary"] if ck else None
        todo["yx"] = bool(ck) and not ck["summary"]["done"]
        todo["notice"] = bool(ck) and ck["summary"]["notice_need"] and not ck["summary"]["notice_done"]
        g["todo"] = [k for k, _l, _a in STAGES if todo[k]]
        g["next"] = next((a for k, _l, a in STAGES if todo[k]), "")
        g["codes"] = [i["factory_code"] for i in alive]
        out.append(g)
    out.sort(key=lambda g: (not g["diff_pending"], g["expected_date"], g["warehouse"], g["inbound_no"]))
    return out


@momo_bp.route("/api/momo/orders")
def api_orders():
    conn = get_conn()
    try:
        out = overview(conn, request.args)
        every = conn.execute("SELECT COUNT(DISTINCT inbound_no) AS n, MIN(expected_date) AS d1, MAX(expected_date) AS d2 FROM momo_orders").fetchone()
        whs = sorted({r["warehouse"] for r in _rows(conn.execute("SELECT DISTINCT warehouse FROM momo_orders")) if r["warehouse"]})
    finally:
        conn.close()
    stages = [{"key": k, "label": lab, "count": sum(1 for g in out if k in g["todo"])} for k, lab, _a in STAGES]
    stages.append({"key": "all", "label": "全部入庫單", "count": len(out)})
    for g in out:                                         # 清單只要幾個欄位；整筆品項右側面板另外拿
        g["items"] = [{"id": i["id"], "factory_code": i["factory_code"], "product_name": i["product_name"], "qty": i["qty"], "unit_name": i["unit_name"],
                       "line": i["line"], "removed": i["removed"]} for i in g["items"]]
    return jsonify({"orders": out, "stages": stages, "warehouses": whs,
                    "all_orders": {"count": every["n"] or 0, "first": every["d1"] or "", "last": every["d2"] or ""},
                    "summary": {"orders": len(out), "items": sum(g["item_count"] for g in out), "qty": sum(g["qty"] for g in out)}})


def _ids(conn, no):
    return [r["id"] for r in _rows(conn.execute("SELECT id FROM momo_orders WHERE inbound_no = ?", (no,)))]


@momo_bp.route("/api/momo/diffs/confirm", methods=["POST"])
def api_diffs_confirm():
    p = request.get_json(silent=True) or {}
    nos = [str(x).strip() for x in (p.get("inbound_nos") or []) if str(x).strip()]
    if not nos:
        return jsonify({"error": "請選擇要確認的入庫單。"}), 400
    operator, stamp, done = _operator(), now(), 0
    conn = get_conn()
    try:
        for no in nos:
            ids = _ids(conn, no)
            if not ids:
                continue
            marks = ",".join("?" * len(ids))
            n = conn.execute(f"SELECT COUNT(*) AS n FROM momo_diffs WHERE confirmed_at = '' AND order_id IN ({marks})", ids).fetchone()["n"]
            if not n:
                continue
            conn.execute(f"UPDATE momo_diffs SET confirmed_by = ?, confirmed_at = ? WHERE confirmed_at = '' AND order_id IN ({marks})", [operator, stamp] + ids)
            conn.execute(f"UPDATE momo_orders SET diff_pending = 0, updated_at = ? WHERE id IN ({marks})", [stamp] + ids)
            line = _row(conn.execute("SELECT line FROM momo_orders WHERE id = ?", (ids[0],)))["line"]
            _log(conn, line, no, "", "", "momo_diff_confirm", "確認 MOMO 差異", "", f"{n} 項差異", operator, "manual")
            done += 1
        conn.commit()
    finally:
        conn.close()
    return jsonify({"ok": True, "confirmed": done})


@momo_bp.route("/api/momo/ship_method", methods=["POST"])
def api_ship_method():
    """整張入庫單改出貨方式（原廠直出／竹運出貨）。已產出 EIP 的要先取消產出，免得同一批貨又下採又從竹運出。"""
    p = request.get_json(silent=True) or {}
    nos = [str(x).strip() for x in (p.get("inbound_nos") or []) if str(x).strip()]
    method = str(p.get("ship_method") or "").strip()
    if not nos:
        return jsonify({"error": "請選擇入庫單。"}), 400
    if method not in SHIP_METHODS:
        return jsonify({"error": f"出貨方式只能選：{'、'.join(SHIP_METHODS)}。"}), 400
    operator, stamp, changed, blocked = _operator(), now(), 0, []
    conn = get_conn()
    try:
        for no in nos:
            ids = _ids(conn, no)
            if not ids:
                continue
            marks = ",".join("?" * len(ids))
            if conn.execute(f"SELECT COUNT(*) AS n FROM momo_eip_items WHERE order_id IN ({marks})", ids).fetchone()["n"]:
                blocked.append(no)
                continue
            olds = sorted({r["ship_method"] for r in _rows(conn.execute(f"SELECT ship_method FROM momo_orders WHERE id IN ({marks})", ids))})
            if olds == [method]:
                continue
            conn.execute(f"UPDATE momo_orders SET ship_method = ?, updated_at = ? WHERE id IN ({marks})", [method, stamp] + ids)
            line = _row(conn.execute("SELECT line FROM momo_orders WHERE id = ?", (ids[0],)))["line"]
            _log(conn, line, no, "", "", "momo_ship_method", "出貨方式", "、".join(olds), method, operator, "manual")
            changed += 1
        conn.commit()
    finally:
        conn.close()
    if blocked and not changed:
        return jsonify({"error": f"{'、'.join(blocked)} 已產出 EIP 採購單，請先在右側面板取消產出，再改出貨方式。"}), 400
    return jsonify({"ok": True, "count": changed, "blocked": blocked})


@momo_bp.route("/api/momo/uploads")
def api_uploads():
    conn = get_conn()
    try:
        rows = _rows(conn.execute("""SELECT id, filenames, rows_total, new_count, changed_count, same_count, removed_count, errors_json, uploaded_by,
                                     uploaded_at, committed_at FROM momo_uploads WHERE status = 'ok' ORDER BY id DESC LIMIT 200"""))
        files = collections.defaultdict(list)
        if rows:
            marks = ",".join("?" * len(rows))
            for f in _rows(conn.execute(f"SELECT id, upload_id, filename FROM momo_upload_files WHERE upload_id IN ({marks}) ORDER BY id",
                                        [r["id"] for r in rows])):
                files[f["upload_id"]].append({"id": f["id"], "filename": f["filename"]})
    finally:
        conn.close()
    for r in rows:
        r["errors"] = len(json.loads(r.pop("errors_json") or "[]"))
        r["files"] = files.get(r["id"], [])
    return jsonify({"uploads": rows})


@momo_bp.route("/api/momo/upload_files/<int:fid>")
def api_upload_file(fid):
    conn = get_conn()
    try:
        f = _row(conn.execute("SELECT filename, content_b64 FROM momo_upload_files WHERE id = ?", (fid,)))
    finally:
        conn.close()
    if f is None or not f["content_b64"]:
        return jsonify({"error": "找不到這個原始檔。"}), 404
    return send_file(io.BytesIO(base64.b64decode(f["content_b64"])), as_attachment=True, download_name=f["filename"] or f"momo_{fid}.xls")


def portal_summary(conn, today_iso):
    """YFS 訂單系統首頁的 MOMO 卡片用。"""
    def one(sql, params=()):
        r = conn.execute(sql, params).fetchone()
        return (next(iter(dict(r).values())) if r is not None else 0) or 0
    import datetime as _dt
    d0 = _dt.date.fromisoformat(today_iso) - _dt.timedelta(days=29)
    rows = conn.execute("SELECT SUBSTR(created_at, 1, 10) AS d, COUNT(DISTINCT inbound_no) AS n FROM momo_orders WHERE created_at >= ? "
                        "GROUP BY SUBSTR(created_at, 1, 10)", (d0.isoformat(),)).fetchall()
    by_day = {r["d"]: int(r["n"]) for r in rows}
    return {"daily_new": [{"date": (d0 + _dt.timedelta(days=i)).isoformat(), "value": by_day.get((d0 + _dt.timedelta(days=i)).isoformat(), 0)}
                          for i in range(30)],
            "new_today": one("SELECT COUNT(DISTINCT inbound_no) AS n FROM momo_orders WHERE created_at >= ?", (today_iso,)),
            "pending_pos": one("SELECT COUNT(DISTINCT inbound_no) AS n FROM momo_orders WHERE expected_date >= ? AND removed = 0", (today_iso,)),
            "diff_pos": one("SELECT COUNT(DISTINCT inbound_no) AS n FROM momo_orders WHERE diff_pending = 1"),
            "last_import": one("SELECT MAX(committed_at) AS m FROM momo_uploads WHERE status = 'ok'") or ""}


__all__ = ["plan", "apply", "okey", "overview", "STAGES", "portal_summary"]

"""② 訂單明細的匯入：多檔合併解析 → 預覽（比對新增／有變／消失）→ 確認寫入。"""
from .common import *  # noqa: F401,F403 — 共用工具、Flask、db、openpyxl 都從這裡來


def _diff_import(conn, rows):
    products = _products_map(conn)
    new, updated, identical, missing = [], [], [], {}
    for r in rows:
        if r["barcode"] and r["barcode"] not in products:
            missing.setdefault(r["barcode"], {
                "barcode": r["barcode"], "sku_id": r["sku_id"], "yf_sku": r["yf_sku"],
                "brand": r["brand"], "product_name": r["product_name"],
                "box_size": r["box_size"], "line": r["line"] or ""})
        existing = _row(conn.execute(
            "SELECT * FROM mst_orders WHERE po_number = ? AND sku_id = ?",
            (r["po_number"], r["sku_id"])))
        if existing is None:
            new.append(r); continue
        changes = []
        mapping = {
            "line": r["line"] or "", "barcode": r["barcode"], "yf_sku": r["yf_sku"],
            "brand": r["brand"], "product_name": r["product_name"], "warehouse": r["warehouse"],
            "order_type": r["order_type"], "unit": r["unit"], "unit_price": r["unit_price"],
            "qty_coupang": r["qty_coupang"], "qty_file_ship": r["qty_file_ship"],
            "box_size_file": r["box_size"], "delivery_date_file": r["delivery_date"],
        }
        for field, value in mapping.items():
            if not _same(existing.get(field), value):
                changes.append({"field": field, "label": COUPANG_FIELDS[field],
                                "old": existing.get(field), "new": value})
        if existing["missing_in_file"]:
            changes.append({"field": "missing_in_file", "label": "品項重新出現",
                            "old": "檔案已無此品項", "new": "恢復"})
        if changes:
            updated.append({
                "existing_id": existing["id"], "changes": changes,
                "po_number": r["po_number"], "sku_id": r["sku_id"], "barcode": r["barcode"],
                "product_name": r["product_name"], "line": r["line"] or "",
                "qty_ship_overridden": existing["qty_ship_overridden"],
                "delivery_date_overridden": existing["delivery_date_overridden"],
            })
        else:
            identical.append(r)

    # 酷澎把品項數量下修到 0 時，匯出的表那一列直接消失。這張 PO 有在檔案裡、
    # 底下某個品項卻沒帶到 → 判定被拿掉：不刪列、出貨歸 0、標記。
    file_keys = {(r["po_number"], r["sku_id"]) for r in rows}
    removed = []
    for po in sorted({r["po_number"] for r in rows}):
        for ex in _rows(conn.execute("SELECT * FROM mst_orders WHERE po_number = ?", (po,))):
            if (po, ex["sku_id"]) in file_keys or ex["missing_in_file"]:
                continue
            removed.append({"id": ex["id"], "po_number": po, "sku_id": ex["sku_id"],
                            "barcode": ex["barcode"], "product_name": ex["product_name"],
                            "line": ex["line"], "qty_coupang": ex["qty_coupang"],
                            "qty_ship": ex["qty_ship"], "delivery_date": ex["delivery_date"]})
    return new, updated, identical, list(missing.values()), removed


def _parse_uploads(uploads):
    """把多份上傳檔各自 parse_workbook 再合併。回傳 (rows, warnings)；任一份壞掉
    就回傳 ((錯誤訊息, 400), None)，錯誤訊息點名是哪一份壞的，整批都不收。"""
    merged, warnings, seen = [], [], {}
    many = len(uploads) > 1
    for up in uploads:
        try:
            rows, warns = parse_workbook(io.BytesIO(up.read()), up.filename)
        except ImportError_ as exc:
            msg = f"「{up.filename}」：{exc}" if many else str(exc)
            if many:
                msg += "　（有檔案解析失敗，這次全部都沒匯入；修好再一起丟，或分開丟。）"
            return (msg, 400), None
        for w in warns:
            warnings.append(f"「{up.filename}」{w}" if many else w)
        dup = 0
        for r in rows:
            key = (r["po_number"], r["sku_id"])
            if key in seen:
                dup += 1
                merged = [x for x in merged if (x["po_number"], x["sku_id"]) != key]
            seen[key] = up.filename
            r["source_file"] = up.filename
            merged.append(r)
        if dup:
            warnings.append(f"「{up.filename}」有 {dup} 筆 PO／SKU 跟前面的檔案重複，以這份為準。")
    return merged, warnings


@master_bp.route("/api/master/import/preview", methods=["POST"])
def api_import_preview():
    """一次可以丟好幾份訂單彙總表（例如 7 月一份、8 月一份）。每份各自解析後
    合成一批預覽；跨檔案撞到同一個 (PO, SKU) 時以後面那份為準並提醒，跟單一
    檔案內重複的處理方式一致。"""
    uploads = [u for u in request.files.getlist("file") if u is not None and u.filename]
    if not uploads:
        return jsonify({"error": "沒有收到檔案。"}), 400
    rows, warnings = _parse_uploads(uploads)
    if isinstance(rows, tuple):          # (error message, status)
        return jsonify({"error": rows[0]}), rows[1]
    filename = "、".join(u.filename for u in uploads)
    cfg = _line_groups()
    conn = get_conn()
    try:
        new, updated, identical, missing, removed = _diff_import(conn, rows)
        payload = {"rows": rows, "missing_products": missing, "removed": removed,
                   "filename": filename}
        cur = conn.execute(
            """INSERT INTO mst_import_batches
               (filename, operator, rows_total, rows_new, rows_updated, rows_identical,
                committed, payload_json, created_at) VALUES (?,?,?,?,?,?,0,?,?)""",
            (filename, _operator(), len(rows), len(new), len(updated), len(identical),
             json.dumps(payload, ensure_ascii=False), now()))
        conn.commit()
        batch_id = cur.lastrowid
    finally:
        conn.close()

    by_raw = collections.Counter(r["line"] or "" for r in rows)
    by_group = collections.Counter(_group_of(r["line"], cfg) for r in rows)
    dates = sorted({r["delivery_date"] for r in rows if r["delivery_date"]})
    warnings = list(warnings)
    if by_raw.get(""):
        warnings.append(f"有 {by_raw['']} 筆沒有「線別」，歸到「{UNCLASSIFIED}」，畫面上會標紅；"
                        "請在訂單明細裡確認這幾筆。")
    return jsonify({
        "batch_id": batch_id, "filename": filename, "files": [u.filename for u in uploads],
        "rows_total": len(rows),
        "new_count": len(new), "updated_count": len(updated), "identical_count": len(identical),
        "lines_raw": dict(by_raw), "lines": dict(by_group), "dates": dates,
        "po_count": len({r["po_number"] for r in rows}), "warnings": warnings,
        "new_preview": new[:300], "updated": updated[:300],
        "missing_products": missing, "removed": removed, "removed_count": len(removed),
    })


def _commit_add_missing_products(conn, missing, operator, stamp, fname):
    """預覽時勾了「總表裡還沒有的商品也建進主檔」：用整合表的資料先建一筆、標 auto_created。"""
    added = 0
    for m in missing:
        if _row(conn.execute("SELECT id FROM mst_products WHERE barcode = ?", (m["barcode"],))):
            continue
        conn.execute(
            """INSERT INTO mst_products
               (barcode, sku_id, yf_sku, brand, product_name, box_size, note,
                auto_created, lines_seen, updated_by, updated_at)
               VALUES (?,?,?,?,?,?,'',1,?,?,?)""",
            (m["barcode"], m.get("sku_id") or "", m.get("yf_sku") or "", m.get("brand") or "",
             m.get("product_name") or "", m.get("box_size"), m.get("line") or "", operator, stamp))
        _log(conn, m.get("line"), "", m.get("sku_id"), m["barcode"], "product_new", "新增主檔",
             "", m.get("box_size"), operator, "import", fname)
        added += 1
    return added


def _commit_mark_removed(conn, removed, batch_id, operator, stamp):
    """PO 還在檔案裡、底下某品項不見了：列保留、出貨歸 0、標 missing_in_file（再出現會自動恢復）。"""
    n = 0
    for rm in removed:
        ex = _row(conn.execute("SELECT * FROM mst_orders WHERE id = ?", (rm["id"],)))
        if ex is None or ex["missing_in_file"]:
            continue
        conn.execute(
            """UPDATE mst_orders SET qty_ship = 0, missing_in_file = 1, last_seen_at = ?,
               updated_at = ?, last_batch_id = ?, last_batch_changes = ?, version = version + 1
               WHERE id = ?""",
            (stamp, stamp, batch_id, f"檔案已無此品項，出貨 {_short(ex['qty_ship'])}→0", ex["id"]))
        _log(conn, ex["line"], ex["po_number"], ex["sku_id"], ex["barcode"], "qty_ship", "出貨數量",
             ex["qty_ship"], 0, operator, "import", "這次的整合表裡這張 PO 已沒有這個品項，出貨數量歸 0")
        n += 1
    return n


def _commit_insert_row(conn, r, batch_id, stamp):
    """第一次看到這個 (PO, SKU)：整列寫進去，first_batch_id 記住是哪一批帶進來的。"""
    ship = r["qty_file_ship"] if r["qty_file_ship"] is not None else r["qty_coupang"]
    conn.execute(
        """INSERT INTO mst_orders
           (po_number, sku_id, line, barcode, yf_sku, brand, product_name, warehouse,
            order_type, unit, unit_price, qty_coupang, qty_file_ship, qty_ship,
            box_size_file, delivery_date_file, delivery_date, remarks_file, remarks,
            source_file, first_seen_at, last_seen_at, updated_at, version,
            first_batch_id, last_batch_id, last_batch_changes, address, quote_note)
           VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,1,?,?,'新增',?,?)""",
        (r["po_number"], r["sku_id"], r["line"] or "", r["barcode"], r["yf_sku"], r["brand"],
         r["product_name"], r["warehouse"], r["order_type"], r["unit"], r["unit_price"],
         r["qty_coupang"], r["qty_file_ship"], ship, r["box_size"], r["delivery_date"],
         r["delivery_date"], r["remarks_file"], "", r["source_file"], stamp, stamp, stamp,
         batch_id, batch_id, r.get("address") or "", r.get("quote_note") or ""))


def _commit_update_row(conn, existing, r, batch_id, operator, stamp, fname):
    """既有的 (PO, SKU)：整合表帶來的欄位照更新；出貨數量／交貨日若人改過就不動，只記一筆
    「整合表變了但畫面不動」；每個變動用人看得懂的字記進 last_batch_changes。回傳 True 表示有變。"""
    line = r["line"] or ""
    ship = r["qty_file_ship"] if r["qty_file_ship"] is not None else r["qty_coupang"]
    sets, vals, changed, notes = [], [], False, []
    mapping = {
        "line": line, "barcode": r["barcode"], "yf_sku": r["yf_sku"], "brand": r["brand"],
        "product_name": r["product_name"], "warehouse": r["warehouse"],
        "order_type": r["order_type"], "unit": r["unit"], "unit_price": r["unit_price"],
        "qty_coupang": r["qty_coupang"], "qty_file_ship": r["qty_file_ship"],
        "box_size_file": r["box_size"], "delivery_date_file": r["delivery_date"],
        "remarks_file": r["remarks_file"],
    }
    for field, value in mapping.items():
        if not _same(existing.get(field), value):
            sets.append(f"{field} = ?"); vals.append(value); changed = True
            if field in COUPANG_FIELDS:
                _log(conn, line, r["po_number"], r["sku_id"], r["barcode"], field,
                     COUPANG_FIELDS[field], existing.get(field), value, operator, "import", fname)
                if field not in ("qty_file_ship", "delivery_date_file"):   # 這兩個下面用人看得懂的名字記
                    notes.append(f"{COUPANG_FIELDS[field]} {_short(existing.get(field))}→{_short(value)}")
    if not existing["qty_ship_overridden"] and not _same(existing["qty_ship"], ship):
        sets.append("qty_ship = ?"); vals.append(ship); changed = True
        notes.append(f"出貨數量 {_short(existing['qty_ship'])}→{_short(ship)}")
        _log(conn, line, r["po_number"], r["sku_id"], r["barcode"], "qty_ship", "出貨數量",
             existing["qty_ship"], ship, operator, "import", "隨整合表更新")
    elif existing["qty_ship_overridden"] and not _same(existing["qty_file_ship"], r["qty_file_ship"]):
        notes.append(f"整合表出貨 {_short(existing['qty_file_ship'])}→{_short(r['qty_file_ship'])}（人工調整過，畫面不動）")
    if not existing["delivery_date_overridden"] and not _same(existing["delivery_date"], r["delivery_date"]):
        sets.append("delivery_date = ?"); vals.append(r["delivery_date"]); changed = True
        notes.append(f"交貨日 {_short(existing['delivery_date'])}→{_short(r['delivery_date'])}")
        _log(conn, line, r["po_number"], r["sku_id"], r["barcode"], "delivery_date", "交貨日",
             existing["delivery_date"], r["delivery_date"], operator, "import", "隨整合表更新")
    elif existing["delivery_date_overridden"] and not _same(existing["delivery_date_file"], r["delivery_date"]):
        notes.append(f"整合表交貨日 {_short(existing['delivery_date_file'])}→{_short(r['delivery_date'])}（人工改期過，畫面不動）")
    if existing["missing_in_file"]:
        sets.append("missing_in_file = 0"); changed = True
        notes.append("品項重新出現")
        if not existing["qty_ship_overridden"] and "qty_ship = ?" not in sets:
            sets.append("qty_ship = ?"); vals.append(ship)
        _log(conn, line, r["po_number"], r["sku_id"], r["barcode"], "missing_in_file",
             "品項重新出現", 1, 0, operator, "import", "整合表裡又有這個品項了")
    sets.append("last_seen_at = ?"); vals.append(stamp)
    sets.append("source_file = ?"); vals.append(r["source_file"])
    # 地址、報價備註：瑪氏拆單要用，靜靜跟著整合表更新，不算「有變」、不記歷程（跟出貨無關）
    for field in ("address", "quote_note"):
        if field in r and not _same(existing.get(field) or "", r.get(field) or ""):
            sets.append(f"{field} = ?"); vals.append(r.get(field) or "")
    if changed:
        sets.append("updated_at = ?"); vals.append(stamp)
        sets.append("last_batch_id = ?"); vals.append(batch_id)
        sets.append("last_batch_changes = ?"); vals.append("；".join(notes) or "有變")
        sets.append("version = version + 1")
    conn.execute(f"UPDATE mst_orders SET {', '.join(sets)} WHERE id = ?", vals + [existing["id"]])
    return changed


def _learn_lines_seen(conn, rows):
    """主檔學線別：這個國條出現在哪些原始線別的訂單裡（空白線別不算）。"""
    seen_by_bc = collections.defaultdict(set)
    for r in rows:
        if r["barcode"]:
            seen_by_bc[r["barcode"]].add(r["line"] or "")
    for bc, lines in seen_by_bc.items():
        p = _row(conn.execute("SELECT id, lines_seen FROM mst_products WHERE barcode = ?", (bc,)))
        if p is None:
            continue
        merged = sorted(set(_split_lines(p["lines_seen"])) | {l for l in lines if l})
        if ",".join(merged) != (p["lines_seen"] or ""):
            conn.execute("UPDATE mst_products SET lines_seen = ? WHERE id = ?", (",".join(merged), p["id"]))


@master_bp.route("/api/master/import/commit", methods=["POST"])
def api_import_commit():
    """把預覽那批真的寫進去。順序：補主檔 → 標記檔案裡消失的品項 → 每列新增或更新 →
    主檔學線別 → 批次標記已確認。整批在一個交易裡，中途炸掉全部回滾。"""
    payload = request.get_json(silent=True) or {}
    batch_id = norm_int(payload.get("batch_id"))
    add_missing = payload.get("add_missing_products", True)
    operator = _operator()
    if batch_id is None:
        return jsonify({"error": "缺少 batch_id。"}), 400
    conn = get_conn()
    try:
        batch = _row(conn.execute("SELECT * FROM mst_import_batches WHERE id = ?", (batch_id,)))
        if batch is None:
            return jsonify({"error": "找不到這次的預覽，請重新上傳。"}), 404
        if batch["committed"]:
            return jsonify({"error": "這批已經匯入過了，不能重複確認。"}), 409
        data = json.loads(batch["payload_json"] or "{}")
        rows = data.get("rows", [])
        fname = data.get("filename", "")
        stamp = now()

        products_added = _commit_add_missing_products(conn, data.get("missing_products", []), operator, stamp, fname) if add_missing else 0
        removed_n = _commit_mark_removed(conn, data.get("removed", []), batch_id, operator, stamp)
        inserted = updated_n = identical_n = 0
        for r in rows:
            existing = _row(conn.execute(
                "SELECT * FROM mst_orders WHERE po_number = ? AND sku_id = ?", (r["po_number"], r["sku_id"])))
            if existing is None:
                _commit_insert_row(conn, r, batch_id, stamp); inserted += 1
            elif _commit_update_row(conn, existing, r, batch_id, operator, stamp, fname):
                updated_n += 1
            else:
                identical_n += 1
        _learn_lines_seen(conn, rows)

        conn.execute(
            """UPDATE mst_import_batches SET committed = 1, committed_at = ?,
               rows_new = ?, rows_updated = ?, rows_identical = ? WHERE id = ?""",
            (stamp, inserted, updated_n, identical_n, batch_id))
        conn.commit()
    except Exception:
        conn.rollback(); raise
    finally:
        conn.close()
    return jsonify({"ok": True, "inserted": inserted, "updated": updated_n, "identical": identical_n,
                    "products_added": products_added, "removed": removed_n})

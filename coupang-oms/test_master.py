"""業績總表自動化的端到端驢證：拿真實的訂單彙總表（7 月單線別、9 月多線別）與
總表範例跑完整流程。表全部是 mst_ 開頭，不碰訂單管理與採購表轉換的資料。

執行：python test_master.py
"""
import io
import os
import sys
import tempfile
from urllib.parse import unquote

import openpyxl

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SAMPLES = os.path.join(BASE_DIR, "samples", "master")
JUL_XLSX = os.path.join(SAMPLES, "訂單彙總表範例.xlsx")
SEP_XLSX = os.path.join(SAMPLES, "訂單彙總表_9月多線別範例.xlsx")
MASTER_XLSX = os.path.join(SAMPLES, "總表範例.xlsx")

_tmp = tempfile.mkdtemp(prefix="oms_master_test_")
sys.path.insert(0, BASE_DIR)

import db  # noqa: E402
db.DATA_DIR = os.path.join(_tmp, "資料與設定")
db.DB_PATH = os.path.join(db.DATA_DIR, "test.db")
db.BACKUP_DIR = os.path.join(db.DATA_DIR, "backups")

import app as app_module  # noqa: E402

PASS, FAIL = [], []


def check(name, condition, detail=""):
    (PASS if condition else FAIL).append(name)
    print(f"  {'✓' if condition else '✗'} {name}" + (f"  — {detail}" if detail else ""))


def sign_in(client, username="小真", password="changeme123"):
    res = client.post("/login", data={"username": username, "password": password})
    assert res.status_code == 302


def upload(client, url, path, **form):
    with open(path, "rb") as fh:
        data = fh.read()
    payload = dict(form); payload["file"] = (io.BytesIO(data), os.path.basename(path))
    return client.post(url, data=payload, content_type="multipart/form-data")


def jput(client, url, body):
    return client.put(url, json=body)


def orders(client, **params):
    from urllib.parse import urlencode
    return client.get("/api/master/orders?" + urlencode(params)).get_json()


def main():
    db.init_db()
    app_module.app.config["TESTING"] = True
    client = app_module.app.test_client()
    sign_in(client)

    print("\n【1】頁面與入口")
    res = client.get("/master"); html = res.get_data(as_text=True)
    check("頁面打得開", res.status_code == 200)
    check("分頁順序照流程：① 商品主檔 → ② 訂單明細 → ③ 總表", html.index("① 商品主檔") < html.index("② 訂單明細") < html.index("③ 總表"))
    check("頂端沒有全域的線別選單（線別是篩選，不是模式）", 'id="sel-line"' not in html)
    check("OMS 首頁有「業績總表自動化」按鈕", "業績總表自動化" in client.get("/").get_data(as_text=True))
    meta = client.get("/api/master/lines").get_json()
    check("還沒有資料時線別群組是空的（不寫死任何名字）", meta["groups"] == [], str(meta["groups"]))

    print("\n【2】① 先匯總表鋪主檔（不問線別）")
    res = upload(client, "/api/master/products/import", MASTER_XLSX)
    pi = res.get_json()
    check("總表匯入成功、抓到品類／COGS／箱入數／Note 等欄", res.status_code == 200 and {"barcode", "box_size", "category", "cost_price", "note"} <= set(pi["columns_found"]), str(pi))
    prods = client.get("/api/master/products").get_json()["products"]
    p = next((x for x in prods if x["barcode"] == "4987176340894"), None)
    check("ARIEL 4987176340894：箱入數 6、品類 Fabric、COGS 202、PG code", p and p["box_size"] == 6 and p["category"] == "Fabric" and p["cost_price"] == 202 and p["pgcode"] == "80864486", str(p and (p["box_size"], p["category"], p["cost_price"])))
    check("主檔的線別欄空的（還沒出現在任何訂單）", all(x["line_groups"] == [] for x in prods))

    print("\n【3】② 匯 9 月多線別訂單彙總表：預覽")
    res = upload(client, "/api/master/import/preview", SEP_XLSX)
    pv = res.get_json()
    check("讀到 861 個品項、258 張 PO", pv["rows_total"] == 861 and pv["po_count"] == 258, f"{pv['rows_total']} / {pv['po_count']}")
    check("原始線別四種 + 空白", set(pv["lines_raw"]) == {"瑪氏", "寶僑", "CPG-潔品", "CPG-紙品", ""}, str(pv["lines_raw"]))
    check("群組化：CPG-潔品＋CPG-紙品 → 紙潔（253 筆），空白 → 未分類（4 筆）",
          pv["lines"].get("紙潔") == 253 and pv["lines"].get("未分類") == 4 and pv["lines"].get("寶僑") == 315, str(pv["lines"]))
    check("沒有線別的 4 筆有提醒", any("沒有「線別」" in w for w in pv["warnings"]))
    check("主檔沒有的國條列為待建（總表範例只有 80 列）", len(pv["missing_products"]) > 100, str(len(pv["missing_products"])))
    res = client.post("/api/master/import/commit", json={"batch_id": pv["batch_id"]})
    cm = res.get_json()
    check("寫入 861 筆", cm["inserted"] == 861, str(cm))
    meta = client.get("/api/master/lines").get_json()
    check("線別群組從資料長出來：寶僑／瑪氏／紙潔／未分類", set(meta["groups"]) == {"寶僑", "瑪氏", "紙潔", "未分類"}, str(meta["groups"]))
    check("月份清單有 2026-09", "2026-09" in meta["months"])

    print("\n【4】訂單明細：篩選面與線別群組")
    od = orders(client, month="2026-09")
    check("9 月 861 筆全帶出", od["count"] == 861)
    lines = {x["line"]: x["rows"] for x in od["facets"]["lines"]}
    check("線別篩選面 = 群組（紙潔 253）", lines.get("紙潔") == 253 and lines.get("未分類") == 4, str(lines))
    check("日期篩選面 23 天", len([d for d in od["facets"]["dates"] if d["date"]]) == 23, str(len(od["facets"]["dates"])))
    check("倉別篩選面 11 個", len(od["facets"]["warehouses"]) == 11, str(len(od["facets"]["warehouses"])))
    cross = [x for x in od["facets"]["pos"] if len(x["lines"]) > 1]
    check("同一 PO 跨線別的 PO 在篩選面只出現一次、標兩個線別", len(cross) == 0 and any(
        {r["line"] for r in od["rows"] if r["po_number"] == "13000000452582"} == {"CPG-潔品", "CPG-紙品"} for _ in [0]),
          "跨線別 PO 應被群組成紙潔一個線別")
    cross_po = "13000000452582"
    check("PO 13000000452582 兩個原始線別都歸紙潔", {r["line_group"] for r in od["rows"] if r["po_number"] == cross_po} == {"紙潔"})
    od_pg = orders(client, month="2026-09", lines="寶僑")
    check("勾寶僑只回寶僑 315 筆", od_pg["count"] == 315 and all(r["line_group"] == "寶僑" for r in od_pg["rows"]))
    od_cpg = orders(client, month="2026-09", lines="紙潔")
    check("勾紙潔回 253 筆（潔品＋紙品）", od_cpg["count"] == 253 and {r["line"] for r in od_cpg["rows"]} == {"CPG-潔品", "CPG-紙品"})
    check("勾紙潔＋倉別交叉篩", orders(client, month="2026-09", lines="紙潔", warehouses="TAO1")["count"] < 253)
    unc = orders(client, month="2026-09", lines="未分類")
    check("未分類 4 筆帶出、標沒線別", unc["count"] == 4 and all(r["line"] == "" for r in unc["rows"]))
    check("換到 7 月是空的（資料照交貨日分月）", orders(client, month="2026-07")["count"] == 0)
    BC = "4987176340863"   # 總表範例有（Fabric、COGS 202、Note 新品）、9 月寶僑訂單也有
    prods = client.get("/api/master/products").get_json()["products"]
    p = next(x for x in prods if x["barcode"] == BC)
    check("主檔從訂單學到線別：ARIEL 4987176340863 屬於寶僑", p["line_groups"] == ["寶僑"], str(p["line_groups"]))
    p_never = next(x for x in prods if x["barcode"] == "4987176340894")
    check("沒出現在任何訂單的商品，線別空著（尚未出現在訂單）", p_never["line_groups"] == [])
    check("主檔可用線別群組篩", all("紙潔" in x["line_groups"] for x in client.get("/api/master/products?line=紙潔").get_json()["products"]))

    print("\n【5】就地編輯與防互蓋")
    row = next(r for r in od_pg["rows"] if r["cases"] is not None and r["box_size"] and r["box_size"] > 1)
    oid, ver = row["id"], row["version"]
    res = jput(client, f"/api/master/orders/{oid}", {"version": ver, "qty_ship": row["box_size"] * 3})
    check("改出貨數量 → 箱數 3、立旗標", res.status_code == 200 and res.get_json()["row"]["cases"] == 3 and res.get_json()["row"]["qty_ship_overridden"] == 1, str(res.get_json())[:100])
    check("舊版本再存被擋（409）", jput(client, f"/api/master/orders/{oid}", {"version": ver, "qty_ship": 1}).status_code == 409)
    v2 = res.get_json()["row"]["version"]
    res = jput(client, f"/api/master/orders/{oid}", {"version": v2, "remarks": "打單缺貨"})
    check("備註存成 OP 備註", res.get_json()["row"]["remarks"] == "打單缺貨" and res.get_json()["row"]["remarks_overridden"] == 1)
    check("匯出備註 = Note＋OP 備註（若 Note 為空就只有 OP）", res.get_json()["row"]["export_note"].endswith("打單缺貨"))
    v3 = res.get_json()["row"]["version"]
    res = jput(client, f"/api/master/orders/{oid}", {"version": v3, "remarks": ""})
    check("備註清空 → 旗標歸零", res.get_json()["row"]["remarks_overridden"] == 0)

    print("\n【6】PO 改期：跨線別整張一起搬、批次、視窗儲存")
    before = orders(client, month="2026-09", pos=cross_po)
    n_rows = before["count"]; old_date = before["rows"][0]["delivery_date"]
    res = jput(client, "/api/master/pos/date", {"po_numbers": [cross_po], "delivery_date": "2026-09-27", "expected": {cross_po: old_date}})
    check(f"跨線別 PO 改期：潔品＋紙品 {n_rows} 個品項一起搬", res.status_code == 200 and res.get_json()["moved"] == n_rows, str(res.get_json()))
    after = orders(client, month="2026-09", pos=cross_po)
    check("全部在 9/27", all(r["delivery_date"] == "2026-09-27" for r in after["rows"]))
    check("用舊日期當依據再改被擋（409）", jput(client, "/api/master/pos/date", {"po_numbers": [cross_po], "delivery_date": "2026-09-28", "expected": {cross_po: old_date}}).status_code == 409)
    two = [x["po_number"] for x in od_pg["facets"]["pos"][:2]]
    exp = {x["po_number"]: x["date"] for x in od_pg["facets"]["pos"][:2]}
    res = jput(client, "/api/master/pos/date", {"po_numbers": two, "delivery_date": "2026-09-30", "expected": exp})
    check("批次改兩張 PO", res.status_code == 200 and res.get_json()["po_count"] == 2 and res.get_json()["moved"] > 0, str(res.get_json()))
    check("批次改完都在 9/30", all(r["delivery_date"] == "2026-09-30" for r in orders(client, month="2026-09", pos=",".join(two))["rows"]))
    res = jput(client, "/api/master/pos/date", {"po_numbers": two, "reset": True})
    check("批次恢復整合表日期", res.status_code == 200 and all(r["delivery_date_overridden"] == 0 for r in orders(client, month="2026-09", pos=",".join(two))["rows"]))
    detail = client.get(f"/api/master/pos/{cross_po}").get_json()
    check("PO 視窗：兩個原始線別、歸紙潔、有歷程", detail["lines"] == ["紙潔"] and set(detail["lines_raw"]) == {"CPG-潔品", "CPG-紙品"} and any(l["field"] == "delivery_date" for l in detail["logs"]))
    items = [{"id": r["id"], "version": r["version"], "qty_ship": 0, "remarks": "視窗改的"} for r in detail["rows"][:2]]
    res = jput(client, f"/api/master/pos/{cross_po}/save", {"items": items, "delivery_date": "2026-09-25", "expected_date": "2026-09-27"})
    check("PO 視窗一次儲存：改期＋兩個品項", res.status_code == 200 and res.get_json()["changed"] >= n_rows + 4, str(res.get_json()))
    d2 = client.get(f"/api/master/pos/{cross_po}").get_json()
    check("視窗儲存後：日期 9/25、兩筆出貨 0 且備註「視窗改的」", d2["delivery_date"] == "2026-09-25" and sum(1 for r in d2["rows"] if r["qty_ship"] == 0 and r["remarks"] == "視窗改的") == 2)
    res = jput(client, f"/api/master/pos/{cross_po}/save", {"items": [{"id": detail["rows"][0]["id"], "version": detail["rows"][0]["version"], "qty_ship": 5}]})
    check("視窗用舊版本存被擋、整批不寫", res.status_code == 409)

    print("\n【7】再匯入：人改過的不覆蓋；消失的品項留列歸 0")
    res = upload(client, "/api/master/import/preview", SEP_XLSX); pv2 = res.get_json()
    check("同一份再匯：0 新增", pv2["new_count"] == 0)
    client.post("/api/master/import/commit", json={"batch_id": pv2["batch_id"]})
    r_now = next(r for r in orders(client, month="2026-09", pos=row["po_number"])["rows"] if r["id"] == oid)
    check("人工改的出貨數量沒被蓋掉", r_now["qty_ship"] == row["box_size"] * 3 and r_now["qty_ship_overridden"] == 1)
    check("人工改的日期沒被蓋掉（9/25）", client.get(f"/api/master/pos/{cross_po}").get_json()["delivery_date"] == "2026-09-25")
    wb = openpyxl.load_workbook(SEP_XLSX); ws = wb.active
    hdr = [c.value for c in ws[1]]; ci = {h: i + 1 for i, h in enumerate(hdr)}
    gone = None
    for r in range(ws.max_row, 1, -1):
        if str(ws.cell(r, ci["PO單號"]).value) == cross_po:
            gone = str(ws.cell(r, ci["SKU ID"]).value); ws.delete_rows(r); break
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    res = client.post("/api/master/import/preview", data={"file": (buf, "少一筆.xlsx")}, content_type="multipart/form-data"); pv3 = res.get_json()
    check("PO 在檔案裡、品項消失 → 預覽抓到 1 筆", pv3["removed_count"] == 1 and pv3["removed"][0]["sku_id"] == gone)
    client.post("/api/master/import/commit", json={"batch_id": pv3["batch_id"]})
    g = next(r for r in client.get(f"/api/master/pos/{cross_po}").get_json()["rows"] if r["sku_id"] == gone)
    check("那筆留著、出貨 0、標記檔案已無此品項", g["qty_ship"] == 0 and g["missing_in_file"] == 1)
    res = upload(client, "/api/master/import/preview", SEP_XLSX); pv4 = res.get_json()
    check("完整檔再匯：預覽標「品項重新出現」", any(c["field"] == "missing_in_file" for u in pv4["updated"] for c in u["changes"]))
    client.post("/api/master/import/commit", json={"batch_id": pv4["batch_id"]})
    g2 = next(r for r in client.get(f"/api/master/pos/{cross_po}").get_json()["rows"] if r["sku_id"] == gone)
    check("重新出現：標記解除", g2["missing_in_file"] == 0)

    print("\n【8】7 月單線別檔也能一起進（資料照月份分）")
    res = upload(client, "/api/master/import/preview", JUL_XLSX); pvj = res.get_json()
    check("7 月檔 53 筆全部新增", pvj["new_count"] == 53)
    client.post("/api/master/import/commit", json={"batch_id": pvj["batch_id"]})
    check("7 月看得到 53 筆、9 月還是 861", orders(client, month="2026-07")["count"] == 53 and orders(client, month="2026-09")["count"] == 861)
    row7 = next(r for r in orders(client, month="2026-07")["rows"] if r["barcode"] == "6903148182406")
    check("62 袋 ÷ 24 = 2.5833 箱，不湊整", abs(row7["cases"] - 2.5833) < 1e-3)

    print("\n【9】線別群組設定可改")
    res = jput(client, "/api/master/line_groups", {"rules": [{"prefix": "CPG", "group": "紙潔"}], "map": {"瑪氏": "MARS"}})
    check("把瑪氏指定成 MARS", res.status_code == 200)
    check("篩選面跟著變", "MARS" in {x["line"] for x in orders(client, month="2026-09")["facets"]["lines"]})
    jput(client, "/api/master/line_groups", {"rules": [{"prefix": "CPG", "group": "紙潔"}], "map": {}})
    check("改回來", "瑪氏" in {x["line"] for x in orders(client, month="2026-09")["facets"]["lines"]})

    print("\n【10】總表（現算）")
    s = client.get("/api/master/summary?line=紙潔&month=2026-09").get_json()
    check("紙潔總表：有日期欄、有商品", len(s["dates"]) > 5 and len(s["rows"]) > 50, f"{len(s['dates'])} 天 / {len(s['rows'])} 商品")
    check("紙潔的商品標出原始線別（潔品／紙品）", any(set(r["lines_raw"]) & {"CPG-潔品", "CPG-紙品"} for r in s["rows"]))
    check("整月合計 = 每日合計相加", abs(s["month_total"] - sum(s["totals_by_date"].values())) < 1e-6)
    s_pg = client.get("/api/master/summary?line=寶僑&month=2026-09").get_json()
    r_a = next(x for x in s_pg["rows"] if x["barcode"] == BC)
    check("寶僑總表帶品類 Fabric、Note 新品", r_a["category"] == "Fabric" and r_a["note"] == "新品", str((r_a["category"], r_a["note"])))
    check("總表只列該線別：紙潔的表裡沒有寶僑 ARIEL", all(x["barcode"] != BC for x in s["rows"]))
    check("沒出現在訂單的商品不進總表（系統不知道它是誰的）", all(x["barcode"] != "4987176340894" for x in s_pg["rows"]))

    print("\n【11】匯出")
    res = client.get("/api/master/export?line=寶僑&month=2026-09")
    wbx = openpyxl.load_workbook(io.BytesIO(res.data))
    check("匯出總表：三欄品類在", wbx.sheetnames[0] == "總表" and "品類" in [c.value for c in wbx["總表"][1]])
    cpg_date = next(d["date"] for d in orders(client, month="2026-09", lines="紙潔")["facets"]["dates"] if d["date"] and d["rows"] >= 5)
    cpg_tab = f"{int(cpg_date[5:7]):02d}{int(cpg_date[8:10]):02d}交貨"; cpg_lbl = f"{int(cpg_date[5:7])}/{int(cpg_date[8:10])}交貨"
    res = client.get(f"/api/master/export/daily?month=2026-09&lines=紙潔&dates={cpg_date}")
    wbd = openpyxl.load_workbook(io.BytesIO(res.data))
    check(f"報價檔格式：篩紙潔 + {cpg_date} → 只有 {cpg_tab} 一個分頁", wbd.sheetnames == [cpg_tab], str(wbd.sheetnames))
    cd = unquote(res.headers.get("Content-Disposition", ""))
    check("檔名帶線別與日期", "紙潔" in cd and cpg_tab in cd, cd[:100])
    wsd = wbd[cpg_tab]; hd = [c.value for c in wsd[1]]
    check("A～R 欄位順序照報價檔", hd[0] == "SKU ID" and hd[3] == "品類" and hd[10] == "單價(含稅)" and hd[14] == "備註" and hd[17] == "交貨日")
    pos_cells = [wsd.cell(r, 18).value for r in range(2, wsd.max_row + 1) if wsd.cell(r, 18).value]
    check("每張 PO 第一列 R 欄寫 PO_日期(倉別)", all(f"_{cpg_lbl}(" in v for v in pos_cells) and len(pos_cells) >= 1, str(pos_cells[:2]))
    yellow = [r for r in range(2, wsd.max_row + 1) if wsd.cell(r, 1).value is None and wsd.cell(r, 1).fill.fgColor.rgb in ("00FFFF00", "FFFFFF00")]
    check("PO 之間黃色空白列", len(yellow) == len(pos_cells) - 1, f"{len(yellow)} / {len(pos_cells)}")
    check("紙潔的分頁裡只有 CPG 的品項（沒有寶僑）", all("Pampers" not in str(wsd.cell(r, 6).value or "") for r in range(2, wsd.max_row + 1)))
    pg_date = next(r["delivery_date"] for r in orders(client, month="2026-09", lines="寶僑")["rows"] if r["barcode"] == BC)
    res = client.get(f"/api/master/export/daily?month=2026-09&lines=寶僑&dates={pg_date}")
    wsp = openpyxl.load_workbook(io.BytesIO(res.data)).worksheets[0]
    row_a = next((r for r in wsp.iter_rows(min_row=2, values_only=True) if str(r[2]) == BC), None)
    check(f"寶僑 {pg_date} 報價檔：ARIEL 品類 Fabric、單價 = COGS 202、備註 = Note 新品", row_a and row_a[3] == "Fabric" and row_a[10] == 202 and row_a[14] == "新品", str(row_a and (row_a[3], row_a[10], row_a[14])))
    check("整合表自帶的 MPO_ 備註沒進報價檔", all("MPO_" not in str(r[14] or "") for r in wsp.iter_rows(min_row=2, values_only=True)))
    res = client.get("/api/master/export/daily?month=2026-09&warehouses=NOPE")
    check("篡到沒資料的檔案有說明", openpyxl.load_workbook(io.BytesIO(res.data)).sheetnames == ["無資料"])

    print("\n【12】歷程")
    logs = client.get(f"/api/master/logs?po={cross_po}").get_json()["logs"]
    check("PO 的歷程含改期、出貨數量、備註", {l["field"] for l in logs} >= {"delivery_date", "qty_ship", "remarks"})
    check("歷程可用關鍵字查", len(client.get("/api/master/logs?q=打單缺貨").get_json()["logs"]) >= 1)

    if db.IS_POSTGRES:
        print("\n【13】v1 舊資料庫升級（SQLite 專用，PostgreSQL 模式略過）")
        print(f"\n通過 {len(PASS)} 項，失敗 {len(FAIL)} 項")
        if FAIL:
            print("失敗項目："); [print("  -", f) for f in FAIL]; sys.exit(1)
        return
    print("\n【13】v1 舊資料庫升級")
    import sqlite3
    old = os.path.join(_tmp, "v1.db")
    c = sqlite3.connect(old)
    c.executescript("""
      CREATE TABLE mst_products (id INTEGER PRIMARY KEY AUTOINCREMENT, line TEXT NOT NULL, barcode TEXT NOT NULL,
        sku_id TEXT DEFAULT '', yf_sku TEXT DEFAULT '', brand TEXT DEFAULT '', product_name TEXT DEFAULT '',
        box_size INTEGER, note TEXT DEFAULT '', updated_by TEXT DEFAULT '', updated_at TEXT DEFAULT '', UNIQUE(line, barcode));
      INSERT INTO mst_products (line, barcode, brand, product_name, box_size, note) VALUES
        ('寶僑','111','ARIEL','A',6,'新品'), ('紙潔','111','ARIEL','A',6,''), ('寶僑','222','P','B',1,'由匯入自動建立，箱入數請核對');
      CREATE TABLE mst_orders (id INTEGER PRIMARY KEY AUTOINCREMENT, line TEXT NOT NULL, po_number TEXT NOT NULL, sku_id TEXT NOT NULL,
        barcode TEXT DEFAULT '', yf_sku TEXT DEFAULT '', brand TEXT DEFAULT '', product_name TEXT DEFAULT '', warehouse TEXT DEFAULT '',
        order_type TEXT DEFAULT '', unit TEXT DEFAULT '', unit_price REAL, qty_coupang INTEGER, qty_file_ship INTEGER, qty_ship INTEGER,
        qty_ship_overridden INTEGER NOT NULL DEFAULT 0, box_size_file INTEGER, delivery_date_file TEXT DEFAULT '', delivery_date TEXT DEFAULT '',
        delivery_date_overridden INTEGER NOT NULL DEFAULT 0, remarks_file TEXT DEFAULT '', remarks TEXT DEFAULT '', remarks_overridden INTEGER NOT NULL DEFAULT 0,
        source_file TEXT DEFAULT '', first_seen_at TEXT DEFAULT '', last_seen_at TEXT DEFAULT '', updated_at TEXT DEFAULT '', version INTEGER NOT NULL DEFAULT 1,
        UNIQUE(line, po_number, sku_id));
      INSERT INTO mst_orders (line, po_number, sku_id, barcode, qty_coupang, qty_ship, delivery_date) VALUES ('寶僑','P1','S1','111',10,10,'2026-07-01'), ('CPG-紙品','P2','S2','333',5,5,'2026-07-02');
    """); c.commit(); c.close()
    saved = db.DB_PATH; db.DB_PATH = old
    try:
        db.init_db()
        c = sqlite3.connect(old); c.row_factory = sqlite3.Row
        prods = [dict(r) for r in c.execute("SELECT * FROM mst_products ORDER BY barcode")]
        check("主檔合併成國條唯一：2 筆（111 合併）", [p["barcode"] for p in prods] == ["111", "222"], str([p["barcode"] for p in prods]))
        p111 = next(p for p in prods if p["barcode"] == "111")
        check("合併後 lines_seen 記兩個線別", set(p111["lines_seen"].split(",")) == {"寶僑", "紙潔"}, p111["lines_seen"])
        p222 = next(p for p in prods if p["barcode"] == "222")
        check("舊的「自動建立」Note 搬成旗標、Note 清空", p222["note"] == "" and p222["auto_created"] == 1)
        ords = [dict(r) for r in c.execute("SELECT * FROM mst_orders ORDER BY po_number")]
        check("訂單搬到新表、線別保留原始值", len(ords) == 2 and ords[1]["line"] == "CPG-紙品")
        check("有 missing_in_file 新欄位", "missing_in_file" in ords[0])
        check("schema_version 記為 2", c.execute("SELECT value FROM mst_meta WHERE key='schema_version'").fetchone()[0] == "2")
        c.close()
    finally:
        db.DB_PATH = saved

    print(f"\n通過 {len(PASS)} 項，失敗 {len(FAIL)} 項")
    if FAIL:
        print("失敗項目："); [print("  -", f) for f in FAIL]
        sys.exit(1)


if __name__ == "__main__":
    main()

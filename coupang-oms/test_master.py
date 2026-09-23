"""商品主檔自動化的端到端驢證：拿真實的訂單彙總表（7 月單線別、9 月多線別）與
總表範例跑完整流程。表全部是 mst_ 開頭，不碰訂單管理與採購表轉換的資料。

執行：python test_master.py
"""
import io
import re
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
PG_MASTER_XLSX = os.path.join(SAMPLES, "酷澎主檔範例_寶僑.xlsx")
CPG_MASTER_XLSX = os.path.join(SAMPLES, "酷澎主檔範例_CPG.xlsx")
PG_SHEET_XLSX = os.path.join(SAMPLES, "寶僑總表範例.xlsx")

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
    check("用的是商品主檔自己的藍色 logo，不是訂單管理的紅色 logo", "logo_master.png" in html, "拆成套件後曾因路徑算錯退回 logo.png")
    check("分頁順序照流程：① 商品主檔 → ② 訂單明細 → ③ 總表", html.index("① 商品主檔") < html.index("② 訂單明細") < html.index("③ 總表"))
    check("頂端沒有全域的線別選單（線別是篩選，不是模式）", 'id="sel-line"' not in html)
    check("OMS 首頁有「商品主檔自動化」按鈕", "商品主檔自動化" in client.get("/").get_data(as_text=True))
    meta = client.get("/api/master/lines").get_json()
    check("還沒有資料時線別清單是空的（線別是從資料長出來的）", meta["groups"] == [], str(meta["groups"]))

    print("\n【1b】模組建表失敗時整個停用（503），訂單管理不受影響")
    saved_ready, saved_err = db.MASTER_READY, db.MASTER_ERROR
    db.MASTER_READY, db.MASTER_ERROR = False, "測試用：假裝建表失敗"
    try:
        check("API 回 503 並講清楚原因", client.get("/api/master/lines").status_code == 503 and "假裝建表失敗" in client.get("/api/master/lines").get_json()["error"])
        check("頁面也回 503", client.get("/master").status_code == 503)
        check("訂單管理首頁照常開", client.get("/").status_code == 200)
    finally:
        db.MASTER_READY, db.MASTER_ERROR = saved_ready, saved_err
    check("恢復後 API 正常", client.get("/api/master/lines").status_code == 200)

    print("\n【2】① 先匯總表鋪主檔（不問線別）")
    res = upload(client, "/api/master/products/import", MASTER_XLSX)
    pi = res.get_json()
    check("總表匯入成功、抓到品類／COGS／箱入數／Note 等欄", res.status_code == 200 and {"barcode", "box_size", "category", "cost_price", "note"} <= set(pi["columns_found"]), str(pi))
    prods = client.get("/api/master/products").get_json()["products"]
    p = next((x for x in prods if x["barcode"] == "4987176340894"), None)
    check("ARIEL 4987176340894：箱入數 6、品類 Fabric、COGS 202、PG code", p and p["box_size"] == 6 and p["category"] == "Fabric" and p["cost_price"] == 202 and p["pgcode"] == "80864486", str(p and (p["box_size"], p["category"], p["cost_price"])))
    check("主檔的線別欄空的（還沒出現在任何訂單）", all(x["line_groups"] == [] for x in prods))

    print("\n【2b】① 酷澎主檔格式（一個線別一份，有線別欄）")
    res = upload(client, "/api/master/products/import", PG_MASTER_XLSX); pm = res.get_json()
    check("寶僑主檔匯入成功、抓到線別／單位／業務報價單價／效期／啟用／報價備註", res.status_code == 200 and {"master_line", "unit", "cost_price", "shelf_days", "active", "note"} <= set(pm["columns_found"]), str(pm))
    check("寶僑主檔 599 筆全進來", pm["added"] + pm["updated"] + pm["unchanged"] == 599, str(pm))
    prods = client.get("/api/master/products").get_json()["products"]
    p = next((x for x in prods if x["barcode"] == "4902430732949"), None)
    check("潘婷 4902430732949：線別寶僑、單位瓶、箱入數 12、單價 137、效期 1095、啟用 Y", p and p["master_line"] == "寶僑" and p["unit"] == "瓶" and p["box_size"] == 12 and p["cost_price"] == 137 and p["shelf_days"] == 1095 and p["active"] == "Y", str(p and {k: p[k] for k in ("master_line", "unit", "box_size", "cost_price", "shelf_days", "active")}))
    check("主檔的線別立刻分得出來（不用等訂單）", p and p["line_groups"] == ["寶僑"], str(p and p["line_groups"]))
    check("線別下拉已經有寶僑", "寶僑" in client.get("/api/master/lines").get_json()["groups"])
    p2 = next((x for x in prods if x["barcode"] == "4987176340894"), None)
    check("總表先鋪的品類／PG code 沒被酷澎主檔洗掉（只更新有值的欄位）", p2 and p2["category"] == "Fabric" and p2["pgcode"] == "80864486", str(p2 and (p2["category"], p2["pgcode"])))
    res = upload(client, "/api/master/products/import", CPG_MASTER_XLSX); pc = res.get_json()
    check("CPG 主檔匯入（120 列、119 個國條，重複的那筆算更新）", res.status_code == 200 and pc["added"] == 119 and pc["updated"] == 1, str(pc))
    prods = client.get("/api/master/products").get_json()["products"]
    p3 = next((x for x in prods if x["barcode"] == "4710104000917"), None)
    check("得意 4710104000917：線別原始值 CPG-紙品、畫面歸紙潔、報價備註進 Note", p3 and p3["master_line"] == "CPG-紙品" and p3["line_groups"] == ["紙潔"] and p3["note"].startswith("包裝升級"), str(p3 and (p3["master_line"], p3["line_groups"], p3["note"][:10])))
    check("主檔可用線別篩到紙潔的 119 筆", len(client.get("/api/master/products?line=紙潔").get_json()["products"]) == 119)
    res = client.post("/api/master/products", json={"barcode": "4710104000917", "unit": "袋", "active": "n", "shelf_days": "1825", "box_size": 8, "master_line": "CPG-紙品"})
    check("手動存主檔可以改單位／啟用／效期", res.status_code == 200 and res.get_json()["product"]["active"] == "N" and res.get_json()["product"]["shelf_days"] == 1825, res.get_data(as_text=True)[:200])

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
    check("線別從資料長出來：寶僑／瑪氏／紙潔／未分類", set(meta["groups"]) == {"寶僑", "瑪氏", "紙潔", "未分類"}, str(meta["groups"]))
    check("月份清單有 2026-09", "2026-09" in meta["months"])

    print("\n【4】訂單明細：篩選面與線別")
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
    check("沒出現在任何訂單的商品，線別由酷澎主檔給（寶僑）", p_never["line_groups"] == ["寶僑"], str(p_never["line_groups"]))
    check("主檔可用線別篩", all("紙潔" in x["line_groups"] for x in client.get("/api/master/products?line=紙潔").get_json()["products"]))

    print("\n【5】就地編輯與防互蓋")
    row = next(r for r in od_pg["rows"] if r["cases"] is not None and r["box_size"] and r["box_size"] > 1)
    oid, ver = row["id"], row["version"]
    res = jput(client, f"/api/master/orders/{oid}", {"version": ver, "qty_ship": row["box_size"] * 3})
    unchanged = next(r for r in client.get(f"/api/master/pos/{row['po_number']}").get_json()["rows"] if r["id"] == oid)
    check("改出貨數量沒選原因 → 擋下（400）、數字沒動", res.status_code == 400 and "原因" in res.get_json()["error"] and unchanged["qty_ship"] == row["qty_ship"], str(res.get_json())[:100])
    res = jput(client, f"/api/master/orders/{oid}", {"version": ver, "qty_ship": row["box_size"] * 3, "reason": "車不夠"})
    check("原因不在清單裡 → 擋下（400）", res.status_code == 400)
    res = jput(client, f"/api/master/orders/{oid}", {"version": ver, "qty_ship": row["box_size"] * 3, "reason": "其他"})
    check("選「其他」沒寫說明 → 擋下（400）", res.status_code == 400 and "其他" in res.get_json()["error"])
    res = jput(client, f"/api/master/orders/{oid}", {"version": ver, "qty_ship": row["box_size"] * 3, "reason": "缺貨"})
    check("帶原因改出貨數量 → 箱數 3、立旗標", res.status_code == 200 and res.get_json()["row"]["cases"] == 3 and res.get_json()["row"]["qty_ship_overridden"] == 1, str(res.get_json())[:100])
    lg = [l for l in client.get(f"/api/master/pos/{row['po_number']}").get_json()["logs"] if l["field"] == "qty_ship"]
    check("歷程記了原因「缺貨」", lg and lg[0]["reason"] == "缺貨", str(lg[:1]))
    check("舊版本再存被擋（409）", jput(client, f"/api/master/orders/{oid}", {"version": ver, "qty_ship": 1, "reason": "缺貨"}).status_code == 409)
    v2 = res.get_json()["row"]["version"]
    res = jput(client, f"/api/master/orders/{oid}", {"version": v2, "remarks": "打單缺貨"})
    check("改備註不用選原因", res.status_code == 200)
    check("備註存成 OP 備註", res.get_json()["row"]["remarks"] == "打單缺貨" and res.get_json()["row"]["remarks_overridden"] == 1)
    check("匯出備註 = Note＋OP 備註（若 Note 為空就只有 OP）", res.get_json()["row"]["export_note"].endswith("打單缺貨"))
    v3 = res.get_json()["row"]["version"]
    res = jput(client, f"/api/master/orders/{oid}", {"version": v3, "remarks": ""})
    check("備註清空 → 旗標歸零", res.get_json()["row"]["remarks_overridden"] == 0)

    print("\n【6】PO 改期：跨線別整張一起搬、批次、視窗儲存")
    before = orders(client, month="2026-09", pos=cross_po)
    n_rows = before["count"]; old_date = before["rows"][0]["delivery_date"]
    res = jput(client, "/api/master/pos/date", {"po_numbers": [cross_po], "delivery_date": "2026-09-27", "expected": {cross_po: old_date}})
    check("改期沒選原因 → 擋下（400）、日期沒動", res.status_code == 400 and orders(client, month="2026-09", pos=cross_po)["rows"][0]["delivery_date"] == old_date, str(res.get_json()))
    res = jput(client, "/api/master/pos/date", {"po_numbers": [cross_po], "delivery_date": "2026-09-27", "expected": {cross_po: old_date}, "reason": "沒車"})
    check(f"跨線別 PO 改期：潔品＋紙品 {n_rows} 個品項一起搬", res.status_code == 200 and res.get_json()["moved"] == n_rows, str(res.get_json()))
    lg = [l for l in client.get(f"/api/master/pos/{cross_po}").get_json()["logs"] if l["field"] == "delivery_date"]
    check("改期歷程每一筆都記了原因「沒車」", lg and all(l["reason"] == "沒車" for l in lg), str(lg[:1]))
    after = orders(client, month="2026-09", pos=cross_po)
    check("全部在 9/27", all(r["delivery_date"] == "2026-09-27" for r in after["rows"]))
    check("用舊日期當依據再改被擋（409）", jput(client, "/api/master/pos/date", {"po_numbers": [cross_po], "delivery_date": "2026-09-28", "expected": {cross_po: old_date}, "reason": "沒車"}).status_code == 409)
    two = [x["po_number"] for x in od_pg["facets"]["pos"][:2]]
    exp = {x["po_number"]: x["date"] for x in od_pg["facets"]["pos"][:2]}
    res = jput(client, "/api/master/pos/date", {"po_numbers": two, "delivery_date": "2026-09-30", "expected": exp, "reason": "其他", "reason_note": "倉庫說要併車"})
    check("批次改兩張 PO", res.status_code == 200 and res.get_json()["po_count"] == 2 and res.get_json()["moved"] > 0, str(res.get_json()))
    check("批次改完都在 9/30", all(r["delivery_date"] == "2026-09-30" for r in orders(client, month="2026-09", pos=",".join(two))["rows"]))
    lg = client.get("/api/master/logs?q=倉庫說要併車").get_json()["logs"]
    check("「其他」的說明接在歷程說明後面、原因記「其他」", lg and lg[0]["reason"] == "其他" and "批次改期" in lg[0]["note"], str(lg[:1]))
    res = jput(client, "/api/master/pos/date", {"po_numbers": two, "reset": True})
    check("恢復整合表日期不用選原因", res.status_code == 200)
    check("批次恢復整合表日期", res.status_code == 200 and all(r["delivery_date_overridden"] == 0 for r in orders(client, month="2026-09", pos=",".join(two))["rows"]))
    detail = client.get(f"/api/master/pos/{cross_po}").get_json()
    check("PO 視窗：兩個原始線別、歸紙潔、有歷程", detail["lines"] == ["紙潔"] and set(detail["lines_raw"]) == {"CPG-潔品", "CPG-紙品"} and any(l["field"] == "delivery_date" for l in detail["logs"]))
    items = [{"id": r["id"], "version": r["version"], "qty_ship": 0, "remarks": "視窗改的"} for r in detail["rows"][:2]]
    res = jput(client, f"/api/master/pos/{cross_po}/save", {"items": items, "delivery_date": "2026-09-25", "expected_date": "2026-09-27"})
    check("PO 視窗沒選原因 → 整批擋下（400），一筆都沒存", res.status_code == 400 and client.get(f"/api/master/pos/{cross_po}").get_json()["delivery_date"] == "2026-09-27", str(res.get_json()))
    res = jput(client, f"/api/master/pos/{cross_po}/save", {"items": items, "delivery_date": "2026-09-25", "expected_date": "2026-09-27", "reason": "酷澎要求"})
    check("PO 視窗一次儲存：改期＋兩個品項", res.status_code == 200 and res.get_json()["changed"] >= n_rows + 4, str(res.get_json()))
    d2 = client.get(f"/api/master/pos/{cross_po}").get_json()
    check("視窗儲存後：日期 9/25、兩筆出貨 0 且備註「視窗改的」", d2["delivery_date"] == "2026-09-25" and sum(1 for r in d2["rows"] if r["qty_ship"] == 0 and r["remarks"] == "視窗改的") == 2)
    res = jput(client, f"/api/master/pos/{cross_po}/save", {"items": [{"id": detail["rows"][0]["id"], "version": detail["rows"][0]["version"], "qty_ship": 5}], "reason": "缺貨"})
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

    print("\n【7b】④ 改單統計：幾張 PO、幾張被改過、改幾次、為什麼")
    st = client.get("/api/master/stats?month=2026-09").get_json()
    all_pos = {r["po_number"] for r in orders(client, month="2026-09")["rows"]}
    check("9 月 PO 張數 = 訂單明細裡不同 PO 數", st["total"]["pos"] == len(all_pos), f"{st['total']['pos']} vs {len(all_pos)}")
    check("有品項數與 SKU 數，SKU 數 ≤ 品項數", st["total"]["items"] >= st["total"]["skus"] > 0)
    check("被改過的 PO ≥ 1、佔比是整數百分比", st["total"]["changed_pos"] >= 1 and 0 < st["total"]["changed_pct"] <= 100)
    check("改單次數 ≥ 被改過的 PO 數", st["total"]["events"] >= st["total"]["changed_pos"])
    check("分得出酷澎改的（再匯入少一筆）與我們改的", st["total"]["coupang"] >= 1 and st["total"]["manual"] >= 1, str({k: st["total"][k] for k in ("coupang", "manual")}))
    rs = st["total"]["reasons"]
    check("原因分佈：缺貨／沒車／酷澎要求／其他都有、清單含「未填」", all(rs.get(k, 0) >= 1 for k in ("缺貨", "沒車", "酷澎要求", "其他")) and "未填" in st["reasons"], str(rs))
    check("原因加總 = 我們改的次數", sum(rs.values()) == st["total"]["manual"], f"{sum(rs.values())} vs {st['total']['manual']}")
    check("整張 PO 改期（很多品項）只算一次改單", st["total"]["events"] < len(orders(client, month="2026-09")["rows"]))
    check("改最多次的 PO 排在前面、有原因字串", st["top_pos"] and st["top_pos"][0]["events"] >= st["top_pos"][-1]["events"] and any(t["reasons"] for t in st["top_pos"]), str(st["top_pos"][:1]))
    st_pg = client.get("/api/master/stats?month=2026-09&line=寶僑").get_json()
    check("線別篩選：寶僑的 PO 數 < 全部", 0 < st_pg["total"]["pos"] < st["total"]["pos"])
    st_span = client.get("/api/master/stats?month=2026-07&month_to=2026-09").get_json()
    check("跨月：每個月一列、合計另算", len(st_span["months"]) >= 1 and st_span["total"]["pos"] >= st["total"]["pos"] and all(m["month"] <= "2026-09" for m in st_span["months"]))
    check("月份格式不對回 400", client.get("/api/master/stats?month=2026/09").status_code == 400)
    ev = client.get("/api/master/stats/events?month=2026-09").get_json()["events"]
    check("明細：一次改單一列，筆數 = 改單次數", len(ev) == st["total"]["events"], f"{len(ev)} vs {st['total']['events']}")
    check("明細最新的在最上面", ev == sorted(ev, key=lambda e: (e["when"], e["po_number"]), reverse=True))
    man = [e for e in ev if e["source"] == "manual"]
    check("我們改的每一筆都有原因、有誰改、有改了什麼", man and all(e["reason"] and e["operator"] and e["summary"] and e["kind_labels"] for e in man), str(man[:1]))
    q0 = [e for e in man if "qty" in e["kinds"] and "→0" in e["summary"]]
    check("出貨改成 0 算下修", q0 and all(e["qty_dir"] == "down" for e in q0), str(q0[:1]))
    other = [e for e in man if e["reason"] == "其他"]
    check("「其他」的說明帶在明細裡", other and any("倉庫說要併車" in e["note"] for e in other), str(other[:1]))
    cp = [e for e in ev if e["source"] == "import"]
    check("酷澎改的那筆是「品項被拿掉」、沒有原因", cp and all("gone" in e["kinds"] and e["reason"] == "" for e in cp), str(cp[:1]))
    k = st["total"]["kinds"]
    check("改了什麼：改數量（我們）= 明細裡我們改數量的筆數", k["qty_m"] == sum(1 for e in man if "qty" in e["kinds"]), str(k))
    check("改了什麼：改交期（我們）= 明細裡我們改交期的筆數", k["date_m"] == sum(1 for e in man if "date" in e["kinds"]), str(k))
    check("改了什麼：品項被拿掉（酷澎）≥ 1、下修 ≥ 1", k["gone_c"] >= 1 and k["qty_down"] >= 1, str(k))
    check("下修＋上修＋有上有下 = 改數量總次數", k["qty_down"] + k["qty_up"] + k["qty_mixed"] == k["qty_c"] + k["qty_m"], str(k))
    # 「這是測試」：不用選原因、記成「測試」、統計預設不算
    tpo = [x for x in od_pg["facets"]["pos"] if x["po_number"] not in (cross_po,) + tuple(two)][0]["po_number"]
    tdet = client.get(f"/api/master/pos/{tpo}").get_json(); tr0 = tdet["rows"][0]
    import time as _t; _t.sleep(1.05)
    res = jput(client, f"/api/master/orders/{tr0['id']}", {"version": tr0["version"], "qty_ship": (tr0["qty_ship"] or 0) + 1, "is_test": True})
    check("勾「這是測試」不用選原因也能存", res.status_code == 200, str(res.get_json())[:100])
    tl = [l for l in client.get(f"/api/master/pos/{tpo}").get_json()["logs"] if l["field"] == "qty_ship"]
    check("歷程原因記成「測試」", tl and tl[0]["reason"] == "測試", str(tl[:1]))
    st2 = client.get("/api/master/stats?month=2026-09").get_json()
    check("統計預設不算測試：改單次數沒變、另外回報 1 次測試", st2["total"]["events"] == st["total"]["events"] and st2["test_events"] == 1, f"{st2['total']['events']} vs {st['total']['events']}, test={st2['test_events']}")
    check("原因清單裡預設沒有「測試」", "測試" not in st2["reasons"])
    st3 = client.get("/api/master/stats?month=2026-09&include_test=1").get_json()
    check("勾含測試：多 1 次、原因表多「測試」一列", st3["total"]["events"] == st["total"]["events"] + 1 and st3["total"]["reasons"].get("測試") == 1 and st3["test_events"] == 0, str(st3["total"]["reasons"]))
    ev2 = client.get("/api/master/stats/events?month=2026-09").get_json()["events"]
    check("明細預設也不含測試", len(ev2) == len(ev) and not any(e["reason"] == "測試" for e in ev2))
    wbt = openpyxl.load_workbook(io.BytesIO(client.get("/api/master/stats/export?month=2026-09").data))
    check("Excel 說明頁寫了不含測試、有幾次", any("不含" in str(c.value) and "1 次" in str(c.value) for row in wbt["怎麼算的"].iter_rows() for c in row))
    res = client.get("/api/master/stats/export?month=2026-09&minutes=15")
    from urllib.parse import unquote
    check("匯出統計 Excel：200、是 xlsx、檔名帶月份", res.status_code == 200 and "spreadsheetml" in res.mimetype and "改單統計_2026-09" in unquote(res.headers.get("Content-Disposition", "")), res.headers.get("Content-Disposition", ""))
    wbs = openpyxl.load_workbook(io.BytesIO(res.data))
    check("六個分頁：每月統計／改單原因／改了什麼／改最多次的 PO／明細／怎麼算的", wbs.sheetnames == ["每月統計", "改單原因", "改了什麼", "改最多次的 PO", "明細", "怎麼算的"], str(wbs.sheetnames))
    wsd = wbs["明細"]
    check("明細分頁一次改單一列（含表頭）", wsd.max_row == len(ev) + 1 and [c.value for c in wsd[1]][:3] == ["時間", "月份", "PO 單號"], f"{wsd.max_row} vs {len(ev) + 1}")
    wsk = wbs["改了什麼"]
    check("改了什麼分頁最後一列是合計、下修欄對得上", wsk.cell(row=wsk.max_row, column=1).value == "合計" and wsk.cell(row=wsk.max_row, column=4).value == k["qty_down"])
    ws1 = wbs["每月統計"]; hdr1 = [c.value for c in ws1[1]]
    check("填了分鐘數就多一欄估計工時，最後一列是合計", "估計工時（每次 15 分鐘）" in hdr1 and ws1.cell(row=ws1.max_row, column=1).value == "合計")
    check("合計列的 PO 張數對得上", ws1.cell(row=ws1.max_row, column=2).value == st["total"]["pos"])
    res0 = client.get("/api/master/stats/export?month=2026-09")
    check("沒填分鐘數就沒有估計工時欄", "估計工時" not in "".join(str(c.value) for c in openpyxl.load_workbook(io.BytesIO(res0.data))["每月統計"][1]))

    print("\n【8】7 月單線別檔也能一起進（資料照月份分）")
    res = upload(client, "/api/master/import/preview", JUL_XLSX); pvj = res.get_json()
    check("7 月檔 53 筆全部新增", pvj["new_count"] == 53)
    client.post("/api/master/import/commit", json={"batch_id": pvj["batch_id"]})
    check("7 月看得到 53 筆、9 月還是 861", orders(client, month="2026-07")["count"] == 53 and orders(client, month="2026-09")["count"] == 861)
    row7 = next(r for r in orders(client, month="2026-07")["rows"] if r["barcode"] == "6903148182406")
    check("62 袋 ÷ 24 = 2.5833 箱，不湊整", abs(row7["cases"] - 2.5833) < 1e-3)

    print("\n【9】線別顯示名寫死，沒有設定面板")
    check("線別群組設定 API 已拿掉", client.get("/api/master/line_groups").status_code == 404)
    check("頁面上沒有「線別群組」按鈕", "線別群組" not in client.get("/master").get_data(as_text=True))
    facet_lines = {x["line"] for x in orders(client, month="2026-09")["facets"]["lines"]}
    check("CPG-潔品／CPG-紙品 在畫面上叫紙潔、原名不出現", "紙潔" in facet_lines and not any(l.startswith("CPG") for l in facet_lines), str(facet_lines))

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

    print("\n【10b】寶僑總表：匯進主檔那份就是匯出的樣子，含順序")
    check("還沒匯過總表時 API 說沒有", client.get("/api/master/template?line=寶僑").get_json()["exists"] is False)
    res = upload(client, "/api/master/products/import", PG_SHEET_XLSX); pi2 = res.get_json()
    check("寶僑總表匯進主檔：認出是總表、自動記成寶僑的樣式", res.status_code == 200 and pi2["template_line"] == "寶僑", str({k: pi2.get(k) for k in ("template_line", "sheet", "added", "updated")}))
    tm = client.get("/api/master/template?line=寶僑").get_json()
    check("記到 Sheet1、7／8／9 月日期欄與 3 個空欄", tm["exists"] and tm["sheet"] == "Sheet1" and tm["months"]["9"] == {"dates": 12, "spare": 3}, str(tm.get("months")))
    check("酷澎主檔（沒日期欄）匯入不會被當成總表樣式", upload(client, "/api/master/products/import", CPG_MASTER_XLSX).get_json()["template_line"] == "")
    check("有寫歷程", any(l["field"] == "template" for l in client.get("/api/master/logs?q=寶僑總表範例&limit=1000").get_json()["logs"]))
    s_pg = client.get("/api/master/summary?line=寶僑&month=2026-09").get_json()
    res = client.get("/api/master/export?line=寶僑&month=2026-09")
    cd_s = unquote(res.headers.get("Content-Disposition", ""))
    check("匯出檔名就叫「總表.xlsx」：不跟底稿同名、也不帶月份（整份總表不是某月的）", "''總表.xlsx" in cd_s and "寶僑總表範例" not in cd_s, cd_s)
    wbt = openpyxl.load_workbook(io.BytesIO(res.data)); wst = wbt["Sheet1"]
    tpl_ws = openpyxl.load_workbook(PG_SHEET_XLSX)["Sheet1"]
    check("原本三個分頁都在、多「系統看板」「系統填入說明」", wbt.sheetnames == ["工作表1", "Sheet1", "工作表2", "系統看板", "系統填入說明"], str(wbt.sheetnames))
    check("列數與商品順序跟底稿完全一樣", [wst.cell(row=r, column=1).value for r in range(1, 300)] == [tpl_ws.cell(row=r, column=1).value for r in range(1, 300)])
    def same_cell(a, b):   # Excel 存的 183.00000000000003 轉一手會變 183，這不算動到
        if isinstance(a, (int, float)) and isinstance(b, (int, float)):
            return abs(float(a) - float(b)) < 1e-6
        return a == b
    check("A～J 欄（業務維護的商品欄）一格都沒動", all(same_cell(wst.cell(row=r, column=c).value, tpl_ws.cell(row=r, column=c).value) for r in range(1, 245) for c in range(1, 11)))
    tpl_vals = openpyxl.load_workbook(PG_SHEET_XLSX, data_only=True)["Sheet1"]
    def ku_ok(r, c):
        a = wst.cell(row=r, column=c).value
        if isinstance(a, str) and a.startswith("="):                      # 沒動的格：公式一字不差
            return a == tpl_ws.cell(row=r, column=c).value
        return tpl_vals.cell(row=r, column=c).value is None or same_cell(a, tpl_vals.cell(row=r, column=c).value)
    check("K～U（GIV／NIV／供需）填的是系統的值，而系統的值就是這份總表帶進來的 → 數字跟原本一樣（9 月 VLOOKUP 那些格變成值，其他月份的公式不動）",
          all(ku_ok(r, c) for r in range(2, 245) for c in range(11, 22)),
          str([(r, c, tpl_ws.cell(row=1, column=c).value, str(wst.cell(row=r, column=c).value)[:30], tpl_vals.cell(row=r, column=c).value, wst.cell(row=r, column=5).value) for r in range(2, 245) for c in range(11, 22) if not ku_ok(r, c)][:6]))
    o_col = [str(c.value or "").replace("\n", " ") for c in wst[1]].index("Supply_CS (Sep)") + 1
    check("Supply_CS (Sep) 原本 124 格 VLOOKUP 外部檔 → 有數字的全部變成值（含只在總表、沒出過貨的商品）", not any(str(wst.cell(row=r, column=o_col).value).startswith("=") for r in range(2, 297) if isinstance(tpl_vals.cell(row=r, column=o_col).value, (int, float))),
          str([(r, wst.cell(row=r, column=5).value, str(wst.cell(row=r, column=o_col).value)[:40]) for r in range(2, 297) if str(wst.cell(row=r, column=o_col).value).startswith("=") and isinstance(tpl_vals.cell(row=r, column=o_col).value, (int, float))][:5]))
    bd_pg = client.get("/api/master/board?line=寶僑&month=2026-09").get_json(); bmap = {b["brand"]: b for b in bd_pg["brands"]}
    check("總表右邊品牌區的 目標／REBATE目標 匯進來當這一季（7～9 月）的品牌目標", bd_pg["quarter"]["label"] == "7～9 月" and bmap["Pampers 幫寶適"]["target_giv"] == 40000000 and bmap["ARIEL"]["rebate_target"] == 17600000, str({k: (v["target_giv"], v["rebate_target"]) for k, v in bmap.items() if k in ("Pampers 幫寶適", "ARIEL")}))
    check("ARIEL 的達成 = 這季下單 GIV 合計 ÷ 目標（目標還沒填 → 空）、DIFF = REBATE目標 − 這季 COGS 含稅合計", bmap["ARIEL"]["target_pct"] is None and bmap["ARIEL"]["rebate_diff"] == round(17600000 - bmap["ARIEL"]["q_cogs_tax"], 2))
    jput(client, "/api/master/brand_targets", {"line": "寶僑", "month": "2026-09", "brand": "ARIEL", "target_giv": 1234567})
    wst2 = openpyxl.load_workbook(io.BytesIO(client.get("/api/master/export?line=寶僑&month=2026-09").data))["Sheet1"]
    hdr2 = [str(c.value or "").strip() for c in wst2[1]]
    c_goal = hdr2.index("目標") + 1; c_brand = max(i for i, h in enumerate(hdr2[:c_goal]) if h.lower() == "brand") + 1
    row_ariel = next(r for r in range(2, 30) if str(wst2.cell(row=r, column=c_brand).value).strip() == "ARIEL")
    row_pamp = next(r for r in range(2, 30) if str(wst2.cell(row=r, column=c_brand).value).startswith("Pampers"))
    check("匯出時品牌區的 目標 填回系統的值（ARIEL 1,234,567；幫寶適維持 40,000,000）", wst2.cell(row=row_ariel, column=c_goal).value == 1234567 and wst2.cell(row=row_pamp, column=c_goal).value == 40000000, str((wst2.cell(row=row_ariel, column=c_goal).value, wst2.cell(row=row_pamp, column=c_goal).value)))
    jput(client, "/api/master/brand_targets", {"line": "寶僑", "month": "2026-09", "brand": "ARIEL", "target_giv": ""})
    check("7、8 月的日期欄沒被動到", all(same_cell(wst.cell(row=r, column=c).value, tpl_ws.cell(row=r, column=c).value) for r in range(1, 245) for c in range(22, 41)))
    hdr_t = [wst.cell(row=1, column=c).value for c in range(1, wst.max_column + 1)]
    sep_hdrs = [h for h in hdr_t if h and str(h).startswith("9/")]
    check("9/12、9/14 底稿沒有 → 自動插欄，插在日期順序的位置（9/11 → 9/12 → 9/14 → 9/17）", sep_hdrs.index("9/12交貨") == sep_hdrs.index("9/11交貨") + 1 and sep_hdrs.index("9/14交貨") == sep_hdrs.index("9/12交貨") + 1 and sep_hdrs.index("9/17交貨") == sep_hdrs.index("9/14交貨") + 1, str(sep_hdrs))
    check("業務的「9/交貨」空欄照樣留著 3 個", hdr_t.count("9/交貨") == 3)
    check("欄數多 2 欄（101 → 103）", wst.max_column == tpl_ws.max_column + 2, str(wst.max_column))
    ttl_i = next(i for i, h in enumerate(hdr_t) if h and "9月TTL" in str(h)) + 1
    check("9 月加總公式改成涵蓋整個區塊（含新欄）", str(wst.cell(row=2, column=ttl_i).value) == "=SUM(AO2:BE2)", str(wst.cell(row=2, column=ttl_i).value))
    check("右邊業務的公式跟著位移（PG剩餘 =O2-BD2 → =O2-BF2）", str(wst.cell(row=2, column=ttl_i + 1).value) == "=O2-BF2", str(wst.cell(row=2, column=ttl_i + 1).value))
    check("外部檔的 VLOOKUP 一字不動", str(wst.cell(row=2, column=ttl_i + 4).value) == str(tpl_ws["BH2"].value), str(wst.cell(row=2, column=ttl_i + 4).value)[:60])
    check("合併儲存格跟著移（BZ → CB）", sorted(str(m) for m in wst.merged_cells.ranges) == ["CB15:CB16", "CB3:CB6", "CB9:CB12"], str(sorted(str(m) for m in wst.merged_cells.ranges)))
    r0 = next(x for x in s_pg["rows"] if x["barcode"] == "4987176405289")
    rowi = next(r for r in range(2, wst.max_row + 1) if str(wst.cell(row=r, column=5).value).strip() == "4987176405289")
    def cell_for(d):
        col = next(i + 1 for i, h in enumerate(hdr_t) if h and str(h).startswith(f"9/{int(d[8:10])}交貨")); return wst.cell(row=rowi, column=col).value
    check("箱數填到對的商品列、對的日期欄（含 56.75 不湊整）", all(cell_for(d) == v for d, v in r0["by_date"].items()) and r0["by_date"].get("2026-09-21") == 56.75, str({d: cell_for(d) for d in r0["by_date"]}))
    info_rows = [[str(x) if x is not None else "" for x in row] for row in wbt["系統填入說明"].iter_rows(values_only=True)]
    flat = "\n".join("|".join(r) for r in info_rows)
    check("說明分頁列出總表沒有的商品（DNU 開頭那兩個）與自動新增的日期欄", "DNU4987176386724" in flat and "9/12、9/14" in flat, flat[:300])
    check("總表沒有的商品沒混進 Sheet1", not any(str(wst.cell(row=r, column=5).value).startswith("DNU") for r in range(2, wst.max_row + 1)))
    check("沒匯過總表的線別（瑪氏）匯出還是系統格式", openpyxl.load_workbook(io.BytesIO(client.get("/api/master/export?line=瑪氏&month=2026-09").data)).sheetnames[0] == "總表")

    print("\n【10c】總表底稿長得不一樣也要能填：用假的寶僑總表（欄少、公式少、只有 5 個日期）")
    fake_sheet = os.path.join(SAMPLES, "fake", "假_寶僑總表.xlsx")
    upload(client, "/api/master/products/import", fake_sheet)     # 換成假總表當底稿（線別多數決會是寶僑）
    tm = client.get("/api/master/template?line=寶僑").get_json()
    check("假總表接手成寶僑的底稿", tm["exists"] and tm["filename"] == "假_寶僑總表.xlsx", str(tm.get("filename")))
    res = client.get("/api/master/export?line=寶僑&month=2026-09"); wbf = openpyxl.load_workbook(io.BytesIO(res.data)); wsf = wbf["Sheet1"]
    hdr_f = [c.value for c in wsf[1]]; sep_f = [h for h in hdr_f if h and str(h).startswith("9/")]
    days = [int(h.split("/")[1].replace("交貨", "")) for h in sep_f if h.split("/")[1].replace("交貨", "").isdigit()]
    check("9 月日期欄照順序（含新插的）", days == sorted(days) and len(days) >= 8, str(sep_f))
    check("兩個「9/交貨」空欄留著", hdr_f.count("9/交貨") == 2)
    ttl_f = next(i for i, h in enumerate(hdr_f) if h and "9月TTL" in str(h)) + 1
    from openpyxl.utils import get_column_letter as _L
    first_sep = min(i for i, h in enumerate(hdr_f) if h and str(h).startswith("9/")) + 1
    check("TTL 公式涵蓋整個 9 月區塊", str(wsf.cell(row=2, column=ttl_f).value) == f"=SUM({_L(first_sep)}2:{_L(ttl_f - 1)}2)", str(wsf.cell(row=2, column=ttl_f).value))
    check("PG剩餘 公式仍指向 Supply(Sep) − TTL", str(wsf.cell(row=2, column=ttl_f + 1).value) == f"=O2-{_L(ttl_f)}2", str(wsf.cell(row=2, column=ttl_f + 1).value))
    check("樞紐分頁還在", "工作表1" in wbf.sheetnames)
    # 底稿完全沒有的月份：直接呼叫填入函式，要補一整個區塊（日期欄＋TTL）
    import master
    conn = db.get_conn(); tplrow = master._template_row(conn, "寶僑"); conn.close()
    r0 = next(x for x in s_pg["rows"] if x["by_date"])
    fake_sum = {"line": "寶僑", "month": "2026-10", "dates": ["2026-10-03", "2026-10-07"], "totals_by_date": {"2026-10-03": 5, "2026-10-07": 7},
                "rows": [{"barcode": r0["barcode"], "sku_id": r0["sku_id"], "by_date": {"2026-10-03": 5, "2026-10-07": 7}, "month_total": 12, "product_name": "", "brand": "", "box_size": 1}]}
    wb10, rep = master._fill_template(tplrow, fake_sum, "2026-10", "測試")
    h10 = [c.value for c in wb10["Sheet1"][1]]
    check("沒有 10 月區塊時自動補：10/3、10/7、10月TTL 依序接在 9 月後面", [h for h in h10 if h and (str(h).startswith("10/") or "10月TTL" in str(h))] == ["10/3交貨", "10/7交貨", "10月TTL下單總箱數"] and h10.index("10/3交貨") > h10.index("9月TTL下單總箱數"), str([h for h in h10 if h and "10" in str(h)]))
    check("補出來的 TTL 有 SUM 公式", str(wb10["Sheet1"].cell(row=2, column=h10.index("10月TTL下單總箱數") + 1).value).startswith("=SUM("))
    upload(client, "/api/master/products/import", PG_SHEET_XLSX)     # 換回真的總表給後面的測試

    print("\n【11】匯出")
    res = client.get("/api/master/export?line=紙潔&month=2026-09")   # 紙潔沒匯過總表 → 系統格式
    wbx = openpyxl.load_workbook(io.BytesIO(res.data))
    check("匯出總表（系統格式）：三欄品類在", wbx.sheetnames[0] == "總表" and "品類" in [c.value for c in wbx["總表"][1]])
    cpg_date = next(d["date"] for d in orders(client, month="2026-09", lines="紙潔")["facets"]["dates"] if d["date"] and d["rows"] >= 5)
    cpg_tab = f"{int(cpg_date[5:7]):02d}{int(cpg_date[8:10]):02d}交貨"; cpg_lbl = f"{int(cpg_date[5:7])}/{int(cpg_date[8:10])}交貨"
    res = client.get(f"/api/master/export/daily?month=2026-09&lines=紙潔&dates={cpg_date}")
    wbd = openpyxl.load_workbook(io.BytesIO(res.data))
    check(f"報價檔格式：篩紙潔 + {cpg_date} → 只有 {cpg_tab} 一個分頁", wbd.sheetnames == [cpg_tab], str(wbd.sheetnames))
    cd = unquote(res.headers.get("Content-Disposition", ""))
    check("檔名就叫專案報價檔", "專案報價檔.xlsx" in cd, cd[:100])
    wsd = wbd[cpg_tab]; hd = [c.value for c in wsd[1]]
    check("A～R 欄位順序照報價檔", hd[0] == "SKU ID" and hd[3] == "品類" and hd[10] == "單價(含稅)" and hd[14] == "備註" and hd[17] == "交貨日")
    pos_cells = [wsd.cell(r, 18).value for r in range(2, wsd.max_row + 1) if wsd.cell(r, 18).value]
    check("每張 PO 第一列 R 欄寫 PO_日期(倉別)", all(f"_{cpg_lbl}(" in v for v in pos_cells) and len(pos_cells) >= 1, str(pos_cells[:2]))
    yellow = [r for r in range(2, wsd.max_row + 1) if wsd.cell(r, 1).value is None and wsd.cell(r, 1).fill.fgColor.rgb in ("00FFFF00", "FFFFFF00")]
    check("PO 之間黃色空白列", len(yellow) == len(pos_cells) - 1, f"{len(yellow)} / {len(pos_cells)}")
    first_data = next(r for r in range(2, wsd.max_row + 1) if wsd.cell(r, 1).value)
    check("J／M／N 是活的 Excel 公式（出貨箱＝H/I、箱單價＝L×I、總計＝M×J），不是數字（Chloe 2026-09-17）",
          str(wsd.cell(first_data, 10).value).startswith("=") and f"H{first_data}/I{first_data}" in str(wsd.cell(first_data, 10).value)
          and f"L{first_data}*I{first_data}" in str(wsd.cell(first_data, 13).value) and f"M{first_data}*J{first_data}" in str(wsd.cell(first_data, 14).value),
          str((wsd.cell(first_data, 10).value, wsd.cell(first_data, 13).value, wsd.cell(first_data, 14).value)))
    check("數字欄不設「#,##0.##」之類的格式（整數會顯示成 100. 帶逗號，Chloe 2026-09-18）",
          all(wsd.cell(first_data, c).number_format in ("General", "@") for c in (7, 8, 9, 10, 11, 12, 13, 14)),
          str([wsd.cell(first_data, c).number_format for c in (7, 8, 9, 10, 11, 12, 13, 14)]))
    multi = [m for m in wsd.merged_cells.ranges if m.min_col == 18 and m.max_col == 18]
    check("同一張 PO 的交貨日欄合併成一格（有 2 個品項以上的 PO 才會合併）",
          len(multi) >= 1 and all(wsd.cell(m.min_row, 18).value and wsd.cell(m.min_row, 1).value for m in multi)
          and all(wsd.cell(rr, 1).value for m in multi for rr in range(m.min_row, m.max_row + 1)),
          str([str(m) for m in multi][:3]))
    check("紙潔的分頁裡只有 CPG 的品項（沒有寶僑）", all("Pampers" not in str(wsd.cell(r, 6).value or "") for r in range(2, wsd.max_row + 1)))
    pg_date = next(r["delivery_date"] for r in orders(client, month="2026-09", lines="寶僑")["rows"] if r["barcode"] == BC)
    res = client.get(f"/api/master/export/daily?month=2026-09&lines=寶僑&dates={pg_date}")
    wsp = openpyxl.load_workbook(io.BytesIO(res.data)).worksheets[0]
    row_a = next((r for r in wsp.iter_rows(min_row=2, values_only=True) if str(r[2]) == BC), None)
    check(f"寶僑 {pg_date} 報價檔：ARIEL 品類 Fabric、單價 = COGS 202、備註 = Note 新品", row_a and row_a[3] == "Fabric" and row_a[10] == 202 and row_a[14] == "新品", str(row_a and (row_a[3], row_a[10], row_a[14])))
    check("整合表自帶的 MPO_ 備註沒進報價檔", all("MPO_" not in str(r[14] or "") for r in wsp.iter_rows(min_row=2, values_only=True)))
    res = client.get("/api/master/export/daily?month=2026-09&warehouses=NOPE")
    check("篡到沒資料的檔案有說明", openpyxl.load_workbook(io.BytesIO(res.data)).sheetnames == ["無資料"])
    check("檔名就叫專案報價檔", "專案報價檔.xlsx" in unquote(res.headers.get("Content-Disposition", "")), res.headers.get("Content-Disposition"))
    # 跨月：7 月～9 月一次匯，分頁 0904交貨、0903交貨…新的在前，7 月的也在
    res = client.get("/api/master/export/daily?month=2026-07&month_to=2026-09")
    names = openpyxl.load_workbook(io.BytesIO(res.data)).sheetnames
    check("跨月匯出同時有 9 月與 7 月的分頁", any(n.startswith("09") for n in names) and any(n.startswith("07") for n in names), str(names[:3] + names[-3:]))
    check("分頁名稱是 MMDD交貨、新的在前", all(re.fullmatch(r"\d{4}交貨|未排日期", n) for n in names) and names[0] > names[-1], str(names[:2]))
    res = client.get("/api/master/export/daily?month=2026-09&month_to=2026-07")
    check("起訖顛倒也照樣匯（自動對調）", len(openpyxl.load_workbook(io.BytesIO(res.data)).sheetnames) == len(names))

    print("\n【12】歷程")
    logs = client.get(f"/api/master/logs?po={cross_po}").get_json()["logs"]
    check("PO 的歷程含改期、出貨數量、備註", {l["field"] for l in logs} >= {"delivery_date", "qty_ship", "remarks"})
    check("歷程可用關鍵字查", len(client.get("/api/master/logs?q=打單缺貨").get_json()["logs"]) >= 1)

    print("\n【12c】匯出範圍＝某一次匯入：分得出這次新增／有變了哪些")
    imps = client.get("/api/master/imports").get_json()["batches"]
    check("匯入歷程列出每次確認匯入（含 7 月那批）", len(imps) >= 2 and all(b["committed_at"] for b in imps), str([(b["id"], b["new_now"], b["months"]) for b in imps]))
    check("每次匯入也算 PO 張數（new_pos／changed_pos／removed_pos），PO 數不會大於品項數",
          all({"new_pos", "changed_pos", "removed_pos"} <= set(b) and b["new_pos"] <= b["new_now"] and (b["new_pos"] > 0) == (b["new_now"] > 0) for b in imps),
          str([(b["new_pos"], b["new_now"]) for b in imps]))
    od_span = orders(client, month="2026-07", month_to="2026-09")
    check("訂單 API 帶 month_to 可以一次看好幾個月（7～9 月筆數 = 各月加總）",
          od_span["count"] == orders(client, month="2026-07")["count"] + orders(client, month="2026-08")["count"] + orders(client, month="2026-09")["count"] and od_span["count"] > orders(client, month="2026-09")["count"],
          str(od_span["count"]))
    check("每次匯入都有 新增／有變／消失 三個數字，而且互斥（有變不含消失）",
          all({"new_now", "changed_now", "removed_now"} <= set(b) for b in imps)
          and all(sum(x["new"] + x["changed"] + x["removed"] for x in b["month_counts"]) == b["new_now"] + b["changed_now"] + b["removed_now"] for b in imps),
          str([(b["new_now"], b["changed_now"], b["removed_now"]) for b in imps]))
    gone = client.get("/api/master/orders?month=2026-09").get_json()["rows"]
    gone_ids = {r["last_batch_id"] for r in gone if r["missing_in_file"]}
    check("有品項消失的那批，removed_now 跟訂單裡 missing_in_file 的筆數對得上",
          all(b["removed_now"] == sum(1 for r in gone if r["missing_in_file"] and r["last_batch_id"] == b["id"]) for b in imps if b["id"] in gone_ids),
          str([(b["id"], b["removed_now"]) for b in imps if b["id"] in gone_ids]))
    jul = next((b for b in imps if b["months"] == ["2026-07"]), None)
    check("7 月那批：新增 53、落在 2026-07", jul and jul["new_now"] == 53, str(jul and (jul["new_now"], jul["months"], jul["changed_now"])))
    od = orders(client, month="2026-07", batch=jul["id"])
    check("只看這批新增：7 月 53 筆", od["count"] == 53, str(od["count"]))
    check("換到 9 月看同一批 → 0 筆（那批沒有 9 月的單）", orders(client, month="2026-09", batch=jul["id"])["count"] == 0)
    check("不存在的批次 → 0 筆，不會整月都出來", orders(client, month="2026-07", batch=999999)["count"] == 0)
    res = client.get(f"/api/master/export/daily?month=2026-07&month_to=2026-09&batch={jul['id']}&batch_scope=new")
    names = openpyxl.load_workbook(io.BytesIO(res.data)).sheetnames
    check("匯出這批的專案報價檔：跨月範圍裡只剩 7 月的分頁", names and all(n.startswith("07") for n in names), str(names))
    sep = max(imps, key=lambda b: b["new_now"])
    check("9 月那批：預設範圍（新增＋有變）不少於只看新增", orders(client, month="2026-09", batch=sep["id"])["count"] >= orders(client, month="2026-09", batch=sep["id"], batch_scope="new")["count"] > 0)
    wbb = openpyxl.load_workbook(io.BytesIO(client.get(f"/api/master/export/daily?month=2026-07&batch={jul['id']}").data))
    hdr_b = [c.value for c in wbb[wbb.sheetnames[0]][1]]
    check("匯出某一次匯入時多一欄「本次變動」、新增的寫「新增」", hdr_b[-1] == "本次變動" and wbb[wbb.sheetnames[0]].cell(row=2, column=len(hdr_b)).value == "新增", str(hdr_b[-3:]))
    check("匯出目前畫面時沒有那一欄", "本次變動" not in [c.value for c in openpyxl.load_workbook(io.BytesIO(client.get("/api/master/export/daily?month=2026-07").data)).worksheets[0][1]])
    chg = next((b for b in imps if b["changed_now"]), None)
    check("品項消失那次匯入算「有變」", chg is not None, str([(b["filename"], b["changed_now"]) for b in imps]))
    if chg:
        rows_c = [r for r in orders(client, month="2026-09", batch=chg["id"])["rows"] if r["last_batch_id"] == chg["id"] and r["first_batch_id"] != chg["id"]]
        check("有變的列記得改了什麼（出貨數量 →／品項重新出現／檔案已無此品項）", rows_c and all(any(k in (r["last_batch_changes"] or "") for k in ("→", "品項重新出現", "檔案已無此品項")) for r in rows_c), str([r.get("last_batch_changes") for r in rows_c][:2]))
        wbc = openpyxl.load_workbook(io.BytesIO(client.get(f"/api/master/export/daily?month=2026-09&batch={chg['id']}").data)); wsc = wbc.worksheets[0]
        col = [c.value for c in wsc[1]].index("本次變動") + 1
        vals = [wsc.cell(row=r, column=col).value for r in range(2, wsc.max_row + 1) if wsc.cell(row=r, column=col).value]
        check("匯出檔「本次變動」欄寫出變動內容", any(any(k in str(v) for k in ("→", "品項重新出現", "檔案已無此品項")) for v in vals), str(vals[:3]))
    html_m = client.get("/master").get_data(as_text=True)
    check("首頁不再有對帳算式；匯入紀錄是右邊抽屜，不是擋住畫面的視窗",
          "對帳" not in html_m and 'id="dlg-imports"' not in html_m and 'id="imp-drawer"' in html_m)

    print("\n【10d】總表 TTL 欄用位置認，標題寫錯月份也找得到（Chloe 的 9 月區塊寫成 12月TTL）")
    from master.summary import _sheet_layout
    wbl = openpyxl.Workbook(); wsl = wbl.active
    wsl.append(["skuid", "Barcode", "箱入數", "8/4交貨_1", "11月TTL下單總箱數", "9/2交貨", "9/3交貨", "9/交貨", "12月TTL下單總箱數", "Note"])
    lay9 = _sheet_layout(wsl, 1, 9); lay8 = _sheet_layout(wsl, 1, 8)
    check("9 月區塊的 TTL 是第 9 欄（標題雖寫 12月）", lay9["ttl_col"] == 9, str(lay9["ttl_col"]))
    check("8 月區塊的 TTL 是第 5 欄（標題雖寫 11月）", lay8["ttl_col"] == 5, str(lay8["ttl_col"]))
    check("月份區塊的欄位清單含日期欄與 TTL 欄", set(lay9["month_cols"]) == {4, 5, 6, 7, 8, 9}, str(lay9["month_cols"]))

    print("\n【10e】填總表只動系統有訂單的那幾天，業務手填在別天的數字不能被洗掉（Chloe 2026-09-18）")
    from master.summary import _write_cases, _row_index
    wbw = openpyxl.Workbook(); wsw = wbw.active
    wsw.append(["skuid", "Barcode", "箱入數", "10/2交貨_1", "10/5交貨_1", "10/6交貨_1", "10/交貨_1", "10月TTL下單總箱數"])
    wsw.append(["S1", "4987176340894", 6, 618, 30, 99, None, "=SUM(D2:G2)"])
    wsw.append(["S2", "4987176340863", 6, 154, None, None, None, "=SUM(D3:G3)"])
    layw = _sheet_layout(wsw, 1, 10); by_bc, by_sku = _row_index(wsw, 1, layw["bc_col"], layw["sku_col"])
    summ = {"dates": ["2026-10-06"], "rows": [
        {"barcode": "4987176340894", "sku_id": "S1", "by_date": {"2026-10-06": 1}},
        {"barcode": "4987176340863", "sku_id": "S2", "by_date": {}}]}
    _write_cases(wsw, layw, layw["month_cols"], by_bc, by_sku, summ)
    check("10/2、10/5 業務手填的數字原封不動", wsw["D2"].value == 618 and wsw["E2"].value == 30 and wsw["D3"].value == 154, str((wsw["D2"].value, wsw["E2"].value, wsw["D3"].value)))
    check("10/6 系統有訂單：第一個商品填 1，第二個商品那天沒出貨 → 舊值清成空白", wsw["F2"].value == 1 and wsw["F3"].value is None, str((wsw["F2"].value, wsw["F3"].value)))

    print("\n【12b】清除資料（只有管理員）")
    before_orders = orders(client, month="2026-09")["count"]
    before_prods = len(client.get("/api/master/products").get_json()["products"])
    check("清之前有訂單也有主檔", before_orders > 0 and before_prods > 0)
    app_module._write_users(app_module.get_users(), {"Jerry"})   # 小真暫時不是管理員
    res = client.post("/api/master/reset", json={"confirm": "清空資料"})
    check("不是管理員 → 403", res.status_code == 403, str(res.status_code))
    app_module._write_users(app_module.get_users(), {"Jerry", "小真"})
    res = client.post("/api/master/reset", json={"confirm": "清空"})
    check("沒照著打「清空資料」→ 400", res.status_code == 400)
    res = client.post("/api/master/reset", json={"confirm": "清空資料", "orders": False, "products": False})
    check("一個範圍都沒勾 → 400", res.status_code == 400)
    check("擋掉的那幾次什麼都沒清", orders(client, month="2026-09")["count"] == before_orders)
    res = client.post("/api/master/reset", json={"confirm": "清空資料", "orders": True, "products": False, "keep_logs": True})
    check("清訂單明細成功", res.status_code == 200 and res.get_json()["ok"], res.get_data(as_text=True))
    check("訂單清空了", orders(client, month="2026-09")["count"] == 0)
    check("主檔沒被動到", len(client.get("/api/master/products").get_json()["products"]) == before_prods)
    logs_after = client.get("/api/master/logs?q=清除資料").get_json()["logs"]
    check("保留歷程時多記一筆「清除資料」系統紀錄", any(l["field"] == "reset" and l["source"] == "system" and l["operator"] == "小真" for l in logs_after))
    check("舊歷程還在", len(client.get("/api/master/logs?q=打單缺貨").get_json()["logs"]) >= 1)
    check("月份清單變空、線別還在（線別是主檔學來的）", client.get("/api/master/lines").get_json()["months"] == [] and client.get("/api/master/lines").get_json()["groups"] != [])
    res = client.post("/api/master/reset", json={"confirm": "清空資料", "orders": False, "products": True})
    check("再清主檔（含歷程）", res.status_code == 200 and "商品主檔" in res.get_json()["message"])
    check("主檔清空了", client.get("/api/master/products").get_json()["products"] == [])
    check("歷程一併清空", client.get("/api/master/logs").get_json()["logs"] == [])
    check("主檔也清了之後線別下拉才變空", client.get("/api/master/lines").get_json()["groups"] == [])
    # 清完再匯一次要能照常用，才算真的清乾淨、沒留下卡住的殘骸
    res = upload(client, "/api/master/products/import", MASTER_XLSX)
    check("清完主檔可以重新匯總表", res.status_code == 200, res.get_data(as_text=True)[:200])
    # 一次丟兩份（7 月＋9 月）：兩份都要讀到，不能只吃第一份
    files = [(io.BytesIO(open(pth, "rb").read()), os.path.basename(pth)) for pth in (JUL_XLSX, SEP_XLSX)]
    res = client.post("/api/master/import/preview", data={"file": files}, content_type="multipart/form-data")
    pv = res.get_json()
    check("一次上傳兩份彙總表都讀到（53 + 861）", res.status_code == 200 and pv["rows_total"] == 53 + 861, str(pv.get("rows_total")))
    check("預覽列出兩個檔名", len(pv.get("files", [])) == 2 and "、" in pv["filename"], str(pv.get("files")))
    check("預覽月份同時有 7 月和 9 月", {d[:7] for d in pv["dates"]} >= {"2026-07", "2026-09"}, str(sorted({d[:7] for d in pv["dates"]})))
    res = client.post("/api/master/import/commit", json={"batch_id": pv["batch_id"], "add_missing_products": True})
    check("清完訂單可以重新匯訂單彙總表（兩個月一起）", res.status_code == 200 and orders(client, month="2026-09")["count"] == 861 and orders(client, month="2026-07")["count"] == 53, res.get_data(as_text=True)[:200])
    # 兩份檔案撞到同一張 PO／SKU：以後面那份為準並提醒
    files = [(io.BytesIO(open(JUL_XLSX, "rb").read()), "a_7月.xlsx"), (io.BytesIO(open(JUL_XLSX, "rb").read()), "b_7月再一份.xlsx")]
    pv = client.post("/api/master/import/preview", data={"file": files}, content_type="multipart/form-data").get_json()
    check("跨檔案重複的 PO／SKU 不會重複算", pv["rows_total"] == 53, str(pv["rows_total"]))
    check("有提醒哪一份跟前面重複", any("b_7月再一份.xlsx" in w and "重複" in w for w in pv["warnings"]), str(pv["warnings"][-1:]))
    # 其中一份壞掉：整批不收、點名是哪一份
    files = [(io.BytesIO(open(JUL_XLSX, "rb").read()), "好的.xlsx"), (io.BytesIO(b"not excel"), "壞的.xlsx")]
    res = client.post("/api/master/import/preview", data={"file": files}, content_type="multipart/form-data")
    check("一份壞掉整批不收，錯誤訊息點名檔名", res.status_code == 400 and "壞的.xlsx" in res.get_json()["error"], res.get_data(as_text=True)[:200])

    print("\n【12d】外部來源檔：總表帶 GIV／NIV／供需、supply 表、Coupang Master、庫存銷售表")
    FAKE = os.path.join(SAMPLES, "fake")
    res = upload(client, "/api/master/products/import", os.path.join(FAKE, "假_寶僑總表.xlsx")); d = res.get_json()
    check("假總表匯入成功、認出是總表（kind=sheet）", res.status_code == 200 and d.get("kind") == "sheet" and d["template_line"] == "寶僑", str(d)[:160])
    check("總表順便帶進 GIV／NIV 與 7～9 月 Supply", d["extras"]["price_updated"] > 0 and {"2026-07", "2026-08", "2026-09"} <= set(d["extras"]["months"]), str(d["extras"]))
    prods = {p["barcode"]: p for p in client.get("/api/master/products").get_json()["products"]}
    fk = openpyxl.load_workbook(os.path.join(FAKE, "假_寶僑總表.xlsx")).worksheets[0]
    fhdr = [c.value for c in fk[1]]; frows = [dict(zip(fhdr, [c.value for c in r])) for r in fk.iter_rows(min_row=2) if r[4].value]
    bc0, bc1, bc3, bc_last = (str(frows[i]["Barcode"]) for i in (0, 1, 3, len(frows) - 1))
    check("主檔 GIV 等於總表 K 欄（原樣存，不四捨五入）", prods[bc0]["giv"] == float(frows[0]["GIV"]), f"{prods[bc0]['giv']} vs {frows[0]['GIV']}")
    ms = client.get("/api/master/month_stats?month=2026-07").get_json()
    row0 = next((r for r in ms["rows"] if r["barcode"] == bc0), None)
    check("月統計有 7 月 Supply（來自總表 Supply_CS (Jul)）", row0 and row0["supply_cs"] == float(frows[0]["Supply_CS (Jul)"]), str(row0)[:120])

    res = upload(client, "/api/master/products/import", os.path.join(FAKE, "假_寶僑supply表.xlsx")); d = res.get_json()
    check("supply 表：認出（kind=supply）、對到 18 個商品、主檔沒有的 1 個略過", res.status_code == 200 and d.get("kind") == "supply" and d["matched"] == 18 and d["not_in_master"] == 1, str(d)[:200])
    check("supply 表：月份從表頭抓到 10～12 月（年份自己推）", d["months"] == ["2026-10", "2026-11", "2026-12"], str(d["months"]))
    check("supply 表：#N/A 那格算「來源沒值」略過", d["skipped"] >= 1)
    ms10 = {r["barcode"]: r for r in client.get("/api/master/month_stats?month=2026-10").get_json()["rows"]}
    check("10 月：第一個商品 Supply 是 #N/A → 空，demand 有值", ms10[bc0]["supply_cs"] is None and ms10[bc0]["demand_cs"] is not None, str(ms10[bc0]))
    check("10 月：其他商品 Supply、demand 都有", ms10[bc1]["supply_cs"] is not None and ms10[bc1]["demand_cs"] is not None)
    check("PG 最大剩餘可供貨量 = Supply − 該月下單（還沒下單 → 等於 Supply）", ms10[bc1]["pg_remaining_supply"] == ms10[bc1]["supply_cs"])

    giv_before = prods[bc1]["giv"]
    res = upload(client, "/api/master/products/import", os.path.join(FAKE, "假_CoupangMaster.xlsx")); d = res.get_json()
    check("Coupang Master：跳過隱藏分頁與說明頁，認出「4月」（kind=master_price）", res.status_code == 200 and d.get("kind") == "master_price" and d["sheet"] == "4月", str(d)[:200])
    check("Coupang Master：對到 17 個（最後一個商品 Master 沒有）", d["matched"] == 17 and d["not_in_master"] == 0, str(d)[:120])
    prods = {p["barcode"]: p for p in client.get("/api/master/products").get_json()["products"]}
    check("GIV 照 Master 更新（×1.1）", abs(prods[bc0]["giv"] - round(float(frows[0]["GIV"]) * 1.1, 2)) < 1e-6, f"{prods[bc0]['giv']}")
    check("Master 是 #N/A 的商品 GIV 不被蓋掉", prods[bc1]["giv"] == giv_before, f"{prods[bc1]['giv']} vs {giv_before}")
    check("GIV 變動有記歷程", any(l["field"] == "giv" for l in client.get("/api/master/logs?q=假_CoupangMaster").get_json()["logs"]))

    res = upload(client, "/api/master/products/import", os.path.join(FAKE, "假_庫存銷售表.xlsx")); d = res.get_json()
    check("庫存銷售表：認出（kind=stock）、年份從 Y26 來、4～10 月", res.status_code == 200 and d.get("kind") == "stock" and d["months"] == [f"2026-{m:02d}" for m in range(4, 11)], str(d)[:200])
    st = openpyxl.load_workbook(os.path.join(FAKE, "假_庫存銷售表.xlsx")).worksheets[0]
    shdr = [c.value for c in st[1]]; srows = {str(r[5].value): dict(zip(shdr, [c.value for c in r])) for r in st.iter_rows(min_row=2)}
    r1 = srows[bc1]
    ordered = sum(r1[f"Y26 {m}月實際下單總箱數"] for m in range(4, 9)) + r1["Y26 9月目前下單總箱數(在途)"]
    sold = sum(r1[f"Y26 {m}月實銷總箱數"] for m in range(4, 9)) + r1["Y26 9/1-9/20實銷總箱數"]
    ms9 = {r["barcode"]: r for r in client.get("/api/master/month_stats?month=2026-09").get_json()["rows"]}
    check("剩餘庫存 = 累計下單 − 累計實銷（照庫存表 AK 欄公式）", ms9[bc1]["remaining_cs"] == round(ordered - sold, 2), f"{ms9[bc1]['remaining_cs']} vs {ordered - sold}")
    check("「9/1-9/20實銷」涵蓋 20 天 → 庫存天數 = 剩餘 ÷ (實銷÷20)（AM 欄）", ms9[bc1]["sold_days"] == 20 and ms9[bc1]["stock_days"] == round((ordered - sold) / (r1["Y26 9/1-9/20實銷總箱數"] / 20), 1), str(ms9[bc1]))
    daily = r1["Y26 9/1-9/20實銷總箱數"] / 20
    check("預計到月底銷售 = 日均 × 30、到月底剩餘 = 剩餘 − 日均 × 10（AN／AO 欄）", ms9[bc1]["month_sales_proj"] == round(daily * 30, 2) and ms9[bc1]["remaining_eom"] == round((ordered - sold) - daily * 10, 2), str(ms9[bc1]))
    ms4 = {r["barcode"]: r for r in client.get("/api/master/month_stats?month=2026-04").get_json()["rows"]}
    check("4 月下單是 #REF! 的商品 → 4 月 ordered 空、不當 0", ms4[bc3]["ordered_cs"] is None and ms4[bc3]["sold_cs"] is not None, str(ms4[bc3]))
    src = client.get("/api/master/sources").get_json()["sources"]
    check("來源狀態：四種都有「上次上傳」", [x["kind"] for x in src] == ["supply", "master_price", "stock", "sheet"] and all(x["last"] for x in src), str([(x["kind"], bool(x["last"])) for x in src]))
    check("來源上傳有記歷程", len([l for l in client.get("/api/master/logs?q=對到").get_json()["logs"] if l["field"] == "source_upload"]) >= 4)
    res = client.post("/api/master/products/import", data={"file": (io.BytesIO(b"PK\x03\x04junk"), "怪檔.xlsx")}, content_type="multipart/form-data")
    check("認不出來的檔一樣回 400", res.status_code == 400)

    print("\n【12e】③ 總表看板：一列一商品，數字照總表公式算；品牌層目標可填")
    pv = upload(client, "/api/master/import/preview", os.path.join(FAKE, "假_訂單彙總表_9月_第一次.xlsx")).get_json()
    check("假 9 月訂單匯進去", client.post("/api/master/import/commit", json={"batch_id": pv["batch_id"], "add_missing_products": True}).status_code == 200)
    check("沒給線別 → 400", client.get("/api/master/board?month=2026-09").status_code == 400)
    check("月份格式錯 → 400", client.get("/api/master/board?line=寶僑&month=2026/9").status_code == 400)
    bd = client.get("/api/master/board?line=寶僑&month=2026-09").get_json()
    rows = {r["barcode"]: r for r in bd["rows"]}
    ordered_rows = [r for r in bd["rows"] if r["ttl"]]
    check("9 月看板有列、有下單的商品不只一個", len(bd["rows"]) >= 10 and len(ordered_rows) >= 5, f"{len(bd['rows'])} rows, {len(ordered_rows)} ordered")
    bd10 = client.get("/api/master/board?line=寶僑&month=2026-10").get_json()
    check("10 月還沒下單：supply 表有 Supply 的商品照樣列出來（ttl=0、剩餘可供貨 = Supply）", len(bd10["rows"]) >= 10 and all(r["ttl"] == 0 for r in bd10["rows"])
          and all(r["remaining_supply"] == r["supply"] for r in bd10["rows"] if r["supply"] is not None), f"{len(bd10['rows'])} rows")
    r = next(r for r in ordered_rows if r["giv"] and r["cost"] and r["box_size"] and r["niv"])
    check("下單 GIV = 下單 × GIV（總表 BN）", r["ttl_giv"] == round(r["ttl"] * r["giv"], 2), str((r["ttl"], r["giv"], r["ttl_giv"])))
    check("COGS 含稅 = 下單 × COGS × 箱入數（總表 BO）", r["cogs_tax"] == round(r["ttl"] * r["cost"] * r["box_size"], 2), str((r["ttl"], r["cost"], r["box_size"], r["cogs_tax"])))
    check("永豐成本未稅 = 下單 × NIV × 1.02（總表 BP）", r["yf_cost"] == round(r["ttl"] * r["niv"] * 1.02, 2), str((r["ttl"], r["niv"], r["yf_cost"])))
    rs = next(r for r in bd["rows"] if r["supply"] is not None and r["giv"])
    check("Supply GIV = Supply × GIV（總表 BK）、剩餘可供貨 = Supply − 下單（總表 BE）",
          rs["supply_giv"] == round(rs["supply"] * rs["giv"], 2) and rs["remaining_supply"] == round(rs["supply"] - rs["ttl"], 2), str(rs))
    check("每一列的超打旗標跟數字一致", all((("over" in r["flags"]) or ("no_over" in r["flags"])) == (r["supply"] is not None and r["ttl"] > r["supply"]) for r in bd["rows"]))
    check("庫存跟月統計 API 算的是同一個數", rows[bc1]["stock_remaining"] == ms9[bc1]["remaining_cs"] and rows[bc1]["stock_days"] == ms9[bc1]["stock_days"], str((rows[bc1]["stock_remaining"], ms9[bc1]["remaining_cs"])))
    check("庫存天數低於 14 天的都標 low_stock", all(("low_stock" in r["flags"]) == (r["stock_days"] is not None and r["stock_days"] < bd["low_stock_days"]) for r in bd["rows"]))
    t = bd["totals"]
    check("合計 = 各列加總（下單、下單 GIV、COGS 含稅）", t["ttl"] == round(sum(r["ttl"] for r in bd["rows"]), 2) and t["ttl_giv"] == round(sum(r["ttl_giv"] or 0 for r in bd["rows"]), 2)
          and t["cogs_tax"] == round(sum(r["cogs_tax"] or 0 for r in bd["rows"]), 2), str(t))
    check("合計有 supply、有庫存、有幾個商品要注意", t["has_supply"] and t["has_stock"] and t["alerts"] == sum(1 for r in bd["rows"] if r["flags"]))
    check("每日出貨的日期與各日合計還在（原本的寬表）", bd["dates"] and all(d in bd["totals_by_date"] for d in bd["dates"]))
    brands = {b["brand"]: b for b in bd["brands"]}
    b0 = max(bd["brands"], key=lambda b: b["ttl_giv"])
    check("品牌層 = 同品牌商品加總（SUMIF）", b0["ttl_giv"] == round(sum(r["ttl_giv"] or 0 for r in bd["rows"] if r["brand"] == b0["brand"]), 2)
          and b0["products"] == sum(1 for r in bd["rows"] if r["brand"] == b0["brand"]), str(b0))
    check("還沒填目標 → 達成、diff 都空", b0["target_giv"] is None and b0["target_pct"] is None and b0["target_diff"] is None)
    res = jput(client, "/api/master/brand_targets", {"line": "寶僑", "month": "2026-09", "brand": b0["brand"], "target_giv": "50000"})
    check("填品牌目標 200", res.status_code == 200 and res.get_json()["target"]["target_giv"] == 50000.0, res.get_data(as_text=True)[:120])
    b1 = next(b for b in client.get("/api/master/board?line=寶僑&month=2026-09").get_json()["brands"] if b["brand"] == b0["brand"])
    check("達成 = 這季下單 GIV 合計 ÷ 目標、diff = 目標 − 這季下單 GIV 合計（目標是一季的，跟總表 CJ 一樣）", b1["target_pct"] == round(b0["q_ttl_giv"] / 50000 * 100) and b1["target_diff"] == round(50000 - b0["q_ttl_giv"], 2), str(b1))
    q_sum = sum(next((x["ttl_giv"] for x in client.get(f"/api/master/board?line=寶僑&month={m}").get_json()["brands"] if x["brand"] == b0["brand"]), 0) for m in ("2026-07", "2026-08", "2026-09"))
    check("這季下單 GIV 合計 = 7、8、9 月三張看板同品牌下單 GIV 相加", abs(b0["q_ttl_giv"] - q_sum) < 0.02, f"{b0['q_ttl_giv']} vs {q_sum}")
    res = jput(client, "/api/master/brand_targets", {"line": "寶僑", "month": "2026-09", "brand": b0["brand"], "rebate_target": 1234.5})
    b2 = next(b for b in client.get("/api/master/board?line=寶僑&month=2026-09").get_json()["brands"] if b["brand"] == b0["brand"])
    check("REBATE 目標另外填、原本的目標不動、REBATE DIFF = 目標 − 這季 COGS 含稅合計", res.status_code == 200 and b2["target_giv"] == 50000.0 and b2["rebate_target"] == 1234.5
          and b2["rebate_diff"] == round(1234.5 - b2["q_cogs_tax"], 2), str(b2))
    check("目標是一季的：8 月看到同一個目標，10 月（下一季）沒有", next(b for b in client.get("/api/master/board?line=寶僑&month=2026-08").get_json()["brands"] if b["brand"] == b0["brand"])["target_giv"] == 50000.0
          and next((b for b in client.get("/api/master/board?line=寶僑&month=2026-10").get_json()["brands"] if b["brand"] == b0["brand"]), {"target_giv": None})["target_giv"] is None)
    check("看板合計的目標 = 各品牌目標加總", client.get("/api/master/board?line=寶僑&month=2026-09").get_json()["totals"]["target_giv"] == 50000.0)
    res = jput(client, "/api/master/brand_targets", {"line": "寶僑", "month": "2026-09", "brand": b0["brand"], "target_giv": ""})
    b3 = next(b for b in client.get("/api/master/board?line=寶僑&month=2026-09").get_json()["brands"] if b["brand"] == b0["brand"])
    check("空字串 = 清掉目標", res.status_code == 200 and b3["target_giv"] is None and b3["target_pct"] is None and b3["rebate_target"] == 1234.5)
    check("目標不是數字 → 400", jput(client, "/api/master/brand_targets", {"line": "寶僑", "month": "2026-09", "brand": b0["brand"], "target_giv": "五萬"}).status_code == 400)
    check("缺品牌 → 400；沒有要改的欄位 → 400", jput(client, "/api/master/brand_targets", {"line": "寶僑", "month": "2026-09", "target_giv": 1}).status_code == 400
          and jput(client, "/api/master/brand_targets", {"line": "寶僑", "month": "2026-09", "brand": b0["brand"]}).status_code == 400)
    tl = [l for l in client.get("/api/master/logs?q=" + b0["brand"] + "&limit=1000").get_json()["logs"] if l["field"] in ("brand_target_giv", "brand_rebate_target")]
    check("填目標有記歷程（填、改 REBATE、清掉各一筆）", len(tl) >= 3, str([(l["field"], l["old_value"], l["new_value"]) for l in tl][:4]))

    print("\n【12f】匯出總表：畫面上的數字要跟匯出的一樣（GIV／NIV、供需、庫存、品牌目標填回底稿；附系統看板）")
    jput(client, "/api/master/brand_targets", {"line": "寶僑", "month": "2026-09", "brand": b0["brand"], "target_giv": 50000})
    bd = client.get("/api/master/board?line=寶僑&month=2026-09").get_json(); rows = {r["barcode"]: r for r in bd["rows"]}
    res = client.get("/api/master/export?line=寶僑&month=2026-09")
    wbe = openpyxl.load_workbook(io.BytesIO(res.data)); wse = wbe["Sheet1"]
    check("底稿是假總表、多「系統看板」「系統填入說明」", "系統看板" in wbe.sheetnames and wbe.sheetnames[-1] == "系統填入說明", str(wbe.sheetnames))
    hde = [str(c.value or "").replace("\n", " ").strip() for c in wse[1]]
    col = lambda name: next(i + 1 for i, h in enumerate(hde) if h.replace(" ", "").lower() == name.replace(" ", "").lower())   # noqa: E731
    rowe = {str(wse.cell(row=r, column=5).value).strip(): r for r in range(2, wse.max_row + 1) if wse.cell(row=r, column=5).value}
    r1 = rowe[bc1]; x1 = rows[bc1]
    check("GIV／NIV 填成主檔的值（Coupang Master 更新過的，不是底稿貼死的舊值）", wse.cell(row=r1, column=col("GIV")).value == x1["giv"] and wse.cell(row=r1, column=col("NIV")).value == x1["niv"]
          and wse.cell(row=rowe[bc0], column=col("GIV")).value == rows[bc0]["giv"] != frows[0]["GIV"], str((wse.cell(row=r1, column=col("GIV")).value, x1["giv"])))
    check("Supply_CS (Sep) 填成 supply 資料的值", wse.cell(row=r1, column=col("Supply_CS (Sep)")).value == x1["supply"], str((wse.cell(row=r1, column=col("Supply_CS (Sep)")).value, x1["supply"])))
    fk_r1 = next(r for r in range(2, fk.max_row + 1) if str(fk.cell(row=r, column=5).value) == bc1)
    check("Supply_CS (Oct) 不是這個月的，不動", wse.cell(row=r1, column=col("Supply_CS (Oct)")).value == fk.cell(row=fk_r1, column=19).value)
    check("目前庫存數量 填成系統算的剩餘庫存（不是底稿的舊數字）", wse.cell(row=r1, column=col("目前庫存數量")).value == x1["stock_remaining"] == ms9[bc1]["remaining_cs"], str((wse.cell(row=r1, column=col("目前庫存數量")).value, x1["stock_remaining"])))
    check("每一個對得到的商品：GIV、Supply、庫存都跟看板一樣", all(
        (x["giv"] is None or wse.cell(row=rowe[bc], column=col("GIV")).value == x["giv"]) and (x["supply"] is None or wse.cell(row=rowe[bc], column=col("Supply_CS (Sep)")).value == x["supply"])
        and (x["stock_remaining"] is None or wse.cell(row=rowe[bc], column=col("目前庫存數量")).value == x["stock_remaining"]) for bc, x in rows.items() if bc in rowe))
    check("底稿裡的公式（PG剩餘、TTL）還是公式，沒被貼成值", str(wse.cell(row=r1, column=col("PG剩餘 可供貨量(CS)")).value).startswith("=") and str(wse.cell(row=r1, column=col("9月TTL下單總箱數")).value).startswith("=SUM("))
    kb = wbe["系統看板"]; kb_txt = "\n".join("|".join(str(c) for c in row if c is not None) for row in kb.iter_rows(values_only=True))
    check("系統看板：品牌列、商品列跟畫面一樣多，目標與季累計在", all(b["brand"] in kb_txt for b in bd["brands"]) and all(x["barcode"] in kb_txt for x in bd["rows"]) and "累計" in kb_txt and "50,000" not in kb_txt, kb_txt[:200])
    kb_rows = list(kb.iter_rows(values_only=True))
    brow = next(r for r in kb_rows if r[0] == b0["brand"])
    check("系統看板品牌列：季累計下單 GIV、目標、達成、差距是數字且等於看板", brow[8] == b0["q_ttl_giv"] and brow[9] == 50000 and brow[10] == round(b0["q_ttl_giv"] / 50000 * 100) and brow[11] == round(50000 - b0["q_ttl_giv"], 2), str(brow[:12]))
    prow = next(r for r in kb_rows if r[1] == bc1)
    check("系統看板商品列：COGS／GIV／Supply／下單／剩餘庫存／下單 GIV 都等於看板", prow[7] == x1["cost"] and prow[8] == x1["giv"] and prow[10] == x1["supply"] and prow[12] == x1["ttl"] and prow[14] == x1["stock_remaining"] and prow[18] == x1["ttl_giv"], str(prow[:20]))
    info_txt = "\n".join("|".join(str(c) for c in row if c is not None) for row in wbe["系統填入說明"].iter_rows(values_only=True))
    check("填入說明寫了 GIV／NIV、供需、庫存三欄各填幾格；品牌區底稿沒有就說跳過", "GIV／NIV|填了" in info_txt and "Supply_CS／demand_CS (9月)|填了" in info_txt and "到月底庫存天數|填了" in info_txt and "底稿沒有品牌區" in info_txt, info_txt[:400])
    wbn = openpyxl.load_workbook(io.BytesIO(client.get("/api/master/export?line=紙潔&month=2026-09").data))
    check("沒底稿的線別（紙潔）系統格式也附系統看板", wbn.sheetnames == ["總表", "系統看板", "每日合計"], str(wbn.sheetnames))

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
        check("schema_version 記為 4", c.execute("SELECT value FROM mst_meta WHERE key='schema_version'").fetchone()[0] == "4")
        check("升級後訂單有 first_batch_id／last_batch_id", {"first_batch_id", "last_batch_id"} <= set(ords[0].keys()))
        check("升級後主檔有 unit／master_line／active 新欄位", {"unit", "master_line", "active", "shelf_days", "date_format"} <= set(prods[0].keys()), str(sorted(prods[0].keys())))
        c.close()
    finally:
        db.DB_PATH = saved

    print(f"\n通過 {len(PASS)} 項，失敗 {len(FAIL)} 項")
    if FAIL:
        print("失敗項目："); [print("  -", f) for f in FAIL]
        sys.exit(1)


if __name__ == "__main__":
    main()

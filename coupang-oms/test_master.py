"""業績總表自動化的端到端驗證：拿真實的訂單彙總表與總表範例跑完整流程。

跟訂單管理（test_flow.py）、採購表轉換（test_purchase.py）一樣獨立成一支，
表全部是 mst_ 開頭，不會碰到另外兩套的資料。

執行：python test_master.py
"""
import io
import os
import sys
import tempfile

import openpyxl

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SAMPLES = os.path.join(BASE_DIR, "samples", "master")
ORDERS_XLSX = os.path.join(SAMPLES, "訂單彙總表範例.xlsx")
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
    mark = "✓" if condition else "✗"
    print(f"  {mark} {name}" + (f"  — {detail}" if detail else ""))


def sign_in(client, username="小真", password="changeme123"):
    res = client.post("/login", data={"username": username, "password": password})
    assert res.status_code == 302, f"測試帳號登入失敗：{username}"


def upload(client, url, path, **form):
    with open(path, "rb") as fh:
        data = fh.read()
    payload = dict(form)
    payload["file"] = (io.BytesIO(data), os.path.basename(path))
    return client.post(url, data=payload, content_type="multipart/form-data")


def jput(client, url, body):
    return client.put(url, json=body)


def main():
    db.init_db()
    app_module.app.config["TESTING"] = True
    client = app_module.app.test_client()
    sign_in(client)
    LINE, MONTH = "寶僑", "2026-07"

    print("\n【1】頁面與入口")
    res = client.get("/master")
    check("頁面打得開", res.status_code == 200)
    check("頁面標題是「業績總表自動化」", "業績總表自動化" in res.get_data(as_text=True))
    check("頁面用藍白主題、載入自己的 logo", "logo_master.png" in res.get_data(as_text=True))
    home = client.get("/").get_data(as_text=True)
    check("OMS 首頁有「業績總表自動化」按鈕", "業績總表自動化" in home and "/master" in home)
    lines = client.get("/api/master/lines").get_json()["lines"]
    check("線別預設含寶僑／紙潔／瑪氏", {"寶僑", "紙潔", "瑪氏"} <= set(lines))

    print("\n【2】匯入訂單彙總表：預覽")
    res = upload(client, "/api/master/import/preview", ORDERS_XLSX)
    assert res.status_code == 200, res.get_json()
    pv = res.get_json()
    check("讀到 53 個品項（跟 Excel 資料列一致）", pv["rows_total"] == 53, f"實際 {pv['rows_total']}")
    check("全部判定為新增", pv["new_count"] == 53)
    check("線別自動歸到「寶僑」", pv["lines"] == {"寶僑": 53}, str(pv["lines"]))
    check("交貨日有三天（7/21、7/23、7/24）",
          pv["dates"] == ["2026-07-21", "2026-07-23", "2026-07-24"], str(pv["dates"]))
    check("主檔還沒有任何國條，全部列為待建",
          len(pv["missing_products"]) > 40, f"{len(pv['missing_products'])} 個")
    missing_no_box = [m for m in pv["missing_products"] if not m["box_size"]]
    check("兩個整合表沒有箱入數的新品被標出來", len(missing_no_box) == 2,
          str([m["barcode"] for m in missing_no_box]))

    print("\n【3】匯入訂單彙總表：確認寫入（同時自動建主檔）")
    res = client.post("/api/master/import/commit", json={"batch_id": pv["batch_id"]})
    assert res.status_code == 200, res.get_json()
    cm = res.get_json()
    check("新增 53 筆", cm["inserted"] == 53, str(cm))
    check("主檔自動建立的筆數 = 待建數", cm["products_added"] == len(pv["missing_products"]))
    res = client.post("/api/master/import/commit", json={"batch_id": pv["batch_id"]})
    check("同一批不能重複確認（409）", res.status_code == 409)

    print("\n【4】訂單明細：分組、箱數、篩選面")
    res = client.get(f"/api/master/orders?line={LINE}&month={MONTH}")
    od = res.get_json()
    check("53 筆全部帶出", od["count"] == 53)
    dates = {d["date"]: d for d in od["facets"]["dates"]}
    check("日期篩選面有三天", set(dates) == {"2026-07-21", "2026-07-23", "2026-07-24"}, str(list(dates)))
    check("PO 篩選面有 6 張", len(od["facets"]["pos"]) == 6)
    row = next(r for r in od["rows"] if r["barcode"] == "6903148182406" and r["po_number"] == "13000000370676")
    check("62 袋 ÷ 24 = 2.5833 箱，不湊整、不四捨五入", abs(row["cases"] - 2.5833) < 1e-3, str(row["cases"]))
    check("箱入數來源標為主檔（匯入時已自動建）", row["box_source"] == "master")
    nobox = [r for r in od["rows"] if r["cases"] is None]
    check("沒有箱入數的兩筆算不出箱數（不是 0）", len(nobox) == 2, str([r["barcode"] for r in nobox]))
    d24 = dates["2026-07-24"]
    check("7/24 那天的箱數合計有算出來", d24["cases"] > 0 and d24["po_count"] >= 1, str(d24))
    res = client.get(f"/api/master/orders?line={LINE}&month={MONTH}&dates=2026-07-21")
    check("勾單一日期只回那天的品項", all(r["delivery_date"] == "2026-07-21" for r in res.get_json()["rows"])
          and res.get_json()["count"] > 0)
    res = client.get(f"/api/master/orders?line={LINE}&month={MONTH}&q=ARIEL")
    check("搜尋 ARIEL 只回 ARIEL", res.get_json()["count"] > 0 and
          all("ARIEL" in (r["brand"] + r["product_name"]).upper() for r in res.get_json()["rows"]))

    print("\n【5】就地編輯：出貨數量、備註、防互蓋")
    oid, ver = row["id"], row["version"]
    res = jput(client, f"/api/master/orders/{oid}", {"version": ver, "qty_ship": 48})
    check("改出貨數量成功", res.status_code == 200, str(res.get_json()))
    fresh = res.get_json()["row"]
    check("48 ÷ 24 = 2 箱", fresh["cases"] == 2.0, str(fresh["cases"]))
    check("立了人工調整旗標", fresh["qty_ship_overridden"] == 1)
    res = jput(client, f"/api/master/orders/{oid}", {"version": ver, "qty_ship": 50})
    check("用舊版本再存被擋（409 防互蓋）", res.status_code == 409)
    res = jput(client, f"/api/master/orders/{oid}", {"version": fresh["version"], "remarks": "不可超打"})
    check("改備註成功且立旗標", res.status_code == 200 and res.get_json()["row"]["remarks_overridden"] == 1)
    ver2 = res.get_json()["row"]["version"]
    res = jput(client, f"/api/master/orders/{oid}", {"version": ver2, "reset_qty_ship": True})
    check("恢復整合表數字：回到 62、旗標歸零",
          res.get_json()["row"]["qty_ship"] == 62 and res.get_json()["row"]["qty_ship_overridden"] == 0)
    res = jput(client, f"/api/master/orders/{oid}", {"version": res.get_json()["row"]["version"], "qty_ship": 48})
    assert res.status_code == 200

    print("\n【6】整張 PO 改交貨日")
    po = "13000000370676"
    before = client.get(f"/api/master/orders?line={LINE}&month={MONTH}&pos={po}").get_json()
    n_po = before["count"]
    res = jput(client, "/api/master/pos/date", {"line": LINE, "po_number": po,
                                               "delivery_date": "2026-07-30", "expected_date": "2026-07-24"})
    check("改期成功，品項全部一起搬", res.status_code == 200 and res.get_json()["moved"] == n_po,
          str(res.get_json()))
    after = client.get(f"/api/master/orders?line={LINE}&month={MONTH}&pos={po}").get_json()
    check("那張 PO 所有品項都在 7/30", all(r["delivery_date"] == "2026-07-30" for r in after["rows"]))
    check("日曆多出 7/30、7/24 箱數減少",
          "2026-07-30" in {d["date"] for d in after["facets"]["dates"]}
          and {d["date"]: d for d in after["facets"]["dates"]}["2026-07-24"]["cases"] < d24["cases"])
    res = jput(client, "/api/master/pos/date", {"line": LINE, "po_number": po,
                                               "delivery_date": "2026-07-31", "expected_date": "2026-07-24"})
    check("拿舊日期當依據再改被擋（409）", res.status_code == 409)
    res = jput(client, "/api/master/pos/date", {"line": LINE, "po_number": po, "delivery_date": "2026-08-05",
                                               "expected_date": "2026-07-30"})
    check("改到下個月：這個月的列表就看不到它", res.status_code == 200 and
          client.get(f"/api/master/orders?line={LINE}&month={MONTH}&pos={po}").get_json()["count"] == 0)
    res = jput(client, "/api/master/pos/date", {"line": LINE, "po_number": po, "reset": True,
                                               "expected_date": "2026-08-05"})
    check("恢復整合表日期：回到 7/24、旗標歸零", res.status_code == 200 and res.get_json()["delivery_date"] == "2026-07-24"
          and all(r["delivery_date_overridden"] == 0 for r in
                  client.get(f"/api/master/orders?line={LINE}&month={MONTH}&pos={po}").get_json()["rows"]))

    print("\n【7】再匯入同一份檔：人改過的不覆蓋，沒改過的照更新")
    res = upload(client, "/api/master/import/preview", ORDERS_XLSX)
    pv2 = res.get_json()
    check("第二次匯入：0 新增、全部相同（整合表沒變）", pv2["new_count"] == 0 and pv2["identical_count"] == 53, str({k: pv2[k] for k in ("new_count","updated_count","identical_count")}))
    check("主檔已齊全，沒有待建", pv2["missing_products"] == [])
    client.post("/api/master/import/commit", json={"batch_id": pv2["batch_id"]})
    row2 = next(r for r in client.get(f"/api/master/orders?line={LINE}&month={MONTH}").get_json()["rows"] if r["id"] == oid)
    check("人工改的出貨數量 48 沒被整合表的 62 蓋掉", row2["qty_ship"] == 48 and row2["qty_ship_overridden"] == 1)
    check("人工備註沒被蓋掉", row2["remarks"] == "不可超打")

    # 模擬酷澎改單：把整合表裡那筆的下單數量改成 200、交期改成 7/25，再匯一次
    wb = openpyxl.load_workbook(ORDERS_XLSX)
    ws = wb.active
    hdr = [c.value for c in ws[1]]
    ci = {h: i + 1 for i, h in enumerate(hdr)}
    target = None
    for r in range(2, ws.max_row + 1):
        if str(ws.cell(r, ci["PO單號"]).value) == po and str(ws.cell(r, ci["條碼(國條)"]).value) == "6903148182406":
            target = r; break
    ws.cell(target, ci["下單數量(酷澎單位)"]).value = 200
    ws.cell(target, ci["出貨數量"]).value = 200
    other = next(r for r in range(2, ws.max_row + 1) if str(ws.cell(r, ci["PO單號"]).value) == "13000000370675")
    ws.cell(other, ci["出貨數量"]).value = int(ws.cell(other, ci["出貨數量"]).value or 0) + 9
    changed = io.BytesIO(); wb.save(changed); changed.seek(0)
    res = client.post("/api/master/import/preview", data={"file": (changed, "變更後.xlsx")}, content_type="multipart/form-data")
    pv3 = res.get_json()
    check("酷澎改了兩筆 → 預覽列出 2 筆有變", pv3["updated_count"] == 2, str(pv3["updated_count"]))
    u = next(x for x in pv3["updated"] if x["barcode"] == "6903148182406")
    check("預覽標出這筆出貨數量是人工調整過的", u["qty_ship_overridden"] == 1)
    client.post("/api/master/import/commit", json={"batch_id": pv3["batch_id"]})
    rows3 = client.get(f"/api/master/orders?line={LINE}&month={MONTH}").get_json()["rows"]
    r_a = next(r for r in rows3 if r["id"] == oid)
    r_b = next(r for r in rows3 if r["po_number"] == "13000000370675" and r["qty_ship_overridden"] == 0 and r["qty_file_ship"] != r["qty_coupang"])
    check("人工調整過的那筆：下單數量更新成 200，出貨數量仍是人工的 48",
          r_a["qty_coupang"] == 200 and r_a["qty_ship"] == 48, f"{r_a['qty_coupang']} / {r_a['qty_ship']}")
    check("沒人工調整的那筆：出貨數量跟著整合表變", r_b["qty_ship"] == r_b["qty_file_ship"])

    print("\n【8】商品主檔：從總表匯入、手動維護、孤兒國條")
    res = upload(client, "/api/master/products/import", MASTER_XLSX, line=LINE)
    check("總表匯入主檔成功", res.status_code == 200, str(res.get_json()))
    pi = res.get_json()
    check("用標題找到 Sheet1 的國條與箱入數", pi["sheet"] == "Sheet1" and pi["added"] + pi["updated"] + pi["unchanged"] > 60, str(pi))
    prods = client.get(f"/api/master/products?line={LINE}").get_json()
    p = next((x for x in prods["products"] if x["barcode"] == "4987176340894"), None)
    check("總表第一列 ARIEL 4987176340894 箱入數 6 進了主檔", p is not None and p["box_size"] == 6, str(p and p["box_size"]))
    check("總表 Note（新品／不可超打）帶進備註", any(x["note"] in ("新品", "不可超打") for x in prods["products"]))
    res = client.post("/api/master/products", json={"line": LINE, "barcode": "6903148355923", "box_size": 16, "product_name": "好自在 無痕早安褲 XL"})
    check("手動補上新品的箱入數", res.status_code == 200 and res.get_json()["product"]["box_size"] == 16)
    rows4 = client.get(f"/api/master/orders?line={LINE}&month={MONTH}").get_json()["rows"]
    r_new = next(r for r in rows4 if r["barcode"] == "6903148355923")
    check("補完主檔，原本算不出的那筆現在有箱數了", r_new["cases"] is not None and r_new["box_source"] == "master", str(r_new["cases"]))
    res = client.post("/api/master/products", json={"line": LINE, "barcode": "123", "box_size": 0})
    check("箱入數 0 被擋", res.status_code == 400)
    check("孤兒清單：主檔都有了所以是空的", client.get(f"/api/master/products?line={LINE}").get_json()["orphans"] == [])

    print("\n【9】配額與總表")
    s = client.get(f"/api/master/summary?line={LINE}&month={MONTH}").get_json()
    check("總表日期欄 = 這個月有出貨的交貨日", s["dates"] == ["2026-07-21", "2026-07-23", "2026-07-24"], str(s["dates"]))
    r_sum = next(r for r in s["rows"] if r["barcode"] == "6903148182406")
    check("國條 6903148182406 在 7/24 的箱數 = 2（48÷24）", r_sum["by_date"].get("2026-07-24") == 2.0, str(r_sum["by_date"]))
    check("月加總 = 各日期加總", abs(r_sum["month_total"] - sum(r_sum["by_date"].values())) < 1e-6)
    check("沒設配額時剩餘顯示空（不是 0）", r_sum["quota"] is None and r_sum["remaining"] is None)
    total_before = s["month_total"]
    check("整月合計 = 每日合計相加", abs(total_before - sum(s["totals_by_date"].values())) < 1e-6)
    res = jput(client, "/api/master/quota", {"line": LINE, "barcode": "6903148182406", "month": MONTH, "qty_cases": 350})
    check("設配額 350", res.status_code == 200)
    s2 = client.get(f"/api/master/summary?line={LINE}&month={MONTH}").get_json()
    r_sum2 = next(r for r in s2["rows"] if r["barcode"] == "6903148182406")
    check("剩餘可供貨 = 350 − 2 = 348（總表 BE = O − BD）", r_sum2["remaining"] == 348.0, str(r_sum2["remaining"]))
    res = jput(client, f"/api/master/orders/{oid}", {"version": r_a["version"], "qty_ship": 240})
    s3 = client.get(f"/api/master/summary?line={LINE}&month={MONTH}").get_json()
    r_sum3 = next(r for r in s3["rows"] if r["barcode"] == "6903148182406")
    check("改出貨數量後總表立刻重算：240÷24=10，剩餘 340（不用重上傳）",
          r_sum3["by_date"]["2026-07-24"] == 10.0 and r_sum3["remaining"] == 340.0, str(r_sum3["remaining"]))
    res = client.post("/api/master/quota/import", data={"line": LINE, "month": MONTH, "text": "4987176340894 354\n6903148182406\t400\n"})
    check("貼文字匯入配額", res.status_code == 200 and res.get_json()["written"] == 2, str(res.get_json()))
    res = upload(client, "/api/master/quota/import", MASTER_XLSX, line=LINE, month=MONTH)
    qc = res.get_json()
    check("上傳總表沒指定欄位 → 先回可選欄位", qc.get("need_column") and "Supply_CS (Jul)" in qc["columns"], str(qc.get("columns", []))[:120])
    res = upload(client, "/api/master/quota/import", MASTER_XLSX, line=LINE, month=MONTH, column="Supply_CS (Jul)")
    check("指定 Supply_CS (Jul) 欄寫入配額", res.status_code == 200 and res.get_json()["written"] > 20, str(res.get_json()))
    s4 = client.get(f"/api/master/summary?line={LINE}&month={MONTH}").get_json()
    r_ariel = next(r for r in s4["rows"] if r["barcode"] == "4987176340894")
    check("ARIEL 的 7 月配額 = 總表 M 欄 354", r_ariel["quota"] == 354.0, str(r_ariel["quota"]))
    res = client.get(f"/api/master/summary?line=紙潔&month={MONTH}").get_json()
    check("換線別：紙潔沒有資料、總表是空的、不會撈到寶僑", res["rows"] == [] and res["month_total"] == 0)

    print("\n【10】匯出 Excel")
    res = client.get(f"/api/master/export?line={LINE}&month={MONTH}")
    check("匯出成功", res.status_code == 200 and res.mimetype.endswith("sheet"))
    wb = openpyxl.load_workbook(io.BytesIO(res.data))
    check("三個工作表：總表、每日合計、訂單明細", wb.sheetnames == ["總表", "每日合計", "訂單明細"], str(wb.sheetnames))
    ws = wb["總表"]; hdr = [c.value for c in ws[1]]
    check("總表日期欄用「7/24交貨」寫法，最右邊是 TTL／配額／剩餘",
          "7/24交貨" in hdr and hdr[-4:-1] == ["7月TTL下單總箱數", "7月配額(CS)", "剩餘可供貨量(CS)"], str(hdr[-5:]))
    bc_col = hdr.index("國條"); rem_col = hdr.index("剩餘可供貨量(CS)"); d24_col = hdr.index("7/24交貨")
    r_x = next(r for r in ws.iter_rows(min_row=2, values_only=True) if str(r[bc_col]) == "6903148182406")
    check("匯出的數字是值不是公式：6903148182406 7/24 = 10、剩餘 = 390", r_x[d24_col] == 10 and r_x[rem_col] == 390, f"{r_x[d24_col]} / {r_x[rem_col]}")
    ws3 = wb["訂單明細"]
    check("訂單明細有 53 列", ws3.max_row - 1 == 53, str(ws3.max_row - 1))

    print("\n【10b】匯出專案報價檔格式（一天一個分頁）")
    res = client.get(f"/api/master/export/daily?line={LINE}&month={MONTH}")
    check("匯出成功", res.status_code == 200)
    wbd = openpyxl.load_workbook(io.BytesIO(res.data))
    check("分頁 = 有出貨的交貨日，命名像 0724交貨，最新在前",
          wbd.sheetnames == ["0724交貨", "0723交貨", "0721交貨"], str(wbd.sheetnames))
    wsd = wbd["0724交貨"]
    hdrd = [c.value for c in wsd[1]]
    check("欄位順序照報價檔（A SKU ID … J 出貨數量(箱) … R 交貨日）",
          hdrd[0] == "SKU ID" and hdrd[2] == "國條" and hdrd[9] == "出貨數量(箱)" and hdrd[17] == "交貨日", str(hdrd[:4]))
    r_cells = [wsd.cell(r, 18).value for r in range(2, wsd.max_row + 1)]
    po_cells = [v for v in r_cells if v]
    check("R 欄只在每張 PO 第一列寫「PO_日期(倉別)」，7/24 有 2 張 PO",
          len(po_cells) == 2 and all("_7/24交貨(" in v for v in po_cells), str(po_cells))
    yellow_rows = [r for r in range(2, wsd.max_row + 1)
                   if wsd.cell(r, 1).fill.fgColor.rgb in ("00FFFF00", "FFFFFF00") and wsd.cell(r, 1).value is None]
    check("PO 之間有一列黃色空白列隔開", len(yellow_rows) == 1, str(yellow_rows))
    row_x = next(r for r in wsd.iter_rows(min_row=2, values_only=True) if str(r[2]) == "6903148182406")
    check("出貨數量(箱) 是值：240 ÷ 24 = 10，總計 = 單價×箱入數×箱數", row_x[9] == 10 and row_x[13] == row_x[10] * row_x[8] * 10, str(row_x[7:14]))

    print("\n【11】歷程")
    logs = client.get(f"/api/master/logs?line={LINE}&po={po}").get_json()["logs"]
    check("PO 改期、出貨數量調整都留下歷程", any(l["field"] == "delivery_date" for l in logs) and any(l["field"] == "qty_ship" for l in logs))
    check("歷程分得出手動與匯入", {l["source"] for l in logs} >= {"manual", "import"})

    print(f"\n通過 {len(PASS)} 項，失敗 {len(FAIL)} 項")
    if FAIL:
        print("失敗項目："); [print("  -", f) for f in FAIL]
        sys.exit(1)


if __name__ == "__main__":
    main()

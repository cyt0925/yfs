"""蝦皮特選寄倉訂單系統（/shopee）端到端測試。
第一段：上傳採購單／入庫單、重傳差異、版本、確認、改履約方式、總覽。
第二段：價格本、竹運採購進貨／供應商直送產 EIP 採購單＋採購下採明細、除不盡改數量、回填 EIP 單號。
瑪氏直送：瑪氏 EIP 採購單（單位數量照蝦皮、依品類／單位／中標分檔）、回填後產瑪氏制式採購單、蝦皮倉資料。
竹運出貨：EC 庫存＋竹運效期表＋商品主檔 → 效期判斷、逐列扣庫存、批次與人工出貨量 → 訂單系統拋單檔＋下修刪單清單。

測試檔照蝦皮後台下載檔的表頭用程式造（真的業務檔不放進 repo）。
執行：python test_shopee.py（設了 DATABASE_URL 就跑 PostgreSQL）
"""
import io
import os
import sys
import tempfile
import zipfile

import openpyxl
import xlrd

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)

_tmp = tempfile.mkdtemp(prefix="oms_shopee_test_")
import db  # noqa: E402
db.DATA_DIR = os.path.join(_tmp, "資料與設定")
db.DB_PATH = os.path.join(db.DATA_DIR, "test.db")
db.BACKUP_DIR = os.path.join(db.DATA_DIR, "backups")

import app as app_module  # noqa: E402
from shopee.common import line_of, split_supplier_sku  # noqa: E402
from shopee.parse import ParseError, parse_file  # noqa: E402

PASS, FAIL = [], []


def check(name, cond, detail=""):
    (PASS if cond else FAIL).append(name)
    print(("  ✓ " if cond else "  ✗ ") + name + ("" if cond or not detail else f"  — {detail}"))


PO_HDR = ["PR ID", "PO ID", "Is Jit", "Warehouse", "PR Creation Date", "PO Creation Date", "Expected Delivery Time", "Shopee SKU ID",
          "Supplier SKU ID", "Shopee SKU Name", "Shopee Requested Qty (unit)", "Shopee Requested Qty (pcs)", "Unit Name",
          "Shopee Request Price (before-tax)", "Shopee Request Price (after-tax)", "Tax %", "Shopee Requested Value (before-tax)",
          "Shopee Requested Value (after-tax)", "Supplier Confirmed Qty (unit)", "Supplier Confirmed Qty (pcs)",
          "Supplier Confirmed Price (before-tax)", "Supplier Confirmed Price (after-tax)", "Supplier Confirmed Value (after-tax)",
          "Brand", "EAN / UPC", "EAN / UPC2", "MOQ", "Sourcing Status", "Next Available Date", "Selling Type", "Inbound ID"]
IN_HDR = ["Inbound ID", "is_jit", "Shopee SKU ID", "Supplier SKU ID", "SKU Name", "PO ID", "Estimated purchase date", "Actual purchase date",
          "Estimated purchase quantity", "Warehouse acceptance quantity", "ASN Status", "Is the Payment Requested", "Warehouse",
          "Selling Type", "Expected Qty (Unit)", "Expected Qty (Pcs)", "Shortage (Pcs)", "Surplus (Pcs)", "Damaged (Pcs)",
          "Received (Pcs)", "Arrived (Pcs)", "Unit Name", "Purchase Type", "Batch ID", "Production Date", "Expiration Date", "Lifecycle"]


def purchase_xlsx(rows, pr="PRTWCOTWA202610010001", po=None, wh="TWA", date="2026-10-19"):
    """rows = [(Shopee SKU ID, Supplier SKU ID, 品名, 數量, Selling Type)]"""
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Sheet1"; ws.append(PO_HDR)
    for sku, sup, name, qty, st in rows:
        r = dict.fromkeys(PO_HDR)
        r.update({"PR ID": pr, "PO ID": po, "Is Jit": "No", "Warehouse": wh, "PR Creation Date": "2026-10-01 13:16",
                  "Expected Delivery Time": date, "Shopee SKU ID": sku, "Supplier SKU ID": sup, "Shopee SKU Name": name,
                  "Shopee Requested Qty (unit)": str(qty), "Shopee Requested Qty (pcs)": str(qty), "Unit Name": "Pieces",
                  "Tax %": "5", "Brand": "NON-BRAND", "EAN / UPC": sup.split("_")[0], "MOQ": "1", "Sourcing Status": "Available",
                  "Selling Type": st})
        ws.append([r[h] for h in PO_HDR])
    buf = io.BytesIO(); wb.save(buf)
    return buf.getvalue()


def inbound_xlsx(rows, po, wh="TWA", date="2026-10-19"):
    """rows = [(Inbound ID, Shopee SKU ID, Supplier SKU ID, 品名, 數量)]"""
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "ASN Detail"; ws.append(IN_HDR)
    for inb, sku, sup, name, qty in rows:
        r = dict.fromkeys(IN_HDR, "/")
        r.update({"Inbound ID": inb, "is_jit": "No", "Shopee SKU ID": sku, "Supplier SKU ID": sup, "SKU Name": name, "PO ID": po,
                  "Estimated purchase date": date, "Actual purchase date": "", "Estimated purchase quantity": qty,
                  "Warehouse acceptance quantity": 0, "ASN Status": "Pending Delivery", "Is the Payment Requested": "No",
                  "Warehouse": wh, "Selling Type": "Pcs", "Expected Qty (Unit)": qty, "Expected Qty (Pcs)": qty,
                  "Unit Name": "Pieces", "Purchase Type": "Consignment"})
        ws.append([r[h] for h in IN_HDR])
    buf = io.BytesIO(); wb.save(buf)
    return buf.getvalue()


def price_xlsx(sheets):
    """價格本：{工作表名: [(料號3, 單位, 品名, 箱入數)]}，表頭照 PG價格本整合／紙潔價格本整合。"""
    wb = openpyxl.Workbook(); wb.remove(wb.active)
    for title, rows in sheets.items():
        ws = wb.create_sheet(title); ws.append(["料號", "單位", "品名", "料號3", "箱入數", "進價"])
        for code3, unit, name, box in rows:
            ws.append([code3[:-1], unit, name, code3, box, 100])
    buf = io.BytesIO(); wb.save(buf)
    return buf.getvalue()


def xls_rows(data):
    sh = xlrd.open_workbook(file_contents=data).sheet_by_index(0)
    head = [sh.cell_value(0, c) for c in range(10)]
    rows = [[sh.cell_value(r, c) for c in range(10)] for r in range(1, sh.nrows) if sh.cell_value(r, 1) != ""]
    return head, rows


PG_ROWS = [("41511463369_285962283475", "4987176232878_CNN", "幫寶適 清新幫 拉拉褲(XXL)26片", 240, "Pcs"),
           ("50403302885_340318306491", "4987176232878_RNN", "幫寶適 清新幫 拉拉褲 箱購(XXL)26片X4包", 26, "Carton"),
           ("43811463251_265962855996", "6903148358214_CNN", "幫寶適 極上守護 拉拉褲(L+)30片", 216, "Pcs")]


def main():
    db.init_db()
    app_module.app.config["TESTING"] = True
    c = app_module.app.test_client()
    with c.session_transaction() as s:
        s["user"] = "Jerry"

    def preview(data, name):
        return c.post("/api/shopee/import/preview", data={"file": (io.BytesIO(data), name)}, content_type="multipart/form-data")

    def orders(**q):
        q.setdefault("from", "2026-01-01")
        return c.get("/api/shopee/orders", query_string=q).get_json()

    print("\n【0】頁面、首頁卡片、選單")
    page = c.get("/shopee").get_data(as_text=True)
    check("蝦皮特選頁打得開、蝦皮橘、有拖曳框和訂單總覽", "蝦皮特選訂單管理系統" in page and "#ee4d2d" in page and 'id="dz"' in page and 'id="tbl"' in page)
    home = c.get("/").get_data(as_text=True)
    check("首頁左側選單的蝦皮特選連到 /shopee", 'href="/shopee"' in home)
    ch = next(x for x in c.get("/api/portal/summary").get_json()["channels"] if x["key"] == "shopee")
    check("首頁蝦皮卡片：上線、連 /shopee、還沒資料時數字是 0、沒有箱數", ch["live"] and ch["url"] == "/shopee" and ch["new_today"] == 0 and ch["month_boxes"] is None, str(ch))

    print("\n【1】讀檔：認檔案種類、線別、單位代碼")
    check("料號拆底線：國條、瑪氏帶 -1、沒有底線", split_supplier_sku("4987176232878_CNN ") == ("4987176232878", "C")
          and split_supplier_sku("M60002536-1_CMN") == ("M60002536-1", "C") and split_supplier_sku("ABC") == ("ABC", ""))
    check("線別：M 開頭瑪氏、7 碼紙潔、13 碼與 14 碼寶僑、其他認不出", line_of("M10273409") == "瑪氏" and line_of("1022257") == "紙潔"
          and line_of("4987176232878") == "寶僑" and line_of("14987176176172") == "寶僑" and line_of("DH56M") == "")
    d = parse_file(purchase_xlsx(PG_ROWS), "PurchaseOrder_x.xlsx")
    check("採購單：看表頭認得、3 列、單位代碼 C＝包 R＝箱、到貨日、數量轉數字", d["kind"] == "purchase" and len(d["rows"]) == 3
          and [r["unit_name"] for r in d["rows"]] == ["包", "箱", "包"] and d["rows"][0]["expected_date"] == "2026-10-19"
          and d["rows"][0]["qty"] == 240 and d["lines"] == {"寶僑": 3}, str(d["lines"]))
    mars = parse_file(purchase_xlsx([("1_1", "M10273409_QMN", "M&M'S", 48, "Pcs"), ("1_2", "DH56M_QMN", "士力架 5入", 45, "Pcs")]), "x.xlsx")
    check("認不出線別的（DH56M）用同檔多數補、還提醒", [r["line"] for r in mars["rows"]] == ["瑪氏", "瑪氏"] and any("DH56M" in e for e in mars["errors"]), str(mars["errors"]))
    ind = parse_file(inbound_xlsx([("IN1", "s1", "1022257_PNN", "橘子工坊", 30), ("IN2", "s1", "1022257_PNN", "橘子工坊", 18)], "POX"), "InboundOrder_x.xlsx")
    check("入庫單：同一張 PO 同一品項兩列（雙效期）合成一列，數量加總、入庫單號串起來", ind["kind"] == "inbound" and len(ind["rows"]) == 1
          and ind["rows"][0]["qty"] == 48 and ind["rows"][0]["inbound_id"] == "IN1,IN2" and ind["rows"][0]["line"] == "紙潔", str(ind["rows"]))
    wb = openpyxl.Workbook(); wb.active.append(["隨便", "一個", "表"]); b = io.BytesIO(); wb.save(b)
    for name, data in (("不是蝦皮的表", b.getvalue()), ("壞掉的檔", b"not excel")):
        try:
            parse_file(data, "x.xlsx"); ok = False
        except ParseError as e:
            ok = "Excel" in str(e) or "表頭" in str(e)
        check(f"{name} → 講清楚、不出現英文錯誤", ok)
    r = preview(b"not excel", "x.xlsx")
    check("上傳壞檔 → 400", r.status_code == 400 and "Traceback" not in r.get_data(as_text=True))

    print("\n【2】第一次上傳採購單（PR 階段，還沒有 PO）")
    raw = purchase_xlsx(PG_ROWS)
    pv = preview(raw, "PurchaseOrder_20261001.xlsx").get_json()
    check("預覽：採購單、新增 3、要選履約方式、還沒寫進訂單", pv["kind"] == "purchase" and pv["counts"]["new"] == 3 and pv["needs_fulfil"]
          and orders()["summary"]["orders"] == 0, str(pv["counts"]))
    r = c.post("/api/shopee/import/commit", json={"upload_id": pv["upload_id"]})
    check("沒選履約方式 → 400", r.status_code == 400 and "履約方式" in r.get_json()["error"])
    r = c.post("/api/shopee/import/commit", json={"upload_id": pv["upload_id"], "fulfil": "竹運採購進貨"})
    check("選了竹運採購進貨 → 匯入 3 項", r.status_code == 200 and r.get_json()["counts"]["new"] == 3, str(r.get_json()))
    check("同一個預覽再按一次 → 404（不會重複匯入）", c.post("/api/shopee/import/commit", json={"upload_id": pv["upload_id"], "fulfil": "竹運出貨"}).status_code == 404)
    o = orders()
    g = o["orders"][0]
    check("總覽：一張單、訂單編號先用 PR、線別寶僑、倉 TWA、到貨日、履約方式、原始訂購量 482", len(o["orders"]) == 1 and g["order_no"] == "PRTWCOTWA202610010001"
          and g["line"] == "寶僑" and g["warehouse"] == "TWA" and g["expected_date"] == "2026-10-19" and g["fulfil"] == "竹運採購進貨"
          and g["qty_original"] == 482 and g["status"] == "已匯入", str(g)[:300])
    ups = c.get("/api/shopee/uploads").get_json()["uploads"]
    check("上傳紀錄：檔名、上傳人員、匯入結果", ups[0]["filename"] == "PurchaseOrder_20261001.xlsx" and ups[0]["uploaded_by"] == "Jerry" and ups[0]["new_count"] == 3)
    r = c.get(f"/api/shopee/uploads/{ups[0]['id']}/file")
    check("原始檔下載回來跟上傳的一模一樣", r.status_code == 200 and r.data == raw)
    logs = c.get("/api/master/logs?scope=shopee&limit=50").get_json()["logs"]
    check("歷程視窗（本頁）看得到這次匯入，系統欄是蝦皮特選訂單管理系統", any(l["field"] == "shopee_import" for l in logs)
          and all(l.get("system") in (None, "蝦皮特選訂單管理系統") for l in logs), str([l["field"] for l in logs]))

    print("\n【3】同一份再傳一次：內容完全一樣")
    pv = preview(raw, "PurchaseOrder_20261001.xlsx").get_json()
    check("預覽：3 項無異動、沒有差異、不用選履約方式", pv["counts"]["same"] == 3 and not pv["diffs"] and not pv["needs_fulfil"], str(pv["counts"]))
    r = c.post("/api/shopee/import/commit", json={"upload_id": pv["upload_id"]})
    check("確認後只記上傳紀錄，訂單還是已匯入", r.status_code == 200 and orders()["orders"][0]["status"] == "已匯入"
          and len(c.get("/api/shopee/uploads").get_json()["uploads"]) == 2)

    print("\n【4】重傳：變成 PO、數量改、到貨日改、少一個品項")
    rows2 = [PG_ROWS[0][:3] + (200, "Pcs"), PG_ROWS[1]]           # 第一項 240→200、第三項不見了
    pv = preview(purchase_xlsx(rows2, po="POTWCTWA26100514L", date="2026-10-20"), "PurchaseOrder_20261005.xlsx").get_json()
    kinds = sorted((x["label"], x["old"], x["new"]) for x in pv["diffs"])
    check("預覽列出差異：數量 240→200（-40）、到貨日 ×2、品項移除；PR 變 PO 不算差異", ("數量", 240, 200) in kinds
          and sum(1 for k in kinds if k[0] == "到貨日") == 2 and any(k[0] == "品項" for k in kinds) and not any(k[0] == "PO 號碼" for k in kinds)
          and any(x["qty_delta"] == -40 for x in pv["diffs"]) and pv["counts"]["remove"] == 1, str(kinds))
    c.post("/api/shopee/import/commit", json={"upload_id": pv["upload_id"]})
    o = orders(); g = o["orders"][0]
    check("總覽：訂單編號換成 PO、PR 還留著、匯入差異待確認、目前數量 226、原始訂購量還是 482、移除 1 項",
          g["order_no"] == "POTWCTWA26100514L" and g["pr_id"] == "PRTWCOTWA202610010001" and g["status"] == "匯入差異待確認"
          and g["qty"] == 226 and g["qty_original"] == 482 and g["removed_count"] == 1 and o["pending_orders"] == 1, str(g)[:300])
    it = next(i for i in g["items"] if i["supplier_sku_id"] == "4987176232878_CNN")
    vs = c.get(f"/api/shopee/versions/{it['id']}").get_json()["versions"]
    check("版本紀錄：舊版不刪，第 1、2 版都在（內容一樣的那次重傳不另記一版），第 1 版數量 240、最新版 200", [v["version"] for v in vs] == [2, 1]
          and vs[-1]["data"]["qty"] == 240 and vs[0]["data"]["qty"] == 200, str([(v["version"], v["data"]["qty"]) for v in vs]))
    dd = c.get("/api/shopee/diffs", query_string={"order_no": "POTWCTWA26100514L"}).get_json()["diffs"]
    check("差異視窗：訂單編號、商品料號、欄位、上一版、新版、數量增減、上傳人員、時間、待確認", dd and all(
        x["order_no"] == "POTWCTWA26100514L" and x["supplier_sku_id"] and x["field_label"] and x["uploaded_by"] == "Jerry" and x["uploaded_at"]
        and not x["confirmed"] for x in dd), str(dd[0]) if dd else "")
    t = c.get("/api/portal/summary").get_json()["tasks"]
    check("首頁今日事項出現「蝦皮特選重傳後有差異 1 張單」", any(x["key"] == "shopee_diff" and x["count"] == 1 and x["url"] == "/shopee" for x in t))
    check("篩「只看差異待確認」只剩這張", [x["order_no"] for x in orders(pending="1")["orders"]] == ["POTWCTWA26100514L"])

    print("\n【5】確認差異")
    r = c.post("/api/shopee/diffs/confirm", json={"order_nos": ["POTWCTWA26100514L"]})
    g = orders()["orders"][0]
    dd = c.get("/api/shopee/diffs", query_string={"order_no": "POTWCTWA26100514L"}).get_json()["diffs"]
    check("確認後回到已匯入、差異記下確認人和時間", r.status_code == 200 and g["status"] == "已匯入" and all(x["confirmed"] and x["confirmed_by"] == "Jerry" for x in dd))
    check("首頁今日事項的蝦皮差異消失", not any(x["key"] == "shopee_diff" for x in c.get("/api/portal/summary").get_json()["tasks"]))
    check("沒選訂單就按確認 → 400", c.post("/api/shopee/diffs/confirm", json={"order_nos": []}).status_code == 400)

    print("\n【6】入庫單：補入庫單號、數量不一樣算差異；只有入庫單的 PO 直接建")
    inb = inbound_xlsx([("INTWA0002610050005", PG_ROWS[0][0], PG_ROWS[0][1], PG_ROWS[0][2], 190),
                        ("INTWA0002610050006", PG_ROWS[1][0], PG_ROWS[1][1], PG_ROWS[1][2], 26)], "POTWCTWA26100514L", date="2026-10-20")
    pv = preview(inb, "InboundOrder_20261006.xlsx").get_json()
    check("入庫單預覽：找到採購單那兩筆、數量 200→190 算差異、不用選履約方式", pv["kind"] == "inbound" and pv["counts"]["new"] == 0
          and [(x["label"], x["old"], x["new"]) for x in pv["diffs"]] == [("數量", 200, 190)] and not pv["needs_fulfil"], str(pv["diffs"]))
    c.post("/api/shopee/import/commit", json={"upload_id": pv["upload_id"]})
    g = orders()["orders"][0]
    it = next(i for i in g["items"] if i["supplier_sku_id"] == "4987176232878_CNN")
    check("補上入庫單號、入庫單數量 190，又變差異待確認", it["inbound_ids"] == "INTWA0002610050005" and it["inbound_qty"] == 190 and g["status"] == "匯入差異待確認", str(it)[:200])
    c.post("/api/shopee/diffs/confirm", json={"order_nos": ["POTWCTWA26100514L"]})
    only = inbound_xlsx([("INTWA0002610050266", "56402432052_307219345020", "M60003196_QMN", "希寶 誘惑泥 55入", 48),
                         ("INTWA0002610050265", "53902436305_360235791151", "M10281171_RMN", "偉嘉 貓罐頭 箱購", 18)], "POTWCTWA26100514N", date="2026-10-13")
    pv = preview(only, "InboundOrder_瑪氏.xlsx").get_json()
    check("只有入庫單、找不到採購單 → 新增 2 項、要選履約方式", pv["counts"]["new"] == 2 and pv["needs_fulfil"])
    c.post("/api/shopee/import/commit", json={"upload_id": pv["upload_id"], "fulfil": "供應商直送"})
    g2 = next(x for x in orders()["orders"] if x["order_no"] == "POTWCTWA26100514N")
    check("建成一張 PO：瑪氏、供應商直送、2 項、單位 盒／箱", g2["line"] == "瑪氏" and g2["fulfil"] == "供應商直送" and g2["item_count"] == 2
          and sorted(i["unit_name"] for i in g2["items"]) == ["盒", "箱"])
    pr_only = purchase_xlsx([("70000000001_1", "1022257_PNN", "橘子工坊浴廁清潔劑", 360, "Pcs")], pr="PRTWCOTWX202610010777", wh="TWX", date="2026-10-21")
    c.post("/api/shopee/import/commit", json={"upload_id": preview(pr_only, "PurchaseOrder_紙潔.xlsx").get_json()["upload_id"], "fulfil": "竹運出貨"})
    pv = preview(inbound_xlsx([("INTWX1", "70000000001_1", "1022257_PNN", "橘子工坊浴廁清潔劑", 360)], "POTWCTWX26100599A", wh="TWX", date="2026-10-21"), "InboundOrder_紙潔.xlsx").get_json()
    check("採購單還在 PR 階段時，入庫單用「同商品、同倉、同到貨日」對回去，不另外建", pv["counts"]["new"] == 0 and not pv["diffs"], str(pv["counts"]))
    c.post("/api/shopee/import/commit", json={"upload_id": pv["upload_id"]})
    g3 = next(x for x in orders()["orders"] if x["pr_id"] == "PRTWCOTWX202610010777")
    check("那張 PR 補上 PO 號碼、入庫單號，線別紙潔", g3["order_no"] == "POTWCTWX26100599A" and g3["items"][0]["inbound_ids"] == "INTWX1" and g3["line"] == "紙潔")

    print("\n【7】改履約方式：要原因、留紀錄")
    bad = c.post("/api/shopee/fulfil", json={"order_no": "POTWCTWA26100514N", "fulfil": "竹運出貨", "reason": ""})
    check("沒寫原因 → 400", bad.status_code == 400 and "原因" in bad.get_json()["error"])
    check("亂填履約方式 → 400", c.post("/api/shopee/fulfil", json={"order_no": "POTWCTWA26100514N", "fulfil": "宅配", "reason": "x"}).status_code == 400)
    r = c.post("/api/shopee/fulfil", json={"order_no": "POTWCTWA26100514N", "fulfil": "竹運出貨", "reason": "竹運有現貨"})
    g2 = next(x for x in orders()["orders"] if x["order_no"] == "POTWCTWA26100514N")
    logs = [l for l in c.get("/api/master/logs?scope=shopee&limit=100").get_json()["logs"] if l["field"] == "shopee_fulfil"]
    check("改成竹運出貨，歷程記下原本、修改後、修改人員、時間、原因", r.get_json()["changed"] and g2["fulfil"] == "竹運出貨" and logs
          and logs[0]["old_value"] == "供應商直送" and logs[0]["new_value"] == "竹運出貨" and logs[0]["reason"] == "竹運有現貨"
          and logs[0]["operator"] == "Jerry" and logs[0]["changed_at"], str(logs[:1]))
    check("找不到的單 → 404", c.post("/api/shopee/fulfil", json={"order_no": "NOPE", "fulfil": "竹運出貨", "reason": "x"}).status_code == 404)

    print("\n【8】總覽篩選、首頁數字")
    check("篩線別：瑪氏只剩一張", [x["order_no"] for x in orders(line="瑪氏")["orders"]] == ["POTWCTWA26100514N"])
    check("篩履約方式：竹運採購進貨只剩寶僑那張", [x["order_no"] for x in orders(fulfil="竹運採購進貨")["orders"]] == ["POTWCTWA26100514L"])
    check("搜尋入庫單號找得到那張", [x["order_no"] for x in orders(q="INTWX1")["orders"]] == ["POTWCTWX26100599A"])
    check("到貨日區間：10/13 只剩瑪氏那張", [x["order_no"] for x in orders(**{"from": "2026-10-13", "to": "2026-10-13"})["orders"]] == ["POTWCTWA26100514N"])
    ch = next(x for x in c.get("/api/portal/summary").get_json()["channels"] if x["key"] == "shopee")
    check("首頁蝦皮卡片：今日新進 3 張、有最後匯入時間", ch["new_today"] == 3 and ch["last_import"], str(ch))
    recent = c.get("/api/portal/summary").get_json()["recent"]
    check("首頁最近操作紀錄標出蝦皮特選訂單管理系統", any(x.get("system") == "蝦皮特選訂單管理系統" for x in recent), str([x.get("system") for x in recent]))
    r = c.post("/api/shopee/import/cancel", json={"upload_id": preview(raw, "x.xlsx").get_json()["upload_id"]})
    check("預覽後按取消 → 不寫訂單、不進上傳紀錄", r.status_code == 200 and len(c.get("/api/shopee/uploads").get_json()["uploads"]) == 7,
          str(len(c.get("/api/shopee/uploads").get_json()["uploads"])))

    stage2(c, preview, orders)

    print("\n" + "=" * 62)
    print(f"通過 {len(PASS)} 項，失敗 {len(FAIL)} 項")
    if FAIL:
        print("失敗項目："); [print("  -", n) for n in FAIL]
        sys.exit(1)


def stage2(c, preview, orders):
    def up(url, data, name):
        return c.post(url, data={"file": (io.BytesIO(data), name)}, content_type="multipart/form-data")

    def imp(data, fulfil):
        return c.post("/api/shopee/import/commit", json={"upload_id": preview(data, "PurchaseOrder_t.xlsx").get_json()["upload_id"], "fulfil": fulfil}).get_json()

    def plan(fulfil):
        return {g["order_no"]: g for g in c.get("/api/shopee/purchase/plan", query_string={"fulfil": fulfil}).get_json()["orders"]}

    def gen(fulfil, nos):
        return c.post("/api/shopee/purchase/generate", json={"fulfil": fulfil, "order_nos": nos})

    def item(g, sup):
        return next(p for p in g["items"] if p["supplier_sku_id"] == sup)

    print("\n【9】價格本（轉換率）")
    page = c.get("/shopee").get_data(as_text=True)
    check("頁籤：訂單匯入與總覽、竹運採購進貨、直送訂單、採購單產出；價格本上傳、修改數量視窗", all(x in page for x in (
        'data-tab="orders"', 'data-f="竹運採購進貨"', 'data-f="供應商直送"', 'data-tab="eip"', 'id="dz-conv"', 'id="b-tbl"', 'id="e-tbl"', 'id="dlg-adj"')))
    pg = price_xlsx({"價格本整合": [("4987176232878C", "包", "幫寶適 清新幫 拉拉褲", 4), ("4987176232878R", "箱", "幫寶適 清新幫 拉拉褲 箱購", 1),
                                   ("6903148358214C", "包", "幫寶適 極上守護 拉拉褲", 3)]})
    r = up("/api/shopee/conv/import", pg, "PG價格本整合.xlsx").get_json()
    check("上傳寶僑價格本：寶僑 3 筆", r["ok"] and r["lines"] == {"寶僑": {"count": 3, "before": 0, "blank": 0}}, str(r))
    paper = price_xlsx({"紙品價格本": [("1099999P", "包", "紙品 沒填箱入數", None)],
                        "潔品價格本": [("1022257P", "瓶", "橘子工坊浴廁清潔劑", 12), ("1022248C", "包", "橘子工坊洗碗精補充包", 12),
                                      ("1021315R", "箱", "橘子工坊洗衣精補充包 箱購", 1), ("AAA0001C", "包", "組合品", 6)],
                        "工作表4": [("1022257P", "瓶", "兩頁合起來的，不另外讀", 99)]})
    r = up("/api/shopee/conv/import", paper, "紙潔價格本整合.xlsx").get_json()
    check("上傳紙潔價格本：組合品料號（AAA…）也算紙潔、「工作表4」不重複讀、箱入數空白 1 筆", r["lines"] == {"紙潔": {"count": 5, "before": 0, "blank": 1}}
          and not r["warnings"], str(r))
    st = c.get("/api/shopee/conv").get_json()["lines"]
    check("價格本狀態：兩個線別、筆數、檔名、上傳人員", st["寶僑"]["count"] == 3 and st["紙潔"]["count"] == 5 and st["紙潔"]["file"] == "紙潔價格本整合.xlsx"
          and st["寶僑"]["by"] == "Jerry", str(st))
    r = up("/api/shopee/conv/import", pg, "PG價格本整合.xlsx").get_json()
    check("同一線別再傳一次：整份換掉，不會變兩倍", r["lines"]["寶僑"] == {"count": 3, "before": 3, "blank": 0}, str(r["lines"]))
    wb = openpyxl.Workbook(); wb.active.append(["隨便", "一個", "表"]); b = io.BytesIO(); wb.save(b)
    r = up("/api/shopee/conv/import", b.getvalue(), "x.xlsx")
    check("不是價格本 → 400、講清楚要哪個欄位", r.status_code == 400 and "料號3" in r.get_json()["error"])
    check("歷程記下價格本更新", any(l["field"] == "shopee_conv" for l in c.get("/api/master/logs?scope=shopee&limit=200").get_json()["logs"]))

    print("\n【10】竹運採購進貨：除不盡要先改數量、整批分箱單位／小單位兩個檔")
    imp(purchase_xlsx([("80000000001_1", "6903148358214_CNN", "幫寶適 極上守護 拉拉褲(L+)30片", 30, "Pcs")],
                      po="POTWCTWX26100700B", pr="PRTWCOTWX202610070001", wh="TWX", date="2026-10-21"), "竹運採購進貨")
    pl = plan("竹運採購進貨")
    L, B = pl.get("POTWCTWA26100514L"), pl.get("POTWCTWX26100700B")
    p1 = item(L, "4987176232878_CNN") if L else {}
    check("計畫：兩張寶僑單；190 包除以轉換率 4 除不盡（餘 2），擋下來", L and B and L["blocking"] == 1 and p1["rem"] == 2 and p1["blocking"]
          and "除不盡" in p1["problems"][0] and not B["blocking"], str(p1)[:300])
    p2 = item(L, "4987176232878_RNN")
    check("轉換率 1 的放箱單位、不是 1 的放小單位；竹運進貨單位數量照蝦皮", p2["unit_group"] == "箱單位" and (p2["out_unit"], p2["out_qty"]) == ("箱", 26)
          and item(B, "6903148358214_CNN")["unit_group"] == "小單位" and (item(B, "6903148358214_CNN")["out_unit"], item(B, "6903148358214_CNN")["out_qty"]) == ("包", 30))
    r = gen("竹運採購進貨", ["POTWCTWA26100514L", "POTWCTWX26100700B"])
    check("有除不盡的品項 → 整批不產（400），列出哪一品、什麼原因", r.status_code == 400 and any("4987176232878_CNN" in d and "除不盡" in d for d in r.get_json()["details"]), str(r.get_json()))
    check("沒勾訂單 → 400；履約方式不對 → 400", gen("竹運採購進貨", []).status_code == 400 and gen("竹運出貨", ["POTWCTWA26100514L"]).status_code == 400)
    oid = p1["order_id"]
    r = c.post("/api/shopee/purchase/adjust", json={"order_id": oid, "qty": 191, "reason": "x"})
    check("改成 191 還是除不盡 → 400，提示 188 或 192", r.status_code == 400 and "188" in r.get_json()["error"] and "192" in r.get_json()["error"], r.get_json()["error"])
    check("沒寫原因 → 400", c.post("/api/shopee/purchase/adjust", json={"order_id": oid, "qty": 188, "reason": ""}).status_code == 400)
    r = c.post("/api/shopee/purchase/adjust", json={"order_id": oid, "qty": 188, "reason": "與業務確認改為整箱"})
    p1 = item(plan("竹運採購進貨")["POTWCTWA26100514L"], "4987176232878_CNN")
    logs = [l for l in c.get("/api/master/logs?scope=shopee&limit=200").get_json()["logs"] if l["field"] == "shopee_qty_adj"]
    check("改成 188：不再擋、下採數量 188、留原始數量、轉換率、人員、時間、原因", r.status_code == 200 and not p1["blocking"] and p1["use_qty"] == 188
          and p1["qty"] == 190 and p1["adj"]["qty_before"] == 190 and p1["adj"]["conv"] == 4 and p1["adj"]["operator"] == "Jerry" and p1["adj"]["updated_at"]
          and logs and logs[0]["old_value"] == "190" and logs[0]["new_value"] == "188" and logs[0]["reason"] == "與業務確認改為整箱", str(p1)[:300])
    r = gen("竹運採購進貨", ["POTWCTWA26100514L", "POTWCTWX26100700B"])
    z = zipfile.ZipFile(io.BytesIO(r.data)) if r.status_code == 200 else None
    names = sorted(z.namelist()) if z else []
    z = z if len(names) == 3 else None
    check("產出 zip：整批一組（跨 10/20、10/21 和兩個倉）＝箱單位、小單位兩個 EIP 檔＋一份下採明細；檔名寫該檔實際的日期和倉", names == sorted([
        "產品採購表上傳-寶僑竹運進貨_1020到觀音-箱單位.xls", "產品採購表上傳-寶僑竹運進貨_1020-1021到觀音、安南三-小單位.xls",
        "蝦皮特選-商品下採_寶僑竹運進貨_1020-1021到觀音、安南三.xlsx"]) and r.headers["X-Shopee-Eip-Files"] == "2", str(names))
    if z:
        head, rows = xls_rows(z.read("產品採購表上傳-寶僑竹運進貨_1020-1021到觀音、安南三-小單位.xls"))
        check("小單位檔：料號、蝦皮品名、單位包、數量用改過的 188、備註蝦皮特選、第 10 欄放轉換率且標題空白",
              head[:4] == ["項目", "料號", "品名", "單位"] and head[9] == "" and [r_[1] for r_ in rows] == ["4987176232878", "6903148358214"]
              and [(r_[3], r_[6], r_[8], r_[9]) for r_ in rows] == [("包", 188, "蝦皮特選", 4), ("包", 30, "蝦皮特選", 3)], str(rows))
        head, rows = xls_rows(z.read("產品採購表上傳-寶僑竹運進貨_1020到觀音-箱單位.xls"))
        check("箱單位檔：箱 26、轉換率 1", [(r_[1], r_[3], r_[6], r_[9]) for r_ in rows] == [("4987176232878", "箱", 26, 1)], str(rows))
        wb = openpyxl.load_workbook(io.BytesIO(z.read("蝦皮特選-商品下採_寶僑竹運進貨_1020-1021到觀音、安南三.xlsx")))
        small = [list(r_) for r_ in wb["小單位"].iter_rows(values_only=True)]
        check("下採明細：箱、小單位兩頁；欄位永豐料號／品名／單位／進貨／轉換率／數量每箱（公式）", wb.sheetnames == ["箱", "小單位"]
              and small[0] == ["永豐料號", "品名", "單位", "進貨", "轉換率", "數量/箱"] and small[1][3] == 188 and small[1][5] == "=D2/E2"
              and [list(r_) for r_ in wb["箱"].iter_rows(values_only=True)][1][3] == 26, str(small))
    pl = plan("竹運採購進貨")
    ov = next(g for g in orders()["orders"] if g["order_no"] == "POTWCTWA26100514L")
    check("產出後：計畫和訂單總覽的採購單產出狀態都是「已產出」", pl["POTWCTWA26100514L"]["po_status"] == "已產出" and ov["po_status"] == "已產出", ov["po_status"])
    eips = c.get("/api/shopee/eips").get_json()["eips"]
    small = next(e for e in eips if e["unit_group"] == "小單位")
    check("採購單產出清單：2 個檔、倉別兩個倉、對應的蝦皮訂單", len(eips) == 2 and small["warehouse_name"] == "觀音、安南三"
          and small["orders"] == ["POTWCTWA26100514L", "POTWCTWX26100700B"] and small["rows"] == 2 and small["qty"] == 218, str(small))
    r = gen("竹運採購進貨", ["POTWCTWA26100514L", "POTWCTWX26100700B"])
    check("還沒回填單號可以重產，舊的換掉不會變 4 個", r.status_code == 200 and len(c.get("/api/shopee/eips").get_json()["eips"]) == 2)
    eips = c.get("/api/shopee/eips").get_json()["eips"]
    small = next(e for e in eips if e["unit_group"] == "小單位"); box = next(e for e in eips if e["unit_group"] == "箱單位")
    bad = c.put(f"/api/shopee/eips/{small['id']}", json={"eip_po": "12345"})
    check("回填單號格式不對 → 400", bad.status_code == 400 and "PO" in bad.get_json()["error"])
    r = c.put(f"/api/shopee/eips/{small['id']}", json={"eip_po": "po202610003"})
    check("回填 PO202610003（小寫也行）", r.status_code == 200 and r.get_json()["eip_po"] == "PO202610003")
    check("同一個單號填到別張 → 400", c.put(f"/api/shopee/eips/{box['id']}", json={"eip_po": "PO202610003"}).status_code == 400)
    r = c.get(f"/api/shopee/eips/{small['id']}/file")
    check("重新下載：檔名帶 EIP 單號、內容一樣", r.status_code == 200 and "PO202610003" in r.headers["Content-Disposition"]
          and xls_rows(r.data)[1][0][6] == 188)
    pl = plan("竹運採購進貨")
    p1 = item(pl["POTWCTWA26100514L"], "4987176232878_CNN")
    check("回填後鎖住：那張檔裡的品項不能重產，說要先清單號", p1["blocking"] and "已回填" in p1["problems"][0]
          and gen("竹運採購進貨", ["POTWCTWA26100514L"]).status_code == 400, str(p1["problems"]))
    check("箱單位那張還沒回填 → 訂單狀態「已產出」；兩張都回填才算「已回填」", pl["POTWCTWA26100514L"]["po_status"] == "已產出" and pl["POTWCTWX26100700B"]["po_status"] == "已回填")
    c.put(f"/api/shopee/eips/{small['id']}", json={"eip_po": ""})
    check("清掉單號就能重產", gen("竹運採購進貨", ["POTWCTWA26100514L", "POTWCTWX26100700B"]).status_code == 200)
    logs = c.get("/api/master/logs?scope=shopee&limit=300").get_json()["logs"]
    check("歷程記下產出 EIP 和回填單號", any(l["field"] == "shopee_eip" and l["po_number"] == "POTWCTWA26100514L" for l in logs)
          and any(l["field"] == "shopee_eip_po" and l["new_value"] == "PO202610003" for l in logs))
    r = c.post("/api/shopee/purchase/adjust", json={"order_id": oid, "qty": ""})
    check("修改數量清空＝取消修改，又變回除不盡", r.get_json().get("cleared") and item(plan("竹運採購進貨")["POTWCTWA26100514L"], "4987176232878_CNN")["rem"] == 2)

    print("\n【11】供應商直送：一天一倉一個檔；寶僑小單位換箱、紙潔一律換箱（需求 10.2、10.3）")
    imp(purchase_xlsx([("81_1", "1022257_PNN", "橘子工坊浴廁清潔劑480ml", 360, "Pcs"), ("81_2", "1022248_CNN", "橘子工坊洗碗精補充包500ml", 1392, "Pcs"),
                       ("81_3", "1021315_RNN", "橘子工坊洗衣精補充包 箱購", 42, "Carton")],
                      po="POTWCTWA26100800C", pr="PRTWCOTWA202610080001", date="2026-10-22"), "供應商直送")
    imp(purchase_xlsx([("82_1", "4987176232878_CNN", "幫寶適 清新幫 拉拉褲", 216, "Pcs"), ("82_2", "4987176232878_RNN", "幫寶適 清新幫 拉拉褲 箱購", 10, "Carton")],
                      po="POTWCTWG26100800D", pr="PRTWCOTWG202610080002", wh="TWG", date="2026-10-23"), "供應商直送")
    imp(purchase_xlsx([("83_1", "1088888_PNN", "價格本沒有的紙潔商品", 24, "Pcs"), ("83_2", "1099999_PNN", "箱入數空白的紙潔商品", 24, "Pcs")],
                      po="POTWCTWA26100800E", pr="PRTWCOTWA202610080003", date="2026-10-22"), "供應商直送")
    c.post("/api/shopee/fulfil", json={"order_no": "POTWCTWA26100514N", "fulfil": "供應商直送", "reason": "改回原廠直送"})
    pl = plan("供應商直送")
    E, N = pl["POTWCTWA26100800E"], pl["POTWCTWA26100514N"]
    check("價格本沒有的、箱入數空白的 → 擋下來並說要更新哪個價格本", E["blocking"] == 2 and "請更新紙潔價格本" in item(E, "1088888_PNN")["problems"][0]
          and "沒有箱入數" in item(E, "1099999_PNN")["problems"][0], str([p["problems"] for p in E["items"]]))
    check("瑪氏商品總表還沒有這些料號 → 擋下來、說要到瑪氏出貨頁更新商品總表", N["supported"] and N["blocking"] == 2
          and "瑪氏商品總表找不到" in N["items"][0]["problems"][0], str(N["items"][0]["problems"]))
    r = gen("供應商直送", ["POTWCTWA26100800C", "POTWCTWA26100514N"])
    check("勾到被擋的瑪氏單 → 400", r.status_code == 400)
    r = gen("供應商直送", ["POTWCTWA26100800C", "POTWCTWG26100800D"])
    z = zipfile.ZipFile(io.BytesIO(r.data)) if r.status_code == 200 else None
    names = sorted(z.namelist()) if z else []
    z = z if len(names) == 5 else None
    check("產出：紙潔一個檔、寶僑分箱單位／小單位，各自一天一倉；zip 名稱帶到貨日", names == sorted([
        "產品採購表上傳-紙潔直送_1022到觀音.xls", "產品採購表上傳-寶僑直送_1023到高鐵南-箱單位.xls", "產品採購表上傳-寶僑直送_1023到高鐵南-小單位.xls",
        "蝦皮特選-商品下採_紙潔直送_1022到觀音.xlsx", "蝦皮特選-商品下採_寶僑直送_1023到高鐵南.xlsx"])
          and "1022-1023" in r.headers.get("Content-Disposition", ""), str(names))
    if z:
        head, rows = xls_rows(z.read("產品採購表上傳-紙潔直送_1022到觀音.xls"))
        check("紙潔：一律換成箱（360÷12＝30、1392÷12＝116、箱購 42），備註空白，第 10 欄標題不動",
              [(r_[1], r_[3], r_[6], r_[8]) for r_ in rows] == [("1022257", "箱", 30, ""), ("1022248", "箱", 116, ""), ("1021315", "箱", 42, "")]
              and head[9] != "", str(rows))
        head, rows = xls_rows(z.read("產品採購表上傳-寶僑直送_1023到高鐵南-小單位.xls"))
        check("寶僑直送小單位：單位寫箱、數量 216÷4＝54、轉換率 4", [(r_[1], r_[3], r_[6], r_[9]) for r_ in rows] == [("4987176232878", "箱", 54, 4)], str(rows))
        head, rows = xls_rows(z.read("產品採購表上傳-寶僑直送_1023到高鐵南-箱單位.xls"))
        check("寶僑直送箱單位：箱 10", [(r_[3], r_[6]) for r_ in rows] == [("箱", 10)], str(rows))
    check("瑪氏出貨頁的歷程不會混進蝦皮特選的紀錄", not any(l["field"].startswith("shopee_") for l in c.get("/api/master/logs?scope=mars&limit=500").get_json()["logs"]))

    print("\n【12】瑪氏直送：EIP 單位數量照蝦皮、依品類／單位／中標分檔；回填後產瑪氏採購單（需求 10.4、第十一章）")
    cn = db.get_conn()
    for row in (("M60003196", "盒", "60003196", "Petcare", "希寶 誘惑泥 55入", 794.29, 24, 24, "Q", 24, "V", "", ""),
                ("M10281171", "箱", "10281171", "Petcare", "偉嘉 貓罐頭 箱購", 500.0, 1, 24, "R", 1, "", "11/22 轉貨號", "需加工貼小白標")):
        cn.execute("INSERT INTO mst_mars_products (yf_sku, unit, mars_code, category, name, price, inner_per_case, pcs_per_case, unit_code, box_qty, "
                   "inner_label, note, po_case_note) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)", row)
    cn.commit(); cn.close()
    check("價格本區塊多一個瑪氏（用瑪氏出貨的商品總表）", c.get("/api/shopee/conv").get_json()["lines"].get("瑪氏", {}).get("count") == 2)
    imp(purchase_xlsx([("84_1", "M60003196_QMN", "希寶 誘惑泥 55入", 50, "Pcs")], po="POTWCTWX26101100F", pr="PRTWCOTWX202610110001", wh="TWX", date="2026-10-14"),
        "供應商直送")
    imp(purchase_xlsx([("85_1", "M60003196_QMN", "希寶 誘惑泥 55入", 48, "Pcs")], po="POTWCTWA26101100G", pr="PRTWCOTWA202610110002", date="2026-10-14"),
        "竹運採購進貨")
    pl = plan("供應商直送")
    N, F = pl["POTWCTWA26100514N"], pl["POTWCTWX26101100F"]
    a, b_ = item(N, "M60003196_QMN"), item(N, "M10281171_RMN")
    check("瑪氏：單位數量照蝦皮（48 盒、18 箱），箱數 2／18，依品類、單位、中標分檔，備註 MARS入倉＋到貨日＋倉",
          not N["blocking"] and (a["out_unit"], a["out_qty"], a["cases"], a["unit_group"]) == ("盒", 48, 2, "盒_PET_需貼中標")
          and (b_["out_unit"], b_["out_qty"], b_["cases"], b_["unit_group"]) == ("箱", 18, 18, "箱_PET_不貼中標")
          and a["remark"] == "MARS入倉20261013台灣瑪氏送蝦皮特選-觀音", str(a)[:300])
    f1 = item(F, "M60003196_QMN")
    check("50 盒除以箱入數 24 除不盡 → 擋下來", F["blocking"] == 1 and f1["rem"] == 2 and "除不盡" in f1["problems"][0], str(f1["problems"]))
    r = c.post("/api/shopee/purchase/adjust", json={"order_id": f1["order_id"], "qty": 48, "reason": "瑪氏整箱出貨"})
    check("瑪氏也能改數量（轉換率用商品總表的箱入數）", r.status_code == 200 and not plan("供應商直送")["POTWCTWX26101100F"]["blocking"], str(r.get_json()))
    g_ = plan("竹運採購進貨")["POTWCTWA26101100G"]
    check("瑪氏的竹運採購進貨 → 擋下來（目前只開放供應商直送）", g_["blocking"] and "只開放供應商直送" in g_["items"][0]["problems"][0], str(g_["items"][0]["problems"]))
    r = gen("供應商直送", ["POTWCTWA26100514N", "POTWCTWX26101100F"])
    z = zipfile.ZipFile(io.BytesIO(r.data)) if r.status_code == 200 else None
    names = sorted(z.namelist()) if z else []
    check("產出：一份＝一倉一到貨日一品類一單位一中標，共 3 個 EIP 檔、沒有採購下採明細", names == sorted([
        "產品採購表上傳-瑪氏直送_1013到觀音-盒_PET_需貼中標.xls", "產品採購表上傳-瑪氏直送_1013到觀音-箱_PET_不貼中標.xls",
        "產品採購表上傳-瑪氏直送_1014到安南三-盒_PET_需貼中標.xls"]), str(names))
    if len(names) == 3:
        wb_ = xlrd.open_workbook(file_contents=z.read("產品採購表上傳-瑪氏直送_1013到觀音-盒_PET_需貼中標.xls"))
        _, rows = xls_rows(z.read("產品採購表上傳-瑪氏直送_1013到觀音-盒_PET_需貼中標.xls"))
        check("瑪氏 EIP：套瑪氏範本（產品領用單）、料號、單位盒、數量 48、備註", wb_.sheet_names()[0] == "產品領用單"
              and [(r_[1], r_[3], r_[6], r_[8]) for r_ in rows] == [("M60003196", "盒", 48, "MARS入倉20261013台灣瑪氏送蝦皮特選-觀音")], str(rows))
    eips = {e["filename"]: e for e in c.get("/api/shopee/eips").get_json()["eips"] if e["line"] == "瑪氏"}
    box_ = eips.get("產品採購表上傳-瑪氏直送_1013到觀音-盒_PET_需貼中標.xls", {})
    case_ = eips.get("產品採購表上傳-瑪氏直送_1013到觀音-箱_PET_不貼中標.xls", {})
    check("採購單產出清單標出瑪氏", box_.get("mars") and case_.get("mars"))
    r = c.get(f"/api/shopee/eips/{box_['id']}/mars_po")
    check("還沒回填 EIP 採購單號 → 不能產瑪氏採購單（400）", r.status_code == 400 and "回填" in r.get_json()["error"])
    check("沒回填的不能打包下載（400）", c.get("/api/shopee/mars_po.zip").status_code == 400)
    c.put(f"/api/shopee/eips/{box_['id']}", json={"eip_po": "PO202610101"})
    c.put(f"/api/shopee/eips/{case_['id']}", json={"eip_po": "PO202610102"})
    r = c.get(f"/api/shopee/eips/{box_['id']}/mars_po")
    import datetime as _dt
    ws = openpyxl.load_workbook(io.BytesIO(r.data)).active if r.status_code == 200 else None
    check("瑪氏採購單檔名：EIP 採購單號＋到貨日＋倉＋單位＋品類＋中標", "PO202610101" in r.headers.get("Content-Disposition", "")
          and "_1013_" in __import__("urllib.parse").parse.unquote(r.headers.get("Content-Disposition", "")), r.headers.get("Content-Disposition"))
    if ws:
        check("表頭：下單日 10/8（往前 2 個工作天，跳過 10/9 補假和週末）、配送日 10/13、EIP 採購單號、觀音倉資料、1/2效期以上、品類",
              ws["C3"].value == _dt.datetime(2026, 10, 8) and ws["F3"].value == _dt.datetime(2026, 10, 13) and ws["C5"].value == "PO202610101"
              and ws["C6"].value == "永豐商店蝦皮特選-觀音倉(TWA)" and ws["F6"].value == "1/2效期以上" and ws["F7"].value == 17617040
              and ws["F8"].value == "進貨組" and ws["B9"].value == "Petcare寵物", str([ws[k].value for k in ("C3", "F3", "C5", "C6", "F6", "F7", "F8", "B9")]))
        check("作業備註：蝦皮特選嘜頭＋採購單、需貼中盒標（盒單位不寫貼板嘜）", ws["F5"].value == "蝦皮特選嘜頭(入庫通知單)+採購單\n需貼中盒標", repr(ws["F5"].value))
        check("明細：永豐料號、瑪氏貨號、單位需求 48、採購數量 2 箱、中標 V、每箱數", [ws.cell(12, k).value for k in (1, 2, 4, 5, 8, 10, 11)]
              == ["M60003196", "60003196", 48, 2, "V", 24, 24], str([ws.cell(12, k).value for k in range(1, 14)]))
    ws = openpyxl.load_workbook(io.BytesIO(c.get(f"/api/shopee/eips/{case_['id']}/mars_po").data)).active
    check("箱單位那份：作業備註加貼板嘜、指定品需加工貼小白標；備註、小白標欄照商品總表",
          ws["F5"].value == "蝦皮特選嘜頭(入庫通知單)+採購單\n貼板嘜\n指定品需加工貼小白標" and ws.cell(12, 12).value == "11/22 轉貨號"
          and ws.cell(12, 13).value == "需加工貼小白標" and ws.cell(12, 5).value == 18, repr(ws["F5"].value))
    r = c.get("/api/shopee/mars_po.zip")
    zn = sorted(zipfile.ZipFile(io.BytesIO(r.data)).namelist()) if r.status_code == 200 else []
    check("打包下載只放已回填的 2 份", zn == ["PO202610101永豐Mars採購單_1013_觀音_盒_PET_需貼中標.xlsx", "PO202610102永豐Mars採購單_1013_觀音_箱_PET_不貼中標.xlsx"], str(zn))
    whs = {w["code"]: w for w in c.get("/api/shopee/warehouses").get_json()["warehouses"]}
    check("蝦皮倉資料：觀音、安南三照需求書，其他倉 ship-to 空著", whs["TWX"]["ship_to"] == "17647732" and whs["TWX"]["contact"] == "曾小姐"
          and whs["TWG"]["address"].startswith("桃園市中壢區") and whs["TWG"]["ship_to"] == "", str(whs["TWG"]))
    r = c.put("/api/shopee/warehouses/TWG", json={"ship_to": "17600001", "contact": "高鐵南進貨組"})
    whs = {w["code"]: w for w in c.get("/api/shopee/warehouses").get_json()["warehouses"]}
    logs = [l for l in c.get("/api/master/logs?scope=shopee&limit=500").get_json()["logs"] if l["field"] == "shopee_wh"]
    check("改高鐵南的 ship-to：存好、只記有改的那一欄", r.get_json()["changed"] == 1 and whs["TWG"]["ship_to"] == "17600001" and whs["TWG"]["edited"]
          and len(logs) == 1 and logs[0]["new_value"] == "17600001", str(logs))
    check("沒有的倉 → 404", c.put("/api/shopee/warehouses/XXX", json={"ship_to": "1"}).status_code == 404)
    ship_tests(c, preview, orders, up)


def ship_tests(c, preview, orders, up):
    import xlwt
    from shopee.ship import ratio_for, DEFAULT_SET
    print("\n【13】竹運出貨：照 Alice 的核對檔判斷效期與庫存，逐列扣庫存，產拋單檔＋下修刪單清單")
    page = c.get("/shopee").get_data(as_text=True)
    check("竹運出貨頁籤打開了、有庫存上傳、設定、清單、批次視窗", all(x in page for x in ('data-tab="zhu"', 'id="dz-stock"', 'id="s-rule"', 'id="z-tbl"', 'id="dlg-ship"')))
    st = dict(DEFAULT_SET)
    check("效期比例照蝦皮進倉規範：90 天內 2/3、365 天內 2/5、其他 1/3；自訂就用自訂", ratio_for(90, st)[1] == "2/3" and ratio_for(300, st)[1] == "2/5"
          and ratio_for(1095, st)[1] == "1/3" and ratio_for(1095, dict(st, rule="custom", ratio=0.66))[0] == 0.66)
    def imp_in(rows, po, date="2026-10-20"):
        pv = preview(inbound_xlsx(rows, po, date=date), "InboundOrder_z.xlsx").get_json()
        return c.post("/api/shopee/import/commit", json={"upload_id": pv["upload_id"], "fulfil": "竹運出貨"}).get_json()
    imp_in([("INZA01", "z1", "4987176377708_CNN", "幫寶適 安睡褲", 81), ("INZA02", "z2", "6903148224090_VNN", "好自在 衛生棉", 144),
            ("INZA03", "z3", "4987176372093_TNN", "好自在 雲枕組", 12), ("INZA04", "z4", "4902430056519_PNN", "歐蕾 護膚霜", 102),
            ("INZA05", "z5", "4987176386700_CNN", "幫寶適 新品", 12), ("INZA06", "z6", "4987176000000_CNN", "主檔沒有的品", 10)], "POTWCTWA26101500H")
    imp_in([("INZB01", "z7", "4987176377708_CNN", "幫寶適 安睡褲", 80)], "POTWCTWA26101500J")
    q = {"from": "2026-10-20", "to": "2026-10-20"}
    def plan():
        return c.get("/api/shopee/ship/plan", query_string=q).get_json()
    def it(pl, no, sup):
        return next(x for g in pl["orders"] if g["order_no"] == no for x in g["items"] if x["supplier_sku_id"] == sup)
    pl = plan()
    check("還沒上傳庫存 → 每品都擋下來，說要先上傳 EC 即時庫存、竹運庫存效期表", all(x["blocking"] for g in pl["orders"] for x in g["items"])
          and "EC 即時庫存" in it(pl, "POTWCTWA26101500H", "4987176377708_CNN")["problems"][0], str(it(pl, "POTWCTWA26101500H", "4987176377708_CNN")["problems"]))
    wbx = xlwt.Workbook(); sh = wbx.add_sheet("全商品")
    for j, h in enumerate(["結餘日期", "料號", "單位", "品名", "EC可用量"]):
        sh.write(0, j, h)
    for i, row in enumerate([("2026/10/19", "4987176377708", "包", "安睡褲", 132), ("2026/10/19", "6903148224090", "袋", "衛生棉", 187),
                             ("2026/10/19", "4987176372093", "組", "雲枕組", 1), ("2026/10/19", "4902430056519", "瓶", "護膚霜", 100),
                             ("2026/10/19", "4987176000000", "包", "沒主檔", 50), ("2026/10/19", "4987176377708", "箱", "安睡褲箱", 9)], start=1):
        for j, v in enumerate(row):
            sh.write(i, j, v)
    bx = io.BytesIO(); wbx.save(bx)
    r = up("/api/shopee/stock/import", bx.getvalue(), "全商品_20261019.xls").get_json()
    check("上傳 EC 即時庫存（舊版 .xls）：認得、記結餘日期", r.get("kind") == "ec" and r["stock"]["ec"]["file_date"] == "2026/10/19" and r["stock"]["ec"]["today"], str(r))
    wbl = openpyxl.Workbook(); ws = wbl.active; ws.title = "sheet1"
    ws.append(["儲區", "客戶品號", "儲位", "可出貨量", "庫存量", "效期", "製造日期"])
    for row in (("良品區", "4987176377708C", "A1", 100, 100, "20290522", "20260523"), ("良品區", "4987176377708C", "A2", 38, 38, "20290522", ""),
                ("良品區", "4987176377708C", "A3", 41, 41, "20290502", ""), ("不良品區", "4987176377708C", "X1", 500, 500, "20300101", ""),
                ("良品區", "6903148224090V", "B1", 96, 96, "20290724", ""), ("良品區", "6903148224090V", "B2", 91, 91, "20290601", ""),
                ("良品區", "4987176372093T", "C1", 2, 2, "20271217", ""), ("良品區", "4902430056519P", "D1", 50, 50, "20270601", ""),
                ("良品區", "4902430056519P", "D2", 20, 20, "20290101", ""), ("良品區", "4987176000000C", "E1", 50, 50, "20290101", "")):
        ws.append(list(row))
    bl = io.BytesIO(); wbl.save(bl)
    r = up("/api/shopee/stock/import", bl.getvalue(), "效期表.xlsx").get_json()
    check("上傳竹運庫存效期表：認得、不良品區不算", r.get("kind") == "lots" and "非良品區 1 列不計" in r["notes"], str(r))
    def master(extra=()):
        wbm = openpyxl.Workbook(); wbm.active.title = "全商品"; wbm.active.append(["料號", "單位", "EC可用量"])
        wsm = wbm.create_sheet("商品主檔"); wsm.append(["客戶品號", "商品名稱", "國際條碼", "供應商代碼", "單位", "長", "寬", "高", "重", "箱入數", "商品年限(日)"])
        for code in ("4987176377708C", "6903148224090V", "4987176372093T", "4902430056519P") + tuple(extra):
            wsm.append([code, "", "", "", "", "", "", "", "", 1, 1095])
        wsm.append(["9999999999999C", "年限空白", "", "", "", "", "", "", "", 1, ""])
        bm = io.BytesIO(); wbm.save(bm)
        return bm.getvalue()
    r = up("/api/shopee/stock/import", master(), "核對竹運出貨效期及庫存V5.xlsm").get_json()
    check("上傳 Alice 的核對檔：只匯入商品主檔、年限空白的列出來", r.get("kind") == "life" and r["items"] == 4
          and any("只匯入商品主檔" in n for n in r["notes"]) and any("空白" in n for n in r["notes"]), str(r))
    wb_ = openpyxl.Workbook(); wb_.active.append(["隨便", "一個"]); b_ = io.BytesIO(); wb_.save(b_)
    r = up("/api/shopee/stock/import", b_.getvalue(), "x.xlsx")
    check("認不出的檔 → 400、說要哪些欄位", r.status_code == 400 and "EC可用量" in r.get_json()["error"])
    pl = plan()
    H, J = "POTWCTWA26101500H", "POTWCTWA26101500J"
    a = it(pl, H, "4987176377708_CNN")
    check("同效期不同儲位加總（138）；單一效期挑最早、量又夠的那批（2029-05-22 出 81）；合格基準＝到貨日＋1095×1/3",
          a["final"] == 81 and a["alloc"] == [{"exp": "2029-05-22", "qty": 81}] and a["threshold"] == "2027-10-20" and a["ratio_text"] == "1/3"
          and a["book"] == 179 and not a["blocking"], str({k: a[k] for k in ("final", "alloc", "threshold", "book")}))
    b_it = it(pl, H, "6903148224090_VNN")
    check("單一效期都不夠 → 先出量最多那批 96，缺 48，提醒合併效期最多可出 144", b_it["final"] == 96 and b_it["short"] == 48
          and b_it["alloc"] == [{"exp": "2029-07-24", "qty": 96}] and any("最多可出 144" in w for w in b_it["warnings"]), str(b_it["warnings"]))
    c_it = it(pl, H, "4987176372093_TNN")
    check("EC 只剩 1 → 出 1、提醒少於 12 件蝦皮會拒收", c_it["final"] == 1 and any("少於 12 件" in w for w in c_it["warnings"]), str(c_it["warnings"]))
    d = it(pl, H, "4902430056519_PNN")
    check("有一批效期不合格（2027-06-01）：只出合格的 20，缺貨原因效期不足＋庫存不足", d["final"] == 20 and d["qualified"] == 20 and d["book"] == 70
          and d["reasons"] == ["效期不足", "庫存不足"] and [l["ok"] for l in d["lots"]] == [False, True], str(d["reasons"]))
    e_ = it(pl, H, "4987176386700_CNN")
    check("完全沒貨、主檔也沒有 → 不擋，算庫存不足、出 0", e_["final"] == 0 and e_["reasons"] == ["庫存不足"] and not e_["blocking"], str(e_["problems"]))
    f_ = it(pl, H, "4987176000000_CNN")
    check("有貨但商品主檔沒有年限 → 擋下來要補主檔（不當成 0 天）", f_["blocking"] and "商品主檔" in f_["problems"][0], str(f_["problems"]))
    j = it(pl, J, "4987176377708_CNN")
    check("逐列扣庫存：後面那張單只剩 EC 51、2029-05-22 剩 57 → 出 51、缺 29", j["ec"] == 51 and j["final"] == 51 and j["short"] == 29
          and j["alloc"] == [{"exp": "2029-05-22", "qty": 51}], str({k: j[k] for k in ("ec", "final", "alloc")}))
    r = c.post("/api/shopee/ship/generate", json={"order_nos": [H, J], **q})
    check("有要處理的（主檔沒年限）→ 整批不產、列出原因", r.status_code == 400 and any("4987176000000_CNN" in x for x in r.get_json()["details"]), str(r.get_json()))
    up("/api/shopee/stock/import", master(("4987176000000C",)), "產品基本資料.xlsx")
    adj = lambda body: c.post("/api/shopee/ship/adjust", json={**q, **body})  # noqa: E731
    r = adj({"order_id": b_it["order_id"], "lots": [{"exp": "2029-07-24", "qty": 96}, {"exp": "2029-06-01", "qty": 48}], "reason": ""})
    check("改出貨量沒寫原因 → 400", r.status_code == 400 and "原因" in r.get_json()["error"])
    r = adj({"order_id": b_it["order_id"], "lots": [{"exp": "2029-06-01", "qty": 92}], "reason": "x"})
    check("選的數量超過那批剩的 → 400", r.status_code == 400 and "只剩 91" in r.get_json()["error"], r.get_json()["error"])
    r = adj({"order_id": d["order_id"], "lots": [{"exp": "2027-06-01", "qty": 10}], "reason": "x"})
    check("選效期不合格的批次 → 400", r.status_code == 400 and "不符合" in r.get_json()["error"])
    r = adj({"order_id": b_it["order_id"], "lots": [{"exp": "2029-07-24", "qty": 96}, {"exp": "2029-06-01", "qty": 48}], "reason": "業務確認兩個效期分開入庫"})
    r2 = adj({"order_id": c_it["order_id"], "lots": [], "reason": "低於 12 件會被拒收，刪單"})
    pl = plan()
    b_it, c_it = it(pl, H, "6903148224090_VNN"), it(pl, H, "4987176372093_TNN")
    logs = [l for l in c.get("/api/master/logs?scope=shopee&limit=500").get_json()["logs"] if l["field"] == "shopee_ship_qty"]
    check("人工合併兩個效期 → 出 144、不缺；記人員、原因、提醒入庫單要依效期拆", r.status_code == 200 and b_it["final"] == 144 and not b_it["short"]
          and b_it["adj"]["reason"] == "業務確認兩個效期分開入庫" and any("依效期拆開" in w for w in b_it["warnings"]) and logs, str(b_it["warnings"]))
    check("改成 0 → 刪單，缺 12", r2.status_code == 200 and c_it["final"] == 0 and c_it["short"] == 12)
    r = c.put("/api/shopee/ship/settings", json={"rule": "custom", "ratio": "abc"})
    check("效期比例亂填 → 400", r.status_code == 400)
    c.put("/api/shopee/ship/settings", json={"rule": "custom", "ratio": "66"})
    d2 = it(plan(), H, "4902430056519_PNN")
    check("改成自訂比例 66（當 0.66）→ 基準日＝到貨日＋722 天", d2["ratio_text"] == "自訂 0.66" and d2["threshold"] == "2028-10-11", str((d2["ratio_text"], d2["threshold"])))
    c.put("/api/shopee/ship/settings", json={"rule": "shopee"})
    r = c.post("/api/shopee/ship/generate", json={"order_nos": [H, J], **q})
    z = zipfile.ZipFile(io.BytesIO(r.data)) if r.status_code == 200 else None
    names = sorted(z.namelist()) if z else []
    check("產出 zip：拋單檔（一天一倉一線別）＋下修刪單清單", names == ["指定到貨日20261020蝦皮(觀音)_PG.xlsx", "蝦皮特選下修刪單清單_1020.xlsx"], str(names))
    if len(names) == 2:
        rows = list(openpyxl.load_workbook(io.BytesIO(z.read(names[0]))).active.iter_rows(values_only=True))
        check("拋單檔照範例：表頭、訂單單號＝入庫單號、蝦皮-觀音、範例地址電話、料號＝國條、出貨量 0 的不放", rows[0] == ("訂單單號", "姓名", "地址", "電話", "手機", "出貨備註", "料號", "產品名稱", "單位", "數量")
              and rows[1][:4] == ("INZA01", "蝦皮-觀音", "桃園市觀音區寶倉街108號C棟1樓", "(03)262-6560") and rows[1][6:] == ("4987176377708", "幫寶適 安睡褲", "包", 81)
              and sorted((r_[0], r_[9]) for r_ in rows[1:]) == [("INZA01", 81), ("INZA02", 144), ("INZA04", 20), ("INZA06", 10), ("INZB01", 51)], str(rows[1:]))
        srows = list(openpyxl.load_workbook(io.BytesIO(z.read(names[1]))).active.iter_rows(values_only=True))
        acts = {(r_[5], r_[0], r_[10], r_[11], r_[12]) for r_ in srows[1:]}
        check("下修刪單清單：刪單／下修、下修數量、缺貨原因、調整原因", ("4987176372093_TNN", "刪單", 12, "庫存不足", "低於 12 件會被拒收，刪單") in acts
              and ("4902430056519_PNN", "下修", 82, "效期不足、庫存不足", None) in acts and ("4987176386700_CNN", "刪單", 12, "庫存不足", None) in acts
              and ("4987176377708_CNN", "下修", 29, "庫存不足", None) in acts and len(srows) == 5, str(acts))
    pl = plan()
    check("產出後兩張都「已產出」、品項照快照（不再分配庫存）", all(g["dispatch_status"] == "已產出" for g in pl["orders"])
          and it(pl, H, "6903148224090_VNN")["final"] == 144 and it(pl, H, "6903148224090_VNN")["dispatch"])
    ov = next(g for g in orders()["orders"] if g["order_no"] == H)
    check("訂單總覽的出貨量、缺貨量帶出來", ov["qty_ship"] == 255 and ov["qty_short"] == 106, str((ov["qty_ship"], ov["qty_short"])))
    check("再產一次同樣的單 → 400（要先取消產出）", c.post("/api/shopee/ship/generate", json={"order_nos": [H, J], **q}).status_code == 400)
    r = c.post("/api/shopee/ship/adjust", json={**q, "order_id": b_it["order_id"], "lots": [{"exp": "2029-07-24", "qty": 96}], "reason": "x"})
    check("已產出的不能改出貨量 → 400", r.status_code == 400 and "取消產出" in r.get_json()["error"])
    r = c.post("/api/shopee/ship/undo", json={"order_no": J})
    pl = plan()
    check("取消產出：那張回到未產出、重新照庫存判斷（前面那張已產出不再扣 → 這次可出 80）", r.get_json()["removed"] == 1
          and next(g for g in pl["orders"] if g["order_no"] == J)["dispatch_status"] == "未產出" and it(pl, J, "4987176377708_CNN")["final"] == 80,
          str(it(pl, J, "4987176377708_CNN")["final"]))
    logs = c.get("/api/master/logs?scope=shopee&limit=500").get_json()["logs"]
    check("歷程記下上傳庫存、產出拋單檔、取消產出", any(l["field"] == "shopee_stock" for l in logs)
          and any(l["field"] == "shopee_ship" and "出貨" in (l["new_value"] or "") for l in logs) and any(l["field"] == "shopee_ship" and l["new_value"] == "未產出" for l in logs))
    w = {x["code"]: x for x in c.get("/api/shopee/dispatch_wh").get_json()["warehouses"]}
    r = c.put("/api/shopee/dispatch_wh/TWA", json={"phone": "03-2600500#5924"})
    check("拋單檔收件資料：7 個倉、觀音照範例，可以改", len(w) == 7 and w["TWA"]["phone"] == "(03)262-6560" and r.get_json()["changed"] == 1
          and {x["code"]: x for x in c.get("/api/shopee/dispatch_wh").get_json()["warehouses"]}["TWA"]["phone"] == "03-2600500#5924")


if __name__ == "__main__":
    main()

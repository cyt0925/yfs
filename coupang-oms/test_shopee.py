"""蝦皮特選寄倉訂單系統（/shopee）端到端測試。
第一段：上傳採購單／入庫單、重傳差異、版本、確認、改履約方式、總覽。
第二段：價格本、竹運採購進貨／供應商直送產 EIP 採購單＋採購下採明細、除不盡改數量、回填 EIP 單號。

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
    check("瑪氏直送：先列出但不能勾（之後做）", not N["supported"] and "瑪氏" in N["items"][0]["problems"][0])
    r = gen("供應商直送", ["POTWCTWA26100800C", "POTWCTWA26100514N"])
    check("勾到瑪氏 → 400", r.status_code == 400)
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


if __name__ == "__main__":
    main()

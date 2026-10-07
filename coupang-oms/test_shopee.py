"""蝦皮特選寄倉訂單系統（/shopee）第一段的端到端測試：上傳採購單／入庫單、重傳差異、版本、確認、改履約方式、總覽。

測試檔照蝦皮後台下載檔的表頭用程式造（真的業務檔不放進 repo）。
執行：python test_shopee.py（設了 DATABASE_URL 就跑 PostgreSQL）
"""
import io
import os
import sys
import tempfile

import openpyxl

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

    print("\n" + "=" * 62)
    print(f"通過 {len(PASS)} 項，失敗 {len(FAIL)} 項")
    if FAIL:
        print("失敗項目："); [print("  -", n) for n in FAIL]
        sys.exit(1)


if __name__ == "__main__":
    main()

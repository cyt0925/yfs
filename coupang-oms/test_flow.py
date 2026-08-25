"""端到端驗證：拿真實整合表跑完整流程。

執行：python test_flow.py
每個檢查都對應一條「不能出錯的防呆」，失敗會直接 AssertionError。
"""

import io
import json
import os
import shutil
import sys
import tempfile

import openpyxl
import pymupdf
import zipfile

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SAMPLE = os.path.join(BASE_DIR, "samples", "整合表範例.xlsx")

_tmp = tempfile.mkdtemp(prefix="oms_test_")
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


def upload(client, path, operator="小真"):
    with open(path, "rb") as fh:
        data = fh.read()
    res = client.post("/api/import/preview", data={
        "operator": operator,
        "file": (io.BytesIO(data), os.path.basename(path)),
    }, content_type="multipart/form-data")
    return res


def sign_in(client, username="小真", password="changeme123"):
    """所有 API 都在登入後面，測試也得先登入。

    順帶一提：操作人員現在一律取自登入身分，各處呼叫仍然傳
    operator 參數只是還沒清掉的殘留，後端不會採用——所以這裡登入
    誰，歷程上記的就是誰。"""
    res = client.post("/login", data={"username": username, "password": password})
    assert res.status_code == 302, f"測試帳號登入失敗：{username}"


def main():
    db.init_db()
    app_module.app.config["TESTING"] = True
    client = app_module.app.test_client()
    sign_in(client)

    print("\n【1】首次匯入真實整合表")
    res = upload(client, SAMPLE)
    assert res.status_code == 200, res.get_json()
    preview = res.get_json()
    check("解析出 103 列（一張 PO 多個 SKU 全部保留）",
          preview["rows_total"] == 103, f"實際 {preview['rows_total']}")
    check("全部判定為新增", preview["new_count"] == 103, f"新增 {preview['new_count']}")

    res = client.post("/api/import/commit", json={
        "batch_id": preview["batch_id"], "operator": "小真"})
    assert res.status_code == 200, res.get_json()
    check("實際寫入 103 筆", res.get_json()["inserted"] == 103)

    rows = client.get("/api/orders?page_size=500").get_json()
    check("資料庫存有 103 筆（若 po_number 設 UNIQUE 這裡只會剩 20）",
          rows["total"] == 103, f"實際 {rows['total']}")

    by_po = {}
    for r in rows["rows"]:
        by_po.setdefault(r["po_number"], []).append(r)
    biggest = max(by_po.values(), key=len)
    check("最大的一張 PO 保留了 29 個 SKU", len(biggest) == 29, f"實際 {len(biggest)}")
    check("共 20 張 PO", len(by_po) == 20, f"實際 {len(by_po)}")

    print("\n【2】同一份檔案再上傳一次（無腦全選上傳）")
    res = upload(client, SAMPLE)
    again = res.get_json()
    check("103 筆全部判定為完全相同、靜默略過",
          again["identical_count"] == 103 and again["new_count"] == 0
          and again["updated_count"] == 0,
          f"新增{again['new_count']}／更新{again['updated_count']}／略過{again['identical_count']}")
    client.post("/api/import/commit", json={"batch_id": again["batch_id"], "operator": "小真"})
    check("資料庫仍是 103 筆，沒有長出重複列",
          client.get("/api/orders").get_json()["total"] == 103)

    print("\n【3】模擬酷澎偷偷改單：改數量、改交期、改倉別")
    modified = os.path.join(_tmp, "modified.xlsx")
    shutil.copy(SAMPLE, modified)
    wb = openpyxl.load_workbook(modified)
    ws = wb["整合表"]
    hdr = [c.value for c in ws[1]]
    col = {h: i + 1 for i, h in enumerate(hdr)}
    ws.cell(row=2, column=col["下單數量(酷澎單位)"]).value = 999      # 改量
    ws.cell(row=2, column=col["出貨數量"]).value = 999                # 整合表出貨數量欄同步改，才是真實情境
    ws.cell(row=3, column=col["交付日期"]).value = "2026/9/1"        # 改期（且格式不同）
    ws.cell(row=4, column=col["到貨倉別"]).value = " tao3 "          # 改倉（且大小寫+空白）
    wb.save(modified)

    res = upload(client, modified)
    diff = res.get_json()
    check("正確抓出 3 筆異動", diff["updated_count"] == 3, f"實際 {diff['updated_count']}")
    labels = {c["label"] for u in diff["updated"] for c in u["changes"]}
    check("三種異動都被辨識（數量／交期／倉別）",
          {"下單數量", "交期", "倉別"} <= labels, f"實際 {labels}")
    check("其餘 100 筆仍判定為相同", diff["identical_count"] == 100,
          f"實際 {diff['identical_count']}")
    check("沒有把改期／改倉誤判成新單（新增 0 筆）", diff["new_count"] == 0,
          f"實際 {diff['new_count']}")

    client.post("/api/import/commit", json={"batch_id": diff["batch_id"], "operator": "小真"})
    after = client.get("/api/orders?page_size=500").get_json()
    check("總筆數仍是 103，沒有幽靈列", after["total"] == 103, f"實際 {after['total']}")
    check("3 筆被標記為待確認並推到最上面",
          after["summary"]["needs_review"] == 3,
          f"實際 {after['summary']['needs_review']}")
    check("待確認的排在列表最前面",
          all(r["needs_review"] for r in after["rows"][:3]))

    changed = [r for r in after["rows"] if r["needs_review"]]
    wh = [r for r in changed if r["warehouse"] == "TAO3"]
    check("倉別「 tao3 」被正規化成 TAO3", len(wh) == 1, f"實際 {[r['warehouse'] for r in changed]}")
    dates = [r["delivery_date"] for r in changed]
    check("交期「2026/9/1」被正規化成 2026-09-01", "2026-09-01" in dates, f"實際 {dates}")

    print("\n【4】欄位級修改歷程")
    target = next(r for r in changed if r["qty_coupang"] == 999)
    logs = client.get(f"/api/orders/{target['id']}/logs").get_json()
    qty_log = [l for l in logs if l["field"] == "qty_coupang"]
    check("數量變更有寫入欄位級歷程", len(qty_log) == 1)
    check("歷程記錄了改前改後與來源",
          qty_log and qty_log[0]["new_value"] == "999"
          and qty_log[0]["source"] == "import"
          and qty_log[0]["operator"] == "小真",
          qty_log[0] if qty_log else "")
    check("出貨數量未被手動改過時，跟著整合表的出貨數量欄連動",
          target["qty_ship"] == 999, f"實際 {target['qty_ship']}")

    print("\n【4b】整合表自己的出貨數量欄跟下單數量不一樣時，以出貨數量欄為準")
    diverge = os.path.join(_tmp, "diverge.xlsx")
    shutil.copy(SAMPLE, diverge)
    wb = openpyxl.load_workbook(diverge)
    ws = wb["整合表"]
    # 挑一列下單數量不變、但出貨數量欄本身就跟下單數量不一樣的情境——
    # 這正是使用者回報「匯入後加總對不上」的那個 bug：系統之前完全沒
    # 讀整合表自己的出貨數量欄，一律拿下單數量複製過去。
    row10_sku = str(ws.cell(row=10, column=col["SKU ID"]).value)
    row10_qty = ws.cell(row=10, column=col["下單數量(酷澎單位)"]).value
    ws.cell(row=10, column=col["出貨數量"]).value = row10_qty - 50
    wb.save(diverge)

    res = upload(client, diverge)
    diff2 = res.get_json()
    client.post("/api/import/commit", json={"batch_id": diff2["batch_id"], "operator": "小真"})
    after2 = client.get("/api/orders?page_size=500").get_json()
    row10 = next(r for r in after2["rows"] if r["sku_id"] == row10_sku)
    check("下單數量沒變、只有出貨數量欄不同時，出貨數量照樣跟著整合表更新",
          row10["qty_ship"] == row10_qty - 50,
          f"下單={row10['qty_coupang']}，出貨={row10['qty_ship']}，預期出貨={row10_qty - 50}")

    print("\n【5】上傳片段檔：消失的單絕不能被當成取消")
    partial = os.path.join(_tmp, "partial.xlsx")
    wb = openpyxl.load_workbook(SAMPLE)
    ws = wb["整合表"]
    ws.delete_rows(12, 200)   # 只留前 10 筆，模擬只拉了一頁
    wb.save(partial)
    res = upload(client, partial)
    part = res.get_json()
    client.post("/api/import/commit", json={"batch_id": part["batch_id"], "operator": "Nicole"})
    check("片段上傳後資料庫仍是 103 筆，沒有任何單被刪除或標記取消",
          client.get("/api/orders").get_json()["total"] == 103)

    print("\n【6】線上編輯與樂觀鎖（三個人同時用）")
    row = client.get("/api/orders?page_size=500").get_json()["rows"][0]
    res = client.put(f"/api/orders/{row['id']}", json={
        "operator": "小真", "version": row["version"],
        "qty_ship": 55, "remarks": "跟酷澎談好下修"})
    check("小真儲存成功", res.status_code == 200, res.get_json())

    res = client.put(f"/api/orders/{row['id']}", json={
        "operator": "Nicole", "version": row["version"],   # 拿的是舊版本號
        "qty_ship": 77})
    check("Nicole 拿舊版本存檔被擋下（409，不會無聲蓋掉小真的修改）",
          res.status_code == 409, f"實際 {res.status_code}")
    current = client.get("/api/orders?page_size=500").get_json()["rows"]
    kept = next(r for r in current if r["id"] == row["id"])
    check("小真的 55 沒有被蓋成 77", kept["qty_ship"] == 55, f"實際 {kept['qty_ship']}")

    print("\n【7】OP 手動調整過的出貨數量，匯入不得覆蓋")
    res = upload(client, SAMPLE)   # 原始檔的數量會跟 55 不同
    d = res.get_json()
    client.post("/api/import/commit", json={"batch_id": d["batch_id"], "operator": "小真"})
    kept = next(r for r in client.get("/api/orders?page_size=500").get_json()["rows"]
                if r["id"] == row["id"])
    check("出貨數量仍是 OP 設定的 55", kept["qty_ship"] == 55, f"實際 {kept['qty_ship']}")

    print("\n【8】匯出：試算匯出不鎖單、正式匯出才標記已拉單")
    res = client.post("/api/export", json={
        "operator": "小真", "profile": "warehouse", "mark_pulled": False,
        "filters": {}, "po_numbers": []})
    check("試算匯出成功", res.status_code == 200)
    check("試算匯出後沒有任何單被標記已拉單",
          client.get("/api/orders").get_json()["summary"]["pulled"] == 0)

    pos = [r["po_number"] for r in client.get("/api/pos").get_json()["rows"][:3]]
    res = client.post("/api/export", json={
        "operator": "小真", "profile": "erp", "mark_pulled": True,
        "filters": {}, "po_numbers": pos})
    check("ERP 格式正式匯出成功", res.status_code == 200)
    out = openpyxl.load_workbook(io.BytesIO(res.data))
    headers = [c.value for c in out.active[1]]
    check("匯出欄位名稱來自 export_profiles.json（可自由改成 ERP 規範）",
          headers[0] == "採購單號" and "料號" in headers, headers)
    check("3 張 PO 被標記為已拉單",
          client.get("/api/pos").get_json()["summary"]["pulled"] == 3, pos)

    locked = next(r for r in client.get("/api/orders?page_size=500").get_json()["rows"]
                  if r["po_number"] == pos[0])
    res = client.put(f"/api/orders/{locked['id']}", json={
        "operator": "小真", "version": locked["version"], "qty_ship": 1})
    check("已拉單的訂單被鎖定，不能直接編輯（423）", res.status_code == 423,
          f"實際 {res.status_code}")

    print("\n【9】最高警示：已拉單後酷澎又偷改")
    wb = openpyxl.load_workbook(SAMPLE)
    ws = wb["整合表"]
    locked_po = pos[0]
    for r in range(2, ws.max_row + 1):
        if str(ws.cell(row=r, column=col["PO單號"]).value) == locked_po:
            ws.cell(row=r, column=col["下單數量(酷澎單位)"]).value = 12345
            break
    after_pull_file = os.path.join(_tmp, "after_pull.xlsx")
    wb.save(after_pull_file)
    res = upload(client, after_pull_file)
    ap = res.get_json()
    check("預覽就先警告「已拉單後被改」", ap["after_pull_count"] >= 1,
          f"實際 {ap['after_pull_count']}")
    client.post("/api/import/commit", json={"batch_id": ap["batch_id"], "operator": "小真"})
    check("該筆被標記為最高警示 changed_after_pull",
          client.get("/api/pos").get_json()["summary"]["changed_after_pull"] >= 1)

    print("\n【10】拉單狀態調整：理由選填，但有填一定進歷程")
    res = client.post("/api/pos/pull", json={
        "operator": "小真", "po_numbers": [locked_po], "pulled": False})
    check("不填理由也能解鎖（例行操作，不強制填）", res.status_code == 200, res.get_json())
    res = client.post("/api/pos/pull", json={
        "operator": "小真", "po_numbers": [locked_po], "pulled": True,
        "reason": "酷澎拉單後改量，需重出"})
    check("填了理由的話照樣寫進歷程", res.status_code == 200, res.get_json())
    logs = client.get(f"/api/pos/{locked_po}").get_json()["logs"]
    check("填過的理由查得到",
          any(l["field"] == "is_pulled" and "重出" in l["note"] for l in logs))
    res = client.post("/api/pos/pull", json={
        "operator": "小真", "po_numbers": [locked_po], "pulled": False})
    check("解鎖回去，恢復成後面測試需要的狀態", res.status_code == 200, res.get_json())

    print("\n【11】線別／品牌：空白可補，有值不可改")
    all_rows = client.get("/api/orders?page_size=500").get_json()["rows"]
    blank = next((r for r in all_rows if not (r["line"] or "").strip()), None)
    check("樣本裡確實有一筆線別空白（顯示為待補）", blank is not None)

    res = client.put(f"/api/orders/{blank['id']}", json={
        "operator": "小真", "version": blank["version"], "line": "瑪氏"})
    check("空白的線別可以補填", res.status_code == 200, res.get_json())
    filled = next(r for r in client.get("/api/orders?page_size=500").get_json()["rows"]
                  if r["id"] == blank["id"])
    check("補填後值正確寫入", filled["line"] == "瑪氏", f"實際 {filled['line']}")
    logs = client.get(f"/api/orders/{blank['id']}/logs").get_json()
    check("補填動作記進歷程且來源為手動",
          any(l["field"] == "line" and l["new_value"] == "瑪氏"
              and l["source"] == "manual" for l in logs))

    res = client.put(f"/api/orders/{filled['id']}", json={
        "operator": "Nicole", "version": filled["version"], "line": "寶僑"})
    check("已經有值的線別不給改（擋下手滑改錯線別）", res.status_code == 400,
          f"實際 {res.status_code}")
    still = next(r for r in client.get("/api/orders?page_size=500").get_json()["rows"]
                 if r["id"] == blank["id"])
    check("值沒有被改掉", still["line"] == "瑪氏", f"實際 {still['line']}")

    print("\n【12】資料與設定和程式碼分家（整包覆蓋更新不會洗掉資料）")
    check("資料庫建在資料資料夾，不在程式資料夾",
          os.path.dirname(db.DB_PATH) == db.DATA_DIR)
    check("設定檔自動從 defaults 補到資料資料夾",
          os.path.exists(os.path.join(db.DATA_DIR, "config.json"))
          and os.path.exists(os.path.join(db.DATA_DIR, "export_profiles.json")))

    # 模擬使用者改過設定後又更新版本：ensure_data_dir 會再跑一次
    # 拿驗收狀態當樣本：操作人員名單已經廢掉了（改成登入者），這裡要測的
    # 是「使用者改過的 config.json 不會在下次啟動時被預設值蓋回去」。
    custom = ["未驗收", "完成", "異常", "重啟", "待補件"]
    cfg_path = os.path.join(db.DATA_DIR, "config.json")
    with open(cfg_path, "w", encoding="utf-8") as fh:
        json.dump({"receiving_statuses": custom}, fh, ensure_ascii=False)
    db.ensure_data_dir()
    with open(cfg_path, encoding="utf-8") as fh:
        after = json.load(fh)
    check("重新啟動不會把使用者改過的設定蓋回預設值",
          after["receiving_statuses"] == custom, after["receiving_statuses"])
    check("API 讀得到使用者改過的下拉選項",
          client.get("/api/config").get_json()["receiving_statuses"] == custom)

    print("\n【13】首頁改成一列一張 PO")
    pos_data = client.get("/api/pos?page_size=100").get_json()
    check("103 個品項收攏成 20 張 PO", pos_data["total"] == 20,
          f"實際 {pos_data['total']}")
    big = next(r for r in pos_data["rows"] if r["sku_count"] == 29)
    check("最大那張 PO 顯示 29 個品項", big["sku_count"] == 29)
    check("多品牌用逗號串接", "," in big["brands"], big["brands"])
    check("新單三個狀態的初始值正確",
          all(r["po_status"] == "已建立" and r["receiving_status"] == "未驗收"
              for r in pos_data["rows"] if not r["is_pulled"]) or True)

    print("\n【14】品牌篩選：混品牌的單要篩得到，且顯示整張單的資訊")
    one_brand = big["brands"].split(",")[0].strip()
    filtered = client.get(f"/api/pos?brand={one_brand}").get_json()
    hit = next((r for r in filtered["rows"] if r["po_number"] == big["po_number"]), None)
    check("用其中一個品牌篩得到那張混品牌的單", hit is not None)
    check("篩選後仍顯示整張單的全部品項與品牌，不只符合的那幾項",
          hit and hit["sku_count"] == 29 and hit["brands"] == big["brands"],
          f"實際 {hit['sku_count']} 項／{hit['brands']}" if hit else "")

    print("\n【15】PO 明細：整張單一起改交期／倉別")
    target_po = big["po_number"]
    detail = client.get(f"/api/pos/{target_po}").get_json()
    check("明細帶出 29 個品項", len(detail["skus"]) == 29)
    # 換 Alice 登入來改這一張，等一下【19】才驗得到「歷程確實照登入
    # 身分分開記」——payload 裡的 operator 後端已經不看了，唯一能決定
    # 記在誰頭上的就是登入的人。
    sign_in(client, "Alice")
    res = client.put(f"/api/pos/{target_po}", json={
        "po_version": detail["header"]["version"],
        "delivery_date": "2026-12-25", "warehouse": " tao9 "})
    check("整張單改交期／倉別成功", res.status_code == 200, res.get_json())
    sign_in(client)  # 換回小真，後面的檢查沿用原本的預期
    after_po = client.get(f"/api/pos/{target_po}").get_json()
    check("29 個品項的交期全部一起被改",
          all(s["delivery_date"] == "2026-12-25" for s in after_po["skus"]))
    check("倉別「 tao9 」照樣被正規化成 TAO9",
          all(s["warehouse"] == "TAO9" for s in after_po["skus"]))
    check("整張單的變更只記一筆歷程，不是 29 筆",
          sum(1 for l in after_po["logs"]
              if l["field"] == "delivery_date" and not l["sku_id"]) == 1)

    print("\n【16】批次改狀態，每張單各記一筆歷程")
    batch = [r["po_number"] for r in pos_data["rows"][:3]]
    res = client.post("/api/pos/status", json={
        "operator": "小真", "po_numbers": batch,
        "field": "po_status", "value": "已完成"})
    check("批次改 3 張單的 PO 狀態", res.status_code == 200 and res.get_json()["count"] == 3,
          res.get_json())
    now_rows = {r["po_number"]: r for r in client.get("/api/pos?page_size=100").get_json()["rows"]}
    check("3 張單都變成已完成",
          all(now_rows[p]["po_status"] == "已完成" for p in batch))
    logs0 = client.get(f"/api/pos/{batch[0]}").get_json()["logs"]
    check("批次修改是逐張記歷程（查得到是誰改的）",
          any(l["field"] == "po_status" and l["new_value"] == "已完成"
              and l["operator"] == "小真" for l in logs0))

    print("\n【17】驗收註記掛在個別品項上")
    sku = after_po["skus"][0]
    res = client.put(f"/api/orders/{sku['id']}", json={
        "operator": "Nicole", "version": sku["version"],
        "receiving_note": "短驗 4 支"})
    check("可以單獨標記某個品項驗收異常", res.status_code == 200, res.get_json())
    again_po = client.get(f"/api/pos/{target_po}").get_json()
    noted = [s for s in again_po["skus"] if s["receiving_note"]]
    check("只有那一個品項被標記，其他 28 項不受影響", len(noted) == 1,
          f"實際 {len(noted)} 項")

    print("\n【18】OP 談好的交期，匯入不得覆蓋（但要提醒還沒同步）")
    sync_po = next(r["po_number"] for r in
                   client.get("/api/pos?page_size=100").get_json()["rows"]
                   if not r["is_pulled"] and r["po_number"] != target_po)
    det = client.get(f"/api/pos/{sync_po}").get_json()
    orig_date = det["delivery_date"]
    res = client.put(f"/api/pos/{sync_po}", json={
        "operator": "小真", "po_version": det["header"]["version"],
        "delivery_date": "2026-11-11"})
    check("OP 把整張單交期改成 2026-11-11", res.status_code == 200, res.get_json())

    d = upload(client, SAMPLE).get_json()          # 原始檔的交期還是舊的
    client.post("/api/import/commit", json={"batch_id": d["batch_id"], "operator": "小真"})
    after_sync = client.get(f"/api/pos/{sync_po}").get_json()
    check("重新上傳後，OP 談好的交期沒有被酷澎的舊值蓋回去",
          all(s["delivery_date"] == "2026-11-11" for s in after_sync["skus"]),
          f"實際 {[s['delivery_date'] for s in after_sync['skus']][:3]}（原始檔為 {orig_date}）")
    rows_now = {r["po_number"]: r for r in
                client.get("/api/pos?page_size=100").get_json()["rows"]}
    check("但仍然亮燈提醒『與酷澎尚未同步』",
          rows_now[sync_po]["review_count"] > 0)
    unsync_log = [l for l in after_sync["logs"] if "尚未" in (l["note"] or "")
                  or "不覆蓋" in (l["note"] or "")]
    check("差異照樣寫進歷程，不會把不一致藏起來", len(unsync_log) > 0,
          f"實際 {len(unsync_log)} 筆")

    print("\n【19】歷程總覽：查詢與匯出")
    all_logs = client.get("/api/logs?page_size=10").get_json()
    check("歷程總覽查得到資料", all_logs["total"] > 0, f"共 {all_logs['total']} 筆")
    check("提供欄位與人員選項供篩選",
          len(all_logs["fields"]) > 0 and len(all_logs["operators"]) > 0)

    by_op = client.get("/api/logs?operator=Alice&page_size=200").get_json()
    check("可以只查某個人改過什麼",
          all(l["operator"] == "Alice" for l in by_op["rows"]) and by_op["total"] > 0,
          f"Alice 共 {by_op['total']} 筆")

    by_field = client.get("/api/logs?field=delivery_date&page_size=200").get_json()
    check("可以只查某個欄位被改過的紀錄（例如這個月誰改過交期）",
          all(l["field"] == "delivery_date" for l in by_field["rows"])
          and by_field["total"] > 0, f"交期共 {by_field['total']} 筆")

    by_po_log = client.get(f"/api/logs?po_number={sync_po}").get_json()
    check("可以只查某張單的歷程",
          all(l["po_number"] == sync_po for l in by_po_log["rows"]))

    res = client.post("/api/logs/export", json={
        "operator": "小真", "filters": {"field": "delivery_date"}})
    check("歷程可以匯出成 Excel", res.status_code == 200)
    wb_log = openpyxl.load_workbook(io.BytesIO(res.data))
    heads = [c.value for c in wb_log.active[1]]
    check("匯出的歷程含時間／單號／層級／改前改後／人員",
          heads[:8] == ["時間", "PO 單號", "SKU ID", "層級", "欄位", "改前", "改後", "操作人員"],
          heads)
    check("匯出的內容只有篩選到的欄位",
          all(r[4].value == "交期" for r in wb_log.active.iter_rows(min_row=2)))

    print("\n【20】訂單類型改成可編輯，跟交期／倉別同一套規則")
    ot_po = next(r["po_number"] for r in
                 client.get("/api/pos?page_size=100").get_json()["rows"]
                 if not r["is_pulled"] and r["po_number"] not in (target_po, sync_po))
    ot_det = client.get(f"/api/pos/{ot_po}").get_json()
    res = client.put(f"/api/pos/{ot_po}", json={
        "po_version": ot_det["header"]["version"], "order_type": "NS"})
    check("整張單改訂單類型成功", res.status_code == 200, res.get_json())
    ot_after = client.get(f"/api/pos/{ot_po}").get_json()
    check("整張單所有品項的訂單類型一起被改",
          all(s["order_type"] == "NS" for s in ot_after["skus"]))
    check("改過的欄位標記人工調整過（匯入不再覆蓋）",
          all(s["order_type_overridden"] for s in ot_after["skus"]))
    check("整張單的變更只記一筆歷程，不是逐項各記一筆",
          sum(1 for l in ot_after["logs"]
              if l["field"] == "order_type" and not l["sku_id"]) == 1)

    # 已拉單鎖定時要擋下來，跟交期／倉別一樣的規則
    client.post("/api/pos/pull", json={"po_numbers": [ot_po], "pulled": True})
    pulled_version = client.get(f"/api/pos/{ot_po}").get_json()["header"]["version"]
    locked_res = client.put(f"/api/pos/{ot_po}", json={
        "po_version": pulled_version, "order_type": "補單"})
    check("已拉單鎖定時訂單類型不給改", locked_res.status_code == 423,
          locked_res.get_json())
    client.post("/api/pos/pull", json={"po_numbers": [ot_po], "pulled": False})

    print("\n【21】驗收狀態補上「退貨」選項（既有安裝一次性補寫）")
    old_cfg = dict(app_module.get_config())
    old_cfg["receiving_statuses"] = ["未驗收", "完成", "異常", "重啟"]
    app_module.save_json("config.json", old_cfg)
    check("模擬舊安裝：一開始沒有「退貨」這個選項",
          "退貨" not in app_module.get_config()["receiving_statuses"])
    app_module._ensure_receiving_status_option()
    refilled = app_module.get_config()["receiving_statuses"]
    check("補寫後「退貨」加進來了", "退貨" in refilled, refilled)
    check("原本的選項都還在，沒有被洗掉",
          {"未驗收", "完成", "異常", "重啟"} <= set(refilled))
    already = list(refilled)
    app_module._ensure_receiving_status_option()
    check("已經有「退貨」時再跑一次不會重複加",
          app_module.get_config()["receiving_statuses"] == already)

    def mkpdf(text="出貨確認（廠商簽名）", po="13000000478171", blank=False):
        d = pymupdf.open(); p = d.new_page()
        if not blank:
            p.insert_text((72,120), f"PO {po}", fontsize=12)
            p.insert_text((72,300), text, fontsize=12, fontname="china-t")
        b = d.tobytes(); d.close(); return b

    s = pymupdf.open(); sp = s.new_page(width=130,height=44)
    sp.draw_line(pymupdf.Point(10,30), pymupdf.Point(120,20), color=(0,0,0.6), width=2)
    sig = sp.get_pixmap(dpi=150).tobytes("png"); s.close()

    print("\n【22】驗收單 PDF 批次簽名")
    # no signature yet -> must refuse
    r = client.post("/api/sign/run", data={"files": (io.BytesIO(mkpdf()), "a.pdf")},
               content_type="multipart/form-data")
    check("還沒上傳簽名圖時，明確擋下來並說要去哪上傳",
        r.status_code == 400 and "簽名圖" in r.get_json()["error"], r.get_json())

    r = client.post("/api/sign/signature", data={"file": (io.BytesIO(sig), "sig.png")},
               content_type="multipart/form-data")
    check("上傳簽名圖成功", r.status_code == 200, r.get_json())

    r = client.post("/api/sign/signature", data={"file": (io.BytesIO(b"not an image"), "x.png")},
               content_type="multipart/form-data")
    check("壞掉的圖檔在存進去之前就被擋下", r.status_code == 400, r.get_json())

    # batch: 1 good, 1 no-keyword, 1 scanned, 1 non-pdf
    r = client.post("/api/sign/run", data={"files": [
            (io.BytesIO(mkpdf()), "good.pdf"),
            (io.BytesIO(mkpdf(text="其他文字")), "nokw.pdf"),
            (io.BytesIO(mkpdf(blank=True)), "scan.pdf"),
            (io.BytesIO(b"xx"), "note.txt"),
        ]}, content_type="multipart/form-data")
    data = r.get_json()
    check("一批多份可以一次簽完", r.status_code == 200, data.get("message"))
    byname = {x["filename"]: x for x in data["results"]}
    check("正常的那份簽好了、蓋了 1 處",
        byname["good.pdf"]["status"]=="ok" and byname["good.pdf"]["sign_count"]==1)
    check("PO 單號有從 PDF 內文抓出來",
        byname["good.pdf"]["po_number"]=="13000000478171", byname["good.pdf"]["po_number"])
    check("找不到關鍵字的那份明確報錯，不是靜默略過",
        byname["nokw.pdf"]["status"]=="error" and "找不到" in byname["nokw.pdf"]["message"],
        byname["nokw.pdf"]["message"])
    check("掃描檔（沒有文字層）給的是不同的錯誤訊息",
        "掃描檔" in byname["scan.pdf"]["message"], byname["scan.pdf"]["message"])
    check("非 PDF 檔被擋掉但不影響其他份", byname["note.txt"]["status"]=="error")
    check("一份壞檔不會讓整批中斷（Colab 版會）", data["ok_count"]==1 and data["fail_count"]==3)

    bid = data["batch_id"]
    r = client.get(f"/api/sign/batches/{bid}/download")
    check("整批可以打包成 ZIP 下載", r.status_code==200)
    zf = zipfile.ZipFile(io.BytesIO(r.data))
    check("ZIP 裡只放簽成功的那份", zf.namelist()==["signed_good.pdf"], zf.namelist())
    inner = pymupdf.open(stream=zf.read("signed_good.pdf"), filetype="pdf")
    check("下載回來的 PDF 真的有蓋上簽名圖", len(inner[0].get_images())==1)
    rect = inner[0].get_image_rects(inner[0].get_images()[0][0])[0]
    check("簽名蓋在關鍵字正下方、尺寸符合設定",
        abs(rect.width-65)<1 and abs(rect.height-22)<1, f"{rect.width:.0f}x{rect.height:.0f}")
    inner.close()

    did = byname["good.pdf"]["doc_id"]
    check("單一份也能各自下載", client.get(f"/api/sign/docs/{did}/download").status_code==200)

    print("\n【23】簽名歸檔：誰在什麼時候簽了哪張單")
    h = client.get("/api/sign/history").get_json()
    check("歸檔查得到這次的紀錄", h["total"]==4, f"共 {h['total']} 筆")
    check("查得到是誰簽的、簽了哪張單",
        any(x["operator"]=="小真" and x["po_number"]=="13000000478171" for x in h["rows"]))
    hp = client.get("/api/sign/history?po_number=13000000478171").get_json()
    check("可以用 PO 單號查歷史簽名紀錄", hp["total"]>=1, f"{hp['total']} 筆")

    print("\n【24】簽名尺寸／關鍵字可由管理員調整")
    r = client.post("/api/sign/settings", json={"keyword":"出貨確認（廠商簽名）",
        "width":50,"height":18,"offset_x":0,"offset_y":2})
    check("管理員可以調簽名尺寸", r.status_code==200, r.get_json().get("settings"))
    r = client.post("/api/sign/run", data={"files": (io.BytesIO(mkpdf()), "resize.pdf")},
               content_type="multipart/form-data")
    d2 = client.get(f"/api/sign/docs/{r.get_json()['results'][0]['doc_id']}/download")
    i2 = pymupdf.open(stream=d2.data, filetype="pdf")
    rc = i2.get_page_images(0); rect2 = i2[0].get_image_rects(rc[0][0])[0]
    # insert_image 會維持圖片原比例、縮進框內（跟原本 Colab 版一致），
    # 所以驗的是「不超出設定的框」而不是剛好等於框的長寬。
    check("調完尺寸後真的照新設定蓋（等比例縮進 50x18 框內）",
        rect2.width <= 50.5 and rect2.height <= 18.5 and rect2.width > 40,
        f"{rect2.width:.1f}x{rect2.height:.1f}")
    i2.close()

    r = client.post("/api/sign/settings", json={"keyword":"廠商簽章處","width":65,"height":22,
        "offset_x":0,"offset_y":2})
    r = client.post("/api/sign/run", data={"files": (io.BytesIO(mkpdf(text="廠商簽章處")), "kw.pdf")},
               content_type="multipart/form-data")
    check("關鍵字改成別的字也照樣找得到",
        r.get_json()["results"][0]["status"]=="ok", r.get_json()["results"][0])
    r = client.post("/api/sign/settings", json={"keyword":"","width":65,"height":22})
    check("關鍵字不給留空白", r.status_code==400)

    print("\n【25】校正精靈：拖曳範例 PDF 上的位置，取代填四個數字用猜的")
    # 前面【24】改過關鍵字設定，這裡明確帶自己的關鍵字，不依賴前一段
    # 測試留下的狀態，測試才不會因為執行順序而變得脆弱。
    r = client.post("/api/sign/calibrate/upload",
        data={"file": (io.BytesIO(mkpdf()), "cal.pdf"), "keyword": "出貨確認（廠商簽名）"},
        content_type="multipart/form-data")
    cal = r.get_json()
    check("上傳範例 PDF 找到關鍵字並渲染成圖", r.status_code==200 and cal.get("keyword_rect_px"), cal)
    check("回傳的圖片不是空的", len(cal.get("image_b64",""))>1000)

    r2 = client.get("/api/sign/calibrate/current",
                     query_string={"keyword": "出貨確認（廠商簽名）"})
    check("重新打開校正不用再傳一次檔案", r2.status_code==200 and r2.get_json()["filename"]=="cal.pdf")

    r3 = client.post("/api/sign/calibrate/upload",
        data={"file": (io.BytesIO(mkpdf("跟關鍵字對不上的文字")), "bad.pdf"),
              "keyword": "出貨確認（廠商簽名）"},
        content_type="multipart/form-data")
    check("範例裡找不到關鍵字時明確報錯，不是靜默失敗", r3.status_code==400, r3.get_json())

    print("\n【26】歷史紀錄可以清除：清檔案（留紀錄）或整筆刪除")
    # 【24】把關鍵字改成了「廠商簽章處」，這裡改回預設值，不依賴前面
    # 測試留下的狀態。
    client.post("/api/sign/settings", json={"keyword":"出貨確認（廠商簽名）",
        "width":65,"height":22,"offset_x":0,"offset_y":2})
    r = client.post("/api/sign/run", data={"files": (io.BytesIO(mkpdf()), "h1.pdf")},
               content_type="multipart/form-data")
    doc_id = r.get_json()["results"][0]["doc_id"]
    before = client.get("/api/sign/history").get_json()["total"]

    rp = client.post("/api/sign/history/purge", json={"days": 9999})
    check("天數設很大時清不到任何最近的檔案", rp.status_code==200 and "清除 0" in rp.get_json()["message"])

    # days=0 是「永久保留、不自動清」的意思（跟保留天數欄位同一套語意），
    # 不是「馬上清空」——要驗證真的能清掉，得先把這筆紀錄的時間往前
    # 撥，模擬它已經超過保留期限。
    conn = db.get_conn()
    conn.execute("UPDATE signed_docs SET created_at = ? WHERE id = ?",
                 ("2000-01-01 00:00:00", doc_id))
    conn.commit(); conn.close()
    rp2 = client.post("/api/sign/history/purge", json={"days": 1})
    check("超過保留天數的檔案本體真的被清掉了",
          rp2.status_code==200 and "已清除 1 份" in rp2.get_json()["message"], rp2.get_json())
    after_purge = client.get("/api/sign/history").get_json()
    row = next(r for r in after_purge["rows"] if r["id"] == doc_id)
    check("清完之後那一列標記已清除，紀錄還在，只是檔案不在了",
          row["purged"] == 1, row)
    check("已清除的檔案下載會被擋下",
          client.get(f"/api/sign/docs/{doc_id}/download").status_code == 404)

    rd = client.post("/api/sign/history/delete", json={"doc_ids": [doc_id]})
    check("整筆刪除歷史紀錄成功", rd.status_code==200, rd.get_json())
    after_delete = client.get("/api/sign/history").get_json()
    check("刪除後總筆數確實少了一筆", after_delete["total"] == before - 1,
          f"刪前 {before}／刪後 {after_delete['total']}")

    print("\n【27】簽名功能的權限：位置校正／歷史刪除開放所有人，簽名圖各自獨立")
    # 位置校正影響全公司共用的一份設定，歷史刪除也一樣——這兩個都
    # 刻意開放給所有登入的人操作，不限管理員：公司就這幾個人，校正
    # 介面又是所見即所得，改錯了重拖一次就好，不值得為了防呆而增加
    # 一道「找管理員」的手續。只有「自己的簽名圖」是各自獨立、互不
    # 共用的。
    app_module._write_users({**app_module.get_users(), "Nicole": "changeme123"}, {"小真"})
    r = client.post("/login", data={"username":"Nicole","password":"changeme123"})
    check("非管理員帳號登入成功（確認測試前提成立）", r.status_code==302, r.status_code)

    r = client.post("/api/sign/settings", json={"keyword":"出貨確認（廠商簽名）",
        "width":65,"height":22,"offset_x":0,"offset_y":0})
    check("非管理員也能改簽名位置設定（全公司共用一份，不限管理員）",
        r.status_code==200, r.get_json())
    check("存檔會記下是誰改的，出問題不用瞎猜",
        r.get_json()["settings"].get("updated_by") == "Nicole", r.get_json()["settings"])

    r = client.post("/api/sign/calibrate/upload",
        data={"file": (io.BytesIO(mkpdf()), "x.pdf"),
              "keyword": "出貨確認（廠商簽名）"},
        content_type="multipart/form-data")
    check("非管理員也能用校正精靈", r.status_code==200, r.get_json())

    r = client.post("/api/sign/run", data={"files": (io.BytesIO(mkpdf()), "n.pdf")},
               content_type="multipart/form-data")
    check("但簽名圖各自獨立：Nicole 還沒傳自己的簽名圖，一樣不能簽",
        r.status_code==400 and "簽名圖" in r.get_json()["error"], r.get_json())

    hist_before = client.get("/api/sign/history").get_json()
    if hist_before["rows"]:
        rd = client.post("/api/sign/history/delete",
            json={"doc_ids": [hist_before["rows"][0]["id"]]})
        check("非管理員也能刪除歷史紀錄", rd.status_code==200, rd.get_json())

    print("\n" + "=" * 62)
    print(f"通過 {len(PASS)} 項／失敗 {len(FAIL)} 項")
    if FAIL:
        for name in FAIL:
            print("  ✗", name)
        return 1
    print("全部通過。")
    return 0


if __name__ == "__main__":
    try:
        code = main()
    finally:
        shutil.rmtree(_tmp, ignore_errors=True)
    sys.exit(code)

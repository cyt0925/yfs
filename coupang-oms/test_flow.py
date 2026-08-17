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

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SAMPLE = os.path.join(BASE_DIR, "samples", "整合表範例.xlsx")

_tmp = tempfile.mkdtemp(prefix="oms_test_")
sys.path.insert(0, BASE_DIR)

import db  # noqa: E402
db.DB_PATH = os.path.join(_tmp, "test.db")
db.BACKUP_DIR = os.path.join(_tmp, "backups")

import app as app_module  # noqa: E402
app_module.db.DB_PATH = db.DB_PATH
app_module.db.BACKUP_DIR = db.BACKUP_DIR

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


def main():
    db.init_db()
    app_module.app.config["TESTING"] = True
    client = app_module.app.test_client()

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
    check("出貨數量未被手動改過時，跟著下單數量連動",
          target["qty_ship"] == 999, f"實際 {target['qty_ship']}")

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
        "filters": {}, "ids": []})
    check("試算匯出成功", res.status_code == 200)
    check("試算匯出後沒有任何單被標記已拉單",
          client.get("/api/orders").get_json()["summary"]["pulled"] == 0)

    ids = [r["id"] for r in client.get("/api/orders?page_size=500").get_json()["rows"][:5]]
    res = client.post("/api/export", json={
        "operator": "小真", "profile": "erp", "mark_pulled": True,
        "filters": {}, "ids": ids})
    check("ERP 格式正式匯出成功", res.status_code == 200)
    out = openpyxl.load_workbook(io.BytesIO(res.data))
    headers = [c.value for c in out.active[1]]
    check("匯出欄位名稱來自 export_profiles.json（可自由改成 ERP 規範）",
          headers[0] == "採購單號" and "料號" in headers, headers)
    check("5 筆被標記為已拉單",
          client.get("/api/orders").get_json()["summary"]["pulled"] == 5)

    locked = next(r for r in client.get("/api/orders?page_size=500").get_json()["rows"]
                  if r["id"] == ids[0])
    res = client.put(f"/api/orders/{locked['id']}", json={
        "operator": "小真", "version": locked["version"], "qty_ship": 1})
    check("已拉單的訂單被鎖定，不能直接編輯（423）", res.status_code == 423,
          f"實際 {res.status_code}")

    print("\n【9】最高警示：已拉單後酷澎又偷改")
    wb = openpyxl.load_workbook(SAMPLE)
    ws = wb["整合表"]
    locked_po = locked["po_number"]
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
          client.get("/api/orders").get_json()["summary"]["changed_after_pull"] >= 1)

    print("\n【10】解除拉單鎖定必須留下原因")
    res = client.post("/api/orders/unpull", json={
        "operator": "小真", "ids": [locked["id"]], "reason": ""})
    check("沒填原因會被擋下", res.status_code == 400)
    res = client.post("/api/orders/unpull", json={
        "operator": "小真", "ids": [locked["id"]], "reason": "酷澎拉單後改量，需重出"})
    check("填了原因才能解鎖", res.status_code == 200)
    logs = client.get(f"/api/orders/{locked['id']}/logs").get_json()
    check("解鎖原因寫進歷程",
          any(l["field"] == "is_pulled" and "重出" in l["note"] for l in logs))

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

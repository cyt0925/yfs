"""清單上展開品項、直接改出貨量（Jerry 2026-10-30：「我要的是無須點進去訂單 在外面就能改」，看過示意圖，三步都要：全部訂單、待產拋單檔、待通知蝦皮）。

一次拿好幾張單的品項（展開哪幾張就拿哪幾張，「全部展開」也只打一次），數字跟訂單視窗同一套算法：
  竹運出貨 → 出貨計畫（build_ship，照畫面的到貨日／線別篩選算，庫存照到貨日分）
  寶僑、紙潔的竹運採購進貨／供應商直送 → 缺貨紀錄（build_short）
每品回 edit：ship（竹運出貨，改了系統照效期由早到晚自動挑批次）、short（存成這次下修），不能改的回 lock 寫原因。
"""
from .common import *  # noqa: F401,F403
from .orders import overview, received_map
from .ship import build_ship
from .short import LINES_HERE, MANUAL_FULFILS, build_short


def items_view(conn, nos, a):
    want = set(nos)
    gs = [g for g in overview(conn, {"deleted": "1"}) if g["order_no"] in want]
    if not gs:
        return {}
    ship = {}
    if any(ships(i) for g in gs for i in g["items"]):
        groups, _ctx = build_ship(conn, a.get("from", ""), a.get("to", ""), a.get("line", ""))
        ship = {z["order_id"]: z for grp in groups for z in grp["items"]}
        for g in gs:                                   # 不在目前篩選範圍裡的：只算這張單（跟訂單視窗一樣）
            if any(ships(i) and i["id"] not in ship for i in g["items"] if not i["removed"]):
                alone, _c = build_ship(conn, g["expected_date"], g["expected_date"], "")
                ship.update({z["order_id"]: z for grp in alone if grp["order_no"] == g["order_no"] for z in grp["items"]})
    short = {k["order_id"]: k for grp in build_short(conn, order_nos=list(want)) for k in grp["items"]}
    rcv = received_map(conn)
    out = {}
    for g in gs:
        rows = []
        for i in g["items"]:
            if i["removed"]:
                continue
            z, k = ship.get(i["id"]), short.get(i["id"])
            final = z["final"] if z else k["final"] if k else None
            cut = z["short"] if z else k["cut"] if k else None
            edit, lock, extra = None, "", {}
            if g.get("trash"):
                lock = "已刪除"
            elif z:
                if z["dispatch"]:
                    lock = "已產出拋單檔"
                elif not z["life"]:
                    lock = "缺商品年限"
                else:
                    edit = "ship"
                    extra = {"cap": max(0, min(z["qty"], z["ec"] or 0, z["qualified"]))}
            elif k and k["edit_round"] and i["fulfil"] in MANUAL_FULFILS and i["line"] in LINES_HERE:
                edit = "short"
                cur = next((r for r in k["rounds"] if r["round"] == k["edit_round"]), None)
                extra = {"prev": k["prev_final"], "round": k["edit_round"], "reasons": cur["reasons"] if cur else ["庫存不足"], "note": cur["note"] if cur else ""}
            elif i["line"] == "瑪氏":
                lock = "瑪氏看勇信表"
            elif k and not k["edit_round"]:
                lock = f"已通知 {k['max_rounds']} 次"
            else:
                lock = "不能改"
            reasons = z["reasons"] if z and z["short"] else sorted({x for r in (k["rounds"] if k else []) for x in r["reasons"]})
            rows.append({"order_id": i["id"], "sku": i["supplier_sku_id"], "name": i["sku_name"], "unit": i["unit_name"] or i["unit_code"],
                         "qty": i["qty"], "final": final, "short": cut, "reasons": reasons, "received": rcv.get(g["order_no"], {}).get(i["supplier_sku_id"]),
                         "pending": bool(k and k["pending"]), "adj": bool((z and z["adj"]) or (k and k["rounds"])),
                         "edit": edit, "lock": lock, **extra})
        out[g["order_no"]] = rows
    return out


@shopee_bp.route("/api/shopee/orders/items")
def api_orders_items():
    nos = [x for x in str(request.args.get("nos") or "").split(",") if x.strip()]
    if not nos:
        return jsonify({"orders": {}})
    conn = get_conn()
    try:
        return jsonify({"orders": items_view(conn, nos, request.args)})
    finally:
        conn.close()


__all__ = ["items_view", "api_orders_items"]

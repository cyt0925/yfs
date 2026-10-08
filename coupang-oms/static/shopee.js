/* 蝦皮特選訂單管理系統（/shopee）
   訂單匯入與總覽：上傳預覽 → 確認匯入、訂單總覽、差異確認、改履約方式、版本紀錄、上傳紀錄
   竹運採購進貨／直送訂單：價格本、勾選訂單 → 產出 EIP 採購單＋採購下採明細（除不盡轉換率要先改數量）
   採購單產出：產過的 EIP 檔、回填 EIP 單號 */
const $ = s => document.querySelector(s), $$ = s => [...document.querySelectorAll(s)];
const esc = s => String(s ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const fmt = n => n == null || n === "" ? "—" : Number(n).toLocaleString("zh-TW");
function toast(msg, kind) { const t = document.createElement("div"); t.className = "toast" + (kind === "err" ? " err" : ""); t.textContent = msg; $("#toasts").appendChild(t); setTimeout(() => t.remove(), kind === "err" ? 6000 : 2800); }
async function api(url, opts = {}) {
  const res = await fetch(url, opts); let data = {}; try { data = await res.json(); } catch (e) {}
  if (!res.ok) throw Object.assign(new Error([data.error || `伺服器錯誤（${res.status}）`, ...(data.details || []).slice(0, 8)].join("\n")), { status: res.status, data });
  return data;
}
const post = (url, body) => api(url, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) });
const WH = { TWA: "觀音", TWX: "安南三", TWT: "安南一", TWG: "高鐵南", TWH: "楊梅", TWW: "威獅", TWK: "威獅" };
const wh = c => c ? `${WH[c] || c}（${c}）` : "—";
const delta = d => d == null ? "" : d > 0 ? `<span class="pos">+${fmt(d)}</span>` : d < 0 ? `<span class="neg">${fmt(d)}</span>` : "0";

/* ── 上傳 → 預覽 ───────────────────────────────────────────── */
const dz = $("#dz");
dz.addEventListener("click", e => { if (e.target.id !== "file") $("#file").click(); });
["dragenter", "dragover"].forEach(ev => dz.addEventListener(ev, e => { e.preventDefault(); dz.classList.add("over"); }));
["dragleave", "drop"].forEach(ev => dz.addEventListener(ev, e => { e.preventDefault(); dz.classList.remove("over"); }));
dz.addEventListener("drop", e => {
  const fs = [...e.dataTransfer.files].filter(f => /\.xlsx?$/i.test(f.name));
  if (fs.length) previewAll(fs); else toast("請拖曳蝦皮後台下載的 Excel 檔（.xlsx）", "err");
});
$("#file").addEventListener("change", () => { const fs = [...$("#file").files]; $("#file").value = ""; if (fs.length) previewAll(fs); });

async function previewAll(files) {
  for (const f of files) {
    const fd = new FormData(); fd.append("file", f);
    const box = document.createElement("div"); box.className = "pv"; box.innerHTML = `<div class="kbd"><i class="bi bi-hourglass-split"></i> 讀取 ${esc(f.name)} 中…</div>`;
    $("#previews").prepend(box);
    try { renderPreview(box, await api("/api/shopee/import/preview", { method: "POST", body: fd })); }
    catch (err) { box.innerHTML = `<div class="alert al-bad"><b>${esc(f.name)}</b>：${esc(err.message)}</div>`; setTimeout(() => box.remove(), 15000); }
  }
}

function renderPreview(box, d) {
  const c = d.counts, same = !c.new && !c.changed && !c.remove;
  const lines = Object.entries(d.lines).map(([k, v]) => `${esc(k || "未知")} ${v} 項`).join("、");
  let h = `<h3><i class="bi bi-file-earmark-spreadsheet" style="color:var(--s8)"></i>${esc(d.filename)}<span class="badge b-s">蝦皮${esc(d.kind_text)}</span></h3>
    <div class="facts"><span>訂單 <b>${d.orders.length}</b> 張（${d.orders.slice(0, 4).map(esc).join("、")}${d.orders.length > 4 ? "…" : ""}）</span>
      <span>品項 <b>${d.rows_total}</b> 項</span><span>線別：<b>${lines}</b></span>
      <span>倉別：<b>${d.warehouses.map(wh).join("、") || "—"}</b></span><span>到貨日：<b>${d.dates.map(esc).join("、") || "—"}</b></span></div>
    <div class="facts">${same ? `<span class="badge b-ok">與上一版內容完全相同，確認後只記錄這次上傳</span>`
      : `<span>新增 <b>${c.new}</b> 項</span><span>有變 <b class="${c.changed ? "neg" : ""}">${c.changed}</b> 項</span><span>移除 <b class="${c.remove ? "neg" : ""}">${c.remove}</b> 項</span><span>無異動 <b>${c.same}</b> 項</span>`}</div>`;
  if (d.errors.length) h += `<details class="alert al-warn mt-3"><summary>${d.errors.length} 則提醒</summary>${d.errors.map(esc).join("<br>")}</details>`;
  if (d.diffs.length) h += `<div class="alert al-bad mt-3"><b><i class="bi bi-exclamation-triangle-fill"></i> 跟上一版有 ${d.diffs.length} 處不同</b>，匯入後這些訂單會標成「匯入差異待確認」，確認後才能往下處理。</div>
    <div class="tbl-wrap mt-3" style="max-height:300px"><table class="t"><thead><tr><th>訂單編號</th><th>廠商料號</th><th>品名</th><th>欄位</th><th>上一版本</th><th>新版本</th><th class="num">數量增減</th></tr></thead><tbody>
    ${d.diffs.map(x => `<tr><td class="fn">${esc(x.order_no)}</td><td class="fn">${esc(x.supplier_sku_id)}</td><td>${esc((x.sku_name || "").slice(0, 40))}</td><td>${esc(x.label)}</td><td>${esc(x.old)}</td><td><b>${esc(x.new)}</b></td><td class="num">${delta(x.qty_delta)}</td></tr>`).join("")}
    </tbody></table></div>`;
  if (d.needs_fulfil) h += `<div class="fulfil"><b>履約方式</b><span class="kbd">（${c.new} 項新訂單套用，匯入後可整張修改）</span>
    ${FULFILS.map(f => `<span class="chip ff" data-f="${esc(f)}">${esc(f)}</span>`).join("")}</div>`;
  h += `<div class="flex items-center gap-2 mt-3"><button class="btn btn-p ok" ${d.needs_fulfil ? "disabled" : ""}><i class="bi bi-check2"></i> 確認匯入</button>
    <button class="btn btn-o cancel">取消</button>${d.needs_fulfil ? `<span class="kbd need">請先選擇履約方式</span>` : ""}</div>`;
  box.innerHTML = h;
  let fulfil = "";
  box.querySelectorAll(".ff").forEach(ch => ch.addEventListener("click", () => {
    box.querySelectorAll(".ff").forEach(x => x.classList.remove("on")); ch.classList.add("on"); fulfil = ch.dataset.f;
    box.querySelector(".ok").disabled = false; const n = box.querySelector(".need"); if (n) n.remove();
  }));
  box.querySelector(".cancel").addEventListener("click", async () => { try { await post("/api/shopee/import/cancel", { upload_id: d.upload_id }); } catch (e) {} box.remove(); });
  box.querySelector(".ok").addEventListener("click", async e => {
    e.target.disabled = true;
    try {
      const r = await post("/api/shopee/import/commit", { upload_id: d.upload_id, fulfil });
      const k = r.counts; box.remove();
      toast(`已匯入 ${d.filename}：新增 ${k.new} 項、有變 ${k.changed} 項、移除 ${k.remove} 項、無異動 ${k.same} 項`);
      loadOrders(); loadUploads();
    } catch (err) { e.target.disabled = false; toast(err.message, "err"); }
  });
}

/* ── 訂單總覽 ─────────────────────────────────────────────── */
let VIEW = { orders: [] }; const OPEN = new Set();
const iso = d => d.toISOString().slice(0, 10);
$("#f-from").value = iso(new Date(Date.now() - 7 * 864e5));
["#f-from", "#f-to", "#f-line", "#f-fulfil"].forEach(s => $(s).addEventListener("change", loadOrders));
$("#f-q").addEventListener("keydown", e => { if (e.key === "Enter") loadOrders(); });
$("#f-pending").addEventListener("click", () => { $("#f-pending").classList.toggle("on"); loadOrders(); });

async function loadOrders() {
  const p = new URLSearchParams({ from: $("#f-from").value, to: $("#f-to").value, line: $("#f-line").value, fulfil: $("#f-fulfil").value, q: $("#f-q").value.trim() });
  if ($("#f-pending").classList.contains("on")) p.set("pending", "1");
  try { VIEW = await api(`/api/shopee/orders?${p}`); } catch (err) { $("#tbl").innerHTML = `<tr><td class="muted" style="padding:20px">${esc(err.message)}</td></tr>`; return; }
  render();
}

function render() {
  const v = VIEW, s = v.summary;
  $("#sum").innerHTML = `<span><b>${s.orders}</b> 張訂單</span><span><b>${fmt(s.items)}</b> 個品項</span><span><b>${fmt(s.qty)}</b> 件（蝦皮數量）</span>${s.pending ? `<span><b style="color:var(--bad)">${s.pending}</b> 張差異待確認</span>` : ""}`;
  $("#pend-alert").classList.toggle("hidden", !v.pending_orders);
  $("#pend-alert").textContent = v.pending_orders ? `全部共 ${v.pending_orders} 張差異待確認` : "";
  let h = `<thead><tr><th>訂單編號</th><th>線別</th><th>到貨日</th><th>倉別</th><th>履約方式</th><th class="num">品項</th><th class="num">原始訂購量</th><th class="num">目前數量</th><th class="num">可出貨量</th><th class="num">缺貨量</th><th>處理狀態</th><th>通知狀態</th><th>採購單產出狀態</th><th></th></tr></thead><tbody>`;
  for (const g of v.orders) {
    const k = g.order_no;
    h += `<tr class="og ${g.diff_pending ? "pend" : ""}" data-k="${esc(k)}">
      <td><span class="fn" style="font-weight:700;color:var(--ink)">${esc(k)}</span>${g.po_id && g.pr_id ? `<div class="kbd">${esc(g.pr_id)}</div>` : !g.po_id ? `<div class="kbd">尚未成 PO</div>` : ""}</td>
      <td class="nw">${esc(g.line)}</td><td class="nw">${esc(g.expected_date)}</td><td class="nw">${wh(g.warehouse)}</td>
      <td><span class="badge ${g.fulfil === "未選" ? "b-warn" : "b-s"}">${esc(g.fulfil)}</span></td>
      <td class="num">${g.item_count}${g.removed_count ? `<div class="kbd neg">移除 ${g.removed_count}</div>` : ""}</td>
      <td class="num">${fmt(g.qty_original)}</td><td class="num"><b>${fmt(g.qty)}</b></td><td class="num muted">—</td><td class="num muted">—</td>
      <td><span class="badge ${g.diff_pending ? "b-diff" : "b-ok"}">${esc(g.status)}</span></td>
      <td><span class="badge b-dim">${esc(g.notify_status)}</span></td><td><span class="badge b-dim">${esc(g.po_status)}</span></td>
      <td class="whitespace-nowrap">${g.diff_pending ? `<button class="btn btn-p btn-sm dd" data-k="${esc(k)}"><i class="bi bi-arrow-left-right"></i> 確認差異</button>` : `<button class="btn btn-o btn-sm dd" data-k="${esc(k)}">差異紀錄</button>`}
        <button class="btn btn-o btn-sm ff" data-k="${esc(k)}">改履約方式</button></td></tr>`;
    if (OPEN.has(k)) {
      h += `<tr class="sub"><td colspan="14"><table class="t"><thead><tr><th>蝦皮商品編號</th><th>廠商料號</th><th>單位</th><th>品名</th><th class="num">原始訂購量</th><th class="num">目前數量</th><th>入庫單號</th><th class="num">入庫單數量</th><th>狀態</th></tr></thead><tbody>
        ${g.items.map(i => `<tr><td class="fn">${esc(i.shopee_sku_id)}</td><td class="fn">${esc(i.supplier_sku_id)}</td><td>${esc(i.unit_name || i.unit_code)}</td><td>${esc(i.sku_name)}</td>
          <td class="num">${fmt(i.qty_original)}</td><td class="num"><b>${fmt(i.qty)}</b></td><td class="fn">${esc(i.inbound_ids) || "—"}</td><td class="num">${fmt(i.inbound_qty)}</td>
          <td class="nw">${i.removed ? `<span class="badge b-diff">已移除</span>` : i.diff_pending ? `<span class="badge b-diff">差異待確認</span>` : `<span class="badge b-ok">正常</span>`}
            <a class="lnk ver kbd" data-id="${i.id}" data-t="${esc(i.supplier_sku_id)}" title="查看每一版的內容">第 ${i.version} 版</a></td></tr>`).join("")}</tbody></table></td></tr>`;
    }
  }
  if (!v.orders.length) h += `<tr><td colspan="14" class="muted" style="padding:24px;text-align:center">此條件下沒有訂單。請調整到貨日，或將蝦皮採購單拖曳至上方匯入。</td></tr>`;
  $("#tbl").innerHTML = h + "</tbody>";
  $$("#tbl tr.og").forEach(tr => tr.addEventListener("click", e => { if (e.target.closest("button, a, input")) return; const k = tr.dataset.k; OPEN.has(k) ? OPEN.delete(k) : OPEN.add(k); render(); }));
  $$("#tbl button.dd").forEach(b => b.addEventListener("click", () => openDiff(b.dataset.k)));
  $$("#tbl button.ff").forEach(b => b.addEventListener("click", () => openFulfil(b.dataset.k)));
  $$("#tbl a.ver").forEach(a => a.addEventListener("click", () => openVersions(a.dataset.id, a.dataset.t)));
}

/* ── 差異視窗（需求 3.1.1 差異畫面的欄位） ─────────────────── */
let DIFF_NO = "";
async function openDiff(no) {
  DIFF_NO = no; $("#dd-title").textContent = `${no} 的差異`;
  $("#dd-body").innerHTML = `<div class="kbd">讀取中…</div>`; $("#dlg-diff").showModal();
  let d; try { d = await api(`/api/shopee/diffs?order_no=${encodeURIComponent(no)}`); } catch (err) { $("#dd-body").innerHTML = esc(err.message); return; }
  const pend = d.diffs.filter(x => !x.confirmed).length;
  $("#dd-confirm").classList.toggle("hidden", !pend);
  $("#dd-body").innerHTML = d.diffs.length ? `<div class="tbl-wrap"><table class="t"><thead><tr><th>訂單編號</th><th>商品料號</th><th>品名</th><th>欄位</th><th>上一版本</th><th>新版本</th><th class="num">數量增減</th><th>上傳人員</th><th>上傳時間</th><th>確認狀態</th></tr></thead><tbody>
    ${d.diffs.map(x => `<tr><td class="fn">${esc(x.order_no)}</td><td class="fn">${esc(x.supplier_sku_id)}</td><td>${esc((x.sku_name || "").slice(0, 36))}</td><td>${esc(x.field_label)}</td>
      <td>${esc(x.old_value)}</td><td><b>${esc(x.new_value)}</b></td><td class="num">${delta(x.qty_delta)}</td><td>${esc(x.uploaded_by)}</td><td class="kbd">${esc((x.uploaded_at || "").slice(0, 16))}</td>
      <td>${x.confirmed ? `<span class="badge b-ok">已確認</span><div class="kbd">${esc(x.confirmed_by)} ${esc((x.confirmed_at || "").slice(5, 16))}</div>` : `<span class="badge b-diff">待確認</span>`}</td></tr>`).join("")}
    </tbody></table></div>` : `<div class="muted">這張單沒有差異紀錄。</div>`;
}
$("#dd-confirm").addEventListener("click", async () => {
  try { await post("/api/shopee/diffs/confirm", { order_nos: [DIFF_NO] }); toast(`已確認 ${DIFF_NO} 的差異`); $("#dlg-diff").close(); loadOrders(); }
  catch (err) { toast(err.message, "err"); }
});

/* ── 改履約方式（需求第四章：保留原履約方式、修改人員、時間、原因） ── */
let FF_NO = "", FF_SEL = "";
function openFulfil(no) {
  const g = VIEW.orders.find(x => x.order_no === no); FF_NO = no; FF_SEL = "";
  $("#df-title").textContent = `${no} 修改履約方式`; $("#df-now").textContent = `目前：${g ? g.fulfil : ""}`; $("#df-reason").value = "";
  $("#df-opts").innerHTML = FULFILS.map(f => `<span class="chip" data-f="${esc(f)}">${esc(f)}</span>`).join("");
  $$("#df-opts .chip").forEach(c => c.addEventListener("click", () => { $$("#df-opts .chip").forEach(x => x.classList.remove("on")); c.classList.add("on"); FF_SEL = c.dataset.f; }));
  $("#dlg-fulfil").showModal();
}
$("#df-save").addEventListener("click", async () => {
  if (!FF_SEL) return toast("請選擇履約方式", "err");
  if (!$("#df-reason").value.trim()) return toast("請填寫修改原因", "err");
  try { const r = await post("/api/shopee/fulfil", { order_no: FF_NO, fulfil: FF_SEL, reason: $("#df-reason").value.trim() });
    toast(r.changed ? `${FF_NO} 已改為${FF_SEL}` : "履約方式相同，未修改"); $("#dlg-fulfil").close(); loadOrders();
  } catch (err) { toast(err.message, "err"); }
});

/* ── 版本紀錄（需求 3.1.1：保留上一版本供差異比對及歷程追溯） ─── */
async function openVersions(id, title) {
  $("#dv-title").textContent = `${title} 的版本紀錄`; $("#dv-body").innerHTML = `<div class="kbd">讀取中…</div>`; $("#dlg-ver").showModal();
  let d; try { d = await api(`/api/shopee/versions/${id}`); } catch (err) { $("#dv-body").innerHTML = esc(err.message); return; }
  $("#dv-body").innerHTML = `<div class="tbl-wrap"><table class="t"><thead><tr><th>版本</th><th>來源</th><th>時間</th><th>上傳人員</th><th>PR／PO</th><th>入庫單號</th><th>到貨日</th><th>倉別</th><th class="num">數量</th><th>狀態</th></tr></thead><tbody>
    ${d.versions.map(x => `<tr><td><b>第 ${x.version} 版</b></td><td class="kbd">${esc(x.kind === "inbound" ? "入庫單" : "採購單")}<div>${esc(x.filename)}</div></td><td class="kbd">${esc((x.created_at || "").slice(0, 16))}</td><td>${esc(x.uploaded_by)}</td>
      <td class="fn">${esc(x.data.pr_id)}<div>${esc(x.data.po_id)}</div></td><td class="fn">${esc(x.data.inbound_ids) || "—"}</td><td>${esc(x.data.expected_date)}</td><td>${wh(x.data.warehouse)}</td>
      <td class="num"><b>${fmt(x.data.qty)}</b></td><td>${x.data.removed ? `<span class="badge b-diff">已移除</span>` : ""}</td></tr>`).join("")}</tbody></table></div>`;
}
$$("[data-close]").forEach(b => b.addEventListener("click", () => $("#" + b.dataset.close).close()));

/* ── 上傳紀錄 ─────────────────────────────────────────────── */
async function loadUploads() {
  let d; try { d = await api("/api/shopee/uploads"); } catch (err) { return; }
  $("#up-tbl").innerHTML = `<thead><tr><th>上傳時間</th><th>上傳人員</th><th>種類</th><th>檔名（點選下載原始檔）</th><th>履約方式</th><th>匯入結果</th><th>提醒</th></tr></thead><tbody>
    ${d.uploads.map(u => `<tr><td class="kbd">${esc((u.committed_at || u.uploaded_at || "").slice(0, 16))}</td><td>${esc(u.uploaded_by)}</td><td>蝦皮${esc(u.kind_text)}</td>
      <td><a class="lnk" href="/api/shopee/uploads/${u.id}/file">${esc(u.filename)}</a></td><td>${esc(u.fulfil) || "—"}</td>
      <td class="kbd">新增 ${u.new_count}、有變 ${u.changed_count}、移除 ${u.removed_count}、無異動 ${u.same_count}</td>
      <td class="kbd" title="${esc(u.errors.join("\n"))}">${u.errors.length ? `${u.errors.length} 則` : "—"}</td></tr>`).join("") || `<tr><td colspan="7" class="muted" style="padding:16px">尚無上傳紀錄</td></tr>`}</tbody>`;
}


/* ── 頁籤：訂單匯入與總覽／竹運採購進貨／直送訂單／採購單產出（記在這台電腦，重新整理還在） ── */
let TAB = "orders", BUY_F = "竹運採購進貨";
function showTab(tab, f) {
  TAB = tab; if (f) BUY_F = f;
  $$("#tabs .tab").forEach(t => t.classList.toggle("on", t.dataset.tab === tab && (tab !== "buy" || t.dataset.f === BUY_F)));
  ["orders", "buy", "eip"].forEach(k => $(`#pane-${k}`).classList.toggle("hidden", k !== tab));
  try { localStorage.setItem("shopee_tab", JSON.stringify([tab, BUY_F])); } catch (e) { /* 無痕視窗存不了就算了 */ }
  if (tab === "orders") { loadOrders(); loadUploads(); }
  else if (tab === "buy") { BSEL.clear(); BOPEN.clear(); $("#b-err").innerHTML = ""; loadConv(); loadPlan(); }
  else loadEips();
}
$$("#tabs .tab").forEach(t => t.addEventListener("click", () => showTab(t.dataset.tab, t.dataset.f)));

/* ── 價格本 ───────────────────────────────────────────────── */
async function loadConv() {
  let d; try { d = await api("/api/shopee/conv"); } catch (err) { $("#conv-st").innerHTML = `<span class="neg">${esc(err.message)}</span>`; return; }
  $("#conv-st").innerHTML = ["寶僑", "紙潔"].map(l => {
    const x = d.lines[l];
    return x ? `<div class="it"><b>${l}</b> <span class="kbd">${fmt(x.count)} 筆${x.blank ? `，<span class="neg">箱入數空白 ${fmt(x.blank)} 筆</span>` : ""}</span>
      <div class="kbd" title="${esc(x.file || "")}">${esc((x.updated_at || "").slice(0, 16).replace("T", " "))}　${esc(x.by || "")}</div></div>`
      : `<div class="it"><b>${l}</b> <span class="badge b-warn">尚未上傳</span></div>`;
  }).join("");
}
const dzc = $("#dz-conv");
dzc.addEventListener("click", e => { if (e.target.id !== "conv-file") $("#conv-file").click(); });
["dragenter", "dragover"].forEach(ev => dzc.addEventListener(ev, e => { e.preventDefault(); dzc.classList.add("over"); }));
["dragleave", "drop"].forEach(ev => dzc.addEventListener(ev, e => { e.preventDefault(); dzc.classList.remove("over"); }));
dzc.addEventListener("drop", e => { const f = [...e.dataTransfer.files].find(x => /\.xlsx$/i.test(x.name)); f ? uploadConv(f) : toast("請拖曳價格本 Excel 檔（.xlsx）", "err"); });
$("#conv-file").addEventListener("change", () => { const f = $("#conv-file").files[0]; $("#conv-file").value = ""; if (f) uploadConv(f); });
async function uploadConv(f) {
  const fd = new FormData(); fd.append("file", f);
  try {
    const r = await api("/api/shopee/conv/import", { method: "POST", body: fd });
    toast(Object.entries(r.lines).map(([l, x]) => `${l}價格本已更新：${fmt(x.count)} 筆（原 ${fmt(x.before)} 筆）${x.blank ? `，其中 ${fmt(x.blank)} 筆箱入數空白` : ""}`).join("\n")
      + (r.warnings.length ? `\n${r.warnings.slice(0, 3).join("\n")}` : ""));
    loadConv(); loadPlan();
  } catch (err) { toast(err.message, "err"); }
}

/* ── 勾選訂單 → 產出 EIP 採購單 ───────────────────────────── */
let PLAN = { orders: [] }; const BOPEN = new Set(), BSEL = new Set();
const HINT = { "竹運採購進貨": "向供應商採購、進竹運倉。寶僑分「箱單位」「小單位」兩個檔，單位與數量照蝦皮；備註欄放轉換率。",
               "供應商直送": "供應商直接送到蝦皮倉，一天一倉一個檔。寶僑小單位換成箱數；紙潔一律換成箱數（需求 10.2、10.3）。" };
$("#b-from").value = iso(new Date(Date.now() - 7 * 864e5));
["#b-from", "#b-to", "#b-line"].forEach(s => $(s).addEventListener("change", () => { BSEL.clear(); loadPlan(); }));
$("#b-todo").addEventListener("click", () => { $("#b-todo").classList.toggle("on"); renderPlan(); });
const canPick = g => g.supported && !g.blocking;

async function loadPlan() {
  $("#b-title").textContent = BUY_F === "供應商直送" ? "直送訂單" : "竹運採購進貨";
  $("#b-hint").textContent = HINT[BUY_F];
  const p = new URLSearchParams({ fulfil: BUY_F, from: $("#b-from").value, to: $("#b-to").value, line: $("#b-line").value });
  try { PLAN = await api(`/api/shopee/purchase/plan?${p}`); } catch (err) { $("#b-tbl").innerHTML = `<tr><td class="muted" style="padding:20px">${esc(err.message)}</td></tr>`; return; }
  const ok = new Set(PLAN.orders.filter(canPick).map(g => g.order_no));
  [...BSEL].forEach(k => { if (!ok.has(k)) BSEL.delete(k); });
  renderPlan();
}

function renderPlan() {
  const hide = $("#b-todo").classList.contains("on");
  const list = PLAN.orders.filter(g => !(hide && g.po_status === "已回填"));
  const picks = list.filter(canPick), blk = list.filter(g => g.supported && g.blocking).length;
  $("#b-sum").innerHTML = `<span><b>${list.length}</b> 張訂單</span><span><b>${picks.length}</b> 張可產出</span>${blk ? `<span><b style="color:var(--bad)">${blk}</b> 張需先處理</span>` : ""}`
    + `<span><b>${list.filter(g => g.po_status === "已產出").length}</b> 張已產出待回填</span>`;
  const all = picks.length && picks.every(g => BSEL.has(g.order_no));
  let h = `<thead><tr><th><input type="checkbox" class="ck" id="b-all" ${all ? "checked" : ""} ${picks.length ? "" : "disabled"} title="全選可產出的訂單"></th>
    <th>訂單編號</th><th>線別</th><th>到貨日</th><th>倉別</th><th class="num">品項</th><th class="num">蝦皮數量</th><th>檢查結果</th><th>採購單產出狀態</th></tr></thead><tbody>`;
  for (const g of list) {
    const k = g.order_no, pick = canPick(g);
    const st = !g.supported ? `<span class="badge b-dim">${esc(g.line === "瑪氏" ? "瑪氏直送之後開放" : "認不出線別")}</span>`
      : g.blocking ? `<span class="badge b-diff">${g.blocking} 項需處理</span>` : `<span class="badge b-ok">可產出</span>`;
    h += `<tr class="og ${g.supported && g.blocking ? "blk" : ""} ${g.supported ? "" : "lock"}" data-k="${esc(k)}">
      <td><input type="checkbox" class="ck b-ck" data-k="${esc(k)}" ${BSEL.has(k) ? "checked" : ""} ${pick ? "" : "disabled"}></td>
      <td><span class="fn" style="font-weight:700;color:var(--ink)">${esc(k)}</span>${g.po_id && g.pr_id ? `<div class="kbd">${esc(g.pr_id)}</div>` : ""}</td>
      <td class="nw">${esc(g.line)}</td><td class="nw">${esc(g.expected_date)}</td><td class="nw">${wh(g.warehouse)}</td>
      <td class="num">${g.items.length}</td><td class="num">${fmt(g.qty)}</td><td>${st}</td>
      <td><span class="badge ${g.po_status === "未產出" ? "b-dim" : g.po_status === "已回填" ? "b-ok" : "b-s"}">${esc(g.po_status)}</span>${eipOf(g)}</td></tr>`;
    if (BOPEN.has(k)) h += planItems(g);
  }
  if (!list.length) h += `<tr><td colspan="9" class="muted" style="padding:24px;text-align:center">此條件下沒有${esc(BUY_F)}的訂單。請調整到貨日，或在「訂單匯入與總覽」匯入時選擇此履約方式。</td></tr>`;
  $("#b-tbl").innerHTML = h + "</tbody>";
  $$("#b-tbl tr.og").forEach(tr => tr.addEventListener("click", e => { if (e.target.closest("button, a, input")) return; const k = tr.dataset.k; BOPEN.has(k) ? BOPEN.delete(k) : BOPEN.add(k); renderPlan(); }));
  $$("#b-tbl input.b-ck").forEach(c => c.addEventListener("change", () => { c.checked ? BSEL.add(c.dataset.k) : BSEL.delete(c.dataset.k); renderPlan(); }));
  const ba = $("#b-all"); if (ba) ba.addEventListener("change", () => { picks.forEach(g => ba.checked ? BSEL.add(g.order_no) : BSEL.delete(g.order_no)); renderPlan(); });
  $$("#b-tbl button.adj").forEach(b => b.addEventListener("click", () => openAdj(Number(b.dataset.id))));
  const n = [...BSEL].length;
  $("#b-go").disabled = !n;
  $("#b-go-hint").textContent = n ? `已勾選 ${n} 張訂單，會產出 EIP 採購單＋採購下採明細（zip）。` : "請勾選要產出的訂單；有「需處理」的訂單請先點開修正。";
}
function eipOf(g) {
  const ps = [...new Set(g.items.map(p => p.eip && p.eip.eip_po).filter(Boolean))];
  return ps.length ? `<div class="kbd fn">${ps.map(esc).join("、")}</div>` : "";
}
function planItems(g) {
  const bao = g.line === "寶僑";
  return `<tr class="sub"><td colspan="9"><table class="t"><thead><tr><th>廠商料號</th><th>品名</th><th>單位</th><th class="num">蝦皮數量</th><th class="num">轉換率</th>
    <th class="num">下採數量</th><th>EIP 單位</th><th class="num">EIP 數量</th>${bao ? "<th>檔案</th>" : ""}<th>狀態</th><th></th></tr></thead><tbody>
    ${g.items.map(p => `<tr><td class="fn">${esc(p.supplier_sku_id)}</td><td>${esc((p.name || "").slice(0, 46))}</td><td class="nw">${esc(p.unit)}</td>
      <td class="num">${fmt(p.qty)}</td><td class="num">${fmt(p.conv)}</td>
      <td class="num">${p.adj ? `<b>${fmt(p.use_qty)}</b><div class="kbd">已修改</div>` : fmt(p.use_qty)}</td>
      <td class="nw">${esc(p.out_unit) || "—"}</td><td class="num"><b>${p.out_qty == null ? "—" : fmt(p.out_qty)}</b></td>${bao ? `<td class="nw kbd">${esc(p.unit_group) || "—"}</td>` : ""}
      <td>${p.problems.length ? p.problems.map(x => `<div class="prob"><i class="bi bi-exclamation-circle"></i> ${esc(x)}</div>`).join("") : `<span class="badge b-ok">正常</span>`}</td>
      <td class="nw">${p.rem || p.adj ? `<button class="btn ${p.rem ? "btn-p" : "btn-o"} btn-sm adj" data-id="${p.order_id}"><i class="bi bi-pencil"></i> 修改數量</button>` : ""}</td></tr>`).join("")}
    </tbody></table></td></tr>`;
}

$("#b-go").addEventListener("click", async () => {
  const nos = [...BSEL]; if (!nos.length) return;
  const b = $("#b-go"); b.disabled = true; $("#b-err").innerHTML = "";
  try {
    const res = await fetch("/api/shopee/purchase/generate", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ fulfil: BUY_F, order_nos: nos }) });
    if (!res.ok) {
      let d = {}; try { d = await res.json(); } catch (e) {}
      $("#b-err").innerHTML = `<div class="alert al-bad"><b>${esc(d.error || `伺服器錯誤（${res.status}）`)}</b>${(d.details || []).length ? `<div class="mt-1">${d.details.map(esc).join("<br>")}</div>` : ""}</div>`;
      return;
    }
    const blob = await res.blob(); const m = (res.headers.get("Content-Disposition") || "").match(/filename\*=UTF-8''([^;]+)/);
    const a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = m ? decodeURIComponent(m[1]) : "蝦皮特選採購單.zip"; document.body.appendChild(a); a.click(); a.remove();
    toast(`已產出 ${res.headers.get("X-Shopee-Eip-Files")} 個 EIP 採購單檔（共 ${res.headers.get("X-Shopee-Files")} 個檔案）。上傳 EIP 後，請至「採購單產出」回填單號。`);
    BSEL.clear(); loadPlan();
  } catch (err) { toast(err.message, "err"); }
  finally { b.disabled = ![...BSEL].length; }
});

/* ── 修改下採數量（需求 10.5：除不盡不自動進位或捨去，人改＋寫原因） ── */
let ADJ = null;
function openAdj(id) {
  let p = null, g = null;
  for (const x of PLAN.orders) { const y = x.items.find(i => i.order_id === id); if (y) { p = y; g = x; break; } }
  if (!p) return;
  ADJ = p;
  const c = p.conv, q = p.qty, down = Math.floor(q / c) * c, up = Math.ceil(q / c) * c;
  $("#da-title").textContent = `${g.order_no} 修改下採數量`;
  $("#da-why").textContent = q % c ? `蝦皮數量 ${fmt(q)} 除以轉換率 ${c} 除不盡，系統不會自動進位或捨去，請填寫修改後數量與原因。` : `蝦皮數量 ${fmt(q)} 可以整除轉換率 ${c}；如不需修改，可按「取消修改」回到原始數量。`;
  $("#da-facts").innerHTML = `<span>商品料號 <b class="fn">${esc(p.supplier_sku_id)}</b></span><span>${esc((p.name || "").slice(0, 40))}</span>
    <span>原始數量 <b>${fmt(q)}</b> ${esc(p.unit)}</span><span>轉換率 <b>${c}</b></span><span>計算結果 <b>${(q / c).toFixed(2).replace(/\.?0+$/, "")}</b> 箱</span><span>餘數 <b class="${q % c ? "neg" : ""}">${q % c}</b></span>`;
  $("#da-down").textContent = `${fmt(down)}（${down / c} 箱）`; $("#da-down").dataset.v = down; $("#da-down").classList.toggle("hidden", !down || down === q);
  $("#da-up").textContent = `${fmt(up)}（${up / c} 箱）`; $("#da-up").dataset.v = up; $("#da-up").classList.toggle("hidden", up === q);
  $("#da-qty").value = p.adj ? p.adj.qty_after : ""; $("#da-reason").value = p.adj ? p.adj.reason : "";
  $("#da-who").textContent = p.adj ? `上次修改：${p.adj.operator}　${(p.adj.updated_at || "").slice(0, 16).replace("T", " ")}` : "";
  $("#da-clear").classList.toggle("hidden", !p.adj);
  $("#dlg-adj").showModal();
}
["#da-down", "#da-up"].forEach(s => $(s).addEventListener("click", () => { $("#da-qty").value = $(s).dataset.v; $("#da-reason").focus(); }));
$("#da-save").addEventListener("click", async () => {
  if (!$("#da-qty").value.trim()) return toast("請填寫修改後數量", "err");
  try { const r = await post("/api/shopee/purchase/adjust", { order_id: ADJ.order_id, qty: $("#da-qty").value.trim(), reason: $("#da-reason").value.trim() });
    toast(`${ADJ.supplier_sku_id} 下採數量已改為 ${fmt(r.qty)}`); $("#dlg-adj").close(); loadPlan();
  } catch (err) { toast(err.message, "err"); }
});
$("#da-clear").addEventListener("click", async () => {
  try { await post("/api/shopee/purchase/adjust", { order_id: ADJ.order_id, qty: "" }); toast(`${ADJ.supplier_sku_id} 已回到原始數量 ${fmt(ADJ.qty)}`); $("#dlg-adj").close(); loadPlan(); }
  catch (err) { toast(err.message, "err"); }
});

/* ── 採購單產出：EIP 檔清單、回填 EIP 單號 ─────────────────── */
let EIPS = [];
$("#e-from").value = iso(new Date(Date.now() - 30 * 864e5));
["#e-from", "#e-to"].forEach(s => $(s).addEventListener("change", loadEips));
$("#e-todo").addEventListener("click", () => { $("#e-todo").classList.toggle("on"); renderEips(); });
async function loadEips() {
  const p = new URLSearchParams({ from: $("#e-from").value, to: $("#e-to").value });
  try { EIPS = (await api(`/api/shopee/eips?${p}`)).eips; } catch (err) { $("#e-tbl").innerHTML = `<tr><td class="muted" style="padding:20px">${esc(err.message)}</td></tr>`; return; }
  renderEips();
}
function renderEips() {
  const list = EIPS.filter(e => !($("#e-todo").classList.contains("on") && e.eip_po));
  $("#e-tbl").innerHTML = `<thead><tr><th>到貨日</th><th>倉別</th><th>線別</th><th>履約方式</th><th>檔名（點選重新下載）</th><th class="num">品項</th><th class="num">數量</th><th>蝦皮訂單</th><th>產出</th><th>EIP 採購單號</th></tr></thead><tbody>
    ${list.map(e => `<tr><td class="nw">${esc(e.expected_date)}</td><td class="nw">${esc(e.warehouse_name)}</td><td class="nw">${esc(e.line)}</td><td class="nw">${esc(e.fulfil)}</td>
      <td><a class="lnk" href="/api/shopee/eips/${e.id}/file">${esc(e.filename)}</a></td><td class="num">${e.rows}</td><td class="num">${fmt(e.qty)}</td>
      <td class="fn" title="${esc(e.orders.join("\n"))}">${esc(e.orders.slice(0, 2).join("、"))}${e.orders.length > 2 ? `…等 ${e.orders.length} 張` : ""}</td>
      <td class="kbd nw">${esc(e.created_by)}<div>${esc((e.created_at || "").slice(5, 16).replace("T", " "))}</div></td>
      <td class="nw"><input class="inp fn e-po" data-id="${e.id}" value="${esc(e.eip_po)}" placeholder="例：PO202610001" style="width:140px">
        <button class="btn btn-o btn-sm e-save" data-id="${e.id}">儲存</button>${e.eip_po ? `<div class="kbd">${esc(e.updated_by)} ${esc((e.updated_at || "").slice(5, 16).replace("T", " "))}</div>` : ""}</td></tr>`).join("")
      || `<tr><td colspan="10" class="muted" style="padding:24px;text-align:center">此條件下沒有 EIP 採購單。請至「竹運採購進貨」或「直送訂單」勾選訂單產出。</td></tr>`}</tbody>`;
  $$("#e-tbl button.e-save").forEach(b => b.addEventListener("click", () => saveEip(b.dataset.id)));
  $$("#e-tbl input.e-po").forEach(i => i.addEventListener("keydown", e => { if (e.key === "Enter") saveEip(i.dataset.id); }));
}
async function saveEip(id) {
  const v = $(`#e-tbl input.e-po[data-id="${id}"]`).value.trim();
  try { const r = await api(`/api/shopee/eips/${id}`, { method: "PUT", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ eip_po: v }) });
    toast(r.eip_po ? `已回填 EIP 採購單號 ${r.eip_po}` : "已清除 EIP 採購單號，該批訂單可重新產出"); loadEips();
  } catch (err) { toast(err.message, "err"); }
}

let START = ["orders", "竹運採購進貨"];
try { const v = JSON.parse(localStorage.getItem("shopee_tab") || "null"); if (Array.isArray(v) && ["orders", "buy", "eip"].includes(v[0])) START = v; } catch (e) { /* 讀不到就從總覽開始 */ }
showTab(START[0], START[1] === "供應商直送" ? "供應商直送" : "竹運採購進貨");

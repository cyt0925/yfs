/* 蝦皮特選訂單管理系統（/shopee）
   2026-10-13 改版（Jerry：「真的要分為六個介面嗎？」）：六個頁籤合成一個畫面。
   - 最上面一個入口收所有檔案（/api/shopee/drop 看內容認：蝦皮採購單／入庫單、價格本、庫存資料；PDF 是勇信配送明細表）
   - 一排作業流程格子（差異待確認 → 待產採購單 → 待回填 EIP 採購單號 → 待產拋單檔 → 待通知蝦皮 → 全部訂單），
     數字來自訂單總覽算的每張單待辦（todo）；點一格，下面的清單換成那一步、只出現那一步要按的按鈕
   - 到貨日、線別、搜尋一組篩選，各步共用
   - 價格本、庫存資料、效期標準、倉資料、收件資料、上傳紀錄、通知紀錄收進「資料與設定」 */
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

/* ── 上傳：一個入口收所有檔案 ─────────────────────────────── */
const dz = $("#dz");
dz.addEventListener("click", e => { if (!e.target.closest(".fresh") && e.target.id !== "file") $("#file").click(); });
["dragenter", "dragover"].forEach(ev => dz.addEventListener(ev, e => { e.preventDefault(); dz.classList.add("over"); }));
["dragleave", "drop"].forEach(ev => dz.addEventListener(ev, e => { e.preventDefault(); dz.classList.remove("over"); }));
dz.addEventListener("drop", e => {
  const fs = [...e.dataTransfer.files].filter(f => /\.(xlsx|xls|xlsm|pdf)$/i.test(f.name));
  if (fs.length) handleFiles(fs); else toast("請拖曳 Excel 檔或勇信配送明細表（PDF）", "err");
});
$("#file").addEventListener("change", () => { const fs = [...$("#file").files]; $("#file").value = ""; if (fs.length) handleFiles(fs); });
$("#fresh").addEventListener("click", e => { e.preventDefault(); e.stopPropagation(); showSettings(true); });

async function handleFiles(files) {
  const pdfs = files.filter(f => /\.pdf$/i.test(f.name)), others = files.filter(f => !/\.pdf$/i.test(f.name));
  let changed = false;
  for (const f of others) changed = (await dropOne(f)) || changed;
  if (pdfs.length) await yxUpload(pdfs);
  if (changed) { loadFresh(); refresh(); }
}
async function dropOne(f) {
  const fd = new FormData(); fd.append("file", f);
  const box = document.createElement("div"); box.className = "pv"; box.innerHTML = `<div class="kbd"><i class="bi bi-hourglass-split"></i> 讀取 ${esc(f.name)} 中…</div>`;
  $("#previews").prepend(box);
  let res, d = {};
  try { res = await fetch("/api/shopee/drop", { method: "POST", body: fd }); try { d = await res.json(); } catch (e) {} }
  catch (err) { box.innerHTML = `<div class="alert al-bad"><b>${esc(f.name)}</b>：${esc(err.message)}</div>`; return false; }
  if (!res.ok) {
    box.innerHTML = `<div class="alert al-bad"><b>${esc(f.name)}</b>：${esc(d.error || `伺服器錯誤（${res.status}）`)}${(d.details || []).length ? `<div class="mt-1">${d.details.slice(0, 8).map(esc).join("<br>")}</div>` : ""}</div>`;
    setTimeout(() => box.remove(), 15000); return false;
  }
  const kind = res.headers.get("X-Shopee-Kind");
  if (kind === "orders") { renderPreview(box, d); return false; }
  box.remove();
  if (kind === "conv") toast(Object.entries(d.lines).map(([l, x]) => `${l}價格本已更新：${fmt(x.count)} 筆（原 ${fmt(x.before)} 筆）${x.blank ? `，其中 ${fmt(x.blank)} 筆箱入數空白` : ""}`).join("\n")
    + (d.warnings.length ? `\n${d.warnings.slice(0, 3).join("\n")}` : ""));
  else if (kind === "stock") toast(`已更新${d.label}：${f.name}（${fmt(d.rows)} 列）` + (d.notes.length ? `\n${d.notes.join("\n")}` : ""));
  else if (kind === "yx") { showYx(d); return false; }
  return true;
}

/* 上傳區右邊：價格本、庫存資料什麼時候更新的（點了到「資料與設定」） */
async function loadFresh() {
  let c, k; try { [c, k] = await Promise.all([api("/api/shopee/conv"), api("/api/shopee/stock")]); } catch (err) { return; }
  const md = t => t ? `${Number(t.slice(5, 7))}/${Number(t.slice(8, 10))}` : "";
  const chips = ["寶僑", "紙潔"].map(l => { const x = c.lines[l]; return x ? `<span title="${esc(x.file || "")}"><i class="bi bi-check-circle" style="color:var(--ok)"></i> ${l}價格本 ${md(x.updated_at)}</span>`
    : `<span class="w"><i class="bi bi-exclamation-circle"></i> ${l}價格本未上傳</span>`; });
  for (const kk of KIND_ORDER) {
    const x = k.stock[kk];
    if (!x) { chips.push(`<span class="w"><i class="bi bi-exclamation-circle"></i> ${KIND_LABEL[kk]}未上傳</span>`); continue; }
    const old = kk !== "life" && !x.today;
    chips.push(`<span class="${old ? "w" : ""}" title="${esc(x.filename)}"><i class="bi ${old ? "bi-exclamation-circle" : "bi-check-circle"}" ${old ? "" : `style="color:var(--ok)"`}></i> ${KIND_LABEL[kk]} ${x.today ? "今天" : md(x.uploaded_at)}</span>`);
  }
  $("#fresh").innerHTML = chips.join("");
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
      refresh();
    } catch (err) { e.target.disabled = false; toast(err.message, "err"); }
  });
}

/* ── 共用篩選、作業流程格子 ─────────────────────────────── */
let VIEW = { orders: [], stages: [] }, ORD = new Map(), STAGE = "", EIP_TODO = 0; const OPEN = new Set();
const iso = d => d.toISOString().slice(0, 10);
$("#f-from").value = iso(new Date(Date.now() - 7 * 864e5));
["#f-from", "#f-to", "#f-line"].forEach(s => $(s).addEventListener("change", () => { BSEL.clear(); ZSEL.clear(); KSEL.clear(); refresh(); }));
$("#f-q").addEventListener("keydown", e => { if (e.key === "Enter") refresh(); });
const fq = () => ({ from: $("#f-from").value, to: $("#f-to").value, line: $("#f-line").value });
const inStage = no => STAGE === "all" || ((ORD.get(no) || {}).todo || []).includes(STAGE);
const STAGE_DEF = [
  { key: "diff", t: "差異待確認", d: "蝦皮重傳後內容有變", act: "確認差異", bad: true,
    hint: "蝦皮重傳後數量、到貨日或倉別有變；確認差異後，訂單才能往下處理。" },
  { key: "buy", t: "待產採購單", d: "竹運採購進貨、直送", act: "產出 EIP 採購單",
    hint: "勾選訂單產出 EIP 採購單；「需處理」的訂單請點該列，在右側修正。" },
  { key: "eip", t: "待回填 EIP 採購單號", d: "上傳 EIP 後回填", act: "回填 EIP 採購單號",
    hint: "上傳 EIP 後回填採購單號，回填後即鎖定；瑪氏回填後可下載瑪氏採購單。" },
  { key: "ship", t: "待產拋單檔", d: "竹運出貨", act: "產出拋單檔",
    hint: "系統依 EC 庫存、竹運效期判斷可出貨量；勾選訂單產出拋單檔，缺貨會自動列入「待通知蝦皮」。" },
  { key: "notify", t: "待通知蝦皮", d: "有缺貨、尚未通知", act: "通知蝦皮",
    hint: "勾選訂單產生 LINE 通知文字。瑪氏缺貨請上傳勇信配送明細表；寶僑、紙潔缺貨請點該單，在右側填寫。" },
  { key: "all", t: "全部訂單", d: "", hint: "點一張單，右側看品項、採購單、缺貨、差異與歷程；採購來信缺貨時，搜尋訂單後在右側填寫。" },
];
const SDEF = Object.fromEntries(STAGE_DEF.map(x => [x.key, x]));

async function loadOrders() {
  const p = new URLSearchParams({ ...fq(), q: $("#f-q").value.trim() });
  try { VIEW = await api(`/api/shopee/orders?${p}`); } catch (err) { toast(err.message, "err"); return; }
  ORD = new Map(VIEW.orders.map(g => [g.order_no, g]));
}
async function loadEipCount() {
  try { EIPS = (await api(`/api/shopee/eips?${new URLSearchParams(fq())}`)).eips; } catch (err) { EIPS = []; }
  EIP_TODO = eipList(false).length;
}
function renderStages() {
  const cnt = Object.fromEntries((VIEW.stages || []).map(x => [x.key, x.count]));
  cnt.eip = EIP_TODO; cnt.all = VIEW.orders.length;
  $("#stages").innerHTML = STAGE_DEF.map((x, i) => `${i ? `<span class="arr"><i class="bi bi-chevron-right"></i></span>` : ""}
    <div class="st ${STAGE === x.key ? "on" : ""} ${x.key === "all" ? "all" : !cnt[x.key] ? "zero" : x.bad ? "bad" : ""}" data-s="${x.key}" role="button" tabindex="0">
      <div class="n">${fmt(cnt[x.key] || 0)}</div><div class="t">${x.t}</div><div class="d">${x.key === "all" ? `其中 ${fmt(cnt.done || 0)} 張已處理` : x.d}</div></div>`).join("");
  $$("#stages .st").forEach(el => {
    el.addEventListener("click", () => setStage(el.dataset.s));
    el.addEventListener("keydown", e => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); setStage(el.dataset.s); } });
  });
}
function setStage(k) {
  if (STAGE !== k) { BSEL.clear(); ZSEL.clear(); KSEL.clear(); OPEN.clear(); BOPEN.clear(); ZOPEN.clear(); KOPEN.clear(); }
  STAGE = k;
  try { localStorage.setItem("shopee_stage", k); } catch (e) { /* 無痕視窗存不了就算了 */ }
  renderStages(); loadStage();
}
async function loadStage() {
  const d = SDEF[STAGE];
  $("#st-title").textContent = d.t; $("#st-hint").textContent = d.hint;
  const pane = { diff: "orders", all: "orders", buy: "buy", eip: "eip", ship: "ship", notify: "notify" }[STAGE];
  ["orders", "buy", "eip", "ship", "notify"].forEach(k => $(`#st-${k}`).classList.toggle("hidden", k !== pane));
  if (pane === "orders") render();
  else if (pane === "buy") { $("#b-err").innerHTML = ""; loadPlan(); }
  else if (pane === "eip") renderEips();
  else if (pane === "ship") { $("#z-err").innerHTML = ""; loadShip(); }
  else loadShort();
}
async function refresh() {
  await Promise.all([loadOrders(), loadEipCount()]);
  if (!STAGE) {        // 一進來：網址指定的那一步（首頁今日事項連 /shopee#diff）→ 上次看的那一步 → 第一個有待辦的
    let saved = ""; try { saved = localStorage.getItem("shopee_stage") || ""; } catch (e) { /* 讀不到就算了 */ }
    if (SDEF[location.hash.slice(1)]) saved = location.hash.slice(1);
    const cnt = Object.fromEntries((VIEW.stages || []).map(x => [x.key, x.count])); cnt.eip = EIP_TODO;
    STAGE = SDEF[saved] ? saved : (["diff", "buy", "eip", "ship", "notify"].find(k => cnt[k]) || "all");
  }
  renderStages(); loadStage();
  if (SIDE_NO) openSide(SIDE_NO);       // 做完事，右側面板也換成最新的
}

/* ── 差異待確認／全部訂單：一張單一列，點開看品項與操作 ─────── */
const NEXT_CLS = { diff: "b-diff", buy: "b-s", eip: "b-s", ship: "b-s", notify: "b-warn" };
function nextBadges(g) {
  if (!g.todo.length) return `<span class="badge b-ok">已處理</span>`;
  return `<span class="nx">${g.todo.map(k => `<span class="badge ${NEXT_CLS[k]} go-st" data-s="${k}" title="前往「${SDEF[k].t}」" style="cursor:pointer">${SDEF[k].act}</span>`).join("")}</span>`;
}
function render() {
  const list = VIEW.orders.filter(g => inStage(g.order_no));
  $("#sum").innerHTML = `<span><b>${list.length}</b> 張訂單</span><span><b>${fmt(list.reduce((a, g) => a + g.item_count, 0))}</b> 個品項</span><span><b>${fmt(list.reduce((a, g) => a + g.qty, 0))}</b> 件（蝦皮數量）</span>`;
  let h = `<thead><tr><th>訂單編號</th><th>線別</th><th>到貨日</th><th>倉別</th><th class="opt">履約方式</th><th class="num opt">品項</th><th class="num">蝦皮數量</th><th class="num opt">可出貨量</th><th class="num opt">缺貨量</th><th>下一步</th></tr></thead><tbody>`;
  for (const g of list) {
    const k = g.order_no;
    h += `<tr class="og ${g.diff_pending ? "pend" : ""}" data-k="${esc(k)}">
      <td><span class="fn" style="font-weight:700;color:var(--ink)">${esc(k)}</span>${g.po_id && g.pr_id ? `<div class="kbd">${esc(g.pr_id)}</div>` : !g.po_id ? `<div class="kbd">尚未成 PO</div>` : ""}</td>
      <td class="nw">${esc(g.line)}</td><td class="nw">${esc(g.expected_date)}</td><td class="nw">${wh(g.warehouse)}</td>
      <td class="nw opt">${esc(g.fulfil)}</td>
      <td class="num opt">${g.item_count}${g.removed_count ? `<div class="kbd neg">移除 ${g.removed_count}</div>` : ""}</td>
      <td class="num"><b>${fmt(g.qty)}</b>${g.qty_original !== g.qty ? `<div class="kbd">原 ${fmt(g.qty_original)}</div>` : ""}</td>
      <td class="num opt ${g.qty_ship == null ? "muted" : ""}">${fmt(g.qty_ship)}</td><td class="num opt ${g.qty_short ? "neg" : g.qty_short == null ? "muted" : ""}">${fmt(g.qty_short)}</td>
      <td>${STAGE === "diff" ? `<button class="btn btn-p btn-sm dd" data-k="${esc(k)}"><i class="bi bi-arrow-left-right"></i> 確認差異</button>` : nextBadges(g)}</td></tr>`;
  }
  if (!list.length) h += `<tr><td colspan="10" class="muted" style="padding:24px;text-align:center">${STAGE === "all" ? "此條件下沒有訂單。請調整到貨日，或將蝦皮採購單拖曳至上方上傳。" : "這一步目前沒有要處理的訂單。"}</td></tr>`;
  $("#tbl").innerHTML = h + "</tbody>";
  $$("#tbl tr.og").forEach(tr => tr.addEventListener("click", e => { if (e.target.closest("button, a, input, .go-st")) return; openSide(tr.dataset.k); }));
  $$("#tbl button.dd").forEach(b => b.addEventListener("click", () => openDiff(b.dataset.k)));
  $$("#tbl .go-st").forEach(b => b.addEventListener("click", () => setStage(b.dataset.s)));
  markSel();
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
  try { await post("/api/shopee/diffs/confirm", { order_nos: [DIFF_NO] }); toast(`已確認 ${DIFF_NO} 的差異`); $("#dlg-diff").close(); refresh(); }
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
    toast(r.changed ? `${FF_NO} 已改為${FF_SEL}` : "履約方式相同，未修改"); $("#dlg-fulfil").close(); refresh();
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


/* ── 價格本 ───────────────────────────────────────────────── */
async function loadConv() {
  let d; try { d = await api("/api/shopee/conv"); } catch (err) { $("#conv-st").innerHTML = `<span class="neg">${esc(err.message)}</span>`; return; }
  $("#conv-st").innerHTML = ["寶僑", "紙潔", "瑪氏"].map(l => {
    const x = d.lines[l], src = l === "瑪氏" ? `<div class="kbd">商品總表請至<a class="lnk" href="/mars">瑪氏出貨</a>上傳</div>` : "";
    return x ? `<div class="it"><b>${l}</b> <span class="kbd">${fmt(x.count)} 筆${x.blank ? `，<span class="neg">箱入數空白 ${fmt(x.blank)} 筆</span>` : ""}</span>
      <div class="kbd" title="${esc(x.file || "")}">${esc((x.updated_at || "").slice(0, 16).replace("T", " "))}　${esc(x.by || "")}</div>${src}</div>`
      : `<div class="it"><b>${l}</b> <span class="badge b-warn">尚未上傳</span>${src}</div>`;
  }).join("");
}

/* ── 勾選訂單 → 產出 EIP 採購單 ───────────────────────────── */
let PLAN = { orders: [] }; const BOPEN = new Set(), BSEL = new Set();
const BUY_FULFILS = ["竹運採購進貨", "供應商直送"];
const canPick = g => g.supported && !g.blocking;

async function loadPlan() {
  try {
    const parts = await Promise.all(BUY_FULFILS.map(f => api(`/api/shopee/purchase/plan?${new URLSearchParams({ ...fq(), fulfil: f })}`)
      .then(d => d.orders.map(g => ({ ...g, fulfil: f })))));
    PLAN = { orders: parts.flat().sort((a, b) => (a.expected_date || "").localeCompare(b.expected_date || "") || a.order_no.localeCompare(b.order_no)) };
  } catch (err) { $("#b-tbl").innerHTML = `<tr><td class="muted" style="padding:20px">${esc(err.message)}</td></tr>`; return; }
  const ok = new Set(PLAN.orders.filter(g => canPick(g) && inStage(g.order_no)).map(g => g.order_no));
  [...BSEL].forEach(k => { if (!ok.has(k)) BSEL.delete(k); });
  renderPlan();
}

function renderPlan() {
  const list = PLAN.orders.filter(g => inStage(g.order_no));
  const picks = list.filter(canPick), blk = list.filter(g => g.supported && g.blocking).length;
  $("#b-sum").innerHTML = `<span><b>${list.length}</b> 張訂單</span><span><b>${picks.length}</b> 張可產出</span>${blk ? `<span><b style="color:var(--bad)">${blk}</b> 張需先處理</span>` : ""}`;
  const all = picks.length && picks.every(g => BSEL.has(g.order_no));
  let h = `<thead><tr><th><input type="checkbox" class="ck" id="b-all" ${all ? "checked" : ""} ${picks.length ? "" : "disabled"} title="全選可產出的訂單"></th>
    <th>訂單編號</th><th>線別</th><th>到貨日</th><th>倉別</th><th>履約方式</th><th class="num opt">品項</th><th class="num">蝦皮數量</th><th>檢查結果</th></tr></thead><tbody>`;
  for (const g of list) {
    const k = g.order_no, pick = canPick(g);
    const st = !g.supported ? `<span class="badge b-dim">${esc(g.line === "瑪氏" ? "瑪氏目前只開放供應商直送" : "認不出線別")}</span>`
      : g.blocking ? `<span class="badge b-diff">${g.blocking} 項需處理</span>` : `<span class="badge b-ok">可產出</span>`;
    h += `<tr class="og ${g.supported && g.blocking ? "blk" : ""} ${g.supported ? "" : "lock"}" data-k="${esc(k)}">
      <td><input type="checkbox" class="ck b-ck" data-k="${esc(k)}" ${BSEL.has(k) ? "checked" : ""} ${pick ? "" : "disabled"}></td>
      <td><span class="fn" style="font-weight:700;color:var(--ink)">${esc(k)}</span>${g.po_id && g.pr_id ? `<div class="kbd">${esc(g.pr_id)}</div>` : ""}</td>
      <td class="nw">${esc(g.line)}</td><td class="nw">${esc(g.expected_date)}</td><td class="nw">${wh(g.warehouse)}</td><td class="nw">${esc(g.fulfil === "供應商直送" ? "直送" : g.fulfil)}</td>
      <td class="num opt">${g.items.length}</td><td class="num">${fmt(g.qty)}</td><td>${st}</td></tr>`;
  }
  if (!list.length) h += `<tr><td colspan="9" class="muted" style="padding:24px;text-align:center">這一步目前沒有要處理的訂單。</td></tr>`;
  $("#b-tbl").innerHTML = h + "</tbody>";
  $$("#b-tbl tr.og").forEach(tr => tr.addEventListener("click", e => { if (e.target.closest("button, a, input")) return; openSide(tr.dataset.k); }));
  $$("#b-tbl input.b-ck").forEach(c => c.addEventListener("change", () => { c.checked ? BSEL.add(c.dataset.k) : BSEL.delete(c.dataset.k); renderPlan(); }));
  const ba = $("#b-all"); if (ba) ba.addEventListener("change", () => { picks.forEach(g => ba.checked ? BSEL.add(g.order_no) : BSEL.delete(g.order_no)); renderPlan(); });
  markSel();
  const n = [...BSEL].length;
  $("#b-go").disabled = !n;
  $("#b-go-hint").textContent = n ? `已勾選 ${n} 張訂單，會產出 EIP 採購單（竹運採購進貨另附採購下採明細）。` : "請勾選要產出的訂單；「需處理」的訂單請點該列，在右側修正。";
}
$("#b-go").addEventListener("click", async () => {
  const nos = [...BSEL]; if (!nos.length) return;
  const b = $("#b-go"); b.disabled = true; $("#b-err").innerHTML = "";
  const byF = {}; nos.forEach(k => { const g = PLAN.orders.find(x => x.order_no === k); if (g) (byF[g.fulfil] = byF[g.fulfil] || []).push(k); });
  const done = [], errs = [];
  try {
    for (const [fulfil, list] of Object.entries(byF)) {       // 兩種履約方式各產一包（分檔規則不同）
      const res = await fetch("/api/shopee/purchase/generate", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ fulfil, order_nos: list }) });
      if (!res.ok) {
        let d = {}; try { d = await res.json(); } catch (e) {}
        errs.push(`<div class="alert al-bad"><b>${esc(fulfil)}：${esc(d.error || `伺服器錯誤（${res.status}）`)}</b>${(d.details || []).length ? `<div class="mt-1">${d.details.map(esc).join("<br>")}</div>` : ""}</div>`);
        continue;
      }
      const blob = await res.blob(); const m = (res.headers.get("Content-Disposition") || "").match(/filename\*=UTF-8''([^;]+)/);
      const a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = m ? decodeURIComponent(m[1]) : "蝦皮特選採購單.zip"; document.body.appendChild(a); a.click(); a.remove();
      done.push(`${fulfil} ${res.headers.get("X-Shopee-Eip-Files")} 個 EIP 採購單檔`);
      list.forEach(k => BSEL.delete(k));
    }
    $("#b-err").innerHTML = errs.join("");
    if (done.length) { toast(`已產出 ${done.join("、")}。上傳 EIP 後，請至「待回填 EIP 採購單號」回填。`); refresh(); }
  } catch (err) { toast(err.message, "err"); }
  finally { b.disabled = ![...BSEL].length; }
});

/* ── 修改下採數量（需求 10.5：除不盡不自動進位或捨去，人改＋寫原因） ── */
let ADJ = null;
function openAdj(id) {
  let p = null, g = null;
  for (const x of [SIDE && SIDE.buy, ...PLAN.orders].filter(Boolean)) { const y = x.items.find(i => i.order_id === id); if (y) { p = y; g = x; break; } }
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
    toast(`${ADJ.supplier_sku_id} 下採數量已改為 ${fmt(r.qty)}`); $("#dlg-adj").close(); refresh();
  } catch (err) { toast(err.message, "err"); }
});
$("#da-clear").addEventListener("click", async () => {
  try { await post("/api/shopee/purchase/adjust", { order_id: ADJ.order_id, qty: "" }); toast(`${ADJ.supplier_sku_id} 已回到原始數量 ${fmt(ADJ.qty)}`); $("#dlg-adj").close(); refresh(); }
  catch (err) { toast(err.message, "err"); }
});

/* ── 採購單產出：EIP 檔清單、回填 EIP 單號 ─────────────────── */
let EIPS = [];
$("#e-all").addEventListener("click", () => { $("#e-all").classList.toggle("on"); renderEips(); });
function eipList(withFilled) {
  const q = $("#f-q").value.trim().toLowerCase(), line = $("#f-line").value;
  return EIPS.filter(e => (withFilled || !e.eip_po) && (!line || e.line === line)
    && (!q || [e.filename, e.eip_po, ...e.orders].some(t => String(t || "").toLowerCase().includes(q))))
    .sort((a, b) => (!!a.eip_po - !!b.eip_po) || (a.expected_date || "").localeCompare(b.expected_date || ""));
}
function renderEips() {
  const list = eipList($("#e-all").classList.contains("on"));
  $("#e-mars").classList.toggle("hidden", !EIPS.some(e => e.mars && e.eip_po));
  $("#e-tbl").innerHTML = `<thead><tr><th>到貨日</th><th>倉別</th><th>線別</th><th class="opt">履約方式</th><th>檔名（點選重新下載）</th><th class="num opt">品項</th><th class="num">數量</th><th>蝦皮訂單</th><th class="opt">產出</th><th>EIP 採購單號</th></tr></thead><tbody>
    ${list.map(e => `<tr class="${e.eip_po ? "lock" : ""}"><td class="nw">${esc(e.expected_date)}</td><td class="nw">${esc(e.warehouse_name)}</td><td class="nw">${esc(e.line)}</td><td class="nw opt">${esc(e.fulfil === "供應商直送" ? "直送" : e.fulfil)}</td>
      <td><a class="lnk" href="/api/shopee/eips/${e.id}/file">${esc(e.filename)}</a>${e.mars ? `<div class="mt-1">${e.eip_po
        ? `<a class="btn btn-o btn-sm" href="/api/shopee/eips/${e.id}/mars_po"><i class="bi bi-file-earmark-excel"></i> 瑪氏採購單</a>`
        : `<button class="btn btn-o btn-sm" disabled title="請先回填 EIP 採購單號"><i class="bi bi-file-earmark-excel"></i> 瑪氏採購單</button>`}</div>` : ""}</td><td class="num opt">${e.rows}</td><td class="num">${fmt(e.qty)}</td>
      <td class="fn" title="${esc(e.orders.join("\n"))}">${esc(e.orders.slice(0, 2).join("、"))}${e.orders.length > 2 ? `…等 ${e.orders.length} 張` : ""}</td>
      <td class="kbd nw opt">${esc(e.created_by)}<div>${esc((e.created_at || "").slice(5, 16).replace("T", " "))}</div></td>
      <td class="nw"><input class="inp fn e-po" data-id="${e.id}" value="${esc(e.eip_po)}" placeholder="例：PO202610001" style="width:140px">
        <button class="btn ${e.eip_po ? "btn-o" : "btn-p"} btn-sm e-save" data-id="${e.id}">儲存</button>${e.eip_po ? `<div class="kbd">${esc(e.updated_by)} ${esc((e.updated_at || "").slice(5, 16).replace("T", " "))}</div>` : ""}</td></tr>`).join("")
      || `<tr><td colspan="10" class="muted" style="padding:24px;text-align:center">目前沒有待回填的 EIP 採購單。</td></tr>`}</tbody>`;
  $$("#e-tbl button.e-save").forEach(b => b.addEventListener("click", () => saveEip(b.dataset.id)));
  $$("#e-tbl input.e-po").forEach(i => i.addEventListener("keydown", e => { if (e.key === "Enter") saveEip(i.dataset.id); }));
}
$("#e-mars").addEventListener("click", async () => {
  const p = new URLSearchParams({ from: $("#f-from").value, to: $("#f-to").value });
  const res = await fetch(`/api/shopee/mars_po.zip?${p}`);
  if (!res.ok) { let d = {}; try { d = await res.json(); } catch (e) {} return toast(d.error || `伺服器錯誤（${res.status}）`, "err"); }
  const blob = await res.blob(); const m = (res.headers.get("Content-Disposition") || "").match(/filename\*=UTF-8''([^;]+)/);
  const a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = m ? decodeURIComponent(m[1]) : "蝦皮特選瑪氏採購單.zip"; document.body.appendChild(a); a.click(); a.remove();
  toast(`已下載 ${res.headers.get("X-Shopee-Files")} 份瑪氏採購單`);
});

/* ── 蝦皮倉資料（瑪氏採購單表頭 C6～F8） ─────────────────────── */
const WH_COLS = [["name", "入倉倉別", 230], ["address", "地址", 360], ["phone", "電話", 165], ["ship_to", "Ship-to", 100], ["contact", "送貨聯絡人", 120]];
async function loadWh() {
  let d; try { d = await api("/api/shopee/warehouses"); } catch (err) { $("#wh-tbl").innerHTML = `<tr><td class="muted">${esc(err.message)}</td></tr>`; return; }
  $("#wh-tbl").innerHTML = `<thead><tr><th>倉別</th>${WH_COLS.map(c => `<th>${c[1]}</th>`).join("")}<th></th></tr></thead><tbody>
    ${d.warehouses.map(w => `<tr data-code="${esc(w.code)}"><td class="nw"><b>${esc(w.short)}</b> <span class="kbd">${esc(w.code)}</span></td>
      ${WH_COLS.map(c => `<td><input class="inp wh-f" data-k="${c[0]}" value="${esc(w[c[0]])}" style="width:${c[2]}px"></td>`).join("")}
      <td><button class="btn btn-o btn-sm wh-save">儲存</button></td></tr>`).join("")}</tbody>`;
  $$("#wh-tbl button.wh-save").forEach(b => b.addEventListener("click", () => saveWh(b.closest("tr"))));
}
async function saveWh(tr) {
  const body = {}; tr.querySelectorAll("input.wh-f").forEach(i => { body[i.dataset.k] = i.value.trim(); });
  try { const r = await api(`/api/shopee/warehouses/${tr.dataset.code}`, { method: "PUT", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) });
    toast(r.changed ? `已更新 ${tr.querySelector("b").textContent} 的倉資料` : "內容沒有變更");
  } catch (err) { toast(err.message, "err"); }
}
$("#wh-box").addEventListener("toggle", () => { if ($("#wh-box").open) loadWh(); });

async function saveEip(id, value) {
  const v = String(value ?? $(`#e-tbl input.e-po[data-id="${id}"]`).value).trim();
  try { const r = await api(`/api/shopee/eips/${id}`, { method: "PUT", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ eip_po: v }) });
    toast(r.eip_po ? `已回填 EIP 採購單號 ${r.eip_po}` : "已清除 EIP 採購單號，該批訂單可重新產出"); refresh();
  } catch (err) { toast(err.message, "err"); }
}

/* ── 竹運出貨：庫存資料、設定 ─────────────────────────────── */
const KIND_ORDER = ["ec", "lots", "life"], KIND_LABEL = { ec: "EC 即時庫存", lots: "竹運庫存效期表", life: "竹運商品主檔" };
async function loadStock() {
  let d; try { d = await api("/api/shopee/stock"); } catch (err) { $("#stock-st").innerHTML = `<span class="neg">${esc(err.message)}</span>`; return; }
  $("#stock-st").innerHTML = KIND_ORDER.map(k => {
    const x = d.stock[k];
    if (!x) return `<div class="it"><b>${KIND_LABEL[k]}</b> <span class="badge b-warn">尚未上傳</span></div>`;
    const old = k !== "life" && !x.today;
    return `<div class="it ${old ? "old" : ""}"><b>${KIND_LABEL[k]}</b> <span class="kbd">${fmt(x.rows_total)} 列</span>${old ? ` <span class="badge b-warn">不是今天的檔</span>` : ""}
      <div class="kbd" title="${esc(x.filename)}">${esc((x.uploaded_at || "").slice(0, 16).replace("T", " "))}　${esc(x.uploaded_by)}${x.file_date ? `　結餘日期 ${esc(x.file_date)}` : ""}</div></div>`;
  }).join("");
  const st = d.settings; $("#s-rule").value = st.rule; $("#s-mode").value = st.mode; $("#s-ratio").value = st.ratio;
  $("#s-ratio-box").classList.toggle("hidden", st.rule !== "custom");
}
async function saveShipSet(body) {
  try { await api("/api/shopee/ship/settings", { method: "PUT", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) }); loadStock(); toast("效期標準已更新，竹運出貨與勇信核對會依新標準判斷"); }
  catch (err) { toast(err.message, "err"); loadStock(); }
}
$("#s-rule").addEventListener("change", () => saveShipSet({ rule: $("#s-rule").value }));
$("#s-mode").addEventListener("change", () => saveShipSet({ mode: $("#s-mode").value }));
$("#s-ratio").addEventListener("change", () => saveShipSet({ ratio: $("#s-ratio").value.trim() }));

/* ── 竹運出貨：訂單、出貨量、效期批次 ─────────────────────── */
let SHIP = { orders: [] }; const ZOPEN = new Set(), ZSEL = new Set();
const zq = () => fq();
const zPick = g => g.dispatch_status !== "已產出" && !g.blocking;
async function loadShip() {
  try { SHIP = await api(`/api/shopee/ship/plan?${new URLSearchParams(zq())}`); } catch (err) { $("#z-tbl").innerHTML = `<tr><td class="muted" style="padding:20px">${esc(err.message)}</td></tr>`; return; }
  const ok = new Set(SHIP.orders.filter(g => zPick(g) && inStage(g.order_no)).map(g => g.order_no));
  [...ZSEL].forEach(k => { if (!ok.has(k)) ZSEL.delete(k); });
  renderShip();
}
function renderShip() {
  const list = SHIP.orders.filter(g => inStage(g.order_no));
  const picks = list.filter(zPick), all = picks.length && picks.every(g => ZSEL.has(g.order_no));
  $("#z-sum").innerHTML = `<span><b>${list.length}</b> 張訂單</span><span><b>${fmt(list.reduce((a, g) => a + g.qty, 0))}</b> 訂購</span>`
    + `<span><b>${fmt(list.reduce((a, g) => a + g.ship, 0))}</b> 可出貨</span><span><b style="color:var(--bad)">${fmt(list.reduce((a, g) => a + g.short, 0))}</b> 缺貨</span>`
    + (list.some(g => g.blocking) ? `<span><b style="color:var(--bad)">${list.filter(g => g.blocking).length}</b> 張需先處理</span>` : "");
  let h = `<thead><tr><th><input type="checkbox" class="ck" id="z-all" ${all ? "checked" : ""} ${picks.length ? "" : "disabled"} title="全選可產出的訂單"></th>
    <th>訂單編號</th><th>線別</th><th>到貨日</th><th>倉別</th><th class="num">品項</th><th class="num">訂購量</th><th class="num">出貨量</th><th class="num">缺貨量</th><th>檢查結果</th><th>拋單檔</th></tr></thead><tbody>`;
  for (const g of list) {
    const k = g.order_no, done = g.dispatch_status === "已產出";
    const st = done ? `<span class="badge b-dim">—</span>` : g.blocking ? `<span class="badge b-diff">${g.blocking} 項需處理</span>`
      : `<span class="badge ${g.short ? "b-warn" : "b-ok"}">${g.short ? (g.deleted ? `缺貨（刪單 ${g.deleted} 品）` : "部分缺貨") : "可全數出貨"}</span>${g.warnings ? ` <span class="badge b-warn">${g.warnings} 項提醒</span>` : ""}`;
    h += `<tr class="og ${g.blocking && !done ? "blk" : ""} ${done ? "lock" : ""}" data-k="${esc(k)}">
      <td><input type="checkbox" class="ck z-ck" data-k="${esc(k)}" ${ZSEL.has(k) ? "checked" : ""} ${zPick(g) ? "" : "disabled"}></td>
      <td><span class="fn" style="font-weight:700;color:var(--ink)">${esc(k)}</span>${g.po_id && g.pr_id ? `<div class="kbd">${esc(g.pr_id)}</div>` : ""}</td>
      <td class="nw">${esc(g.line)}</td><td class="nw">${esc(g.expected_date)}</td><td class="nw">${wh(g.warehouse)}</td><td class="num">${g.items.length}</td>
      <td class="num">${fmt(g.qty)}</td><td class="num"><b>${fmt(g.ship)}</b></td><td class="num ${g.short ? "neg" : ""}">${fmt(g.short)}</td><td>${st}</td>
      <td class="nw">${done ? `<span class="badge b-s">已產出</span> <span class="kbd">${esc((g.dispatch_at || "").slice(5, 16).replace("T", " "))}</span> <button class="btn btn-o btn-sm z-undo" data-k="${esc(k)}">取消產出</button>`
        : `<span class="badge b-dim">${esc(g.dispatch_status)}</span>`}</td></tr>`;
  }
  if (!list.length) h += `<tr><td colspan="11" class="muted" style="padding:24px;text-align:center">這一步目前沒有要處理的訂單。</td></tr>`;
  $("#z-tbl").innerHTML = h + "</tbody>";
  $$("#z-tbl tr.og").forEach(tr => tr.addEventListener("click", e => { if (e.target.closest("button, a, input")) return; openSide(tr.dataset.k); }));
  $$("#z-tbl input.z-ck").forEach(c => c.addEventListener("change", () => { c.checked ? ZSEL.add(c.dataset.k) : ZSEL.delete(c.dataset.k); renderShip(); }));
  const za = $("#z-all"); if (za) za.addEventListener("change", () => { picks.forEach(g => za.checked ? ZSEL.add(g.order_no) : ZSEL.delete(g.order_no)); renderShip(); });
  markSel();
  $$("#z-tbl button.z-undo").forEach(b => b.addEventListener("click", () => undoShip(b.dataset.k)));
  const n = ZSEL.size; $("#z-go").disabled = !n;
  $("#z-go-hint").textContent = n ? `已勾選 ${n} 張訂單，會產出訂單系統拋單檔（一天一倉一個檔）＋下修刪單清單（zip）。` : "請勾選要產出的訂單；「需處理」的訂單請點該列，在右側修正。";
}
let SHIPX = null;
function openShip(id) {
  let x = null; for (const g of [SIDE && SIDE.ship, ...SHIP.orders].filter(Boolean)) { x = g.items.find(i => i.order_id === id); if (x) break; }
  if (!x) return; SHIPX = x;
  $("#ds-title").textContent = `${x.supplier_sku_id} 效期批次與出貨量`;
  $("#ds-facts").innerHTML = `<span>${esc((x.name || "").slice(0, 46))}</span><span>訂購量 <b>${fmt(x.qty)}</b> ${esc(x.unit)}</span><span>EC 可用 <b>${fmt(x.ec)}</b>${x.ec < x.qty ? `（不足 ${fmt(x.qty - x.ec)}）` : "（足夠）"}</span>
    <span>竹運帳面 <b>${fmt(x.book)}</b></span><span>符合效期 <b>${fmt(x.qualified)}</b></span><span>系統建議 <b>${fmt(x.suggest)}</b></span>
    <span>商品年限 <b>${fmt(x.life)}</b> 天（${esc(x.life_src)}）</span><span>效期標準 <b>${esc(x.ratio_text)}</b>，效期需在 <b>${esc(x.threshold)}</b> 以後</span>`;
  const sel = {}; x.alloc.forEach(a => { sel[a.exp] = a.qty; });
  $("#ds-lots").innerHTML = `<thead><tr><th>效期</th><th>製造日期</th><th class="num">到貨時剩餘天數</th><th class="num">剩餘比例</th><th>符合</th><th class="num">可出貨量</th><th class="num">建議順位</th><th>選用</th><th class="num">本次數量</th></tr></thead><tbody>
    ${x.lots.map(l => `<tr class="${l.ok ? "" : "no"}"><td class="nw">${esc(l.exp)}</td><td class="nw kbd">${esc(l.mfg) || "—"}</td><td class="num">${fmt(l.days)}</td>
      <td class="num">${l.pct == null ? "—" : Math.round(l.pct * 100) + "%"}</td><td>${l.ok ? `<span class="badge b-ok">符合</span>` : `<span class="badge b-diff">不符合</span>`}</td>
      <td class="num">${fmt(l.qty)}</td><td class="num">${l.rank || "—"}</td>
      <td>${l.ok && l.qty > 0 ? `<input type="checkbox" class="ck ds-ck" data-exp="${esc(l.exp)}" ${sel[l.exp] ? "checked" : ""}>` : ""}</td>
      <td class="num">${l.ok && l.qty > 0 ? `<input class="inp ds-q" data-exp="${esc(l.exp)}" data-max="${l.qty}" value="${sel[l.exp] || ""}" style="width:80px;text-align:right" inputmode="numeric">` : ""}</td></tr>`).join("")
      || `<tr><td colspan="9" class="muted" style="padding:14px">竹運庫存效期表裡沒有這個品號的庫存。</td></tr>`}</tbody>`;
  $$("#ds-lots .ds-ck").forEach(c => c.addEventListener("change", () => {
    const q = $(`#ds-lots .ds-q[data-exp="${c.dataset.exp}"]`);
    if (c.checked && !q.value) { const used = dsTotal(); q.value = Math.max(0, Math.min(Number(q.dataset.max), Math.min(x.qty, x.ec) - used)); }
    if (!c.checked) q.value = ""; dsSum();
  }));
  $$("#ds-lots .ds-q").forEach(q => q.addEventListener("input", () => { $(`#ds-lots .ds-ck[data-exp="${q.dataset.exp}"]`).checked = !!Number(q.value); dsSum(); }));
  $("#ds-cap").textContent = `最多 ${fmt(Math.min(x.qty, x.ec || 0, x.qualified))}（訂購量、EC 可用、符合效期取最小）`;
  $("#ds-reason").value = x.adj ? x.adj.reason : ""; $("#ds-who").textContent = x.adj ? `上次修改：${x.adj.operator}　${(x.adj.updated_at || "").slice(0, 16).replace("T", " ")}` : "";
  $("#ds-clear").classList.toggle("hidden", !x.adj); dsSum(); $("#dlg-ship").showModal();
}
const dsTotal = () => $$("#ds-lots .ds-q").reduce((a, q) => a + (Number(q.value) || 0), 0);
function dsSum() { $("#ds-total").textContent = fmt(dsTotal()); }
$("#ds-zero").addEventListener("click", () => { $$("#ds-lots .ds-q").forEach(q => { q.value = ""; }); $$("#ds-lots .ds-ck").forEach(c => { c.checked = false; }); dsSum(); $("#ds-reason").focus(); });
$("#ds-save").addEventListener("click", async () => {
  const lots = $$("#ds-lots .ds-q").filter(q => Number(q.value) > 0).map(q => ({ exp: q.dataset.exp, qty: Number(q.value) }));
  try { const r = await post("/api/shopee/ship/adjust", { order_id: SHIPX.order_id, lots, reason: $("#ds-reason").value.trim(), ...zq() });
    toast(`${SHIPX.supplier_sku_id} 出貨量已改為 ${fmt(r.qty)}`); $("#dlg-ship").close(); refresh();
  } catch (err) { toast(err.message, "err"); }
});
$("#ds-clear").addEventListener("click", async () => {
  try { await post("/api/shopee/ship/adjust", { order_id: SHIPX.order_id, clear: true }); toast(`${SHIPX.supplier_sku_id} 已回到系統建議`); $("#dlg-ship").close(); refresh(); }
  catch (err) { toast(err.message, "err"); }
});
async function undoShip(no) {
  if (!confirm(`取消 ${no} 的拋單檔產出？\n取消後會照目前的庫存重新判斷。拋單檔已匯入訂單系統的話，請記得到訂單系統刪除。`)) return;
  try { await post("/api/shopee/ship/undo", { order_no: no }); toast(`已取消 ${no} 的拋單檔產出，回到「待產拋單檔」`); refresh(); } catch (err) { toast(err.message, "err"); }
}
$("#z-go").addEventListener("click", async () => {
  const nos = [...ZSEL]; if (!nos.length) return;
  const b = $("#z-go"); b.disabled = true; $("#z-err").innerHTML = "";
  try {
    const res = await fetch("/api/shopee/ship/generate", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ order_nos: nos, ...zq() }) });
    if (!res.ok) {
      let d = {}; try { d = await res.json(); } catch (e) {}
      $("#z-err").innerHTML = `<div class="alert al-bad"><b>${esc(d.error || `伺服器錯誤（${res.status}）`)}</b>${(d.details || []).length ? `<div class="mt-1">${d.details.map(esc).join("<br>")}</div>` : ""}</div>`;
      return;
    }
    const blob = await res.blob(); const m = (res.headers.get("Content-Disposition") || "").match(/filename\*=UTF-8''([^;]+)/);
    const a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = m ? decodeURIComponent(m[1]) : "蝦皮特選竹運出貨.zip"; document.body.appendChild(a); a.click(); a.remove();
    const sh = Number(res.headers.get("X-Shopee-Short") || 0);
    toast(`已產出拋單檔${sh ? `，另有 ${sh} 品缺貨，已列入「待通知蝦皮」` : ""}。`);
    ZSEL.clear(); refresh();
  } catch (err) { toast(err.message, "err"); }
  finally { b.disabled = !ZSEL.size; }
});

/* ── 拋單檔收件資料 ───────────────────────────────────────── */
async function loadDwh() {
  let d; try { d = await api("/api/shopee/dispatch_wh"); } catch (err) { return; }
  $("#dwh-tbl").innerHTML = `<thead><tr><th>倉別</th><th>姓名</th><th>地址</th><th>電話</th><th></th></tr></thead><tbody>
    ${d.warehouses.map(w => `<tr data-code="${esc(w.code)}"><td class="nw"><b>${esc(w.short)}</b> <span class="kbd">${esc(w.code)}</span></td>
      <td><input class="inp dwh-f" data-k="name" value="${esc(w.name)}" style="width:130px"></td><td><input class="inp dwh-f" data-k="address" value="${esc(w.address)}" style="width:360px"></td>
      <td><input class="inp dwh-f" data-k="phone" value="${esc(w.phone)}" style="width:165px"></td><td><button class="btn btn-o btn-sm dwh-save">儲存</button></td></tr>`).join("")}</tbody>`;
  $$("#dwh-tbl button.dwh-save").forEach(b => b.addEventListener("click", async () => {
    const tr = b.closest("tr"), body = {}; tr.querySelectorAll("input.dwh-f").forEach(i => { body[i.dataset.k] = i.value.trim(); });
    try { const r = await api(`/api/shopee/dispatch_wh/${tr.dataset.code}`, { method: "PUT", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) });
      toast(r.changed ? `已更新 ${tr.querySelector("b").textContent} 的收件資料` : "內容沒有變更"); } catch (err) { toast(err.message, "err"); }
  }));
}
$("#dwh-box").addEventListener("toggle", () => { if ($("#dwh-box").open) loadDwh(); });

/* ── 缺貨維護與通知 ───────────────────────────────────────── */
let SHORT = { orders: [] }; const KOPEN = new Set(), KSEL = new Set();
async function loadShortData() {
  try { SHORT = await api(`/api/shopee/short/plan?${new URLSearchParams(fq())}`); } catch (err) { SHORT = { orders: [] }; throw err; }
}
async function loadShort() {
  try { await loadShortData(); } catch (err) { $("#k-tbl").innerHTML = `<tr><td class="muted" style="padding:20px">${esc(err.message)}</td></tr>`; return; }
  const ok = new Set(SHORT.orders.filter(g => g.pending && inStage(g.order_no)).map(g => g.order_no));
  [...KSEL].forEach(k => { if (!ok.has(k)) KSEL.delete(k); });
  renderShort();
}
function renderShort() {
  const list = SHORT.orders.filter(g => inStage(g.order_no));
  const picks = list.filter(g => g.pending), all = picks.length && picks.every(g => KSEL.has(g.order_no));
  $("#k-sum").innerHTML = `<span><b>${list.length}</b> 張訂單</span><span><b>${list.filter(g => g.short_items).length}</b> 張有缺貨</span>`
    + `<span><b style="color:var(--warn)">${list.reduce((a, g) => a + g.pending, 0)}</b> 品待通知</span><span><b style="color:var(--bad)">${fmt(list.reduce((a, g) => a + g.cut, 0))}</b> 累計下修</span>`;
  let h = `<thead><tr><th><input type="checkbox" class="ck" id="k-all" ${all ? "checked" : ""} ${picks.length ? "" : "disabled"} title="全選待通知的訂單"></th>
    <th>訂單編號</th><th>線別</th><th class="opt">履約方式</th><th>到貨日</th><th>倉別</th><th class="num opt">品項</th><th class="num opt">原始訂購量</th><th class="num">最終可出貨量</th><th class="num">累計下修</th><th>通知狀態</th></tr></thead><tbody>`;
  for (const g of list) {
    const k = g.order_no;
    h += `<tr class="og ${g.pending ? "pend" : ""}" data-k="${esc(k)}">
      <td><input type="checkbox" class="ck k-ck" data-k="${esc(k)}" ${KSEL.has(k) ? "checked" : ""} ${g.pending ? "" : "disabled"}></td>
      <td><span class="fn" style="font-weight:700;color:var(--ink)">${esc(k)}</span>${g.po_id && g.pr_id ? `<div class="kbd">${esc(g.pr_id)}</div>` : ""}</td>
      <td class="nw">${esc(g.line)}</td><td class="nw opt">${esc(g.fulfil)}</td><td class="nw">${esc(g.expected_date)}</td><td class="nw">${wh(g.warehouse)}</td><td class="num opt">${g.items.length}</td>
      <td class="num opt">${fmt(g.orig)}</td><td class="num"><b>${fmt(g.final)}</b></td><td class="num ${g.cut ? "neg" : ""}">${fmt(g.cut)}${g.deleted ? `<div class="kbd">刪單 ${g.deleted} 品</div>` : ""}</td>
      <td><span class="badge ${g.status === "待通知" ? "b-warn" : g.status.startsWith("已通知") ? "b-ok" : "b-dim"}">${esc(g.status)}</span></td></tr>`;
  }
  if (!list.length) h += `<tr><td colspan="11" class="muted" style="padding:24px;text-align:center">這一步目前沒有要處理的訂單。</td></tr>`;
  $("#k-tbl").innerHTML = h + "</tbody>";
  $$("#k-tbl tr.og").forEach(tr => tr.addEventListener("click", e => { if (e.target.closest("button, a, input")) return; openSide(tr.dataset.k); }));
  $$("#k-tbl input.k-ck").forEach(c => c.addEventListener("change", () => { c.checked ? KSEL.add(c.dataset.k) : KSEL.delete(c.dataset.k); renderShort(); }));
  const ka = $("#k-all"); if (ka) ka.addEventListener("change", () => { picks.forEach(g => ka.checked ? KSEL.add(g.order_no) : KSEL.delete(g.order_no)); renderShort(); });
  markSel();
  const n = KSEL.size; $("#k-go").disabled = !n;
  $("#k-go-hint").textContent = n ? `已勾選 ${n} 張訂單，會把待通知的品項整理成一則 LINE 訊息，並鎖定這次的下修紀錄。` : "請勾選有「待通知」的訂單。";
}
const roundText = r => `第 ${r.round} 次：${fmt(r.before)} → ${fmt(r.after)}${r.after ? "" : "（刪單）"}`;
let SHORTX = null;
function openShort(id) {
  let x = null; for (const g of [SIDE && SIDE.short, ...SHORT.orders].filter(Boolean)) { x = g.items.find(i => i.order_id === id); if (x) break; }
  if (!x) return; SHORTX = x;
  $("#dk-title").textContent = `${x.supplier_sku_id} 缺貨`;
  $("#dk-facts").innerHTML = `<span>${esc((x.name || "").slice(0, 46))}</span><span>原始訂購量 <b>${fmt(x.orig)}</b> ${esc(x.unit)}</span><span>目前可出貨量 <b>${fmt(x.final)}</b></span><span>累計下修 <b>${fmt(x.cut)}</b></span>`;
  $("#dk-hist-wrap").classList.toggle("hidden", !x.rounds.length);
  $("#dk-hist").innerHTML = `<thead><tr><th>次數</th><th>來源</th><th class="num">之前</th><th class="num">之後</th><th class="num">這次缺</th><th>原因</th><th>說明</th><th>填寫</th><th>通知</th></tr></thead><tbody>
    ${x.rounds.map(r => `<tr><td>第 ${r.round} 次</td><td class="kbd">${esc(r.source)}</td><td class="num">${fmt(r.before)}</td><td class="num"><b>${fmt(r.after)}</b></td><td class="num neg">${fmt(r.short)}</td>
      <td class="kbd">${esc(r.reasons.join("、"))}</td><td class="kbd">${esc(r.note)}</td><td class="kbd nw">${esc(r.operator)}<div>${esc((r.updated_at || "").slice(5, 16).replace("T", " "))}</div></td>
      <td class="nw">${r.notified ? `<span class="badge b-ok">已通知</span><div class="kbd">${esc((r.notify_at || "").slice(5, 16).replace("T", " "))}</div>` : `<span class="badge b-warn">待通知</span>`}</td></tr>`).join("")}</tbody>`;
  const editing = x.rounds.find(r => r.round === x.edit_round);
  $("#dk-form").classList.toggle("hidden", !x.edit_round);
  $("#dk-lock").classList.toggle("hidden", !!x.edit_round);
  $("#dk-lock").textContent = `已通知 ${x.max_rounds} 次（${x.max_rounds === 1 ? "紙潔、瑪氏只下修一次" : "寶僑最多兩次"}），不能再填。`;
  $("#dk-round").textContent = editing ? `修改第 ${x.edit_round} 次（尚未通知）` : `第 ${x.edit_round} 次下修`;
  $("#dk-prev").textContent = `這次之前可出貨 ${fmt(x.prev_final)}`;
  $("#dk-short").value = editing ? editing.short : ""; $("#dk-note").value = editing ? editing.note : "";
  $$(".dk-r").forEach(c => { c.checked = editing ? editing.reasons.includes(c.value) : c.value === "庫存不足"; });
  dkAfter(); $("#dlg-short").showModal();
}
function dkAfter() { const v = Number($("#dk-short").value || 0); $("#dk-after").textContent = SHORTX ? `這次之後可出貨 ${fmt(Math.max(SHORTX.prev_final - v, 0))}${SHORTX.prev_final - v <= 0 ? "（刪單）" : ""}` : ""; }
$("#dk-short").addEventListener("input", dkAfter);
$("#dk-all").addEventListener("click", () => { $("#dk-short").value = SHORTX.prev_final; dkAfter(); });
$("#dk-save").addEventListener("click", async () => {
  try { const r = await post("/api/shopee/short/save", { order_id: SHORTX.order_id, short: $("#dk-short").value.trim() || "0",
      reasons: $$(".dk-r").filter(c => c.checked).map(c => c.value), note: $("#dk-note").value.trim() });
    toast(r.cleared ? `${SHORTX.supplier_sku_id} 已取消這次缺貨` : `${SHORTX.supplier_sku_id} 第 ${r.round} 次下修：可出貨 ${fmt(r.after)}，已列入「待通知蝦皮」`); $("#dlg-short").close(); refresh();
  } catch (err) { toast(err.message, "err"); }
});
function showNotify(n) {
  const msgs = n.messages && n.messages.length ? n.messages : [n.text];
  $("#dn-title").textContent = `蝦皮缺貨通知（${(n.created_at || "").slice(0, 16).replace("T", " ")}，共 ${msgs.length} 則）`;
  $("#dn-xlsx").href = `/api/shopee/short/notify/${n.id}/xlsx`;
  $("#dn-msgs").innerHTML = msgs.map((m, i) => `<div class="pv"><div class="flex items-center gap-2 mb-2"><b>第 ${i + 1} 則</b><span class="kbd">${esc(m.split("\n")[0])}</span><div class="flex-1"></div>
    <button class="btn btn-p btn-sm dn-copy" data-i="${i}"><i class="bi bi-clipboard"></i> 複製</button></div>
    <textarea class="inp dn-text" readonly style="width:100%;height:${Math.min(60 + m.split("\n").length * 21, 420)}px;font-family:inherit;line-height:1.55">${esc(m)}</textarea></div>`).join("");
  $$("#dn-msgs button.dn-copy").forEach(b => b.addEventListener("click", async () => {
    const ta = $$("#dn-msgs textarea.dn-text")[Number(b.dataset.i)];
    try { await navigator.clipboard.writeText(ta.value); } catch (e) { ta.select(); document.execCommand("copy"); }
    b.innerHTML = `<i class="bi bi-check2"></i> 已複製`; toast(`第 ${Number(b.dataset.i) + 1} 則已複製，可以貼到 LINE`);
  }));
  $("#dlg-notify").showModal();
}
$("#k-go").addEventListener("click", async () => {
  const nos = [...KSEL]; if (!nos.length) return;
  if (!confirm(`產生通知後，這次的下修紀錄會鎖定（之後再缺要填下一次）。確定產生？`)) return;
  try { const r = await post("/api/shopee/short/notify", { order_nos: nos }); KSEL.clear(); refresh(); showNotify(await api(`/api/shopee/short/notify/${r.id}`)); }
  catch (err) { toast(err.message, "err"); }
});
/* 瑪氏：勇信配送明細表核對（需求 8.1） */
let YX = null;
const YX_ST = { ok: ["b-ok", "無缺貨"], partial: ["b-warn", "部分缺貨"], none: ["b-diff", "全缺"], over: ["b-diff", "數量異常"] };
async function yxUpload(files) {
  const fd = new FormData(); files.forEach(f => fd.append("file", f));
  try { showYx(await api("/api/shopee/yx/compare", { method: "POST", body: fd })); }
  catch (err) { toast(err.message, "err"); }
}
function showYx(d) { YX = d; renderYx(); if (!$("#dlg-yx").open) $("#dlg-yx").showModal(); }
async function yxRerun(skip) {
  try { YX = await post("/api/shopee/yx/compare", { pages: YX.pages, files: YX.files, skip_eips: skip }); renderYx(); } catch (err) { toast(err.message, "err"); }
}
function renderYx() {
  const d = YX, s = d.summary, al = [];
  if (d.unknown.length) al.push(`<div class="alert al-warn"><b>表上有 ${d.unknown.length} 個 EIP 採購單號在蝦皮特選的瑪氏 EIP 採購單中找不到</b>（可能是酷澎的單，或尚未回填），這幾頁不處理：<br>${d.unknown.map(u => `${esc(u.eip_po)}（第 ${u.pages.join("、")} 頁，${fmt(u.cases)} 箱）`).join("、")}</div>`);
  if (d.no_eip.length) al.push(`<div class="alert al-bad"><b>同一天同一倉還有 ${d.no_eip.length} 份瑪氏 EIP 採購單尚未回填 EIP 採購單號</b>，無法與勇信配送明細表核對（不列為缺貨）。請先至「待回填 EIP 採購單號」回填後重新上傳：<br>${d.no_eip.map(b => esc(b.filename)).join("<br>")}</div>`);
  if (s.absent) al.push(`<div class="alert al-bad"><i class="bi bi-exclamation-octagon-fill"></i> 有 <b>${s.absent}</b> 份已回填 EIP 採購單號，但勇信配送明細表上沒有，視為整份未出貨（全部刪單）。本次原本就不出貨的，請取消勾選。</div>`);
  const extras = d.eips.flatMap(b => b.extra.map(x => `${esc(b.eip_po)}：勇信多出 ${esc(x.mars_code)} ${esc(x.name)} ${fmt(x.cases)} 箱（未訂購）`));
  if (extras.length) al.push(`<div class="alert al-warn"><b>勇信配送明細表上有未訂購的瑪氏貨號</b>（系統不處理，請人工確認）：<br>${extras.join("<br>")}</div>`);
  if (d.page_warnings.length) al.push(`<div class="alert al-warn">${d.page_warnings.map(esc).join("<br>")}</div>`);
  let h = `<div class="sum"><span><b>${fmt(s.pages)}</b> 頁</span><span><b>${fmt(s.files)}</b> 份 EIP</span><span><b>${fmt(s.items)}</b> 品</span>`
    + `<span><b style="color:var(--warn)">${fmt(s.partial)}</b> 部分缺貨</span><span><b style="color:var(--bad)">${fmt(s.none)}</b> 全缺</span>`
    + (s.expiry_short ? `<span><b style="color:var(--bad)">${fmt(s.expiry_short)}</b> 品效期不足</span>` : "")
    + (s.over ? `<span><b style="color:var(--bad)">${fmt(s.over)}</b> 數量異常</span>` : "") + (s.warns ? `<span><b style="color:var(--warn)">${fmt(s.warns)}</b> 項提醒</span>` : "")
    + `<span class="kbd">${esc(d.dates.join("、"))} ${esc(d.warehouses.map(w => WH[w] || w).join("、"))}</span></div>` + al.join("");
  h += `<div class="tbl-wrap" style="max-height:52vh"><table class="t"><thead><tr><th>訂單編號</th><th>廠商料號</th><th>瑪氏貨號</th><th>品名</th><th>單位</th>
    <th class="num">訂購（箱）</th><th class="num">勇信出貨（箱）</th><th>效期（每批）</th><th class="num">缺貨（箱）</th><th>結果</th><th class="num">可出貨量</th></tr></thead><tbody>`;
  for (const b of d.eips) {
    const tag = b.skipped ? `<span class="badge b-dim">這次不核</span>` : b.in_pdf ? `<span class="badge b-s">表上第 ${b.pages.join("、")} 頁</span>` : `<span class="badge b-diff">表上沒有，視為整份未出貨</span>`;
    h += `<tr class="og"><td colspan="11"><label class="nw" style="cursor:pointer"><input type="checkbox" class="ck yx-ck" data-id="${b.eip_id}" ${b.skipped ? "" : "checked"}> 核對</label>
      <b class="fn" style="margin-left:10px">${esc(b.eip_po)}</b> <span class="kbd">${esc(b.expected_date)} 到貨 · ${esc(b.warehouse_name)} · ${esc(b.unit_group)}</span> ${tag}</td></tr>`;
    for (const r of b.items) {
      const [cls, txt] = YX_ST[r.status] || ["b-dim", "—"];
      const res = b.skipped ? `<span class="kbd">這次不核</span>` : `<span class="badge ${cls}">${txt}</span> <span class="kbd nw">${esc(r.action)}</span>${r.reasons.length ? `<div class="kbd nw">原因：${esc(r.reasons.join("、"))}</div>` : ""}${r.was && r.action !== "已通知，不再修改" ? `<div class="kbd">前次寫入：可出貨量 ${fmt(r.was.after)}</div>` : ""}`;
      h += `<tr class="${b.skipped ? "no" : ""}"><td class="fn">${esc(r.order_no)}</td><td class="fn">${esc(r.supplier_sku_id)}</td><td class="fn">${esc(r.mars_code)}</td><td>${esc((r.name || "").slice(0, 36))}</td><td class="nw">${esc(r.unit)}</td>
        <td class="num">${fmt(r.ordered)}</td><td class="num"><b>${fmt(r.shipped)}</b></td>
        <td class="kbd nw">${r.batches.map(x => `<div${x.ok ? "" : ` class="neg"`}>${esc(x.expiry || "未印效期")} × ${fmt(x.qty)}${x.ok ? "" : "（效期不足）"}</div>`).join("") || "—"}${r.need ? `<div>允收 ${esc(r.need)} 以後</div>` : ""}</td>
        <td class="num ${r.short ? "neg" : ""}">${fmt(r.short)}</td><td>${res}${r.warn ? `<div class="kbd" style="color:var(--warn)"><i class="bi bi-info-circle"></i> ${esc(r.warn)}</div>` : ""}</td>
        <td class="num nw">${r.final != null && r.final < r.qty ? `${fmt(r.qty)} → <b class="neg">${fmt(r.final)}</b>` : fmt(r.qty)}</td></tr>`;
    }
  }
  h += `</tbody></table></div><div class="flex items-center gap-3 flex-wrap"><button id="yx-go" class="btn btn-p" ${s.changes ? "" : "disabled"}><i class="bi bi-check2-circle"></i> 確認，寫入缺貨紀錄</button>
    <span class="kbd">${s.changes ? `將寫入或更新 ${s.changes} 筆，完成後於「待通知蝦皮」勾選訂單產生蝦皮通知。` : "無需寫入的缺貨。"}${s.locked ? `已通知的 ${s.locked} 品不再修改。` : ""}</span>
    <div class="flex-1"></div><button id="yx-clear" class="btn btn-o btn-sm">關閉</button></div>`;
  $("#yx-out").innerHTML = h;
  $$("#yx-out input.yx-ck").forEach(c => c.addEventListener("change", () => yxRerun($$("#yx-out input.yx-ck").filter(x => !x.checked).map(x => Number(x.dataset.id)))));
  $("#yx-clear").addEventListener("click", () => { YX = null; $("#yx-out").innerHTML = ""; $("#dlg-yx").close(); });
  $("#yx-go").addEventListener("click", async () => {
    try { const r = await post("/api/shopee/yx/apply", { pages: YX.pages, files: YX.files, skip_eips: YX.skip_eips });
      toast(`已寫入 ${r.written} 筆缺貨${r.cleared ? `、取消 ${r.cleared} 筆` : ""}，請於「待通知蝦皮」勾選訂單通知蝦皮`);
      await yxRerun(YX.skip_eips); refresh();
    } catch (err) { toast(err.message, "err"); }
  });
}
async function loadNotifies() {
  let d; try { d = await api("/api/shopee/short/notifies"); } catch (err) { return; }
  $("#k-hist").innerHTML = `<thead><tr><th>通知時間</th><th>產生人員</th><th class="num">訂單</th><th class="num">品項</th><th></th></tr></thead><tbody>
    ${d.notifies.map(n => `<tr><td class="kbd">${esc((n.created_at || "").slice(0, 16).replace("T", " "))}</td><td>${esc(n.created_by)}</td><td class="num">${n.orders}</td><td class="num">${n.items}</td>
      <td class="nw"><button class="btn btn-o btn-sm k-n" data-id="${n.id}">再看一次／複製</button> <a class="btn btn-o btn-sm" href="/api/shopee/short/notify/${n.id}/xlsx">明細 Excel</a></td></tr>`).join("")
      || `<tr><td colspan="5" class="muted" style="padding:16px">尚無通知紀錄</td></tr>`}</tbody>`;
  $$("#k-hist button.k-n").forEach(b => b.addEventListener("click", async () => { try { showNotify(await api(`/api/shopee/short/notify/${b.dataset.id}`)); } catch (err) { toast(err.message, "err"); } }));
}

/* ── 右側面板：一張單的全部（改版第二次） ─────────────────────── */
let SIDE = null, SIDE_NO = "";
$("#sd-close").addEventListener("click", closeSide);
document.addEventListener("keydown", e => { if (e.key === "Escape" && SIDE_NO && !document.querySelector("dialog[open]")) closeSide(); });
function markSel() { $$("table.t tr.og[data-k]").forEach(tr => tr.classList.toggle("sel", !!SIDE_NO && tr.dataset.k === SIDE_NO)); }
function closeSide() { SIDE = null; SIDE_NO = ""; $("#side").classList.add("hidden"); document.body.classList.remove("side-open"); markSel(); }
async function openSide(no) {
  const again = SIDE_NO === no && SIDE;
  SIDE_NO = no; $("#side").classList.remove("hidden"); document.body.classList.add("side-open"); markSel();
  if (!again) { $("#sd-title").innerHTML = `<b class="fn" style="font-size:15px;color:var(--ink)">${esc(no)}</b>`; $("#sd-body").innerHTML = `<div class="kbd" style="padding:12px 0">讀取中…</div>`; }
  let d; try { d = await api(`/api/shopee/order?${new URLSearchParams({ ...fq(), no })}`); }
  catch (err) { if (SIDE_NO === no) $("#sd-body").innerHTML = `<div class="alert al-bad">${esc(err.message)}</div>`; return; }
  if (SIDE_NO !== no) return;                      // 讀的時候又點了別張
  const top = $("#sd-body").scrollTop; SIDE = d; renderSide(); if (again) $("#sd-body").scrollTop = top;
}
const byId = grp => new Map(((grp || {}).items || []).map(x => [x.order_id, x]));
function renderSide() {
  const d = SIDE, g = d.order, buy = byId(d.buy), ship = byId(d.ship), sh = byId(d.short);
  const pendDiffs = d.diffs.filter(x => !x.confirmed);
  $("#sd-title").innerHTML = `<b class="fn" style="font-size:15px;color:var(--ink)">${esc(g.order_no)}</b>${g.po_id && g.pr_id ? `<div class="kbd fn">${esc(g.pr_id)}</div>` : ""}
    <div class="facts" style="margin-top:4px"><span>${esc(g.line)}</span><span>${wh(g.warehouse)}</span><span>到貨 <b>${esc(g.expected_date)}</b></span>
      <span>${esc(g.fulfil)} <a class="lnk sd-ff" title="修改履約方式">改</a></span></div>
    <div class="mt-1">${nextBadges(g)}</div>`;
  let h = "";
  if (pendDiffs.length) h += `<div class="alert al-bad"><b><i class="bi bi-exclamation-triangle-fill"></i> 蝦皮重傳後有 ${pendDiffs.length} 處不同</b>，確認後才能往下處理：
      <div class="mt-1">${pendDiffs.map(x => `<div>${esc(x.supplier_sku_id)}　${esc(x.field_label)}：${esc(x.old_value)} → <b>${esc(x.new_value)}</b>${x.qty_delta ? `（${delta(x.qty_delta)}）` : ""}</div>`).join("")}</div>
      <button class="btn btn-p btn-sm mt-2 sd-dconf"><i class="bi bi-check2-circle"></i> 確認差異</button></div>`;

  // 品項：每一品把各步的資料放在一起，該改的按鈕就在旁邊
  h += `<div class="sd-sec"><div class="h"><i class="bi bi-box-seam"></i> 品項（${g.item_count}${g.removed_count ? `，移除 ${g.removed_count}` : ""}）</div><div class="b">`;
  for (const i of g.items) {
    const p = buy.get(i.id), z = ship.get(i.id), k = sh.get(i.id), unit = i.unit_name || i.unit_code;
    const kv = [`原始 <b>${fmt(i.qty_original)}</b>`, i.inbound_ids ? `入庫單 <span class="fn">${esc(i.inbound_ids)}</span>` : "", `<a class="lnk ver" data-id="${i.id}" data-t="${esc(i.supplier_sku_id)}">第 ${i.version} 版</a>`];
    let more = "", acts = "";
    if (p) {
      more += `<div class="sd-kv"><span>轉換率 <b>${fmt(p.conv)}</b></span><span>下採 <b>${fmt(p.use_qty)}</b>${p.adj ? "（已修改）" : ""}</span>
        <span>EIP <b>${p.out_qty == null ? "—" : fmt(p.out_qty)}</b> ${esc(p.out_unit || "")}${p.cases != null && p.out_unit !== "箱" ? `（${fmt(p.cases)} 箱）` : ""}</span>${p.unit_group ? `<span>${esc(p.unit_group)}</span>` : ""}</div>`;
      if (g.todo.includes("buy")) more += p.problems.map(x => `<div class="prob"><i class="bi bi-exclamation-circle"></i> ${esc(x)}</div>`).join("");   // 已產出的單不再列產出前的檢查
      if ((p.rem || p.adj) && !(p.eip && p.eip.eip_po)) acts += `<button class="btn ${p.rem ? "btn-p" : "btn-o"} btn-sm sd-adj" data-id="${i.id}"><i class="bi bi-pencil"></i> 修改數量</button>`;
    }
    if (z) {
      more += `<div class="sd-kv">${z.dispatch ? "" : `<span>EC 可用 <b>${fmt(z.ec)}</b></span><span>符合效期 <b>${fmt(z.qualified)}</b></span>`}<span>出貨 <b>${fmt(z.final)}</b>${z.adj ? "（已修改）" : ""}</span>
        ${z.short ? `<span class="neg">缺 ${fmt(z.short)}${z.reasons.length ? `（${esc(z.reasons.join("、"))}）` : ""}</span>` : ""}${z.alloc.length ? `<span>效期 ${z.alloc.map(a => `${esc(a.exp)}×${fmt(a.qty)}`).join("、")}</span>` : ""}</div>`;
      more += z.problems.map(x => `<div class="prob"><i class="bi bi-exclamation-circle"></i> ${esc(x)}</div>`).join("") + z.warnings.map(x => `<div class="kbd" style="color:var(--warn)"><i class="bi bi-info-circle"></i> ${esc(x)}</div>`).join("");
      if (!z.dispatch && z.life && z.lots.length) acts += `<button class="btn ${z.warnings.length ? "btn-p" : "btn-o"} btn-sm sd-ship" data-id="${i.id}"><i class="bi bi-box-seam"></i> 效期批次／改出貨量</button>`;
    }
    if (k) {
      if (k.rounds.length) more += `<div class="sd-kv"><span>可出貨 <b>${fmt(k.final)}</b></span><span class="neg">累計下修 ${fmt(k.cut)}</span>
        ${k.pending ? `<span class="badge b-warn">待通知</span>` : k.notified_round ? `<span class="badge b-ok">已通知第 ${k.notified_round} 次</span>` : ""}</div>
        <div class="kbd">${k.rounds.map(r => `${roundText(r)}（${esc(r.source)}${r.reasons.length ? "，" + esc(r.reasons.join("、")) : ""}）`).join("<br>")}</div>`;
      if (k.edit_round) acts += `<button class="btn ${k.pending ? "btn-p" : "btn-o"} btn-sm sd-short" data-id="${i.id}"><i class="bi bi-dash-circle"></i> ${k.rounds.length && !k.pending ? `第 ${k.edit_round} 次缺貨` : k.pending ? "修改缺貨" : "填缺貨"}</button>`;
      else if (g.line === "瑪氏" && !i.removed) more += `<div class="kbd">缺貨由勇信配送明細表核對（PDF 丟最上方上傳）</div>`;
    }
    h += `<div class="sd-it ${i.removed ? "muted" : ""}"><div class="r"><span class="fn" style="font-weight:700;color:var(--ink)">${esc(i.supplier_sku_id)}</span>
        <span style="flex:1;min-width:0">${esc((i.sku_name || "").slice(0, 40))}</span><span class="nw"><b style="font-size:15px">${fmt(i.qty)}</b> ${esc(unit)}</span></div>
      <div class="sd-kv">${kv.filter(Boolean).map(x => `<span>${x}</span>`).join("")}${i.removed ? `<span class="badge b-diff">已移除</span>` : i.diff_pending ? `<span class="badge b-diff">差異待確認</span>` : ""}</div>
      ${more}${acts ? `<div class="sd-acts">${acts}</div>` : ""}</div>`;
  }
  h += `</div></div>`;

  // EIP 採購單（竹運採購進貨、直送）
  if (d.buy) h += `<div class="sd-sec"><div class="h"><i class="bi bi-file-earmark-spreadsheet"></i> EIP 採購單 <span class="kbd">${d.eips.length ? `${d.eips.length} 份` : "尚未產出"}</span></div><div class="b">
    ${d.eips.map(e => `<div class="sd-it"><a class="lnk" href="/api/shopee/eips/${e.id}/file">${esc(e.filename)}</a>
      <div class="flex items-center gap-2 mt-1 flex-wrap"><input class="inp fn sd-po" data-id="${e.id}" value="${esc(e.eip_po)}" placeholder="EIP 採購單號，例：PO202610001" style="width:220px">
        <button class="btn ${e.eip_po ? "btn-o" : "btn-p"} btn-sm sd-po-save" data-id="${e.id}">儲存</button>
        ${e.mars ? (e.eip_po ? `<a class="btn btn-o btn-sm" href="/api/shopee/eips/${e.id}/mars_po"><i class="bi bi-file-earmark-excel"></i> 瑪氏採購單</a>` : `<span class="kbd">回填後可下載瑪氏採購單</span>`) : ""}</div>
      ${e.eip_po ? `<div class="kbd">${esc(e.updated_by)} ${esc((e.updated_at || "").slice(5, 16).replace("T", " "))} 回填</div>` : ""}</div>`).join("")
      || `<div class="kbd">尚未產出。</div>`}
    ${g.todo.includes("buy") ? (d.buy.blocking ? `<div class="kbd neg mt-1">先修正上面標紅的品項，才能產出 EIP 採購單。</div>`
        : `<div class="sd-acts"><button class="btn btn-p btn-sm sd-regen"><i class="bi bi-file-earmark-zip"></i> 產出這張單的 EIP 採購單</button></div>`)
      : d.eips.some(e => !e.eip_po) ? `<div class="sd-acts"><button class="btn btn-o btn-sm sd-regen"><i class="bi bi-arrow-repeat"></i> 重新產出 EIP 採購單</button><span class="kbd">改過數量後用；還沒回填單號的那份會換掉。</span></div>` : ""}</div></div>`;

  // 竹運出貨：拋單檔
  if (d.ship) h += `<div class="sd-sec"><div class="h"><i class="bi bi-truck"></i> 拋單檔 <span class="kbd">${d.ship.dispatch_status === "已產出" ? `已產出 ${esc((d.ship.dispatch_at || "").slice(5, 16).replace("T", " "))}` : "尚未產出"}</span></div><div class="b">
    ${d.ship.alone ? `<div class="kbd" style="color:var(--warn)">這張單不在目前的到貨日／線別範圍內，出貨量只就這張單試算；要跟別張一起分配庫存，請調整上方篩選。</div>` : ""}
    ${d.ship.dispatch_status === "已產出" ? `<button class="btn btn-o btn-sm sd-undo"><i class="bi bi-arrow-counterclockwise"></i> 取消拋單檔產出</button> <span class="kbd">拋單檔已匯入訂單系統的話，請記得到訂單系統刪除。</span>`
      : d.ship.blocking ? `<div class="kbd neg">先處理上面標紅的品項，才能產出拋單檔。</div>`
      : `<div class="sd-acts"><button class="btn btn-p btn-sm sd-zgen"><i class="bi bi-file-earmark-zip"></i> 產出這張單的拋單檔</button></div>`}</div></div>`;

  // 缺貨待通知：這張單可以直接產通知
  if (d.short && d.short.pending) h += `<div class="sd-sec"><div class="h"><i class="bi bi-chat-dots"></i> 通知蝦皮 <span class="kbd">${d.short.pending} 品待通知</span></div><div class="b">
    <div class="sd-acts"><button class="btn btn-p btn-sm sd-notify"><i class="bi bi-chat-dots"></i> 產生這張單的蝦皮通知（LINE）</button><span class="kbd">產生後這次的下修紀錄會鎖定。</span></div></div></div>`;

  // 差異紀錄、歷程（平常收起來）
  if (d.diffs.length) h += `<details class="sd-sec"><summary class="h"><i class="bi bi-arrow-left-right"></i> 差異紀錄 <span class="kbd">${d.diffs.length} 筆</span></summary><div class="b">
    ${d.diffs.map(x => `<div class="sd-it"><span class="fn">${esc(x.supplier_sku_id)}</span>　${esc(x.field_label)}：${esc(x.old_value)} → <b>${esc(x.new_value)}</b>
      <div class="kbd">${esc(x.uploaded_by)} ${esc((x.uploaded_at || "").slice(5, 16))} 上傳　${x.confirmed ? `${esc(x.confirmed_by)} ${esc((x.confirmed_at || "").slice(5, 16))} 確認` : `<span class="neg">待確認</span>`}</div></div>`).join("")}</div></details>`;
  h += `<details class="sd-sec"><summary class="h"><i class="bi bi-clock-history"></i> 歷程 <span class="kbd">${d.logs.length} 筆</span></summary><div class="b">
    ${d.logs.map(l => `<div class="sd-it"><span class="kbd">${esc((l.changed_at || "").slice(5, 16).replace("T", " "))}　${esc(l.operator)}</span>　<b>${esc(l.field_label)}</b>
      ${l.old_value || l.new_value ? `<div>${esc(l.old_value)}${l.old_value && l.new_value ? " → " : ""}${esc(l.new_value)}</div>` : ""}${l.note ? `<div class="kbd">${esc(l.note)}</div>` : ""}</div>`).join("")
      || `<div class="kbd">尚無紀錄</div>`}</div></details>`;
  $("#sd-body").innerHTML = h;

  $$("#side .sd-ff").forEach(b => b.addEventListener("click", () => openFulfil(g.order_no)));
  $$("#side .sd-dconf").forEach(b => b.addEventListener("click", async () => {
    try { await post("/api/shopee/diffs/confirm", { order_nos: [g.order_no] }); toast(`已確認 ${g.order_no} 的差異`); refresh(); } catch (err) { toast(err.message, "err"); }
  }));
  $$("#side a.ver").forEach(a => a.addEventListener("click", () => openVersions(a.dataset.id, a.dataset.t)));
  $$("#side .sd-adj").forEach(b => b.addEventListener("click", () => openAdj(Number(b.dataset.id))));
  $$("#side .sd-ship").forEach(b => b.addEventListener("click", () => openShip(Number(b.dataset.id))));
  $$("#side .sd-short").forEach(b => b.addEventListener("click", () => openShort(Number(b.dataset.id))));
  $$("#side .sd-po-save").forEach(b => b.addEventListener("click", () => saveEip(b.dataset.id, $(`#side .sd-po[data-id="${b.dataset.id}"]`).value)));
  $$("#side .sd-po").forEach(i => i.addEventListener("keydown", e => { if (e.key === "Enter") saveEip(i.dataset.id, i.value); }));
  $$("#side .sd-undo").forEach(b => b.addEventListener("click", () => undoShip(g.order_no)));
  $$("#side .sd-zgen").forEach(b => b.addEventListener("click", async () => {
    b.disabled = true;
    const q = d.ship.alone ? { from: g.expected_date, to: g.expected_date, line: "" } : fq();
    try {
      const res = await fetch("/api/shopee/ship/generate", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ order_nos: [g.order_no], ...q }) });
      if (!res.ok) { let e = {}; try { e = await res.json(); } catch (x) {} toast([e.error || `伺服器錯誤（${res.status}）`, ...(e.details || []).slice(0, 5)].join("\n"), "err"); return; }
      const blob = await res.blob(); const m = (res.headers.get("Content-Disposition") || "").match(/filename\*=UTF-8''([^;]+)/);
      const a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = m ? decodeURIComponent(m[1]) : "蝦皮特選竹運出貨.zip"; document.body.appendChild(a); a.click(); a.remove();
      const sh = Number(res.headers.get("X-Shopee-Short") || 0);
      toast(`已產出 ${g.order_no} 的拋單檔${sh ? `，另有 ${sh} 品缺貨，已列入「待通知蝦皮」` : ""}。`); refresh();
    } catch (err) { toast(err.message, "err"); } finally { b.disabled = false; }
  }));
  $$("#side .sd-notify").forEach(b => b.addEventListener("click", async () => {
    if (!confirm(`產生 ${g.order_no} 的蝦皮通知後，這次的下修紀錄會鎖定（之後再缺要填下一次）。確定產生？`)) return;
    try { const r = await post("/api/shopee/short/notify", { order_nos: [g.order_no] }); refresh(); showNotify(await api(`/api/shopee/short/notify/${r.id}`)); }
    catch (err) { toast(err.message, "err"); }
  }));
  $$("#side .sd-regen").forEach(b => b.addEventListener("click", async () => {
    b.disabled = true;
    try {
      const res = await fetch("/api/shopee/purchase/generate", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ fulfil: d.buy.fulfil, order_nos: [g.order_no] }) });
      if (!res.ok) { let e = {}; try { e = await res.json(); } catch (x) {} toast([e.error || `伺服器錯誤（${res.status}）`, ...(e.details || []).slice(0, 5)].join("\n"), "err"); return; }
      const blob = await res.blob(); const m = (res.headers.get("Content-Disposition") || "").match(/filename\*=UTF-8''([^;]+)/);
      const a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = m ? decodeURIComponent(m[1]) : "蝦皮特選採購單.zip"; document.body.appendChild(a); a.click(); a.remove();
      toast(`已產出 ${g.order_no} 的 EIP 採購單。上傳 EIP 後，請在這裡或「待回填 EIP 採購單號」回填。`); refresh();
    } catch (err) { toast(err.message, "err"); } finally { b.disabled = false; }
  }));
  $$("#side .go-st").forEach(b => b.addEventListener("click", () => setStage(b.dataset.s)));
}

/* ── 資料與設定 ───────────────────────────────────────────── */
function showSettings(on) {
  $("#view-work").classList.toggle("hidden", on); $("#view-set").classList.toggle("hidden", !on); $("#btn-set").classList.toggle("on", on);
  if (on) { loadConv(); loadStock(); loadUploads(); loadNotifies(); window.scrollTo(0, 0); }
  else { loadFresh(); refresh(); }
}
$("#btn-set").addEventListener("click", () => showSettings($("#view-set").classList.contains("hidden")));
$("#set-back").addEventListener("click", () => showSettings(false));

loadFresh(); refresh();

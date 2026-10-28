/* MOMO 訂單管理系統（/momo）——版面照蝦皮特選（Jerry 2026-10-28 看過示意圖「好 開始做」）
   - 最上面一個入口收檔案：MOMO 下載檔可一次選好幾個（合在一起預覽）；價格本、瑪氏商品總表送去原本那支（跟蝦皮特選共用）
   - 作業流程格子：差異待確認 → 待產 EIP 採購單 → 待回填 EIP 採購單號 →（待核缺貨與效期、待畫入庫通知單：下一批）→ 全部入庫單
   - 到貨日、上傳日、線別、倉別、搜尋一組篩選，各步共用；改了馬上出結果
   - 點一張入庫單，右邊打開這張單的全部（品項與下採、EIP、差異、歷程） */
const $ = s => document.querySelector(s), $$ = s => [...document.querySelectorAll(s)];
const esc = s => String(s ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const fmt = n => n == null || n === "" ? "—" : Number(n).toLocaleString("zh-TW");
const md = t => t ? `${Number(t.slice(5, 7))}/${Number(t.slice(8, 10))}` : "";
const WD = "日一二三四五六";
const mdw = t => t ? `${md(t)}（${WD[new Date(t + "T00:00:00").getDay()]}）` : "—";
const todayIso = () => new Date(Date.now() - new Date().getTimezoneOffset() * 6e4).toISOString().slice(0, 10);
const when = t => t ? `${t.slice(0, 10) === todayIso() ? "今天" : md(t)}${t.length >= 16 ? " " + t.slice(11, 16) : ""}` : "";
function toast(msg, kind) {
  const t = document.createElement("div"); t.className = "toast" + (kind === "err" ? " err" : ""); t.textContent = msg; $("#toasts").appendChild(t);
  setTimeout(() => t.remove(), kind === "err" ? 7000 : msg.length > 60 ? 7000 : 2800);
}
async function api(url, opts = {}) {
  const res = await fetch(url, opts); let data = {}; try { data = await res.json(); } catch (e) {}
  if (!res.ok) throw Object.assign(new Error([data.error || `伺服器錯誤（${res.status}）`, ...(data.details || []).slice(0, 8)].join("\n")), { status: res.status, data });
  return data;
}
const post = (url, body, method = "POST") => api(url, { method, headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) });
async function download(url, opts) {
  const res = await fetch(url, opts);
  if (!res.ok) { let d = {}; try { d = await res.json(); } catch (e) {} throw new Error([d.error || `伺服器錯誤（${res.status}）`, ...(d.details || []).slice(0, 8)].join("\n")); }
  const cd = res.headers.get("Content-Disposition") || "", m = cd.match(/filename\*=UTF-8''([^;]+)/i) || cd.match(/filename="?([^";]+)"?/i);
  const blob = await res.blob(), a = document.createElement("a");
  a.href = URL.createObjectURL(blob); a.download = m ? decodeURIComponent(m[1]) : "download"; document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 4000);
  return res;
}

/* ── 上傳：一個入口，MOMO 下載檔一次好幾個 ───────────────── */
const dz = $("#dz");
dz.addEventListener("click", e => { if (!e.target.closest(".fresh") && e.target.id !== "file") $("#file").click(); });
["dragenter", "dragover"].forEach(ev => dz.addEventListener(ev, e => { e.preventDefault(); dz.classList.add("over"); }));
["dragleave", "drop"].forEach(ev => dz.addEventListener(ev, e => { e.preventDefault(); dz.classList.remove("over"); }));
dz.addEventListener("drop", e => {
  const fs = [...e.dataTransfer.files].filter(f => /\.(xlsx|xls|xlsm)$/i.test(f.name));
  if (fs.length) handleFiles(fs); else toast("請拖曳 Excel 檔（MOMO 下載檔、價格本、瑪氏商品總表）", "err");
});
$("#file").addEventListener("change", () => { const fs = [...$("#file").files]; $("#file").value = ""; if (fs.length) handleFiles(fs); });
$("#fresh").addEventListener("click", e => { e.preventDefault(); e.stopPropagation(); showSettings(true); });

async function handleFiles(files) {
  const fd = new FormData(); files.forEach(f => fd.append("file", f));
  let kinds;
  try { kinds = (await api("/api/momo/kinds", { method: "POST", body: fd })).files; } catch (err) { toast(err.message, "err"); return; }
  const orders = [], refs = [];
  files.forEach((f, i) => {
    const k = kinds[i] && kinds[i].kind;
    if (k === "orders") orders.push(f); else if (k === "conv" || k === "mars") refs.push([f, k]);
    else toast(`${f.name}：認不出是哪一種檔案。可以上傳 MOMO 下載的未出貨清單、價格本、瑪氏商品總表。`, "err");
  });
  for (const [f, k] of refs) await uploadRef(f, k);
  if (orders.length) await previewOrders(orders);
  if (refs.length) { loadFresh(); refresh(); }
}
async function uploadRef(f, kind) {
  const fd = new FormData(); fd.append("file", f);
  try {
    const d = await api(kind === "conv" ? "/api/shopee/conv/import" : "/api/mars/products/import", { method: "POST", body: fd });
    toast(kind === "conv" ? Object.entries(d.lines).map(([l, x]) => `${l}價格本已更新：${fmt(x.count)} 筆（原 ${fmt(x.before)} 筆）`).join("\n")
      : `已更新瑪氏商品總表：${fmt(d.codes)} 個料號（${fmt(d.rows)} 列）`);
  } catch (err) { toast(`${f.name}：${err.message}`, "err"); }
}
async function previewOrders(files) {
  const fd = new FormData(); files.forEach(f => fd.append("file", f));
  const box = document.createElement("div"); box.className = "pv"; box.innerHTML = `<div class="kbd"><i class="bi bi-hourglass-split"></i> 讀取 ${files.length} 個檔中…</div>`;
  $("#previews").prepend(box);
  try { renderPreview(box, await api("/api/momo/import/preview", { method: "POST", body: fd })); }
  catch (err) { box.innerHTML = `<div class="alert al-bad">${esc(err.message).replace(/\n/g, "<br>")}</div>`; setTimeout(() => box.remove(), 15000); }
}
function renderPreview(box, d) {
  const c = d.counts, lines = Object.entries(d.lines).map(([l, n]) => `${esc(l)} ${n} 品`).join("、");
  const errs = d.errors || [], warns = d.warnings || [];
  box.innerHTML = `<h3><i class="bi bi-file-earmark-spreadsheet" style="color:var(--s8)"></i> MOMO 下載檔 ${d.filenames.length} 個
      <span class="kbd" style="font-weight:400">${d.filenames.map(esc).join("、")}</span></h3>
    <div class="facts"><span>入庫單 <b>${d.orders.length}</b> 張</span><span>品項 <b>${d.rows_total}</b></span><span>${lines}</span>
      <span>倉別 <b>${d.warehouses.map(esc).join("、")}</b></span><span>到貨日 <b>${d.dates.map(md).join("、")}</b></span></div>
    <div class="facts"><span>新增 <b>${c.new}</b></span><span>有變 <b style="color:${c.changed ? "var(--bad)" : "inherit"}">${c.changed}</b></span>
      <span>移除 <b>${c.remove}</b></span><span>沒有變 <b>${c.same}</b></span>${d.dup ? `<span class="badge b-warn">同一筆在檔案裡出現兩次 ${d.dup} 筆，以後面的為準</span>` : ""}</div>
    ${d.diffs.length ? `<div class="alert al-warn mt-2"><b>MOMO 改過的地方</b>（匯入後這幾張單要先確認差異才能往下做）<br>${d.diffs.slice(0, 12).map(x =>
      `${esc(x.inbound_no)} ${esc(x.factory_code)}：${esc(x.label)} ${esc(x.old)} → <b>${esc(x.new)}</b>`).join("<br>")}${d.diffs.length > 12 ? `<br>…共 ${d.diffs.length} 項` : ""}</div>` : ""}
    ${errs.length ? `<div class="alert al-bad mt-2"><b>${errs.length} 列讀不進來</b>（其他照常匯入）<div class="tbl-wrap mt-1" style="max-height:200px"><table class="t"><tr><th>檔名</th><th>工作表</th><th>列</th><th>倉別</th><th>原因</th></tr>
      ${errs.map(e => `<tr><td>${esc(e.file)}</td><td>${esc(e.sheet)}</td><td>${esc(e.row)}</td><td>${esc(e.warehouse)}</td><td>${esc(e.reason)}</td></tr>`).join("")}</table></div></div>` : ""}
    ${warns.length ? `<div class="alert al-warn mt-2">${warns.map(esc).join("<br>")}</div>` : ""}
    <div class="flex gap-2 mt-3"><button class="btn btn-p" data-a="ok"><i class="bi bi-check2"></i> 確認匯入</button><button class="btn btn-o" data-a="no">取消</button></div>`;
  box.querySelector("[data-a=ok]").addEventListener("click", async e => {
    e.target.disabled = true;
    try { const r = await post("/api/momo/import/commit", { upload_id: d.upload_id }); box.remove();
      toast(`已匯入：新增 ${r.counts.new}、有變 ${r.counts.changed}、移除 ${r.counts.remove}、沒有變 ${r.counts.same}`); loadFresh(); refresh(); }
    catch (err) { e.target.disabled = false; toast(err.message, "err"); }
  });
  box.querySelector("[data-a=no]").addEventListener("click", async () => { try { await post("/api/momo/import/cancel", { upload_id: d.upload_id }); } catch (e) {} box.remove(); });
}
async function loadPending() {
  try { const d = await api("/api/momo/import/pending");
    d.pending.forEach(p => { const box = document.createElement("div"); box.className = "pv"; $("#previews").append(box); renderPreview(box, p); }); } catch (e) {}
}
async function loadFresh() {
  let d; try { d = await api("/api/momo/fresh"); } catch (e) { return; }
  const ok = (label, t, title) => `<span title="${esc(title || "")}"><i class="bi bi-check-circle" style="color:var(--ok)"></i> ${label} ${when(t)}</span>`;
  const miss = label => `<span class="w"><i class="bi bi-exclamation-circle"></i> ${label}未上傳</span>`;
  const chips = [d.orders ? ok("MOMO 下載檔", d.orders.committed_at, d.orders.filenames) : `<span class="w"><i class="bi bi-exclamation-circle"></i> MOMO 下載檔未匯入</span>`];
  // 寶僑不一定有在 MOMO 賣（問題清單第 18 題），沒傳不亮黃燈
  [["紙潔", "紙潔價格本"], ["寶僑", "寶僑價格本"], ["瑪氏", "瑪氏商品總表"]].forEach(([k, l]) => { if (d.conv[k]) chips.push(ok(l, d.conv[k].updated_at, d.conv[k].file)); else if (k !== "寶僑") chips.push(miss(l)); });
  $("#fresh").innerHTML = chips.join(""); $("#conv-st").innerHTML = chips.slice(1).join("");
}

/* ── 篩選：改了馬上出結果 ───────────────────────────────── */
const PICKS = { line: new Set(), wh: new Set() };
let WH_LIST = [];
function pickHtml(key, label, opts) {
  const on = PICKS[key];
  return `<button class="inp" data-pk="${key}" style="cursor:pointer;${on.size ? "border-color:var(--s8);color:var(--s9);font-weight:700" : ""}">${on.size ? [...on].join("、") : label} ▾</button>
    <div class="pick-pop hidden">${opts.map(o => `<label><input type="checkbox" class="ck" value="${esc(o)}" ${on.has(o) ? "checked" : ""}> ${esc(o)}</label>`).join("") || `<span class="kbd">還沒有資料</span>`}</div>`;
}
function drawPicks() {
  $("#pk-line").innerHTML = pickHtml("line", "全部線別", ["紙品", "潔品", "瑪氏", "寶僑"]);
  $("#pk-wh").innerHTML = pickHtml("wh", "全部倉別", WH_LIST);
}
document.addEventListener("click", e => {
  if (e.target.closest(".pick-pop")) return;                 // 在選單裡勾選不要關掉，才能連勾好幾個
  const b = e.target.closest("[data-pk]");
  $$(".pick-pop").forEach(p => { if (!b || p !== b.nextElementSibling) p.classList.add("hidden"); });
  if (b) b.nextElementSibling.classList.toggle("hidden");
});
document.addEventListener("change", e => {
  const pop = e.target.closest(".pick-pop"); if (!pop) return;
  const key = pop.previousElementSibling.dataset.pk;
  e.target.checked ? PICKS[key].add(e.target.value) : PICKS[key].delete(e.target.value);
  const btn = pop.previousElementSibling; btn.textContent = (PICKS[key].size ? [...PICKS[key]].join("、") : key === "line" ? "全部線別" : "全部倉別") + " ▾";
  refresh();
});
let qTimer;
$("#f-q").addEventListener("input", () => { clearTimeout(qTimer); qTimer = setTimeout(refresh, 350); });
["#f-from", "#f-to", "#f-up-from", "#f-up-to"].forEach(s => $(s).addEventListener("change", refresh));
function qs(extra = {}) {
  const p = new URLSearchParams({ from: $("#f-from").value, to: $("#f-to").value, up_from: $("#f-up-from").value, up_to: $("#f-up-to").value,
    line: [...PICKS.line].join(","), warehouse: [...PICKS.wh].join(","), q: $("#f-q").value.trim(), ...extra });
  [...p.keys()].forEach(k => { if (!p.get(k)) p.delete(k); });
  return p.toString();
}

/* ── 作業流程格子 ───────────────────────────────────────── */
const STAGE_TEXT = {
  diff: ["差異待確認", "MOMO 重新下載後內容有變", "點一張單，右側確認差異後才能往下做。"],
  buy: ["待產 EIP 採購單", "新匯入的 MOMO 入庫單", "勾選入庫單產出 EIP 採購單；「需處理」的單點該列，右側會列出原因和修正方式。"],
  eip: ["待回填 EIP 採購單號", "回填後就能下載瑪氏採購單／直送採購單", "上傳 EIP 後回填單號；回填後同一列可以下載瑪氏採購單、直送採購單，寄給採購。"],
  yx: ["待核缺貨與效期", "瑪氏：等勇信配送明細表（下一批上線）", ""],
  notice: ["待畫入庫通知單", "瑪氏：有不出貨的品項（下一批上線）", ""],
  all: ["全部入庫單", "", "所有入庫單；勾選可以批次改出貨方式。"],
};
let STAGE = "", DATA = null, PLAN = {}, SEL = new Set(), SIDE = "";
function drawStages(stages) {
  const n = Object.fromEntries(stages.map(s => [s.key, s.count]));
  if (!STAGE) STAGE = n.diff ? "diff" : n.buy ? "buy" : n.eip ? "eip" : "buy";
  const keys = ["diff", "buy", "eip", "yx", "notice", "all"].filter(k => k !== "diff" || n.diff || STAGE === "diff");
  $("#stages").innerHTML = keys.map((k, i) => {
    const soon = k === "yx" || k === "notice", c = soon ? "—" : n[k] ?? 0;
    const cls = ["st", soon ? "soon" : "", STAGE === k ? "on" : "", !soon && !c ? "zero" : "", k === "diff" && c ? "bad" : "", k === "all" ? "all" : ""].join(" ");
    return (i ? `<i class="bi bi-chevron-right arr"></i>` : "") + `<div class="${cls}" data-st="${k}" title="${soon ? "下一批上線" : ""}"><div class="n">${c}</div>
      <div class="t">${STAGE_TEXT[k][0]}</div><div class="d">${k === "all" ? `其中 ${fmt(n.all - (DATA ? DATA.orders.filter(g => g.todo.length).length : 0))} 張已處理完` : STAGE_TEXT[k][1]}</div></div>`;
  }).join("");
  $$("#stages .st").forEach(el => el.addEventListener("click", () => { if (el.classList.contains("soon")) return; STAGE = el.dataset.st; SEL.clear(); refresh(); }));
}

async function refresh() {
  let d; try { d = await api("/api/momo/orders?" + qs()); } catch (err) { toast(err.message, "err"); return; }
  DATA = d;
  const whs = [...new Set([...d.warehouses, ...PICKS.wh])].sort();
  if (whs.join() !== WH_LIST.join()) { WH_LIST = whs; drawPicks(); }
  drawStages(d.stages);
  $("#st-title").textContent = STAGE_TEXT[STAGE][0]; $("#st-hint").textContent = STAGE_TEXT[STAGE][2];
  $("#pane-orders").classList.toggle("hidden", STAGE === "eip"); $("#pane-eip").classList.toggle("hidden", STAGE !== "eip");
  if (STAGE === "eip") return drawEips();
  if (STAGE === "buy") {
    try { PLAN = Object.fromEntries((await api("/api/momo/purchase/plan?" + qs())).orders.map(g => [g.inbound_no, g])); } catch (err) { PLAN = {}; toast(err.message, "err"); }
  }
  drawOrders();
}

function rowsOf() {
  const all = DATA.orders;
  return STAGE === "all" ? all : all.filter(g => g.todo.includes(STAGE));
}
function drawOrders() {
  const rows = rowsOf(), buy = STAGE === "buy";
  [...SEL].forEach(no => { if (!rows.some(g => g.inbound_no === no)) SEL.delete(no); });
  const ready = rows.filter(g => buy && PLAN[g.inbound_no] && !PLAN[g.inbound_no].blocking).length;
  $("#sum").innerHTML = `<span><b>${rows.length}</b>張入庫單</span>` + (buy ? `<span><b style="color:var(--ok)">${ready}</b>張可產出</span><span><b style="color:var(--warn)">${rows.length - ready}</b>張需先處理</span>` : "")
    + `<span><b>${fmt(rows.reduce((a, g) => a + g.qty, 0))}</b>MOMO 數量</span>`;
  if (!rows.length) {
    const ao = DATA.all_orders;
    $("#tbl").innerHTML = `<tr><td class="kbd" style="padding:24px;text-align:center">${ao.count ? `這一步目前沒有入庫單。${STAGE !== "all" ? "" : "可能被篩選條件擋掉了，到貨日共有 " + md(ao.first) + "～" + md(ao.last) + "。"}` : "還沒有匯入 MOMO 下載檔，請從上方「上傳檔案」上傳。"}</td></tr>`;
  } else {
    const head = `<tr>${buy || STAGE === "all" ? `<th><input type="checkbox" class="ck" id="ck-all"></th>` : ""}<th>MOMO 入庫單號</th><th>線別</th><th>到貨日</th><th>倉別</th><th class="opt">出貨方式</th>
      <th class="num opt">品項</th><th class="num">MOMO 數量</th>${buy ? `<th class="num">箱數</th>` : ""}<th class="opt">上傳日</th><th>${buy ? "檢查結果" : "進度"}</th></tr>`;
    $("#tbl").innerHTML = head + rows.map(g => {
      const p = PLAN[g.inbound_no], blk = buy && p && p.blocking, probs = buy && p ? [...new Set(p.items.flatMap(x => x.problems.filter(t => !t.startsWith("贈品或"))))] : [];
      const ckOk = STAGE === "all" || (p && !p.blocking);
      const res = buy ? (p ? (p.blocking ? `<span class="badge b-warn">${p.blocking} 項需處理</span><div class="prob">${esc(probs[0] || "")}${probs.length > 1 ? ` 等 ${probs.length} 種` : ""}</div>` : `<span class="badge b-ok">可產出</span>`) : "")
        : `${g.diff_pending ? `<span class="badge b-bad">差異待確認</span> ` : ""}<span class="badge ${g.po_status === "已回填" ? "b-ok" : g.po_status === "已產出" ? "b-s" : "b-dim"}">${g.po_status === "不需產出" ? "不用下採" : "EIP " + g.po_status}</span>${g.eip_pos.length ? ` <span class="fn">${g.eip_pos.map(esc).join("、")}</span>` : ""}`;
      return `<tr class="og ${blk ? "blk" : ""} ${g.diff_pending ? "pend" : ""} ${SIDE === g.inbound_no ? "sel" : ""}" data-no="${esc(g.inbound_no)}">
        ${buy || STAGE === "all" ? `<td onclick="event.stopPropagation()"><input type="checkbox" class="ck rk" value="${esc(g.inbound_no)}" ${SEL.has(g.inbound_no) ? "checked" : ""} ${ckOk ? "" : "disabled"}></td>` : ""}
        <td class="nw"><b class="fn" style="font-size:12.5px;color:var(--ink)">${esc(g.inbound_no)}</b></td>
        <td>${g.lines.map(l => `<span class="badge ${l === "認不出" ? "b-bad" : "b-s"}">${esc(l)}</span>`).join(" ")}</td>
        <td class="nw">${mdw(g.expected_date)}</td><td>${esc(g.warehouse)}${g.wh_known ? "" : ` <span class="badge b-warn" title="倉別主檔沒有這個倉，採購單的地址、電話會空白">主檔沒有</span>`}</td>
        <td class="opt nw">${g.ship_method === "竹運出貨" ? `<span class="badge b-blue">竹運出貨</span>` : esc(g.ship_method)}</td>
        <td class="num opt">${g.item_count}${g.removed_count ? `<span class="kbd">（移除 ${g.removed_count}）</span>` : ""}</td><td class="num">${fmt(g.qty)}</td>
        ${buy ? `<td class="num">${p && p.boxes != null && !p.blocking ? fmt(p.boxes) : "—"}</td>` : ""}<td class="opt kbd">${md(g.upload_date)}</td><td>${res}</td></tr>`;
    }).join("");
    $$("#tbl tr.og").forEach(tr => tr.addEventListener("click", () => openSide(tr.dataset.no)));
    $$("#tbl .rk").forEach(ck => ck.addEventListener("change", () => { ck.checked ? SEL.add(ck.value) : SEL.delete(ck.value); drawBar(); }));
    const all = $("#ck-all");
    if (all) all.addEventListener("change", () => { $$("#tbl .rk:not(:disabled)").forEach(ck => { ck.checked = all.checked; all.checked ? SEL.add(ck.value) : SEL.delete(ck.value); }); drawBar(); });
  }
  $("#buy-go").classList.toggle("hidden", !buy);
  drawBar();
}
function drawBar() {
  const n = SEL.size;
  $("#b-go").disabled = !n; $("#b-go-hint").textContent = n ? `已勾選 ${n} 張入庫單。` : "請勾選要產出的入庫單（「需處理」的單要先處理才能勾）。";
  const bar = $("#a-bar");
  bar.classList.toggle("hidden", !n); bar.className = n ? "abar" : "hidden";
  bar.innerHTML = n ? `已勾選 <b>${n}</b> 張　出貨方式改成 ${SHIP_METHODS.map(m => `<button class="btn btn-o btn-sm" data-sm="${esc(m)}">${esc(m)}</button>`).join(" ")}
    <span class="kbd">竹運出貨的單不用產 EIP（竹運出貨第二批上線）</span>` : "";
  $$("#a-bar [data-sm]").forEach(b => b.addEventListener("click", () => setShip([...SEL], b.dataset.sm)));
}
async function setShip(nos, method) {
  try { const r = await post("/api/momo/ship_method", { inbound_nos: nos, ship_method: method });
    toast(`已改 ${r.count} 張為${method}` + (r.blocked.length ? `\n${r.blocked.join("、")} 已產出 EIP，請先取消產出再改` : "")); SEL.clear(); refresh(); if (SIDE) openSide(SIDE); }
  catch (err) { toast(err.message, "err"); }
}
$("#b-go").addEventListener("click", async () => {
  const btn = $("#b-go"); btn.disabled = true;
  try { const res = await download("/api/momo/purchase/generate", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ inbound_nos: [...SEL] }) });
    toast(`已產出 ${res.headers.get("X-Momo-Files")} 份 EIP 採購單（壓縮檔）。上傳 EIP 後，到「待回填 EIP 採購單號」回填單號。`); SEL.clear(); refresh(); }
  catch (err) { toast(err.message, "err"); btn.disabled = false; }
});

/* ── 待回填 EIP 採購單號：一份 EIP 一列 ─────────────────────── */
let SHOW_FILLED = false;
$("#e-all").addEventListener("click", () => { SHOW_FILLED = !SHOW_FILLED; $("#e-all").classList.toggle("on", SHOW_FILLED); drawEips(); });
$("#e-mars").addEventListener("click", () => download(`/api/momo/po.zip?kind=mars&${qs()}`).catch(err => toast(err.message, "err")));
$("#e-direct").addEventListener("click", () => download(`/api/momo/po.zip?kind=direct&${qs()}`).catch(err => toast(err.message, "err")));
function eipActs(e) {
  return `<input class="inp fn eip-in" data-id="${e.id}" value="${esc(e.eip_po)}" placeholder="PO2026…" style="width:130px"> <button class="btn btn-o btn-sm eip-save" data-id="${e.id}">${e.eip_po ? "更新" : "回填"}</button>
    <a class="lnk kbd eip-dl" data-u="/api/momo/eips/${e.id}/file">EIP</a>
    ${e.eip_po ? `<a class="lnk kbd eip-dl" data-u="/api/momo/eips/${e.id}/${e.mars ? "mars_po" : "direct_po"}"><b>${e.mars ? "瑪氏採購單" : "直送採購單"}</b></a>`
      : `<a class="lnk kbd eip-cancel" data-id="${e.id}" title="這份 EIP 不要了：刪掉，入庫單退回待產 EIP 採購單">取消產出</a>`}`;
}
function bindEipActs(root, after) {
  root.querySelectorAll(".eip-save").forEach(b => b.addEventListener("click", async () => {
    const v = root.querySelector(`.eip-in[data-id="${b.dataset.id}"]`).value.trim();
    try { await post(`/api/momo/eips/${b.dataset.id}`, { eip_po: v }, "PUT"); toast(v ? `已回填 ${v.toUpperCase()}` : "已清除 EIP 採購單號"); after(); refresh(); }
    catch (err) { toast(err.message, "err"); }
  }));
  root.querySelectorAll(".eip-dl").forEach(a => a.addEventListener("click", () => download(a.dataset.u).catch(err => toast(err.message, "err"))));
  root.querySelectorAll(".eip-cancel").forEach(a => a.addEventListener("click", async () => {
    if (!confirm("取消產出這份 EIP 採購單？裡面的入庫單會退回「待產 EIP 採購單」。MOMO 後台已經上傳的 EIP 要自己撤。")) return;
    try { const r = await api(`/api/momo/eips/${a.dataset.id}`, { method: "DELETE" }); toast(`已取消產出，${r.orders.join("、")} 退回待產 EIP 採購單`); after(); refresh(); }
    catch (err) { toast(err.message, "err"); }
  }));
}
async function drawEips() {
  let d; try { d = await api(`/api/momo/eips?${qs(SHOW_FILLED ? {} : { open: "1" })}`); } catch (err) { toast(err.message, "err"); return; }
  $("#e-tbl").innerHTML = d.eips.length ? `<tr><th>檔名</th><th>線別</th><th>倉別</th><th>到貨日</th><th class="num">列</th><th class="num">箱數</th><th>入庫單</th><th>EIP 採購單號／下載</th></tr>`
    + d.eips.map(e => `<tr><td class="fn" style="max-width:360px;white-space:normal">${esc(e.filename)}</td><td><span class="badge b-s">${esc(e.line)}</span></td><td>${esc(e.warehouse)}</td>
      <td class="nw">${mdw(e.expected_date)}</td><td class="num">${e.rows}</td><td class="num">${fmt(e.boxes)}</td>
      <td class="fn">${e.orders.map(no => `<a class="lnk" data-open="${esc(no)}">${esc(no)}</a>`).join("<br>")}</td><td class="nw">${eipActs(e)}</td></tr>`).join("")
    : `<tr><td class="kbd" style="padding:24px;text-align:center">${SHOW_FILLED ? "這段到貨日沒有產出過的 EIP 採購單。" : "沒有待回填的 EIP 採購單。按「顯示已回填」可以看回填過的、下載採購單。"}</td></tr>`;
  bindEipActs($("#e-tbl"), () => {});
  $$("#e-tbl [data-open]").forEach(a => a.addEventListener("click", () => openSide(a.dataset.open)));
}

/* ── 右側面板：一張入庫單的全部 ─────────────────────────── */
$("#sd-close").addEventListener("click", closeSide);
document.addEventListener("keydown", e => { if (e.key === "Escape" && SIDE && !document.querySelector("dialog[open]")) closeSide(); });
function closeSide() { SIDE = ""; $("#side").classList.add("hidden"); document.body.classList.remove("side-open"); $$("#tbl tr.og").forEach(t => t.classList.remove("sel")); }
async function openSide(no) {
  SIDE = no; $$("#tbl tr.og").forEach(t => t.classList.toggle("sel", t.dataset.no === no));
  $("#side").classList.remove("hidden"); document.body.classList.add("side-open");
  let d; try { d = await api("/api/momo/order?no=" + encodeURIComponent(no)); } catch (err) { $("#sd-body").innerHTML = `<div class="alert al-bad">${esc(err.message)}</div>`; return; }
  if (SIDE !== no) return;
  const g = d.order, plan = d.plan;
  $("#sd-title").innerHTML = `<div class="fn" style="font-size:15px;color:var(--ink);font-weight:700">${esc(g.inbound_no)}</div>
    <div class="kbd" style="font-size:12.5px">${esc(g.warehouse)}・${mdw(g.expected_date)} 到貨・${g.lines.map(esc).join("、")}・${g.item_count} 品・MOMO 數量 ${fmt(g.qty)}</div>`;
  const open = d.diffs.filter(x => !x.confirmed);
  const items = plan ? plan.items : [];
  const sec = (title, body, extra = "") => `<div class="sd-sec"><div class="h">${title}${extra}</div><div class="b">${body}</div></div>`;
  let html = sec("出貨方式", `<div class="flex gap-2 flex-wrap">${SHIP_METHODS.map(m => `<span class="chip ${g.ship_method === m ? "on" : ""}" data-ship="${esc(m)}">${esc(m)}</span>`).join("")}</div>
    <div class="kbd mt-1">預設原廠直出；竹運出貨的單不用產 EIP（竹運出貨第二批上線）。已產出 EIP 的要先取消產出才能改。</div>`);
  if (open.length) html += sec(`<i class="bi bi-arrow-left-right" style="color:var(--bad)"></i> MOMO 改過（${open.length} 項）`,
    open.map(x => `<div class="sd-it">${esc(x.factory_code)} ${esc((x.product_name || "").slice(0, 26))}<div class="sd-kv">${esc(x.field_label)}：${esc(x.old_value)} → <b>${esc(x.new_value)}</b>　<span>${esc(x.uploaded_by)} ${when(x.uploaded_at)}</span></div></div>`).join("")
    + `<button class="btn btn-p btn-sm mt-2" id="sd-diff-ok"><i class="bi bi-check2-circle"></i> 確認差異</button><span class="kbd"> 確認後才能產 EIP 採購單</span>`);
  html += sec(`品項與下採（${items.length} 品）`, items.length ? items.map(x => `<div class="sd-it">
      <div class="flex gap-2 items-baseline"><span class="fn">${esc(x.factory_code)}</span><span style="flex:1">${esc(x.name)}</span></div>
      <div class="sd-kv"><span>MOMO <b>${fmt(x.qty)} ${esc(x.unit)}</b></span><span>箱入數 <b>${x.conv ?? "—"}</b></span>
        <span>下採 <b>${x.skip ? "不下採（0）" : x.out_qty != null ? fmt(x.out_qty) + " 箱" : "—"}</b></span>${x.unit_group ? `<span class="kbd">${esc(x.unit_group.replaceAll("_", "・"))}</span>` : ""}
        ${x.eip ? `<span class="badge ${x.eip.eip_po ? "b-ok" : "b-s"}">${x.eip.eip_po ? esc(x.eip.eip_po) : "已產出"}</span>` : ""}</div>
      ${x.adj ? `<div class="kbd">已改下採數量 ${fmt(x.adj.qty_before)} → <b>${fmt(x.adj.qty_after)}</b>：${esc(x.adj.reason)}（${esc(x.adj.operator)} ${when(x.adj.updated_at)}）</div>` : ""}
      ${x.problems.map(t => `<div class="prob">${esc(t)}</div>`).join("")}
      ${!(x.eip && x.eip.eip_po) ? `<div class="mt-1"><button class="btn btn-o btn-sm" data-adj='${esc(JSON.stringify({ id: x.order_id, qty: x.qty, conv: x.conv, unit: x.unit, code: x.factory_code, name: x.name, adj: x.adj, why: x.problems.join("；") }))}'><i class="bi bi-calculator"></i> 改數量</button></div>` : ""}
    </div>`).join("") : `<div class="kbd">${g.ship_method === "竹運出貨" ? "這張單是竹運出貨，不用下採。" : "沒有要下採的品項。"}</div>`);
  if (d.eips.length) html += sec(`EIP 採購單（${d.eips.length} 份）`, d.eips.map(e => `<div class="sd-it"><div class="fn" style="white-space:normal">${esc(e.filename)}</div>
      <div class="sd-kv"><span>${e.rows} 列・${fmt(e.boxes)} 箱</span><span>${esc(e.created_by)} ${when(e.created_at)} 產出</span></div><div class="mt-1 nw">${eipActs(e)}</div></div>`).join(""));
  html += sec("倉別資料", d.warehouse ? `<div class="kbd" style="white-space:pre-line;color:var(--ink2)">${esc(d.warehouse.note)}</div>`
    : `<div class="alert al-warn">倉別主檔沒有「${esc(g.warehouse)}」，瑪氏採購單、直送採購單的地址、電話會空白。請告訴系統管理員加進主檔。</div>`);
  html += `<details class="sd-sec"><summary class="h">歷程（${d.logs.length}）</summary><div class="b">${d.logs.map(l => `<div class="sd-it"><b>${esc(l.field_label)}</b>　${esc(l.old_value)}${l.old_value ? " → " : ""}${esc(l.new_value)}
    <div class="kbd">${esc(l.operator)} ${when(l.changed_at)}${l.note ? "・" + esc(l.note).slice(0, 120) : ""}</div></div>`).join("") || `<div class="kbd">沒有紀錄</div>`}</div></details>`;
  $("#sd-body").innerHTML = html;
  $$("#sd-body [data-ship]").forEach(c => c.addEventListener("click", () => { if (!c.classList.contains("on")) setShip([no], c.dataset.ship); }));
  const ok = $("#sd-diff-ok");
  if (ok) ok.addEventListener("click", async () => { try { await post("/api/momo/diffs/confirm", { inbound_nos: [no] }); toast("已確認差異"); refresh(); openSide(no); } catch (err) { toast(err.message, "err"); } });
  $$("#sd-body [data-adj]").forEach(b => b.addEventListener("click", () => openAdj(JSON.parse(b.dataset.adj))));
  bindEipActs($("#sd-body"), () => openSide(no));
}

/* ── 改下採數量 ─────────────────────────────────────────── */
let ADJ = null;
$$("[data-close]").forEach(b => b.addEventListener("click", () => $("#" + b.dataset.close).close()));
function openAdj(x) {
  ADJ = x;
  $("#da-title").textContent = `修改下採數量　${x.code}`;
  $("#da-why").textContent = x.why || "目前沒有問題；改數量要寫原因，會留紀錄。";
  $("#da-facts").innerHTML = `<span>${esc(x.name)}</span><span>MOMO 數量 <b>${fmt(x.qty)} ${esc(x.unit)}</b></span><span>箱入數 <b>${x.conv ?? "查不到"}</b></span>`;
  const c = x.conv;
  [["#da-down", c ? Math.floor(x.qty / c) * c : null], ["#da-up", c ? Math.ceil(x.qty / c) * c : null]].forEach(([s, v]) => {
    const el = $(s); el.classList.toggle("hidden", !v || v === x.qty); el.textContent = v ? `${fmt(v)}（${v / c} 箱）` : ""; el.onclick = () => { $("#da-qty").value = v; };
  });
  $("#da-zero").onclick = () => { $("#da-qty").value = 0; if (!$("#da-reason").value) $("#da-reason").value = "不下採"; };
  $("#da-qty").value = x.adj ? x.adj.qty_after : ""; $("#da-reason").value = x.adj ? x.adj.reason : "";
  $("#da-who").textContent = x.adj ? `上次修改：${x.adj.operator} ${when(x.adj.updated_at)}` : "";
  $("#da-clear").classList.toggle("hidden", !x.adj);
  $("#dlg-adj").showModal();
}
$("#da-save").addEventListener("click", async () => {
  try { await post("/api/momo/purchase/adjust", { order_id: ADJ.id, qty: $("#da-qty").value.trim(), reason: $("#da-reason").value.trim() });
    $("#dlg-adj").close(); toast("已儲存下採數量"); refresh(); if (SIDE) openSide(SIDE); }
  catch (err) { toast(err.message, "err"); }
});
$("#da-clear").addEventListener("click", async () => {
  try { await post("/api/momo/purchase/adjust", { order_id: ADJ.id, qty: "" }); $("#dlg-adj").close(); toast("已取消修改"); refresh(); if (SIDE) openSide(SIDE); }
  catch (err) { toast(err.message, "err"); }
});

/* ── 資料與設定 ─────────────────────────────────────────── */
$("#btn-set").addEventListener("click", () => showSettings(true));
$("#set-back").addEventListener("click", () => showSettings(false));
function showSettings(on) {
  $("#view-work").classList.toggle("hidden", on); $("#view-set").classList.toggle("hidden", !on);
  if (on) { closeSide(); loadWh(); loadUploads(); }
}
let WH_LABELS = {};
async function loadWh() {
  let d; try { d = await api("/api/momo/warehouses"); } catch (err) { toast(err.message, "err"); return; }
  WH_LABELS = d.labels;
  const cols = ["address", "hours", "phone", "contact", "ship_to", "mars_name"];
  $("#wh-tbl").innerHTML = `<tr><th>倉別</th>${cols.map(c => `<th>${esc(WH_LABELS[c])}</th>`).join("")}</tr>` + d.warehouses.map(w => `<tr data-wh="${esc(w.name)}">
    <td class="nw"><b>${esc(w.name)}</b>${w.aliases.length ? `<div class="kbd">也認：${w.aliases.map(esc).join("、")}</div>` : ""}${w.edited ? `<div><span class="badge b-s">改過</span></div>` : ""}</td>
    ${cols.map(c => `<td class="wh-edit" data-f="${c}" style="${c === "hours" ? "min-width:260px" : ""}">${esc(w[c]) || `<span class="kbd">（空白）</span>`}</td>`).join("")}</tr>`).join("");
  $$("#wh-tbl td.wh-edit").forEach(td => td.addEventListener("click", () => {
    if (td.querySelector("textarea")) return;
    const name = td.parentElement.dataset.wh, f = td.dataset.f, old = td.querySelector(".kbd") ? "" : td.textContent;
    td.innerHTML = `<textarea class="inp">${esc(old)}</textarea>`; const ta = td.querySelector("textarea"); ta.focus();
    ta.addEventListener("blur", async () => {
      const v = ta.value.trim();
      if (v !== old) { try { await post(`/api/momo/warehouses/${encodeURIComponent(name)}`, { [f]: v }, "PUT"); toast(`已更新 ${name} ${WH_LABELS[f]}`); } catch (err) { toast(err.message, "err"); } }
      loadWh();
    });
  }));
}
async function loadUploads() {
  let d; try { d = await api("/api/momo/uploads"); } catch (err) { return; }
  $("#up-tbl").innerHTML = `<tr><th>匯入時間</th><th>誰</th><th>檔案</th><th class="num">品項</th><th>結果</th></tr>` + d.uploads.map(u => `<tr><td class="nw">${esc((u.committed_at || "").slice(0, 16))}</td><td>${esc(u.uploaded_by)}</td>
    <td>${u.files.map(f => `<a class="lnk fn" href="/api/momo/upload_files/${f.id}">${esc(f.filename)}</a>`).join("<br>")}</td><td class="num">${u.rows_total}</td>
    <td class="kbd">新增 ${u.new_count}、有變 ${u.changed_count}、移除 ${u.removed_count}、沒有變 ${u.same_count}${u.errors ? `、讀不進來 ${u.errors} 列` : ""}</td></tr>`).join("");
}

drawPicks(); loadFresh(); loadPending(); refresh();

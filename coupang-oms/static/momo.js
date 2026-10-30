/* MOMO 訂單管理系統（/momo）——版面照蝦皮特選（Jerry 2026-10-28 看過示意圖「好 開始做」）
   - 最上面一個入口收檔案：MOMO 下載檔可一次選好幾個（合在一起預覽）；價格本、瑪氏商品總表送去原本那支（跟蝦皮特選共用）
   - 作業流程格子：差異待確認 → 待產 EIP 採購單 → 待回填 EIP 採購單號 → 待核缺貨與效期（瑪氏）→（待畫入庫通知單：下一批）→ 全部入庫單
   - 待核缺貨與效期（2026-10-28 示意圖「OK 照這樣做」）：勇信配送明細表 PDF、MOMO 效期匯入檔都丟同一個上傳區
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
  const fs = [...e.dataTransfer.files].filter(f => /\.(xlsx|xls|xlsm|pdf)$/i.test(f.name));
  if (fs.length) handleFiles(fs); else toast("請拖曳 Excel 或 PDF 檔（MOMO 下載檔、勇信配送明細表、效期匯入檔、價格本、瑪氏商品總表）", "err");
});
$("#file").addEventListener("change", () => { const fs = [...$("#file").files]; $("#file").value = ""; if (fs.length) handleFiles(fs); });
$("#fresh").addEventListener("click", e => { e.preventDefault(); e.stopPropagation(); showSettings(true); });

async function handleFiles(files) {
  const fd = new FormData(); files.forEach(f => fd.append("file", f));
  let kinds;
  try { kinds = (await api("/api/momo/kinds", { method: "POST", body: fd })).files; } catch (err) { toast(err.message, "err"); return; }
  const orders = [], refs = [], pdfs = [], exps = [], notices = [];
  files.forEach((f, i) => {
    const k = kinds[i] && kinds[i].kind;
    if (k === "orders") orders.push(f); else if (k === "conv" || k === "mars") refs.push([f, k]);
    else if (k === "yx") pdfs.push(f); else if (k === "notice") notices.push(f); else if (k === "expfill") exps.push(f);
    else toast(`${f.name}：認不出是哪一種檔案。可以上傳 MOMO 下載的未出貨清單、勇信配送明細表、價格本、瑪氏商品總表。`, "err");
  });
  for (const [f, k] of refs) await uploadRef(f, k);
  if (pdfs.length) await uploadYx(pdfs);
  if (notices.length) await uploadNotice(notices);
  exps.forEach(askExpfill);
  if (orders.length) await previewOrders(orders);
  if (refs.length) { loadFresh(); refresh(); }
}
/* 勇信配送明細表：上傳完在上傳區下面列出每張入庫單核對的結果 */
async function uploadYx(files) {
  const fd = new FormData(); files.forEach(f => fd.append("file", f));
  const box = document.createElement("div"); box.className = "pv"; box.innerHTML = `<div class="kbd"><i class="bi bi-hourglass-split"></i> 讀取勇信配送明細表 ${files.length} 份…</div>`;
  $("#previews").prepend(box);
  let d; try { d = await api("/api/momo/yx", { method: "POST", body: fd }); }
  catch (err) { box.innerHTML = `<div class="alert al-bad">${esc(err.message).replace(/\n/g, "<br>")}</div>`; setTimeout(() => box.remove(), 15000); return; }
  const line = o => `<div><b class="fn" style="color:var(--ink)">${esc(o.inbound_no)}</b>（${esc(o.warehouse)}）`
    + [o.noship_bought ? `缺貨 ${o.noship_bought} 品` : "", o.ship ? `可寫入效期 ${o.ship} 品` : "", o.pending ? `<b style="color:var(--warn)">要人決定 ${o.pending} 品</b>` : "",
       o.wait ? `還沒收到勇信表 ${o.wait} 品` : ""].filter(Boolean).join("、") + `</div>`;
  box.innerHTML = `<h3><i class="bi bi-truck" style="color:var(--s8)"></i> 勇信配送明細表 ${files.length} 份、對到 ${d.eip_pos.length} 份 EIP
      <span class="kbd" style="font-weight:400">${files.map(f => esc(f.name)).join("、")}</span><span style="flex:1"></span><button class="btn btn-o btn-sm" data-a="x">關閉</button></h3>
    <div class="mt-2" style="font-size:13px">${d.orders.map(line).join("") || `<span class="kbd">沒有對到 MOMO 的入庫單。</span>`}</div>
    ${d.unknown.length ? `<div class="alert al-warn mt-2">這幾個收貨單號不是 MOMO 瑪氏 EIP（可能是蝦皮、酷澎的，或 EIP 單號還沒回填）：${d.unknown.map(esc).join("、")}</div>` : ""}
    ${d.warnings.length ? `<div class="alert al-warn mt-2">${d.warnings.map(esc).join("<br>")}</div>` : ""}`;
  box.querySelector("[data-a=x]").addEventListener("click", () => box.remove());
  if (d.orders.length && STAGE !== "yx") STAGE = "yx";
  refresh(); if (SIDE) openSide(SIDE);
}
/* ④ MOMO 入庫通知單 PDF：收進來就能下載畫好的 */
async function uploadNotice(files) {
  const fd = new FormData(); files.forEach(f => fd.append("file", f));
  const box = document.createElement("div"); box.className = "pv"; box.innerHTML = `<div class="kbd"><i class="bi bi-hourglass-split"></i> 讀取入庫通知單 ${files.length} 份…</div>`;
  $("#previews").prepend(box);
  let d; try { d = await api("/api/momo/notice", { method: "POST", body: fd }); }
  catch (err) { box.innerHTML = `<div class="alert al-bad">${esc(err.message).replace(/\n/g, "<br>")}</div>`; setTimeout(() => box.remove(), 15000); return; }
  const line = o => o.kind === "mark" ? `<div class="flex items-center gap-2 flex-wrap" style="margin:3px 0"><span class="badge b-s">嘜頭</span><b class="fn" style="color:var(--ink)">${esc(o.inbound_no)}</b>（${esc(o.warehouse)}）
      ${o.rows} 頁：填共幾箱 ${o.fill} 頁${o.drop ? `、不出貨拿掉 ${o.drop} 頁` : ""}
      ${o.problems.length ? `<span class="badge b-warn">還有 ${o.problems.length} 品要先決定</span>` : ""}${o.unmatched.length ? `<span class="badge b-bad">系統找不到 ${o.unmatched.map(esc).join("、")}（照原樣留著）</span>` : ""}
      ${o.ready ? `<a class="btn btn-p btn-sm" href="/api/momo/mark/${encodeURIComponent(o.inbound_no)}/file"><i class="bi bi-download"></i> 下載填好的嘜頭</a>` : ""}</div>`
    : `<div class="flex items-center gap-2 flex-wrap" style="margin:3px 0"><span class="badge b-s">入庫通知單</span><b class="fn" style="color:var(--ink)">${esc(o.inbound_no)}</b>（${esc(o.warehouse)}）
      ${o.need ? `要畫：${esc(o.marks)}` : `<span class="kbd">全部照訂單出，沒有要畫的</span>`}
      ${o.problems.length ? `<span class="badge b-warn">還有 ${o.problems.length} 品要先決定</span>` : ""}${o.unmatched.length ? `<span class="badge b-bad">PDF 上找不到 ${o.unmatched.map(esc).join("、")}</span>` : ""}
      ${o.ready && o.need ? `<a class="btn btn-p btn-sm" href="/api/momo/notice/${encodeURIComponent(o.inbound_no)}/file"><i class="bi bi-download"></i> 下載畫好的</a>` : ""}</div>`;
  box.innerHTML = `<h3><i class="bi bi-file-earmark-pdf" style="color:var(--s8)"></i> 入庫通知單／嘜頭 ${d.orders.length} 份<span style="flex:1"></span><button class="btn btn-o btn-sm" data-a="x">關閉</button></h3>
    <div class="mt-2" style="font-size:13px">${d.orders.map(line).join("")}</div>
    ${d.errors.length ? `<div class="alert al-warn mt-2">${d.errors.map(esc).join("<br>")}</div>` : ""}
    ${d.orders.some(o => o.warnings.length) ? `<div class="alert al-warn mt-2">${d.orders.flatMap(o => o.warnings).map(esc).join("<br>")}</div>` : ""}`;
  box.querySelector("[data-a=x]").addEventListener("click", () => box.remove());
  box.querySelectorAll("a.btn").forEach(a => a.addEventListener("click", () => setTimeout(() => { refresh(); if (SIDE) openSide(SIDE); }, 1500)));
  if (d.orders.some(o => o.kind === "notice") && STAGE !== "notice") STAGE = "notice";
  refresh(); if (SIDE) openSide(SIDE);
}
/* MOMO 效期匯入檔跟 MOMO 下載檔長得一樣：裡面的單都核過勇信表時，讓人選要填效期還是當訂單匯入 */
function askExpfill(f) {
  const box = document.createElement("div"); box.className = "pv";
  box.innerHTML = `<h3><i class="bi bi-calendar-check" style="color:var(--s8)"></i> ${esc(f.name)}</h3>
    <div class="kbd" style="font-size:12.5px">這份檔裡的入庫單都已經核過勇信配送明細表，應該是 MOMO 的效期匯入檔。</div>
    <div class="flex gap-2 mt-2"><button class="btn btn-p" data-a="fill"><i class="bi bi-download"></i> 填好效期下載</button>
      <button class="btn btn-o" data-a="imp">當作 MOMO 下載檔匯入</button><button class="btn btn-o" data-a="x">取消</button></div>`;
  $("#previews").prepend(box);
  box.querySelector("[data-a=fill]").addEventListener("click", async () => { if (await fillExp(f)) box.remove(); });
  box.querySelector("[data-a=imp]").addEventListener("click", () => { box.remove(); previewOrders([f]); });
  box.querySelector("[data-a=x]").addEventListener("click", () => box.remove());
}
async function fillExp(f) {
  const fd = new FormData(); fd.append("file", f);
  try { const res = await download("/api/momo/expfill", { method: "POST", body: fd });
    toast(`已填好 ${res.headers.get("X-Momo-Filled")} 筆效期，請把下載的檔上傳 MOMO。`); refresh(); if (SIDE) openSide(SIDE); return true; }
  catch (err) { toast(err.message, "err"); return false; }
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
  diff: ["差異待確認", "MOMO 重新下載後內容有變", "點一張單，在視窗裡確認差異後才能往下做。"],
  buy: ["待產 EIP 採購單", "新匯入的 MOMO 入庫單", "勾選入庫單產出 EIP 採購單；「需處理」的單點該列，視窗裡會列出原因和修正方式。"],
  eip: ["待回填 EIP 採購單號", "回填後就能下載瑪氏採購單／直送採購單", "上傳 EIP 後回填單號；回填後同一列可以下載瑪氏採購單、直送採購單，寄給採購。"],
  yx: ["待核缺貨與效期", "瑪氏：勇信表核缺貨、填效期", "瑪氏、已回填 EIP 單號的入庫單。丟勇信配送明細表進來就自動核對；再丟 MOMO 效期匯入檔，下載填好效期的檔上傳 MOMO。"],
  notice: ["待畫入庫通知單", "瑪氏：有不出貨或短出的品項", "效期做完、有不出貨或短出的單。從 MOMO 後台下載入庫通知單 PDF 丟進上傳區，系統畫好線、寫實際出貨數、蓋章，下載後印出來。"],
  all: ["全部入庫單", "", "所有入庫單；勾選可以批次改出貨方式。"],
};
let STAGE = "", DATA = null, PLAN = {}, SEL = new Set(), SIDE = "";
function drawStages(stages) {
  const n = Object.fromEntries(stages.map(s => [s.key, s.count]));
  if (!STAGE) STAGE = n.diff ? "diff" : n.buy ? "buy" : n.eip ? "eip" : n.yx ? "yx" : n.notice ? "notice" : "buy";
  const keys = ["diff", "buy", "eip", "yx", "notice", "all"].filter(k => k !== "diff" || n.diff || STAGE === "diff");
  $("#stages").innerHTML = keys.map((k, i) => {
    const soon = false, c = n[k] ?? 0;
    const cls = ["st", soon ? "soon" : "", STAGE === k ? "on" : "", !c ? "zero" : "", k === "diff" && c ? "bad" : "", k === "all" ? "all" : ""].join(" ");
    return (i ? `<i class="bi bi-chevron-right arr"></i>` : "") + `<div class="${cls}" data-st="${k}"><div class="n">${c}</div>
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
  $("#yx-help").classList.toggle("hidden", STAGE !== "yx"); $("#nt-help").classList.toggle("hidden", STAGE !== "notice");
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
  if (STAGE === "yx") return drawChecks();
  if (STAGE === "notice") return drawNotices();
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
/* ④ 待畫入庫通知單：一張入庫單一列 */
function drawNotices() {
  const rows = rowsOf();
  $("#sum").innerHTML = `<span><b>${rows.length}</b>張入庫單</span><span><b style="color:var(--warn)">${rows.filter(g => !g.check.notice).length}</b>張還沒收到入庫通知單</span>`;
  $("#tbl").innerHTML = !rows.length ? `<tr><td class="kbd" style="padding:24px;text-align:center">這一步目前沒有入庫單。效期做完、有不出貨或短出品項的單會出現在這裡。</td></tr>`
    : `<tr><th>MOMO 入庫單號</th><th>倉別</th><th>到貨日</th><th class="num">品項</th><th>不出貨</th><th>短出</th><th>入庫通知單</th><th></th></tr>` + rows.map(g => {
      const c = g.check, nt = c.notice;
      return `<tr class="og ${SIDE === g.inbound_no ? "sel" : ""}" data-no="${esc(g.inbound_no)}">
        <td class="nw"><b class="fn" style="font-size:12.5px;color:var(--ink)">${esc(g.inbound_no)}</b></td><td>${esc(g.warehouse)}</td><td class="nw">${mdw(g.expected_date)}</td>
        <td class="num">${c.items}</td><td>${c.noship ? `<span class="badge b-bad">${c.noship} 品</span>${c.skip ? `<span class="kbd">（含不下採 ${c.skip}）</span>` : ""}` : `<span class="kbd">—</span>`}</td>
        <td>${c.short ? `<span class="badge b-warn">${c.short} 品</span>` : `<span class="kbd">—</span>`}</td>
        <td>${nt ? `<span class="badge b-ok">已收到 ${when(nt.uploaded_at)}</span>${nt.drawn_at ? `<div class="kbd">已畫 ${when(nt.drawn_at)}</div>` : ""}` : `<span class="badge b-blue">還沒收到</span><div class="kbd">從 MOMO 後台下載後丟上傳區</div>`}</td>
        <td class="nw" onclick="event.stopPropagation()">${nt ? `<a class="btn btn-p btn-sm" href="/api/momo/notice/${encodeURIComponent(g.inbound_no)}/file"><i class="bi bi-download"></i> 下載畫好的</a>` : ""}</td></tr>`;
    }).join("");
  $$("#tbl tr.og").forEach(tr => tr.addEventListener("click", () => openSide(tr.dataset.no)));
  $$("#tbl a.btn").forEach(a => a.addEventListener("click", () => setTimeout(() => { refresh(); if (SIDE) openSide(SIDE); }, 1500)));
  $("#buy-go").classList.add("hidden");
  drawBar();
}
/* ③ 待核缺貨與效期：一張入庫單一列（示意圖的欄位） */
function drawChecks() {
  const rows = rowsOf();
  $("#sum").innerHTML = `<span><b>${rows.length}</b>張入庫單</span><span><b style="color:var(--warn)">${rows.filter(g => g.check.pending).length}</b>張要人決定</span>`
    + `<span><b>${rows.filter(g => g.check.wait === g.check.items).length}</b>張還沒收到勇信表</span>`;
  $("#tbl").innerHTML = !rows.length ? `<tr><td class="kbd" style="padding:24px;text-align:center">這一步目前沒有入庫單。瑪氏的 EIP 單號回填後會出現在這裡。</td></tr>`
    : `<tr><th>MOMO 入庫單號</th><th>倉別</th><th>到貨日</th><th class="num">品項</th><th>勇信配送明細表</th><th>缺貨</th><th>效期</th><th>效期匯入檔</th></tr>` + rows.map(g => {
      const c = g.check, none = c.wait === c.items;
      const yx = c.yx_at ? `<span class="badge b-ok">已核對 ${when(c.yx_at)}</span>` : none ? `<span class="badge b-blue">還沒收到</span>` : `<span class="badge b-warn">還差 ${c.wait} 品</span>`;
      const short = none ? "—" : (c.noship_bought ? `<span class="badge b-bad">${c.noship_bought} 品不出貨</span>` : c.ask_qty ? "" : `<span class="kbd">沒有缺貨</span>`)
        + (c.ask_qty ? ` <span class="badge b-warn">${c.ask_qty} 品數量不對</span><div class="kbd">要人決定</div>` : "") + (c.skip ? `<div class="kbd">另有 ${c.skip} 品不下採</div>` : "");
      const ask = c.pending - c.ask_qty;
      const exp = none ? "—" : `${c.ship} 品可寫入` + (ask > 0 ? `<div class="prob" style="color:var(--warn)">${ask} 品效期要人看</div>` : "");
      const fill = c.filled_at ? `<span class="badge b-ok">已填 ${when(c.filled_at)}</span>` : none ? "—" : c.ship ? `<span class="badge b-warn">還沒填</span>` : `<span class="kbd">不用填</span>`;
      return `<tr class="og ${c.pending ? "ask" : ""} ${SIDE === g.inbound_no ? "sel" : ""}" data-no="${esc(g.inbound_no)}">
        <td class="nw"><b class="fn" style="font-size:12.5px;color:var(--ink)">${esc(g.inbound_no)}</b></td><td>${esc(g.warehouse)}</td><td class="nw">${mdw(g.expected_date)}</td>
        <td class="num">${c.items}</td><td>${yx}</td><td>${short}</td><td>${exp}</td><td>${fill}</td></tr>`;
    }).join("");
  $$("#tbl tr.og").forEach(tr => tr.addEventListener("click", () => openSide(tr.dataset.no)));
  $("#buy-go").classList.add("hidden");
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

/* ── 入庫單視窗：一張單的全部（2026-10-30 從右側面板改成畫面中間的大視窗，跟蝦皮特選一樣；品項一品一列，下採和 ③ 放同一列） ── */
let SD = null, SDQ = "", SDF = "all";
$("#sd-close").addEventListener("click", closeSide); $("#sd-close2").addEventListener("click", closeSide);
$("#side").addEventListener("mousedown", e => { if (e.target.id === "side") closeSide(); });          // 點視窗外面的灰色區域也關
document.addEventListener("keydown", e => {
  if (!SIDE || document.querySelector("dialog[open]")) return;
  if (e.key === "Escape") closeSide();
  else if ((e.key === "ArrowLeft" || e.key === "ArrowRight") && !e.target.closest("input, textarea, select")) { e.preventDefault(); sideStep(e.key === "ArrowLeft" ? -1 : 1); }
});
$("#sd-prev").addEventListener("click", () => sideStep(-1)); $("#sd-next").addEventListener("click", () => sideStep(1));
/* 上一張／下一張：照目前這一步清單上看得到的順序 */
const sideList = () => [...new Set($$("#tbl tr.og[data-no], #e-tbl [data-open]").filter(el => el.offsetParent !== null).map(el => el.dataset.no || el.dataset.open))];
function sideStep(dir) { const l = sideList(), i = l.indexOf(SIDE); if (i < 0 || !l[i + dir]) return; openSide(l[i + dir]); }
function sideNav() {
  const l = sideList(), i = l.indexOf(SIDE);
  $("#sd-prev").disabled = i <= 0; $("#sd-next").disabled = i < 0 || i >= l.length - 1;
  $("#sd-pos").textContent = i >= 0 ? `第 ${i + 1}／${l.length} 張` : "";
}
function closeSide() { SIDE = ""; SD = null; $("#side").classList.add("hidden"); document.body.classList.remove("side-open"); $$("#tbl tr.og").forEach(t => t.classList.remove("sel")); }
async function openSide(no) {
  const again = SIDE === no && SD;
  if (!again) { SDQ = $("#f-q").value.trim(); SDF = "all"; }      // 換一張單：品項搜尋帶上面搜尋框的字，篩選回到全部
  SIDE = no; $$("#tbl tr.og").forEach(t => t.classList.toggle("sel", t.dataset.no === no));
  $("#side").classList.remove("hidden"); document.body.classList.add("side-open"); sideNav();
  if (!again) { $("#sd-title").innerHTML = `<b class="fn" style="font-size:16px;color:var(--ink)">${esc(no)}</b>`; $("#sd-body").innerHTML = `<div class="kbd" style="padding:12px 0">讀取中…</div>`; }
  let d; try { d = await api("/api/momo/order?no=" + encodeURIComponent(no)); }
  catch (err) { if (SIDE === no) $("#sd-body").innerHTML = `<div class="alert al-bad">${esc(err.message)}</div>`; return; }
  if (SIDE !== no) return;                         // 讀的時候又點了別張
  const top = $("#sd-body").scrollTop; SD = d; renderSide(); if (again) $("#sd-body").scrollTop = top;
}
/* 一品一列：下採計畫和 ③ 核對結果用訂單編號接起來 */
function sideItems(d) {
  const ck = new Map(((d.check || {}).items || []).map(x => [x.order_id, x]));
  const rows = ((d.plan || {}).items || []).map(p => ({ p, c: ck.get(p.order_id) }));
  const seen = new Set(rows.map(r => r.p.order_id));
  ck.forEach((c, id) => { if (!seen.has(id)) rows.push({ p: null, c }); });
  return rows;
}
const isBadRow = r => !!((r.p && r.p.blocking) || (r.c && r.c.decision === "pending"));
const isCutRow = r => !!(r.c && (r.c.decision === "noship" || (r.c.decision === "ship" && r.c.actual != null && r.c.actual !== (r.c.qty || 0))));
function renderSide() {
  const d = SD, g = d.order, no = g.inbound_no, ck = d.check, rows = sideItems(d);
  const fact = (k, v) => `<div><div class="k">${k}</div><div class="v">${v}</div></div>`;
  const stg = (g.todo || [])[0];
  $("#sd-title").innerHTML = `<div class="flex items-baseline gap-3 flex-wrap"><b class="fn" style="font-size:17px;color:var(--ink)">${esc(no)}</b>
      <span class="kbd">${esc(g.warehouse)}・${mdw(g.expected_date)} 到貨・${g.lines.map(esc).join("、")}</span>
      ${stg && STAGE_TEXT[stg] ? `<span class="badge b-warn">${esc(STAGE_TEXT[stg][0])}</span>` : `<span class="badge b-ok">已處理完</span>`}</div>
    <div class="sd-facts">${fact("到貨日", mdw(g.expected_date))}${fact("倉別", esc(g.warehouse))}${fact("品項", `${g.item_count} 品`)}${fact("MOMO 數量", fmt(g.qty))}
      ${fact("出貨方式", `<span class="flex gap-1 flex-wrap" style="font-size:12.5px">${SHIP_METHODS.map(m => `<span class="chip ${g.ship_method === m ? "on" : ""}" style="padding:1px 8px" data-ship="${esc(m)}">${esc(m)}</span>`).join("")}</span>`)}</div>`;
  const open = d.diffs.filter(x => !x.confirmed);
  const sec = (title, body, extra = "") => `<div class="sd-sec"><div class="h">${title}${extra}</div><div class="b">${body}</div></div>`;
  let html = "";
  if (open.length) html += sec(`<i class="bi bi-arrow-left-right" style="color:var(--bad)"></i> MOMO 改過（${open.length} 項）`,
    open.map(x => `<div class="sd-it">${esc(x.factory_code)} ${esc((x.product_name || "").slice(0, 40))}<div class="sd-kv">${esc(x.field_label)}：${esc(x.old_value)} → <b>${esc(x.new_value)}</b>　<span>${esc(x.uploaded_by)} ${when(x.uploaded_at)}</span></div></div>`).join("")
    + `<button class="btn btn-p btn-sm mt-2" id="sd-diff-ok"><i class="bi bi-check2-circle"></i> 確認差異</button><span class="kbd"> 確認後才能產 EIP 採購單</span>`);
  const nBad = rows.filter(isBadRow).length, nCut = rows.filter(isCutRow).length;
  html += sec(`品項與下採（${rows.length} 品）${nBad ? `<span class="badge b-bad">${nBad} 項需處理，排在最前面</span>` : ""}`, rows.length ? `
    <div class="it-tools"><input class="inp" id="sd-q" placeholder="搜尋這張單的原廠編號、品名" value="${esc(SDQ)}" style="width:280px">
      <span class="chip sd-fl ${SDF === "all" ? "on" : ""}" data-f="all">全部 ${rows.length}</span>${nBad ? `<span class="chip sd-fl ${SDF === "bad" ? "on" : ""}" data-f="bad">需處理 ${nBad}</span>` : ""}${nCut
      ? `<span class="chip sd-fl ${SDF === "cut" ? "on" : ""}" data-f="cut">不出貨／短出 ${nCut}</span>` : ""}<span class="kbd" id="sd-found"></span></div>
    <div class="tbl-wrap" style="max-height:none"><table class="t it-t" id="sd-items"></table></div>
    ${ck ? `<div class="kbd mt-1">不出貨的可以改成出貨（例如採購說會補出）；改了會記誰改的。</div>` : ""}`
    : `<div class="kbd">${g.ship_method === "竹運出貨" ? "這張單是竹運出貨，不用下採。" : "沒有要下採的品項。"}</div>`);
  html += `<div class="sd-grid">`;
  if (ck) html += checkSec(ck, sec);
  if (d.eips.length) html += sec(`EIP 採購單（${d.eips.length} 份）`, d.eips.map(e => `<div class="sd-it"><div class="fn" style="white-space:normal">${esc(e.filename)}</div>
      <div class="sd-kv"><span>${e.rows} 列・${fmt(e.boxes)} 箱</span><span>${esc(e.created_by)} ${when(e.created_at)} 產出</span></div><div class="mt-1 nw">${eipActs(e)}</div></div>`).join(""));
  html += sec("倉別資料", d.warehouse ? `<div class="kbd" style="white-space:pre-line;color:var(--ink2)">${esc(d.warehouse.note)}</div>`
    : `<div class="alert al-warn">倉別主檔沒有「${esc(g.warehouse)}」，瑪氏採購單、直送採購單的地址、電話會空白。請到「資料與設定」的倉別主檔加進來。</div>`);
  html += `<details class="sd-sec"><summary class="h">歷程（${d.logs.length}）</summary><div class="b">${d.logs.map(l => `<div class="sd-it"><b>${esc(l.field_label)}</b>　${esc(l.old_value)}${l.old_value ? " → " : ""}${esc(l.new_value)}
    <div class="kbd">${esc(l.operator)} ${when(l.changed_at)}${l.note ? "・" + esc(l.note).slice(0, 120) : ""}</div></div>`).join("") || `<div class="kbd">沒有紀錄</div>`}</div></details></div>`;
  $("#sd-body").innerHTML = html;
  $$("#sd-title [data-ship]").forEach(c => c.addEventListener("click", () => { if (!c.classList.contains("on")) setShip([no], c.dataset.ship); }));
  const ok = $("#sd-diff-ok");
  if (ok) ok.addEventListener("click", async () => { try { await post("/api/momo/diffs/confirm", { inbound_nos: [no] }); toast("已確認差異"); refresh(); openSide(no); } catch (err) { toast(err.message, "err"); } });
  if (rows.length) {
    $("#sd-q").addEventListener("input", () => { SDQ = $("#sd-q").value.trim(); drawItems(); });
    $$("#side .sd-fl").forEach(c => c.addEventListener("click", () => { SDF = c.dataset.f; $$("#side .sd-fl").forEach(x => x.classList.toggle("on", x === c)); drawItems(); }));
    drawItems();
  }
  bindEipActs($("#sd-body"), () => openSide(no));
  bindCheck(no);
}

/* 品項表：一品一列。左半是下採（MOMO 數量、箱入數、下採箱數、改數量），有勇信表後右半是 ③（勇信出／效期、MOMO 最低、結果） */
const RES = { ship: ["b-ok", "寫入效期"], noship: ["b-bad", "不出貨"], pending: ["b-warn", "要決定"], wait: ["b-dim", "等勇信表"] };
const slash = t => (t || "").replaceAll("-", "/");
function drawItems() {
  const d = SD; if (!d || !$("#sd-items")) return;
  const ck = d.check, q = SDQ.toLowerCase();
  const isHit = r => !!q && [r.p && r.p.factory_code, r.p && r.p.name, r.c && r.c.factory_code, r.c && r.c.name].some(v => String(v || "").toLowerCase().includes(q));
  let list = sideItems(d).map((r, k) => [r, k]).sort((a, b) => (isBadRow(a[0]) ? 0 : 1) - (isBadRow(b[0]) ? 0 : 1) || a[1] - b[1]).map(x => x[0]);
  if (q) list = list.filter(isHit);
  if (SDF === "bad") list = list.filter(isBadRow); else if (SDF === "cut") list = list.filter(isCutRow);
  $("#sd-found").textContent = q ? `　找到 ${list.length} 品（符合「${SDQ}」）` : SDF !== "all" ? `　顯示 ${list.length} 品` : "";
  let h = `<tr><th>商品</th><th>MOMO 數量</th><th>箱入數</th><th>下採</th>${ck ? `<th>勇信出／效期</th><th>MOMO 最低</th><th>③ 結果</th>` : ""}</tr>`;
  h += list.map(r => {
    const p = r.p, x = r.c, b = p || x;
    const unit = b.unit, qty = p ? p.qty : x.qty;
    let left = `<td class="nm2"><div style="font-weight:600">${esc(b.name)}</div><div class="fn kbd">${esc(b.factory_code)}${p && p.unit_group ? `・${esc(p.unit_group.replaceAll("_", "・"))}` : ""}</div>
        ${p && p.adj ? `<div class="kbd">已改下採數量 ${fmt(p.adj.qty_before)} → <b>${fmt(p.adj.qty_after)}</b>：${esc(p.adj.reason)}（${esc(p.adj.operator)} ${when(p.adj.updated_at)}）</div>` : ""}
        ${p && p.note ? `<div class="kbd" style="color:var(--warn)">${esc(p.note)}</div>` : ""}
        ${p ? p.problems.map(t => `<div class="prob">${esc(t)}</div>`).join("") : ""}</td>
      <td class="xs nw"><span class="big">${fmt(qty)}</span> ${esc(unit)}</td><td class="nw">${p ? (p.conv ?? "—") : (x.conv ?? "—")}</td>
      <td class="xs" style="min-width:110px"><b class="nw">${p ? (p.skip ? "不下採（0）" : p.out_qty != null ? fmt(p.out_qty) + " 箱" : "—") : x.skip ? "不下採" : fmt(x.ordered) + " 箱"}</b>
        ${p && p.eip ? `<div><span class="badge ${p.eip.eip_po ? "b-ok" : "b-s"}">${p.eip.eip_po ? esc(p.eip.eip_po) : "已產出"}</span></div>` : ""}
        ${p && !(p.eip && p.eip.eip_po) ? `<div class="mt-1"><button class="btn btn-o btn-sm" data-adj='${esc(JSON.stringify({ id: p.order_id, qty: p.qty, conv: p.conv, unit: p.unit, code: p.factory_code, name: p.name, adj: p.adj, why: p.problems.join("；") }))}'><i class="bi bi-calculator"></i> 改數量</button></div>` : ""}</td>`;
    let right = "";
    if (ck) right = x ? checkCells(x) : `<td colspan="3" class="kbd">—</td>`;
    const cls = ["sd-it", p && p.blocking ? "bad" : "", x && x.decision === "pending" ? "ask" : "", x && x.decision === "noship" ? "x" : "", isHit(r) ? "hit" : ""].join(" ");
    return `<tr class="${cls}">${left}${right}</tr>`;
  }).join("");
  if (!list.length) h += `<tr><td colspan="7" class="kbd" style="padding:18px;text-align:center">${q ? `這張單沒有符合「${esc(SDQ)}」的品項。` : "沒有符合的品項。"}</td></tr>`;
  $("#sd-items").innerHTML = h;
  bindRows(SIDE);
}
/* ③ 那三格：勇信出幾箱、效期、結果；不出貨可以改、效期可以換 */
function checkCells(x) {
  const r = RES[x.decision], x_ = x.decision === "noship", can = x.status !== "wait" && x.status !== "skip", ov = x.override;
  const pick = can && !x_ && x.batches.length > 1;          // 兩批效期：點一下就選那批
  const bt = x.batches.length ? x.batches.map(b => `<div class="nw">${x.batches.length > 1 ? `${b.qty} 箱 ` : ""}${pick && b.expiry ? `<a class="lnk" data-pick="${x.order_id}" data-v="${esc(b.expiry)}" title="用這批的效期">${slash(b.expiry)}</a>` : slash(b.expiry) || "沒印效期"}</div>`).join("") : "—";
  return `<td><b class="nw">${x.shipped == null ? "—" : fmt(x.shipped) + " 箱"}</b>${x.shipped ? `<div class="kbd">${bt}</div>` : ""}</td><td class="xs nw">${slash(x.min_expiry) || "—"}</td>
    <td style="min-width:150px"><span class="badge ${r[0]}">${r[1]}</span>${x.decision === "ship" && x.expiry ? ` <span class="nw">${slash(x.expiry)}</span>` : ""}
      ${x.why && x.decision !== "ship" ? `<div class="kbd" style="color:${x.decision === "pending" ? "var(--warn)" : ""}">${esc(x.why)}</div>` : ""}
      ${can ? `<div class="flex gap-1 mt-1 flex-wrap"><span class="chip ${x.decision === "ship" ? "on" : ""}" style="padding:1px 8px;font-size:12px" data-ship1="${x.order_id}">出貨</span>
        <span class="chip ${x_ ? "on" : ""}" style="padding:1px 8px;font-size:12px" data-ship0="${x.order_id}">不出貨</span></div>
        ${!x_ && x.decision !== "pending" && !(ov && ov.expiry) ? `<a class="lnk kbd" data-showexp="${x.order_id}">換效期</a>` : ""}
        ${!x_ ? `<input type="date" class="inp mt-1 ${x.decision === "pending" || (ov && ov.expiry) ? "" : "hidden"}" style="padding:2px 4px;font-size:12px;width:130px" data-exp="${x.order_id}" data-ov="${ov && ov.ship != null ? ov.ship : ""}" value="${esc(x.expiry)}" title="換一個效期">` : ""}` : ""}
      ${ov ? `<div class="kbd">${esc(ov.operator)} ${when(ov.at)} 改過・<a class="lnk" data-clr="${x.order_id}">回到系統判斷</a></div>` : ""}</td>`;
}
/* ③ 區塊：勇信表收到沒有、填效期匯入檔；④ 入庫通知單與嘜頭 */
function checkSec(ck, sec) {
  const s = ck.summary;
  const top = s.yx_at ? `<div class="kbd mb-1"><i class="bi bi-check2-circle" style="color:var(--ok)"></i> 勇信配送明細表 ${when(s.yx_at)} 上傳</div>`
    : `<div class="alert al-warn mb-2">${s.wait === s.items ? "還沒收到這張單的勇信配送明細表。" : `還差 ${s.wait} 品的勇信配送明細表。`}收到後丟進最上面的上傳區。</div>`;
  const sum = `<div class="sd-kv"><span>寫入效期 <b>${s.ship}</b> 品</span>${s.noship ? `<span>不出貨 <b>${s.noship}</b> 品</span>` : ""}${s.pending ? `<span style="color:var(--warn)">要決定 <b>${s.pending}</b> 品</span>` : ""}${s.wait ? `<span>等勇信表 <b>${s.wait}</b> 品</span>` : ""}</div>`;
  const foot = s.ship ? `<div class="flex items-center gap-2 flex-wrap mt-2"><button class="btn btn-p btn-sm" id="ck-fill" ${s.pending || s.wait ? "disabled" : ""}><i class="bi bi-download"></i> 下載填好效期的匯入檔（${s.ship} 筆）</button>
      <span class="kbd">${s.pending ? "還有要決定的品項（上面品項表標黃的）。" : "先從 MOMO 後台下載這張的效期匯入檔，按這裡選檔（或丟進上傳區）。"}${s.filled_at ? ` 上次填好：${when(s.filled_at)}` : ""}</span>
      <input type="file" id="ck-file" accept=".xls,.xlsx" class="hidden"></div>` : "";
  const nb = s.noship_bought + s.skip;
  return sec(`<span style="color:var(--s8)">③</span> 缺貨與效期`, top + sum + foot)
    + (s.yx_at ? sec(`<span style="color:var(--s8)">④</span> 入庫通知單與嘜頭`, ((nb || s.short) ? noticeBody(ck) : `<div class="kbd">全部照訂單出，入庫通知單不用畫。</div>`) + markBody(ck)) : "");
}
function noticeBody(ck) {
  const s = ck.summary, nt = s.notice;
  const lines = ck.items.filter(x => x.decision === "noship" || (x.decision === "ship" && x.actual != null && x.actual !== (x.qty || 0)))
    .map(x => `<div class="sd-it"><span class="fn">${esc(x.factory_code)}</span> ${esc((x.name || "").slice(0, 24))}
      <div class="sd-kv">${x.decision === "noship" ? `<span class="badge b-bad">不出貨</span><span>整列畫線、商品有效日期那格蓋章</span>`
        : `<span class="badge b-warn">短出</span><span>數量 ${fmt(x.qty)} 畫掉，寫實際 <b>${fmt(x.actual)}</b> ${esc(x.unit)}（勇信 ${fmt(x.shipped)} 箱 × ${x.conv}），蓋章</span>`}</div></div>`).join("");
  const state = !s.done ? `<div class="alert al-warn">③ 還沒做完（效期檔還沒填或還有要決定的），畫出來的會不準。</div>` : "";
  const file = nt ? `<div class="kbd mt-1"><i class="bi bi-check2-circle" style="color:var(--ok)"></i> 入庫通知單 ${esc(nt.filename)}，${esc(nt.uploaded_by)} ${when(nt.uploaded_at)} 上傳${nt.drawn_at ? `；${esc(nt.drawn_by)} ${when(nt.drawn_at)} 畫過` : ""}</div>
      <div class="flex items-center gap-2 flex-wrap mt-2"><a class="btn btn-p btn-sm nt-dl" href="/api/momo/notice/${encodeURIComponent(ck.inbound_no || SIDE)}/file"><i class="bi bi-download"></i> 下載畫好的入庫通知單</a>
      <span class="kbd">${STAMP_HAS ? "" : "還沒上傳業務專用章，會畫紅框代替；到「資料與設定」上傳。"}</span></div>`
    : `<div class="alert al-blue mt-1" style="background:#eef2ff;color:#3730a3;border:1px solid #c7d2fe">還沒收到這張的入庫通知單：從 MOMO 後台下載 PDF，丟進最上面的上傳區。</div>`;
  return state + lines + file;
}
function markBody(ck) {
  const s = ck.summary, mk = s.mark, no = ck.inbound_no || SIDE;
  const head = `<div class="sd-it"><b>嘜頭</b>：「共＿箱」填勇信出的箱數，「第＿箱」留白給倉庫寫；不出貨的那頁拿掉。</div>`;
  return head + (mk ? `<div class="kbd"><i class="bi bi-check2-circle" style="color:var(--ok)"></i> 嘜頭 ${esc(mk.filename)}，${esc(mk.uploaded_by)} ${when(mk.uploaded_at)} 上傳${mk.drawn_at ? `；${esc(mk.drawn_by)} ${when(mk.drawn_at)} 下載過` : ""}</div>
      <div class="mt-2"><a class="btn btn-p btn-sm nt-dl" href="/api/momo/mark/${encodeURIComponent(no)}/file"><i class="bi bi-download"></i> 下載填好的嘜頭</a></div>`
    : `<div class="kbd">還沒收到這張的嘜頭：從 MOMO 後台下載嘜頭 PDF，丟進最上面的上傳區。</div>`);
}
let STAMP_HAS = true;
function bindCheck(no) {
  $$("#sd-body a.nt-dl").forEach(a => a.addEventListener("click", () => setTimeout(() => { refresh(); openSide(no); }, 1500)));
  const btn = $("#ck-fill");
  if (btn) { btn.addEventListener("click", () => $("#ck-file").click());
    $("#ck-file").addEventListener("change", () => { const f = $("#ck-file").files[0]; if (f) fillExp(f); }); }
}
/* 品項表每一列的按鈕（表會因為搜尋、篩選重畫，所以每次畫完重綁） */
function bindRows(no) {
  const T = "#sd-items";
  $$(`${T} [data-adj]`).forEach(b => b.addEventListener("click", () => openAdj(JSON.parse(b.dataset.adj))));
  const ck = (body, msg) => post("/api/momo/check", body).then(() => { toast(msg); refresh(); openSide(no); }).catch(err => toast(err.message, "err"));
  const exp = id => { const el = $(`${T} [data-exp="${id}"]`); return el ? el.value : ""; };
  $$(`${T} [data-ship1]`).forEach(b => b.addEventListener("click", () => ck({ order_id: b.dataset.ship1, ship: 1, expiry: exp(b.dataset.ship1) }, "已改成出貨")));
  $$(`${T} [data-ship0]`).forEach(b => b.addEventListener("click", () => ck({ order_id: b.dataset.ship0, ship: 0 }, "已改成不出貨")));
  const setExp = (id, v) => { const i = $(`${T} [data-exp="${id}"]`); ck({ order_id: id, ship: i ? i.dataset.ov : "", expiry: v }, `效期改成 ${slash(v)}`); };
  $$(`${T} [data-exp]`).forEach(i => i.addEventListener("change", () => i.value && setExp(i.dataset.exp, i.value)));
  $$(`${T} [data-showexp]`).forEach(a => a.addEventListener("click", () => { a.remove(); $(`${T} [data-exp="${a.dataset.showexp}"]`).classList.remove("hidden"); }));
  $$(`${T} [data-pick]`).forEach(a => a.addEventListener("click", () => setExp(a.dataset.pick, a.dataset.v)));
  $$(`${T} [data-clr]`).forEach(a => a.addEventListener("click", () => ck({ order_id: a.dataset.clr }, "已回到系統判斷")));
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
/* 清除資料（限管理員；Jerry 2026-10-29，照蝦皮特選） */
if ($("#rs-go")) $("#rs-go").addEventListener("click", async () => {
  const body = { logs: $("#rs-logs").checked, confirm: $("#rs-confirm").value.trim() };
  if (body.confirm !== "清空資料") return toast("請輸入「清空資料」以確認", "err");
  if (!confirm("確定清除 MOMO 的訂單資料？清除後無法復原。")) return;
  try { const r = await post("/api/momo/reset", body); toast(r.message); $("#rs-confirm").value = ""; closeSide(); refresh(); loadFresh(); loadUploads(); }
  catch (err) { toast(err.message, "err"); }
});
function showSettings(on) {
  $("#view-work").classList.toggle("hidden", on); $("#view-set").classList.toggle("hidden", !on);
  if (on) { closeSide(); loadWh(); loadUploads(); loadStamp(); }
}
async function loadStamp() {
  let d; try { d = await api("/api/momo/stamp"); } catch (err) { return; }
  STAMP_HAS = d.has;
  $("#stamp-st").innerHTML = d.has ? `<img src="/api/momo/stamp.png?t=${Date.now()}" style="height:64px"><span class="kbd">${esc(d.updated_by)} ${when(d.updated_at)} 上傳</span>` : `<span class="badge b-warn">還沒上傳</span><span class="kbd">沒上傳的話畫出來的入庫通知單會用紅框寫「業務專用章」代替。</span>`;
}
$("#stamp-file").addEventListener("change", async () => {
  const f = $("#stamp-file").files[0]; $("#stamp-file").value = ""; if (!f) return;
  const fd = new FormData(); fd.append("file", f);
  try { await api("/api/momo/stamp", { method: "POST", body: fd }); toast("業務專用章已更新"); loadStamp(); } catch (err) { toast(err.message, "err"); }
});
$("#stamp-btn").addEventListener("click", () => $("#stamp-file").click());
let WH_LABELS = {};
async function loadWh() {
  let d; try { d = await api("/api/momo/warehouses"); } catch (err) { toast(err.message, "err"); return; }
  WH_LABELS = d.labels;
  // MOMO 檔出現過、主檔還沒有的倉：點一下加成新的倉，或當成某個倉的別名（倉名不給人打，才會跟 MOMO 檔一字不差；Jerry 2026-10-29）
  const opts = d.warehouses.map(w => `<option value="${esc(w.name)}">${esc(w.name)}</option>`).join("");
  $("#wh-missing").innerHTML = d.missing.length ? `<div class="alert al-warn"><b>MOMO 檔出現過、主檔還沒有的倉（${d.missing.length} 個）</b>：沒加的話採購單的地址、電話會空白。
      ${d.missing.map(m => `<div class="flex items-center gap-2 flex-wrap mt-2" data-miss="${esc(m.name)}"><b>${esc(m.name)}</b><span class="kbd">${m.orders} 張入庫單</span>
        <button class="btn btn-p btn-sm wm-add"><i class="bi bi-plus-lg"></i> 加成新的倉</button><span class="kbd">或是</span>
        <select class="inp wm-of" style="padding:3px 6px"><option value="">選一個倉…</option>${opts}</select><button class="btn btn-o btn-sm wm-alias">的別名</button></div>`).join("")}</div>`
    : `<div class="kbd"><i class="bi bi-check-circle" style="color:var(--ok)"></i> MOMO 檔裡的倉別主檔都有。</div>`;
  $$("#wh-missing [data-miss]").forEach(row => {
    const name = row.dataset.miss;
    row.querySelector(".wm-add").addEventListener("click", async () => {
      try { await post("/api/momo/warehouses", { name }); toast(`已加入 ${name}，請把地址、電話等填上`); loadWh(); refresh(); } catch (err) { toast(err.message, "err"); }
    });
    row.querySelector(".wm-alias").addEventListener("click", async () => {
      const of = row.querySelector(".wm-of").value; if (!of) return toast("請先選是哪個倉", "err");
      try { await post("/api/momo/warehouses", { name, alias_of: of }); toast(`以後「${name}」就當 ${of}`); loadWh(); refresh(); } catch (err) { toast(err.message, "err"); }
    });
  });
  const cols = ["aliases", "address", "hours", "phone", "contact", "ship_to", "mars_name"];
  const show = (w, c) => c === "aliases" ? w.aliases.join("、") : w[c];
  $("#wh-tbl").innerHTML = `<tr><th>倉別</th>${cols.map(c => `<th>${esc(WH_LABELS[c])}</th>`).join("")}</tr>` + d.warehouses.map(w => `<tr data-wh="${esc(w.name)}">
    <td class="nw"><b>${esc(w.name)}</b>${w.custom ? `<div><span class="badge b-blue">畫面上加的</span></div>` : ""}${w.edited ? `<div><span class="badge b-s">改過</span></div>` : ""}</td>
    ${cols.map(c => `<td class="wh-edit" data-f="${c}" style="${c === "hours" ? "min-width:260px" : ""}" ${c === "aliases" ? 'title="MOMO 檔寫成別的樣子也認得；好幾個用頓號分開"' : ""}>${esc(show(w, c)) || `<span class="kbd">（空白）</span>`}</td>`).join("")}</tr>`).join("");
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

drawPicks(); loadFresh(); loadPending(); refresh(); api("/api/momo/stamp").then(d => { STAMP_HAS = d.has; }).catch(() => {});

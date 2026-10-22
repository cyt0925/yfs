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
function toast(msg, kind) {
  const t = document.createElement("div"); t.className = "toast" + (kind === "err" ? " err" : kind === "warn" ? " warn" : ""); t.textContent = msg; $("#toasts").appendChild(t);
  // 有要人處理的（warn：價格本上傳後還有品項找不到）留久一點、點一下才關；長的訊息也多留幾秒讀完
  if (kind === "warn") { t.title = "點一下關閉"; t.addEventListener("click", () => t.remove()); setTimeout(() => t.remove(), 30000); }
  else setTimeout(() => t.remove(), kind === "err" ? 6000 : msg.length > 60 ? 7000 : 2800);
}
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
  if (kind === "conv") toast(Object.entries(d.lines).map(([l, x]) => `${l}價格本已更新：${fmt(x.count)} 筆（原 ${fmt(x.before)} 筆）${x.blank ? `，其中 ${fmt(x.blank)} 筆箱入數空白` : ""}`
      + convPending((d.pending || {})[l])).join("\n")
    + (d.warnings.length ? `\n${d.warnings.slice(0, 3).join("\n")}` : ""), (Object.values(d.pending || {}).some(p => p.missing.length || p.no_box.length)) ? "warn" : undefined);
  else if (kind === "stock") toast(`已更新${d.label}：${f.name}（${fmt(d.rows)} 列）` + (d.notes.length ? `\n${d.notes.join("\n")}` : ""));
  else if (kind === "mars") toast(`已更新瑪氏商品總表：${fmt(d.codes)} 個料號（${fmt(d.rows)} 列）${d.added || d.removed ? `，新增 ${d.added}、移除 ${d.removed}` : ""}`
    + ((d.kept_manual || []).length ? `\n保留在蝦皮特選補登的 ${d.kept_manual.length} 品（Excel 裡還沒有）` : "") + (d.warnings.length ? `\n${d.warnings.slice(0, 3).join("\n")}` : ""));
  else if (kind === "yx") { showYx(d); return false; }
  return true;
}

/* 價格本上傳後：還在等價格、不能產採購單的品項找到了沒有（Becky 2026-10-16「更新價格本了 是不是得重拋採購單？」——不用重拋，這裡直接講） */
function convPending(p) {
  if (!p || !p.waiting) return "";
  const list = xs => xs.slice(0, 5).map(x => `${x.code3}${x.name ? `（${x.name.slice(0, 14)}）` : ""}`).join("、") + (xs.length > 5 ? ` 等 ${xs.length} 個` : "");
  const out = [];
  if (p.was_blocked) out.push(`原本找不到價格的 ${p.was_blocked} 個品項，這次找到 ${p.fixed.length} 個`);
  if (p.missing.length) out.push(`還有 ${p.missing.length} 個價格本裡沒有：${list(p.missing)}`);
  if (p.no_box.length) out.push(`${p.no_box.length} 個價格本有、但沒有箱入數：${list(p.no_box)}`);
  if (!out.length || (!p.missing.length && !p.no_box.length)) out.push(`待產採購單的品項都找得到價格了，不用重拋訂單，直接到「待產採購單」產出即可`);
  return "\n" + out.join("\n");
}

/* 上傳區右邊：蝦皮採購單／入庫單最後一次匯入、價格本、庫存資料什麼時候更新的（點了到「資料與設定」）。
   蝦皮採購單排第一個：只列參考資料會讓人以為全部傳好了（Jerry 2026-10-13：「上面有寫要上傳蝦皮採購單嗎??」） */
async function loadFresh() {
  let c, k, u, m; try { [c, k, u, m] = await Promise.all([api("/api/shopee/conv"), api("/api/shopee/stock"), api("/api/shopee/uploads"), api("/api/mars/status")]); } catch (err) { return; }
  const md = t => t ? `${Number(t.slice(5, 7))}/${Number(t.slice(8, 10))}` : "";
  const today = iso(new Date(Date.now() - new Date().getTimezoneOffset() * 6e4));
  const chips = [["purchase", "蝦皮採購單"], ["inbound", "蝦皮入庫單"]].map(([kind, label]) => {
    const x = u.uploads.find(r => r.kind === kind), at = x && (x.committed_at || x.uploaded_at);
    if (x) return `<span title="${esc(x.filename)}"><i class="bi bi-check-circle" style="color:var(--ok)"></i> ${label} ${at.slice(0, 10) === today ? "今天" : md(at)}</span>`;
    return kind === "purchase" ? `<span class="w"><i class="bi bi-exclamation-circle"></i> ${label}未匯入</span>`
      : `<span title="竹運出貨產拋單檔需要入庫單號"><i class="bi bi-dash-circle"></i> ${label}未匯入</span>`;
  });
  chips.push(...["寶僑", "紙潔"].map(l => { const x = c.lines[l]; return x ? `<span title="${esc(x.file || "")}"><i class="bi bi-check-circle" style="color:var(--ok)"></i> ${l}價格本 ${md(x.updated_at)}</span>`
    : `<span class="w"><i class="bi bi-exclamation-circle"></i> ${l}價格本未上傳</span>`; }));
  const mu = m.last_upload;     // 瑪氏商品總表跟瑪氏出貨頁共用，這裡也能傳
  chips.push(mu ? `<span title="${esc(mu.filename)}${(m.manual || []).length ? `；蝦皮特選補登 ${m.manual.length} 品` : ""}"><i class="bi bi-check-circle" style="color:var(--ok)"></i> 瑪氏商品總表 ${mu.uploaded_at.slice(0, 10) === today ? "今天" : md(mu.uploaded_at)}</span>`
    : `<span class="w"><i class="bi bi-exclamation-circle"></i> 瑪氏商品總表未上傳</span>`);
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
  // 履約方式每個線別各選一次（Jerry 2026-10-13：採購單常常三個線別混在一起）。瑪氏預設供應商直送；竹運採購進貨瑪氏還沒開放
  const nl = d.new_lines && Object.keys(d.new_lines).length ? d.new_lines : { "": c.new }, pick = {};
  if (d.needs_fulfil) h += `<div class="fulfil" style="display:block"><b>履約方式</b><span class="kbd">（每個線別各選一次，套用到這個檔的新訂單；匯入後可整張修改）</span>
    ${Object.entries(nl).map(([ln, n]) => {
      if (ln === "瑪氏") pick[ln] = "供應商直送";
      return `<div class="ffl" data-l="${esc(ln)}"><span class="ffl-l"><b>${esc(ln || "新訂單")}</b> ${fmt(n)} 項</span>${FULFILS.map(f => {
        const off = ln === "瑪氏" && f === "竹運採購進貨";
        return `<span class="chip ff ${pick[ln] === f ? "on" : ""} ${off ? "off" : ""}" data-f="${esc(f)}" ${off ? 'title="瑪氏目前只開放竹運出貨、供應商直送"' : ""}>${esc(f)}</span>`;
      }).join("")}</div>`; }).join("")}</div>`;
  const need = () => Object.keys(nl).filter(ln => !pick[ln]);
  h += `<div class="flex items-center gap-2 mt-3"><button class="btn btn-p ok"><i class="bi bi-check2"></i> 確認匯入</button>
    <button class="btn btn-o cancel">取消</button><span class="kbd need"></span></div>`;
  box.innerHTML = h;
  const sync = () => {
    const miss = d.needs_fulfil ? need() : [];
    box.querySelector(".ok").disabled = miss.length > 0;
    box.querySelector(".need").textContent = miss.length ? `請先選擇${miss.map(x => x || "新訂單").join("、")}的履約方式` : "";
  };
  box.querySelectorAll(".ffl").forEach(row => row.querySelectorAll(".ff:not(.off)").forEach(ch => ch.addEventListener("click", () => {
    row.querySelectorAll(".ff").forEach(x => x.classList.remove("on")); ch.classList.add("on"); pick[row.dataset.l] = ch.dataset.f; sync();
  })));
  sync();
  box.querySelector(".cancel").addEventListener("click", async () => { try { await post("/api/shopee/import/cancel", { upload_id: d.upload_id }); } catch (e) {} box.remove(); });
  box.querySelector(".ok").addEventListener("click", async e => {
    e.target.disabled = true;
    try {
      const r = await post("/api/shopee/import/commit", { upload_id: d.upload_id, fulfils: pick });
      const k = r.counts; box.remove();
      toast(`已匯入 ${d.filename}：新增 ${k.new} 項、有變 ${k.changed} 項、移除 ${k.remove} 項、無異動 ${k.same} 項`);
      loadFresh(); refresh();
    } catch (err) { e.target.disabled = false; toast(err.message, "err"); }
  });
}

/* ── 共用篩選、作業流程格子 ─────────────────────────────── */
let VIEW = { orders: [], stages: [] }, ORD = new Map(), STAGE = "", EIP_TODO = 0; const OPEN = new Set();
const iso = d => d.toISOString().slice(0, 10);
$("#f-from").value = iso(new Date(Date.now() - 7 * 864e5));
["#f-from", "#f-to", "#f-up-from", "#f-up-to", "#f-line"].forEach(s => $(s).addEventListener("change", () => { BSEL.clear(); ZSEL.clear(); KSEL.clear(); refresh(); }));
$("#f-q").addEventListener("keydown", e => { if (e.key === "Enter") refresh(); });
const fq = () => ({ from: $("#f-from").value, to: $("#f-to").value, up_from: $("#f-up-from").value, up_to: $("#f-up-to").value, line: $("#f-line").value });
const inStage = no => STAGE === "all" || ((ORD.get(no) || {}).todo || []).includes(STAGE);
const STAGE_DEF = [
  { key: "diff", t: "差異待確認", d: "蝦皮重傳後內容有變", act: "確認差異", bad: true,
    hint: "蝦皮重傳後數量、到貨日或倉別有變；確認差異後，訂單才能往下處理。" },
  { key: "buy", t: "待產採購單", d: "竹運採購進貨、直送", act: "產出 EIP 採購單",
    hint: "勾選訂單產出 EIP 採購單；「需處理」的訂單請點該列，在右側修正。" },
  { key: "eip", t: "待回填 EIP 採購單號", d: "上傳 EIP 後回填", act: "回填 EIP 採購單號",
    hint: "上傳 EIP 後回填採購單號，回填後即鎖定；回填後瑪氏可下載瑪氏採購單，寶僑、紙潔直送可下載直送採購單（給採購轉原廠）。" },
  { key: "ship", t: "待產拋單檔", d: "竹運出貨", act: "產出拋單檔",
    hint: "系統依 EC 庫存、竹運效期判斷可出貨量；勾選訂單產出拋單檔，缺貨會自動列入「待通知蝦皮」。" },
  { key: "notify", t: "待通知蝦皮", d: "有缺貨、尚未通知", act: "通知蝦皮",
    hint: "勾選訂單產生 LINE 通知文字。瑪氏缺貨請上傳勇信配送明細表；寶僑、紙潔缺貨請點該單，在右側填寫。" },
  { key: "all", t: "全部訂單", d: "", hint: "點一張單，右側看品項、採購單、缺貨、差異與歷程；採購來信缺貨時，搜尋訂單後在右側填寫。" },
];
const SDEF = Object.fromEntries(STAGE_DEF.map(x => [x.key, x]));

/* ── 表頭篩選（Jerry 2026-10-19「可以做個篩選嗎，比較好篩出需要的部分」，看示意圖選「跟 Excel 一樣」；
   隔天「全部都要」：每一欄都有 ▾，不只倉別、履約方式這幾欄）──
   點 ▾ 勾要看的值（後面寫幾張，照其他欄的篩選算）；值多的有搜尋框；數字、日期可以排序。到貨日、線別、搜尋還是右上角那組（每一步共用），
   表頭再篩是在那組的結果裡縮小。表頭篩選只管目前這一步，換一步就清掉。被篩掉的訂單同時取消勾選，免得產出畫面上看不到的單。
   右側面板開著時有幾欄會收起來，所以數字列寫出篩了哪幾欄，不會不知道為什麼少了。 */
const HF = {};                                   // {欄位: Set(要看的值)}；沒有這個欄位＝不篩
let HS = null;                                   // 排序 {k: 欄位, dir: 1 由小到大／-1 由大到小}
let HF_COLS = {}, HF_ROWS = [], HF_REDRAW = null, HF_UNIT = "張", HF_POP = null, HF_KEY = "";
const hfPass = (r, skip) => Object.entries(HF).every(([k, s]) => k === skip || !HF_COLS[k] || HF_COLS[k].get(r).some(v => s.has(v)));
const hfCmp = (a, b) => typeof a === "number" && typeof b === "number" ? a - b : String(a).localeCompare(String(b), "zh-Hant", { numeric: true });
function hfUse(rows, cols, redraw, unit = "張") {  // 每個清單畫之前呼叫，回篩過、排過的列
  HF_COLS = cols; HF_ROWS = rows; HF_REDRAW = redraw; HF_UNIT = unit;
  Object.keys(HF).forEach(k => { if (!cols[k]) delete HF[k]; });
  if (HS && !(cols[HS.k] && cols[HS.k].sort)) HS = null;
  const out = rows.filter(r => hfPass(r));
  if (!HS) return out;
  const f = cols[HS.k].sort;
  return out.map((r, i) => [r, i]).sort((a, b) => hfCmp(f(a[0]), f(b[0])) * HS.dir || a[1] - b[1]).map(x => x[0]);
}
const hfShow = (k, v) => HF_COLS[k] && HF_COLS[k].num && v !== "—" ? fmt(Number(v)) : v;
const hfWords = c => c.date ? ["由早到晚", "由晚到早"] : ["由小到大", "由大到小"];
function hfTh(k, cls = "") {
  const c = HF_COLS[k], s = HS && HS.k === k ? `<i class="bi ${HS.dir > 0 ? "bi-sort-up" : "bi-sort-down"} hf-sorted" title="${hfWords(c)[HS.dir > 0 ? 0 : 1]}"></i>` : "";
  return `<th class="${cls}${HF[k] ? " hf-on" : ""}">${esc(c.label)}${s}<button type="button" class="hf-b" data-hf="${k}" title="篩選${esc(c.label)}" aria-label="篩選${esc(c.label)}"><i class="bi bi-caret-down-fill"></i></button></th>`;
}
function hfNote(n) {          // 數字列最後：篩了哪幾欄（右側面板開著時那幾欄可能收起來了）、怎麼排的
  const parts = Object.entries(HF).filter(([k]) => HF_COLS[k]).map(([k, s]) => `${HF_COLS[k].label}（${s.size <= 2 ? [...s].map(v => hfShow(k, v)).join("、") : `${s.size} 項`}）`);
  const srt = HS ? `依${HF_COLS[HS.k].label}${hfWords(HF_COLS[HS.k])[HS.dir > 0 ? 0 : 1]}排` : "";
  if (!parts.length && !srt) return "";
  return `<span class="hf-note">${parts.length ? `已篩選：${esc(parts.join("、"))}，篩選後 ${fmt(n)} ${HF_UNIT}` : ""}${parts.length && srt ? "；" : ""}${esc(srt)} · <a href="#" class="hf-clr">清除表頭篩選</a></span>`;
}
const hfEmpty = cols => `<tr><td colspan="${cols}" class="muted" style="padding:24px;text-align:center">表頭篩選後沒有符合的${HF_UNIT === "張" ? "訂單" : "資料"}。
  <div class="mt-2"><button class="btn btn-o btn-sm hf-clr"><i class="bi bi-funnel"></i> 清除表頭篩選</button></div></td></tr>`;
const hfKeep = (sel, shown) => { const v = new Set(shown.map(g => g.order_no)); [...sel].forEach(k => { if (!v.has(k)) sel.delete(k); }); };
/* 上面「13 張可產出」「7 張需先處理」：點了等於那一欄只勾這幾個值，再點一次取消 */
function hfQuick(k, vals, html) {
  const on = !!HF[k] && HF[k].size === vals.length && vals.every(v => HF[k].has(v));
  return vals.length ? `<span class="hf-q${on ? " on" : ""}" data-hf="${k}" data-v="${esc(JSON.stringify(vals))}" role="button" tabindex="0" title="${on ? "取消篩選" : "只看這些"}">${html}</span>` : `<span>${html}</span>`;
}
const hfClear = () => { Object.keys(HF).forEach(k => delete HF[k]); HS = null; hfClose(); };
function hfClose() { if (HF_POP) HF_POP.remove(); HF_POP = null; HF_KEY = ""; }
function hfOpen(btn) {
  const k = btn.dataset.hf, c = HF_COLS[k]; if (!c) return;
  if (HF_KEY === k) return hfClose();
  hfClose(); HF_KEY = k;
  const rows = HF_ROWS.filter(r => hfPass(r, k)), cnt = new Map();      // 張數照其他欄的篩選算（跟 Excel 一樣）
  rows.forEach(r => new Set(c.get(r)).forEach(v => cnt.set(v, (cnt.get(v) || 0) + 1)));
  const first = c.first || [], grp = c.group || (() => "");
  const byVal = c.num || c.date || c.byValue;   // 數字、日期、單號照值排；其他照張數多到少
  const vals = [...cnt.keys()].sort((a, b) => ((first.indexOf(a) + 1 || 99) - (first.indexOf(b) + 1 || 99))
    || (byVal ? (a === "—") - (b === "—") || hfCmp(c.num ? Number(a) : a, c.num ? Number(b) : b) : cnt.get(b) - cnt.get(a) || hfCmp(a, b)));
  const on = v => !HF[k] || HF[k].has(v), seen = new Set();
  const opt = (v, g) => `<label class="hf-o${g ? " hf-in" : ""}"><input type="checkbox" class="ck hf-v" value="${esc(v)}" data-g="${esc(g)}" ${on(v) ? "checked" : ""}><span>${esc(hfShow(k, v))}</span><span class="hf-c">${fmt(cnt.get(v))} ${HF_UNIT}</span></label>`;
  let h = "";
  for (const v of vals) {
    const g = grp(v);
    if (!g) { h += opt(v, ""); continue; }
    if (seen.has(g)) continue;
    seen.add(g);
    const kids = vals.filter(x => grp(x) === g), n = rows.filter(r => c.get(r).some(x => grp(x) === g)).length;   // 一張單可能同時有好幾個原因
    h += `<label class="hf-o hf-gl" data-g="${esc(g)}"><input type="checkbox" class="ck hf-g" data-g="${esc(g)}" ${kids.every(on) ? "checked" : ""}><span>${esc(g)}</span><span class="hf-c">${fmt(n)} ${HF_UNIT}</span></label>`
      + kids.map(x => opt(x, g)).join("");
  }
  const w = hfWords(c), sorted = d => HS && HS.k === k && HS.dir === d ? " on" : "";
  const pop = document.createElement("div"); pop.id = "hf-pop"; pop.setAttribute("role", "dialog"); pop.setAttribute("aria-label", `篩選${c.label}`);
  pop.innerHTML = `<div class="hf-t">${esc(c.label)}</div>`
    + (c.sort ? `<div class="hf-s"><button type="button" class="chip hf-sort${sorted(1)}" data-d="1"><i class="bi bi-sort-up"></i> ${w[0]}</button><button type="button" class="chip hf-sort${sorted(-1)}" data-d="-1"><i class="bi bi-sort-down"></i> ${w[1]}</button></div>` : "")
    + (vals.length > 8 ? `<input class="inp hf-find" placeholder="搜尋${esc(c.label)}" aria-label="搜尋${esc(c.label)}">` : "")
    + `<div class="hf-list">${h || `<div class="muted">目前沒有可以篩的值。</div>`}</div>
    <div class="hf-a"><button type="button" class="btn btn-o btn-sm hf-all">全選</button><button type="button" class="btn btn-p btn-sm hf-ok">套用</button></div>`;
  document.body.appendChild(pop); HF_POP = pop;
  const r = btn.closest("th").getBoundingClientRect();
  pop.style.left = `${Math.max(8, Math.min(r.left, document.documentElement.clientWidth - pop.offsetWidth - 8))}px`; pop.style.top = `${r.bottom + 4}px`;
  const boxes = () => [...pop.querySelectorAll("input.hf-v")], shown = b => b.closest("label").style.display !== "none";
  const sync = () => {
    pop.querySelectorAll("label.hf-gl").forEach(gl => { const kids = boxes().filter(b => b.dataset.g === gl.dataset.g);
      gl.querySelector("input").checked = kids.every(b => b.checked); gl.style.display = kids.some(shown) ? "" : "none"; });
    pop.querySelector(".hf-ok").disabled = !boxes().some(b => b.checked);
  };
  pop.querySelectorAll("input.hf-g").forEach(gb => gb.addEventListener("change", () => { boxes().filter(b => b.dataset.g === gb.dataset.g && shown(b)).forEach(b => { b.checked = gb.checked; }); sync(); }));
  boxes().forEach(b => b.addEventListener("change", sync));
  const find = pop.querySelector(".hf-find");
  if (find) {           // 跟 Excel 一樣：打字後只勾找到的，套用就只看找到的
    find.addEventListener("input", () => {
      const q = find.value.trim().toLowerCase();
      boxes().forEach(b => { const hit = !q || String(hfShow(k, b.value)).toLowerCase().includes(q) || b.value.toLowerCase().includes(q);
        b.closest("label").style.display = hit ? "" : "none"; b.checked = hit; });
      sync();
    });
    find.addEventListener("keydown", e => { if (e.key === "Enter" && !pop.querySelector(".hf-ok").disabled) pop.querySelector(".hf-ok").click(); });
    setTimeout(() => find.focus(), 0);
  }
  pop.querySelector(".hf-all").addEventListener("click", () => { boxes().filter(shown).forEach(b => { b.checked = true; }); sync(); });
  pop.querySelectorAll(".hf-sort").forEach(sb => sb.addEventListener("click", () => {
    const d = Number(sb.dataset.d); HS = HS && HS.k === k && HS.dir === d ? null : { k, dir: d };   // 再點一次取消排序
    hfClose(); if (HF_REDRAW) HF_REDRAW();
  }));
  pop.querySelector(".hf-ok").addEventListener("click", () => {
    const pick = boxes().filter(b => b.checked).map(b => b.value);
    if (pick.length === boxes().length) delete HF[k]; else HF[k] = new Set(pick);
    hfClose(); if (HF_REDRAW) HF_REDRAW();
  });
  sync();
}
document.addEventListener("click", e => {
  const b = e.target.closest(".hf-b");
  if (b) { e.preventDefault(); e.stopPropagation(); hfOpen(b); return; }
  const clr = e.target.closest(".hf-clr");
  if (clr) { e.preventDefault(); hfClear(); if (HF_REDRAW) HF_REDRAW(); return; }
  const q = e.target.closest(".hf-q");
  if (q) {
    const vals = JSON.parse(q.dataset.v); hfClose();
    if (q.classList.contains("on")) delete HF[q.dataset.hf]; else HF[q.dataset.hf] = new Set(vals);
    if (HF_REDRAW) HF_REDRAW(); return;
  }
  if (HF_POP && !e.target.closest("#hf-pop")) hfClose();
});
document.addEventListener("keydown", e => {
  if (e.key === "Escape" && HF_POP) { hfClose(); return; }
  const q = e.target.closest && e.target.closest(".hf-q");
  if (q && (e.key === "Enter" || e.key === " ")) { e.preventDefault(); q.click(); }
});
window.addEventListener("scroll", e => { if (HF_POP && !HF_POP.contains(e.target)) hfClose(); }, true);
window.addEventListener("resize", hfClose);
/* 各清單共用的欄位（g 是一張訂單；待回填那步是一個 EIP 檔，欄名一樣的也共用） */
const HF_NO = { label: "訂單編號", get: g => [g.order_no], byValue: true };
const HF_LINE = { label: "線別", get: g => [g.line || "—"] };
const HF_DATE = { label: "到貨日", get: g => [g.expected_date || "—"], sort: g => g.expected_date || "", date: true };
const HF_WH = { label: "倉別", get: g => [wh(g.warehouse)] };
const HF_FF = { label: "履約方式", get: g => [g.fulfil || "—"] };
const hfNum = (label, f) => ({ label, num: true, get: g => [f(g) == null ? "—" : String(f(g))], sort: g => f(g) == null ? -Infinity : Number(f(g)) });
/* 上傳日＝那張單第一次進系統那天；永豐料號＝蝦皮採購單的 Supplier SKU ID（OP 2026-10-21，Jerry 看過示意圖） */
const HF_UP = { label: "上傳日", get: g => [g.upload_date || "—"], sort: g => g.upload_date || "", date: true };
const codesOf = g => g.codes || [];
const HF_CODES = { label: "永豐料號", get: g => codesOf(g).length ? codesOf(g) : ["—"], byValue: true };
const codesTd = g => {        // 一張單先列 3 個，其餘「+N」滑過去看全部；完整的在右側面板
  const c = codesOf(g); if (!c.length) return `<span class="muted">—</span>`;
  return `<div class="fn codes">${c.slice(0, 3).map(esc).join("<br>")}${c.length > 3 ? ` <span class="more" title="${esc(c.join("、"))}">+${c.length - 3}</span>` : ""}</div>`;
};

/* 清單是空的：分清楚「還沒匯入」「被篩選條件篩掉」「這一步沒事」，並告訴人下一步做什麼 */
function emptyRow(cols) {
  const a = VIEW.all_orders || {}, md = t => t ? `${Number(t.slice(5, 7))}/${Number(t.slice(8, 10))}` : "";
  let m = "這一步目前沒有要處理的訂單。";
  if (!a.count) m = `<b>尚未匯入蝦皮訂單。</b><br>請將蝦皮採購單（檔名 PurchaseOrder 開頭的 Excel）拖曳至上方「上傳檔案」，選擇履約方式後按「確認匯入」。`;
  else if (!VIEW.orders.length) m = `目前的到貨日、上傳日、線別或搜尋條件下沒有訂單（系統共有 ${fmt(a.count)} 張，到貨日 ${md(a.first)}～${md(a.last)}）。
    <div class="mt-2"><button class="btn btn-o btn-sm clr-f"><i class="bi bi-funnel"></i> 清除篩選，顯示全部訂單</button></div>`;
  return `<tr><td colspan="${cols}" class="muted" style="padding:24px;text-align:center">${m}</td></tr>`;
}
document.addEventListener("click", e => {
  if (!e.target.closest(".clr-f")) return;
  ["#f-from", "#f-to", "#f-up-from", "#f-up-to", "#f-line", "#f-q"].forEach(s => { $(s).value = ""; });
  BSEL.clear(); ZSEL.clear(); KSEL.clear(); refresh();
});

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
  if (STAGE !== k) { ASEL.clear(); BSEL.clear(); ZSEL.clear(); KSEL.clear(); OPEN.clear(); BOPEN.clear(); ZOPEN.clear(); KOPEN.clear(); hfClear(); }
  STAGE = k;
  try { localStorage.setItem("shopee_stage", k); } catch (e) { /* 無痕視窗存不了就算了 */ }
  renderStages(); loadStage();
  followSide();
}
/* 右側卡片跟著清單走（Jerry 2026-10-21「從待產拋單檔按到待通知蝦皮，右邊卡片還停在待產拋單檔的訂單」）：
   卡片那張單不在這一步的清單裡，就換成這一步的第一張；清單空了就關掉 */
function followSide() {
  if (!SIDE_NO) return;
  if (ORD.has(SIDE_NO) && inStage(SIDE_NO)) return;
  const first = firstInStage(); if (first) openSide(first); else closeSide();
}
async function loadStage() {
  const d = SDEF[STAGE];
  $("#st-title").textContent = d.t; $("#st-hint").textContent = d.hint;
  const pane = { diff: "orders", all: "orders", buy: "buy", eip: "eip", ship: "ship", notify: "notify" }[STAGE];
  ["orders", "buy", "eip", "ship", "notify"].forEach(k => $(`#st-${k}`).classList.toggle("hidden", k !== pane));
  if (pane === "orders") render();
  else if (pane === "buy") { $("#b-err").innerHTML = ""; await loadPlan(); }
  else if (pane === "eip") renderEips();
  else if (pane === "ship") { $("#z-err").innerHTML = ""; await loadShip(); }
  else await loadShort();
}
const firstInStage = () => (VIEW.orders.find(g => inStage(g.order_no)) || {}).order_no || "";
async function refresh() {
  await Promise.all([loadOrders(), loadEipCount()]);
  if (!STAGE) {        // 一進來：網址指定的那一步（首頁今日事項連 /shopee#diff）→ 上次看的那一步 → 第一個有待辦的
    let saved = ""; try { saved = localStorage.getItem("shopee_stage") || ""; } catch (e) { /* 讀不到就算了 */ }
    const cnt = Object.fromEntries((VIEW.stages || []).map(x => [x.key, x.count])); cnt.eip = EIP_TODO;
    const hash = location.hash.slice(1);
    STAGE = SDEF[hash] ? hash : (SDEF[saved] && (saved === "all" || cnt[saved])) ? saved
      : (["diff", "buy", "eip", "ship", "notify"].find(k => cnt[k]) || "all");
  }
  SHIP = { orders: [] };               // 拋單檔計畫重算前先清掉，右側面板才不會拿到舊的
  renderStages();
  const pane = loadStage();
  if (STAGE === "ship") await pane;      // 待產拋單檔：等清單算好，面板直接拿來用；其他步驟清單和面板同時更新，不互等
  // 右側面板跟著左邊動（Jerry 2026-10-20「我左邊在操作的時候右邊要跟著動啦」）：開著的那張還在清單裡就重畫（搜尋到的品項排最前面）；
  // 被篩掉了就換成清單第一張；清單空了就關掉；搜尋只剩一張就直接打開
  const q = $("#f-q").value.trim();
  if (SIDE_NO && !(ORD.has(SIDE_NO) && inStage(SIDE_NO))) followSide();   // 被篩掉、或做完事離開這一步了 → 換成這一步的第一張
  else if (SIDE_NO) openSide(SIDE_NO);       // 做完事，右側面板也換成最新的
  else if (q && VIEW.orders.filter(g => inStage(g.order_no)).length === 1) openSide(firstInStage());
}

/* ── 差異待確認／全部訂單：一張單一列，點開看品項與操作 ─────── */
const NEXT_CLS = { diff: "b-diff", buy: "b-s", eip: "b-s", ship: "b-s", notify: "b-warn" };
function nextBadges(g) {
  if (!g.todo.length) return `<span class="badge b-ok">已處理</span>`;
  return `<span class="nx">${g.todo.map(k => `<span class="badge ${NEXT_CLS[k]} go-st" data-s="${k}" title="前往「${SDEF[k].t}」" style="cursor:pointer">${SDEF[k].act}</span>`).join("")}</span>`;
}
function render() {
  const all = VIEW.orders.filter(g => inStage(g.order_no));
  const cols = { no: HF_NO, line: HF_LINE, date: HF_DATE, up: HF_UP, wh: HF_WH, ff: HF_FF, codes: HF_CODES, items: hfNum("品項", g => g.item_count), qty: hfNum("蝦皮數量", g => g.qty),
    ship: hfNum("可出貨量", g => g.qty_ship), short: hfNum("缺貨量", g => g.qty_short),
    next: { label: "下一步", get: g => STAGE === "diff" ? ["確認差異"] : g.todo.length ? g.todo.map(k => SDEF[k].act) : ["已處理"] } };
  const list = hfUse(all, cols, render);
  hfKeep(ASEL, list);
  const aAll = list.length && list.every(g => ASEL.has(g.order_no));
  $("#sum").innerHTML = `<span><b>${all.length}</b> 張訂單</span><span><b>${fmt(all.reduce((a, g) => a + g.item_count, 0))}</b> 個品項</span><span><b>${fmt(all.reduce((a, g) => a + g.qty, 0))}</b> 件（蝦皮數量）</span>${hfNote(list.length)}`;
  let h = `<thead><tr><th><input type="checkbox" class="ck" id="a-all" ${aAll ? "checked" : ""} ${list.length ? "" : "disabled"} title="全選"></th>${hfTh("no")}${hfTh("line")}${hfTh("date")}${hfTh("up", "opt")}${hfTh("wh")}${hfTh("ff")}${hfTh("codes", "opt")}${hfTh("items", "num opt")}${hfTh("qty", "num")}${hfTh("ship", "num opt")}${hfTh("short", "num opt")}${hfTh("next")}</tr></thead><tbody>`;
  for (const g of list) {
    const k = g.order_no;
    h += `<tr class="og ${g.diff_pending ? "pend" : ""}" data-k="${esc(k)}">
      <td><input type="checkbox" class="ck a-ck" data-k="${esc(k)}" ${ASEL.has(k) ? "checked" : ""}></td>
      <td><span class="fn" style="font-weight:700;color:var(--ink)">${esc(k)}</span>${g.po_id && g.pr_id ? `<div class="kbd">${esc(g.pr_id)}</div>` : !g.po_id ? `<div class="kbd">尚未成 PO</div>` : ""}</td>
      <td class="nw">${esc(g.line)}</td><td class="nw">${esc(g.expected_date)}</td><td class="nw opt">${esc(g.upload_date || "—")}</td><td class="nw">${wh(g.warehouse)}</td>
      ${ffTd(g)}<td class="opt">${codesTd(g)}</td>
      <td class="num opt">${g.item_count}${g.removed_count ? `<div class="kbd neg">移除 ${g.removed_count}</div>` : ""}</td>
      <td class="num"><b>${fmt(g.qty)}</b>${g.qty_original !== g.qty ? `<div class="kbd">原 ${fmt(g.qty_original)}</div>` : ""}</td>
      <td class="num opt ${g.qty_ship == null ? "muted" : ""}">${fmt(g.qty_ship)}</td><td class="num opt ${g.qty_short ? "neg" : g.qty_short == null ? "muted" : ""}">${fmt(g.qty_short)}</td>
      <td>${STAGE === "diff" ? `<button class="btn btn-p btn-sm dd" data-k="${esc(k)}"><i class="bi bi-arrow-left-right"></i> 確認差異</button>` : nextBadges(g)}</td></tr>`;
  }
  if (!list.length) h += all.length ? hfEmpty(13) : emptyRow(13);
  $("#tbl").innerHTML = h + "</tbody>";
  $$("#tbl tr.og").forEach(tr => tr.addEventListener("click", e => { if (e.target.closest("button, a, input, .go-st, .ff-cell")) return; openSide(tr.dataset.k); }));
  $$("#tbl button.dd").forEach(b => b.addEventListener("click", () => { LAST_K = b.dataset.k; openSide(b.dataset.k); openDiff(b.dataset.k); }));
  $$("#tbl .go-st").forEach(b => b.addEventListener("click", () => setStage(b.dataset.s)));
  $$("#tbl input.a-ck").forEach(c => c.addEventListener("change", () => tick(c, ASEL, render)));
  const aa = $("#a-all"); if (aa) aa.addEventListener("change", () => { list.forEach(g => aa.checked ? ASEL.add(g.order_no) : ASEL.delete(g.order_no)); render(); });
  renderBatchBar("#a-bar", ASEL, render);
  markSel();
}
/* ── 清單是動手的地方、右側卡片是看的地方（Jerry 2026-10-21）：
   履約方式就地改（那一格點了跳三個選項，選了就存、不用原因）；勾好幾張，清單下面一條批次改履約方式（每一步都有）、差異待確認多一起確認；
   勾哪一張、改哪一張，右側卡片就切到那張（LAST_K）。Alice 2026-10-21：批次點選確認、改出貨方式不用理由。 ── */
const ASEL = new Set(); let BB_FF = "", BB_REASON = "", LAST_K = "";
const REASON_CHIPS = ["竹運庫存不足", "效期不足", "低於 12 件拒收，刪單", "客戶要求"];
const ffTd = (g, cls = "nw") => `<td class="${cls} ff-cell" data-k="${esc(g.order_no)}" title="點一下改履約方式"><span>${esc(g.fulfil || "竹運出貨")}</span> <i class="bi bi-caret-down-fill ff-i"></i></td>`;
let FF_POP = null;
function ffClose() { if (FF_POP) FF_POP.remove(); FF_POP = null; }
function ffOpen(cell) {
  ffClose(); const k = cell.dataset.k, cur = cell.textContent.trim();
  const pop = document.createElement("div"); pop.id = "ff-pop"; pop.setAttribute("role", "dialog"); pop.setAttribute("aria-label", "改履約方式");
  pop.innerHTML = `<div class="hf-t">${esc(k)}　履約方式改成</div><div class="flex gap-2 flex-wrap">${FULFILS.map(f => `<span class="chip ff-opt ${cur.startsWith(f) ? "on" : ""}" data-f="${esc(f)}">${esc(f)}</span>`).join("")}</div>
    <div class="kbd" style="margin-top:6px">選了就存，不用填原因；歷程會記原本是什麼、改成什麼。</div>`;
  document.body.appendChild(pop); FF_POP = pop;
  const r = cell.getBoundingClientRect();
  pop.style.left = `${Math.max(8, Math.min(r.left, document.documentElement.clientWidth - pop.offsetWidth - 8))}px`; pop.style.top = `${r.bottom + 4}px`;
  pop.querySelectorAll(".ff-opt").forEach(c => c.addEventListener("click", async () => {
    ffClose();
    try {
      const f = c.dataset.f, chk = await ffCheck([k], f); if (!chk) return;
      const d = await post("/api/shopee/fulfil", { order_nos: [k], fulfil: f });
      ffToast(d, f, chk, k); SIDE_NO = k; refresh();
    } catch (err) { toast(err.message, "err"); }
  }));
  LAST_K = k; openSide(k);
}
document.addEventListener("click", e => {
  const c = e.target.closest(".ff-cell");
  if (c) { e.preventDefault(); hfClose(); ffOpen(c); return; }
  if (FF_POP && !e.target.closest("#ff-pop")) ffClose();
});
document.addEventListener("keydown", e => { if (e.key === "Escape" && FF_POP) ffClose(); });
/* 改履約方式前先問系統會動到什麼（已做過的 EIP、產過的拋單檔），有的話跳提醒、講人話（Jerry 2026-10-22「怕他們不知怎麼用」）。
   回 null＝按了「先不要」；回 {orders}＝可以改（沒影響時 orders 是空的，不跳視窗） */
const ffStage = f => f === "竹運出貨" ? "待產拋單檔" : "待產採購單";
const ffMd = t => t ? `${+t.slice(5, 7)}/${+t.slice(8, 10)} ${t.slice(11, 16)}` : "";
async function ffCheck(nos, f) {
  const chk = await post("/api/shopee/fulfil/check", { order_nos: nos, fulfil: f });
  if (!chk.orders.length) return chk;
  const one = x => {
    let h = "", sys = [], todo = [];
    const box = e => `<div class="ff-file">${esc(e.filename)}${e.eip_po ? `（${esc(e.eip_po)}）` : ""}<br>${e.others ? `（這份檔裡還有另外 ${e.others} 張單）` : "（這份檔只有這張單）"}</div>`;
    const open_ = x.eips.filter(e => !e.eip_po), filled = x.eips.filter(e => e.eip_po);
    if (open_.length) h += `<div>這張單已經做過 EIP 採購單${open_.length > 1 ? `（${open_.length} 份）` : ""}：</div>${open_.map(box).join("")}`;
    if (filled.length) h += `<div>這張單的 EIP 採購單已經回填單號 <b>${filled.map(e => esc(e.eip_po)).join("、")}</b>，表示<b>原廠那邊已經下單了</b>：</div>${filled.map(box).join("")}`;
    const n = x.eips.length, whole = x.eips.filter(e => !e.others).length;
    if (n) sys.push(whole === n ? (n > 1 ? `這 ${n} 份 EIP 都只有這張單，整份取消` : "那份 EIP 只有這張單，整份取消")
      : !whole ? `把這張單從${n > 1 ? `這 ${n} 份` : "那份"} EIP 拿掉，其他單照舊、重新做好`
      : "只有這張單的那份整份取消；有其他單的那份把這張拿掉、重新做好");
    if (open_.some(e => e.others)) todo.push("EIP 如果<b>已經上傳到蝦皮 EIP</b>，請到「待回填 EIP 採購單號」重新下載有標「重新做過」的那份，再上傳一次，蓋掉舊的。還沒上傳就不用管。");
    if (open_.some(e => !e.others)) todo.push(`整份取消的 EIP 如果<b>已經上傳到蝦皮 EIP</b>，請到蝦皮 EIP 把它刪掉。還沒上傳就不用管。`);
    if (filled.length) {
      todo.push(`通知採購：到 EIP 系統把這張單的品項從 ${filled.map(e => esc(e.eip_po)).join("、")} 拿掉（或改數量），原廠才不會照舊出貨。`);
      if (filled.some(e => e.direct || e.mars)) todo.push(`直送的單：${filled.some(e => e.mars) ? "瑪氏採購單" : "直送採購單"}請重新下載，交給採購。`);
    }
    if (x.dispatch) {
      h += `<div>這張單已經產過竹運拋單檔（${esc(ffMd(x.dispatch.created_at))} 產出）。</div>`;
      sys.push("取消這張單的拋單檔");
      todo.push("拋單檔如果<b>已經匯進訂單系統</b>，請到訂單系統把這張單刪掉，竹運才不會照舊出貨。");
    }
    sys.push(`這張單移到「${ffStage(f)}」`);
    return `<div class="ff-sec"><div class="ff-no">${esc(x.order_no)}</div>${h}
      <div class="ff-h">按「確定改」之後，系統會自動：</div><ul class="ff-ul">${sys.map(t => `<li>${t}</li>`).join("")}</ul>
      <div class="ff-todo"><div class="ff-h" style="color:var(--warn)">你要做的事</div>${todo.length > 1 ? `<ol class="ff-ul ff-ol">${todo.map(t => `<li>${t}</li>`).join("")}</ol>` : todo.join("")}</div></div>`;
  };
  $("#df-title").textContent = `改成「${f}」？`;
  $("#df-body").innerHTML = (nos.length > 1 ? `<div>勾選的 ${nos.length} 張裡，有 ${chk.orders.length} 張已經做過 EIP 或產過拋單檔，按「確定改」會一起處理：</div>` : "")
    + chk.orders.map(one).join("");
  const dlg = $("#dlg-ff"); dlg.showModal();
  return await new Promise(res => {
    const done = v => { $("#df-ok").onclick = $("#df-no").onclick = dlg.onclose = null; if (dlg.open) dlg.close(); res(v); };
    $("#df-ok").onclick = () => done(chk); $("#df-no").onclick = () => done(null); dlg.onclose = () => done(null);
  });
}
function ffToast(r, f, chk, k) {
  if (!r.count) return toast(k ? "履約方式相同，未修改" : "履約方式都相同，未修改");
  const reup = chk.orders.some(x => x.eips.some(e => !e.eip_po && e.others));
  const msg = `${k ? `${k} ` : `${r.count} 張`}已改成${f}，移到「${ffStage(f)}」。` + (r.handled.length ? `${r.handled.join("；")}。` : "")
    + (reup ? "如果已經上傳到蝦皮 EIP，記得重新上傳。" : "");
  toast(msg, r.handled.length ? "warn" : "");
}
const tick = (c, sel, redraw) => { const k = c.dataset.k; if (c.checked) { sel.add(k); LAST_K = k; openSide(k); } else sel.delete(k); redraw(); };
function renderBatchBar(barSel, sel, redraw) {
  const bar = $(barSel), n = sel.size;
  bar.classList.toggle("hidden", !n);
  if (!n) { BB_FF = ""; return; }
  bar.innerHTML = `<div class="abar"><span>已勾選 <b>${n}</b> 張訂單</span><span class="kbd">履約方式改成：</span>${FULFILS.map(f => `<span class="chip bb-ff ${BB_FF === f ? "on" : ""}" data-f="${esc(f)}">${esc(f)}</span>`).join("")}
    <input class="inp bb-reason" placeholder="原因（選填）" value="${esc(BB_REASON)}" style="width:200px"><button class="btn btn-p btn-sm bb-apply" ${BB_FF ? "" : "disabled"}>套用到 ${n} 張</button>
    ${STAGE === "diff" ? `<span class="kbd">｜</span><button class="btn btn-o btn-sm bb-confirm"><i class="bi bi-check2-all"></i> 勾選的一起確認差異</button>` : ""}
    <a href="#" class="lnk bb-clear">取消勾選</a></div>`;
  bar.querySelectorAll(".bb-ff").forEach(c => c.addEventListener("click", () => { BB_FF = c.dataset.f; renderBatchBar(barSel, sel, redraw); }));
  bar.querySelector(".bb-reason").addEventListener("input", e => { BB_REASON = e.target.value; });
  bar.querySelector(".bb-apply").addEventListener("click", async () => {
    const nos = [...sel]; if (!BB_FF || !nos.length) return;
    try {
      const f = BB_FF, chk = await ffCheck(nos, f); if (!chk) return;
      const r = await post("/api/shopee/fulfil", { order_nos: nos, fulfil: f, reason: BB_REASON.trim() });
      ffToast(r, f, chk, nos.length === 1 ? nos[0] : ""); sel.clear(); BB_FF = ""; BB_REASON = ""; if (LAST_K) SIDE_NO = LAST_K; refresh();
    } catch (err) { toast(err.message, "err"); }
  });
  const bc = bar.querySelector(".bb-confirm"); if (bc) bc.addEventListener("click", async () => {
    const nos = [...sel]; if (!nos.length) return;
    try { const r = await post("/api/shopee/diffs/confirm", { order_nos: nos }); toast(`已確認 ${r.confirmed} 張的差異`); sel.clear(); if (LAST_K) SIDE_NO = LAST_K; refresh(); }
    catch (err) { toast(err.message, "err"); }
  });
  bar.querySelector(".bb-clear").addEventListener("click", e => { e.preventDefault(); sel.clear(); redraw(); });
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
  const ok = new Set(PLAN.orders.filter(g => inStage(g.order_no)).map(g => g.order_no));
  [...BSEL].forEach(k => { if (!ok.has(k)) BSEL.delete(k); });
  renderPlan();
}

/* 「需處理」要寫出卡在哪（Jerry 2026-10-13：「為何產出 EIP 採購單勾不起來」——只寫「3 項需處理」、勾選框反灰，看不出原因）。
   清單上寫原因的摘要；勾選框換成紅色驚嘆號，點這一列右側列出每一品的原因和怎麼修。 */
const WHY = [[/^認不出線別/, "認不出線別"], [/^瑪氏目前只開放/, "瑪氏只開放供應商直送"], [/差異待確認/, "差異待確認"],
  [/^已回填 EIP/, "已回填 EIP 採購單號"], [/^瑪氏商品總表(找不到|有 .*，沒有)/, "瑪氏商品總表找不到"], [/^價格本找不到/, "價格本找不到"],
  [/沒有箱入數/, "沒有箱入數"], [/除不盡/, "除不盡，需改數量"], [/^請先上傳/, "庫存資料未上傳"], [/^竹運商品主檔沒有/, "竹運商品主檔沒有年限"],
  [/已超過目前可用的庫存/, "出貨量超過庫存"], [/^還沒有入庫單號/, "還沒有入庫單號"], [/張入庫單/, "一品有多張入庫單"]];
const whyKeys = items => {         // 表頭篩選「檢查結果」用：一張單卡在哪幾種原因
  const s = new Set(); for (const p of items) if (p.blocking) for (const m of p.problems) s.add((WHY.find(([re]) => re.test(m)) || [0, m])[1]);
  return s.size ? [...s] : ["其他問題"];
};
const NEED = "需先處理";
function whyText(items) {
  const n = new Map();
  for (const p of items) if (p.blocking) for (const m of p.problems) { const w = (WHY.find(([re]) => re.test(m)) || [0, m])[1]; n.set(w, (n.get(w) || 0) + 1); }
  return [...n].map(([w, c]) => `${w} ${c} 項`).join("、");
}
const blockCell = g => `<span class="badge b-diff">${g.blocking} 項需處理</span><div class="kbd" style="color:var(--bad);margin-top:3px">${esc(whyText(g.items))}</div>`;
const BLK_ICON = `<i class="bi bi-exclamation-circle blk-i" title="需先處理：點這一列，右側會列出每一品的原因"></i>`;
const blkHint = n => n ? `有 ${n} 張需先處理：點該列，右側會列出每一品的原因和修正方式。` : "";

const PLAN_FREE = new Set(["可產出", "瑪氏目前只開放供應商直送", "認不出線別"]);
const planChk = g => !g.supported ? [g.line === "瑪氏" ? "瑪氏目前只開放供應商直送" : "認不出線別"] : g.blocking ? whyKeys(g.items) : ["可產出"];
function renderPlan() {
  const every = PLAN.orders.filter(g => inStage(g.order_no));
  const list = hfUse(every, { no: HF_NO, line: HF_LINE, date: HF_DATE, up: HF_UP, wh: HF_WH, ff: HF_FF, codes: HF_CODES, items: hfNum("品項", g => g.items.length), qty: hfNum("蝦皮數量", g => g.qty),
    chk: { label: "檢查結果", get: planChk, first: ["可產出"], group: v => PLAN_FREE.has(v) ? "" : NEED } }, renderPlan);
  hfKeep(BSEL, list);
  const picks = list.filter(canPick), blkList = every.filter(g => g.supported && g.blocking), blk = blkList.length;
  $("#b-sum").innerHTML = `<span><b>${every.length}</b> 張訂單</span>${hfQuick("chk", every.some(canPick) ? ["可產出"] : [], `<b>${every.filter(canPick).length}</b> 張可產出`)}`
    + (blk ? hfQuick("chk", [...new Set(blkList.flatMap(g => whyKeys(g.items)))], `<b style="color:var(--bad)">${blk}</b> 張需先處理`) : "") + hfNote(list.length);
  const all = list.length && list.every(g => BSEL.has(g.order_no));
  let h = `<thead><tr><th><input type="checkbox" class="ck" id="b-all" ${all ? "checked" : ""} ${list.length ? "" : "disabled"} title="全選"></th>
    ${hfTh("no")}${hfTh("line")}${hfTh("date")}${hfTh("up", "opt")}${hfTh("wh")}${hfTh("ff")}${hfTh("codes", "opt")}${hfTh("items", "num opt")}${hfTh("qty", "num")}${hfTh("chk")}</tr></thead><tbody>`;
  for (const g of list) {
    const k = g.order_no, pick = canPick(g);
    const st = !g.supported ? `<span class="badge b-dim">${esc(g.line === "瑪氏" ? "瑪氏目前只開放供應商直送" : "認不出線別")}</span>`
      : g.blocking ? blockCell(g) : `<span class="badge b-ok">可產出</span>`;
    h += `<tr class="og ${g.supported && g.blocking ? "blk" : ""} ${g.supported ? "" : "lock"}" data-k="${esc(k)}">
      <td class="nw"><input type="checkbox" class="ck b-ck" data-k="${esc(k)}" ${BSEL.has(k) ? "checked" : ""}>${!pick && g.supported && g.blocking ? " " + BLK_ICON : ""}</td>
      <td><span class="fn" style="font-weight:700;color:var(--ink)">${esc(k)}</span>${g.po_id && g.pr_id ? `<div class="kbd">${esc(g.pr_id)}</div>` : ""}</td>
      <td class="nw">${esc(g.line)}</td><td class="nw">${esc(g.expected_date)}</td><td class="nw opt">${esc(g.upload_date || "—")}</td><td class="nw">${wh(g.warehouse)}</td>${ffTd(g)}
      <td class="opt">${codesTd(g)}</td><td class="num opt">${g.items.length}</td><td class="num">${fmt(g.qty)}</td><td>${st}</td></tr>`;
  }
  if (!list.length) h += every.length ? hfEmpty(11) : emptyRow(11);
  $("#b-tbl").innerHTML = h + "</tbody>";
  $$("#b-tbl tr.og").forEach(tr => tr.addEventListener("click", e => { if (e.target.closest("button, a, input, .ff-cell")) return; openSide(tr.dataset.k); }));
  $$("#b-tbl input.b-ck").forEach(c => c.addEventListener("change", () => tick(c, BSEL, renderPlan)));
  const ba = $("#b-all"); if (ba) ba.addEventListener("change", () => { list.forEach(g => ba.checked ? BSEL.add(g.order_no) : BSEL.delete(g.order_no)); renderPlan(); });
  markSel();
  renderBatchBar("#b-bar", BSEL, renderPlan);
  const n = BSEL.size, pn = bPicked().length;        // 勾的可以拿去批次改履約方式；產出只算可產出的
  $("#b-go").disabled = !pn;
  $("#b-go-hint").textContent = n ? `已勾選 ${n} 張訂單${pn < n ? `，其中 ${pn} 張可產出（其他要先處理）` : ""}，會產出 EIP 採購單（竹運採購進貨另附採購下採明細）。`
    : picks.length ? `請勾選要產出的訂單。${blkHint(blk)}` : blkHint(blk) || "這一步目前沒有可產出的訂單。";
}
const bPicked = () => [...BSEL].filter(k => { const g = PLAN.orders.find(x => x.order_no === k); return g && canPick(g); });
$("#b-go").addEventListener("click", async () => {
  const nos = bPicked(); if (!nos.length) return;
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
/* 補商品資料：瑪氏新品不在商品總表，直接在這裡補（Jerry 2026-10-14：不用改 Excel、再到瑪氏出貨頁重傳） */
let MP = null, MP_LABEL = null;
async function openMarsProduct(id) {
  let d; try { d = await api(`/api/shopee/mars_product?order_id=${id}`); } catch (err) { return toast(err.message, "err"); }
  if (d.exists) { toast(`商品總表已經有 ${d.code}（${d.unit}），重新整理後即可產出。`); return refresh(); }
  MP = { id, ...d }; const sib = d.siblings[0];
  $("#dm-title").textContent = `補商品資料：${d.code}（${d.unit}）`;
  $("#dm-facts").innerHTML = `<span>永豐料號 <b class="fn">${esc(d.code)}</b></span><span>單位 <b>${esc(d.unit)}</b></span><span>${esc((d.name || "").slice(0, 60))}</span>`;
  $("#dm-sib").innerHTML = sib ? `商品總表已有這個料號的${d.siblings.map(x => esc(x.unit)).join("、")}，已帶入品類、中盒貼標、瑪氏貨號、價格、效期天；<b>箱入數請依${esc(d.unit)}填寫</b>。` : "";
  $("#dm-cat").innerHTML = `<option value="">請選擇</option>` + d.categories.map(c => `<option value="${esc(c.value)}">${esc(c.label)}</option>`).join("");
  $("#dm-cat").value = sib ? (d.categories.find(c => c.value.toLowerCase() === (sib.category || "").toLowerCase()) || {}).value || "" : "";
  $$("#dlg-mp .dm-u").forEach(x => { x.textContent = d.unit; });
  $("#dm-box").value = ""; $("#dm-mc").value = sib ? sib.mars_code || "" : ""; $("#dm-price").value = sib && sib.price != null ? sib.price : "";
  $("#dm-days").value = sib && sib.shelf_days != null ? sib.shelf_days : "";
  MP_LABEL = sib ? ((sib.inner_label || "").toUpperCase() === "V" ? "V" : "") : null;
  $("#dm-lv").classList.toggle("on", MP_LABEL === "V"); $("#dm-ln").classList.toggle("on", MP_LABEL === "");
  $("#dlg-mp").showModal(); $("#dm-cat").focus();
}
["#dm-lv", "#dm-ln"].forEach(s => $(s).addEventListener("click", () => {
  MP_LABEL = $(s).dataset.l; $("#dm-lv").classList.toggle("on", MP_LABEL === "V"); $("#dm-ln").classList.toggle("on", MP_LABEL === "");
}));
$("#dm-save").addEventListener("click", async () => {
  if (!$("#dm-cat").value) return toast("請選擇品類", "err");
  if (!$("#dm-box").value.trim()) return toast("請填寫箱入數", "err");
  if (MP_LABEL === null) return toast("請選擇中盒貼標", "err");
  try {
    await post("/api/shopee/mars_product", { order_id: MP.id, category: $("#dm-cat").value, box_qty: $("#dm-box").value.trim(), inner_label: MP_LABEL,
      mars_code: $("#dm-mc").value.trim(), price: $("#dm-price").value.trim(), shelf_days: $("#dm-days").value.trim() });
    toast(`${MP.code}（${MP.unit}）已補進瑪氏商品總表`); $("#dlg-mp").close(); refresh();
  } catch (err) { toast(err.message, "err"); }
});

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
  const withFilled = $("#e-all").classList.contains("on"), every = eipList(withFilled);
  const cols = { date: HF_DATE, wh: { label: "倉別", get: e => [e.warehouse_name || "—"] }, line: HF_LINE, ff: HF_FF,
    file: { label: "檔名（點選重新下載）", get: e => [e.filename], byValue: true }, rows: hfNum("品項", e => e.rows), qty: hfNum("數量", e => e.qty),
    orders: { label: "蝦皮訂單", get: e => e.orders.length ? e.orders : ["—"], byValue: true },
    by: { label: "產出", get: e => [e.created_by || "—"], sort: e => e.created_at || "", date: true },
    po: { label: "EIP 採購單號", get: e => [e.eip_po || "未回填"], first: ["未回填"] } };
  const list = hfUse(every, cols, renderEips, "個檔");
  $("#e-hf").innerHTML = hfNote(list.length);
  $("#e-mars").classList.toggle("hidden", !EIPS.some(e => e.mars && e.eip_po));
  $("#e-direct").classList.toggle("hidden", !EIPS.some(e => e.direct && e.eip_po));
  $("#e-tbl").innerHTML = `<thead><tr>${hfTh("date")}${hfTh("wh")}${hfTh("line")}${hfTh("ff")}${hfTh("file")}${hfTh("rows", "num opt")}${hfTh("qty", "num")}${hfTh("orders")}${hfTh("by", "opt")}${hfTh("po")}</tr></thead><tbody>
    ${list.map(e => `<tr class="${e.eip_po ? "lock" : ""}"><td class="nw">${esc(e.expected_date)}</td><td class="nw">${esc(e.warehouse_name)}</td><td class="nw">${esc(e.line)}</td><td class="nw opt">${esc(e.fulfil)}</td>
      <td><a class="lnk" href="/api/shopee/eips/${e.id}/file">${esc(e.filename)}</a>${e.note ? `<div class="ff-note">${esc(e.note)}</div>` : ""}${e.mars ? `<div class="mt-1">${e.eip_po
        ? `<a class="btn btn-o btn-sm" href="/api/shopee/eips/${e.id}/mars_po"><i class="bi bi-file-earmark-excel"></i> 瑪氏採購單</a>`
        : `<button class="btn btn-o btn-sm" disabled title="請先回填 EIP 採購單號"><i class="bi bi-file-earmark-excel"></i> 瑪氏採購單</button>`}</div>` : ""}${e.direct ? `<div class="mt-1">${e.eip_po
        ? `<a class="btn btn-o btn-sm" href="/api/shopee/eips/${e.id}/direct_po"><i class="bi bi-file-earmark-excel"></i> 直送採購單</a>`
        : `<button class="btn btn-o btn-sm" disabled title="請先回填 EIP 採購單號"><i class="bi bi-file-earmark-excel"></i> 直送採購單</button>`}</div>` : ""}</td><td class="num opt">${e.rows}</td><td class="num">${fmt(e.qty)}</td>
      <td class="fn" title="${esc(e.orders.join("\n"))}">${esc(e.orders.slice(0, 2).join("、"))}${e.orders.length > 2 ? `…等 ${e.orders.length} 張` : ""}</td>
      <td class="kbd nw opt">${esc(e.created_by)}<div>${esc((e.created_at || "").slice(5, 16).replace("T", " "))}</div></td>
      <td class="nw"><input class="inp fn e-po" data-id="${e.id}" value="${esc(e.eip_po)}" placeholder="例：PO202610001" style="width:140px">
        <button class="btn ${e.eip_po ? "btn-o" : "btn-p"} btn-sm e-save" data-id="${e.id}">儲存</button>${e.eip_po ? `<div class="kbd">${esc(e.updated_by)} ${esc((e.updated_at || "").slice(5, 16).replace("T", " "))}</div>` : ""}
        <div class="mt-1"><button class="btn btn-o btn-sm e-cancel" data-id="${e.id}" ${e.eip_po ? `disabled title="已回填 EIP 採購單號，請先清掉單號再取消產出"` : `title="這份檔刪掉，裡面的訂單退回「待產採購單」"`}><i class="bi bi-arrow-counterclockwise"></i> 取消產出</button></div></td></tr>`).join("")
      || (every.length ? hfEmpty(10) : `<tr><td colspan="10" class="muted" style="padding:24px;text-align:center">目前沒有待回填的 EIP 採購單。</td></tr>`)}</tbody>`;
  $$("#e-tbl button.e-save").forEach(b => b.addEventListener("click", () => saveEip(b.dataset.id)));
  $$("#e-tbl input.e-po").forEach(i => i.addEventListener("keydown", e => { if (e.key === "Enter") saveEip(i.dataset.id); }));
  $$("#e-tbl button.e-cancel").forEach(b => b.addEventListener("click", () => cancelEip(b.dataset.id, EIPS.find(x => String(x.id) === b.dataset.id))));
}
async function dlZip(url, fallback, what) {
  const res = await fetch(`${url}?${new URLSearchParams(fq())}`);
  if (!res.ok) { let d = {}; try { d = await res.json(); } catch (e) {} return toast(d.error || `伺服器錯誤（${res.status}）`, "err"); }
  const blob = await res.blob(); const m = (res.headers.get("Content-Disposition") || "").match(/filename\*=UTF-8''([^;]+)/);
  const a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = m ? decodeURIComponent(m[1]) : fallback; document.body.appendChild(a); a.click(); a.remove();
  toast(`已下載 ${res.headers.get("X-Shopee-Files")} 份${what}`);
}
$("#e-mars").addEventListener("click", () => dlZip("/api/shopee/mars_po.zip", "蝦皮特選瑪氏採購單.zip", "瑪氏採購單"));
$("#e-direct").addEventListener("click", () => dlZip("/api/shopee/direct_po.zip", "蝦皮特選直送採購單.zip", "直送採購單"));   // 給採購轉原廠（Alice 2026-10-21）

/* ── 蝦皮倉資料（瑪氏採購單表頭 C6～F8） ─────────────────────── */
const WH_COLS = [["name", "入倉倉別", 230], ["address", "地址", 360], ["phone", "電話", 165], ["ship_to", "Ship-to", 100], ["contact", "送貨聯絡人", 120], ["notes", "進倉說明（直送採購單用）", 420]];
async function loadWh() {
  let d; try { d = await api("/api/shopee/warehouses"); } catch (err) { $("#wh-tbl").innerHTML = `<tr><td class="muted">${esc(err.message)}</td></tr>`; return; }
  $("#wh-tbl").innerHTML = `<thead><tr><th>倉別</th>${WH_COLS.map(c => `<th>${c[1]}</th>`).join("")}<th></th></tr></thead><tbody>
    ${d.warehouses.map(w => `<tr data-code="${esc(w.code)}"><td class="nw"><b>${esc(w.short)}</b> <span class="kbd">${esc(w.code)}</span></td>
      ${WH_COLS.map(c => c[0] === "notes" ? `<td><textarea class="inp wh-f" data-k="notes" rows="4" style="width:${c[2]}px;font-size:12px">${esc(w.notes)}</textarea></td>`
        : `<td><input class="inp wh-f" data-k="${c[0]}" value="${esc(w[c[0]])}" style="width:${c[2]}px"></td>`).join("")}
      <td><button class="btn btn-o btn-sm wh-save">儲存</button></td></tr>`).join("")}</tbody>`;
  $$("#wh-tbl button.wh-save").forEach(b => b.addEventListener("click", () => saveWh(b.closest("tr"))));
}
async function saveWh(tr) {
  const body = {}; tr.querySelectorAll(".wh-f").forEach(i => { body[i.dataset.k] = i.value.trim(); });
  try { const r = await api(`/api/shopee/warehouses/${tr.dataset.code}`, { method: "PUT", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) });
    toast(r.changed ? `已更新 ${tr.querySelector("b").textContent} 的倉資料` : "內容沒有變更");
  } catch (err) { toast(err.message, "err"); }
}
$("#wh-box").addEventListener("toggle", () => { if ($("#wh-box").open) loadWh(); });

/* 取消產出 EIP（Jerry 2026-10-21「產出 EIP 之後反悔，回到待產採購單」）：跟拋單檔的「取消產出」同一種做法 */
async function cancelEip(id, e) {
  const n = e && e.orders ? `檔裡的 ${e.orders.length} 張單` : "檔裡的訂單";
  if (!confirm(`取消這份 EIP 的產出？\n${e ? e.filename : ""}\n${n}會回到「待產採購單」。已上傳到蝦皮 EIP 的，請自行到蝦皮撤回。`)) return;
  try { const r = await api(`/api/shopee/eips/${id}`, { method: "DELETE" }); toast(`已取消產出，${r.orders.length} 張單回到「待產採購單」`); refresh(); }
  catch (err) { toast(err.message, "err"); }
}
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
  const st = d.settings; $("#s-rule").value = st.rule; $("#s-ratio").value = st.ratio;
  $("#s-ratio-box").classList.toggle("hidden", st.rule !== "custom");
}
async function saveShipSet(body) {
  try { await api("/api/shopee/ship/settings", { method: "PUT", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) }); loadStock(); toast("效期標準已更新，竹運出貨與勇信核對會依新標準判斷"); }
  catch (err) { toast(err.message, "err"); loadStock(); }
}
$("#s-rule").addEventListener("change", () => saveShipSet({ rule: $("#s-rule").value }));
$("#s-ratio").addEventListener("change", () => saveShipSet({ ratio: $("#s-ratio").value.trim() }));

/* ── 竹運出貨：訂單、出貨量、效期批次 ─────────────────────── */
let SHIP = { orders: [] }; const ZOPEN = new Set(), ZSEL = new Set();
const zq = () => fq();
const zPick = g => g.dispatch_status !== "已產出" && !g.blocking;
async function loadShip() {
  try { SHIP = await api(`/api/shopee/ship/plan?${new URLSearchParams(zq())}`); SHIP.q = JSON.stringify(zq()); } catch (err) { $("#z-tbl").innerHTML = `<tr><td class="muted" style="padding:20px">${esc(err.message)}</td></tr>`; return; }
  const ok = new Set(SHIP.orders.filter(g => zPick(g) && inStage(g.order_no)).map(g => g.order_no));
  [...ZSEL].forEach(k => { if (!ok.has(k)) ZSEL.delete(k); });
  renderShip();
}
const SHIP_FREE = new Set(["已產出拋單檔", "可全數出貨", "部分缺貨", "缺貨（有刪單）"]);
const shipChk = g => g.dispatch_status === "已產出" ? ["已產出拋單檔"] : g.blocking ? whyKeys(g.items) : [g.short ? (g.deleted ? "缺貨（有刪單）" : "部分缺貨") : "可全數出貨"];
function renderShip() {
  const every = SHIP.orders.filter(g => inStage(g.order_no));
  const list = hfUse(every, { no: HF_NO, line: HF_LINE, date: HF_DATE, up: HF_UP, wh: HF_WH, ff: { label: "履約方式", get: g => [g.fulfil || "竹運出貨"] }, codes: HF_CODES, items: hfNum("品項", g => g.items.length), qty: hfNum("訂購量", g => g.qty),
    ship: hfNum("出貨量", g => g.ship), short: hfNum("缺貨量", g => g.short),
    chk: { label: "檢查結果", get: shipChk, first: ["可全數出貨", "部分缺貨", "缺貨（有刪單）"], group: v => SHIP_FREE.has(v) ? "" : NEED },
    disp: { label: "拋單檔", get: g => [g.dispatch_status] } }, renderShip);
  hfKeep(ZSEL, list);
  const picks = list.filter(zPick), open_ = list.filter(g => g.dispatch_status !== "已產出"), all = open_.length && open_.every(g => ZSEL.has(g.order_no));
  const blkList = every.filter(g => g.blocking && g.dispatch_status !== "已產出");
  $("#z-sum").innerHTML = `<span><b>${every.length}</b> 張訂單</span><span><b>${fmt(every.reduce((a, g) => a + g.qty, 0))}</b> 訂購</span>`
    + `<span><b>${fmt(every.reduce((a, g) => a + g.ship, 0))}</b> 可出貨</span><span><b style="color:var(--bad)">${fmt(every.reduce((a, g) => a + g.short, 0))}</b> 缺貨</span>`
    + (blkList.length ? hfQuick("chk", [...new Set(blkList.flatMap(g => whyKeys(g.items)))], `<b style="color:var(--bad)">${blkList.length}</b> 張需先處理`) : "") + hfNote(list.length);
  let h = `<thead><tr><th><input type="checkbox" class="ck" id="z-all" ${all ? "checked" : ""} ${open_.length ? "" : "disabled"} title="全選還沒產出的訂單"></th>
    ${hfTh("no")}${hfTh("line")}${hfTh("date")}${hfTh("up", "opt")}${hfTh("wh")}${hfTh("ff")}${hfTh("codes", "opt")}${hfTh("items", "num")}${hfTh("qty", "num")}${hfTh("ship", "num")}${hfTh("short", "num")}${hfTh("chk")}${hfTh("disp")}</tr></thead><tbody>`;
  for (const g of list) {
    const k = g.order_no, done = g.dispatch_status === "已產出";
    const st = done ? `<span class="badge b-dim">—</span>` : g.blocking ? blockCell(g)
      : `<span class="badge ${g.short ? "b-warn" : "b-ok"}">${g.short ? (g.deleted ? `缺貨（刪單 ${g.deleted} 品）` : "部分缺貨") : "可全數出貨"}</span>${g.warnings ? ` <span class="badge b-warn">${g.warnings} 項提醒</span>` : ""}`;
    h += `<tr class="og ${g.blocking && !done ? "blk" : ""} ${done ? "lock" : ""}" data-k="${esc(k)}">
      <td class="nw"><input type="checkbox" class="ck z-ck" data-k="${esc(k)}" ${ZSEL.has(k) ? "checked" : ""} ${done ? "disabled" : ""}>${!done && g.blocking ? " " + BLK_ICON : ""}</td>
      <td><span class="fn" style="font-weight:700;color:var(--ink)">${esc(k)}</span>${g.po_id && g.pr_id ? `<div class="kbd">${esc(g.pr_id)}</div>` : ""}</td>
      <td class="nw">${esc(g.line)}</td><td class="nw">${esc(g.expected_date)}</td><td class="nw opt">${esc(g.upload_date || "—")}</td><td class="nw">${wh(g.warehouse)}</td>${ffTd(g)}<td class="opt">${codesTd(g)}</td><td class="num">${g.items.length}</td>
      <td class="num">${fmt(g.qty)}</td><td class="num"><b>${fmt(g.ship)}</b></td><td class="num ${g.short ? "neg" : ""}">${fmt(g.short)}</td><td>${st}</td>
      <td class="nw">${done ? `<span class="badge b-s">已產出</span> <span class="kbd">${esc((g.dispatch_at || "").slice(5, 16).replace("T", " "))}</span> <button class="btn btn-o btn-sm z-undo" data-k="${esc(k)}">取消產出</button>`
        : `<span class="badge b-dim">${esc(g.dispatch_status)}</span>`}</td></tr>`;
  }
  if (!list.length) h += every.length ? hfEmpty(14) : emptyRow(14);
  $("#z-tbl").innerHTML = h + "</tbody>";
  $$("#z-tbl tr.og").forEach(tr => tr.addEventListener("click", e => { if (e.target.closest("button, a, input, .ff-cell")) return; openSide(tr.dataset.k); }));
  $$("#z-tbl input.z-ck").forEach(c => c.addEventListener("change", () => tick(c, ZSEL, renderShip)));
  const za = $("#z-all"); if (za) za.addEventListener("change", () => { open_.forEach(g => za.checked ? ZSEL.add(g.order_no) : ZSEL.delete(g.order_no)); renderShip(); });
  markSel();
  $$("#z-tbl button.z-undo").forEach(b => b.addEventListener("click", () => { LAST_K = b.dataset.k; openSide(b.dataset.k); undoShip(b.dataset.k); }));
  renderBatchBar("#z-bar", ZSEL, renderShip);
  const n = ZSEL.size, pn = zPicked().length; $("#z-go").disabled = !pn;
  const zb = blkList.length;
  $("#z-go-hint").textContent = n ? `已勾選 ${n} 張訂單${pn < n ? `，其中 ${pn} 張可產出（其他要先處理）` : ""}，會產出訂單系統拋單檔（一天一倉一個檔）＋下修刪單清單（zip）。`
    : picks.length ? `請勾選要產出的訂單。${blkHint(zb)}` : blkHint(zb) || "這一步目前沒有可產出的訂單。";
}
const zPicked = () => [...ZSEL].filter(k => { const g = SHIP.orders.find(x => x.order_no === k); return g && zPick(g); });
let SHIPX = null;
// 檔案是多少、到貨日較早的幾張單先用了多少（Jerry 2026-10-20：EC 可用 231，訂單系統卻是 315，不講清楚會以為是單位換算）
const usedNote = (file, used, n) => used ? `<span class="kbd">（檔案 ${fmt(file)}，到貨日較早的 ${n} 張單先用 ${fmt(used)}）</span>` : "";
function openShip(id) {
  let x = null; for (const g of [SIDE && SIDE.ship, ...SHIP.orders].filter(Boolean)) { x = g.items.find(i => i.order_id === id); if (x) break; }
  if (!x) return; SHIPX = x;
  $("#ds-title").textContent = `${x.supplier_sku_id} 效期批次與出貨量`;
  $("#ds-facts").innerHTML = `<span>${esc((x.name || "").slice(0, 46))}</span><span>訂購量 <b>${fmt(x.qty)}</b> ${esc(x.unit)}</span><span>EC 可用 <b>${fmt(x.ec)}</b>${x.ec < x.qty ? `（不足 ${fmt(x.qty - x.ec)}）` : "（足夠）"}${usedNote(x.ec_file, x.ec_used, x.ec_users)}</span>
    <span>竹運帳面 <b>${fmt(x.book)}</b>${usedNote(x.book_file, x.book_used, x.book_users)}</span><span>符合效期 <b>${fmt(x.qualified)}</b></span><span>系統建議 <b>${fmt(x.suggest)}</b></span>
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
  $("#ds-who").textContent = x.adj ? `上次修改：${x.adj.operator}　${(x.adj.updated_at || "").slice(0, 16).replace("T", " ")}` : "";
  $("#ds-clear").classList.toggle("hidden", !x.adj); dsSum(); $("#dlg-ship").showModal();
}
const dsTotal = () => $$("#ds-lots .ds-q").reduce((a, q) => a + (Number(q.value) || 0), 0);
function dsSum() { $("#ds-total").textContent = fmt(dsTotal()); }
$("#ds-zero").addEventListener("click", () => { $$("#ds-lots .ds-q").forEach(q => { q.value = ""; }); $$("#ds-lots .ds-ck").forEach(c => { c.checked = false; }); dsSum(); });
$("#ds-save").addEventListener("click", async () => {
  const lots = $$("#ds-lots .ds-q").filter(q => Number(q.value) > 0).map(q => ({ exp: q.dataset.exp, qty: Number(q.value) }));
  try { const r = await post("/api/shopee/ship/adjust", { order_id: SHIPX.order_id, lots, reason: "", ...zq() });
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
  const nos = zPicked(); if (!nos.length) return;
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
  const every = SHORT.orders.filter(g => inStage(g.order_no));
  const list = hfUse(every, { no: HF_NO, line: HF_LINE, ff: HF_FF, date: HF_DATE, up: HF_UP, wh: HF_WH, codes: HF_CODES, items: hfNum("品項", g => g.items.length),
    orig: hfNum("原始訂購量", g => g.orig), final: hfNum("最終可出貨量", g => g.final), cut: hfNum("累計下修", g => g.cut),
    st: { label: "通知狀態", get: g => [g.status] } }, renderShort);
  hfKeep(KSEL, list);
  const picks = list.filter(g => g.pending), all = list.length && list.every(g => KSEL.has(g.order_no));
  $("#k-sum").innerHTML = `<span><b>${every.length}</b> 張訂單</span><span><b>${every.filter(g => g.short_items).length}</b> 張有缺貨</span>`
    + `<span><b style="color:var(--warn)">${every.reduce((a, g) => a + g.pending, 0)}</b> 品待通知</span><span><b style="color:var(--bad)">${fmt(every.reduce((a, g) => a + g.cut, 0))}</b> 累計下修</span>${hfNote(list.length)}`;
  let h = `<thead><tr><th><input type="checkbox" class="ck" id="k-all" ${all ? "checked" : ""} ${list.length ? "" : "disabled"} title="全選"></th>
    ${hfTh("no")}${hfTh("line")}${hfTh("ff")}${hfTh("date")}${hfTh("up", "opt")}${hfTh("wh")}${hfTh("codes", "opt")}${hfTh("items", "num opt")}${hfTh("orig", "num opt")}${hfTh("final", "num")}${hfTh("cut", "num")}${hfTh("st")}</tr></thead><tbody>`;
  for (const g of list) {
    const k = g.order_no;
    h += `<tr class="og ${g.pending ? "pend" : ""}" data-k="${esc(k)}">
      <td><input type="checkbox" class="ck k-ck" data-k="${esc(k)}" ${KSEL.has(k) ? "checked" : ""}></td>
      <td><span class="fn" style="font-weight:700;color:var(--ink)">${esc(k)}</span>${g.po_id && g.pr_id ? `<div class="kbd">${esc(g.pr_id)}</div>` : ""}</td>
      <td class="nw">${esc(g.line)}</td>${ffTd(g)}<td class="nw">${esc(g.expected_date)}</td><td class="nw opt">${esc(g.upload_date || "—")}</td><td class="nw">${wh(g.warehouse)}</td><td class="opt">${codesTd(g)}</td><td class="num opt">${g.items.length}</td>
      <td class="num opt">${fmt(g.orig)}</td><td class="num"><b>${fmt(g.final)}</b></td><td class="num ${g.cut ? "neg" : ""}">${fmt(g.cut)}${g.deleted ? `<div class="kbd">刪單 ${g.deleted} 品</div>` : ""}</td>
      <td><span class="badge ${g.status === "待通知" ? "b-warn" : g.status.startsWith("已通知") ? "b-ok" : "b-dim"}">${esc(g.status)}</span></td></tr>`;
  }
  if (!list.length) h += every.length ? hfEmpty(13) : emptyRow(13);
  $("#k-tbl").innerHTML = h + "</tbody>";
  $$("#k-tbl tr.og").forEach(tr => tr.addEventListener("click", e => { if (e.target.closest("button, a, input, .ff-cell")) return; openSide(tr.dataset.k); }));
  $$("#k-tbl input.k-ck").forEach(c => c.addEventListener("change", () => tick(c, KSEL, renderShort)));
  const ka = $("#k-all"); if (ka) ka.addEventListener("change", () => { list.forEach(g => ka.checked ? KSEL.add(g.order_no) : KSEL.delete(g.order_no)); renderShort(); });
  markSel();
  renderBatchBar("#k-bar", KSEL, renderShort);
  const n = KSEL.size, pn = kPicked().length; $("#k-go").disabled = !pn;
  $("#k-go-hint").textContent = n ? `已勾選 ${n} 張訂單${pn < n ? `，其中 ${pn} 張有待通知的品項` : ""}，會把待通知的品項整理成一則 LINE 訊息，並鎖定這次的下修紀錄。` : "請勾選有「待通知」的訂單。";
}
const kPicked = () => [...KSEL].filter(k => { const g = SHORT.orders.find(x => x.order_no === k); return g && g.pending; });
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
  const nos = kPicked(); if (!nos.length) return;
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
  // 清單剛算好同一組篩選的拋單檔計畫就直接用，不叫伺服器再算一次（算一次要把範圍內每張單都扣過庫存，搜尋才會慢）
  const have = SHIP.q === JSON.stringify(fq()) ? SHIP.orders.find(g => g.order_no === no) : null;
  let d; try { d = await api(`/api/shopee/order?${new URLSearchParams({ ...fq(), no, ...(have ? { ship: "0" } : {}) })}`); }
  catch (err) { if (SIDE_NO === no) $("#sd-body").innerHTML = `<div class="alert al-bad">${esc(err.message)}</div>`; return; }
  if (SIDE_NO !== no) return;                      // 讀的時候又點了別張
  if (have) d.ship = have;
  const top = $("#sd-body").scrollTop; SIDE = d; renderSide(); if (again && !$("#f-q").value.trim()) $("#sd-body").scrollTop = top;
}
const byId = grp => new Map(((grp || {}).items || []).map(x => [x.order_id, x]));
function renderSide() {
  const d = SIDE, g = d.order, buy = byId(d.buy), ship = byId(d.ship), sh = byId(d.short);
  const pendDiffs = d.diffs.filter(x => !x.confirmed);
  const under12 = i => { const z = ship.get(i.id); return !!(z && !z.dispatch && z.final > 0 && z.final < 12); };   // 少於 12 件蝦皮會拒收：改成 0 放這裡（Jerry 2026-10-21）
  const under = g.items.filter(under12);
  $("#sd-title").innerHTML = `<b class="fn" style="font-size:15px;color:var(--ink)">${esc(g.order_no)}</b>${g.po_id && g.pr_id ? `<div class="kbd fn">${esc(g.pr_id)}</div>` : ""}
    <div class="facts" style="margin-top:4px"><span>${esc(g.line)}</span><span>${wh(g.warehouse)}</span><span>到貨 <b>${esc(g.expected_date)}</b></span>
      <span>履約方式 <b>${esc(g.fulfil)}</b> <span class="kbd">（在清單那一格點一下可改）</span></span></div>
    <div class="mt-1">${nextBadges(g)}${under.length ? ` <button class="btn btn-o btn-sm sd-zero-all" data-ids="${esc(under.map(i => i.id).join(","))}"><i class="bi bi-x-circle"></i> 這張少於 12 件的 ${under.length} 項改成 0（刪單）</button>` : ""}</div>`;
  let h = "";
  if (pendDiffs.length) h += `<div class="alert al-bad"><b><i class="bi bi-exclamation-triangle-fill"></i> 蝦皮重傳後有 ${pendDiffs.length} 處不同</b>，確認後才能往下處理：
      <div class="mt-1">${pendDiffs.map(x => `<div>${esc(x.supplier_sku_id)}　${esc(x.field_label)}：${esc(x.old_value)} → <b>${esc(x.new_value)}</b>${x.qty_delta ? `（${delta(x.qty_delta)}）` : ""}</div>`).join("")}</div>
      <button class="btn btn-p btn-sm mt-2 sd-dconf"><i class="bi bi-check2-circle"></i> 確認差異</button></div>`;

  // 品項：每一品把各步的資料放在一起，該改的按鈕就在旁邊
  // 需處理的品項排最前面（瑪氏一張單常常四十幾品，擋下來的那幾品不能讓人一品一品捲著找）
  const isBad = i => { const p = buy.get(i.id), z = ship.get(i.id); return !!((g.todo.includes("buy") && p && p.blocking) || (z && z.blocking && !z.dispatch)); };
  // 左邊搜尋框有字：符合的品項排最前面並標起來（Jerry 2026-10-20：搜一個料號，右邊 28 品還要自己捲著找）
  const q = $("#f-q").value.trim().toLowerCase();
  const isHit = i => !!q && [i.supplier_sku_id, i.sku_name, i.shopee_sku_id, i.inbound_ids, i.barcode, i.base_code].some(v => String(v || "").toLowerCase().includes(q));
  const rank = i => (isHit(i) ? 0 : 2) + (isBad(i) ? 0 : 1);
  const items = g.items.map((i, k) => [i, k]).sort((a, b) => rank(a[0]) - rank(b[0]) || a[1] - b[1]).map(x => x[0]);
  const nBad = g.items.filter(isBad).length, nHit = g.items.filter(isHit).length;
  h += `<div class="sd-sec"><div class="h"><i class="bi bi-box-seam"></i> 品項（${g.item_count}${g.removed_count ? `，移除 ${g.removed_count}` : ""}）${nHit
    ? `<span style="font-weight:400;font-size:13px;color:var(--s9)">　符合「${esc($("#f-q").value.trim())}」的 ${nHit} 品排在最前面</span>` : ""}${nBad
    ? `<span class="neg" style="font-weight:400;font-size:13px">　${nBad} 項需處理${nHit ? "" : "，排在最前面"}</span>` : ""}</div><div class="b">
    ${g.line === "瑪氏" && d.short ? `<div class="kbd" style="padding:6px 0">瑪氏的缺貨依勇信配送明細表核對，PDF 請拖曳至上方「上傳檔案」。</div>` : ""}`;
  for (const i of items) {
    const p = buy.get(i.id), z = ship.get(i.id), k = sh.get(i.id), unit = i.unit_name || i.unit_code;
    const kv = [`原始訂購量 <b>${fmt(i.qty_original)}</b>`, i.inbound_ids ? `入庫單號 <span class="fn">${esc(i.inbound_ids)}</span>` : "",
      i.version > 1 ? `<a class="lnk ver" data-id="${i.id}" data-t="${esc(i.supplier_sku_id)}">蝦皮重傳過 ${i.version - 1} 次，查看各版本</a>` : ""];
    let more = "", acts = "";
    if (p) {
      more += `<div class="sd-kv"><span>轉換率 <b>${fmt(p.conv)}</b></span><span>下採數量 <b>${fmt(p.use_qty)}</b>${p.adj ? "（已修改）" : ""}</span>
        <span>EIP 採購數量 <b>${p.out_qty == null ? "—" : fmt(p.out_qty)}</b> ${esc(p.out_unit || "")}${p.cases != null && p.out_unit !== "箱" ? `（${fmt(p.cases)} 箱）` : ""}</span>${p.unit_group ? `<span>EIP 檔：${esc(p.unit_group)}</span>` : ""}</div>`;
      if (g.todo.includes("buy")) more += p.problems.map(x => `<div class="prob"><i class="bi bi-exclamation-circle"></i> ${esc(x)}</div>`).join("");
      if (g.todo.includes("buy") && p.problems.some(x => /^瑪氏商品總表(找不到|有 )/.test(x))) acts += `<button class="btn btn-p btn-sm sd-mp" data-id="${i.id}"><i class="bi bi-plus-square"></i> 補商品資料</button>`;   // 已產出的單不再列產出前的檢查
      if ((p.rem || p.adj) && !(p.eip && p.eip.eip_po)) acts += `<button class="btn ${p.rem ? "btn-p" : "btn-o"} btn-sm sd-adj" data-id="${i.id}"><i class="bi bi-pencil"></i> 修改數量</button>`;
    }
    if (z) {
      more += `<div class="sd-kv">${z.dispatch ? "" : `<span>EC 可用量 <b>${fmt(z.ec)}</b>${usedNote(z.ec_file, z.ec_used, z.ec_users)}</span><span>符合效期量 <b>${fmt(z.qualified)}</b></span>`}<span>出貨量 <b>${fmt(z.final)}</b>${z.adj ? "（已修改）" : ""}</span>
        ${z.short ? `<span class="neg">缺貨 ${fmt(z.short)}${z.reasons.length ? `（${esc(z.reasons.join("、"))}）` : ""}</span>` : ""}${z.alloc.length ? `<span>出貨效期 ${z.alloc.map(a => `${esc(a.exp)}（${fmt(a.qty)}）`).join("、")}</span>` : ""}</div>`;
      more += z.problems.map(x => `<div class="prob"><i class="bi bi-exclamation-circle"></i> ${esc(x)}</div>`).join("") + z.warnings.map(x => `<div class="kbd" style="color:var(--warn)"><i class="bi bi-info-circle"></i> ${esc(x)}</div>`).join("");
      if (!z.dispatch && z.life && z.lots.length) acts += `<button class="btn ${z.warnings.length ? "btn-p" : "btn-o"} btn-sm sd-ship" data-id="${i.id}"><i class="bi bi-box-seam"></i> 效期批次／改出貨量</button>`;
      if (under12(i)) acts += `<button class="btn btn-o btn-sm sd-zero" data-ids="${i.id}"><i class="bi bi-x-circle"></i> 改成 0（刪單）</button>`;
    }
    if (k) {
      if (k.rounds.length) more += `<div class="sd-kv"><span>可出貨量 <b>${fmt(k.final)}</b></span><span class="neg">累計下修 ${fmt(k.cut)}</span>
        ${k.pending ? `<span class="badge b-warn">待通知</span>` : k.notified_round ? `<span class="badge b-ok">已通知第 ${k.notified_round} 次</span>` : ""}</div>
        <div class="kbd">${k.rounds.map(r => `${roundText(r)}（${esc(r.source)}${r.reasons.length ? "，" + esc(r.reasons.join("、")) : ""}）`).join("<br>")}</div>`;
      if (k.edit_round) acts += `<button class="btn ${k.pending ? "btn-p" : "btn-o"} btn-sm sd-short" data-id="${i.id}"><i class="bi bi-dash-circle"></i> ${k.rounds.length && !k.pending ? `填第 ${k.edit_round} 次缺貨` : k.pending ? "修改缺貨" : "填缺貨"}</button>`;
    }
    h += `<div class="sd-it ${i.removed ? "muted" : ""}${isHit(i) ? " hit" : ""}"><div class="r"><span class="fn" style="font-weight:700;color:var(--ink)" title="永豐料號（蝦皮採購單的 Supplier SKU ID）"><span class="kbd" style="font-weight:400">永豐料號</span> ${esc(i.supplier_sku_id)}</span>
        <span style="flex:1;min-width:0">${esc((i.sku_name || "").slice(0, 40))}</span><span class="nw"><b style="font-size:15px">${fmt(i.qty)}</b> ${esc(unit)}</span></div>
      <div class="sd-kv">${kv.filter(Boolean).map(x => `<span>${x}</span>`).join("")}${i.removed ? `<span class="badge b-diff">已移除</span>` : i.diff_pending ? `<span class="badge b-diff">差異待確認</span>` : ""}</div>
      ${more}${acts ? `<div class="sd-acts">${acts}</div>` : ""}</div>`;
  }
  h += `</div></div>`;

  // EIP 採購單（竹運採購進貨、直送）
  if (d.buy) h += `<div class="sd-sec"><div class="h"><i class="bi bi-file-earmark-spreadsheet"></i> EIP 採購單 <span class="kbd">${d.eips.length ? `${d.eips.length} 份` : "尚未產出"}</span></div><div class="b">
    ${d.eips.map(e => `<div class="sd-it"><a class="lnk" href="/api/shopee/eips/${e.id}/file">${esc(e.filename)}</a>${e.note ? `<div class="ff-note">${esc(e.note)}</div>` : ""}
      <div class="flex items-center gap-2 mt-1 flex-wrap"><input class="inp fn sd-po" data-id="${e.id}" value="${esc(e.eip_po)}" placeholder="EIP 採購單號，例：PO202610001" style="width:220px">
        <button class="btn ${e.eip_po ? "btn-o" : "btn-p"} btn-sm sd-po-save" data-id="${e.id}">儲存單號</button>
        ${e.mars ? (e.eip_po ? `<a class="btn btn-o btn-sm" href="/api/shopee/eips/${e.id}/mars_po"><i class="bi bi-file-earmark-excel"></i> 下載瑪氏採購單</a>` : `<span class="kbd">回填後可下載瑪氏採購單</span>`) : ""}
        ${e.direct ? (e.eip_po ? `<a class="btn btn-o btn-sm" href="/api/shopee/eips/${e.id}/direct_po"><i class="bi bi-file-earmark-excel"></i> 下載直送採購單</a>` : `<span class="kbd">回填後可下載直送採購單（給採購轉原廠）</span>`) : ""}
        <button class="btn btn-o btn-sm sd-eip-cancel" data-id="${e.id}" ${e.eip_po ? `disabled title="已回填 EIP 採購單號，請先清掉單號再取消產出"` : `title="這份檔刪掉，裡面的訂單退回「待產採購單」"`}><i class="bi bi-arrow-counterclockwise"></i> 取消產出</button></div>
      ${e.eip_po ? `<div class="kbd">${esc(e.updated_by)} ${esc((e.updated_at || "").slice(5, 16).replace("T", " "))} 回填</div>` : ""}</div>`).join("")
      || `<div class="kbd">尚未產出。</div>`}
    ${g.todo.includes("buy") ? (d.buy.blocking ? `<div class="kbd neg mt-1">先修正上面標紅的品項，才能產出 EIP 採購單。</div>`
        : `<div class="sd-acts"><button class="btn btn-p btn-sm sd-regen"><i class="bi bi-file-earmark-zip"></i> 產出這張單的 EIP 採購單</button></div>`)
      : d.eips.some(e => !e.eip_po) ? `<div class="sd-acts"><button class="btn btn-o btn-sm sd-regen"><i class="bi bi-arrow-repeat"></i> 重新產出 EIP 採購單</button><span class="kbd">改過數量後用；還沒回填單號的那份會換掉。</span></div>` : ""}</div></div>`;

  // 竹運出貨：拋單檔
  // 調整原因（Jerry 2026-10-21：統一放在右側卡片、拋單檔上面一區）：上面一格填一次套用到這張單全部改過的品項，
  // 只有一品不一樣就改那一品那格；之後再改數量的自動帶這張單的原因（後端 _order_reason）
  const adjd = d.ship ? d.ship.items.filter(x => x.adj && !x.dispatch) : [];
  if (adjd.length) {
    const cnt = {}; adjd.forEach(x => { if (x.adj.reason) cnt[x.adj.reason] = (cnt[x.adj.reason] || 0) + 1; });
    const common = Object.entries(cnt).sort((a, b) => b[1] - a[1])[0], missing = adjd.filter(x => !x.adj.reason).length;
    h += `<div class="sd-sec" id="sd-rsn"><div class="h"><i class="bi bi-chat-square-text"></i> 調整原因 <span class="kbd">${adjd.length} 項改過出貨量${missing ? `，<b style="color:var(--warn)">${missing} 項還沒填</b>` : ""}</span></div><div class="b">
      <div class="flex items-center gap-2 flex-wrap">${REASON_CHIPS.map(r => `<span class="chip rs-chip">${esc(r)}</span>`).join("")}</div>
      <div class="flex items-center gap-2 flex-wrap mt-2"><input class="inp" id="rs-all" placeholder="這張單的調整原因" value="${esc(common ? common[0] : "")}" style="flex:1;min-width:180px">
        <button class="btn btn-p btn-sm" id="rs-apply"><i class="bi bi-check2-all"></i> 套用到全部 ${adjd.length} 項</button></div>
      <div class="kbd mt-1">只有一品不一樣，就改下面那一品的原因（改完離開那格就存）。之後再改數量的品項會自動帶這張單的原因。</div>
      <table class="t mt-2"><tbody>${adjd.map(x => `<tr><td class="fn nw">${esc(x.supplier_sku_id)}</td><td class="num nw">${fmt(x.qty)} → <b>${fmt(x.final)}</b></td>
        <td style="width:55%"><input class="inp rs-one" data-id="${x.order_id}" value="${esc(x.adj.reason)}" placeholder="（沒填）" style="width:100%"></td></tr>`).join("")}</tbody></table>
    </div></div>`;
  }
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

  $$("#side .sd-zero, #side .sd-zero-all").forEach(b => b.addEventListener("click", async () => {
    const ids = b.dataset.ids.split(",").map(Number); if (!ids.length) return;
    if (!confirm(`把這 ${ids.length} 項的出貨量改成 0（刪單）？原因會記「低於 12 件拒收，刪單」。改錯可以在效期視窗按「回到系統建議」。`)) return;
    try { const r = await post("/api/shopee/ship/zero", { order_ids: ids, reason: "低於 12 件拒收，刪單", ...zq() }); toast(`${r.count} 項已改成 0（刪單）`); refresh(); }
    catch (err) { toast(err.message, "err"); }
  }));
  const saveRsn = async (ids, reason) => {
    try { const r = await post("/api/shopee/ship/reasons", { order_ids: ids, reason }); toast(r.count ? `${r.count} 項調整原因已存為「${reason}」` : "原因沒有變"); refresh(); }
    catch (err) { toast(err.message, "err"); }
  };
  $$("#side .rs-chip").forEach(c => c.addEventListener("click", () => { $("#rs-all").value = c.textContent; $$("#side .rs-chip").forEach(x => x.classList.toggle("on", x === c)); }));
  const rsa = $("#rs-apply"); if (rsa) rsa.addEventListener("click", () => { const v = $("#rs-all").value.trim(); if (!v) return toast("請選一個原因，或自己打", "err");
    saveRsn($$("#side .rs-one").map(i => Number(i.dataset.id)), v); });
  $$("#side .rs-one").forEach(i => i.addEventListener("change", () => { const v = i.value.trim(); if (v) saveRsn([Number(i.dataset.id)], v); }));
  $$("#side .sd-dconf").forEach(b => b.addEventListener("click", async () => {
    try { await post("/api/shopee/diffs/confirm", { order_nos: [g.order_no] }); toast(`已確認 ${g.order_no} 的差異`); refresh(); } catch (err) { toast(err.message, "err"); }
  }));
  $$("#side a.ver").forEach(a => a.addEventListener("click", () => openVersions(a.dataset.id, a.dataset.t)));
  $$("#side .sd-adj").forEach(b => b.addEventListener("click", () => openAdj(Number(b.dataset.id))));
  $$("#side .sd-mp").forEach(b => b.addEventListener("click", () => openMarsProduct(Number(b.dataset.id))));
  $$("#side .sd-ship").forEach(b => b.addEventListener("click", () => openShip(Number(b.dataset.id))));
  $$("#side .sd-short").forEach(b => b.addEventListener("click", () => openShort(Number(b.dataset.id))));
  $$("#side .sd-po-save").forEach(b => b.addEventListener("click", () => saveEip(b.dataset.id, $(`#side .sd-po[data-id="${b.dataset.id}"]`).value)));
  $$("#side .sd-eip-cancel").forEach(b => b.addEventListener("click", () => cancelEip(b.dataset.id, (d.eips || []).find(x => String(x.id) === b.dataset.id))));
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
/* 清除資料（限管理員；Jerry 2026-10-14） */
if ($("#rs-go")) $("#rs-go").addEventListener("click", async () => {
  const body = { orders: $("#rs-orders").checked, refs: $("#rs-refs").checked, logs: $("#rs-logs").checked, confirm: $("#rs-confirm").value.trim() };
  if (!body.orders && !body.refs) return toast("請至少勾選一項清除範圍", "err");
  if (body.confirm !== "清空資料") return toast("請輸入「清空資料」以確認", "err");
  if (!confirm(`確定清除${[body.orders && "訂單資料", body.refs && "參考資料"].filter(Boolean).join("、")}？清除後無法復原。`)) return;
  try { const r = await post("/api/shopee/reset", body); toast(r.message); $("#rs-confirm").value = ""; closeSide(); showSettings(true); }
  catch (err) { toast(err.message, "err"); }
});
$("#set-back").addEventListener("click", () => showSettings(false));

/* 上傳了還沒按「確認匯入」就重新整理：把預覽卡片放回來 */
async function loadPending() {
  let d; try { d = await api("/api/shopee/import/pending"); } catch (err) { return; }
  for (const p of d.pending) {
    const box = document.createElement("div"); box.className = "pv"; $("#previews").append(box);
    renderPreview(box, p);
    box.insertAdjacentHTML("afterbegin", `<div class="alert al-warn mb-3"><i class="bi bi-exclamation-triangle-fill"></i> 這個檔今天已上傳，但還沒按「確認匯入」，訂單尚未寫入。</div>`);
  }
}
loadFresh(); loadPending(); refresh();

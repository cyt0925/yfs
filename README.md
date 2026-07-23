<!DOCTYPE html>
<html lang="zh-Hant">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>YFS 營運 SOP 檢索</title>
<style>
:root{
  --ink:#1a2430; --ink-2:#3a4757; --muted:#6b7787;
  --paper:#f4f5f2; --card:#ffffff; --line:#e2e5df; --line-2:#eceee9;
  --accent:#0d7d6f; --accent-ink:#0a5a50; --accent-wash:#e3f0ed;
  --brand-mars:#b4531f; --brand-mars-wash:#f6e9e0;
  --brand-pg:#1d63b8; --brand-pg-wash:#e3eef9;
  --brand-cpg:#2f8f5b; --brand-cpg-wash:#e2f1e7;
  --brand-cross:#8155a7; --brand-cross-wash:#efe7f6;
  --brand-gen:#5c6470; --brand-gen-wash:#ecedf0;
  --m-consign:#b45309; --m-consign-wash:#f7ecdb;
  --m-buyout:#4046c7; --m-buyout-wash:#e6e7fb;
  --m-loan:#7c3aed; --m-loan-wash:#efe6fc;
  --d-direct:#3a7ca5; --d-hct:#5a6b45;
  --shadow:0 1px 2px rgba(26,36,48,.05),0 8px 24px rgba(26,36,48,.06);
  --radius:14px;
  --sans:"PingFang TC","Noto Sans TC","Microsoft JhengHei","Hiragino Sans GB",-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;
  --mono:ui-monospace,"SF Mono","JetBrains Mono","Cascadia Code",Menlo,Consolas,monospace;
}
@media (prefers-color-scheme:dark){
  :root{
    --ink:#e8ebe6; --ink-2:#c2c8cd; --muted:#8b95a1;
    --paper:#12171d; --card:#1a212a; --line:#2a333d; --line-2:#232b34;
    --accent:#37b3a2; --accent-ink:#5eccbc; --accent-wash:#173029;
    --brand-mars:#e2895a; --brand-mars-wash:#331f13;
    --brand-pg:#5ea0e8; --brand-pg-wash:#122436;
    --brand-cpg:#5ec48a; --brand-cpg-wash:#123024;
    --brand-cross:#b491d8; --brand-cross-wash:#241a33;
    --brand-gen:#9aa3ae; --brand-gen-wash:#232830;
    --m-consign:#e0a04a; --m-consign-wash:#33270f;
    --m-buyout:#8a8ff0; --m-buyout-wash:#1e2140;
    --m-loan:#b48ef2; --m-loan-wash:#271a3d;
    --d-direct:#6fb0d6; --d-hct:#96ab7a;
    --shadow:0 1px 2px rgba(0,0,0,.3),0 10px 30px rgba(0,0,0,.35);
  }
}
:root[data-theme="light"]{
  --ink:#1a2430; --ink-2:#3a4757; --muted:#6b7787;
  --paper:#f4f5f2; --card:#ffffff; --line:#e2e5df; --line-2:#eceee9;
  --accent:#0d7d6f; --accent-ink:#0a5a50; --accent-wash:#e3f0ed;
  --brand-mars:#b4531f; --brand-mars-wash:#f6e9e0;
  --brand-pg:#1d63b8; --brand-pg-wash:#e3eef9;
  --brand-cpg:#2f8f5b; --brand-cpg-wash:#e2f1e7;
  --brand-cross:#8155a7; --brand-cross-wash:#efe7f6;
  --brand-gen:#5c6470; --brand-gen-wash:#ecedf0;
  --m-consign:#b45309; --m-consign-wash:#f7ecdb;
  --m-buyout:#4046c7; --m-buyout-wash:#e6e7fb;
  --m-loan:#7c3aed; --m-loan-wash:#efe6fc;
  --d-direct:#3a7ca5; --d-hct:#5a6b45;
  --shadow:0 1px 2px rgba(26,36,48,.05),0 8px 24px rgba(26,36,48,.06);
}
:root[data-theme="dark"]{
  --ink:#e8ebe6; --ink-2:#c2c8cd; --muted:#8b95a1;
  --paper:#12171d; --card:#1a212a; --line:#2a333d; --line-2:#232b34;
  --accent:#37b3a2; --accent-ink:#5eccbc; --accent-wash:#173029;
  --brand-mars:#e2895a; --brand-mars-wash:#331f13;
  --brand-pg:#5ea0e8; --brand-pg-wash:#122436;
  --brand-cpg:#5ec48a; --brand-cpg-wash:#123024;
  --brand-cross:#b491d8; --brand-cross-wash:#241a33;
  --brand-gen:#9aa3ae; --brand-gen-wash:#232830;
  --m-consign:#e0a04a; --m-consign-wash:#33270f;
  --m-buyout:#8a8ff0; --m-buyout-wash:#1e2140;
  --m-loan:#b48ef2; --m-loan-wash:#271a3d;
  --d-direct:#6fb0d6; --d-hct:#96ab7a;
  --shadow:0 1px 2px rgba(0,0,0,.3),0 10px 30px rgba(0,0,0,.35);
}
*{box-sizing:border-box}
body{margin:0;background:var(--paper);color:var(--ink);font-family:var(--sans);line-height:1.65;
  font-feature-settings:"palt";-webkit-font-smoothing:antialiased;}
.wrap{max-width:1280px;margin:0 auto;padding:0 22px}

/* Header */
header{border-bottom:1px solid var(--line);background:
  linear-gradient(180deg,color-mix(in srgb,var(--accent) 7%,var(--paper)),var(--paper));}
.head-inner{padding:30px 0 24px;display:flex;flex-wrap:wrap;gap:22px;align-items:flex-end;justify-content:space-between}
.brandmark{display:flex;align-items:center;gap:14px}
.logo{width:44px;height:44px;border-radius:11px;background:var(--accent);color:#fff;display:grid;place-items:center;
  font-weight:800;font-size:19px;letter-spacing:.5px;flex:none;box-shadow:0 4px 12px color-mix(in srgb,var(--accent) 40%,transparent)}
.title-block h1{margin:0;font-size:23px;font-weight:800;letter-spacing:-.01em}
.title-block p{margin:3px 0 0;color:var(--muted);font-size:13.5px}
.stats{display:flex;gap:10px;flex-wrap:wrap}
.stat{background:var(--card);border:1px solid var(--line);border-radius:11px;padding:9px 15px;min-width:70px;text-align:center;box-shadow:var(--shadow)}
.stat b{display:block;font-size:20px;font-weight:800;line-height:1.1;font-variant-numeric:tabular-nums}
.stat span{font-size:11px;color:var(--muted);letter-spacing:.04em}

/* Search bar */
.searchbar{padding:18px 0 6px}
.search-row{display:flex;gap:12px;align-items:center;flex-wrap:wrap}
.search-input{position:relative;flex:1;min-width:240px}
.search-input svg{position:absolute;left:14px;top:50%;transform:translateY(-50%);width:18px;height:18px;color:var(--muted)}
.search-input input{width:100%;padding:13px 42px 13px 42px;border:1.5px solid var(--line);border-radius:12px;background:var(--card);
  color:var(--ink);font-size:15px;font-family:inherit;transition:border-color .15s,box-shadow .15s}
.search-input input:focus{outline:none;border-color:var(--accent);box-shadow:0 0 0 4px var(--accent-wash)}
.search-input .clear-x{position:absolute;right:8px;top:50%;transform:translateY(-50%);border:0;background:transparent;
  color:var(--muted);cursor:pointer;font-size:20px;padding:6px;line-height:1;border-radius:8px;display:none}
.search-input .clear-x:hover{background:var(--line-2);color:var(--ink)}
.count-pill{font-size:13px;color:var(--muted);white-space:nowrap}
.count-pill b{color:var(--ink);font-variant-numeric:tabular-nums}
.src-bar{display:flex;align-items:center;gap:10px;flex-wrap:wrap;padding:10px 0 2px;font-size:12.5px;color:var(--muted)}
.src-badge{display:inline-flex;align-items:center;gap:6px;padding:4px 11px;border-radius:999px;font-weight:600;font-size:12px;border:1px solid var(--line)}
.src-badge .dot{width:7px;height:7px;border-radius:50%;flex:none}
.src-badge.live{background:var(--accent-wash);color:var(--accent-ink);border-color:transparent}
.src-badge.live .dot{background:var(--accent)}
.src-badge.seed{background:var(--m-consign-wash);color:var(--m-consign);border-color:transparent}
.src-badge.seed .dot{background:var(--m-consign)}
.src-badge.loading{background:var(--line-2);color:var(--muted)}
.src-badge.loading .dot{background:var(--muted);animation:pulse 1s ease-in-out infinite}
@keyframes pulse{0%,100%{opacity:.3}50%{opacity:1}}
.src-refresh{border:1px solid var(--line);background:var(--card);color:var(--ink-2);border-radius:8px;
  padding:4px 11px;font-size:12.5px;cursor:pointer;font-family:inherit;display:inline-flex;align-items:center;gap:5px}
.src-refresh:hover{border-color:var(--accent);color:var(--accent-ink)}
.src-refresh svg{width:13px;height:13px}
.src-note{color:var(--muted)}
@media (prefers-reduced-motion:reduce){.src-badge.loading .dot{animation:none}}
.kw-suggest{display:flex;gap:7px;flex-wrap:wrap;padding:12px 0 4px;align-items:center}
.kw-suggest .lbl{font-size:12px;color:var(--muted);margin-right:2px}
.kw-suggest button{border:1px solid var(--line);background:var(--card);color:var(--ink-2);border-radius:999px;
  padding:5px 12px;font-size:12.5px;cursor:pointer;font-family:inherit;transition:all .13s}
.kw-suggest button:hover{border-color:var(--accent);color:var(--accent-ink)}
.kw-suggest button.on{background:var(--accent);border-color:var(--accent);color:#fff}

/* Layout */
.layout{display:grid;grid-template-columns:250px 1fr;gap:26px;padding:20px 0 60px;align-items:start}
.filters{position:sticky;top:16px;background:var(--card);border:1px solid var(--line);border-radius:var(--radius);
  padding:18px 18px 20px;box-shadow:var(--shadow)}
.filters-head{display:flex;justify-content:space-between;align-items:center;margin-bottom:6px}
.filters-head h2{margin:0;font-size:13px;font-weight:700;letter-spacing:.06em;text-transform:uppercase;color:var(--ink-2)}
.reset{border:0;background:transparent;color:var(--accent-ink);font-size:12.5px;cursor:pointer;font-family:inherit;padding:4px 6px;border-radius:7px}
.reset:hover{background:var(--accent-wash)}
.fgroup{padding:14px 0;border-top:1px solid var(--line-2)}
.fgroup:first-of-type{border-top:0;padding-top:8px}
.fgroup h3{margin:0 0 10px;font-size:12.5px;font-weight:700;color:var(--muted);letter-spacing:.03em}
.fopts{display:flex;flex-wrap:wrap;gap:7px}
.fopt{display:inline-flex;align-items:center;gap:6px;border:1px solid var(--line);background:var(--paper);color:var(--ink-2);
  border-radius:9px;padding:6px 11px;font-size:13px;cursor:pointer;font-family:inherit;transition:all .13s;user-select:none}
.fopt:hover{border-color:var(--ink-2)}
.fopt .n{font-size:11px;color:var(--muted);font-variant-numeric:tabular-nums}
.fopt.on{background:var(--accent-wash);border-color:var(--accent);color:var(--accent-ink);font-weight:600}
.fopt.on .n{color:var(--accent-ink)}
.fopt[data-k="mode"][data-v="寄銷"].on{background:var(--m-consign-wash);border-color:var(--m-consign);color:var(--m-consign)}
.fopt[data-k="mode"][data-v="買斷"].on{background:var(--m-buyout-wash);border-color:var(--m-buyout);color:var(--m-buyout)}
.fopt[data-k="mode"][data-v="領用"].on{background:var(--m-loan-wash);border-color:var(--m-loan);color:var(--m-loan)}

/* Cards */
.grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(340px,1fr));gap:16px;align-items:start}
.card{background:var(--card);border:1px solid var(--line);border-radius:var(--radius);box-shadow:var(--shadow);
  overflow:hidden;transition:border-color .15s,box-shadow .15s}
.card:hover{border-color:color-mix(in srgb,var(--accent) 45%,var(--line))}
.card.open{border-color:var(--accent);grid-column:1/-1}
.card-head{padding:16px 18px;cursor:pointer;display:flex;flex-direction:column;gap:11px}
.ch-top{display:flex;align-items:flex-start;justify-content:space-between;gap:12px}
.code{font-family:var(--mono);font-size:11.5px;color:var(--muted);letter-spacing:.02em;background:var(--paper);
  border:1px solid var(--line-2);padding:2px 7px;border-radius:6px;white-space:nowrap;font-variant-numeric:tabular-nums}
.card h3{margin:0;font-size:16.5px;font-weight:700;letter-spacing:-.01em;line-height:1.45;text-wrap:balance}
.chip-row{display:flex;flex-wrap:wrap;gap:6px;align-items:center}
.chip{font-size:11.5px;font-weight:600;padding:3px 9px;border-radius:999px;white-space:nowrap;border:1px solid transparent}
.chip.brand-MARS{background:var(--brand-mars-wash);color:var(--brand-mars)}
.chip.brand-PG{background:var(--brand-pg-wash);color:var(--brand-pg)}
.chip.brand-CPG{background:var(--brand-cpg-wash);color:var(--brand-cpg)}
.chip.brand-跨品類{background:var(--brand-cross-wash);color:var(--brand-cross)}
.chip.brand-通用{background:var(--brand-gen-wash);color:var(--brand-gen)}
.chip.plat{background:transparent;border-color:var(--line);color:var(--ink-2);font-weight:500}
.chip.mode-寄銷{background:var(--m-consign-wash);color:var(--m-consign)}
.chip.mode-買斷{background:var(--m-buyout-wash);color:var(--m-buyout)}
.chip.mode-領用{background:var(--m-loan-wash);color:var(--m-loan)}
.chip.deliv{background:transparent;border-color:currentColor;font-weight:600}
.chip.deliv.直送{color:var(--d-direct)}
.chip.deliv.竹運{color:var(--d-hct)}
.chip.status{background:var(--m-consign-wash);color:var(--m-consign);font-weight:600}
.ch-foot{display:flex;align-items:center;justify-content:space-between;gap:10px;color:var(--muted);font-size:12px}
.ch-meta{display:flex;gap:12px;flex-wrap:wrap}
.expand-ind{display:flex;align-items:center;gap:5px;color:var(--accent-ink);font-size:12.5px;font-weight:600;white-space:nowrap}
.expand-ind svg{width:15px;height:15px;transition:transform .2s}
.card.open .expand-ind svg{transform:rotate(180deg)}

/* Body (expanded) */
.card-body{display:none;border-top:1px solid var(--line-2);padding:4px 22px 24px}
.card.open .card-body{display:block;animation:fade .25s ease}
@keyframes fade{from{opacity:0;transform:translateY(-4px)}to{opacity:1;transform:none}}
.sec{padding-top:20px}
.sec-label{font-size:11.5px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;color:var(--accent-ink);
  margin:0 0 9px;display:flex;align-items:center;gap:8px}
.sec-label::before{content:"";width:14px;height:2px;background:var(--accent);border-radius:2px}
.purpose{margin:0;color:var(--ink-2);font-size:14px}
.flow{background:var(--paper);border:1px solid var(--line-2);border-radius:11px;padding:13px 15px;font-size:13.5px;color:var(--ink-2);line-height:1.9}
.flow b{color:var(--accent-ink)}
.roles{width:100%;border-collapse:collapse;font-size:13.5px}
.roles td{border:1px solid var(--line-2);padding:8px 12px;vertical-align:top}
.roles tr td:first-child{font-weight:700;color:var(--ink);white-space:nowrap;width:1%;background:var(--paper)}
.roles tr td:last-child{color:var(--ink-2)}
.steps{display:flex;flex-direction:column;gap:8px}
.step{border:1px solid var(--line-2);border-radius:10px;overflow:hidden;background:var(--card)}
.step-h{display:flex;gap:12px;align-items:center;padding:11px 14px;cursor:pointer;user-select:none}
.step-h:hover{background:var(--paper)}
.step-num{flex:none;width:24px;height:24px;border-radius:7px;background:var(--accent-wash);color:var(--accent-ink);
  font-size:12px;font-weight:700;display:grid;place-items:center;font-variant-numeric:tabular-nums}
.step-t{font-size:14px;font-weight:600;flex:1;color:var(--ink)}
.step-h .caret{width:15px;height:15px;color:var(--muted);transition:transform .2s;flex:none}
.step.open .caret{transform:rotate(90deg)}
.step-body{display:none;padding:2px 15px 14px 50px;font-size:13.5px;color:var(--ink-2);white-space:pre-line;line-height:1.8}
.step.open .step-body{display:block}
.step-body.empty{color:var(--muted);font-style:italic}
.exc{list-style:none;margin:0;padding:0;display:flex;flex-direction:column;gap:7px}
.exc li{position:relative;padding:9px 12px 9px 34px;background:color-mix(in srgb,var(--m-consign-wash) 55%,var(--card));
  border:1px solid color-mix(in srgb,var(--m-consign) 22%,var(--line-2));border-radius:9px;font-size:13.5px;color:var(--ink-2)}
.exc li::before{content:"!";position:absolute;left:12px;top:9px;width:15px;height:15px;border-radius:50%;
  background:var(--m-consign);color:#fff;font-size:10.5px;font-weight:800;display:grid;place-items:center}
.tags{display:flex;flex-wrap:wrap;gap:6px}
.tag{font-size:11.5px;color:var(--muted);background:var(--paper);border:1px solid var(--line-2);border-radius:6px;padding:2px 8px}
mark{background:color-mix(in srgb,var(--accent) 28%,transparent);color:inherit;border-radius:3px;padding:0 1px}

.empty-state{grid-column:1/-1;text-align:center;padding:70px 20px;color:var(--muted)}
.empty-state svg{width:44px;height:44px;opacity:.5;margin-bottom:12px}
.empty-state p{margin:0;font-size:15px}
.empty-state button{margin-top:14px;border:1px solid var(--accent);background:transparent;color:var(--accent-ink);
  border-radius:9px;padding:8px 16px;cursor:pointer;font-family:inherit;font-size:13.5px}

footer{border-top:1px solid var(--line);color:var(--muted);font-size:12.5px;padding:20px 0 40px;text-align:center;line-height:1.7}

.mobile-filter-toggle{display:none}
@media (max-width:860px){
  .layout{grid-template-columns:1fr;gap:14px}
  .filters{position:static}
  .mobile-filter-toggle{display:flex;align-items:center;justify-content:space-between;width:100%;border:1px solid var(--line);
    background:var(--card);border-radius:11px;padding:12px 16px;font-family:inherit;font-size:14px;font-weight:600;color:var(--ink);cursor:pointer;box-shadow:var(--shadow)}
  .filters.collapsed .filters-body{display:none}
  .head-inner{padding:22px 0 18px}
  .stats{width:100%}
  .grid{grid-template-columns:1fr}
  .card.open{grid-column:auto}
}
</style>
</head>
<body>

<header>
  <div class="wrap head-inner">
    <div class="brandmark">
      <div class="logo">YFS</div>
      <div class="title-block">
        <h1>營運 SOP 檢索中心</h1>
        <p>電子商務部 · 各品牌與通路出貨作業標準流程 · 全三部分</p>
      </div>
    </div>
    <div class="stats" id="stats"></div>
  </div>
</header>

<div class="wrap searchbar">
  <div class="search-row">
    <div class="search-input">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="11" cy="11" r="7"/><path d="m21 21-4.3-4.3"/></svg>
      <input id="search" type="text" placeholder="搜尋流程、平台、關鍵字（退貨、補貨、嘜頭、EIP、缺貨…）" autocomplete="off">
      <button class="clear-x" id="clearSearch" aria-label="清除搜尋">×</button>
    </div>
    <div class="count-pill"><b id="resultCount">46</b> / <span id="totalCount">46</span> 份文件</div>
  </div>
  <div class="src-bar">
    <span class="src-badge loading" id="srcBadge"><span class="dot"></span><span id="srcText">載入中…</span></span>
    <button class="src-refresh" id="refreshBtn" title="重新從試算表載入最新內容">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 12a9 9 0 1 1-2.64-6.36"/><path d="M21 3v6h-6"/></svg>重新整理</button>
    <span class="src-note" id="srcNote"></span>
  </div>
  <div class="kw-suggest" id="kwSuggest"><span class="lbl">熱門關鍵字</span></div>
</div>

<div class="wrap">
  <div class="layout">
    <button class="mobile-filter-toggle" id="mFilter">篩選條件 <span>＋</span></button>
    <aside class="filters collapsed" id="filters">
      <div class="filters-head">
        <h2>篩選條件</h2>
        <button class="reset" id="resetBtn">重設</button>
      </div>
      <div class="filters-body" id="filtersBody"></div>
    </aside>
    <main class="grid" id="grid"></main>
  </div>
</div>

<footer class="wrap">
  YFS 營運 SOP 檢索 · 共 <span id="footTotal">46</span> 份標準作業程序文件 · 內容依原始 .docx 文件整理<br>
  <span id="footBreak">MARS 瑪氏 8 · PG 寶僑 11 · CPG 紙潔 8 · 跨品類 14 · 通用作業 5</span>
</footer>

<!-- 設定檔：試算表網址放在 config.js（改那個檔就好，不用動這個大檔案）。 -->
<script src="config.js"></script>
<script>
/* 若 config.js 不存在或未設定，預設為「內建範例」模式。 */
const CONFIG = Object.assign(
  { SOPS_CSV_URL:"", STEPS_CSV_URL:"", ROLE_SEP:"｜" },
  (window.CONFIG||{})
);

/* 內建範例資料（種子）＋ 關鍵字詞庫。有設定試算表時會被試算表內容覆蓋。 */
const SEED = [{"code": "SOP-CP-MARS-001", "title": "酷澎－買斷出貨作業流程", "brand": "MARS", "platform": "酷澎", "modes": ["買斷"], "delivery": "直送", "version": "V2.0", "date": "2026/07/17", "author": "黃千容", "purpose": "建立標準化的酷澎出貨流程，確保入庫、補貨與拋檔作業正確，並降低錯誤率。", "overview": "收到業務通知 → 建立資料夾並製作出貨文件 → EIP採購單下採 → 發送出貨採購單給原廠 → 製作EMMA打單資料 → 原廠確認後通知勇信打單 → 製作出貨嘜頭 → 原廠回覆效期後確認嘜頭 → 出貨文件給勇信 → 缺貨另行通知小真下修", "roles": [["業務", "提供補貨PO單號、倉別"], ["企劃(OP)", "製作出貨相關文件（EIP下採檔案、嘜頭、驗收單）"], ["採購", "與原廠下採單，發信通知缺貨"]], "steps": [["步驟 1：收到業務出貨通知", "前置條件：收到業務補貨通知（提供補貨PO單號、倉別）。"], ["步驟 2：建立相關出貨文件", "確認到貨日資料夾（以出貨日命名），新增本次檔案，包含出貨文件、採購單文件、訂單匯入等檔案。"], ["步驟 3：製作採購單", "1. 打開瑪氏採購單檔案，複製檔名至採購單文件（檔名須相同），將永豐料號、採購數量貼至採購單表格，備註填上到貨時間與倉別。\n2. 確認採購單箱數與瑪氏採購單箱數相同。\n3. 於整合表新增工作表，依訂單編號／料號／數量複製（方便貼採購單號並帶入EMMA匯入資料）。"], ["步驟 4：上傳採購表至EIP系統", "將產品採購表上傳EIP，申請後的PO單回填至瑪氏採購單「永豐PO單號」，並在表格名稱前加上。\n注意事項：注意需求日期及寄銷倉庫選擇。"], ["步驟 5：製作Emma出貨表單", "1. 倉儲物流管理 → 酷澎採購入庫處理 → 批次上傳採購單 → 選擇檔案 → 批次匯入。\n2. 整合表預覽匯出（善用篩選：上傳人／倉別／線別／到貨日期），查詢後全選，先選匯出拆單格式EXCEL（選U欄），再匯出酷澎訂單匯入檔（瑪氏），開啟取消保護（密碼1234）。\n3. 把拆單格式的U欄貼至匯入檔K欄；出貨備註A欄帶入EIP採購單號。\n提醒：核對單號數量，同料號不同單號只會吃到第一筆訂單。"], ["步驟 6：發信給瑪氏Lucy打單", "前置作業：目前打單為D+2出貨，下單日前一天下班前將信件發出。\n依瑪氏採購單分倉別，一個PO單發一張MAIL。"], ["步驟 7：確認打單內容", "前置條件：待Lucy確認打單回覆。\nLucy會針對單一採購單回覆打貨狀況，需每封確認；若有缺貨於整合表備註，並於驗收單及酷澎採購單刪除缺品項。"], ["步驟 8：勇信回覆效期後製作嘜頭", "前置條件：勇信約下午15:30後提供效期信。\n整合表產製：Tab3勾選資料列 → 產生嘜頭（系統自動判斷板嘜／箱嘜），確認效期如實帶入，批次下載Zip。\n手動匯入：Tab3下載嘜頭格式Excel編輯（品名／效期／數量），比對無效期品項是否等同缺貨，雙效期需修改序號與箱數，再匯入產生嘜頭。\n確認無誤後將出貨文件壓縮為7z。"], ["步驟 10：回覆出貨文件", "將壓縮7z的出貨文件及進倉規範回覆勇信。"], ["步驟 11：缺貨下修", "最終出貨結果確認，無法出貨商品須通知小真下修；可與賜玟發的缺貨信比對整合表確認。"]], "exceptions": []}, {"code": "SOP-HCT-MARS-001", "title": "HCT新竹物流（瑪氏直送）出貨作業流程", "brand": "MARS", "platform": "HCT新竹物流", "modes": [], "delivery": "竹運", "version": "V1.0", "date": "2025/10/02", "author": "伍龍英", "purpose": "建立標準化的HCT新竹物流（瑪氏直送）出貨流程，確保入庫、補貨與拋檔作業正確，並降低錯誤率。作業時間 D+2.5 天。", "overview": "業務提供總表 → 建立資料夾並整理文件 → 製作採購單 → 訂單系統拋檔(EIP) → 發送檔案給原廠 → 原廠提供效期後填上新竹物流嘜頭 → 發送檔案給採購 → 採購寄給原廠", "roles": [["業務", "提供總表，指定到貨日期"], ["企劃(OP)", "依據製作採購單，處理出貨單及系統拋檔"], ["採購", "驗收與結帳"]], "steps": [["步驟 1：收到業務提供的總表", "開啟總表 → 依 PET/GUM/CHO 與貼標／不貼標 與 箱/盒/包，製作瑪氏採購單（請使用最新版本）。"], ["步驟 2：EIP採購單申請與填入", "將業務提供的採購單填入EIP(Excel)，拋檔匯入永豐系統產生PO號碼，PO號碼填入採購單。"], ["步驟 3：發信給瑪氏（OP）", "將準備完成的採購單發信給相關人員。注意確認收件人及出貨文件正確。"], ["步驟 4：匯入效期並填入嘜頭", "下午15:30勇信提供效期信（下載PDF）→ 採購通知後PDF匯入系統：倉儲物流管理 > 批次上傳配送表PDF > 進倉倉別:竹運 > 選擇效期信(PDF) > 批次匯入。\n點選竹運嘜頭與效期合併，收貨單號貼上PO號碼 → 查詢採購單號；入庫單號跳出POIN即抓到資料 → 載入嘜頭與效期。\n嘜頭產出後全選（未找到效期需手動填入PDF）→ 批次印出合併嘜頭PDF，空白欄位填入效期並回傳採購。"], ["步驟 5：採購提供嘜頭給原廠", ""]], "exceptions": ["採購單或出貨單數量不符，需立即回報業務。", "入庫或裸品嘜頭缺漏，需重新匯出補齊。", "庫存不足，立即與業務確認。", "通知瑪氏後若無法出貨，立即與業務確認，需下修數量（重新拋單）或不出。"]}, {"code": "SOP-MO-MARS-001", "title": "MOMO 出貨作業流程", "brand": "MARS", "platform": "MOMO", "modes": ["寄銷"], "delivery": "直送", "version": "V1.0", "date": "2025/10/03", "author": "伍龍英", "purpose": "建立標準化的Momo出貨流程，確保入庫、補貨與拋檔作業正確，並降低錯誤率。作業時間 D+3.5 天。", "overview": "業務提供總表、採購單 → 建立資料夾並整理文件 → 核對補貨及採購單 → 訂單系統拋檔(EIP) → 發送檔案給原廠 → 原廠提供效期後回填Momo後台 → 下載嘜頭+入庫單給原廠", "roles": [["業務", "提供總表、採購單，指定到貨日期"], ["企劃(OP)", "依據流程處理入庫、出貨單及系統拋檔"]], "steps": [["步驟 1：EIP採購單申請與填入", "將業務提供的採購單填入EIP(Excel)，拋檔匯入永豐系統產生PO號碼，PO號碼填入採購單。"], ["步驟 3：發信給瑪氏（OP）", "將準備完成的採購單發信給相關人員，確認收件人及出貨文件正確。"], ["步驟 4：原廠提供效期後回填Momo（OP）", "原廠提供效期後回填Momo，Momo填入效期、儲存日期並下載嘜頭＋入庫單，再回信給原廠。"], ["步驟 5：回信給瑪氏（OP）", ""]], "exceptions": ["採購單或出貨單數量不符，需立即回報業務。", "入庫嘜頭缺漏，需重新提供補齊。", "庫存不足，立即與業務確認。", "通知瑪氏後若無法出貨，需下修數量（重新拋單）或不出。"]}, {"code": "SOP-PC-MARS-001-直送", "title": "PCHOME 出貨作業流程（直送）", "brand": "MARS", "platform": "PCHOME", "modes": ["寄銷"], "delivery": "直送", "version": "V1.0", "date": "2025/10/02", "author": "伍龍英", "purpose": "建立標準化的Pchome出貨流程，確保入庫、補貨與拋檔作業正確，並降低錯誤率。作業時間 D+3.5 天。", "overview": "業務提供總表、採購單 → Pchome後台回填到貨日與數量 → 合併(嘜頭+二聯單)PDF與下載 → 建立資料夾並整理文件 → 核對補貨及採購單 → 庫存確認 → 訂單系統拋檔 → 發送檔案給原廠", "roles": [["業務", "提供總表、採購單，指定到貨日期"], ["企劃(OP)", "依據流程處理入庫、出貨單及系統拋檔"]], "steps": [["步驟 1：收到業務提供的總表、採購單", "登入Pchome後台回填到貨日與數量(CSV檔)，合併(嘜頭+二聯單)PDF與下載。"], ["步驟 2：EIP採購單申請與填入", "將採購單填入EIP(Excel)，拋檔匯入永豐系統產生PO號碼並填入採購單。"], ["步驟 3：發信給瑪氏（OP）", "將準備完成的檔案（採購單、嘜頭+二聯單等）發信給相關人員。"]], "exceptions": ["採購單或出貨單數量不符，需立即回報業務。", "入庫或裸品嘜頭缺漏，需重新匯出補齊。", "庫存不足，立即與業務確認。", "通知瑪氏後若無法出貨，需下修數量（重新拋單）或不出。"]}, {"code": "SOP-PC-MARS-001-竹運", "title": "PCHOME 出貨作業流程（竹運）", "brand": "MARS", "platform": "PCHOME", "modes": ["寄銷"], "delivery": "竹運", "version": "V1.0", "date": "2026/07/02", "author": "伍龍英", "purpose": "建立標準化的Pchome出貨流程，確保入庫、補貨與拋檔作業正確，並降低錯誤率。作業時間 D+4 天。", "overview": "業務提供補貨列表 → Pchome後台下載(嘜頭+二聯單) → 合併PDF與下載 → 建立資料夾並整理文件 → 核對補貨及採購單 → 庫存確認 → 訂單系統拋檔 → 發送檔案給竹運", "roles": [["業務", "提供補貨列表，指定到貨日期"], ["企劃(OP)", "依據流程處理入庫、出貨單及系統拋檔"]], "steps": [["步驟 1：收到補貨列表並確認庫存", "複製上一份資料夾「指定到貨日」→ 貼上本次B單品項 → 確認庫存 → B單拋檔 → 登入Pchome後台下載(嘜頭+二聯單) → 合併PDF → 發送檔案給竹運。\nB單匯入EIP(Excel)：填上訂單編號、料號、產品名稱、單位、數量（十五倉，姓名地址電話不動）。\n將商品複製到「莉絲的出貨效期及庫存V5.3版」，更新全商品與效期，填入出貨日；Pchome效期2/5填0.4，清除資料後按「開始核對(單一效期)」，系統篩選無法出貨商品貼給業務。"], ["步驟 2：出貨資料檔案匯入", "訂單系統 → 接收電子商城訂單 → 商城名稱:B2B移倉出貨單，成立移倉單。"], ["步驟 3：發信給竹運", "將準備完成的檔案（採購單、嘜頭+二聯單等）發信給相關人員。"]], "exceptions": ["採購單或出貨單數量不符，需立即回報業務。", "入庫或裸品嘜頭缺漏，需重新匯出補齊。", "庫存不足，立即與業務確認。", "通知瑪氏後若無法出貨，需下修數量（重新拋單）或不出。"]}, {"code": "SOP-UBER-001", "title": "優食（UberEats）出貨流程", "brand": "MARS", "platform": "UberEats", "modes": [], "delivery": "竹運", "version": "V1.0", "date": "2025/10/01", "author": "黃千容", "purpose": "建立標準化的優食出貨流程，確保出貨作業正確並能有效追蹤，以降低錯誤率並提升作業效率。", "overview": "店長下單 → 系統通知信 → 建立出貨資料夾 → 依訂單品項填入UBER訂單匯入表 → 竹運庫存表轉換貼上竹運效期 → 核對可出貨 → 系統拋單 → 確認缺貨/異常放行 → 產生O單至優食好物商城訂單匯出 → 確認條碼完整 → 放入優食雲端", "roles": [["業務", "確認庫存效期／商品金額／下採"], ["企劃", "依照流程執行出貨流程及系統拋單"]], "steps": [["步驟 1：收到系統通知信", "1-1 登入UBER訂單匯入表，訂單編號欄位為PO#。\n1-2 Sheet門市資訊依各店輸入。\n1-3 國條輸入後公式自動帶入料號(UB+國條)。\n1-4 UB下單為最小單位，匯入表需換算「箱」單位出貨。\n1-5 UB訂單為未稅價，核對金額時訂單需*1.05。\n1-6 金額不符通知業務確認電商促銷設定；無資料/金額不符通知業務新增價格，否則刪除品項不出貨。\n1-7 依新竹庫存表篩選料號填入Sheet(新竹效期)，效期不足刪單不出並於群組告知。\n1-8 將反黃欄位貼到UBER訂單匯入後拋單（記得貼值）。\n新竹後台：報表查詢 → 共用-儲位及庫存/庫存明細表 → 良品區 → 查詢 → Download。"], ["步驟 2：拋單", "電子商城新訂單 → 商城名稱:優食好物UB → 匯入備註指定到貨日yyyymmdd(當天拋單日期)。"], ["步驟 3：出貨單處理", "3-1 有出貨O單後，至優食好物商城訂單匯出PO單，存至當日出貨單資料夾。\n3-2 出貨單上傳雲端依各分店放入（新增商品需業務更新，否則出貨單空白需人工帶入）。"]], "exceptions": ["拋單後再次確認訂單狀況，是否有缺貨或金額異常要放行。", "必要時附上聯絡窗口資訊。"]}, {"code": "SOP-BK-PG-001", "title": "博客來－竹運出貨作業流程", "brand": "PG", "platform": "博客來", "modes": ["寄銷", "領用"], "delivery": "竹運", "version": "V1.0", "date": "2025/09/30", "author": "杜孟涵", "purpose": "建立標準化的博客來出貨流程，確保入庫、補貨與拋檔作業正確，並降低錯誤率。", "overview": "業務通知 → 建立出貨資料夾與匯入檔 → 收取訂單與匯出明細 → 處理出貨明細與系統拋單 → 列印採購單與填寫出貨資料 → 列印外箱嘜頭 → 貼附有效期限標示 → 發送出貨文件與登記", "roles": [["業務", "提供出貨單，指定到貨日期"], ["企劃(OP)", "依據流程處理入庫、出貨單及系統拋檔"]], "steps": [["步驟 1：建立出貨資料夾與匯入檔", "建立新資料夾（以出貨日命名），複製上次出貨資料並更新為本次檔案。確保檔名格式正確。"], ["步驟 2：收取訂單與匯出明細", "至 [進貨收單作業]-[收取訂單及填寫出貨資料]-[收取訂單] → 勾選本次採購單號 → 收單及匯出 → 核對下載明細。"], ["步驟 3：處理出貨明細與系統拋單", "將出貨明細貼入匯入檔，確認商品庫存後，至訂單系統拋單。"], ["步驟 4：列印採購單與填寫出貨資料", "至 [填寫出貨資料] → 勾選採購單號 → [列印採購單] → [填寫出貨資料]。"], ["步驟 5：列印外箱嘜頭", "至 [列印外箱嘜頭] → 查詢 → 選擇出貨單號 → 列印嘜頭下載PDF。"], ["步驟 6：貼附有效期限標示", "將下載之嘜頭逐張貼上本次出貨商品有效期限（每一張都要貼）。"], ["步驟 7：發送出貨文件與登記", "將出貨文件(採購單及外箱嘜頭)發信給相關人員，並至雲表登記出貨。"], ["領用單出貨", "業務申請領用單(通常為贈品)。執行步驟2-步驟6，不須至訂單系統拋單，將出貨檔案發信給Emma通知出貨。信件內需提供領用單號及出貨清單(料號、品名、單位、數量)。"]], "exceptions": ["更改出貨數量至[填寫出貨數量]修改；回填錯誤選「是否回填數量=是」後更正。", "忘記列印採購單或需修改，至[列印外箱嘜頭]按回復後重新操作。", "幫寶適訂單需入「明治大竹倉」，注意出貨地址不同。", "一次最多合併4張採購單，回填總箱數不得超過200箱。", "庫存不足，立即與業務確認。", "有效期問題無法出貨，需下修數量（重新拋單）或不出。"]}, {"code": "SOP-BK-PG-002", "title": "博客來－直送出貨作業流程", "brand": "PG", "platform": "博客來", "modes": ["寄銷", "領用"], "delivery": "直送", "version": "V1.0", "date": "2025/10/15", "author": "杜孟涵", "purpose": "建立標準化的博客來出貨流程（原廠直送），確保入庫、補貨與拋檔作業正確，並降低錯誤率。", "overview": "業務通知 → 建立資料夾與出貨文件 → 收取訂單與匯出明細 → 製作採購單及匯入檔 → EIP上傳採購單 → 發信 → 核單確認缺貨 → 修改出貨數量 → 列印採購單與出貨資料 → 列印外箱嘜頭 → 貼附有效期限 → 發送出貨文件", "roles": [["業務", "提供出貨單，指定到貨日期"], ["企劃(OP)", "依據流程處理入庫、出貨單及系統拋檔"]], "steps": [["步驟 1：建立出貨資料夾與出貨文件", "建立新資料夾（以出貨日命名），複製上次出貨資料並更新。確保檔名格式正確。"], ["步驟 2：收取訂單與匯出明細", "至 [收取訂單] → 勾選採購單號 → 收單及匯出 → 核對明細。每張採購單上限200箱，超過通知業務拆單。"], ["步驟 3：製作出貨文件（OP）", "將補貨資料填入採購單及匯入檔，匯入檔填入庫存跟均銷，至EIP上傳採購單。\n庫存(百貨庫存查詢)&均銷(百貨接單查詢，訂單日期選一個月，樞紐條碼&接單量小計)。\n用供應商料號來V，V進數字後要/轉換率（以箱為單位）。"], ["步驟 4：EIP系統拋單（OP）", "登入EIP → 申請表單 → 產品領用/採購 → 產品採購申請 → 完成拋單。"], ["步驟 5：發信（OP）", "將採購單發給相關人員，確認收件人及出貨文件正確。"], ["步驟 6：核單確認缺貨", "收到採購缺貨通知信 → 確認缺貨品項 → 至博客來後台下修數量。缺貨務必刪單，未到貨會產生罰款。"], ["步驟 7：列印採購單與填寫出貨資料", "至 [填寫出貨資料] → 勾選採購單號 → [列印採購單] → [填寫出貨資料]。"], ["步驟 8：列印外箱嘜頭", "至 [列印外箱嘜頭] → 查詢 → 選擇出貨單號 → 列印嘜頭下載PDF。"], ["步驟 9：貼附有效期限標示", "將嘜頭逐張貼上有效期限（每一張都要貼）。"], ["步驟 10：發送出貨文件", "將出貨文件(採購單及外箱嘜頭)發信給相關人員。"], ["領用單出貨", "贈品若從原廠直出，作業方式同寄銷流程，待採購壓到貨後業務開立領用單；信件標題須標示贈品領用，領用單後補。核單後確認缺貨，下載外箱嘜頭、採購單回信給採購。"]], "exceptions": ["更改出貨數量至[填寫出貨數量]修改。", "忘記列印採購單，至[列印外箱嘜頭]按回復後重新操作。", "幫寶適訂單需入「明治大竹倉」。", "一次最多合併4張採購單，總箱數不得超過200箱。", "庫存不足，立即與業務確認。", "有效期問題無法出貨，需下修數量或不出。"]}, {"code": "SOP-CP-PG-001", "title": "酷澎－買斷出貨作業流程（直送）", "brand": "PG", "platform": "酷澎", "modes": ["買斷"], "delivery": "直送", "version": "V2.0", "date": "2026/06/29", "author": "林佳慧", "purpose": "建立標準化的酷澎出貨流程，確保入庫、補貨與拋檔作業正確，並降低錯誤率。", "overview": "收到業務補貨通知 → 約倉 → 酷澎後台下載PO單及驗收單 → 建立資料夾並製作出貨文件 → EIP採購單下採 → 發送檔案給相關單位 → 出貨前一天收到缺貨通知 → 製作出貨檔案 → 發送檔案給相關單位", "roles": [["業務", "提供補貨PO單號、倉別"], ["企劃(OP)", "約倉、製作出貨相關文件（EIP下採檔案、嘜頭、驗收單）"], ["採購", "與原廠下採單，發信通知缺貨"]], "steps": [["步驟 1：補貨通知並確認車數及約倉", "收到酷澎小組補貨通知 → 酷澎後台下載PO單及驗收單。"], ["步驟 2：建立出貨資料夾", "建立新資料夾（以出貨日命名），複製上次出貨資料並更新，包PO單、驗收單、出貨資料夾等共5-6個檔案。"], ["步驟 3：製作採購單", "1. 訂單系統批次上傳採購單。2. 到整合表預覽匯出訂單匯入檔，整理表格後分別貼入採購單。"], ["步驟 4：上傳採購表至EIP並發信請Emma轉單", "產品採購表上傳EIP，完成後發信通知。注意需求日期及寄銷倉庫選擇。"], ["步驟 5：製作配置明細", "因PG CODE關係，出貨資料要抓最新的（無論倉別），刪除出貨資料內PDF，產品採購單&訂單匯入表格可抓上一次同倉。"], ["步驟 5：更改嘜頭檔案", "採購打單後會提供單號，把嘜頭檔案改為正確單號。"], ["步驟 6：出貨前一日確認缺貨並發送嘜頭給原廠", "採購打單才知有無缺貨，依缺貨信把驗收單劃掉或修改出貨數並畫押，存嘜頭資料夾，將嘜頭寄至寶僑人員。"], ["步驟 7：製作出貨嘜頭", "待寶僑人員回信提供嘜頭 → 下載檔案覆蓋原嘜頭 → 打開word檔製作出貨嘜頭。"], ["步驟 8：發信通知出貨", "壓縮嘜頭資料夾寄出並提供約倉時段。"]], "exceptions": []}, {"code": "SOP-CP-PG-002", "title": "酷澎買斷－竹運出貨作業流程", "brand": "PG", "platform": "酷澎", "modes": ["買斷"], "delivery": "竹運", "version": "V1.0", "date": "2025/10/08", "author": "杜孟涵", "purpose": "建立標準化的酷澎出貨流程，確保入庫、補貨與拋檔作業正確，並降低錯誤率。", "overview": "收到業務補貨通知 → 酷澎後台下載PO單及驗收單 → 建立資料夾並製作出貨文件 → 約倉 → 確認庫存 → 訂單系統拋檔 → 發送檔案給相關單位", "roles": [["業務", "提供補貨PO單號、倉別"], ["企劃(OP)", "約倉、製作出貨相關文件（訂單匯入檔、嘜頭、驗收單）"]], "steps": [["步驟 1：補貨通知", "收到業務補貨通知 → 酷澎後台下載PO單及驗收單。"], ["步驟 2：建立出貨資料夾", "確認到貨日符合出貨天數規範 → 建立新資料夾（以出貨日命名），複製上次資料並更新，包PO單、驗收單、出貨資料夾等共5-6個檔案。"], ["步驟 3：製作出貨文件", "1. 打開整合表測試檔案、永豐料號對照表、訂單匯入檔案、PO單共3個檔案。把PO單解除跨欄置中貼到整合表測試檔案。\n2. 把B及D欄剖析，第二個sheet「整合表」已代入公式，將對應PO單資料貼進永豐料號對照表。\n注意：永豐料號對照表需自行填入單位。"], ["步驟 4：約倉", "預估車數及約倉，通常以150箱/車估算。路徑：酷澎首頁 → 運輸 → 創建預約出貨 → 選擇倉庫代號及交貨日 → 搜尋勾選訂單約倉。如額滿需與業務確認是否更改倉別。"], ["步驟 5：確認庫存並拋檔", "確認商品皆有庫存後，訂單系統 → 外部訂單 → 電子商城新訂單 → 完成拋單。檔名依指定格式，確認拋單成功。"], ["步驟 6：製作嘜頭", "以PO單號建立嘜頭資料夾，區分板嘜/箱嘜。箱嘜一支SKU一張；板嘜一張最多5支SKU。不同PO請分開資料夾。"], ["步驟 7：驗收單及進貨規範", "將驗收單畫押簽名；若缺貨(不出或下修)通知小真下修數量，簽收單可手動更改或重新下載，放入最新進貨規範。下修數量務必於到貨日D-1完成。"], ["步驟 8：發信通知出貨", "壓縮出貨檔案資料夾，將壓縮檔及料號對照表寄給物流並提供約倉時段。"]], "exceptions": []}, {"code": "SOP-ET-PG-001", "title": "東森－竹運出貨作業流程", "brand": "PG", "platform": "東森", "modes": ["寄銷", "領用"], "delivery": "竹運", "version": "V1.0", "date": "2025/09/26", "author": "杜孟涵", "purpose": "建立標準化的東森出貨流程，確保入庫、補貨與拋檔作業正確，並降低錯誤率。", "overview": "業務提供出貨單 → 東森後台入庫指示書 → 匯出與列印 → 建立資料夾並整理文件 → 核對補貨及採購單 → 庫存確認 → 訂單系統拋檔 → 發送檔案給物流", "roles": [["業務", "提供出貨單，指定到貨日期"], ["企劃(OP)", "依據流程處理入庫、出貨單及系統拋檔"]], "steps": [["步驟 1：建立出貨資料夾", "建立新資料夾（以出貨日命名），複製上次資料並更新，包含出貨單、入庫指示書、裸品嘜頭等共4-5個檔案。"], ["步驟 2：收到業務提供的出貨單", "登入東森後台入庫指示作業，點選入庫指示，數字改為全部數量，按查詢。"], ["步驟 3：匯出入庫資料", "將所有品項匯出Excel，使用樞紐統計採購單數量與商品數量是否足夠。"], ["步驟 4：勾選出貨品項", "依業務出貨單勾選品項，舊單優先使用；同品項多張單需逐一確認數量。"], ["步驟 5：匯出與列印文件", "匯出Excel，列印入庫指示書與裸品嘜頭存PDF。小單位品項必須有裸品嘜頭。"], ["步驟 6：更新補貨及採購單檔案", "依序填入採購單及匯入檔，補齊缺少貨號（商品管理→報表管理→商品資料報表查詢）。"], ["步驟 7：確認庫存並拋檔", "確認庫存後，訂單系統 → 外部訂單 → 接收電子商城訂單 → 完成拋單。"], ["步驟 8：發信及登錄（OP）", "將檔案（出貨單、入庫指示書、裸品嘜頭等）發信給相關人員，並至雲表登記出貨。"], ["領用單出貨", "業務申請領用單(通常為贈品)。執行步驟2-步驟6，不須拋單，將出貨檔案發信給Emma。信件內需提供領用單號及出貨清單。"]], "exceptions": ["採購單或出貨單數量不符，需立即回報業務。", "入庫或裸品嘜頭缺漏，需重新匯出補齊。", "庫存不足，立即與業務確認。", "有效期問題無法出貨，需下修數量（重新拋單）或不出。"]}, {"code": "SOP-PC-PG-001", "title": "PCHOME 寄銷/買斷/領用－原廠直出流程", "brand": "PG", "platform": "PCHOME", "modes": ["寄銷", "買斷", "領用"], "delivery": "直送", "version": "V2.0", "date": "2026/06/29", "author": "林佳慧", "purpose": "規範PCHOME出貨作業流程，確保買斷與寄銷出貨執行一致性，降低錯誤並提升作業效率。", "overview": "業務補貨通知 → 建立公槽資料夾 → 區分買斷/寄銷/贈品流程 → 製作採購單與上傳表 → 上傳EIP系統 & 通知採購 → 出貨前一天核單與缺貨修正 → 下載嘜頭、寄信通知出貨", "roles": [["業務", "提供補貨資料、確認到貨日"], ["採購", "下採單、發信通知缺貨"], ["企劃(OP)", "建資料夾、製作/修改採購單與上傳檔、上傳EIP系統、下載嘜頭、通知出貨"]], "steps": [["步驟 1：接收補貨通知並約倉", "收到業務補貨通知。"], ["步驟 2：建立公槽資料夾", "至公槽新增到貨日資料夾，複製上一筆檔案修改。檔名需符合命名規範。"], ["步驟 3：製作/修改採購單", "依業務補貨資料貼入採購單；有公式欄位不得更動；需帶入「單位」及「箱入數」等必要欄位（需V的資料）。"], ["步驟 4：製作/修改EIP上傳表", "買斷：只需填寫箱數(G欄)。寄銷：需同時填寫庫存(E欄)及實銷(F欄)。庫存來源=借貨單良品庫存；實銷來源=寄倉商品接單統計(一個月)。庫存公式=ROUNDDOWN(良品庫存/轉換率,0)；實銷公式=ROUND(實銷/轉換率,0)。備註更改當次交貨日。"], ["步驟 5：上傳EIP系統與通知採購", "登入EIP → 申請表單 → 產品領用/採購 → 產品採購申請；上傳買斷或寄銷採購單；寄信給採購等待核單。選擇正確寄銷倉別。"], ["步驟 6：出貨前一天核單與修正", "採購發信通知缺貨品項；修改採購單並下修PCHOME後台數量；下載嘜頭，將更新後出貨檔案寄給採購。"], ["步驟 7：贈品領用", "贈品若從原廠直出，作業方式同寄銷流程，待採購壓到貨後業務開立領用單；信件標題須標示贈品領用，領用單後補。贈品領用都是寄銷(借貨單)，皆使用虛擬料號。"]], "exceptions": ["缺貨品項 → 出貨前一天依採購通知修正數量及更改約倉日期。"]}, {"code": "SOP-PC-PG-002", "title": "PCHOME 寄銷/買斷/領用－竹運出貨流程", "brand": "PG", "platform": "PCHOME", "modes": ["寄銷", "買斷", "領用"], "delivery": "竹運", "version": "V1.0", "date": "2025/10/02", "author": "杜孟涵", "purpose": "規範PCHOME出貨作業流程，確保買斷與寄銷出貨執行一致性，降低錯誤並提升作業效率。", "overview": "業務補貨通知 → 約倉（成功則出貨，失敗回報業務）→ 建立公槽資料夾 → 區分買斷/寄銷/贈品流程 → 製作採購單與上傳表 → 上傳訂單系統 → 缺貨修正 → 下載嘜頭、寄信通知出貨", "roles": [["業務", "提供補貨資料、確認到貨日"], ["企劃(OP)", "建資料夾、製作/修改採購單與上傳檔、上傳訂單系統、下載嘜頭、通知出貨"]], "steps": [["步驟 1：接收補貨通知並約倉", "收到業務補貨通知。"], ["步驟 2：建立公槽資料夾", "約倉成功 → 至公槽新增到貨日資料夾，複製上一筆檔案修改。檔名需符合規範，買斷及領用不須製作上傳檔。"], ["步驟 3：製作/修改採購單(買斷/寄銷/領用)", "依業務補貨資料貼入採購單；有公式欄位不得更動；需帶入「單位」及「箱入數」。買斷/寄銷/領用皆須製作採購單。"], ["步驟 4：製作/修改訂單系統上傳表(寄銷)", "依業務補貨資料貼入上傳檔。檔名依指定格式。僅寄銷需自行拋檔入訂單系統。"], ["步驟 5-1：上傳訂單系統與通知出貨(寄銷)", "確認庫存後，訂單系統 → 外部訂單 → 接收電子商城訂單 → 完成拋單。至PCHOME下載嘜頭，將出貨檔案(採購單、嘜頭)寄信給物流。"], ["步驟 5-2：通知出貨(買斷)", "確認庫存後，至PCHOME下載嘜頭，將出貨文件寄信給Emma。買斷信件不須通知物流。"], ["步驟 5-3：通知出貨(領用)", "確認庫存後，至PCHOME下載嘜頭，將出貨文件寄信給Emma。贈品領用都是寄銷(借貨單)，信件需提供領用單號及出貨清單，標題須標示領用單號。"]], "exceptions": ["約不到倉 → 回報業務，確認改到貨日。", "缺貨品項 → 回報業務，確認是否下修不出，並於PCHOME後台修正到貨數量。"]}, {"code": "SOP-SP-PG-001", "title": "蝦皮寄銷－竹運出貨流程", "brand": "PG", "platform": "蝦皮", "modes": ["寄銷"], "delivery": "竹運", "version": "V1.0", "date": "2025/12/08", "author": "林佳慧", "purpose": "建立標準化的蝦皮寄銷出貨流程，確保補貨作業正確、文件齊全並能有效追蹤，以降低錯誤率並提升作業效率。", "overview": "業務群組通知確認採購單 → 後台下載採購單 → 下採 → 匯出檔案 → 製作出貨檔案 → 核對竹運庫存及效期 → 系統拋單 → 發信", "roles": [["業務", "發送補貨通知、提供補貨明細"], ["企劃", "依照流程執行補貨、文件處理與系統拋單"]], "steps": [["步驟 1：收到業務通知", "確認通知內容。寶僑分母嬰跟洗劑線別，可能同天同倉但入庫單不同，須注意合併一起做檔。"], ["步驟 2：蝦皮後台下載採購單", "到蝦皮後台下載採購單，於補貨資料夾建立檔案。自行套轉換率，確認是否超出1200箱。"], ["步驟 3：下採EIP", "整理須補貨內容並完成下採。注意抓商品到貨日。複製上一個出貨資料夾，蝦皮特選-商品下採EXCEL貼入下採量(箱與小單位分開)，注意轉換率，貼入產品採購表上傳-箱單位/包單位，EIP下採。"], ["步驟 4：蝦皮後台下載入庫單", "已下採完成且入庫單已開立 → 到PG補貨資料夾開啟新資料夾，將補貨檔案貼入。注意倉別（蝦皮分觀音倉、安南倉(楊梅轉安南)）。"], ["步驟 5：查詢進倉單", "進蝦皮寄倉後台 → 入庫單 → 貼上採購單ID。"], ["步驟 6：匯出與列印PDF", "勾選入庫單 → 匯出Excel與PDF(入庫單與採購單) → 下載。"], ["步驟 7：製作出貨檔案", "將補貨excel資料貼入(入庫單號、料號、商品名稱、單位、數量)。確認入庫單是否被刪除，如有則入庫單也要一併刪除。"], ["步驟 8：核對庫存及做檔案", "填入並核對竹運庫存，確認無缺貨。若缺貨與業務確認下修數量，修改拋檔數量及出貨清單。"], ["步驟 9：系統拋單", "訂單系統 → 外部訂單 → 接收電子商城訂單 → 完成拋單。檔名依指定格式，確認拋單成功。"], ["步驟 10：發信及登錄", "將入庫單&出貨清單寄給相關人員，雲表登記B單出貨數。附上壓縮的入庫單&出貨清單、出貨規範。"]], "exceptions": []}, {"code": "SOP-YH-PG-001", "title": "雅虎寄銷－竹運出貨流程", "brand": "PG", "platform": "雅虎", "modes": ["寄銷", "領用"], "delivery": "竹運", "version": "V1.0", "date": "2025/09/22", "author": "杜孟涵", "purpose": "建立標準化的雅虎寄銷出貨流程，確保補貨作業正確、文件齊全並能有效追蹤，以降低錯誤率並提升作業效率。", "overview": "業務通知 → 建立補貨資料夾 → 後台查詢進倉單 → 匯出檔案 → 核對竹運庫存 → 系統拋單 → 自編碼確認 → 文件檢查 → 發信", "roles": [["業務", "發送補貨通知、提供補貨明細"], ["企劃", "依照流程執行補貨、文件處理與系統拋單"]], "steps": [["步驟 1：收到業務通知", "判斷補貨商品為一般品或幫寶適（帳號不同）。判斷正確帳號。"], ["步驟 2：建立補貨資料夾", "到雅虎補貨資料夾開啟新資料夾，將補貨檔案貼入。檔案命名需依規範。"], ["步驟 3：查詢進倉單", "進雅虎後台 → 倉儲作業 → 進倉單查詢/取消 → 使用指定進倉日與到貨日期區間查詢。"], ["步驟 4：匯出與列印PDF", "勾選進倉單 → 匯出Excel → 批次列印嘜頭 → 批次列印進倉單。確認進倉單單號及數量正確。"], ["步驟 5：核對庫存", "填入並核對竹運庫存，確認無缺貨。若缺貨立即回報業務。"], ["步驟 6：系統拋單", "訂單系統 → 外部訂單 → 接收電子商城訂單 → 完成拋單。"], ["步驟 7：自編碼確認", "若有自編碼，須加上自編碼請YAHOO協助貼標文字。若無則YAHOO不會主動貼標。"], ["步驟 8：文件檢查", "確認文件齊全（嘜頭、進倉單），若缺件需補齊。"], ["步驟 9：發信及登錄", "將嘜頭與進倉單發信給相關人員，並至雲表登記出貨。"], ["領用單出貨", "業務申請領用單(通常為贈品)。執行步驟3-步驟5，不須拋單，發信給Emma。信件需提供領用單號及出貨清單。"]], "exceptions": ["發現缺貨：立即回報業務，並更新補貨檔案。", "有效期問題無法出貨，需下修數量（重新拋單）或不出。", "缺貨需刪單(按僅顯示供應商可申請取消的進倉單即可刪除)，缺部分數量可不理會，整張單都缺才刪單。", "指定進倉日前後一天都可以到貨。"]}, {"code": "SOP-YH-PG-002", "title": "雅虎寄銷－直送出貨流程", "brand": "PG", "platform": "雅虎", "modes": ["寄銷", "領用"], "delivery": "直送", "version": "V1.0", "date": "2025/10/14", "author": "杜孟涵", "purpose": "建立標準化的雅虎寄銷出貨流程（廠商直送），確保補貨作業正確、文件齊全並能有效追蹤。", "overview": "業務通知 → 建立補貨資料夾 → 後台查詢進倉單 → 匯出檔案 → 製作採購單及匯入檔 → EIP上傳採購單 → 發信 → 核單確認缺貨刪單 → 下載進倉單及嘜頭 → 自編碼確認 → 文件檢查 → 發信", "roles": [["業務", "發送補貨通知、提供補貨明細"], ["企劃", "依照流程執行補貨、文件處理與系統拋單"]], "steps": [["步驟 1：收到業務通知", "判斷補貨商品為一般品或幫寶適（帳號不同）。"], ["步驟 2：建立補貨資料夾", "到雅虎補貨資料夾開啟新資料夾，將補貨檔案貼入。"], ["步驟 3：查詢進倉單", "進雅虎後台 → 倉儲作業 → 進倉單查詢/取消 → 指定進倉日與到貨日期區間查詢 → 勾選欲出貨單號 → 匯出Excel，與業務檔案核對數字。"], ["步驟 4：製作出貨文件", "製作上傳檔案，補貨資料填入採購單及匯入檔，匯入檔填入庫存跟均銷，至EIP上傳採購單。庫存及均銷：倉儲作業→新增進倉單(在庫量良品K欄/近30日銷量O欄)。用供應商料號來V，V進數字後要/轉換率。"], ["步驟 5：EIP系統拋單", "登入EIP → 申請表單 → 產品領用/採購 → 產品採購申請 → 完成拋單。"], ["步驟 6：發信", "將採購單發給相關人員。"], ["步驟 7：核單確認缺貨", "收到採購缺貨通知信 → 確認缺貨品項 → 至雅虎後台刪單。缺貨務必刪單，未到貨會產生罰款。"], ["步驟 8：匯出與列印PDF", "勾選進倉單 → 批次列印嘜頭 → 批次列印進倉單。"], ["步驟 9：自編碼確認", "若有自編碼，須加上自編碼請YAHOO協助貼標。"], ["步驟 10：文件檢查", "確認文件齊全（嘜頭、進倉單）。"], ["步驟 11：發信", "將出貨文件發給相關人員。"], ["領用單出貨", "贈品若從原廠直出，作業方式同寄銷流程，待採購壓到貨後業務開立領用單；信件標題須標示贈品領用。核單後確認缺貨，下載嘜頭、進倉單回信給採購。"]], "exceptions": ["有缺貨需刪單(按僅顯示供應商可申請取消的進倉單即可刪除)，缺部分數量可不理會，整張單都缺才刪單。", "指定進倉日前後一天都可以到貨。"]}, {"code": "SOP-YFS-EC-001", "title": "外倉盤點作業流程", "brand": "通用", "platform": "通用", "modes": [], "delivery": "－", "version": "V1.0", "date": "2025/04/21", "author": "杜孟涵", "purpose": "規範外倉盤點作業流程，確保庫存資料正確、盤點紀錄完整，並提升財務帳務及庫存管理之準確性與一致性。", "overview": "盤點規劃與通知 → 確認盤點時程與人員安排 → 執行外倉盤點作業（含寄銷倉）→ 盤點簽署與文件回收 → 差異處理與報表對帳", "roles": [["PM", "確認盤點人員安排"], ["營管", "統籌盤點規劃、協調財務與IT部門"], ["財務部", "陪同實盤及確認帳差原因並追蹤至結案"], ["IT部門", "盤點當日撈取Oracle系統庫存報表，協助對帳"], ["盤點人員(KA)", "確認盤點日期、執行現場盤點、確認實盤與帳面一致"]], "steps": [["步驟 1：盤點前準備作業", "約每年底財務通知啟動盤點。確認各外倉庫存及預計盤點月份、各通路須為「實盤」或「代盤」、向PM確認負責人員。避開大量進貨與月初/月底結帳期間；盤點當日進貨不列入清單。"], ["步驟 2：通知與預約作業", "發信通知負責人員：盤點日期時間、廠商進場人數、代盤費用、須提供之盤點正本資料（加蓋公司章/發票大小章並簽名）及庫存Excel。同通路多倉須同日盤點以利Oracle對帳。通知財務與IT部門系統準備。"], ["步驟 3：寄銷倉盤點作業", "公司所有寄銷存貨皆須列入盤點。寄銷倉管理人確保實際庫存與系統帳面一致，若有盤差由寄銷倉負責理賠。"], ["步驟 4：簽署與文件回收", "存貨保管人員與盤點人員共同於盤點明細表簽名，正本加蓋發票章或公司大小章，由盤點人員帶回統整後交財務留存。"], ["步驟 5：盤差處理與報表對帳", "IT於盤點當日上午匯出庫存報表；各倉針對差異品項提出說明與佐證；Oracle報表與實盤不符須提出差異說明（附進貨單、出貨單影本）。"]], "exceptions": ["盤點過程系統與實盤差距過大，立即通知PM與財務確認。", "外倉無法配合或人員異動影響時程，負責窗口應重新安排。"]}, {"code": "SOP-YFS-EC-002", "title": "加工組裝作業流程（含貼標改包）", "brand": "通用", "platform": "通用", "modes": [], "delivery": "－", "version": "V1.0", "date": "2025/08/22", "author": "杜孟涵", "purpose": "規範派工組裝流程，確保作業一致性與可追溯性，避免溝通落差與加工異常。", "overview": "業務填寫派工單 → 物流回覆報價及時程 → 業務提出流通加工申請 → 系統核准後物流執行 → 加工完成與回饋 → 業務確認結案", "roles": [["業務", "提出派工申請、確認報價與打樣、提出加工申請與結案確認"], ["物流", "回覆報價、安排加工作業、回覆進度與提供加工紀錄"], ["採購", "提供入庫單號、協助包材採購作業"]], "steps": [["步驟 1：業務派工申請", "業務填寫《永豐商店派工單》Sheet:組合加工單(FRM-YFS-EC-001)。須含加工需求、預計數量、是否打樣、包材需求，完成後提交物流。資料不完整物流得退回補件。"], ["步驟 2：物流回覆與確認", "回覆報價與打樣可行性，提供打樣時程。收到派工單後3個工作天內回覆。若包材不適用需與業務確認採購（決定包材尺寸、請CPG採購詢價）。"], ["步驟 3：加工申請與核准", "業務於EIP系統提出《流通加工申請》，系統簽核後由採購提供《入庫單號》，業務轉交物流執行並於訂單系統建立竹運主檔。"], ["步驟 4：加工執行與進度回覆", "物流依加工單安排作業，回覆預計完成日期。遇人力或包材異常即時通報。"], ["步驟 5：加工完成與回饋", "物流確認完成並通知業務結案，提供加工紀錄（異常/困難、耗損數量、改善建議）。紀錄須存檔備查。"], ["步驟 6：結案", "業務確認加工結果並結案，文件與紀錄歸檔（保存至少一年）。"], ["貼標/改包 步驟 1：業務派工申請", "業務填寫《永豐商店派工單》Sheet:貼標/大拆小加工單。須含加工需求、數量、是否貼標、標籤位置；貼標須提供「貼標示意圖」。示意圖需清楚標示位置避免誤貼。"], ["貼標/改包 步驟 2：物流評估報價", "物流評估作業內容（貼標或拆包）並提供報價，3個工作天內回覆。"], ["貼標/改包 步驟 3：業務確認與執行", "業務同意報價後回覆物流執行，物流依派工單貼標或拆包，業務於訂單系統建立竹運主檔。執行中發現數量或標籤異常即時回報。"]], "exceptions": ["加工過程包材不足或異常，物流立即通知業務協調。", "加工進度延誤超過預期一週以上，須回報主管與業務。", "貼標示意圖不清楚，物流退回業務確認修正。", "加工數量與派工單不符，立即回報業務確認後再作業。", "業務需確認竹運主檔是否已建立完成。"]}, {"code": "SOP-YFS-EC-003-逆物流", "title": "逆物流作業流程（退貨）", "brand": "通用", "platform": "通用", "modes": ["寄銷", "買斷"], "delivery": "－", "version": "V1.0", "date": "2025/08/22", "author": "杜孟涵", "purpose": "建立逆物流（退貨）作業之標準流程，確保退貨資訊完整、處理正確、部門協作順暢，並避免對帳與庫存產生異常。涵蓋寄銷、買斷、進貨驗退、廠退、消費者退貨。", "overview": "退貨分類 → 進貨驗退 → 退廠處理 → 消費者退貨 → 對帳與異常處理", "roles": [["企劃專員", "登錄退貨資訊、申請系統退貨、跨部門協調、處理對帳更正"], ["物流", "收退作業、驗入退貨、寄回原廠或倉庫、退貨通知"], ["業務", "確認通路異常、申請退廠"], ["採購", "處理原廠直出訂單的拒收/退回與折讓、調整進貨數量"], ["財務", "折讓、發票開立、對帳異常處理"]], "steps": [["步驟 1：退貨分類", "退貨依來源及通路區分：寄銷（進貨驗退/廠退/客直退）、買斷（進貨驗退/消費者轉單退貨:宅配退貨、超取退貨）。需先確認通路別與商品屬性，判斷後續作業方式。"], ["步驟 2：進貨驗退流程", "【寄倉通路-竹運】收貨倉庫確認異常→拒收驗退→物流通知業務→企劃登錄「公槽退貨表單」並回覆物流驗入。【寄倉-原廠直出】拒收驗退直接退回原廠，採購調整進貨數量並開立發票或折讓單。\n【買斷-竹運】現場拒收：物流通知業務，酷澎專員於訂單系統辦理「部分退貨」，物流退貨驗入，月底依財務對帳資料進行交易更正。後續驗退：驗入發現異常→發驗退通知信→業務確認後酷澎專員確認退回方式。\n若派車收退，信件需附收退商品資訊、收退倉庫資料、廠商自取單。"], ["步驟 3：退廠流程", "在庫商品異常：通路倉庫申請退廠→企劃登錄公槽退貨表單→確認退回方式與運費付款方式(到付)→信件通知物流驗收。業務申請退廠：業務於通路後台申請→通知企劃→同上流程。"], ["步驟 4：消費者（轉單）退貨流程", "【宅配退貨】企劃於訂單系統申請退貨→系統產生退貨單→宅配收退→倉庫驗入→結案。【店配退貨】未取退回：檢查退貨雲端表格→至平台確認→訂單系統申請退貨→驗入結案。派車收退：通路後台接收退貨申請→訂單系統申請退貨(超取轉宅配收退)→輸入收退資料→系統產生退貨單→宅配收退→驗入結案。"]], "exceptions": ["特殊退貨須由業務、企劃、物流、財務跨部門確認後執行。", "涉及帳務異常由財務協助折讓、發票開立。", "派車收退需附收退商品資訊、倉庫資料與廠商自取單。", "所有退貨資料須即時登錄於「公槽退貨表單」，確保各部門同步掌握。"]}, {"code": "SOP-YFS-EC-003-寄銷退貨", "title": "寄銷通路退貨作業流程", "brand": "通用", "platform": "通用", "modes": ["寄銷"], "delivery": "－", "version": "V1.0", "date": "2025/08/22", "author": "杜孟涵", "purpose": "建立逆物流（退貨）作業之標準流程，確保退貨資訊完整、處理正確、部門協作順暢，並避免對帳與庫存產生異常。適用於寄銷通路之廠退與客直退流程。", "overview": "通路退貨資訊蒐集 → 退貨資料登錄（企劃）→ 採購開立入庫單 → 物流驗收入庫 → 企劃資料檢核與補登 → 財務與企劃月底對帳", "roles": [["企劃專員", "每日確認各通路退貨資訊並登錄、檢查退貨雲端表格並追蹤異常、協助財務釐清對帳"], ["採購", "每日依退貨資訊開立入庫單"], ["物流", "依入庫單驗收入庫、填寫退貨商品明細至雲端"], ["財務", "月底進行退貨對帳、異常轉交企劃確認"]], "steps": [["步驟 1：退貨資訊登錄（企劃專員）", "需可登入各寄銷通路後台。每日確認退貨資料並登記至公槽退貨表單。資料須當日更新避免延遲開單入庫。"], ["步驟 2：入庫單開立（採購）", "企劃已登錄退貨資料後，每日依退貨資料開立退貨入庫單。需確保入庫單品項正確。"], ["步驟 3：退貨驗收入庫（物流）", "已收到實體退貨商品及入庫單，依入庫單驗收入庫並登錄明細至雲端表格。若品項不符需立即回報。"], ["步驟 4：資料檢核與補登（企劃專員）", "檢視雲端表格確認每筆退貨皆有對應入庫單號。若空白須判斷（確認退貨來源、入庫單號是否有誤、是否異常、是否已登記）。若無入庫單須補登記並於業務欄回填「已登記待開單」，物流再依補登資料驗入。"], ["步驟 5：對帳作業（財務 & 企劃）", "財務部月底進行退貨對帳、檢核退貨對帳單。若資料有問題轉交專員確認，專員檢視客戶對帳退貨資料確認異常原因並追蹤。"]], "exceptions": ["退貨資料與實際退貨不符 → 物流回報企劃與採購。", "對帳資料有疑義 → 財務通知企劃釐清。", "退貨商品遺失或短少 → 依公司異常品項流程處理。"]}, {"code": "YFS-EC-SOP-001", "title": "EIP 申請庫存調整步驟", "brand": "通用", "platform": "通用", "modes": [], "delivery": "－", "version": "V1.0", "date": "—", "author": "—", "purpose": "規範於永豐商店 EIP 系統申請庫存調整（盤盈／盤虧／移倉）之操作步驟；常見情境如退貨入不良品倉後之盤盈調整。", "overview": "進入 EIP 申請表單 → 結帳/促銷/請款 → 庫存調整申請單 → 選擇異動類型與異動類別 → 選擇異動倉庫 → 填寫庫存差異調整報表 → 送出申請", "roles": [], "steps": [["步驟 1：進入申請表單", "登入永豐商店 EIP，於首頁點選「申請表單」。\n（首頁功能：申請表單、簽核相關、差勤相關、差旅報支單、報表系統）"], ["步驟 2：選擇結帳/促銷/請款 → 庫存調整申請單", "申請表單 → 結帳/促銷/請款 → 庫存調整申請單。\n（申請表單分類：資訊類、庶務類、結帳/促銷/請款、產品領用/採購、TP費用相關、外倉庫存/銷售；結帳/促銷/請款下含：結帳申請、促銷價格申請、請款申請、發票管理系統、庫存調整申請單）"], ["步驟 3：填寫異動類型與倉庫", "一、庫存調整類型：選擇 盤盈／盤虧／移倉（本例為盤盈）。\n二、異動類別：一般倉庫庫存調整／外倉結帳庫存不足／宅配通儲位調整／宅配通出貨未結庫存不足／(整新/改包)差異。\n三、異動倉庫：因退貨入不良品，故選「YSL_竹運不良品倉」（不良品倉）。"], ["步驟 4：填寫庫存差異調整報表", "1. 料號：貼上要盤盈的料號（品名由系統自動帶出）。\n2. 單位：選擇要盤盈的單位。\n3. 調整後數量：填入調整後的總量 — 例如要盤盈 1、系統庫存為 10，則填入 11。\n4. 系統庫存數量：系統自動帶出目前庫存（例如 10）。\n5. 調整差異數量：此欄自動顯示（11−10＝1），代表要調整盤盈的數量。\n6. 備註：視需要填寫。\n完成後點「新增」增列品項，確認無誤後送出申請。"]], "exceptions": []}, {"code": "SOP-BK-CPG-001", "title": "博客來－直送出貨作業流程", "brand": "CPG", "platform": "博客來", "modes": ["寄銷"], "delivery": "直送", "version": "V1.0", "date": "2025/11/27", "author": "柯秋依", "purpose": "建立標準化的博客來出貨流程（原廠直送），確保入庫、補貨與拋檔作業正確，並降低錯誤率。", "overview": "業務通知 → 建立資料夾與出貨文件 → 收取訂單與匯出明細 → 製作採購單及匯入檔 → EIP上傳採購單 → 發信 → 核單確認缺貨 → 修改出貨數量 → 列印採購單與出貨資料 → 列印外箱嘜頭 → 發送出貨文件", "roles": [["業務", "提供補貨資料，指定到貨日期"], ["企劃(OP)", "依據流程處理入庫、出貨單及系統拋檔"]], "steps": [["步驟 1：建立出貨資料夾與出貨文件", "建立新資料夾（以出貨日命名），複製上次出貨資料並更新為本次檔案。確保檔名格式正確。"], ["步驟 2：製作出貨文件", "製作採購單上傳表及入庫資料表，貼上各對應欄位。"], ["步驟 3：填寫出貨資料", "至 [進貨收單作業]-[收取訂單及填寫出貨資料]-[填寫出貨資料] → 勾選採購單號 → 點選[出貨數量] → 確認品項數量正確 → 儲存。"], ["步驟 4：填寫採購單", "點選 [列印採購單] → [填寫出貨資料]。列印後確認採購單內容正確。"], ["步驟 5：列印外箱嘜頭", "至 [列印外箱嘜頭] → 查詢 → 選擇出貨單號 → 列印嘜頭下載PDF。"], ["步驟 6：EIP系統拋單（OP）", "登入EIP → 申請表單 → 產品領用/採購 → 產品採購申請 → 完成拋單。"], ["步驟 7：發信（OP）", "將採購單發給相關人員，確認收件人及出貨文件正確。"]], "exceptions": []}, {"code": "SOP-MO-CPG-001", "title": "MOMO 寄銷出貨流程", "brand": "CPG", "platform": "MOMO", "modes": ["寄銷"], "delivery": "直送", "version": "V1.0", "date": "2025/10/14", "author": "柯秋依", "purpose": "建立標準化的momo寄銷出貨流程，確保補貨作業正確、文件齊全並能有效追蹤，以降低錯誤率並提升作業效率。", "overview": "業務通知 → 製作上傳表及出貨檔案入庫憑單 → 上傳EIP系統 → 發信給採購 → 待採購提供商品效期，後台抓取嘜頭跟二聯單 → 發信出貨文件與共網登記", "roles": [["業務", "發送補貨通知、提供補貨明細"], ["企劃(OP)", "依照流程執行補貨、文件處理與系統拋單"]], "steps": [["步驟 1：建立出貨資料夾與匯入檔", "確認通知內容、補貨數量及到貨日。判斷檔案、料號是否正確。"], ["步驟 2：建立補貨資料夾", "到momo補貨資料夾開啟新資料夾，放進補貨檔案，製作採購單上傳表及入庫憑單（一個入庫單號一張入庫憑單）。"], ["步驟 3：發信通知採購", ""], ["步驟 4：製作入庫通知單及嘜頭", "採購通知商品效期後，填入效期、下載單據。"], ["步驟 5：自編碼下載", "momo後台 → 搜尋入庫單號 → 勾選對應品項 → 列印條碼5x12 → 儲存PDF。若有自編碼，業務信件必須通知。"], ["步驟 6：上傳EIP系統與通知採購", "EIP申請表單 → 產品領用/採購 → 產品採購申請。\n需求日期→到貨日；採購單位→補貨平台；供應商名稱→紙／潔／紙潔；平台類型→寄銷倉；寄銷倉庫→富邦媒體寄銷倉(MOMO)；幣別→新台幣；總費用→自動帶出。"], ["步驟 7：文件檢查", "確認文件齊全（嘜頭、入庫通知單、自編碼）。"], ["步驟 8：發信及登錄", "將嘜頭與入庫通知單發信給相關人員，並至雲表登記作業時間。"]], "exceptions": ["若缺貨、運能滿，採購會另外通知，業務須改量、改期、改倉等。"]}, {"code": "SOP-PC-CPG-001", "title": "PCHOME 寄銷－原廠直出流程", "brand": "CPG", "platform": "PCHOME", "modes": ["寄銷"], "delivery": "直送", "version": "V1.0", "date": "2025/10/29", "author": "柯秋依", "purpose": "規範PCHOME出貨作業流程，確保寄銷出貨執行一致性，降低錯誤並提升作業效率。", "overview": "業務補貨通知 → 製作上傳表及借貨檔 → 上傳EIP系統 → 發信給採購 → 待採購提供CPG單號，後台抓取嘜頭&入庫通知單 → 發信出貨文件與共網登記", "roles": [["業務", "提供補貨資料、確認到貨日"], ["採購", "下採單、發信通知缺貨"], ["企劃(OP)", "建資料夾、製作/修改採購單與上傳檔、上傳EIP系統、下載嘜頭、通知出貨"]], "steps": [["步驟 1：接收補貨通知並約倉", "確認通知內容、補貨數量及到貨日。"], ["步驟 2：建立資料夾", "到pc補貨資料夾開啟新資料夾，放進補貨檔案，製作採購單上傳表及借貨單。"], ["步驟 3：製作EIP上傳表", "依業務補貨資料貼入上傳表，填入B、C、D、E、F、G欄位。"], ["步驟 4：製作借貨單", "依業務補貨資料貼入借貨單（全部欄位皆需要）。"], ["步驟 5：上傳EIP系統與通知採購", "EIP申請表單 → 產品領用/採購 → 產品採購申請。留意迴轉率過高狀況；選擇正確寄銷倉別。\n需求日期→到貨日；採購單位→補貨平台；供應商名稱→紙／潔／紙潔；平台類型→寄銷倉；寄銷倉庫→PCHOME；幣別→新台幣；總費用→自動帶出。"], ["步驟 6：發信通知採購", "嘜頭和二聯單：潔品單品一張、箱裝一張；紙品則需等採購通知CPG單號後抓。"], ["步驟 7：抓嘜頭和二聯單單據", "採購通知CPG單號後：新增空白頁貼入補貨表料號、單位、數量、借貨單號；回原Sheet插入空白F欄=TEXT(料號,\"0000000\")，K欄=VLOOKUP(F欄,空白頁,4,0)，依訂購單號分類抓單。\n進PChome後台 → 寄倉入庫管理 → 借貨單 → 庫別 → 筆數選500全部顯示 → 用借貨單號尋找並核對料號、數量、入庫日 → 勾選下載。"], ["步驟 8：發信給採購", ""]], "exceptions": []}, {"code": "SOP-SP-CPG-001", "title": "蝦皮寄銷－竹運出貨流程", "brand": "CPG", "platform": "蝦皮", "modes": ["寄銷"], "delivery": "竹運", "version": "V1.0", "date": "2025/10/03", "author": "張富淳", "purpose": "建立標準化的蝦皮寄銷出貨流程，確保補貨作業正確、文件齊全並能有效追蹤，以降低錯誤率並提升作業效率。", "overview": "業務通知 → 建立補貨資料夾 → 製作竹運B單 → 下載入庫單 → 核對竹運庫存 → 系統拋單 → 文件檢查 → 發信", "roles": [["業務", "發送補貨通知、提供補貨明細"], ["企劃", "依照流程執行補貨、文件處理與系統拋單"]], "steps": [["步驟 1：收到業務通知", "確認通知內容、補貨數量及到貨日。判斷檔案、料號是否正確。"], ["步驟 2：建立補貨資料夾", "到蝦皮補貨資料夾建立新資料夾。竹運B單檔名依物流規定。"], ["步驟 3：製作竹運B單", "進蝦皮寄倉 → 庫存資訊 → 入庫單 → 點選採購單ID → 查詢 → 匯出Excel → 加工製作竹運B單。確保料號、補貨數量一致。"], ["步驟 4：下載入庫單", "進蝦皮寄倉 → 庫存資訊 → 入庫單 → 點選採購單ID → 查詢 → 匯出PDF入庫單。進倉單號料號及數量要與補貨明細一致，注意國條是否弄成料號。"], ["步驟 5：核對竹運庫存", "訂單系統 → 庫存管理 → 每日庫存結餘查詢。若數量不夠與業務確認可出數量，再通知蝦皮窗口調整入庫單數量（入庫單需重下載）。"], ["步驟 6：系統拋單", "登入訂單系統 → 外部訂單 → 接收電子商城訂單 → 完成拋單。檔名依指定格式，確認拋單成功。"], ["步驟 7：文件檢查", "確認文件齊全，若缺件需補齊。"], ["步驟 8：發信及登錄", "將入庫單發信給相關人員，並至雲表登記作業。務必備註「一早到貨」。"]], "exceptions": ["改期及缺貨都需通知蝦皮更新入庫單。", "任何變動都須於入庫日2個工作日前提出，避免罰款。"]}, {"code": "SOP-UB-CPG-001", "title": "優食（UberEats）－直送買斷出貨流程", "brand": "CPG", "platform": "UberEats", "modes": ["買斷"], "delivery": "直送", "version": "V1.0", "date": "2025/11/25", "author": "黃千容", "purpose": "建立標準化的優食直出出貨流程，確保補貨作業正確、文件齊全並能有效追蹤，以降低錯誤率並提升作業效率。", "overview": "業務通知 → 建立補貨資料夾 → 製作訂單 → 核對下單金額 → EIP下採 → 文件檢查 → 發信", "roles": [["業務", "確認各店家補貨通知、提供補貨明細"], ["企劃", "依照流程執行補貨、文件處理與通知匯入"]], "steps": [["步驟 1：收到業務通知", "確認通知內容、補貨數量及到貨日。核對各店下單商品、數量、料號是否正確。"], ["步驟 2：建立補貨資料夾", "於優食好物路徑新增到貨資料夾，貼入統倉出貨訂單，製作採購單及EMMA匯入文件。CPG出貨是出至優食統倉後再自行分給各店；買斷訂單最後到貨日為最後四個工作天，故月底不出貨以免無法對帳。"], ["步驟 3：製作補貨文件", "將各分店訂單貼至「複本 UBER訂單匯入-統倉價格」→ 確認數量及金額 → 確認下單料號。下單金額異常需與業務確認後才能進行下一步。"], ["步驟 4：EIP下採", "EIP申請表單 → 產品領用/採購 → 產品採購申請。日期及下採數量正確。"], ["步驟 5：提供Emma匯入文件", "優食訂單匯入表格填寫 → 產生並提供採購申請單號PO單後發信。日期、下採數量及料號要正確。"], ["步驟 6：發信", "將優食訂單匯入單發信給相關人員，確認收件人及出貨文件正確。"]], "exceptions": ["若有新舊料號，需通知EMMA取消後再提供正確文件。"]}, {"code": "SOP-YH-CPG-001-東森", "title": "東森－直送寄銷出貨流程", "brand": "CPG", "platform": "東森", "modes": ["寄銷"], "delivery": "直送", "version": "V1.0", "date": "2025/10/29", "author": "張富淳", "purpose": "建立標準化的東森寄銷出貨流程，確保補貨作業正確、文件齊全並能有效追蹤，以降低錯誤率並提升作業效率。", "overview": "業務通知 → 建立補貨資料夾 → 製作入庫指示書 → EIP下採 → 文件檢查 → 發信", "roles": [["業務", "發送補貨通知、提供補貨明細"], ["企劃", "依照流程執行補貨、文件處理與系統拋單"]], "steps": [["步驟 1：收到業務通知", "確認通知內容、補貨數量及到貨日。判斷檔案、料號是否正確。"], ["步驟 2：建立補貨資料夾", "到東森補貨資料夾開啟新資料夾，貼入補貨檔案，製作採購單及出貨清單。採購單號料號及數量要與補貨明細一致；原單號數量不夠時業務會開第二張單號，下採數量只能少不能多；採購單號有兩個時反黃標示提醒。"], ["步驟 3：製作入庫指示書", "進東森後台 → 進出貨管理 → 商品入庫管理 → 入庫指示作業 → 入庫指示 → 拉至底部查看資料筆數 → 每頁顯示輸入總數量 → 查詢 → 依KA提供的採購單號逐筆查詢勾選 → 列印入庫指示書 → 儲存。"], ["步驟 4：EIP下採", "EIP申請表單 → 產品領用/採購 → 產品採購申請。日期及下採數量正確。"], ["步驟 5：文件檢查", "確認文件齊全（出貨憑單、入庫指示書），若缺件需補齊。"], ["步驟 6：發信及登錄", "將嘜頭與進倉單發信給相關人員，並至雲表登記作業時間。"]], "exceptions": ["若缺貨，採購會通知業務，雲端紀錄即可。"]}, {"code": "SOP-YH-CPG-001-雅虎", "title": "雅虎－直送寄銷出貨流程", "brand": "CPG", "platform": "雅虎", "modes": ["寄銷"], "delivery": "直送", "version": "V1.0", "date": "2025/10/02", "author": "張富淳", "purpose": "建立標準化的雅虎寄銷出貨流程，確保補貨作業正確、文件齊全並能有效追蹤，以降低錯誤率並提升作業效率。", "overview": "業務通知 → 建立補貨資料夾 → 製作進倉單及嘜頭 → 自編碼下載 → EIP下採 → 文件檢查 → 發信", "roles": [["業務", "發送補貨通知、提供補貨明細"], ["企劃", "依照流程執行補貨、文件處理與系統拋單"]], "steps": [["步驟 1：收到業務通知", "確認通知內容、補貨數量及到貨日。判斷檔案、料號是否正確。"], ["步驟 2：建立補貨資料夾", "到雅虎補貨資料夾開啟新資料夾，貼入補貨檔案，製作採購單及進貨憑單。進貨憑單到庫日期及箱數需一致，下採數量只能少不能多。"], ["步驟 3：製作進倉單及嘜頭", "進雅虎後台 → 倉儲作業 → 進倉單查詢/取消 → 狀態選未結案查詢 → 搜尋進倉單號。進倉單號料號及數量要與補貨明細一致。"], ["步驟 4：自編碼下載", "雅虎後台 → 搜尋進倉單號 → 找到對應打勾 → 上方「批次列印條碼」→ 儲存標籤檔案。若有自編碼，業務信件必須通知。"], ["步驟 5：EIP下採", "EIP申請表單 → 產品領用/採購 → 產品採購申請。日期及下採數量正確。"], ["步驟 6：文件檢查", "確認文件齊全（嘜頭、進倉單、進貨憑單、自編碼），若缺件需補齊。"], ["步驟 7：發信及登錄", "將嘜頭與進倉單發信給相關人員，並至雲表登記作業時間。"]], "exceptions": ["若缺貨，採購會通知業務，可雲端紀錄。", "指定進倉日前後一天都可以到貨；如有變動則入庫單需重開重抓。"]}, {"code": "SOP-ET-EC-001", "title": "東森－退廠作業流程", "brand": "跨品類", "platform": "東森", "modes": [], "delivery": "－", "version": "V1.0", "date": "2025/12/02", "author": "朱昱靜", "purpose": "規範東森平台退廠作業流程，使作業標準化、準確且有效率，並確保退廠紀錄完整與物流安排順暢。", "overview": "收到退廠通知 → 回壓退廠日期 → 等待東森回覆 → 系統退廠作業 → 下載退廠單PDF → 發信給東森物流 → 確認回覆 → 登記內部資料 → 退貨異常處理", "roles": [["企劃", "退廠流程操作、資料記錄與異常處理"], ["物流聯繫窗口", "派車安排、自取協作"], ["倉庫人員", "接收退回商品與確認狀況"], ["東森窗口", "退廠通知"]], "steps": [["步驟 1：收到東森退廠信件", "確認信件內容。信件日期為作業起算依據。"], ["步驟 2：回壓退廠日期", "例如星期二收到 → 回壓下星期一退廠。需等待東森回覆確認。"], ["步驟 3：東森回信後開始作業", "可進行系統退廠程序。"], ["步驟 4：登入東森系統操作", "登入 → 進出貨管理 → 商品退廠管理 → 退廠單列表 → 勾選退廠單 → 選擇回收方式（自取／物流寄送）。\n物流寄送收件資料：地址 桃園市觀音區新富路889號；收件人 馮世昀、李淑婷；電話 03-2866577。\n自取需提前派車：派車費3000元／新竹貨運到付150元／超過20件建議派車；需發信給物流確認可派車時間；自取務必攜帶「退廠單」依預退日至退廠碼頭領取。"], ["步驟 5：輸出退廠單PDF", "退廠單狀態「廠商確認」→ 查詢 → 轉出PDF。"], ["步驟 6：發信給東森物流", "附上退廠單PDF。"], ["步驟 7：確認東森回覆", "確認退廠安排。"], ["步驟 8：內部資料登記", "路徑：\\10.37.90.10\\永豐商店\\18.客戶資料\\8E平台退貨\\2025 → 匯出EXCEL → 貼上退倉資訊。"]], "exceptions": ["商品異常 → 留意退貨異常登記表並回覆處理。", "東森聯絡窗口：陳淑婷 (02)2943-7888#7021 / rene.chen@ehsn.com.tw；楊梅二倉 03-4963068、03-4962859#1117/1118；傳真 03-4962905、03-4962855。"]}, {"code": "SOP-GH-EC-001", "title": "遠時（GoHappy）－請款作業流程", "brand": "跨品類", "platform": "遠時", "modes": [], "delivery": "－", "version": "V1.0", "date": "2026/06/26", "author": "張富淳", "purpose": "建立標準化的遠時對帳及上傳發票作業流程，提升一致性與效率。適用於遠時與公司財務OP帳務核對、上傳請款發票及書信往來確認。", "overview": "遠時每月結帳兩次（月初7號前、月底28號前）→ 下載對帳單讓財務OP核對 → 財務OP開立電子發票 → 每月9號前上傳發票 → 審核通過遠時當月25號完成付款 → 下載費用發票給財務OP作帳", "roles": [["財務OP", "依遠時對帳資料開立請款發票"], ["企劃(OP)", "結帳時間內提供財務OP結帳資料，並於時效內上傳請款發票"]], "steps": [["步驟 1：結帳作業", "對帳 → 對帳與結帳 → 看試算總計 → 去結帳 → 列印（儲存PDF）。確認無誤再按，若有異常反應給業務及遠時財務。"], ["步驟 2：存檔結帳資料至公槽", "建立當月請款資料夾（命名：上次請款後端日期隔天–這次結帳最後日期）→ 抓取對帳資料給財務OP → 下載出貨(應稅)及退貨(應稅) → 合併成一個excel → 更改檔名存檔（供應商對帳總表_各費用明細）。路徑：轉單平台\\02. Friday(GOHappy)\\對帳資料。"], ["步驟 3：將對帳資料提供給財務OP核對", "MAIL通知財務OP及相關業務，並CC各線同仁（品牌:CPG紙潔／P&G寶僑／Mars瑪氏）。信件主旨:GH對帳單日期區間；內容:結帳日及區間／結帳單PDF／費用明細excel。"], ["步驟 4：上傳發票等請款資訊", "財務OP核對完畢開立電子發票後：儲存發票資訊到公槽 → 遠時後台上傳發票 → 對帳 → 憑證及請款作業 → 憑證 → 輸入資料確認無誤儲存。發票類型為電子發票，發票日期、號碼、金額稅額務必一致。"], ["步驟 5：通知遠時已完成請款", "MAIL遠時供應商請款信箱，附件夾帶結帳單及發票（建議勾選傳送及讀取回條）。登錄完成→9號遠時財務作業，至後台結算系統查看狀態直到「已完成審核」。"], ["步驟 6：提供費用發票給財務OP", "結帳完約2–3小時下載費用發票 → 下載發票及明細 → MAIL給財務OP。費用發票有誤請聯繫遠時財務Mica徐小姐#13934。"]], "exceptions": ["遠時財務部：電話 (02)7712-3838，聯絡人 Mica徐小姐#13934，Email mica_hsu@friday.tw。", "發票寄送：invoice@friday.tw（主旨：供應商請款【供應商編號／名稱】）。"]}, {"code": "SOP-SP-EC-001-問答卡片", "title": "蝦皮－問答卡片（客服 FAQ）設定作業流程", "brand": "跨品類", "platform": "蝦皮", "modes": [], "delivery": "－", "version": "V1.0", "date": "2025/12/03", "author": "朱昱靜", "purpose": "制定蝦皮聊聊問答卡片設定流程，統一客服自動回覆內容，提升訊息回覆效率與一致性，並降低人工客服負擔。", "overview": "收集常見問題 → 整理成標準問答 → 編輯成問答卡片格式 → 上傳至蝦皮後台（聊聊自動問答設定）→ 定期檢查並更新內容", "roles": [["企劃", "編輯、更新問答卡片內容，確保資訊正確；依卡片內容補充回覆客戶、處理非自動回覆範圍問題"]], "steps": [["步驟 1：整理常見問題與標準回覆", "將買家最常詢問的五大類問題整理標準話術：\n1. 出貨時間相關：下單隔天約2天出貨、出貨後約2–3工作天配達；急單勿下單、訂單成立無法修改、超商店到店超材會取消、離島不配送（宅配用新竹物流）；客服時間 週一至週五 10:00–17:00。\n2. 收到商品數量不對：連假貨量大可能延遲或分批，請先透過聊聊確認，勿直接申請退貨。\n3. 商品破損或漏液：請點「與賣家客服聊聊」並提供商品照片，破損商品及外箱勿丟棄以利索賠。\n4. 離島配送問題：目前賣場無提供離島宅配服務。\n5. 出貨是否附單據/發票：出貨不附金額單據或發票，發票由蝦皮開立電子發票。"], ["步驟 2：將問答內容設定至蝦皮後台", "登入蝦皮賣家中心 → 聊聊設定 → 問答卡片(FAQ) → 新增問答 → 將1–5筆內容依序貼入 → 儲存並啟用 → 測試是否正常自動發送。回覆內容須禮貌、清楚、可讀性強，並不定期檢查是否需更新（物流時效、政策變動、活動期間等）。"]], "exceptions": ["問答卡片未自動彈出 → 檢查蝦皮後台是否有啟用。", "客戶反應無法理解 → 客服需補充解說。", "內容過期（促銷、物流延遲）→ 立即更新話術。"]}, {"code": "SOP-SP-EC-001-大宗", "title": "蝦皮－大宗訂單作業流程", "brand": "跨品類", "platform": "蝦皮", "modes": [], "delivery": "－", "version": "V1.0", "date": "2025/12/02", "author": "朱昱靜", "purpose": "規範蝦皮大宗訂單作業流程，確保備貨、出貨與物流安排一致且正確。", "overview": "確認到貨日、品項、數量、價格、出貨資訊 → 先回壓出貨（蝦皮備貨天數2天）→ 拋單作業與建價日期設定 → 系統訂單備註 → 寄信給物流並提供專車配送需求及訂單資料", "roles": [["企劃", "與業務確認所有訂單條件、回壓出貨、系統處理、備註撰寫"], ["物流窗口", "安排專車、確認配送日期、回覆配運狀態"], ["倉庫", "依指示備貨並分板，留意效期擺放"]], "steps": [["步驟 1：確認到貨日並回壓出貨", "確認訂單條件（到貨日、品項、數量、價格、出貨資訊）後回壓蝦皮到貨日期。因蝦皮備貨天數為2天，需先回壓出貨，不得延誤避免延遲訂單。"], ["步驟 2：拋單與建價設定", "建價截止日設定為拋單當天日期；訂單收件者名稱需加「專車-XXX」。建價日期不得錯誤避免異常。"], ["步驟 3：系統訂單備註與通知物流", "系統備註「X月X日專車送，請勿宅配出貨」，避免物流誤操作宅配出貨。"], ["步驟 4：拋單後通知物流", "Email通知物流訂單編號與出貨資料（需D+3）；備註「請司機幫忙拉貨到對方棧板」，一板放一個效期。"]], "exceptions": ["顧客要求變更到貨日：需立即同步物流與倉庫。", "訂單箱數與品項若與下單實際不符：回報業務與顧客確認。"]}, {"code": "SOP-SP-EC-001-未驗入", "title": "蝦皮－店配訂單未驗入作業流程", "brand": "跨品類", "platform": "蝦皮", "modes": [], "delivery": "－", "version": "V1.0", "date": "2025/10/13", "author": "杜孟涵", "purpose": "規範蝦皮超取訂單未驗入時之處理流程，確保包裹狀態可即時追蹤並妥善完成後續盤點與庫存調整，避免遺失造成營運損失。", "overview": "發現超取未驗入 → 聯繫蝦皮社群小編 → 提供相關出貨與簽收資料 → 等候蝦皮確認包裹狀態 → 若確認遺失 → 業務同意後執行假退入庫與庫調申請", "roles": [["OP", "發現異常並主導處理，與蝦皮及物流窗口溝通、提出庫調申請"], ["物流", "協助提供出貨紀錄、簽收單及出貨影像資料"], ["業務", "確認遺失狀況並核准後續假退入庫及庫存調整"], ["IT", "協助執行訂單異動及系統調整"]], "steps": [["步驟 1：發現異常", "OP每日於蝦皮後台檢視是否有異常訂單，確認未驗入時間已超過合理範圍（如2–3日）。需截圖保留異常證明。"], ["步驟 2：聯繫蝦皮確認包裹狀態", "即時聯繫蝦皮社群小編，依要求提供：配達理貨中心（桃園或嘉義）、配達日期時間、清晰簽收單、出貨影片或包裹外箱/商品照片、寄送箱數、包裹總件數、同批已刷入之訂單編號、配合物流商、同日是否配送至其他物流中心。資料須清楚完整；影像文件留存至結案後至少三個月。"], ["步驟 3：確認包裹狀態", "等候蝦皮確認包裹是否尋獲或無此包裹；若確認無此包裹，後續依盤虧作業處理。未回覆超過三日需主動追問。"], ["步驟 4：執行盤虧及庫調作業", "申請訂單異動、執行「假退入庫」（由IT協助調整訂單）；OP於EIP提交庫存調整申請單：EIP → 申請表單 → 結帳/促銷/請款 → 庫存調整申請單，詳述庫調原因並附佐證。注意：庫調單「數量」欄應填調整後實際庫存數量（非盤虧/盤盈數量）；確認業務及主管核准後方可送出。"]], "exceptions": ["若物流無法提供簽收單或影像，須與物流釐清歸責。", "若蝦皮拒絕受理異常申報，須通報業務與窗口協調。"]}, {"code": "SOP-CP-CP-001-PO確認", "codeShow": "SOP-CP-CP-001", "title": "酷澎 PO 單確認作業流程", "brand": "跨品類", "platform": "酷澎", "modes": ["買斷"], "delivery": "－", "version": "V2.0", "date": "2026/07/23", "author": "簡孝真", "purpose": "酷澎 PO 單狀態由建立至確認(取消)之作業流程，提升一致性與效率。適用於酷澎後台執行 PO 訂單確認(收單)與取消相關作業。", "overview": "平台建立PO(判別新單/舊單/虛擬) → 檢視PO類型(一般/NS收單時效不同) → 修改退貨資訊 → 新PO匯入我司酷澎採購入庫處理系統 → 通知同仁分線別拉取 → PO彙總統合 → 依回覆於時效內更改PO狀態及內容(日期/數量/倉別) → 後台公告(嘜頭/進倉規範/倉庫位置)更新需通知相關同仁", "roles": [["業務/專員", "確認訂單明細調整與整合需求"], ["企劃(OP)", "執行平台後台系統操作，檢核執行PO並避免逾時確認(取消與確認PO)"]], "steps": [["步驟 1：下載新建立訂單並拋入系統", "酷澎訂單管理 → 物流 → PO List → 勾選PO狀態「已建立」下載 → 將新PO匯入我司酷澎採購入庫處理系統。\n收單時效：一般訂單 48 小時內、NS 訂單 96 小時內，逾時系統自動取消。品牌:CPG紙潔／P&G寶僑／Mars瑪氏。"], ["步驟 2：彙整PO和搜尋訂單", "至公槽記錄所有開出的PO → PO總表(PO ID／開單日／倉別／品牌／交貨日)。平台查詢：PO日期=建立日、預計交貨日期=到貨日、FC=交貨倉別、PO ID=查詢單號。"], ["步驟 3：檢閱訂單明細", "單一頁面含：PO單號、基本資訊(訂單狀態)、供應商資訊、退貨資訊(可編輯)、入庫資訊(FC倉別/交貨日/入庫地址)、產品資訊(品項數量、金額、國際條碼、允收日期；未確認前可下修至0)。"], ["步驟 4：管理訂單-產品資訊", "確認前點「產品資訊編輯」可下修交貨數量(需選無法出貨原因)。數量僅可下修不可上調；未確認可直接改，已確認需另提申請。建議下修幾支即儲存，避免系統異常無法儲存。"], ["步驟 5：管理訂單-狀態確認", "評估各線能否於收單時效內回覆，可先收單後調整。點「供應商確認」跳出檢查項目 → 第4點最終確認可一次性確認。逾期未確認系統自動取消，無法恢復。"], ["步驟 6：管理訂單-退貨資訊", "確認前點「退貨資訊編輯」修改後儲存。務必於確認前填妥退貨地址，否則額外費用由廠商吸收。商品實際效期需大於允收效期(交貨日+總效期1/2)，否則酷澎倉庫端會驗退。退貨方式：酷澎指定貨運(黑貓到付)、廠商指派物流、廠商親取(附自取單)。"], ["步驟 7：進倉規範/倉庫位置/嘜頭更新", "於首頁下載；後台公告更新需通知相關單位(各線、物流、採購、竹運)。後台公告不定期更新需留意。"]], "exceptions": ["有任何問題請於 Supplier Hub 供應商管理系統右上角【線上諮詢】依類別提出詢問。"]}, {"code": "SOP-CP-CP-001-PO調整", "codeShow": "SOP-CP-CP-001", "title": "酷澎 PO 單調整(審核)作業流程", "brand": "跨品類", "platform": "酷澎", "modes": ["買斷"], "delivery": "－", "version": "V1.0", "date": "2025/10/17", "author": "簡孝真", "purpose": "酷澎 PO 單確認後之申請調整/審核(取消)作業流程，提升一致性與效率。適用於 PO 訂單確認(收單)後的調整(取消)審核作業。", "overview": "PO確認 → 依需求調整PO(日期/數量/倉別) → PO資料彙整 → 管理供應狀態維護 → 檢視審核是否通過 → 通過則拉PO單及驗收單 → 未通過通知同仁再送審(反覆多遍) → 無法出貨須整張下修不出 → 直到交貨日前一天", "roles": [["業務/專員", "提出調整訂單明細與需求，或因故確認後無法出貨PO"], ["企劃(OP)", "PO確認後依回覆執行後台操作，申請調整明細或取消不出"]], "steps": [["步驟 1：調整倉別", "相關同仁MAIL提出轉倉需求 → 【線上諮詢】→【供應商｜訂單相關】→【進貨倉別調整諮詢】，提供【欲更改訂單號碼】＋【欲更改倉別】＋【更改原因/同天同進倉訂單】。申請轉倉成功限一次，指定倉別需當天有訂單。"], ["步驟 2：下修數量或整張不出", "【物流】→【商品與供應管理】→【+提出申請】→【更改PO單確認數量】。整張無法出貨：數量下修為1並於評論區備註整單不到貨，另【線上諮詢】→【供應商｜訂單相關】→【其他】提供PO ID告知已下修至1整張不到貨。下修無法上調；已確認訂單無法下修至0；下修最晚可D+1申請(到貨日當天或之後鎖定)。"], ["步驟 3：調整交貨日期", "【物流】→【商品與供應管理】→【+提出申請】→【更改PO單交貨日期】。送貨至少前二日完成修改；延遲交貨可申請D+4內；一天審核2次(約11:00/16:00)。"], ["步驟 4：提出拆/併單需求", "(2025/10/03線上諮詢停止支援，改由酷澎採購協助)拆單：【供應商｜訂單相關】→【其他】附拆單申請表；併單：【供應商｜訂單相關】→【其他】。預留1-2工作天；一般/EOP/D_bucket/NS/New selection/NS2PO無法互相合併；同天同倉同類型才可拆併；併單留意原單是否取消。"], ["步驟 5：管理供應狀態維護", "【物流】→【商品與供應管理】→【+提出申請】→【申請短期無法供貨】或【申請商品狀態為停止供貨】。此狀態即使審核通過，酷澎仍可隨時恢復下單(不須經我司同意)。"], ["步驟 6：確認是否審核完成", "【物流】→【商品與供應管理】→【請求列表】檢視所有審核是否通過；未通過需再通知相關同仁。"], ["步驟 7：拉PO單及驗收單", "所有需求審核調整完成後：【物流】→【PO List】→勾選【PO ID】→下載右上角【PO+驗收單】→回覆各線同仁已完成調整。"]], "exceptions": ["有任何問題請於 Supplier Hub 右上角【線上諮詢】依類別提出詢問。"]}, {"code": "SOP-CP-CP-001-對帳請款", "codeShow": "SOP-CP-CP-001", "title": "酷澎 對帳及上傳發票請款作業流程", "brand": "跨品類", "platform": "酷澎", "modes": ["買斷"], "delivery": "－", "version": "V1.0", "date": "2025/10/07", "author": "簡孝真", "purpose": "建立標準化的酷澎對帳及上傳發票作業流程，提升一致性與效率。適用於酷澎與公司財務OP帳務核對、上傳請款發票/折讓單及書信往來確認。", "overview": "酷澎每月7日對帳(當月2-3工作日後台即產生對帳清單) → 拉對帳單給財務OP核對(區間:當月1-31日) → 財務OP開立發票(電子/手開三聯式)/折讓單(退貨) → 每月14號前上傳 → 審核通過酷澎月底完成付款 → 三聯式手開發票正本寄酷澎會計部、折讓單正本蓋章後寄回", "roles": [["財務OP", "依酷澎對帳資料開立請款發票和折讓單"], ["企劃(OP)", "結帳時間內提供財務OP結帳資料，並於時效內上傳請款發票/折讓單"]], "steps": [["步驟 1：下載對帳資料", "結算 → 發票上傳及應收貨款狀態 → 選帳期月份 → 未附加帳單 → 查詢 → 入庫明細下載(頁面最右)。第二列為負向金額(退貨或倉庫誤驗/驗退)，為財務OP開立折讓單依據(物流→退貨清單)。"], ["步驟 2：存檔至公槽並插入總金額欄位", "建立當月請款資料夾 → 打開template_71813 → 插入I欄總金額(E*H)及J欄備註 → 更改檔名存檔。路徑：轉單平台\\酷澎-買斷\\對帳資料-每月7號(14號前提供發票)。"], ["步驟 3：將對帳資料提供給財務OP核對", "對外：可先MAIL通知酷澎窗口(蘇小姐)已下載當月對帳。對內：除財務OP外CC各線同仁(CPG紙潔／P&G寶僑／Mars瑪氏)；主旨「酷澎20XX年XX月份對帳資料」；內容:對帳月份/帳單號碼/總金額(含稅)/對帳附件。"], ["步驟 4：上傳發票和折讓單", "財務OP核對完畢開立電子或手開三聯式發票(退貨開折讓單) → 點「單據上傳」拖曳檔案上傳。發票日期和發票號碼為必填(多張發票填一個代表即可)。"], ["步驟 5：請款資料送出並確認", "後台已回填發票資訊並上傳全部單據後，每列帳單號碼的請款送出都要「點選確認」。"], ["步驟 6：通知酷澎會計部已完成請款", "MAIL酷澎會計部附上傳單據 → 三聯式手開發票正本寄酷澎會計部、折讓單正本蓋章後寄回。上傳完成後2-3工作天至後台結算系統查詢狀態直到「已完成確認」。"]], "exceptions": ["酷澎會計部：發票抬頭 酷澎股份有限公司；統編 91002999；電話 (02)7751-5656 聯絡人蘇小姐#2704；Email chsu5@coupang.com。", "發票寄送：台北市松山區民生東路三段156號4樓(宏泰金融大樓)，收件人 會計部，accounting_cptw@coupang.com(主旨註明廠商編號A1580000HD)。"]}, {"code": "SOP-CP-CP-001-驗收", "codeShow": "SOP-CP-CP-001", "title": "酷澎 驗收確認作業流程", "brand": "跨品類", "platform": "酷澎", "modes": ["買斷"], "delivery": "－", "version": "V1.0", "date": "2025/10/02", "author": "周渝珊、周亭瑜", "purpose": "建立酷澎驗收確認流程，確保驗收、異常處理與退貨申訴作業正確性，降低風險與錯誤率，並提升作業效率。涵蓋到貨驗收、異常登記、退貨申訴及追蹤。", "overview": "進貨驗收 → 退貨處理 → 驗收異常申請 → 雲端表單登記 → 退貨入庫(原廠#EM / 竹運#CUP)", "roles": [["酷澎專員", "登記雲端異常、回填驗收追蹤總表、提出異常申訴、通知財務OP及物流退貨資訊、訂單系統壓退"], ["企劃(OP)", "協助退貨做訂單交易更正"], ["業務(KA)", "協調特殊狀況(補單、拒收單等)"], ["財務OP(Emma)", "依驗收結果開立發票、原廠出貨開立退單"]], "steps": [["步驟 1：進貨驗收與退貨通知", "若破損或不符進倉標準，酷澎發信通知驗退，廠商須於隔日(工作日)14:00前確認並回覆處理方式(第三方物流寄回／廠商自取／委託其他物流)。需在時限內回覆避免影響進貨。"], ["步驟 2：酷澎後台申請退貨", "後台 → 物流 → 商品與供應管理 → 提出申請 → 退貨商品相關問題 → 填寫申請。申請類型：說明(對退貨有疑慮)／退貨資訊(確認並填退貨日期，限通知信+1周內)／退貨商品有誤需申訴(數量錯誤或損壞需錄影佐證)。"], ["步驟 3：驗收異常申請", "驗收時發現短缺或溢多 → 後台 → 物流 → 商品與供應管理 → 提出申請 → 驗收異常申訴 → 選擇訂單與商品 → 填寫申請內容。"], ["步驟 4：雲端驗收完成表格登記", "訂單商品全數驗入後於Google表單填寫驗收完成資訊。有疑慮的PO及品項需以紅字登記於表格下方。"], ["步驟 5：退貨入庫(原廠出貨#EM)", "有疑慮及驗退的PO/品項於訂單系統備註追蹤。酷澎退貨至竹運後將退貨資訊提供Tobey，由Emma協助開退單。退貨入庫單需填退倉資料，採購協助回填入庫單號再提供Tobey。"], ["步驟 6：退貨入庫(竹運出貨#CUP)", "有疑慮及驗退的PO/品項於訂單系統備註追蹤；訂單系統壓退貨並通知物流不需收退貨。"]], "exceptions": ["未於時限內回覆酷澎退貨通知可能導致延誤，需通知業務及酷澎採購協助。", "退貨商品數量或品項有誤，需錄影佐證並申訴。"]}, {"code": "SOP-CP-CP-001-PDF簽名", "codeShow": "SOP-CP-CPG-001", "title": "酷澎 PDF 驗收單批次簽名自動化流程", "brand": "跨品類", "platform": "酷澎", "modes": [], "delivery": "－", "version": "V1.1", "date": "2026/06/02", "author": "周渝珊", "status": "自動化工具", "purpose": "導入 Google Colab + Python 自動化工具，將驗收單「出貨確認（廠商簽名）」欄位簽名流程標準化與自動化，改善原本人工逐份貼簽名的方式，提升效率並確保一致性。", "overview": "Colab設定(僅第一次) → 準備驗收單及簽名圖檔 → 上傳Google Colab → 執行Python自動簽名程式 → 系統搜尋「出貨確認（廠商簽名）」欄位 → 自動插入簽名圖片 → 輸出已簽名PDF → 自動打包ZIP下載 → 存檔歸檔", "roles": [["企劃(OP)", "上傳PDF與執行程式、自動簽名與輸出PDF及結果確認"]], "steps": [["初始設定 Colab", "1. 開啟 https://colab.research.google.com（使用Google帳號）。2. 點「新增筆記本」建立新專案。3. 貼上自動簽名完整程式碼(!pip install pymupdf；使用 fitz 開啟PDF，search_for(\"出貨確認（廠商簽名）\")定位並 insert_image 插入簽名，逐份輸出 signed_*.pdf 後打包 signed_result.zip 自動下載)。"], ["步驟 1：準備檔案", "第一次設定完後可先開啟Colab按執行，下方會出現選檔區。準備1個 .png 簽名檔及要出貨的驗收單PDF。簽名檔命名不限(程式自動辨識唯一PNG)；避免同時上傳多個簽名檔；檔名避免重複版本混淆。"], ["步驟 2：上傳Colab及執行程式", "將PDF與簽名檔上傳Google Colab並執行。系統僅允許一個簽名檔；可上傳多份PDF。"], ["步驟 3：輸出結果", "執行完成後自動下載 ZIP，內含所有已簽名PDF。"], ["微調簽名位置/大小", "只需改四個數字：sig_width=65、sig_height=22、x0、y1+2（例如縮小改 sig_width=50、sig_height=18），不用改其他程式。"]], "exceptions": ["找不到簽名欄位 → 確認PDF是否為掃描檔、文字是否為「出貨確認（廠商簽名）」。", "簽名未出現 → 確認正確上傳PNG、僅有一個簽名檔。", "輸出檔案異常 → 重新開啟Colab並重新執行。"]}, {"code": "SOP-CP-CP-001-約倉", "codeShow": "SOP-CP-CP-001", "title": "酷澎 約倉作業流程", "brand": "跨品類", "platform": "酷澎", "modes": ["買斷"], "delivery": "－", "version": "V1.2", "date": "2026/06/18", "author": "周亭瑜、周渝珊", "purpose": "建立酷澎約倉流程，降低司機等待時間。涵蓋出貨前約倉作業與現場調整安排。", "overview": "推算車數與時長 → 酷澎系統約倉 → 出貨/現場排隊", "roles": [["酷澎專員", "依出貨量推算車數與下貨時長，於酷澎約倉系統完成預約"], ["送貨司機", "交貨日依倉庫現場指示進倉"], ["KA/PM", "協助與酷澎採購協調異常處理"]], "steps": [["步驟 1：推算車數與時長", "紙潔：每車300–400箱、卸貨2小時；建議每時段先約2台車以因應同單跨倉。瑪氏：每車100-150箱(體積小)，多可估200-300箱(大台車)，7板約150箱=21箱/板，半小時約翻2版。寶僑：向原廠詢問板數車數依其資訊約倉，一時段=1台車；竹運出貨每車150箱、25箱/板。"], ["步驟 2：酷澎系統約倉", "紙潔D-2依原廠需求調整並回覆預約時段。步驟：物流 → 運輸 → 創建預約出貨 → 選擇交貨日 → 勾選訂單 → 新增預約時間 → 選擇棧板數量及預約時間 → 預訂。須提前D-5完成以免無可用時段；每時間以半小時為單位。"]], "exceptions": ["系統無可約倉時段 → 現場排隊。", "到倉逾時 → 現場排隊。", "司機回報等候過久且進不了倉 → 回報KA/PM。", "CPG於物流群回報 → 請酷澎採購協助。", "現場卸貨時間超時 → 回報KA/PM。"]}, {"code": "SOP-CP-CP-001-全面前端", "codeShow": "SOP-CP-CP-001", "title": "酷澎 全面前端作業流程", "brand": "跨品類", "platform": "酷澎", "modes": ["買斷"], "delivery": "－", "version": "V1.0", "date": "2026/06/01", "author": "簡孝真、周亭瑜、周渝珊", "status": "未完成草稿", "purpose": "規範酷澎前端(對外部)PO 確認與調整之作業流程，明確各步驟作業內容、負責人員、時效與資料夾歸檔位置，提升一致性並降低處理無效PO的時間。作業天數以交貨日D往前推D-5(不含假日)，調整不過再+2天(收單到出貨最少7個工作天)。", "overview": "時效速記：交貨日D → 往前推D-5(不含假日) → 調整不過再+2天 → 收單到出貨最少7個工作天。分三階段：① 收單(調整)確認 ② 已排單(調整)確認 ③ 出貨(調整/安排)確認", "roles": [["企劃(OP)", "拉單/收單/取消、供應狀態維護、退貨資訊修改、更改PO確認數量、驗收單落款、統整發信提醒、記錄下修SKU(P&G)"], ["酷澎專員", "各線配貨確認/主檔維護、改倉改EDD、約倉、記錄下修SKU"], ["瑪氏業務", "瑪氏線配貨確認"], ["P&G業務", "P&G線配貨確認"]], "steps": [["第一階段·步驟1：拉單/收單/取消(小真-全線)", "交貨日D-5、酷澎介面有新PO待下載 → 下載PO(每天2次早/午)拋入公司採購入庫處理系統 → 發MAIL通知分線別拉取。需覆核有無漏抓上拋的PO。歸檔：0.0 PO原始檔。"], ["第一階段·步驟2：配貨確認/主檔維護(紙潔-Chloe/瑪氏-雅慈/P&G-Nicole)", "確認交貨日期、倉別、商品狀態、貨量及可出貨品項 → 各分線下載彙總檔 → 確認(檔名:PO開單日+開立-線別)放公槽並回覆小真。一併勾選空白(無線別)判斷是否有新增SKU需維護主檔。歸檔：1.確認PO／0.主檔維護。"], ["第一階段·步驟3：維護管理供應狀態(小真-全線)", "J欄永豐料號標示「停產/不出/暫不供貨」直接下修0。暫不供貨/不出 → 下修後申請暫停供貨(抓28-30天)；停產 → 下修後申請停止供貨。歸檔：申請審核。"], ["第一階段·步驟4：退貨資訊修改(小真-全線)", "確定要收單(預計要出貨)的再修改退貨資訊。"], ["第一階段·步驟5：更改交貨日期或物流中心(紙潔-Chloe/瑪氏+P&G-Nicole)", "收單過程發現需改倉或改交期 → 申請改倉、改EDD。有異動主動告知。"], ["第二階段·步驟6：更改PO確認數量(Chloe/Nicole/小真-全線)", "預計可出貨下修，或出貨前缺貨(不出)調整。各線自行拉PO檢視是否再調整(檔名:調整日+調整-線別)放公槽回覆小真。酷澎後台下修及驗收單落款由小真統一作業。※目前僅紙潔於發單前做庫存確認。歸檔：2.已收單-異動PO／3.已排單下修／驗收單。"], ["第二階段·步驟7：約倉/取消(紙潔-Chloe/瑪氏+P&G-Nicole)", "確認已排定即可先約；PO改期改倉需取消。登記雲表(酷澎紙潔約倉表／酷澎約倉表-瑪氏&PG)。歸檔：4.已約倉。"], ["第三階段·步驟8：出貨前下修(基本上為缺貨)", "出貨前核單發現缺貨 → 自行拉PO調整(檔名:PO#_交貨日_倉別-線別-出貨前下修)放公槽回覆小真。後端有效變無效記錄下修SKU：小真(P&G)/Chloe(紙潔)/Nicole(瑪氏)。歸檔：5.出貨前下修。"]], "exceptions": ["調整PO不管自行申請或請採購協助，有異動都要主動告知以利追蹤。正常提出申請約1天調整好，超過視為有問題。", "PO調整不過 → 由小真統整發信提醒(時效以收到確認回覆郵件後D+2計算)，盡量不卡單快轉備案。", "重複無效SKU → 彙總檔J欄標示即直接下修0；暫不供貨/不出先下修並申請暫停供貨(抓28-30天)。", "進倉問題/物流問題 → 進倉當日物流反應相關問題記錄(等候過久、拒收、嘜頭效期等)。"]}, {"code": "SOP-CP-CPG-001-出貨", "codeShow": "SOP-CP-CPG-001", "title": "酷澎 CPG（紙潔）出貨作業流程", "brand": "CPG", "platform": "酷澎", "modes": ["買斷"], "delivery": "－", "version": "V1.0", "date": "2025/10/01", "author": "周渝珊", "status": "待更新", "purpose": "建立標準化的酷澎訂單整理與修改流程，確保作業時間節點正確，並提升跨部門協作效率。涵蓋新單整理、訂單修改、出貨排程。", "overview": "酷澎新單整理 → 確認貨況 → 回報訂單修改需求 → 更改訂單資訊 → 確認PO進度、重複拉PO及驗收單歸檔 → 整合訂單 → 排定出貨 → 約倉 → 出貨文件製作 → 採購作業(原廠) → 送貨", "roles": [["收單/調整窗口(小真)", "接收並提供新訂單資訊、系統修改PO單"], ["業務、PM", "確認新單內容、與原廠(CPG)及酷澎採購協調"], ["酷澎專員", "協助新單整理(Google表單、整合表)、整合訂單需求、當場拒收登記"], ["企劃(OP)", "製作出貨文件(EIP下採、嘜頭、驗收單)及系統拋檔、出貨前缺貨下修"]], "steps": [["步驟 1：接收與整理新單", "接收新單整理，專員協助業務確認並把訂單放到大整合表。整合表格式請至拆單工具複製。存檔:紙品+潔品\\6.酷澎\\訂單。"], ["步驟 2：訂單修改需求", "收到小真分類完訂單後整單，加入料號、轉換率計算箱數箱價，提供業務/PM確認貨況。照原始訂單複製貼上已設公式→帶到拆解訂單Sheet→貼到大整合表-紙潔All。"], ["步驟 3：依貨況回覆修改需求、整合訂單", "依PM當月供貨量(FCST)排定出貨，參考產銷預估表、紙潔配貨表(缺貨品項)排出貨日，依規則調整時間/倉別，MAIL請小真後台申請。潔品配貨表為最小單位需換算箱；用Google表單(拆單_酷澎數量統計)回信小真。"], ["步驟 3-1：拆單需求處理", "整理整合表需求通知小真或酷澎採購修改。製作拆單檔→存檔→MAIL請小真協助後台申請→完成後小真拉單歸檔。"], ["步驟 3-2：併單需求處理", "確認訂單類型是否能併(一般/EOP/D_bucket/NS/New selection/NS2PO無法互相合併)，再MAIL小真協助後台申請→完成後小真拉單歸檔。"], ["步驟 4：排定出貨", "用範例格式補上訂單資訊，整理最後可出貨訂單，填寫線上表格出貨清單，通知企劃出貨。D-5(工作天)完成。"], ["步驟 5：出貨前下修", "採購核單完發缺貨通知，企劃(OP)MAIL請小真下修並改雲表(酷澎)數量及金額，下修完拉新驗收單給採購。需附要改數量的訂單。"], ["步驟 6：拒收", "收到採購通知拒收，酷澎專員修改雲表(酷澎)、大整合表的箱數及金額。"]], "exceptions": ["業績追蹤：將酷澎已下單、已排到倉訂單更新於樞紐sheet供PM追蹤(業績為未稅，已出貨+已排定)。"]}, {"code": "SOP-CP-MARS-001-新單整理", "codeShow": "SOP-CP-MARS-001", "title": "酷澎 MARS 新單整理與 PO 單修改流程", "brand": "MARS", "platform": "酷澎", "modes": ["買斷"], "delivery": "－", "version": "V1.0", "date": "2025/09/30", "author": "周亭瑜", "purpose": "規範酷澎新單整理與訂單修改需求的作業流程，確保訂單處理的正確性與一致性，提升溝通與作業效率。", "overview": "接收新單(小真提供) → 整理新單(專員填Google表單、存PO製作整合表) → 業務確認並回覆整合表 → 如有修改需求業務回填(改交期/改倉/可出貨數量) → 專員依回覆更新訂單並標註修改部分", "roles": [["小真", "接收並提供新訂單資訊、系統修改PO單"], ["業務", "回覆整合表，確認並提出訂單修改需求"], ["酷澎專員", "新單整理(Google表單、整合表)，整合業務訂單需求"]], "steps": [["步驟 1：新單整理", "1. 新單資訊貼於Google表單。2. PO單存於「瑪氏_未確認訂單」建資料夾(依收信日命名)。3. PO內容貼於EXCEL(整合表格式-Mars 3.0整合表)。4. 通知業務可於資料夾確認訂單。表單需填寫完整避免遺漏SKU或數量。"], ["步驟 1-1：製作整合表", "1. PO內容貼至整合表格式(分頁-原始訂單)。2. 到分頁-拆解訂單，將H欄條碼複製貼到記事本再貼回H欄，後面箱入數公式會帶入。"], ["步驟 2：訂單修改需求處理", "1. 修改需求填於整合表：B欄改交期、D欄改倉別、T欄業務回覆可出貨數量。2. 對照N欄出貨數量與T欄回覆數量，有修改標紅字。3. 專員依修改更新PO並回覆。PO品項出貨單位分箱及小單位需分開；超過酷澎限制需通知業務；延遲交期可申請D+4內(含假日)；轉倉成功限一次。"], ["步驟 2-1：拆單需求處理", "同一PO商品出貨單位分箱與小單位(包/盒)需拆單。拆單EXCEL(Coupang-拆單TEMPLATE)依整合表填：拆單編號(原PO ID不同但交貨日倉別同可填同)、預計交貨日、倉別、備註(整合表AC欄)。"], ["步驟 3：訂單修改需求回覆", "1. PO修改資訊填至Google表單(原單改交期/改倉填想改期/想改倉庫；只拆單原單不變填備註即可；只改交期倉別不改數量則不用改PO)。2. 整合表需求整理好通知小真修改。"]], "exceptions": ["表單無法存取 → 通知系統/表單維護者，暫以Excel手動彙整。", "訂單修改需求不清楚或衝突 → 立即聯繫業務確認。", "PO單缺漏 → 向小真或業務確認。"]}, {"code": "SOP-CP-MARS-001-出貨", "codeShow": "SOP-CP-MARS-001", "title": "酷澎 MARS 出貨作業流程", "brand": "MARS", "platform": "酷澎", "modes": ["買斷"], "delivery": "－", "version": "V1.1", "date": "2026/06/10", "author": "周亭瑜", "purpose": "規範出貨作業流程，確保商品正確、及時出貨，提升作業一致性與效率。涵蓋酷澎平台之整合訂單、出貨排程。", "overview": "確認PO進度、重複拉PO及驗收單歸檔 → 整合訂單 → 排定出貨 → 約倉 → 出貨文件製作 → 採購作業(原廠) → 送貨", "roles": [["酷澎專員", "整合訂單、出貨訂單約倉作業"], ["業務(KA)", "協助異常處理、製作出貨採購單"], ["企劃(OP)", "製作出貨文件(EIP下採、嘜頭、驗收單)及系統拋檔、出貨前缺貨下修"]], "steps": [["步驟 1：訂單歸檔", "訂單修改完成後，小真將文件附於信件，將PO及驗收單歸檔(照月份新增資料夾，依交貨日命名)。"], ["步驟 2：併單需求處理", "整理可合併訂單通知小真處理。同類型才可合併(後台可查訂單類型)：系統單Tip_system／採購指定單mid_march_gap／採購預下單需100%到貨(專案)Eop_Tip_system／新品單NS(不能與其他合併)。一般/EOP/D_bucket/NS/New selection無法互相合併。"], ["步驟 3：約倉作業", "1. 可出貨訂單貼整合表統計P欄箱數(PO有問題回報業務)。2. 出貨資訊貼Google表單估箱/板/車數(車數一般100-150箱，多可200-300箱大台車；7板約150箱=21箱/板；半小時翻2版)。3. 約倉：物流→運輸→創建預約出貨→選交貨日→選訂單→新增預約時間→選棧板數量及時間→預訂。4. 完成後將預約時間貼Word存出貨資料夾並通知業務製作採購單。交貨日D-4需拋檔給瑪氏；確認品項是否同單位(整合表Q欄)；金額是否超過5千。"], ["步驟 4：竹運出貨相關文件", "竹運出貨訂單需做EXCEL(永豐料號對照表_交貨-竹運出貨專用)，完成後通知OP製作出貨文件，不需製作採購單。"], ["步驟 5：出貨前下修", "採購核單完發缺貨通知，企劃(OP)MAIL請小真下修。需附下修數量的PO；下修前需先取消該筆PO約倉，下修完成後務必約回。"]], "exceptions": ["系統無可約倉時段 → 現場排隊。", "訂單資料有誤 → 通知業務確認。"]}, {"code": "SOP-CP-PG-001-出貨", "codeShow": "SOP-CP-PG-001", "title": "酷澎 PG 出貨作業流程", "brand": "PG", "platform": "酷澎", "modes": ["買斷"], "delivery": "－", "version": "V1.0", "date": "2026/06/10", "author": "周亭瑜", "purpose": "建立標準化的酷澎訂單整理與修改流程，確保作業時間節點正確，提升作業一致性與處理效率。涵蓋酷澎平台之整合訂單、出貨排程。", "overview": "酷澎新單整理 → 確認貨況 → 回報訂單修改需求 → 更改訂單資訊 → 確認PO進度 → 整合訂單 → 排定出貨 → 約倉 → 出貨文件製作 → 採購作業(原廠) → 送貨", "roles": [["收單/調整窗口(小真)", "接收並提供新訂單資訊、系統修改PO單"], ["業務、PM", "確認新單內容、與原廠(PG)及酷澎採購協調"], ["酷澎專員", "新單整理、整合訂單需求、當場拒收登記"], ["企劃(OP)", "製作出貨文件(EIP下採、嘜頭、驗收單)及系統拋檔、出貨前缺貨下修"]], "steps": [["步驟 1：接收與整理新單", "依開單日建資料夾存訂單，PO依商品類型放入對應Excel(PG專案報價/幫寶適專案報價)。若混商品依多的放進Excel。存檔:P&G寶僑\\價格本\\酷澎報價。"], ["步驟 1-1：放單", "PO依品項貼至對應Excel(Coupang PG專案報價=除幫寶適外品項；Coupang 幫寶適專案報價=幫寶適)。貼上SKU ID、國條、品名、下單數量、出貨數量、酷澎下單價(含稅)，其他欄位公式帶出。R欄建立交貨資訊(PO單號_日期交貨(倉)，不同PO黃底區隔)。K欄單價vs L欄酷澎下單價需核對(酷澎可能自行改價)。"], ["步驟 2：訂單修改需求", "依PG可供貨量排定出貨，參考總表「PG剩餘最大可供貨量(CS)」，依規則調整交期/倉別/數量，PO調整內容MAIL小真。供貨量皆箱單位；同交期倉別出貨最少400箱、同天不超過兩倉別；幫寶適箱單位只能進TXRC8；可進倉別TAO1/TAO3/TAO4/TAO5/TXFC1/TXRC8/RXRC17。"], ["步驟 2-1：PO單需求處理", "數量下修通知小真；改交期倉別系統申請或酷澎採購協助。系統修改EDD/倉別：物流→商品與供應管理→提出申請→更改PO單交貨日期或物流中心→選PO→選交期倉別、申請原因→提出。系統只顯示有空位倉別；提出申請前先取消已有預約出貨。"], ["步驟 3：排定出貨", "整理可出貨訂單通知企劃(OP)做檔，製作Excel統計出貨總箱數，MAIL PG提供預估板數&車數約倉。提供PG附檔:留A~F欄、J欄及R欄，插空白B欄「PG CODE」用總表V資料，分頁命名交貨日_交貨(倉)。D-4(工作天)完成；提供PG的附檔金額一定要刪掉。"], ["步驟 4：出貨前下修", "採購核單完發缺貨通知，企劃(OP)MAIL請小真下修、專員先取消該筆PO約倉，並修改放單Excel出貨數量及總表出貨數量。下修完成後務必約回。"]], "exceptions": ["訂單修改需求不清楚或衝突 → 立即聯繫業務確認。", "PO單缺漏 → 向小真或業務確認。", "系統無可約倉時段 → 現場排隊。"]}, {"code": "SOP-CP-CP-001-QA", "codeShow": "SOP-CP-CP-001", "title": "酷澎 QA 資料庫（知識庫問答）", "brand": "跨品類", "platform": "酷澎", "modes": [], "delivery": "－", "version": "V1.0", "date": "2026/07/23", "author": "簡孝真／周渝珊／周亭瑜", "status": "知識庫", "purpose": "建立與酷澎作業流程相關疑難雜症的知識庫問答，針對操作酷澎使用者問題快速提供精確答案。", "overview": "四大類常見問題：① PO單確認相關 ② 進倉驗收問題 ③ 退貨商品相關問題 ④ 帳務/請款相關", "roles": [["專員/企劃(OP)", "紀錄酷澎會面臨的問題和處理方式"]], "steps": [["1. PO 單確認相關", "• 收單時效：一般48小時、NS 96小時，逾時自動取消。• 確認PO需留意交貨日、倉別、金額、商品狀態、貨量、名稱、條碼、數量、價格、總金額、效期。• SKU 2周內無法供貨需申請暫缺/停產。• 轉倉+改EDD只能依系統倉別/交貨日選擇(無選項=滿倉或無法轉)。• 更改數量/交貨日需先取消約倉(目前非強制)。• 改交貨日/物流中心=預計交貨日3天前(日曆天)提交。• 改確認數量=交貨日2天前提交。• 上修須≤原數量；下修為0視為取消該商品且無法再異動。• 已確認訂單欲取消整張，最少需剩1品項數量1並線上諮詢提供POID。• 所有申請都要檢查是否通過，出貨前確認PO狀態與明細吻合(酷澎會自行異動或取消)。• 拆/併單需洽酷澎採購，作業約D+2。"], ["2. 進倉驗收問題", "• 驗收異常申訴=商品到倉日+7天後仍未驗收完畢；超過30天的申請將被拒絕。• 申請時一次提供完整佐證出貨資料(已包裝出貨照片/影片)。• 佐證不明確酷澎會要求補件；倉庫超過15天未收到補齊回覆則申訴自動結案。• 佐證資料請於「附件」提供，否則無法協助查詢驗收差異。"], ["3. 退貨商品相關問題", "• 提出申請須於收到退貨信件36小時內(不含六日)。• 不符進倉規範商品會走退貨流程。• 說明=對退貨信件有疑慮選此。• 退貨資訊=認同退貨選此填資訊(限收到信件+1周內日期)。• 退貨商品有誤需申訴=數量錯誤/退錯/損壞，須自開箱起錄影否則無法申訴。"], ["4. 帳務/請款相關", "• 對帳區間：進貨當月1–31日(酷澎會計部 accounting_cptw@coupang.com)。• 每月7日對帳(當月2-3工作日後台即產生清單)。• 每月14號前上傳發票(手開3聯式正本蓋章寄出)/折讓單；逾期遞延次月審核不另通知。• 開立與我司系統相符之發票金額(稅額自行回推5%)。• 撥款固定每月30號(遇假日提前)。• 酷澎統編91002999，電話(02)7751-5656，廠商編號A158000…。"]], "exceptions": []}, {"code": "酷澎採購入庫操作手冊", "codeShow": "系統操作手冊", "title": "酷澎採購入庫處理系統 — 操作手冊", "brand": "跨品類", "platform": "酷澎", "modes": [], "delivery": "－", "version": "—", "date": "—", "author": "—", "status": "系統手冊", "purpose": "本系統協助處理 Coupang 訂單入庫與 MARS 效期自動匹配。作業分三階段：資料準備(匯入PO、維護主檔、瑪氏匯入PDF效期表)→ 資料加工(整合表自動匹配效期、換算箱/板數)→ 文檔印製(匯出符合酷澎規範的嘜頭與整合表)。", "overview": "功能一 酷澎採購入庫處理（Tab1批次上傳採購單／Tab2主檔匯入／Tab3整合表預覽匯出／Tab4產生嘜頭） ＋ 功能二 瑪氏商品配送處理（Tab1批次上傳配送表PDF／Tab2配送資料查詢／Tab3竹運嘜頭與效期合併-尚未開發）", "roles": [["企劃(OP)", "操作系統匯入PO/主檔/效期，匯出整合表與嘜頭"]], "steps": [["功能一·Tab1：批次上傳採購單", "點「選擇檔案」選PO單 → 「批次匯入Excel」存入資料庫。須登入(否則使用者ID遺失匯入失敗)；建議一次≤50份PO；重新匯入會刪除原PO再新增。"], ["功能一·Tab2：酷澎主檔匯入", "下載範本Excel → 填SKU ID、條碼、線別、品牌、品名、箱入數、單價、效期天數 → 上傳(不存在新增、存在更新Upsert)。箱入數(ConvertRate)為正整數決定板嘜/箱嘜(=1板嘜、>1箱嘜)；線別為MARS會啟動瑪氏效期匹配；啟用為否(N)會匯入但整合表不呈現。改一支SKU上傳一支即可。"], ["功能一·Tab3：整合表預覽匯出", "設定查詢(線別預設全選、上傳人模糊查詢、PO/SKU/倉別、日期範圍)→查詢→預覽整合表→勾選欲匯出列→匯出Excel。三種匯出：拆單格式(整理線別彙整多張PO配貨用)、嘜頭格式(符合10列限制、板貨一板最多5SKU)、訂單匯入格式(供匯入系統，E~I欄上鎖僅Emma可改，密碼1234)。同SKU多效期請先下載整合表手動改序號欄避免重複。"], ["功能一·Tab4：產生嘜頭", "方式1整合表直接產製：Tab3勾選→「產生嘜頭」系統自動判板/箱嘜→批次下載Zip。方式2手動匯入Excel：Tab3下載嘜頭格式Excel編輯→「匯入Excel」提交產製→批次下載Zip。分組依PO單號+到貨日+倉別；箱嘜(箱入數≠1或瑪氏)逐SKU一張、板嘜(箱入數=1且非瑪氏)每5SKU一張；嘜頭Excel不可超過10列；檔名CoupangLabel_{PO}_{到貨日}_{倉別}_{pal/box}_{序號}。"], ["功能一·MARS 嘜頭特殊規則", "箱嘜=出貨數量÷箱入數後無條件進位(Ceiling)換算總箱數(例:50件÷12=5箱)；板嘜=直接用原始出貨數量不進位。效期自動比對MARS SKU與效期表YFItemCode，僅帶UsedFlag=N未使用效期；產製後不會自動標記已使用，需至瑪氏配送處理手動「批次更新狀態」。同SKU多效期建議先下載整合表改序號再匯入。"], ["功能二·Tab1：批次上傳配送表PDF", "選進貨倉別→選檔案(可多份原廠導出PDF)→「批次匯入PDF」解析永豐料號(YFItemCode)與指送日期(ExpiryDate)存入效期表。僅支援原始格式PDF(勇信配送表)，重印或加密版解析失敗；上傳後初始為未使用(UsedFlag=N)。"], ["功能二·Tab2：配送資料查詢", "查詢條件(上傳人模糊查詢、時間範圍、倉別)→查詢→勾選紀錄→產完嘜頭後回此頁「批次修改狀態」更新為已使用(UsedFlag=Y)。若未更新，下次匯出整合表仍會抓到舊效期。"], ["功能二·Tab3：竹運嘜頭與效期合併（尚未開發）", "規劃未來瑪氏功能開放其他倉別：上傳嘜頭PDF(批次)→系統解析標籤(SKU/效期/數量)→與整合表或主檔比對標記已處理→提供合併結果預覽與匯出。"]], "exceptions": ["匯入失敗 → 確認登入狀態與PO檔格式。", "整合表沒有效期 → 確認已上傳MARS效期表且UsedFlag=N。", "嘜頭檔名重複 → 同SKU多效期時發生，先下載整合表改序號再上傳。"]}];
const KW = {"vocab": ["補貨", "退貨", "缺貨下修", "嘜頭", "效期", "EIP", "拋單", "約倉", "驗收", "驗收單", "領用", "採購單", "庫存核對", "入庫單", "進倉單", "盤點", "加工組裝", "轉換率", "二聯單", "退廠", "請款", "對帳", "客服", "大宗訂單", "盤虧", "庫存調整", "自編碼", "借貨單", "入庫憑單", "PO單", "拆併單", "轉倉", "供應狀態", "自動化", "板嘜", "箱嘜"], "alias": {"缺貨下修": ["缺貨", "下修"], "庫存核對": ["核對", "庫存"], "拋單": ["拋單", "拋檔"], "領用": ["領用"], "客服": ["客服", "問答", "聊聊"], "庫存調整": ["庫調", "庫存調整"], "大宗訂單": ["大宗"], "盤虧": ["盤虧"], "驗收": ["驗收"], "PO單": ["PO單", "PO ID", "PO單號"], "拆併單": ["拆單", "併單"], "轉倉": ["轉倉", "改倉"], "供應狀態": ["停止供貨", "暫停供貨", "供貨"], "自動化": ["自動化", "Colab", "Python"]}};

let DATA = [];   // 實際渲染用的資料，由 loadData() 填入
</script>
<script>
const el=(t,c,h)=>{const e=document.createElement(t);if(c)e.className=c;if(h!=null)e.innerHTML=h;return e;};
const esc=s=>String(s).replace(/[&<>"]/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[m]));

const PLATFORM_ORDER=["酷澎","博客來","PCHOME","MOMO","雅虎","蝦皮","東森","UberEats","遠時","HCT新竹物流","通用"];
const state={q:"",brand:new Set(),platform:new Set(),mode:new Set(),delivery:new Set(),kw:new Set()};

// ---- build filter counts ----
function facetValues(key){
  const m=new Map();
  DATA.forEach(d=>{
    let vals = key==='mode'? d.modes : [d[key]];
    if(key==='mode' && (!vals||!vals.length)) return;
    vals.forEach(v=>{ if(v&&v!=='－') m.set(v,(m.get(v)||0)+1); });
  });
  return m;
}
function orderVals(key,map){
  let ks=[...map.keys()];
  if(key==='platform') ks.sort((a,b)=>PLATFORM_ORDER.indexOf(a)-PLATFORM_ORDER.indexOf(b));
  else if(key==='brand') ks.sort((a,b)=>(BRAND_ORDER[a]??9)-(BRAND_ORDER[b]??9));
  else if(key==='mode') ks.sort((a,b)=>({'寄銷':0,'買斷':1,'領用':2}[a])-({'寄銷':0,'買斷':1,'領用':2}[b]));
  else if(key==='delivery') ks.sort((a,b)=>({'直送':0,'竹運':1}[a])-({'直送':0,'竹運':1}[b]));
  return ks;
}
const brandLabel={'MARS':'MARS 瑪氏','PG':'PG 寶僑','CPG':'CPG 紙潔','跨品類':'跨品類','通用':'通用作業'};
const BRAND_ORDER={'MARS':0,'PG':1,'CPG':2,'跨品類':3,'通用':4};
function buildFilters(){
  const body=document.getElementById('filtersBody');
  const groups=[['brand','品牌 / 類別'],['platform','平台 / 通路'],['mode','寄銷 / 買斷 / 領用'],['delivery','物流方式（直送 / 竹運）']];
  groups.forEach(([key,label])=>{
    const g=el('div','fgroup'); g.appendChild(el('h3',null,label));
    const opts=el('div','fopts');
    const map=facetValues(key);
    orderVals(key,map).forEach(v=>{
      const b=el('button','fopt'); b.dataset.k=key; b.dataset.v=v;
      const disp = key==='brand'?(brandLabel[v]||v):v;
      b.innerHTML=`${esc(disp)} <span class="n">${map.get(v)}</span>`;
      b.onclick=()=>{ state[key].has(v)?state[key].delete(v):state[key].add(v); b.classList.toggle('on'); render(); };
      opts.appendChild(b);
    });
    g.appendChild(opts); body.appendChild(g);
  });
}

// ---- keyword suggestions ----
const HOT_KW=["退貨","補貨","嘜頭","效期","約倉","驗收","PO單","對帳","請款","退廠","EIP","拋單","缺貨","自編碼"];
function buildKw(){
  const box=document.getElementById('kwSuggest');
  HOT_KW.forEach(k=>{
    const b=el('button',null,esc(k)); b.dataset.kw=k;
    b.onclick=()=>{ state.kw.has(k)?state.kw.delete(k):state.kw.add(k); b.classList.toggle('on'); render(); };
    box.appendChild(b);
  });
}

// ---- stats ----
function buildStats(){
  const s=document.getElementById('stats');
  const platN=new Set(DATA.filter(d=>d.platform!=='通用').map(d=>d.platform)).size;
  const brandN=new Set(DATA.filter(d=>['MARS','PG','CPG'].includes(d.brand)).map(d=>d.brand)).size;
  [[String(DATA.length),'SOP 文件'],[String(platN),'電商通路'],[String(brandN),'品牌'],['3','出貨模式']].forEach(([n,l])=>{
    s.appendChild(el('div','stat',`<b>${n}</b><span>${l}</span>`));
  });
}

// ---- filtering ----
function matches(d){
  if(state.brand.size && !state.brand.has(d.brand)) return false;
  if(state.platform.size && !state.platform.has(d.platform)) return false;
  if(state.mode.size && !d.modes.some(m=>state.mode.has(m))) return false;
  if(state.delivery.size && !state.delivery.has(d.delivery)) return false;
  for(const k of state.kw){ if(!d.search.includes(k)) return false; }
  if(state.q){ if(!d.search.toLowerCase().includes(state.q.toLowerCase())) return false; }
  return true;
}
function hl(text){
  if(!state.q) return esc(text);
  const q=state.q.trim(); if(!q) return esc(text);
  const idx=text.toLowerCase().indexOf(q.toLowerCase());
  if(idx<0) return esc(text);
  return esc(text.slice(0,idx))+'<mark>'+esc(text.slice(idx,idx+q.length))+'</mark>'+hl(text.slice(idx+q.length));
}

// ---- card render ----
const openCards=new Set();
function card(d){
  const c=el('div','card'+(openCards.has(d.code)?' open':'')); c.dataset.code=d.code;
  const modeChips=d.modes.map(m=>`<span class="chip mode-${m}">${m}</span>`).join('');
  const delivChip=d.delivery&&d.delivery!=='－'?`<span class="chip deliv ${d.delivery}">${d.delivery}</span>`:'';
  const statusChip=d.status?`<span class="chip status">${esc(d.status)}</span>`:'';
  const head=el('div','card-head');
  head.innerHTML=`
    <div class="ch-top">
      <h3>${hl(d.title)}</h3>
      <span class="code">${esc(d.codeShow||d.code)}</span>
    </div>
    <div class="chip-row">
      <span class="chip brand-${d.brand}">${brandLabel[d.brand]||d.brand}</span>
      <span class="chip plat">${esc(d.platform==='通用'?'通用':d.platform)}</span>
      ${modeChips}${delivChip}${statusChip}
    </div>
    <div class="ch-foot">
      <div class="ch-meta"><span>版本 ${esc(d.version)}</span><span>${esc(d.date)}</span><span>${esc(d.author)}</span></div>
      <span class="expand-ind">${openCards.has(d.code)?'收合':'展開細節'}
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2"><path d="m6 9 6 6 6-6"/></svg></span>
    </div>`;
  head.onclick=()=>{ openCards.has(d.code)?openCards.delete(d.code):openCards.add(d.code); render(); };
  c.appendChild(head);

  if(openCards.has(d.code)){
    const body=el('div','card-body');
    // purpose
    body.appendChild(sec('目的',`<p class="purpose">${hl(d.purpose)}</p>`));
    // overview flow
    const flow=esc(d.overview).replace(/→/g,'<b>→</b>');
    body.appendChild(sec('流程總覽',`<div class="flow">${flow}</div>`));
    // roles
    if(d.roles&&d.roles.length){
      const rows=d.roles.map(r=>`<tr><td>${esc(r[0])}</td><td>${esc(r[1])}</td></tr>`).join('');
      body.appendChild(sec('角色與職責',`<table class="roles"><tbody>${rows}</tbody></table>`));
    }
    // steps
    const steps=el('div','steps');
    d.steps.forEach((st,i)=>{
      const sd=el('div','step');
      const bodyTxt=st[1]&&st[1].trim()?st[1]:'（本步驟以系統操作截圖為主，無文字細節）';
      const isEmpty=!(st[1]&&st[1].trim());
      sd.innerHTML=`
        <div class="step-h">
          <span class="step-num">${i+1}</span>
          <span class="step-t">${hl(st[0])}</span>
          <svg class="caret" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2"><path d="m9 6 6 6-6 6"/></svg>
        </div>
        <div class="step-body${isEmpty?' empty':''}">${hl(bodyTxt)}</div>`;
      sd.querySelector('.step-h').onclick=()=>sd.classList.toggle('open');
      steps.appendChild(sd);
    });
    body.appendChild(sec(`Step-by-Step 操作細節（${d.steps.length} 步）`,'',steps));
    // exceptions
    if(d.exceptions&&d.exceptions.length){
      const lis=d.exceptions.map(e=>`<li>${hl(e)}</li>`).join('');
      body.appendChild(sec('異常處理',`<ul class="exc">${lis}</ul>`));
    }
    // tags
    if(d.tags&&d.tags.length){
      const t=d.tags.map(x=>`<span class="tag">#${esc(x)}</span>`).join('');
      body.appendChild(sec('關鍵字',`<div class="tags">${t}</div>`));
    }
    c.appendChild(body);
  }
  return c;
}
function sec(label,html,node){
  const s=el('div','sec'); s.appendChild(el('div','sec-label',esc(label)));
  if(node) s.appendChild(node); else s.insertAdjacentHTML('beforeend',html);
  return s;
}

function render(){
  const grid=document.getElementById('grid'); grid.innerHTML='';
  const list=DATA.filter(matches);
  document.getElementById('resultCount').textContent=list.length;
  // refresh facet dynamic disabled? keep simple. Update option counts contextually:
  if(!list.length){
    const es=el('div','empty-state');
    es.innerHTML=`<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6"><circle cx="11" cy="11" r="7"/><path d="m21 21-4.3-4.3"/></svg><p>找不到符合條件的 SOP 文件</p>`;
    const btn=el('button',null,'清除所有篩選'); btn.onclick=resetAll; es.appendChild(btn);
    grid.appendChild(es); return;
  }
  list.forEach(d=>grid.appendChild(card(d)));
}
function resetAll(){
  state.q='';state.brand.clear();state.platform.clear();state.mode.clear();state.delivery.clear();state.kw.clear();
  document.getElementById('search').value='';
  document.getElementById('clearSearch').style.display='none';
  document.querySelectorAll('.fopt.on,.kw-suggest button.on').forEach(b=>b.classList.remove('on'));
  render();
}

document.getElementById('search').addEventListener('input',e=>{
  state.q=e.target.value;
  document.getElementById('clearSearch').style.display=e.target.value?'block':'none';
  render();
});
document.getElementById('clearSearch').onclick=()=>{
  state.q='';document.getElementById('search').value='';
  document.getElementById('clearSearch').style.display='none';render();
};
document.getElementById('resetBtn').onclick=resetAll;
document.getElementById('mFilter').onclick=()=>{
  const f=document.getElementById('filters'); f.classList.toggle('collapsed');
  document.querySelector('#mFilter span').textContent=f.classList.contains('collapsed')?'＋':'－';
};

/* ============================ 資料載入層 ============================ */
// 依關鍵字詞庫計算 tags + 全文搜尋字串（與後端產生方式一致）
function enrich(d){
  const parts=[d.title,d.code,d.codeShow||'',d.status||'',d.purpose||'',d.overview||''];
  (d.roles||[]).forEach(r=>{parts.push(r[0],r[1]);});
  (d.steps||[]).forEach(s=>{parts.push(s[0],s[1]);});
  (d.exceptions||[]).forEach(e=>parts.push(e));
  const text=parts.join(' ');
  const tags=[];
  KW.vocab.forEach(k=>{
    const needles=KW.alias[k]||[k];
    if(needles.some(n=>text.includes(n))) tags.push(k);
  });
  d.tags=tags; d.search=text;
  return d;
}

// 穩健的 CSV 解析（處理引號、內嵌逗號與換行）
function parseCSV(text){
  text=text.replace(/^﻿/,'').replace(/\r\n/g,'\n').replace(/\r/g,'\n');
  const rows=[]; let row=[], field='', inQ=false;
  for(let i=0;i<text.length;i++){
    const c=text[i];
    if(inQ){
      if(c==='"'){ if(text[i+1]==='"'){field+='"';i++;} else inQ=false; }
      else field+=c;
    } else {
      if(c==='"') inQ=true;
      else if(c===','){ row.push(field); field=''; }
      else if(c==='\n'){ row.push(field); rows.push(row); row=[]; field=''; }
      else field+=c;
    }
  }
  if(field!==''||row.length){ row.push(field); rows.push(row); }
  return rows.filter(r=>r.length>1||(r.length===1&&r[0].trim()!==''));
}
function rowsToObjects(rows){
  if(!rows.length) return [];
  const head=rows[0].map(h=>h.trim());
  return rows.slice(1).map(r=>{const o={};head.forEach((h,i)=>o[h]=(r[i]??'').trim());return o;});
}

// 由兩個 CSV 分頁組出 SOP 物件陣列（與 SEED 同結構）
function buildFromCSV(sopsCsv,stepsCsv){
  const sops=rowsToObjects(parseCSV(sopsCsv));
  const steps=rowsToObjects(parseCSV(stepsCsv));
  const byCode={};
  const list=sops.filter(s=>s.code).map(s=>{
    const roles=(s.roles||'').split('\n').map(x=>x.trim()).filter(Boolean)
      .map(line=>{const i=line.indexOf(CONFIG.ROLE_SEP);return i<0?[line,'']:[line.slice(0,i).trim(),line.slice(i+1).trim()];});
    const exceptions=(s.exceptions||'').split('\n').map(x=>x.trim()).filter(Boolean);
    const modes=(s.modes||'').split(/[,，、]/).map(x=>x.trim()).filter(Boolean);
    const o={code:s.code,codeShow:s.codeShow||'',title:s.title||'',brand:s.brand||'通用',
      platform:s.platform||'通用',modes,delivery:s.delivery||'－',version:s.version||'',
      date:s.date||'',author:s.author||'',status:s.status||'',purpose:s.purpose||'',
      overview:s.overview||'',roles,exceptions,steps:[]};
    byCode[s.code]=o; return o;
  });
  steps.filter(st=>st.code&&byCode[st.code]).sort((a,b)=>(+a.step_no||0)-(+b.step_no||0))
    .forEach(st=>byCode[st.code].steps.push([st.title||'',st.body||'']));
  return list;
}

function fetchText(url){
  // 加上時間戳避免快取，確保「重新整理」拿到最新內容
  const u=url+(url.includes('?')?'&':'?')+'_t='+Date.now();
  return fetch(u,{cache:'no-store'}).then(r=>{if(!r.ok)throw new Error('HTTP '+r.status);return r.text();});
}

function setBadge(kind,text,note){
  const b=document.getElementById('srcBadge');
  b.className='src-badge '+kind;
  document.getElementById('srcText').textContent=text;
  document.getElementById('srcNote').textContent=note||'';
}
function fmtTime(d){const p=n=>String(n).padStart(2,'0');
  return `${d.getFullYear()}/${p(d.getMonth()+1)}/${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}`;}

async function loadData(){
  const {SOPS_CSV_URL,STEPS_CSV_URL}=CONFIG;
  if(SOPS_CSV_URL && STEPS_CSV_URL){
    setBadge('loading','從試算表載入中…');
    try{
      const [sc,stc]=await Promise.all([fetchText(SOPS_CSV_URL),fetchText(STEPS_CSV_URL)]);
      const list=buildFromCSV(sc,stc);
      if(!list.length) throw new Error('試算表沒有資料列');
      DATA=list.map(enrich);
      setBadge('live','資料來源：Google 試算表','更新於 '+fmtTime(new Date()));
      return;
    }catch(err){
      DATA=SEED.map(enrich);
      setBadge('seed','試算表載入失敗，顯示內建範例','請檢查發佈設定或稍後按重新整理（'+err.message+'）');
      return;
    }
  }
  DATA=SEED.map(enrich);
  setBadge('seed','資料來源：內建範例','尚未連結試算表 · 46 份');
}

// 重新繪製整個頁面（載入或重新整理後）
function updateFooter(){
  document.getElementById('footTotal').textContent=DATA.length;
  const order=['MARS','PG','CPG','跨品類','通用'];
  const c={}; DATA.forEach(d=>c[d.brand]=(c[d.brand]||0)+1);
  const parts=order.filter(b=>c[b]).map(b=>`${brandLabel[b]||b} ${c[b]}`);
  Object.keys(c).filter(b=>!order.includes(b)).forEach(b=>parts.push(`${b} ${c[b]}`));
  document.getElementById('footBreak').textContent=parts.join(' · ');
}
function refreshUI(){
  document.getElementById('totalCount').textContent=DATA.length;
  document.getElementById('stats').innerHTML='';
  document.getElementById('filtersBody').innerHTML='';
  document.getElementById('kwSuggest').innerHTML='<span class="lbl">熱門關鍵字</span>';
  openCards.clear();
  buildStats(); buildFilters(); buildKw(); updateFooter(); render();
}

async function init(){
  await loadData();
  refreshUI();
}
document.getElementById('refreshBtn').onclick=async()=>{
  await loadData();
  // 保留使用者目前的搜尋與篩選狀態，只重繪
  document.getElementById('totalCount').textContent=DATA.length;
  document.getElementById('stats').innerHTML='';
  document.getElementById('filtersBody').innerHTML='';
  buildStats(); buildFilters(); buildKw(); updateFooter();
  // 重新套用已勾選的篩選/關鍵字視覺狀態
  document.querySelectorAll('.fopt').forEach(b=>{if(state[b.dataset.k]&&state[b.dataset.k].has(b.dataset.v))b.classList.add('on');});
  document.querySelectorAll('.kw-suggest button').forEach(b=>{if(state.kw.has(b.dataset.kw))b.classList.add('on');});
  render();
};
init();
</script>

</body>
</html>

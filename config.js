/* YFS 營運 SOP 檢索 —— 設定檔（已接 Google 試算表 + 編輯功能） */
window.CONFIG = {
  // ①「SOPs」分頁
  SOPS_CSV_URL:  "https://docs.google.com/spreadsheets/d/e/2PACX-1vSPNkdlIRUKe29iLg4tYGq7k_2p5Hg9twF4voct3xJPQ-Znom4fdbRlHW1m7YeeEXKSecm73upSrT7E/pub?gid=1268551703&single=true&output=csv",

  // ②「Steps」分頁
  STEPS_CSV_URL: "https://docs.google.com/spreadsheets/d/e/2PACX-1vSPNkdlIRUKe29iLg4tYGq7k_2p5Hg9twF4voct3xJPQ-Znom4fdbRlHW1m7YeeEXKSecm73upSrT7E/pub?gid=1656864444&single=true&output=csv",

  // 角色欄位「角色｜職責」的分隔字元，勿改
  ROLE_SEP: "｜",

  /* ── 編輯功能 ── */
  // ③ Google 登入用戶端 ID
  GOOGLE_CLIENT_ID: "529595948839-8fr03m0pp0sakfu24p3s8botoh3qkm6n.apps.googleusercontent.com",

  // ④ Apps Script 網頁應用程式網址（/exec 結尾）
  APPS_SCRIPT_URL: "https://script.google.com/macros/s/AKfycbwJ6Rz3T_3fcIIA-0q-pK4lR_Bmm8r5VlPrgZ2ExySahHyTatTlNdQvQ1dmn4lLAjgG/exec",

  // ⑤ 會看到編輯按鈕的帳號（真正權限由後端白名單把關）
  EDITOR_EMAILS: ["tcss1299@gmail.com"]
};

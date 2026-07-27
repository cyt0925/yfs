/* YFS 營運 SOP 檢索 —— 設定檔（已接 Google 試算表 + 編輯功能） */
window.CONFIG = {
  // ①「SOPs」分頁
  SOPS_CSV_URL:  "https://docs.google.com/spreadsheets/d/e/2PACX-1vSPNkdlIRUKe29iLg4tYGq7k_2p5Hg9twF4voct3xJPQ-Znom4fdbRlHW1m7YeeEXKSecm73upSrT7E/pub?gid=1268551703&single=true&output=csv",

  // ②「Steps」分頁
  STEPS_CSV_URL: "https://docs.google.com/spreadsheets/d/e/2PACX-1vSPNkdlIRUKe29iLg4tYGq7k_2p5Hg9twF4voct3xJPQ-Znom4fdbRlHW1m7YeeEXKSecm73upSrT7E/pub?gid=1656864444&single=true&output=csv",

  // 角色欄位「角色｜職責」的分隔字元，勿改
  ROLE_SEP: "｜",

  /* ── 編輯功能 ──
     APPS_SCRIPT_URL 填好後，網站右上角會出現「🔑 編輯模式」。
     編輯碼不放在這裡（這個檔是公開的）；編輯碼設在 Apps Script 的 Code.gs
     裡（var EDIT_CODE），把碼私下發給可編輯的同事即可。 */
  APPS_SCRIPT_URL: "https://script.google.com/macros/s/AKfycbwJ6Rz3T_3fcIIA-0q-pK4lR_Bmm8r5VlPrgZ2ExySahHyTatTlNdQvQ1dmn4lLAjgG/exec"
};

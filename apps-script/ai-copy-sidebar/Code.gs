/**
 * 橘子工坊｜廣告文案 x 生圖 x 草稿合成 側邊欄（OpenAI 版，v2）
 *
 * 這一版仍然不會在試算表裡插入任何列或欄。側邊欄只讀取目前選取欄的
 * 「品類」「廣告名稱」「文案」「預算分配」當參考，所有產出只顯示在側邊欄，
 * 要不要用、複製到哪裡，都由你自己決定。
 *
 * 相較 v1 的差異：
 *  - 文案：換大模型、加品牌語氣與受眾、產品知識可放在「產品資料」分頁、
 *          時事先提切角再寫文案（兩步驟）、範例與反例加量。
 *  - 生圖：Prompt 改成「程式填空模板」，主色調、視覺風格、版面由設計師寫死的英文描述組成，
 *          模型只負責寫 1～2 句場景道具，不能改顏色與版面。
 *  - 參考圖：直接在側邊欄拖放、選檔、Ctrl+V 貼上或貼網址，不用先整理雲端硬碟。
 *  - 生圖：走 Responses API 的 image_generation 工具，也就是 ChatGPT 生圖背後同一套機制：
 *          GPT 先看懂中文與圖片，自己寫指令、呼叫生圖、記住上一輪，之後用中文繼續改。
 *          不指定影像模型，交給 OpenAI 用最新的（目前是 gpt-image-2）。
 *  - 合成：側邊欄用 canvas 把 AI 底圖 + 真實去背產品 PNG + 中文標題 + 角標 + logo 疊成草稿，
 *          數字與 logo 保證正確，給設計師接手微調。
 *
 * 安裝：
 * 1. 擴充功能 → Apps Script
 * 2. 這份 Code.gs 貼進「Code.gs」（取代原本內容）
 * 3. 新增 HTML 檔「Sidebar」（不要打副檔名），貼上 Sidebar.html 的內容
 * 4. 存檔、回試算表重新整理
 * 5. 選單「AI 文案工具 → 設定 OpenAI API Key」
 * 6. 選單「AI 文案工具 → 開啟 AI 文案側邊欄」，參考圖與產品圖直接拖進側邊欄即可
 * 7. （選用）「建立產品資料分頁」填每個品類的賣點；「建立雲端硬碟常用素材庫」放每次都會用到的產品去背圖與 logo
 */

// ============================================================
// 基本設定
// ============================================================

const API_KEY_PROP = "OPENAI_API_KEY";

// 模型清單：依序嘗試，不存在或不支援就自動往後退。
// 2026-09 平台上的 id：gpt-6-astra（$10/$50）、gpt-5.6-sol（$4/$20）、gpt-5.6-terra（$2/$12）、gpt-5.6-luna（$0.2/$1.2）。
// 文案與看圖寫指令用 Terra 就夠；想更強把 "gpt-5.6-sol" 放到第一個即可。
const TEXT_MODEL_CANDIDATES = ["gpt-5.6-luna", "gpt-5.6-terra", "gpt-4.1"];   // 文案用 Luna 省錢，不夠好再往前換
const IMAGE_QUALITY = "medium"; // 預設生圖品質 low / medium / high，high 一張成本約 3～4 倍

const LABELS = { PRODUCT: "品類", AD_NAME: "廣告名稱", COPY: "文案", BUDGET: "預算分配", DATE: "上檔日期" };
const PRODUCT_SHEET = "產品資料";

// 雲端硬碟只是「可選」的常用素材庫。沒建也能用，所有圖都可以直接在側邊欄上傳。
const DRIVE_FOLDERS = {
  PRODUCTS: "橘子工坊產品去背圖",   // 每次都會用到的去背 PNG
  REFERENCES: "橘子工坊參考成品",   // 設計師做過的成品
  ASSETS: "橘子工坊品牌素材",       // logo.png
  OUTPUT: "橘子工坊生圖"            // 產出存這裡（存檔時自動建立）
};
// 生圖對話用的模型：GPT 負責看圖、寫指令、呼叫 image_generation 工具（畫圖本身是 gpt-image-2）
const CHAT_IMAGE_MODELS = ["gpt-5.6-terra", "gpt-5.6-luna", "gpt-5", "gpt-4.1"];

// ============================================================
// 品牌規範（文案）
// ============================================================

const BRAND = {
  name: "橘子工坊",
  platform: "momo 購物網",
  audience: "30～45 歲、會在 momo 買家用清潔品的女性上班族與媽媽，也有部分注重成分的年輕小家庭。",
  voice: [
    "台灣口語，像朋友在 IG 限動講話，不是官方新聞稿。",
    "句子短、有畫面、有情緒，一行一個重點。",
    "emoji 是節奏工具，一行最多 2 個，不要連發。",
    "用「你」直接對話，不用「您」。",
    "自信但不誇大，不用「史上最」「全台第一」這類無法證實的字眼。"
  ],
  banned: [
    "大陸用語：視頻、質量（指品質）、優化、給力、立馬、性價比、小夥伴、寶寶（指顧客）",
    "醫療療效宣稱：治療、殺死病毒、100% 殺菌、預防疾病",
    "任何價格、折扣、贈品、mo點、日期等具體數字（這些由業務另外填）"
  ],
  facts: [
    "核心成分是天然橘油（冷壓橘皮萃取），主打天然、無螢光劑、無甲醛、無人工香精。",
    "制菌力 99%（可以說「制菌」「抑菌」，不可說「殺菌 100%」或「殺病毒」）。",
    "主要品項：洗衣精、洗衣膠囊、洗碗精、洗衣槽清潔劑、地板清潔劑、蔬果清潔劑。",
    "適用小孩衣物、敏感肌，是家庭客群常見的購買理由。",
    "常用 hashtag：#橘子工坊 #洗衣膠囊 #洗衣精 #0添加 #天然配方 #無螢光劑 #適用幼兒衣物"
  ]
};

/**
 * 多品牌：這張表不只橘子工坊。key 是品牌名，keywords 用來從品類／廣告名稱／分頁名自動偵測。
 * facts 只寫確定的事，不確定的交給「產品資料」分頁或業務補充。
 */
const BRAND_PROFILES = {
  "橘子工坊": { keywords: ["橘子工坊", "Orange House", "橘油"], facts: BRAND.facts, hashtags: "#橘子工坊 #天然配方 #無螢光劑", note: "家用清潔品牌，主色橘。" },
  "五月花": { keywords: ["五月花", "May Flower", "衛生紙", "抽取式", "廚房紙巾", "紙品", "濕巾"], facts: ["永豐餘旗下家用紙品牌，品項含抽取式衛生紙、廚房紙巾、濕式衛生紙等。", "賣點以「產品資料」分頁或業務補充為準，不可自行編造成分或認證。"], hashtags: "#五月花", note: "家用紙品牌，包裝常見深藍或藍紫底。" },
  "得意": { keywords: ["得意"], facts: ["永豐餘旗下家用紙品牌。", "賣點以「產品資料」分頁或業務補充為準。"], hashtags: "#得意", note: "家用紙品牌。" },
  "蒲公英": { keywords: ["蒲公英"], facts: ["永豐餘旗下環保再生紙品牌。", "賣點以「產品資料」分頁或業務補充為準。"], hashtags: "#蒲公英", note: "環保紙品牌，主色綠。" },
  "其他": { keywords: [], facts: ["品牌資訊以業務補充與產品圖為準。"], hashtags: "", note: "" }
};
function detectBrand(text) {
  text = String(text || "");
  const keys = Object.keys(BRAND_PROFILES);
  for (let i = 0; i < keys.length; i++) {
    if (BRAND_PROFILES[keys[i]].keywords.some(k => text.indexOf(k) !== -1)) return keys[i];
  }
  return "";
}
function resolveBrand(input) {
  const name = (input && input.brand && BRAND_PROFILES[input.brand]) ? input.brand
    : detectBrand([input && input.product, input && input.adName, input && input.sheetName].join(" ")) || "橘子工坊";
  return Object.assign({ name: name }, BRAND_PROFILES[name]);
}

// 「產品資料」分頁不存在或找不到品項時的後備知識
const PRODUCT_FACTS = {
  "洗衣膠囊": {
    sellingPoints: "一球搞定不用計量、天然橘油深入纖維、制菌去味、洗完衣物清新無殘留",
    scenes: "上班累到不想計量、小孩衣服多、宿舍套房洗衣、出差旅行帶幾球",
    hashtags: "#橘子工坊 #洗衣膠囊 #0添加 #天然配方 #無螢光劑 #適用幼兒衣物"
  },
  "洗衣精": {
    sellingPoints: "天然橘油制菌去污、洗淨5大病毒（依法規可用字眼為制菌）、無螢光劑、適用幼兒衣物",
    scenes: "家庭大量洗衣、寶寶衣物、運動後汗臭、梅雨季室曬悶味",
    hashtags: "#橘子工坊 #洗衣精 #天然配方 #無螢光劑 #適用幼兒衣物"
  },
  "洗碗精": {
    sellingPoints: "天然橘油去油、不傷手、沖洗快速無殘留、可洗蔬果餐具",
    scenes: "火鍋後一堆油碗、下班懶得洗碗、便當盒油垢、租屋族小廚房",
    hashtags: "#橘子工坊 #洗碗精 #天然配方 #不傷手"
  },
  "洗衣槽": {
    sellingPoints: "深層清潔洗衣槽黴垢、去除異味來源、定期保養洗衣機",
    scenes: "衣服洗完還是有味道、洗衣機發霉、換季大掃除",
    hashtags: "#橘子工坊 #洗衣槽清潔 #天然配方"
  },
  "地板": {
    sellingPoints: "天然橘油去油污、無化學殘留、小孩寵物爬地板安心",
    scenes: "小孩在地上爬、寵物家庭、廚房油膩地板、年終大掃除",
    hashtags: "#橘子工坊 #地板清潔 #天然配方 #寵物友善"
  }
};

const TONES = ["促銷型", "時事型", "痛點型", "趣味型", "聯名型"];

const TONE_GUIDE = {
  "促銷型": "營造「現在不買會後悔」的急迫感與撿到便宜的爽感，但不能出現任何數字，用「這檔」「囤起來」「年度最狂」這類字眼。",
  "時事型": "把提供的時事或節慶，跟洗衣／清潔的情境找到一個具體連結點（不是硬湊），第一行先提時事，第二三行轉到產品。",
  "痛點型": "第一行具體描述一個讀者會點頭的困擾場景（味道、污漬、時間、麻煩），第二行點出橘子工坊怎麼解，第三行給輕鬆的結果畫面。",
  "趣味型": "第一行埋一個梗（雙關、諧音、生活吐槽），第二行帶產品，第三行回收同一個梗。要有笑點，不要只是可愛。",
  "聯名型": "同時提到聯名雙方，找到兩個品牌共同的價值（健康、安心、美好生活），語氣正面溫暖，不用吐槽梗。"
};

/** 五種調性的真實範例，用來校準 AI 的口吻與節奏 */
const TONE_EXAMPLES = {
  "促銷型": [
    "618不囤？難道要等原價才買😱\n橘子工坊年度最狂大促🎉\n計算機不用按了，這檔絕對賺🤑",
    "今天少猶豫一秒，接下來半年每洗一次衣服👉\n你就省一筆！\n囤起來，你就是家裡的省錢之神😏",
    "618中慶🎉預備備～GO\n橘子工坊 狂歡開搶🔥\n錯過再等一年！",
    "618沒搶到❓沒關係～\n橘子工坊熱銷回歸📣\n最後倒數，這次別再錯過",
    "衣服會一直髒，優惠不會一直在\n橘子工坊 這檔直接囤到明年\n買過的都說：早知道多買幾組😌"
  ],
  "時事型": [
    "熱血世足賽⚽衣起清新應援✨\n對抗頑固污漬，你需要世界級的頂尖防守球員🥅",
    "吶喊可以瘋狂，球衣不能髒！\n橘子工坊天然制菌洗衣膠囊🍊\n洗淨你流的每一滴熱血汗水✨",
    "颱風假在家追劇一整天📺\n衣服堆成山不想面對？\n丟一球橘子工坊，追劇追到洗衣機叫你",
    "開學季，小孩的白襪永遠是灰的🧦\n橘子工坊天然橘油深入纖維\n洗回開學第一天的白",
    "梅雨季連下十天☔衣服曬不乾還有悶味\n橘子工坊制菌洗衣精\n室曬也不怕，乾了就是乾淨的味道"
  ],
  "痛點型": [
    "汗臭、室曬悶味、火鍋味🍲\n橘子工坊直接幫你下架異味✨\n天然橘油深入纖維🍊把臭根源直接OUT👋",
    "走出門3秒就開始爆汗💦\n明明穿著「剛洗好的衣服」，卻隱隱約約聞到昨天的汗酸味🤢\n別再用化學香精硬蓋了！",
    "小孩衣服上永遠有不明來源的黃漬👶\n洗了三次還在，到底是什麼？\n橘子工坊天然橘油，溫和但不手軟",
    "健身完的衣服放到隔天，整個房間都知道🏋️\n橘子工坊制菌99%\n把味道連根拔起，不是蓋過去"
  ],
  "趣味型": [
    "同事甩的鍋🍳客戶畫的餅🍕橘子工坊都洗得掉💦\n天然橘油深入纖維🍊洗掉你被壓榨的一天🤜",
    "夏天一秒爆汗，衣服異味重到連金特務都覺得棘手🕵️\n防範夏日汗味被炎上🔥就用橘子工坊洗衣膠囊🍊\n一球搞定髒污細菌🦠不想變汗味戰警的趕快下單",
    "上班累到大腦拒絕思考，下班只想拒絕內耗、速洗速躺平！\n不用花時間糾結，交給橘子工坊洗碗精！",
    "衣服：我沒有臭\n鼻子：你有\n橘子工坊：我來當公道伯🍊",
    "洗衣精倒多倒少永遠在賭🎲\n橘子工坊洗衣膠囊，一球就是標準答案\n人生很多事沒有正解，但洗衣有"
  ],
  "聯名型": [
    "娘家x橘子工坊 品牌聯合慶\n✨健康雙效守護，家倍安心✨",
    "美体專科x橘子工坊✨美力不橘限✨\n從飲食管理到衣物潔淨，打造清新無負擔的夏日生活！",
    "桂格x橘子工坊 早安家庭日🌅\n一杯營養顧內在，一球乾淨顧外在\n全家的好日子從早上開始"
  ]
};

/** 反例：告訴模型什麼叫「寫壞了」 */
const BAD_EXAMPLES = [
  { text: "橘子工坊洗衣精，優質產品，值得您擁有！", why: "官方腔、用「您」、沒有畫面、沒有情緒。" },
  { text: "夏天到了🌞🌊🍉衣服要洗乾淨✨✨✨快來買🛒🛒", why: "emoji 堆砌、第一行跟產品沒有連結、沒有鉤子。" },
  { text: "性價比超高的洗衣液，質量沒問題", why: "大陸用語。" },
  { text: "滿1099折150！限時三天！", why: "出現數字與優惠條件，這些由業務填，文案不該寫。" },
  { text: "橘子工坊能殺死所有病毒，預防感冒", why: "醫療療效宣稱，違反廣告法規。" }
];

// ============================================================
// 生圖規範（設計師可直接改這裡的英文描述）
// ============================================================

/** 主色調：放在 Prompt 第一句，權重最高。選「自訂」時用使用者輸入的顏色文字。 */
const PALETTES = [
  { key: "品牌橘", en: "Dominant color palette: warm citrus orange (#F28C28) and cream white, with small accents of fresh leaf green. Warm, sunny, clean.", textColor: "#FFFFFF", textStroke: "#B8440A", badgeColor: "#D9520A" },
  { key: "藍色夏日", en: "Dominant color palette: cool summer blues, sky blue (#4FA8E8) to aqua (#7FD3E6) gradient, white highlights, a hint of sunlight yellow. Orange must NOT be a dominant color; at most a tiny accent.", textColor: "#FFFFFF", textStroke: "#1B4F9C", badgeColor: "#FF6B35" },
  { key: "清新綠", en: "Dominant color palette: fresh mint and leaf greens (#7FC97F, #CFEFD4) with white. Natural, botanical, breathable.", textColor: "#FFFFFF", textStroke: "#2E7D4F", badgeColor: "#F28C28" },
  { key: "節慶紅金", en: "Dominant color palette: festive red (#C8102E) and warm gold (#E8B04B) with cream highlights. Celebratory, premium, Lunar New Year mood without any text or characters.", textColor: "#FFF4D6", textStroke: "#8A0F1E", badgeColor: "#E8B04B" },
  { key: "極簡白", en: "Dominant color palette: clean white and very light warm grey, with one soft orange accent. Minimal, airy, studio look.", textColor: "#D9520A", textStroke: "#FFFFFF", badgeColor: "#D9520A" },
  { key: "自訂", en: "", textColor: "#FFFFFF", textStroke: "#333333", badgeColor: "#D9520A" }
];

/**
 * 留白版面：en 給「只畫背景」模式當指令，zones 給第四步 canvas 合成用（畫面比例 0~1）。
 */
const LAYOUTS = [
  {
    key: "上標題．下產品",
    en: "Composition: the upper third of the picture is ONLY plain empty sky, a plain wall, or a smooth soft gradient. Nothing hangs, stands or floats there: no clothes line, no shirt, no plants, no props, no objects of any kind. All objects sit in the lower half of the frame. The lower 40% is a simple flat surface (table, floor, shelf or platform) seen slightly from above, with its center completely clear. Small props only at the far left and far right edges of that surface.",
    zones: { headline: { x: 0.06, y: 0.05, w: 0.88, h: 0.28 }, products: { x: 0.05, y: 0.55, w: 0.90, h: 0.38 }, badge: "top-right", logo: "top-left" }
  },
  {
    key: "左產品．右標題",
    en: "Composition: the entire right half of the picture is ONLY plain empty wall, sky or a smooth soft gradient with no objects at all. All objects are in the left half. The left half is a simple flat surface (table or shelf) whose center is completely clear, with one or two small props at its far left edge only.",
    zones: { headline: { x: 0.52, y: 0.12, w: 0.44, h: 0.40 }, products: { x: 0.03, y: 0.30, w: 0.48, h: 0.60 }, badge: "top-right", logo: "top-left" }
  },
  {
    key: "中央產品．上下標題",
    en: "Composition: the top quarter and the bottom 15% of the picture are ONLY plain empty background (sky, wall or smooth gradient) with no objects. The middle band is an open, empty stage or tabletop with a soft spotlight, its center completely clear. Small props only at the far left and right edges of the stage.",
    zones: { headline: { x: 0.08, y: 0.04, w: 0.84, h: 0.20 }, products: { x: 0.10, y: 0.30, w: 0.80, h: 0.52 }, badge: "top-right", logo: "top-left" }
  }
];

const RATIOS = ["1:1", "9:16", "4:5", "16:9"];   // momo CPAS 素材以 1:1 為主
const SIZE_MAP = { "9:16": "1024x1536", "4:5": "1024x1536", "16:9": "1536x1024", "1:1": "1024x1024" };

// ============================================================
// 選單
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("AI 文案工具")
    .addItem("開啟 AI 文案側邊欄", "showSidebar")
    .addSeparator()
    .addItem("設定 OpenAI API Key", "setApiKey")
    .addItem("建立產品資料分頁（選用）", "setupProductSheet")
    .addItem("建立雲端硬碟常用素材庫（選用）", "setupDriveFolders")
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("Sidebar")
    .setTitle("AI 文案助手")
    .setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
}

function setApiKey() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt(
    "設定 OpenAI API Key",
    "請貼上你的 API Key（在 platform.openai.com/api-keys 建立，開頭是 sk-proj- 或 sk-）：",
    ui.ButtonSet.OK_CANCEL
  );
  if (res.getSelectedButton() !== ui.Button.OK) return;

  const key = res.getResponseText().trim();
  if (!key) { ui.alert("沒有輸入內容，未儲存。"); return; }
  PropertiesService.getScriptProperties().setProperty(API_KEY_PROP, key);
  ui.alert("API Key 已儲存，整份檔案都能直接使用側邊欄。");
}

/** 建立四個雲端硬碟資料夾（已存在就跳過），並告知路徑 */
function setupDriveFolders() {
  const created = [];
  Object.keys(DRIVE_FOLDERS).forEach(k => {
    const name = DRIVE_FOLDERS[k];
    const it = DriveApp.getFoldersByName(name);
    if (!it.hasNext()) { DriveApp.createFolder(name); created.push(name); }
  });
  SpreadsheetApp.getUi().alert(
    "常用素材庫已就位（這是選用功能，圖片也可以直接拖進側邊欄）。\n\n" +
    "・" + DRIVE_FOLDERS.PRODUCTS + "：每次都會用到的去背產品 PNG\n" +
    "・" + DRIVE_FOLDERS.REFERENCES + "：設計師做過的成品\n" +
    "・" + DRIVE_FOLDERS.ASSETS + "：logo.png（檔名含 logo 即可）\n" +
    "・" + DRIVE_FOLDERS.OUTPUT + "：產出圖會存這裡\n\n" +
    (created.length ? "本次新建：" + created.join("、") : "全部都已存在，沒有新建。")
  );
}

/** 建立「產品資料」分頁並填入範本列 */
function setupProductSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(PRODUCT_SHEET);
  if (sh) { SpreadsheetApp.getUi().alert("「" + PRODUCT_SHEET + "」分頁已存在，未變更。"); return; }
  sh = ss.insertSheet(PRODUCT_SHEET);
  const rows = [["品類關鍵字", "賣點/差異點", "常見使用情境", "常用 hashtag", "備註"]];
  Object.keys(PRODUCT_FACTS).forEach(k => {
    const f = PRODUCT_FACTS[k];
    rows.push([k, f.sellingPoints, f.scenes, f.hashtags, ""]);
  });
  sh.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  sh.getRange(1, 1, 1, rows[0].length).setFontWeight("bold").setBackground("#FCE5CD");
  sh.setColumnWidths(1, 5, 220);
  sh.setFrozenRows(1);
  SpreadsheetApp.getUi().alert(
    "已建立「" + PRODUCT_SHEET + "」分頁。\n\n" +
    "A 欄「品類關鍵字」只要是入稿表品類文字裡包含的字（例如「洗衣膠囊」），側邊欄就會自動帶入該列的賣點。\n" +
    "可以自由增刪列，不影響入稿表。"
  );
}

// ============================================================
// 讀取試算表內容
// ============================================================

function findLabelRow(sheet, label) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return -1;
  const values = sheet.getRange(1, 1, lastRow, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === label) return i + 1;
  }
  return -1;
}

function columnToLetter(col) {
  let letter = "";
  while (col > 0) {
    const rem = (col - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
}

/** 從「產品資料」分頁找品項；找不到就用內建 PRODUCT_FACTS */
function getProductInfo(productText) {
  const text = String(productText || "");
  if (!text) return null;

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PRODUCT_SHEET);
  if (sh && sh.getLastRow() >= 2) {
    const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues();
    for (let i = 0; i < rows.length; i++) {
      const kw = String(rows[i][0]).trim();
      if (kw && text.indexOf(kw) !== -1) {
        return { source: PRODUCT_SHEET, key: kw, sellingPoints: String(rows[i][1]), scenes: String(rows[i][2]), hashtags: String(rows[i][3]) };
      }
    }
  }
  const keys = Object.keys(PRODUCT_FACTS);
  for (let i = 0; i < keys.length; i++) {
    if (text.indexOf(keys[i]) !== -1) {
      return Object.assign({ source: "內建", key: keys[i] }, PRODUCT_FACTS[keys[i]]);
    }
  }
  return null;
}

/** 側邊欄開啟或按「重新讀取」時呼叫 */
function getColumnContext() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const cell = sheet.getActiveCell();
  const col = cell.getColumn();

  if (col < 2) {
    return { ok: false, message: "請先點選任一檔期的欄位（B 欄以後），不要停在 A 欄。" };
  }

  const readCell = label => {
    const r = findLabelRow(sheet, label);
    return r === -1 ? "" : String(sheet.getRange(r, col).getValue() || "").trim();
  };
  const adName = readCell(LABELS.AD_NAME);
  let product = readCell(LABELS.PRODUCT);
  if (!product) product = adName;

  return {
    ok: true,
    sheetName: sheet.getName(),
    columnLabel: columnToLetter(col),
    product: product,
    adName: adName,
    copyText: readCell(LABELS.COPY),
    budgetText: readCell(LABELS.BUDGET),
    dateText: readCell(LABELS.DATE),
    brandGuess: detectBrand(product + " " + adName + " " + sheet.getName()),
    brands: Object.keys(BRAND_PROFILES),
    productInfo: getProductInfo(product),
    hasApiKey: !!PropertiesService.getScriptProperties().getProperty(API_KEY_PROP)
  };
}

function getFormOptions() {
  return {
    tones: TONES,
    palettes: PALETTES,
    layouts: LAYOUTS,
    ratios: RATIOS
  };
}

// ============================================================
// OpenAI 共用呼叫
// ============================================================

function requireApiKey() {
  const key = PropertiesService.getScriptProperties().getProperty(API_KEY_PROP);
  if (!key) throw new Error("尚未設定 API Key，請先執行選單「AI 文案工具 → 設定 OpenAI API Key」。");
  return key;
}

/**
 * 400 錯誤裡若指出某參數不被支援（例如 gpt-5.x 不吃 temperature、gpt-image-2 不吃 input_fidelity），
 * 從 payload 頂層或 tools[0] 把它拿掉。回傳 true 表示有拿掉、可以重送。
 */
function stripUnsupportedParam(payload, body) {
  const m = body.match(/does not support the '([a-z_.]+)' parameter/) ||
            body.match(/Unsupported parameter:? '?([a-z_.]+)'?/i) ||
            body.match(/'([a-z_.]+)' is not supported/) ||
            body.match(/Unsupported value:? '([a-z_.]+)'/i);
  if (!m) return false;
  const key = m[1].split(".").pop();
  if (payload[key] !== undefined) { delete payload[key]; return true; }
  if (payload.tools && payload.tools[0] && payload.tools[0][key] !== undefined) { delete payload.tools[0][key]; return true; }
  return false;
}

/** 呼叫 chat completions，依序試 TEXT_MODEL_CANDIDATES，回傳解析後的 JSON */
function callChatJson(prompt, schemaName, schema) {
  const apiKey = requireApiKey();
  const payload = {
    messages: [{ role: "user", content: prompt }],
    temperature: 0.9,
    response_format: {
      type: "json_schema",
      json_schema: { name: schemaName, strict: true, schema: schema }
    }
  };
  const options = {
    method: "post",
    contentType: "application/json",
    headers: { Authorization: "Bearer " + apiKey },
    muteHttpExceptions: true
  };

  let lastError = "";
  for (let i = 0; i < TEXT_MODEL_CANDIDATES.length; i++) {
    const model = TEXT_MODEL_CANDIDATES[i];
    const p = JSON.parse(JSON.stringify(Object.assign({ model: model }, payload)));

    for (let attempt = 0; attempt < 3; attempt++) {
      options.payload = JSON.stringify(p);
      const res = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", options);
      const code = res.getResponseCode();
      const body = res.getContentText();

      if (code === 200) {
        const json = JSON.parse(body);
        return JSON.parse(json.choices[0].message.content);
      }
      if (code === 401) throw new Error("API Key 無效或已被刪除。請重新執行「設定 OpenAI API Key」。");
      if (code === 400 && stripUnsupportedParam(p, body)) continue;   // 拿掉不支援的參數再送一次
      if (code === 404 || (code === 400 && body.indexOf("model") !== -1 && (body.indexOf("does not exist") !== -1 || body.indexOf("not found") !== -1))) {
        lastError = "模型 " + model + " 無法使用";
        break;
      }
      if (code === 429 && body.indexOf("insufficient_quota") !== -1) {
        throw new Error("OpenAI 帳戶額度不足。請到 platform.openai.com → Settings → Billing 儲值。");
      }
      if (code === 429 || code >= 500) {
        lastError = "模型 " + model + " 暫時無法回應（" + code + "）";
        if (attempt === 0) { Utilities.sleep(2000); continue; }
        break;
      }
      throw new Error("API 錯誤（" + code + "）：\n" + body.substring(0, 300));
    }
  }
  throw new Error("生成失敗（" + lastError + "）。稍候再試一次，或把這則錯誤訊息貼給工程師調整模型清單。");
}

// ============================================================
// 文案：共用的品牌段落
// ============================================================

function brandBlock(input) {
  const info = input.productInfo || getProductInfo(input.product);
  const productLines = info
    ? "已知賣點：" + info.sellingPoints + "\n常見情境：" + info.scenes
    : "（沒有這個品項的資料，請依商品名稱合理推斷，不要編造成分）";

  const brand = resolveBrand(input);
  return `【品牌】${brand.name}，投放平台：${BRAND.platform}。${brand.note || ""}
【受眾】${BRAND.audience}
【語氣】
${BRAND.voice.map(v => "- " + v).join("\n")}
【品牌事實（只能用這些，不可自行加成分或功效）】
${brand.facts.map(v => "- " + v).join("\n")}
【禁止】
${BRAND.banned.map(v => "- " + v).join("\n")}

【商品】${input.product}
${productLines}
【業務補充賣點】${input.sellingPoint || "（未提供）"}`;
}

// ============================================================
// 文案步驟一：針對時事／梗料提切角
// ============================================================

function proposeAngles(input) {
  if (!input.product) throw new Error("請先填寫「商品」。");
  if (!input.trend) throw new Error("請先在「本週時事／梗料」填寫內容，才能提切角。");

  const prompt = `${brandBlock(input)}

【任務】
下面是業務提供的本週時事或梗料。請提出 5 個「把這個時事連到 ${input.product} 」的文案切角，讓業務挑選。
每個切角要有：
- title：6～12 字的切角名稱
- hook：這個切角的第一行文案（示範口吻，繁體中文，可含 1 個 emoji）
- why：一句話說明時事跟產品的連結點在哪，為什麼讀者會覺得合理而不是硬湊

5 個切角的視角要不同（例如：當事人視角、旁觀者吐槽、家庭場景、上班族場景、反向操作）。
調性偏好：${input.tone || "不限"}。
${input.tone ? "調性說明：" + TONE_GUIDE[input.tone] : ""}

【本週時事／梗料】
${input.trend}

只回傳 JSON。`;

  const schema = {
    type: "object",
    properties: {
      angles: {
        type: "array",
        items: {
          type: "object",
          properties: { title: { type: "string" }, hook: { type: "string" }, why: { type: "string" } },
          required: ["title", "hook", "why"],
          additionalProperties: false
        }
      }
    },
    required: ["angles"],
    additionalProperties: false
  };
  const out = callChatJson(prompt, "trend_angles", schema);
  return { angles: (out.angles || []).slice(0, 5) };
}

// ============================================================
// 文案步驟二：生成三版文案 + 生圖場景
// ============================================================

function buildCopyPrompt(input) {
  const examples = (TONE_EXAMPLES[input.tone] || [])
    .map((t, i) => `範例${i + 1}：\n${t}`).join("\n\n");
  const bad = BAD_EXAMPLES.map(b => `✗「${b.text}」→ ${b.why}`).join("\n");

  const angleBlock = input.angle
    ? `【業務已選定的切角，三個版本都要圍繞它，只是換情境或視角】\n${input.angle.title}：${input.angle.why}\n參考首句：${input.angle.hook}`
    : "";

  const trendBlock = input.trend
    ? `【本週時事／梗料】\n${input.trend}`
    : "【本週時事／梗料】（未提供）" + (input.tone === "時事型"
      ? "\n→ 時事型但沒有時事，改用普遍的季節或生活情境（天氣、上班、家務、放假），不要臆測近期新聞。"
      : "");

  return `你是「${resolveBrand(input).name}」的社群廣告文案手，專寫 ${BRAND.platform} 的短文案。你只負責「前段情境鉤子」，後段的價格、優惠、贈品由業務另外填寫。

${brandBlock(input)}

【任務】
寫出 3 個「${input.tone}」調性、切角明顯不同的三行式文案草稿，並寫一段生圖用的英文「場景描述」。

【這個調性的寫法】
${TONE_GUIDE[input.tone] || ""}

【這個調性的真實範例，模仿口吻、節奏與 emoji 習慣，但不要抄內容】
${examples}

【寫壞了的樣子，避免】
${bad}

${angleBlock}

${trendBlock}

【文案規則】
- 每則 3 行，繁體中文台灣用語。
- 第一行是鉤子，一定要有畫面或情緒，讓人想往下看；第二行帶到 ${resolveBrand(input).name} 與商品；第三行收尾（結果畫面、金句或回收梗）。
- 三個版本的切角要真的不同（不同情境、不同人物視角、不同情緒），不可以只是換同義詞。
- 絕對不可出現任何價格、折扣、贈品、mo點、日期等具體數字或優惠條件。
- 每行最多 2 個 emoji，可以完全不用。

【生圖場景描述規則（scene 欄位）】
- 英文，1～2 句，只描述「背景環境」：地點（牆面、桌面、陽台、窗邊）、材質、季節感，最多 2 個放在邊緣的小道具（毛巾、摺好的衣物、小植物）。
- 背景是空舞台，不能有主體物件：不要寫掛著的衣服、洗衣機、行李箱、家電、家具。
- 不要寫任何顏色（顏色由設計師另外指定）、不要寫版面配置、不要寫文字或產品、不要寫光線。
- 場景可以輕輕呼應文案情境（例如文案講汗臭，邊緣放一條運動毛巾），但不要把文案的故事整個搬進畫面。

只回傳 JSON。`;
}

/**
 * 側邊欄呼叫：把參考圖分析與文案情境，整理成一段業務看得懂、可以直接改的中文背景描述。
 */
function describeBackgroundZh(input) {
  const ref = input.refAnalysis || null;
  const parts = [];
  if (ref) parts.push("參考圖分析（英文）：\n背景顏色：" + (ref.backgroundPaletteEn || ref.paletteEn) + "\n風格：" + ref.styleEn + "\n環境：" + ref.sceneEn);
  if (input.copyScene) parts.push("文案想呼應的情境（英文）：" + input.copyScene);
  if (!parts.length) return { zh: "" };

  const prompt = `你是電商視覺的美術指導。請把下面的資料整理成一段「背景圖」的中文描述，給業務看、讓業務可以直接修改。
要求：
- 繁體中文，3～4 句，口語、具體，像跟設計師交代。
- 依序講：背景是什麼環境與顏色、光線與質感、邊緣放什麼小道具（最多 2 個）、哪一區要留空給標題與產品。
- 這是「背景」，不能提到產品、包裝、文字、logo，也不能有掛著的衣服、洗衣機、行李箱這類主體物件。
- ${ref ? "以參考圖分析為主；文案情境只拿來挑小道具，不要把文案的故事整個搬進畫面。" : "依文案情境挑一個乾淨的環境，不要把故事整個搬進畫面。"}
- 商品是「${input.product || "橘子工坊清潔用品"}」，道具要跟它的使用情境相關。

${parts.join("\n\n")}

只回傳 JSON。`;
  const schema = { type: "object", properties: { zh: { type: "string" } }, required: ["zh"], additionalProperties: false };
  return callChatJson(prompt, "background_description", schema);
}

/** 側邊欄呼叫：生成三版文案草稿 + 場景 + 組好的生圖 Prompt */
function generateCopy(input) {
  if (!input.product || !input.tone) {
    throw new Error("請至少填寫「商品」與「文案調性」再生成。");
  }
  const schema = {
    type: "object",
    properties: {
      copy1: { type: "string" },
      copy2: { type: "string" },
      copy3: { type: "string" },
      scene: { type: "string" }
    },
    required: ["copy1", "copy2", "copy3", "scene"],
    additionalProperties: false
  };
  return callChatJson(buildCopyPrompt(input), "ad_copy_draft", schema);
}

// ============================================================
// 雲端硬碟素材
// ============================================================

function getFolderByName(name) {
  const it = DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : null;
}

function isImageFile(file) {
  const m = file.getMimeType() || "";
  return m === "image/png" || m === "image/jpeg" || m === "image/webp";
}

function listFolderImages(folderName, withThumb) {
  const folder = getFolderByName(folderName);
  if (!folder) return { exists: false, files: [] };
  const files = [];
  const it = folder.getFiles();
  while (it.hasNext()) {
    const f = it.next();
    if (!isImageFile(f)) continue;
    const item = { id: f.getId(), name: f.getName().replace(/\.(png|jpe?g|webp)$/i, ""), mime: f.getMimeType(), size: f.getSize() };
    if (withThumb) {
      try {
        const th = f.getThumbnail();
        if (th) item.thumb = "data:" + th.getContentType() + ";base64," + Utilities.base64Encode(th.getBytes());
      } catch (e) { /* 縮圖失敗不影響使用 */ }
    }
    files.push(item);
  }
  files.sort((a, b) => a.name.localeCompare(b.name, "zh-Hant"));
  return { exists: true, files: files };
}

/** 側邊欄呼叫：列出去背產品圖、參考成品、logo */
function listAssets() {
  const products = listFolderImages(DRIVE_FOLDERS.PRODUCTS, true);
  const references = listFolderImages(DRIVE_FOLDERS.REFERENCES, true);
  const assets = listFolderImages(DRIVE_FOLDERS.ASSETS, false);
  const logo = assets.files.find(f => /logo/i.test(f.name)) || null;
  return { products: products, references: references, logo: logo };
}

/** 側邊欄呼叫：取單一檔案的 base64（canvas 合成用） */
function getFileBase64(fileId) {
  const f = DriveApp.getFileById(fileId);
  const blob = f.getBlob();
  return { id: fileId, name: f.getName(), dataUrl: "data:" + blob.getContentType() + ";base64," + Utilities.base64Encode(blob.getBytes()) };
}

// ============================================================
// 生圖
// ============================================================

/** 橘子工坊 momo CPAS 素材的固定版型（從設計師歷年成品歸納，設計師可直接改這段） */
const HOUSE_TEMPLATE = [
  "【永豐餘消費品 momo 素材固定版型，除非業務另外指定，一律照此排】",
  "1. 品牌 logo 放左上角（聯名時兩個 logo 並排在上方）。",
  "2. 標題 1～2 行，商品名可直接寫進標題。位置依【文字排法】。字體處理見下方【字體規格】。",
  "3. 三到四個圓形賣點徽章（白底或半透明圓形＋小圖示＋4～6 字），散布在產品旁邊，不能擋到產品正面。",
  "4. 產品放中間偏下，是視覺重心之一但不是唯一主角；多包裝時扇形排開；膠囊、水花、光點帶出動態。",
  "5. 檔期、優惠、價格的擺法依【文字排法】；價格永遠是全圖最大的字，mo 點以圓形 mo 圖示呈現。",
  "6. 背景高彩度、明亮、有光線感，主題呼應文案（水、冰、運動、節慶皆可），但不能雜到蓋住文字。",
  "7. 圖上文字只能有業務給的那幾組，繁體中文，字要大、正確、清楚。",
  "",
  "【字體規格，照做，不要用招牌式的平板字】",
  "- 主標：超粗圓黑體，整組向右上傾斜約 8 度（斜體感），兩行錯落、第二行縮排，字距壓緊到筆畫微微相碰。主色填色＋深色粗描邊（約字高 6%）＋往右下的深色長投影，字要像浮在畫面上。關鍵字（產品名、動作詞）改【強調色】，其餘用主色。",
  "- 價格：全圖最大的字。「$」與數字用另一套斜體的展示型數字字體，用【強調色】，粗描邊；「/顆」「up」「起」等單位縮小三分之一，疊在數字右側。「下殺」「直降」這類動詞用主色、中等字級，放在數字左邊。",
  "- 優惠：「買2組送150點」裡的數字放大一級並改【強調色】，其餘用主色；mo 點一律畫成圓形 mo 圖示加「150點」，不要寫成「mo點」三個字。",
  "- 小字註記（限量須登記、圖片僅供參考等）：最小字級、白色或淺灰、無描邊，放在底部資訊帶的最下緣。",
  "- 除了【文字排法】允許的資訊帶或價格牌以外，任何文字都不可以有色塊底框、膠囊框或橫幅框；文字直接壓在畫面上，靠描邊與陰影和背景分離。",
  "- 中文字與數字要用不同字體，數字更粗更斜；同一組文字內至少有兩種字級。"
].join("\n");

/** 文字排法：預設「自動變化」，避免每張都把字堆在底部 */
const TEXT_LAYOUTS = {
  "auto": "【文字排法：自動變化】每次生成請從下面幾種排法挑一種不同的，不要固定用底部資訊帶：(a) 主標左上兩行，價格做成右下角斜貼的圓形或多邊形價格牌，優惠與檔期一行小字貼在價格牌上方；(b) 主標置中偏上、價格緊接在主標下方並排成一組視覺重心，優惠與檔期放最下緣一行小字，不加底帶；(c) 主標左側直排或左對齊佔左半，右半留給產品，價格與優惠疊在產品下方的斜切色帶上；(d) 自家慣例：底部滿版深色資訊帶，左邊「檔期 momo限定 ▸ 優惠」，右邊特大價格。文字不可全部堆在畫面下方三分之一。",
  "band": "【文字排法：底部資訊帶（自家慣例）】主標左上兩行；底部一條滿版深色資訊帶（冷色系深藍、節慶深紅），左邊「檔期 momo限定 ▸ 優惠」，右邊特大價格；小字註記在資訊帶最下緣。",
  "tag": "【文字排法：主標左上＋價格牌】主標左上兩行；價格做成右下角斜貼約 12 度的圓形或爆炸形價格牌，牌內「$數字」特大、單位小字；優惠與檔期一行放在價格牌正上方；不要底部資訊帶。",
  "center": "【文字排法：置中疊排】主標置中偏上，價格緊接在主標下方置中、跟主標並排成一組視覺重心；產品在下方左右分列；優惠與檔期一行小字放畫面最下緣，不加底帶。",
  "side": "【文字排法：左字右圖】左半部由上到下依序放主標（左對齊）、優惠、價格，右半部放產品與主視覺；價格與優惠壓在一條從左下往右上斜切的色帶上；不要底部資訊帶。"
};
/** 強調色：預設依背景挑對比色，不要每張都黃 */
const ACCENTS = {
  "auto": "【強調色：依背景自動對比】先決定背景主色，再挑強調色：藍色系背景用亮黃或橘；橘色系背景用白或深藍；綠色系背景用白或檸檬黃；紅色或深色背景用金或白；淺色背景用品牌橘或深藍。主色（一般文字）用白或深藍擇一與背景對比最大的。不要習慣性用黃色。",
  "yellow": "【強調色：亮黃】關鍵字與價格數字用亮黃 (#FFD400)，主色白。",
  "orange": "【強調色：品牌橘】關鍵字與價格數字用品牌橘 (#F28C28)，主色白或深藍。",
  "white": "【強調色：白字深邊】所有文字白色，關鍵字只靠放大與加粗強調，描邊深藍或深灰，不用第二個顏色。",
  "navy": "【強調色：深藍】淺色背景用：主色深藍 (#1B3F8B)，關鍵字與價格用品牌橘。"
};

/**
 * 第一輪對話的開場指令。之後的輪次只送業務的中文，靠 previous_response_id 延續。
 * mode: "full" 完整稿（含標題文字，像 ChatGPT 直接生一張廣告）／ "background" 只畫背景（之後用程式疊真實文字與產品）
 */
function buildChatInstruction(input) {
  const ratio = input.ratio || "9:16";
  const lines = [];
  const brand = resolveBrand(input);
  lines.push(`你是永豐餘消費品旗下品牌「${brand.name}」的資深電商視覺設計師，投放平台是 ${BRAND.platform}，受眾是 ${BRAND.audience}`);
  lines.push(`【品牌鐵律】這張圖的品牌是「${brand.name}」${brand.note ? "（" + brand.note + "）" : ""}。畫面中不可出現任何其他品牌的名稱或 logo，尤其不可把它畫成「橘子工坊」。沒有附 logo 圖就不要自己畫 logo，把左上角留空。`);
  lines.push("【數字鐵律】日期、價格、優惠、規格只能用下面業務給的欄位；參考圖、版型圖、產品圖上出現的日期與價格一律不可沿用。");
  lines.push(`請直接使用圖片生成工具產出一張 ${ratio} 的圖，不要先反問，不確定的地方自己做合理判斷。之後我會用中文請你修改，每次修改只動我提到的部分，其他完全保持。`);
  lines.push("");
  if (input.mode === "background") {
    lines.push("【這次要畫的是「背景底圖」】");
    lines.push("- 不要畫任何產品、包裝、文字、logo、價格。真實產品照與中文標題之後會由程式疊上去。");
    lines.push("- 背景是空舞台：不能有掛著的衣服、洗衣機、行李箱、家電等主體物件。小道具只放邊緣，限洗衣情境（毛巾、摺好的衣物、籐籃、小綠植、水珠、泡泡、橘子切片）。");
    const layout = LAYOUTS.find(l => l.key === input.layout) || LAYOUTS[0];
    lines.push("- 留白配置（英文原文）：" + layout.en);
    lines.push("- 單一柔和光源，方便之後合成去背產品。");
  } else {
    lines.push("【這次要畫的是「完整的廣告主視覺」】");
    lines.push(HOUSE_TEMPLATE);
    lines.push("");
    if (input.layoutImages && input.layoutImages.length) lines.push("【文字排法：照版型圖】文字位置、資訊帶樣式、強調色都照附上的版型圖，只換內容。");
    else {
      lines.push(TEXT_LAYOUTS[input.textLayout] || TEXT_LAYOUTS.auto);
      lines.push(ACCENTS[input.accent] || ACCENTS.auto);
    }
    lines.push("");
    lines.push("【圖上的文字，只能有這幾組，不可多加任何字】");
    const headline = input.headline || (input.hookCopy ? String(input.hookCopy).split(/\r?\n/)[0] : "");
    if (headline) lines.push("- 主標：「" + headline + "」");
    if (input.subline) lines.push("- 優惠：「" + input.subline + "」（放底部資訊帶左側）");
    if (input.price) lines.push("- 價格：「" + input.price + "」（放底部資訊帶右側，數字特大亮黃）");
    if (input.badge) lines.push("- 檔期：「" + input.badge + "」（放底部資訊帶最左，接「momo限定 ▸」）");
    if (input.note) lines.push("- 小字註記：「" + input.note + "」（最小字級，資訊帶最下緣）");
    const badges = (input.productInfo && input.productInfo.sellingPoints ? String(input.productInfo.sellingPoints).split(/[、,，]/) : []).map(t => t.trim()).filter(t => t && t.length <= 8).slice(0, 4);
    if (badges.length) lines.push("- 賣點徽章：" + badges.map(b => "「" + b + "」").join(""));
    if (!headline && !input.subline && !input.price && !input.badge) lines.push("- （沒有給文字，這張不要有任何文字）");
    if (input.hookCopy) {
      lines.push("");
      lines.push("【這次的文案，主視覺要扣住它】\n" + input.hookCopy);
      lines.push("先找出文案裡可以視覺化的字眼（例如「一球」就讓膠囊像球一樣飛進畫面），讓主視覺表現那個梗，不要只是把產品放中間。");
    }
    if (input.productImages && input.productImages.length) {
      lines.push("- 附上的真實產品去背圖：必須忠實重現形狀、配色、包裝版面與包裝上的文字，不可重新設計或改字。看不清的小字寧可留白，不要編字。");
    }
    if (input.logoImages && input.logoImages.length) lines.push("- 附上的 logo 圖：原樣放左上角，不可改顏色或變形。");
  }
  lines.push("");
  if (input.referenceImages && input.referenceImages.length) {
    lines.push("【參考圖的用法】只借它的配色、氛圍、光線、構圖密度。不要抄它的文字、logo、價格、他牌產品；若參考圖裡有他牌 logo 或名人角色，一律排除。");
  }
  if (input.layoutImages && input.layoutImages.length) {
    lines.push("【版型圖的用法】這是我們自家設計師的成品，請照它的排版來排：logo 位置、標題位置與字體處理、賣點徽章、產品位置、底部資訊帶的樣式，都照它。只把文字、產品、背景主題換成這次的內容。");
  }
  if (input.refAnalysis) {
    lines.push("【先前對參考圖的分析】" + input.refAnalysis.summaryZh + (input.refAnalysis.backgroundPaletteEn ? "　背景色：" + input.refAnalysis.backgroundPaletteEn : ""));
  }
  const productInfo = input.productInfo;
  if (input.product) lines.push("【商品】" + input.product + (productInfo ? "。賣點：" + productInfo.sellingPoints : ""));
  lines.push("【品牌事實，不可違反】" + brand.facts.slice(0, 3).join("；"));
  lines.push("【禁止】療效宣稱字眼、他牌 logo、真人臉部。");
  return lines.join("\n");
}

function extractResponsesError(code, body) {
  if (code === 401) throw new Error("API Key 無效或已被刪除。請重新執行「設定 OpenAI API Key」。");
  if (body.indexOf("must be verified") !== -1 || (code === 403 && body.indexOf("organization") !== -1)) {
    throw new Error("這個 OpenAI 組織尚未完成生圖功能所需的身份驗證。\n請到 platform.openai.com/settings/organization/general 完成 Organization Verification 後再試。");
  }
  if (code === 429 && body.indexOf("insufficient_quota") !== -1) throw new Error("OpenAI 帳戶額度不足。請到 platform.openai.com → Settings → Billing 儲值。");
  if (code === 429 || code >= 500) throw new Error("服務暫時忙線或已達速率上限，請稍候再試一次。");
  if (body.indexOf("safety") !== -1 || body.indexOf("moderation") !== -1) throw new Error("內容被安全系統擋下。請把可能敏感的字眼改掉再試。");
  throw new Error("API 錯誤（" + code + "）：\n" + body.substring(0, 300));
}

/**
 * 側邊欄呼叫：生圖對話。第一輪帶開場指令與圖片；之後只帶業務的中文與 previousResponseId。
 * input: { text, mode, ratio, quality, previousResponseId, productImages, referenceImages, layoutImages, logoImages,
 *          product, productInfo, headline, subline, price, badge, note, hookCopy, layout, textLayout, accent }
 * 回傳: { responseId, base64, text, revisedPrompt, model }
 */
function chatImage(input) {
  const apiKey = requireApiKey();
  const size = SIZE_MAP[input.ratio] || "1024x1024";
  const quality = ["low", "medium", "high"].indexOf(input.quality) !== -1 ? input.quality : IMAGE_QUALITY;
  const firstTurn = !input.previousResponseId;
  const userText = String(input.text || "").trim();
  if (!userText && firstTurn) throw new Error("請先用中文寫這張圖要長什麼樣。");

  const content = [];
  if (firstTurn) {
    content.push({ type: "input_text", text: buildChatInstruction(input) + "\n\n【業務的要求】\n" + (userText || "依上面的設定產出一張。") });
    (input.productImages || []).forEach((im, i) => {
      content.push({ type: "input_text", text: "【真實產品去背圖 " + (i + 1) + "】" });
      content.push({ type: "input_image", image_url: "data:" + (im.mime || "image/png") + ";base64," + im.base64, detail: "high" });
    });
    (input.logoImages || []).forEach((im, i) => {
      content.push({ type: "input_text", text: "【品牌 logo " + (i + 1) + "，原樣放左上角】" });
      content.push({ type: "input_image", image_url: "data:" + (im.mime || "image/png") + ";base64," + im.base64, detail: "high" });
    });
    (input.layoutImages || []).forEach((im, i) => {
      content.push({ type: "input_text", text: "【版型圖 " + (i + 1) + "，照它的排版】" });
      content.push({ type: "input_image", image_url: "data:" + (im.mime || "image/jpeg") + ";base64," + im.base64, detail: "high" });
    });
    (input.referenceImages || []).forEach((im, i) => {
      content.push({ type: "input_text", text: "【參考圖 " + (i + 1) + "，只借氛圍】" });
      content.push({ type: "input_image", image_url: "data:" + (im.mime || "image/jpeg") + ";base64," + im.base64, detail: "low" });
    });
  } else {
    content.push({ type: "input_text", text: userText + "\n（只改我提到的部分，其他保持與上一張一致。直接產出新圖。）" });
  }

  // 不指定 image model，讓 OpenAI 用目前最新的（實測已是 gpt-image-2）。
  // 舊的 input_fidelity 參數新模型不吃，不送；若日後有參數不被支援，下面的迴圈會自動拿掉重送。
  const tool = { type: "image_generation", size: size, quality: quality, output_format: "png" };

  const basePayload = {
    input: [{ role: "user", content: content }],
    tools: [tool],
    tool_choice: { type: "image_generation" },
    reasoning: { effort: "low" }   // 看圖寫指令不需要深思，省時間；模型不支援會自動拿掉
  };
  if (input.previousResponseId) basePayload.previous_response_id = input.previousResponseId;

  let lastError = "";
  for (let i = 0; i < CHAT_IMAGE_MODELS.length; i++) {
    const model = CHAT_IMAGE_MODELS[i];
    let payload = Object.assign({ model: model }, basePayload);
    for (let attempt = 0; attempt < 2; attempt++) {
      const res = UrlFetchApp.fetch("https://api.openai.com/v1/responses", {
        method: "post", contentType: "application/json",
        headers: { Authorization: "Bearer " + apiKey },
        payload: JSON.stringify(payload), muteHttpExceptions: true
      });
      const code = res.getResponseCode(), body = res.getContentText();
      if (code === 200) {
        const out = JSON.parse(body);
        const imgCall = (out.output || []).find(o => o.type === "image_generation_call");
        const msg = (out.output || []).find(o => o.type === "message");
        const msgText = msg && msg.content ? msg.content.filter(c => c.type === "output_text").map(c => c.text).join("\n") : "";
        if (!imgCall || !imgCall.result) {
          throw new Error("模型這次沒有產圖，只回了文字：\n" + (msgText || "（無）") + "\n請把要求講得更具體，或直接說「請直接產圖」。");
        }
        return { responseId: out.id, base64: imgCall.result, text: msgText, revisedPrompt: imgCall.revised_prompt || "", model: model };
      }
      // tool_choice 不被接受時，拿掉再試一次
      if (code === 400 && body.indexOf("tool_choice") !== -1 && payload.tool_choice) {
        payload = Object.assign({}, payload); delete payload.tool_choice; attempt--; continue;
      }
      // 「模型不支援 X 參數」→ 從頂層或工具設定拿掉再試（例如 gpt-image-2 不吃 input_fidelity、某些模型不吃 reasoning）
      if (code === 400) {
        const copy = JSON.parse(JSON.stringify(payload));
        if (stripUnsupportedParam(copy, body)) { payload = copy; attempt--; continue; }
      }
      if (code === 404 || (code === 400 && body.indexOf("model") !== -1 && (body.indexOf("does not exist") !== -1 || body.indexOf("not found") !== -1))) {
        lastError = "模型 " + model + " 無法使用"; break;
      }
      if (code === 400 && body.indexOf("previous_response_id") !== -1) {
        throw new Error("上一輪對話已失效（可能過期），請按「重新開始對話」。");
      }
      if ((code === 429 && body.indexOf("insufficient_quota") === -1) || code >= 500) {
        lastError = "模型 " + model + " 暫時無法回應（" + code + "）";
        if (attempt === 0) { Utilities.sleep(2500); continue; }
        break;
      }
      extractResponsesError(code, body);
    }
  }
  throw new Error("生圖失敗（" + lastError + "）。稍候再試，或把這則訊息貼給工程師。");
}

/** 側邊欄呼叫：貼網址時由伺服器端抓圖（瀏覽器跨網域抓不到），回傳 dataUrl */
function fetchImageFromUrl(url) {
  url = String(url || "").trim();
  if (!/^https?:\/\//i.test(url)) throw new Error("網址要以 http:// 或 https:// 開頭。");
  const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true, headers: { "User-Agent": "Mozilla/5.0" } });
  if (res.getResponseCode() !== 200) throw new Error("抓不到這個網址（" + res.getResponseCode() + "）。改用右鍵另存圖片再拖進來。");
  const blob = res.getBlob();
  const mime = (blob.getContentType() || "").split(";")[0];
  if (!/^image\/(png|jpeg|webp|gif)$/.test(mime)) throw new Error("這個網址不是圖片檔（" + mime + "）。請對圖片本身按右鍵「複製圖片網址」。");
  const bytes = blob.getBytes();
  if (bytes.length > 15 * 1024 * 1024) throw new Error("圖片超過 15MB，請換小一點的。");
  return { dataUrl: "data:" + mime + ";base64," + Utilities.base64Encode(bytes), mime: mime };
}

/** 側邊欄呼叫：把 base64 圖存進雲端硬碟輸出資料夾 */
function saveImageToDrive(input) {
  const folder = getFolderByName(DRIVE_FOLDERS.OUTPUT) || DriveApp.createFolder(DRIVE_FOLDERS.OUTPUT);
  const bytes = Utilities.base64Decode(input.base64);
  const blob = Utilities.newBlob(bytes, "image/png", (input.filename || "ai-image") + ".png");
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { url: file.getUrl() };
}

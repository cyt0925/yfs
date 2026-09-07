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
 *  - 產品：可從雲端硬碟資料夾勾選去背產品圖與參考成品，走 /v1/images/edits 一起送進模型。
 *  - 合成：側邊欄用 canvas 把 AI 底圖 + 真實去背產品 PNG + 中文標題 + 角標 + logo 疊成草稿，
 *          數字與 logo 保證正確，給設計師接手微調。
 *
 * 安裝：
 * 1. 擴充功能 → Apps Script
 * 2. 這份 Code.gs 貼進「Code.gs」（取代原本內容）
 * 3. 新增 HTML 檔「Sidebar」（不要打副檔名），貼上 Sidebar.html 的內容
 * 4. 存檔、回試算表重新整理
 * 5. 選單「AI 文案工具 → 設定 OpenAI API Key」
 * 6. 選單「AI 文案工具 → 建立雲端硬碟素材資料夾」，把去背產品 PNG、參考成品、logo 丟進對應資料夾
 * 7. （選用）選單「AI 文案工具 → 建立產品資料分頁」，填每個品類的賣點
 * 8. 選單「AI 文案工具 → 開啟 AI 文案側邊欄」
 */

// ============================================================
// 基本設定
// ============================================================

const API_KEY_PROP = "OPENAI_API_KEY";

// 依序嘗試，哪個模型可用就用哪個。三行文案 token 極少，直接用大模型，品質差很多。
const TEXT_MODEL_CANDIDATES = ["gpt-4.1", "gpt-4o", "gpt-4.1-mini"];
const IMAGE_MODEL = "gpt-image-1";
const IMAGE_QUALITY = "medium"; // low / medium / high，high 一張成本約 3～4 倍

const LABELS = { PRODUCT: "品類", AD_NAME: "廣告名稱", COPY: "文案", BUDGET: "預算分配" };
const PRODUCT_SHEET = "產品資料";

const DRIVE_FOLDERS = {
  PRODUCTS: "橘子工坊產品去背圖",   // 去背 PNG，一個檔一個產品
  REFERENCES: "橘子工坊參考成品",   // 設計師做過的成品，當風格參考
  ASSETS: "橘子工坊品牌素材",       // logo.png、mo點圖示等
  OUTPUT: "橘子工坊生圖"            // 產出存這裡
};

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

/** 視覺風格：固定英文描述，模型不能改 */
const VISUAL_STYLES = [
  { key: "清新自然光", en: "Style: bright natural daylight photography, soft diffused window light from the upper left, gentle shadows, realistic materials, airy and clean." },
  { key: "高彩度促銷", en: "Style: high-saturation e-commerce promotional graphic, bold flat color blocks, subtle radial glow behind the product zone, energetic but uncluttered." },
  { key: "情境敘事", en: "Style: lifestyle editorial photography of a real Taiwanese home (laundry corner, balcony, kitchen), warm and relatable, shallow depth of field, no people faces." },
  { key: "極簡棚拍", en: "Style: minimal studio product photography, seamless paper backdrop, single soft key light, one clean floor-to-wall horizon line, no props except one or two." },
  { key: "節慶時事", en: "Style: festive seasonal key visual, celebratory props (streamers, confetti, ribbons) kept to the edges, sparkle highlights, still clean in the center." },
  { key: "3D 立體插畫", en: "Style: soft 3D rendered illustration, rounded shapes, matte pastel materials, gentle ambient occlusion, toy-like friendly look." }
];

/**
 * 版面：同時用在 Prompt 文字與 canvas 合成。
 * zones 為畫面比例 (0~1)：headline 標題區、products 產品區、badge 角標位置。
 */
const LAYOUTS = [
  {
    key: "上標題．下產品",
    en: "Layout: the TOP 35% of the frame must be a calm, low-detail area (plain color or very soft gradient) reserved for a headline. The BOTTOM 40% is a simple flat surface (table, floor, or platform) where products will stand. Decorative elements only in the upper-left and right edges, never in the center.",
    zones: { headline: { x: 0.06, y: 0.05, w: 0.88, h: 0.28 }, products: { x: 0.05, y: 0.55, w: 0.90, h: 0.38 }, badge: "top-right", logo: "top-left" }
  },
  {
    key: "左產品．右標題",
    en: "Layout: the RIGHT 45% of the frame must be a calm, low-detail area (plain color or very soft gradient) reserved for a headline. The LEFT 50% is a simple flat surface where products will stand. Decorative elements only along the bottom edge and top-left corner.",
    zones: { headline: { x: 0.52, y: 0.12, w: 0.44, h: 0.40 }, products: { x: 0.03, y: 0.30, w: 0.48, h: 0.60 }, badge: "top-right", logo: "top-left" }
  },
  {
    key: "中央產品．上下標題",
    en: "Layout: the TOP 25% and BOTTOM 15% must be calm, low-detail bands reserved for text. The CENTER 55% is an open stage with a soft spotlight where products will stand. Decorative elements only at the far left and right edges.",
    zones: { headline: { x: 0.08, y: 0.04, w: 0.84, h: 0.20 }, products: { x: 0.10, y: 0.30, w: 0.80, h: 0.52 }, badge: "top-right", logo: "top-left" }
  }
];

const RATIOS = ["9:16", "1:1", "16:9", "4:5"];
const SIZE_MAP = { "9:16": "1024x1536", "4:5": "1024x1536", "16:9": "1536x1024", "1:1": "1024x1024" };

/** 所有生圖共用的硬規則 */
const IMAGE_HARD_RULES = [
  "This is a BACKGROUND PLATE for an e-commerce ad. Real product photos and Chinese headline text will be composited on top later.",
  "ABSOLUTELY NO text, letters, numbers, logos, watermarks, price tags, labels or signage anywhere in the image.",
  "Do NOT invent or draw any product bottle, box, pouch or packaging unless product reference images are provided; if they are provided, reproduce them faithfully and place them ONLY inside the product zone described in the layout.",
  "Single soft light direction, no harsh multi-source shadows, so cut-out products composite naturally.",
  "Keep the reserved text areas genuinely empty and low-contrast: no busy patterns, no small objects, no strong highlights there.",
  "Photorealistic quality, sharp focus, no blur on the main surfaces, no people faces, no hands."
].join(" ");

// ============================================================
// 選單
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("AI 文案工具")
    .addItem("開啟 AI 文案側邊欄", "showSidebar")
    .addSeparator()
    .addItem("設定 OpenAI API Key", "setApiKey")
    .addItem("建立雲端硬碟素材資料夾", "setupDriveFolders")
    .addItem("建立產品資料分頁", "setupProductSheet")
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
    "雲端硬碟資料夾已就位。\n\n" +
    "・" + DRIVE_FOLDERS.PRODUCTS + "：放去背產品 PNG，一個檔一個產品，檔名就是顯示名稱\n" +
    "・" + DRIVE_FOLDERS.REFERENCES + "：放設計師做過的成品，當風格參考\n" +
    "・" + DRIVE_FOLDERS.ASSETS + "：放 logo.png（檔名含 logo 即可）\n" +
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
    productInfo: getProductInfo(product),
    hasApiKey: !!PropertiesService.getScriptProperties().getProperty(API_KEY_PROP)
  };
}

function getFormOptions() {
  return {
    tones: TONES,
    palettes: PALETTES,
    visualStyles: VISUAL_STYLES.map(s => s.key),
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
    options.payload = JSON.stringify(Object.assign({ model: model }, payload));

    for (let attempt = 0; attempt < 2; attempt++) {
      const res = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", options);
      const code = res.getResponseCode();
      const body = res.getContentText();

      if (code === 200) {
        const json = JSON.parse(body);
        return JSON.parse(json.choices[0].message.content);
      }
      if (code === 401) throw new Error("API Key 無效或已被刪除。請重新執行「設定 OpenAI API Key」。");
      if (code === 404 || (code === 400 && body.indexOf("model") !== -1 && body.indexOf("does not exist") !== -1)) {
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

  return `【品牌】${BRAND.name}，投放平台：${BRAND.platform}
【受眾】${BRAND.audience}
【語氣】
${BRAND.voice.map(v => "- " + v).join("\n")}
【品牌事實（只能用這些，不可自行加成分或功效）】
${BRAND.facts.map(v => "- " + v).join("\n")}
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

  return `你是「${BRAND.name}」的社群廣告文案手，專寫 ${BRAND.platform} 的短文案。你只負責「前段情境鉤子」，後段的價格、優惠、贈品由業務另外填寫。

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
- 第一行是鉤子，一定要有畫面或情緒，讓人想往下看；第二行帶到 ${BRAND.name} 與商品；第三行收尾（結果畫面、金句或回收梗）。
- 三個版本的切角要真的不同（不同情境、不同人物視角、不同情緒），不可以只是換同義詞。
- 絕對不可出現任何價格、折扣、贈品、mo點、日期等具體數字或優惠條件。
- 每行最多 2 個 emoji，可以完全不用。

【生圖場景描述規則（scene 欄位）】
- 英文，1～2 句，只描述「背景場景與道具」：地點、材質、季節感、1～3 個小道具。
- 不要寫任何顏色（顏色由設計師另外指定）、不要寫版面配置、不要寫文字或產品、不要寫光線。
- 場景要跟三則文案的共同情境呼應，例如文案講健身汗臭，場景就是明亮的家庭洗衣角落加一條運動毛巾。

只回傳 JSON。`;
}

function buildImagePrompt(input, scene) {
  const palette = PALETTES.find(p => p.key === input.palette) || PALETTES[0];
  const paletteEn = palette.key === "自訂"
    ? "Dominant color palette: " + (input.customPalette || "designer's choice") + "."
    : palette.en;
  const style = VISUAL_STYLES.find(s => s.key === input.visualStyle) || VISUAL_STYLES[0];
  const layout = LAYOUTS.find(l => l.key === input.layout) || LAYOUTS[0];
  const ratio = input.ratio || "9:16";

  return [
    paletteEn,
    layout.en,
    style.en,
    "Scene: " + (scene || "a clean home laundry corner with a folded white towel and a small green plant."),
    "Aspect ratio " + ratio + ".",
    IMAGE_HARD_RULES
  ].join("\n\n");
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
  const out = callChatJson(buildCopyPrompt(input), "ad_copy_draft", schema);
  out.imagePrompt = buildImagePrompt(input, out.scene);
  return out;
}

/** 側邊欄呼叫：只重組 Prompt（換主色調／版面／風格時不用重跑文案） */
function rebuildImagePrompt(input) {
  return { imagePrompt: buildImagePrompt(input, input.scene) };
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

function appendBytes(dst, src) {
  const CHUNK = 20000;
  for (let i = 0; i < src.length; i += CHUNK) {
    Array.prototype.push.apply(dst, src.slice(i, i + CHUNK));
  }
}

/** 手動組 multipart/form-data，因為 UrlFetchApp 的物件 payload 不支援同名多檔 image[] */
function buildMultipart(fields, files) {
  const boundary = "----OrangeHouseBoundary" + Utilities.getUuid().replace(/-/g, "");
  const enc = s => Utilities.newBlob(s).getBytes();
  const out = [];
  Object.keys(fields).forEach(k => {
    appendBytes(out, enc("--" + boundary + "\r\nContent-Disposition: form-data; name=\"" + k + "\"\r\n\r\n" + fields[k] + "\r\n"));
  });
  files.forEach(f => {
    appendBytes(out, enc("--" + boundary + "\r\nContent-Disposition: form-data; name=\"" + f.field + "\"; filename=\"" + f.name + "\"\r\nContent-Type: " + f.mime + "\r\n\r\n"));
    appendBytes(out, f.bytes);
    appendBytes(out, enc("\r\n"));
  });
  appendBytes(out, enc("--" + boundary + "--\r\n"));
  return { contentType: "multipart/form-data; boundary=" + boundary, bytes: out };
}

function handleImageError(code, body) {
  if (code === 401) throw new Error("API Key 無效或已被刪除。請重新執行「設定 OpenAI API Key」。");
  if (body.indexOf("must be verified") !== -1 || (code === 403 && body.indexOf("organization") !== -1)) {
    throw new Error("這個 OpenAI 組織尚未完成生圖功能所需的身份驗證。\n請到 platform.openai.com/settings/organization/general 完成 Organization Verification 後再試。");
  }
  if (code === 429 && body.indexOf("insufficient_quota") !== -1) {
    throw new Error("OpenAI 帳戶額度不足。請到 platform.openai.com → Settings → Billing 儲值。");
  }
  if (code === 429 || code >= 500) throw new Error("生圖服務暫時忙線或已達速率上限，請稍候再試一次。");
  if (body.indexOf("safety") !== -1 || body.indexOf("moderation") !== -1) {
    throw new Error("Prompt 被安全系統擋下。請把 Prompt 裡可能敏感的字眼改掉再試。");
  }
  throw new Error("生圖失敗（" + code + "）：\n" + body.substring(0, 300));
}

/**
 * 側邊欄呼叫：產生圖片，回傳 base64。
 * input.productFileIds / input.referenceFileIds 有東西 → /v1/images/edits 帶參考圖
 * 否則 → /v1/images/generations 純文字
 */
function generateImage(input) {
  const apiKey = requireApiKey();
  if (!input.imagePrompt) throw new Error("沒有生圖 Prompt 可以使用，請先生成文案。");
  const size = SIZE_MAP[input.ratio] || "1024x1024";

  const productIds = input.productFileIds || [];
  const referenceIds = input.referenceFileIds || [];
  const useEdits = productIds.length + referenceIds.length > 0;

  let prompt = input.imagePrompt;
  if (useEdits) {
    const notes = [];
    if (productIds.length) {
      notes.push("The first " + productIds.length + " attached image(s) are the REAL PRODUCTS (cut-out, transparent background). Reproduce their shape, colors and label design as faithfully as possible, do not redesign them, and place them standing naturally ONLY inside the product zone described in the layout, with correct perspective and a soft contact shadow.");
    }
    if (referenceIds.length) {
      notes.push("The last " + referenceIds.length + " attached image(s) are previous finished ads from our designer. Use them ONLY as a reference for overall mood, composition density and lighting. Do NOT copy their text, numbers, logos or badges.");
    }
    prompt = notes.join(" ") + "\n\n" + prompt;
  }

  let res;
  if (useEdits) {
    const files = [];
    productIds.concat(referenceIds).forEach((id, i) => {
      const f = DriveApp.getFileById(id);
      const blob = f.getBlob();
      const mime = blob.getContentType();
      const ext = mime === "image/jpeg" ? "jpg" : (mime === "image/webp" ? "webp" : "png");
      // 檔名只用 ASCII，避免中文檔名進 multipart 標頭
      files.push({ field: "image[]", name: "image" + (i + 1) + "." + ext, mime: mime, bytes: blob.getBytes() });
    });
    const mp = buildMultipart({ model: IMAGE_MODEL, prompt: prompt, size: size, quality: IMAGE_QUALITY, n: "1" }, files);
    res = UrlFetchApp.fetch("https://api.openai.com/v1/images/edits", {
      method: "post",
      contentType: mp.contentType,
      headers: { Authorization: "Bearer " + apiKey },
      payload: mp.bytes,
      muteHttpExceptions: true
    });
  } else {
    res = UrlFetchApp.fetch("https://api.openai.com/v1/images/generations", {
      method: "post",
      contentType: "application/json",
      headers: { Authorization: "Bearer " + apiKey },
      payload: JSON.stringify({ model: IMAGE_MODEL, prompt: prompt, size: size, quality: IMAGE_QUALITY, n: 1 }),
      muteHttpExceptions: true
    });
  }

  const code = res.getResponseCode();
  const body = res.getContentText();
  if (code !== 200) handleImageError(code, body);

  const json = JSON.parse(body);
  return { base64: json.data[0].b64_json, mode: useEdits ? "edits" : "generations", finalPrompt: prompt };
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

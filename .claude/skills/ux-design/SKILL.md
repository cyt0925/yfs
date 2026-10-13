---
name: ux-design
description: 做新畫面、加功能到既有畫面、或改版之前一定要讀：畫面要照同事實際作業的流程排，不照程式模組切。頁籤、清單、按鈕、詳細資料要放哪，照這份的原則和檢查清單決定，做完先給 Jerry 看示意圖再動手。
---

# 畫面設計規矩（UX）

## 為什麼有這份

2026-10-13 Jerry 看完蝦皮特選訂單管理系統的六個頁籤（訂單匯入與總覽、竹運出貨、竹運採購進貨、直送訂單、缺貨維護與通知、採購單產出）後的原話：

> 「真的要這樣分為六步驟嗎？分為六個介面嗎？ux 的設計簡直為 0 分，一點都不直覺，難道沒有任何可以融合或是簡略的頁面嗎？」

錯在哪：**頁籤是照程式的模組切的**（ship.py 一個頁籤、purchase.py 一個頁籤、short.py 一個頁籤），不是照同事每天怎麼做事切的。
結果同事要自己記「這個檔要去哪個頁籤傳」「這張單現在卡在哪一步、要去哪個頁籤處理」，每一頁都有自己的上傳區、自己的日期篩選、自己的清單。

## 原則（附出處）

1. **畫面繞著「東西」排，不繞著「功能」排。** 同事處理的東西是「一張蝦皮訂單」，產 EIP 採購單、填缺貨、通知都是對這張單做的動作，
   動作要跟著單走，不要各自變成一個頁籤。（Object-Oriented UX：navigation 以名詞為主、動詞附在物件上，
   [uxmastery](https://uxmastery.com/object-focused-vs-task-focused/)、[LogRocket](https://blog.logrocket.com/ux-design/object-oriented-ux-ooux/)）
2. **用「狀態／下一步」當入口，不用模組當入口。** 一進來就看到「待產採購單 3 張、待回填 EIP 採購單號 6 張、待通知蝦皮 1 張」，點一格清單就切過去，
   那一格只出現那一步要按的按鈕。（NN/g 第 1 條「系統狀態要看得見」、第 6 條「讓人用看的，不要用記的」，
   [10 Usability Heuristics](https://www.nngroup.com/articles/ten-usability-heuristics/)；依狀態限制可做的動作，
   [Automation Anywhere work items](https://docs.automationanywhere.com/bundle/enterprise-v11.3/page/enterprise/topics/control-room/workload/work-item-status-actions.html)）
3. **頁籤不拿來放有先後順序的步驟，也不拿來當頁面導覽。** 頁籤適合「彼此獨立、一次只看一組」的內容；
   有順序的流程、需要對照的資料放在同一頁。使用者多半只看預設那個頁籤，其他頁籤的東西會被忽略。
   （[NN/g Tabs, Used Right](https://www.nngroup.com/articles/tabs-used-right/)、[NN/g In-Page Links](https://www.nngroup.com/articles/in-page-links-content-navigation/)、
   [GOV.GR tabs](https://guide.services.gov.gr/docs/components/tabs)、[NICE tabs](https://design-system.nice.org.uk/components/tabs/)）
4. **常用的放前面，偶爾用的收起來。** 價格本、庫存資料、倉資料、收件資料、效期設定這些「參考資料」平常不用看，
   收到「資料與設定」；每天要做的事留在主畫面。（Progressive disclosure，[NN/g](https://www.nngroup.com/articles/progressive-disclosure/)）
5. **看一筆、改一筆用側邊面板，清單不要被蓋住。** 點一張單，右邊打開那張單的全部（品項、採購單、缺貨、版本、歷程），
   該改的在那裡改；不要另開頁籤、不要用蓋住整個畫面的大視窗。（[NN/g Data Tables](https://www.nngroup.com/articles/data-tables/)：
   modal 會擋住旁邊的資料，non-modal 側邊面板可以一邊看表一邊改；[SAP detail panel](https://www.sap.com/design-system/hcm/components/panel/detail-panel-fcl/usage)、
   [saasui list-detail](https://www.saasui.design/blog/saas-list-detail-master-detail-ux-patterns)）
6. **一個入口收所有檔案。** 系統本來就能看表頭認檔案種類，就不要讓人先決定「去哪一頁傳」。拖進來，系統認出是什麼，再問該問的（例如選履約方式）。
7. **流程可以跳、可以回頭。** 複雜工作系統的使用者會跳步驟、回頭改，不要做成只能一路往下的精靈。
   （[NN/g 8 Design Guidelines for Complex Applications](https://www.nngroup.com/articles/complex-application-design/)）

## 動手前的檢查清單

做任何一個新畫面或改版前，先寫下來回答：

1. 同事每天實際的作業順序是什麼？（寫成一行：上傳 → … → 通知）
2. 他們處理的主要「東西」是什麼？（訂單、PO、拆單表…）每個動作是對哪個東西做的？
3. 一進來最想知道什麼？（通常是「還有什麼沒做」，要能一眼看到數字）
4. 哪些是每天用的、哪些是偶爾設定一次的？偶爾用的收起來。
5. 有沒有兩個畫面其實在處理同一批東西？有就合併。
6. 每多一個頁籤、一個上傳區、一個日期篩選，都要問：可以不要嗎？
7. 筆電寬度（1366）看得完整嗎？按鈕會不會被切掉？
8. **先做示意圖給 Jerry 看、確認方向，再寫程式。** 大改分幾次上，每次做完截圖。

## 已經有的好做法（照用）

- 首頁（YFS 訂單系統）「今日事項」：一進來就列出要處理的事，點了直接過去。
- 酷澎訂單管理系統：一張 PO 一列，點開看明細；酷澎修改用顏色標出來、數字在最上面。
- 蝦皮特選訂單管理系統（2026-10-13 改版後）：一個上傳入口、作業流程格子（每格寫幾張要做）、點一張單開右側面板（該單的全部在這裡看、在這裡改）。
  之後新的訂單系統（例如 PChome）照這個樣子做。

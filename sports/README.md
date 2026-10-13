# 今日運彩總覽

一頁看完今日賽事、盤口、傷兵、火燙／冰冷球員、球隊近況、相關新聞與規則式建議。
純靜態網站，零主機費用，資料由 GitHub Actions 每小時自動更新。

## 架構

```
sports/
├── index.html / style.css / app.js   網頁（讀 data/today.json）
├── data/today.json                   排程產生的資料（自動 commit）
├── data/sample.json                  本機測試用的範例資料
└── pipeline/
    ├── fetch.py                      抓資料 + 規則式建議
    ├── names_zh.py                   球隊中文名、狀態中文
    └── make_sample.py                產生範例資料
.github/workflows/sports-data.yml     每小時排程
```

## 資料來源（全部免費、不用 key）

| 資料 | 來源 |
|---|---|
| 賽程、比分、盤口、新聞、排名 | ESPN 公開接口 |
| 傷兵與出賽狀態 | ESPN 各聯盟傷兵名單 |
| 球員近 5 場 vs 季平均（火燙／冰冷） | ESPN 球員比賽紀錄 |
| 球隊近 10 場 | ESPN 球隊賽程 |

盤口是美國莊家（ESPN BET）的參考值，不是台灣運彩。台灣運彩盤口是下一版要接的東西。

## 本機測試

```bash
python3 sports/pipeline/make_sample.py          # 產生 data/sample.json
cd sports && python3 -m http.server 8080        # 開 http://localhost:8080/?data=sample
python3 sports/pipeline/fetch.py                # 真實抓一次（需要能連 ESPN）
```

## 建議怎麼算（規則式，不是 AI 預測）

1. 以美國莊家賠率去掉抽水後的勝率為基準。
2. 看好的一方若有主力缺陣（球隊得分／籃板／助攻王在缺陣名單）扣分；對手主力缺陣加分。
3. 近 10 場勝率 7 成以上加分、3 成以下扣分；對手 3 連勝以上扣分。
4. 主力近 5 場比季平均高 15% 以上算「火燙」加分，低 15% 以上算「冰冷」扣分。
5. 調整後勝率達門檻才給「傾向」，否則寫「觀望」。
6. 大小分：兩隊近 5 場平均總分與盤口差距超過門檻才給方向。

免責：這只是整理資料的工具，不保證任何結果。

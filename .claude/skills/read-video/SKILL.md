---
name: read-video
description: 讀影片。使用者給影片檔（上傳、路徑或網址）要你看內容、整理成 SOP 步驟、擷取截圖、聽語音做逐字稿時用。先用 scripts/read_video.py 把影片拆成關鍵畫面＋縮圖總覽＋逐字稿，再用 Read 看圖、看 summary.md。
---

# 讀影片

Claude 本身看不了影片檔，但看得懂圖片和文字。這個 skill 把影片拆成兩樣東西：
**關鍵畫面**（用 Read 看圖）和**語音逐字稿**（文字），再按時間對齊成一份 `summary.md`。

## 步驟

1. **拿到影片檔的路徑。** 使用者上傳的檔案會出現在對話附的路徑；網址用 `curl -L -o` 下載到
   scratchpad（Google Drive、YouTube 這類網址多半被網路白名單擋住，擋了就直說，請使用者上傳檔案）。
   不要把影片 commit 進 repo。
2. **跑腳本：**
   ```bash
   python3 .claude/skills/read-video/scripts/read_video.py 影片.mp4
   ```
   輸出在 `./.video-cache/<影片檔名>/`（已在 .gitignore）。常用參數：
   - `--scene 0.15`：螢幕錄影（畫面變化小）把場景門檻調低，才抓得到換頁
   - `--interval 3`：每 3 秒固定取樣一張（預設依長度 5–30 秒）
   - `--max-frames 100`：最多抽幾張（預設 60）
   - `--lang auto`：語言自動偵測（預設 zh）；`--model medium` 聽得更準但慢
   - `--no-audio`：只要畫面
3. **先讀 `summary.md`**，再 **Read `sheets/sheet_01.jpg`…**（一張 12 格縮圖，格子下方是編號與時間）。
   這樣一兩張圖就能掌握整支影片在做什麼。
4. **要看細節才 Read `frames/` 裡的單張原尺寸圖**（檔名含編號與時間，和縮圖對得上）。
   想看某個特定時間點，直接再抽一張：
   ```bash
   ffmpeg -y -ss 00:01:23 -i 影片.mp4 -frames:v 1 -q:v 3 /tmp/claude-0/.../scratchpad/t.jpg
   ```
5. 逐字稿在 `transcript.txt`（含時間），`summary.md` 的時間軸已把每張畫面和那段話對好。

## 限制，要老實講

- **語音辨識要能連 huggingface.co 下載 Whisper 模型。** 這個雲端環境的網路白名單目前擋住它，
  腳本會在 `transcript.txt` 寫明原因，畫面部分照常產出。要開通：環境設定 → Network access →
  Allowed domains 加 `huggingface.co`（文件：<https://code.claude.com/docs/en/cloud-environments#network-access>）。
  模型抓過一次會快取在 `~/.cache/huggingface`。
- 抽圖是取樣，不是逐格看。動作很快（例如點一下就消失的選單）可能落在兩張之間，
  懷疑漏掉就用 `--interval` 縮小間隔或手動抽特定秒數。
- 畫面裡的小字要看清楚請 Read 原尺寸圖，不要只看縮圖。
- 沒有 OCR 工具；畫面上的文字是靠 Read 圖片直接看。

## 和 SOP 網站搭配

影片通常是同事錄的操作示範。整理成 SOP 時：
- 每個「換頁／開新視窗／按下送出」的畫面就是一個步驟，截圖從 `frames/` 挑，
  轉成 webp 放進 `img/`（`ffmpeg -i 來源.jpg -q:v 80 img/xxx.webp`），寬度與現有截圖一致（約 1280）。
- 步驟文字以逐字稿為底，改成祈使句、去掉口語。
- 內容更新走 Google 試算表（`SOPs`、`Steps` 分頁），不是改 `index.html`——見 CLAUDE.md 紅線。

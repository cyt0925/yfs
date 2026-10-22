#!/usr/bin/env python3
"""把影片拆成 Claude 讀得懂的東西：關鍵畫面、縮圖總覽、語音逐字稿、對齊後的摘要。

用法：
    python3 read_video.py 影片.mp4 [-o 輸出資料夾] [--interval 秒] [--scene 0.3]
                          [--max-frames 60] [--width 1280] [--no-audio]
                          [--model small] [--lang zh]

輸出（預設在 ./.video-cache/<影片檔名>/）：
    info.json         ffprobe 的長度、解析度、有無聲音
    sheets/sheet_NN.jpg  縮圖總覽（每張 12 格，格子下方印時間）──先讀這個
    frames/NNNN_HHhMMmSSs.jpg  關鍵畫面原尺寸（最寬 --width）──要看細節再讀
    audio.wav         16kHz 單聲道（有聲音才有）
    transcript.txt / transcript.json  語音逐字稿（faster-whisper；模型抓不到會寫原因）
    summary.md        畫面與逐字稿按時間對齊的總表──最後讀這個

只靠 ffmpeg + Pillow 就能抽畫面；語音辨識需要 faster-whisper 和 huggingface.co 能連得到。
"""
import argparse
import json
import math
import os
import re
import shutil
import subprocess
import sys
from pathlib import Path


def run(cmd, **kw):
    return subprocess.run(cmd, capture_output=True, text=True, **kw)


def hms(sec: float) -> str:
    sec = max(0, int(round(sec)))
    h, rem = divmod(sec, 3600)
    m, s = divmod(rem, 60)
    return f"{h:02d}:{m:02d}:{s:02d}"


def fname_time(sec: float) -> str:
    sec = max(0, int(round(sec)))
    h, rem = divmod(sec, 3600)
    m, s = divmod(rem, 60)
    return f"{h:02d}h{m:02d}m{s:02d}s"


def probe(video: Path) -> dict:
    r = run(["ffprobe", "-v", "error", "-show_entries",
             "format=duration,size:stream=codec_type,codec_name,width,height,r_frame_rate",
             "-of", "json", str(video)])
    if r.returncode != 0:
        sys.exit(f"ffprobe 失敗：{r.stderr.strip()}")
    d = json.loads(r.stdout)
    info = {"file": str(video), "duration": float(d["format"].get("duration", 0) or 0),
            "size_bytes": int(d["format"].get("size", 0) or 0),
            "width": None, "height": None, "fps": None, "has_audio": False, "video_codec": None}
    for s in d.get("streams", []):
        if s["codec_type"] == "video" and info["width"] is None:
            info["width"], info["height"] = s.get("width"), s.get("height")
            info["video_codec"] = s.get("codec_name")
            try:
                num, den = s.get("r_frame_rate", "0/1").split("/")
                info["fps"] = round(float(num) / float(den), 3) if float(den) else None
            except ValueError:
                pass
        elif s["codec_type"] == "audio":
            info["has_audio"] = True
    if info["width"] is None:
        sys.exit("這個檔案沒有影像軌。")
    return info


def scene_times(video: Path, threshold: float) -> list[float]:
    """用 ffmpeg 的 scene 偵測找畫面切換點，回傳秒數。"""
    r = run(["ffmpeg", "-hide_banner", "-nostats", "-i", str(video),
             "-vf", f"select='gt(scene,{threshold})',showinfo",
             "-vsync", "vfr", "-f", "null", "-"])
    return [float(t) for t in re.findall(r"pts_time:([0-9.]+)", r.stderr)]


def pick_times(duration: float, scenes: list[float], interval: float, max_frames: int) -> list[float]:
    """場景切換點 + 固定間隔取樣，去掉 1 秒內的重複，超過上限就平均抽稀。"""
    times = {0.2}
    for t in scenes:
        times.add(round(t, 2))
    t = interval
    while t < duration:
        times.add(round(t, 2))
        t += interval
    if duration > 1:
        times.add(round(duration - 0.5, 2))
    ordered = sorted(times)
    merged = []
    for t in ordered:
        if not merged or t - merged[-1] >= 1.0:
            merged.append(t)
    if len(merged) > max_frames:
        step = (len(merged) - 1) / (max_frames - 1)
        merged = [merged[int(i * step)] for i in range(max_frames - 1)] + [merged[-1]]
    return merged


def extract_frames(video: Path, times: list[float], out: Path, width: int) -> list[dict]:
    out.mkdir(parents=True, exist_ok=True)
    frames = []
    for i, t in enumerate(times, 1):
        name = f"{i:04d}_{fname_time(t)}.jpg"
        path = out / name
        r = run(["ffmpeg", "-hide_banner", "-loglevel", "error", "-y", "-ss", f"{t:.3f}",
                 "-i", str(video), "-frames:v", "1",
                 "-vf", f"scale='min({width},iw)':-2", "-q:v", "3", str(path)])
        if r.returncode == 0 and path.exists():
            frames.append({"index": i, "time": t, "hms": hms(t), "file": str(path)})
        else:
            print(f"  抽不到 {hms(t)}：{r.stderr.strip()[:200]}", file=sys.stderr)
    return frames


def contact_sheets(frames: list[dict], out: Path, cols=4, rows=3, cell_w=400) -> list[str]:
    """用 Pillow 拼縮圖總覽，每格下方印編號與時間。"""
    from PIL import Image, ImageDraw, ImageFont
    out.mkdir(parents=True, exist_ok=True)
    font = None
    for cand in ["/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc",
                 "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"]:
        if os.path.exists(cand):
            font = ImageFont.truetype(cand, 22)
            break
    if font is None:
        font = ImageFont.load_default(size=22)
    per = cols * rows
    label_h = 34
    sheets = []
    for n in range(math.ceil(len(frames) / per)):
        batch = frames[n * per:(n + 1) * per]
        thumbs = []
        for f in batch:
            im = Image.open(f["file"]).convert("RGB")
            ratio = cell_w / im.width
            thumbs.append((f, im.resize((cell_w, max(1, int(im.height * ratio))))))
        cell_h = max(t.height for _, t in thumbs) + label_h
        used_rows = math.ceil(len(batch) / cols)
        sheet = Image.new("RGB", (cols * cell_w + (cols + 1) * 6, used_rows * cell_h + (used_rows + 1) * 6), "#222")
        draw = ImageDraw.Draw(sheet)
        for k, (f, th) in enumerate(thumbs):
            x = 6 + (k % cols) * (cell_w + 6)
            y = 6 + (k // cols) * (cell_h + 6)
            sheet.paste(th, (x, y))
            # 標籤放在縮圖下方的獨立色帶，不蓋到畫面
            draw.rectangle([x, y + th.height, x + cell_w, y + th.height + label_h], fill="#000000")
            draw.text((x + 8, y + th.height + 5), f"#{f['index']}  {f['hms']}", fill="#ffff66", font=font)
        path = out / f"sheet_{n + 1:02d}.jpg"
        sheet.save(path, quality=85)
        sheets.append(str(path))
    return sheets


def extract_audio(video: Path, out: Path) -> bool:
    r = run(["ffmpeg", "-hide_banner", "-loglevel", "error", "-y", "-i", str(video),
             "-vn", "-ac", "1", "-ar", "16000", "-c:a", "pcm_s16le", str(out)])
    return r.returncode == 0 and out.exists()


def ensure_faster_whisper() -> bool:
    try:
        import faster_whisper  # noqa: F401
        return True
    except ImportError:
        print("  faster-whisper 沒裝，從 PyPI 安裝中（約 1 分鐘）…", file=sys.stderr)
        r = run([sys.executable, "-m", "pip", "install", "--quiet", "faster-whisper"])
        if r.returncode != 0:
            print(f"  安裝失敗：{r.stderr.strip()[-300:]}", file=sys.stderr)
            return False
        try:
            import faster_whisper  # noqa: F401
            return True
        except ImportError:
            return False


def transcribe(wav: Path, model_name: str, lang: str | None) -> tuple[list[dict], str | None]:
    """回傳 (segments, 錯誤訊息)。segments 是 {start, end, text}。"""
    if not ensure_faster_whisper():
        return [], "faster-whisper 裝不起來（需要 PyPI 可連線）。"
    from faster_whisper import WhisperModel
    try:
        model = WhisperModel(model_name, device="cpu", compute_type="int8")
    except Exception as e:  # 模型下載被擋是最常見的原因
        msg = str(e).splitlines()[-1][:300] if str(e) else repr(e)
        return [], ("Whisper 模型下載失敗：" + msg +
                    "\n模型放在 huggingface.co，這個環境的網路白名單要允許 huggingface.co "
                    "（Claude Code 雲端環境設定 → Network access → Allowed domains）。")
    kwargs = {"beam_size": 5, "vad_filter": True}
    if lang and lang != "auto":
        kwargs["language"] = lang
    segments, info = model.transcribe(str(wav), **kwargs)
    out = []
    for s in segments:
        text = s.text.strip()
        if text:
            out.append({"start": round(s.start, 2), "end": round(s.end, 2), "text": text})
    print(f"  語言：{info.language}（信心 {info.language_probability:.2f}），{len(out)} 段", file=sys.stderr)
    return out, None


def write_summary(out_dir: Path, info: dict, frames: list[dict], sheets: list[str],
                  segs: list[dict], asr_error: str | None, video: Path):
    lines = [f"# 影片解讀：{video.name}", ""]
    lines.append(f"- 長度 {hms(info['duration'])}，{info['width']}×{info['height']}，"
                 f"{info['fps']} fps，{'有' if info['has_audio'] else '無'}聲音，"
                 f"{info['size_bytes'] / 1e6:.1f} MB")
    lines.append(f"- 關鍵畫面 {len(frames)} 張，縮圖總覽 {len(sheets)} 張")
    if asr_error:
        lines.append(f"- ⚠️ 逐字稿沒做出來：{asr_error.splitlines()[0]}")
    elif segs:
        lines.append(f"- 逐字稿 {len(segs)} 段（見 transcript.txt）")
    elif info["has_audio"]:
        lines.append("- 有聲音軌但辨識不出任何語句（可能只是背景音）")
    lines += ["", "## 怎麼讀", "",
              "1. 先 Read `sheets/sheet_01.jpg`…看全貌（每格下方是編號與時間）。",
              "2. 想看清楚某一格，Read `frames/` 裡同編號的原尺寸圖。",
              "3. 下面的時間軸把每張畫面和那段時間說的話對在一起。", ""]
    lines += ["## 時間軸", ""]
    for i, f in enumerate(frames):
        t0 = f["time"]
        t1 = frames[i + 1]["time"] if i + 1 < len(frames) else info["duration"] + 1
        rel = Path(f["file"]).relative_to(out_dir)
        lines.append(f"### #{f['index']} {f['hms']} — `{rel}`")
        said = [s for s in segs if t0 <= s["start"] < t1 or (s["start"] < t0 < s["end"])]
        if said:
            for s in said:
                lines.append(f"- [{hms(s['start'])}] {s['text']}")
        elif segs:
            lines.append("- （這段沒有說話）")
        lines.append("")
    (out_dir / "summary.md").write_text("\n".join(lines), encoding="utf-8")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("video")
    ap.add_argument("-o", "--out", help="輸出資料夾（預設 ./.video-cache/<檔名>）")
    ap.add_argument("--interval", type=float, default=None,
                    help="固定間隔取樣秒數（預設依長度自動：5–30 秒）")
    ap.add_argument("--scene", type=float, default=0.3,
                    help="場景切換門檻 0–1，螢幕錄影建議 0.15（預設 0.3）")
    ap.add_argument("--max-frames", type=int, default=60)
    ap.add_argument("--width", type=int, default=1280, help="關鍵畫面最大寬度")
    ap.add_argument("--no-audio", action="store_true", help="跳過語音辨識")
    ap.add_argument("--model", default="small", help="Whisper 模型：tiny/base/small/medium（預設 small）")
    ap.add_argument("--lang", default="zh", help="語言代碼，auto 為自動偵測（預設 zh）")
    args = ap.parse_args()

    for tool in ("ffmpeg", "ffprobe"):
        if not shutil.which(tool):
            sys.exit(f"找不到 {tool}，請先安裝 ffmpeg。")
    video = Path(args.video).expanduser().resolve()
    if not video.exists():
        sys.exit(f"找不到影片：{video}")
    out_dir = Path(args.out).expanduser().resolve() if args.out else Path.cwd() / ".video-cache" / video.stem
    out_dir.mkdir(parents=True, exist_ok=True)

    print(f"▶ 讀取 {video.name}", file=sys.stderr)
    info = probe(video)
    (out_dir / "info.json").write_text(json.dumps(info, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"  {hms(info['duration'])}，{info['width']}×{info['height']}，"
          f"{'有' if info['has_audio'] else '無'}聲音", file=sys.stderr)

    interval = args.interval or min(30.0, max(5.0, info["duration"] / 30))
    print(f"▶ 偵測場景切換（門檻 {args.scene}）＋每 {interval:.0f} 秒取樣", file=sys.stderr)
    scenes = scene_times(video, args.scene)
    times = pick_times(info["duration"], scenes, interval, args.max_frames)
    print(f"  場景切換 {len(scenes)} 處，共抽 {len(times)} 張", file=sys.stderr)
    frames = extract_frames(video, times, out_dir / "frames", args.width)
    sheets = contact_sheets(frames, out_dir / "sheets")
    (out_dir / "frames.json").write_text(json.dumps(frames, ensure_ascii=False, indent=2), encoding="utf-8")

    segs, asr_error = [], None
    if args.no_audio:
        asr_error = "依參數略過（--no-audio）。"
    elif not info["has_audio"]:
        asr_error = "影片沒有聲音軌。"
    else:
        print("▶ 抽出聲音", file=sys.stderr)
        wav = out_dir / "audio.wav"
        if extract_audio(video, wav):
            print(f"▶ 語音辨識（faster-whisper {args.model}，語言 {args.lang}）", file=sys.stderr)
            segs, asr_error = transcribe(wav, args.model, args.lang)
        else:
            asr_error = "聲音抽不出來。"
    if segs:
        (out_dir / "transcript.txt").write_text(
            "\n".join(f"[{hms(s['start'])} → {hms(s['end'])}] {s['text']}" for s in segs) + "\n", encoding="utf-8")
        (out_dir / "transcript.json").write_text(json.dumps(segs, ensure_ascii=False, indent=2), encoding="utf-8")
    elif asr_error:
        (out_dir / "transcript.txt").write_text(f"（沒有逐字稿）{asr_error}\n", encoding="utf-8")
        print(f"  ⚠️ {asr_error}", file=sys.stderr)

    write_summary(out_dir, info, frames, sheets, segs, asr_error, video)
    print(f"✅ 完成：{out_dir}", file=sys.stderr)
    print(f"   先讀 {out_dir / 'summary.md'}，再讀 {out_dir / 'sheets'} 裡的縮圖總覽。", file=sys.stderr)
    print(str(out_dir))


if __name__ == "__main__":
    main()

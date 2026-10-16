"""錄音轉逐字稿：任何錄音／影片檔 → 依停頓切段 → 每段辨識 → 轉繁體（台灣用語）→ 帶時間的逐字稿。

用法：bash setup.sh（每個新的雲端環境第一次用跑一次）→ python transcribe.py 錄音.m4a [輸出.txt]
模型：SenseVoice（中、英、日、韓、粵）＋ silero 停頓偵測，都在本機跑，錄音不會傳到別的地方。
"""
import os
import subprocess
import sys
import time

import imageio_ffmpeg
import numpy as np
import opencc
import sherpa_onnx

CACHE = os.environ.get("ASR_CACHE", os.path.expanduser("~/.cache/yfs-asr"))      # setup.sh 下載到這裡，不進 repo
MODEL_DIR = os.path.join(CACHE, "sensevoice")
VAD_MODEL = os.path.join(CACHE, "silero_vad.onnx")
SR = 16000


def load_audio(path):
    """用 ffmpeg 轉成 16k 單聲道浮點數；m4a、mp3、amr、mp4 影片都吃。"""
    cmd = [imageio_ffmpeg.get_ffmpeg_exe(), "-nostdin", "-v", "error", "-i", path, "-ac", "1", "-ar", str(SR), "-f", "f32le", "-"]
    raw = subprocess.run(cmd, capture_output=True, check=True).stdout
    return np.frombuffer(raw, dtype=np.float32)


def segments(samples):
    """依停頓切段，回 [(開始秒, 樣本)]；一段最長 20 秒。"""
    cfg = sherpa_onnx.VadModelConfig()
    cfg.silero_vad.model = VAD_MODEL
    cfg.silero_vad.min_silence_duration = 0.4
    cfg.silero_vad.min_speech_duration = 0.25
    cfg.silero_vad.max_speech_duration = 20
    cfg.sample_rate = SR
    vad = sherpa_onnx.VoiceActivityDetector(cfg, buffer_size_in_seconds=60)
    win = cfg.silero_vad.window_size
    out = []
    for i in range(0, len(samples), win):
        vad.accept_waveform(samples[i:i + win])
        while not vad.empty():
            out.append((vad.front.start / SR, np.array(vad.front.samples, dtype=np.float32)))
            vad.pop()
    vad.flush()
    while not vad.empty():
        out.append((vad.front.start / SR, np.array(vad.front.samples, dtype=np.float32)))
        vad.pop()
    return out


def stamp(sec):
    sec = int(sec)
    return f"{sec // 3600:d}:{sec % 3600 // 60:02d}:{sec % 60:02d}" if sec >= 3600 else f"{sec // 60:02d}:{sec % 60:02d}"


def main():
    src = sys.argv[1]
    dst = sys.argv[2] if len(sys.argv) > 2 else os.path.splitext(src)[0] + "_逐字稿.txt"
    t0 = time.time()
    rec = sherpa_onnx.OfflineRecognizer.from_sense_voice(
        model=os.path.join(MODEL_DIR, "model.int8.onnx"), tokens=os.path.join(MODEL_DIR, "tokens.txt"),
        num_threads=os.cpu_count() or 2, use_itn=True, language="auto")
    cc = opencc.OpenCC("s2twp")
    samples = load_audio(src)
    segs = segments(samples)
    lines = []
    for start, seg in segs:
        st = rec.create_stream()
        st.accept_waveform(SR, seg)
        rec.decode_stream(st)
        text = cc.convert(st.result.text.strip())
        if text:
            lines.append(f"[{stamp(start)}] {text}")
    with open(dst, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")
    dur = len(samples) / SR
    print(f"錄音長度 {stamp(dur)}，切成 {len(segs)} 段，花 {time.time() - t0:.0f} 秒；逐字稿：{dst}", file=sys.stderr)
    print("\n".join(lines))


if __name__ == "__main__":
    main()

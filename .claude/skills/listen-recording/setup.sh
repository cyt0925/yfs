#!/usr/bin/env bash
# 聽錄音的工具：裝套件、下載模型（約 2～3 分鐘，下載 1 GB、解開後留 230 MB）。已經裝過就直接結束。
# 模型從 GitHub 的 k2-fsa/sherpa-onnx 發布頁下載；Hugging Face 在這個雲端環境被擋（2026-10-16 實測）。
set -euo pipefail
CACHE="${ASR_CACHE:-$HOME/.cache/yfs-asr}"
REL=https://github.com/k2-fsa/sherpa-onnx/releases/download/asr-models
NAME=sherpa-onnx-sense-voice-zh-en-ja-ko-yue-2024-07-17
python -c "import sherpa_onnx, imageio_ffmpeg, opencc, numpy" 2>/dev/null \
  || pip install -q sherpa-onnx imageio-ffmpeg opencc-python-reimplemented numpy
mkdir -p "$CACHE"
if [ ! -f "$CACHE/sensevoice/model.int8.onnx" ]; then
  curl -sSL "$REL/$NAME.tar.bz2" | tar xjf - -C "$CACHE" "$NAME/tokens.txt" "$NAME/model.int8.onnx" "$NAME/LICENSE"
  rm -rf "$CACHE/sensevoice" && mv "$CACHE/$NAME" "$CACHE/sensevoice"
fi
[ -f "$CACHE/silero_vad.onnx" ] || curl -sSL -o "$CACHE/silero_vad.onnx" "$REL/silero_vad.onnx"
echo "聽錄音工具已就緒：$CACHE"

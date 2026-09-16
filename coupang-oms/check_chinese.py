"""掃 repo 裡的中文有沒有「長得像但沒人用」的怪字。

背景：AI 產生中文時偶爾會把常用字換成筆畫相近的罕見字（「匯」→U+5307、「驟」→U+9A61），
人眼掃過去很難抓到。這些怪字幾乎都不在 Big5（cp950）裡，所以用「能不能編成 Big5」當篩子：
編不過的中文字就列出來、讓 CI 紅燈。

例外：
- 正規表示式的範圍寫法 `一-鿿`（U+4E00～U+9FFF）是合法用途，整段跳過。
- 真的要用罕見字，在同一行尾巴加註解 `rare-ok`。

執行：python check_chinese.py   （在 coupang-oms/ 或 repo 根目錄都可以）
"""
import glob
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PATTERNS = ["coupang-oms/**/*.py", "coupang-oms/**/*.html", "coupang-oms/**/*.js", "coupang-oms/**/*.md",
            "coupang-oms/**/*.txt", "*.md", ".github/workflows/*.yml"]
SKIP = ("node_modules", "static/tailwind", "bootstrap", ".build-tools")
CJK = re.compile(r"[一-鿿]")
RANGE_LITERAL = re.compile(r"[一-鿿]-[一-鿿]")   # 正規式的 一-鿿


def bad_chars(line):
    line = RANGE_LITERAL.sub("", line)
    out = []
    for ch in CJK.findall(line):
        try:
            ch.encode("cp950")
        except UnicodeEncodeError:
            out.append(ch)
    return out


def main():
    problems = []
    for pat in PATTERNS:
        for f in glob.glob(os.path.join(ROOT, pat), recursive=True):
            if any(s in f for s in SKIP):
                continue
            try:
                lines = open(f, encoding="utf-8").read().splitlines()
            except (OSError, UnicodeDecodeError):
                continue
            for i, line in enumerate(lines, 1):
                if "rare-ok" in line:
                    continue
                bad = bad_chars(line)
                if bad:
                    problems.append((os.path.relpath(f, ROOT), i, "".join(sorted(set(bad))), line.strip()[:80]))
    if problems:
        print(f"找到 {len(problems)} 處疑似打錯的罕見中文字（不在 Big5 裡）：")
        for f, i, chars, text in problems:
            print(f"  {f}:{i}  「{chars}」  {text}")
        sys.exit(1)
    print("中文檢查通過：沒有 Big5 以外的罕見字。")


if __name__ == "__main__":
    main()

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""探路用：把玩運彩網頁的結構印到 log，方便決定怎麼抓台灣盤口。
不寫任何檔案，失敗也不影響主流程。"""
import re
import sys
import json
import html as htmlmod
from datetime import datetime, timedelta, timezone

import requests

UA = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/json;q=0.9,*/*;q=0.8",
    "Accept-Language": "zh-TW,zh;q=0.9,en;q=0.8",
    "Referer": "https://www.playsport.cc/",
}
TW = timezone(timedelta(hours=8))
today = datetime.now(TW).strftime("%Y%m%d")
tomorrow = (datetime.now(TW) + timedelta(days=1)).strftime("%Y%m%d")
URLS = [
    ("header", "https://www.playsport.cc/gamesData/result?allianceid=3"),
    ("html", f"https://www.playsport.cc/predict/scale?allianceid=3&gametime={tomorrow}"),
    ("html", "https://www.playsport.cc/predict/games?allianceid=3"),
    ("html", "https://www.playsport.cc/livescore/3"),
    ("json", f"https://ls.playsport.cc/ls_json.php?alliance=3&gamedate={tomorrow}&pbp=1&teamStat=1"),
    ("json", f"https://ls.playsport.cc/ls_json.php?alliance=3&gamedate={today}&pbp=1&teamStat=1"),
    ("rawtable", f"https://www.playsport.cc/gamesData/result?allianceid=3&gametime={today}"),
]


def strip_tags(s):
    s = re.sub(r"<script.*?</script>", " ", s, flags=re.S | re.I)
    s = re.sub(r"<style.*?</style>", " ", s, flags=re.S | re.I)
    s = re.sub(r"<[^>]+>", " ", s)
    s = htmlmod.unescape(s)
    return re.sub(r"\s+", " ", s).strip()


def show_json(text):
    try:
        d = json.loads(text)
    except Exception as e:  # noqa: BLE001
        print("not json:", repr(e), text[:500])
        return
    def walk(o, depth=0, path="$"):
        if depth > 4:
            return
        if isinstance(o, dict):
            print("  " * depth + f"{path}: dict keys={list(o.keys())[:40]}")
            for k in list(o.keys())[:12]:
                walk(o[k], depth + 1, f"{path}.{k}")
        elif isinstance(o, list):
            print("  " * depth + f"{path}: list len={len(o)}")
            if o:
                walk(o[0], depth + 1, f"{path}[0]")
        else:
            print("  " * depth + f"{path} = {str(o)[:120]}")
    walk(d)
    print("json head:", json.dumps(d, ensure_ascii=False)[:3000])


def main():
    for mode, url in URLS:
        print("=" * 100)
        try:
            r = requests.get(url, headers=UA, timeout=25, allow_redirects=True)
        except Exception as e:  # noqa: BLE001
            print("FAIL", url, repr(e))
            continue
        text = r.text or ""
        print("URL", url, "->", r.url, "status", r.status_code, "len", len(text), "type", r.headers.get("content-type"))
        if mode == "json":
            show_json(text)
            continue
        title = re.search(r"<title>(.*?)</title>", text, re.S | re.I)
        print("title:", strip_tags(title.group(1)) if title else None)
        if mode == "header":
            # 聯盟代號對照：抓所有帶 allianceid 或 livescore/<id> 的連結與其文字
            for m in re.finditer(r'<a[^>]+href="([^"]*(?:allianceid=\d+|livescore/\d+)[^"]*)"[^>]*>(.*?)</a>', text, flags=re.S | re.I):
                t = strip_tags(m.group(2))
                if t and len(t) < 30 and "forum/post" not in m.group(1):
                    print("   ", t, "->", m.group(1)[:80])
            continue
        if mode == "rawtable":
            tables = re.findall(r"<table.*?</table>", text, flags=re.S | re.I)
            print("tables:", len(tables))
            if tables:
                print("RAW TABLE 0 (first 6000 chars):")
                print(tables[0][:6000])
            continue
        # html：印出 class 名稱統計、含 odds/盤/讓 的片段、第一個 table 的原始碼
        classes = re.findall(r'class="([^"]+)"', text)
        from collections import Counter
        print("top classes:", Counter(c for cl in classes for c in cl.split()).most_common(60))
        api = sorted(set(re.findall(r'["\']((?:https?:)?\\?/\\?/[^"\'\s<>]*?(?:api|ajax|json|php)[^"\'\s<>]*)["\']', text, re.I)))
        print("api-ish:", api[:30])
        tables = re.findall(r"<table.*?</table>", text, flags=re.S | re.I)
        print("tables:", len(tables))
        for i, t in enumerate(tables[:3]):
            print(f"RAW TABLE {i} (first 5000 chars):")
            print(t[:5000])
        body = strip_tags(text)
        print("text snippet:", body[:2500])
        for m in re.finditer(r"(window\.__[A-Z_]+__|__NEXT_DATA__|__NUXT__|initialState|preloadedState|var\s+\w*(?:game|odds|data)\w*\s*=\s*[\[{])", text, re.I):
            print("inline state:", m.group(0), "at", m.start(), "sample:", text[m.start():m.start() + 600].replace("\n", " "))
    print("=" * 100)
    print("probe done")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:  # noqa: BLE001
        print("probe crashed:", repr(e))
        sys.exit(0)

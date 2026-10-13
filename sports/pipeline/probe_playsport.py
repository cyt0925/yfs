#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""探路用：把玩運彩、台灣運彩網頁的結構印到 log，方便決定怎麼抓台灣盤口。
不寫任何檔案，失敗也不影響主流程。"""
import re
import sys
import html as htmlmod

import requests

UA = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/json;q=0.9,*/*;q=0.8",
    "Accept-Language": "zh-TW,zh;q=0.9,en;q=0.8",
}
URLS = [
    "https://www.playsport.cc/robots.txt",
    "https://www.playsport.cc/",
    "https://www.playsport.cc/predictgame?allianceid=3",
    "https://www.playsport.cc/gamesData/result?allianceid=3",
    "https://www.playsport.cc/livescore?allianceid=3",
    "https://www.sportslottery.com.tw/robots.txt",
    "https://www.sportslottery.com.tw/",
]


def strip_tags(s):
    s = re.sub(r"<script.*?</script>", " ", s, flags=re.S | re.I)
    s = re.sub(r"<style.*?</style>", " ", s, flags=re.S | re.I)
    s = re.sub(r"<[^>]+>", " ", s)
    s = htmlmod.unescape(s)
    return re.sub(r"\s+", " ", s).strip()


def main():
    for url in URLS:
        print("=" * 100)
        try:
            r = requests.get(url, headers=UA, timeout=25, allow_redirects=True)
        except Exception as e:  # noqa: BLE001
            print("FAIL", url, repr(e))
            continue
        text = r.text or ""
        print("URL", url, "->", r.url, "status", r.status_code, "len", len(text), "type", r.headers.get("content-type"))
        if url.endswith("robots.txt"):
            print(text[:3000])
            continue
        title = re.search(r"<title>(.*?)</title>", text, re.S | re.I)
        print("title:", strip_tags(title.group(1)) if title else None)
        hrefs = sorted(set(re.findall(r'href="([^"]+)"', text)))
        inter = [h for h in hrefs if re.search(r"alliance|game|predict|odds|score|lotter|result|bet|sport", h, re.I)]
        print(f"links ({len(hrefs)} total, {len(inter)} interesting):")
        for h in inter[:120]:
            print("   ", h)
        api = sorted(set(re.findall(r'["\']((?:https?:)?/{0,2}[^"\'\s<>]*?(?:api|ajax|json|data)[^"\'\s<>]*)["\']', text, re.I)))
        print("api-ish strings:")
        for a in api[:60]:
            print("   ", a[:200])
        tables = re.findall(r"<table.*?</table>", text, flags=re.S | re.I)
        print("tables:", len(tables))
        for i, t in enumerate(tables[:6]):
            ths = [strip_tags(x) for x in re.findall(r"<th.*?</th>", t, flags=re.S | re.I)]
            trs = re.findall(r"<tr.*?</tr>", t, flags=re.S | re.I)
            print(f"  table {i}: rows={len(trs)} headers={ths[:20]}")
            for tr in trs[:4]:
                tds = [strip_tags(x)[:40] for x in re.findall(r"<t[dh].*?</t[dh]>", tr, flags=re.S | re.I)]
                print("     row:", tds[:16])
        body = strip_tags(text)
        print("text snippet:", body[:1800])
        # 看看有沒有把資料塞在 script 裡的 JSON
        for m in re.finditer(r"(window\.__[A-Z_]+__|__NEXT_DATA__|__NUXT__|initialState|preloadedState)\s*=", text):
            print("inline state:", m.group(0), "at", m.start(), "sample:", text[m.start():m.start() + 400].replace("\n", " "))
    print("=" * 100)
    print("probe done")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:  # noqa: BLE001
        print("probe crashed:", repr(e))
        sys.exit(0)

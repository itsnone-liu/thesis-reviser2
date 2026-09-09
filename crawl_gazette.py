#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""爬取 web.archive.org 上的最高法公报详情页(gongbao.court.gov.cn/Details/*.html)。

输入: cdx 清单(/tmp/cdx.txt, 每行: urlkey ts original ... status ... length)
输出: case_library_src/gbz_<pageid>.json {page_id, url, ts, title, text}
断点续传: 已存在同名输出即跳过。礼貌限速 + 429退避。
"""
import html as H
import json
import os
import re
import sys
import time
import urllib.request

OUT = "case_library_src"
CDX = sys.argv[1] if len(sys.argv) > 1 else "/tmp/cdx.txt"
UA = {"User-Agent": "Mozilla/5.0 (research; caselib builder)"}


def fetch(url, tries=4):
    for k in range(tries):
        try:
            r = urllib.request.urlopen(urllib.request.Request(url, headers=UA), timeout=45)
            return r.read().decode("utf8", "replace")
        except Exception as e:
            if k == tries - 1:
                raise
            time.sleep(8 * (k + 1))


def clean(s):
    i = s.find("End Wayback Rewrite")
    body = H.unescape(s[i:] if i > 0 else s)
    t = re.sub(r"<script.*?</script>", "", body, flags=re.S)
    t = re.sub(r"<style.*?</style>", "", t, flags=re.S)
    t = re.sub(r"<[^>]+>", "\n", t)
    t = re.sub(r"[ \t\u3000]+", " ", t)
    t = re.sub(r"\n\s*\n+", "\n", t)
    return t


def main():
    rows = []
    for line in open(CDX, encoding="utf8", errors="replace"):
        p = line.rstrip("\n").split(" ")
        if len(p) >= 5 and p[4] == "200" and p[2].endswith(".html"):
            rows.append((p[1], p[2]))
    print("待抓", len(rows), flush=True)
    ok = skip = err = 0
    for ts, orig in rows:
        pid = orig.rsplit("/", 1)[1][:-5]
        out = f"{OUT}/gbz_{pid}.json"
        if os.path.exists(out):
            skip += 1
            continue
        wa = f"http://web.archive.org/web/{ts}id_/{orig}"
        try:
            raw = fetch(wa)
        except Exception as e:
            err += 1
            print("ERR", pid, type(e).__name__, flush=True)
            time.sleep(5)
            continue
        t = clean(raw)
        m = re.search(r"<title>([^<]+)</title>", raw)
        title = (m.group(1).split(" - ")[0].strip() if m else "")
        rec = {"page_id": pid, "url": orig, "ts": ts, "title": title, "text": t}
        if len(t) > 800:  # 正文太薄的跳过价值不大
            json.dump(rec, open(out, "w"), ensure_ascii=False)
            ok += 1
        else:
            skip += 1
        time.sleep(1.1)
        if (ok + err) % 50 == 0:
            print(f"进度 ok={ok} skip={skip} err={err}", flush=True)
    print(f"完成 ok={ok} skip={skip} err={err}", flush=True)


if __name__ == "__main__":
    main()

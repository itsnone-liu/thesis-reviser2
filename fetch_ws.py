#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""批量抓取维基文库指导案例官方全文(批量API版) → case_library_src/ws_N.json。

action=query 一次最多50标题, 比逐页抓省50倍请求; 429退避; 断点续抓。
用法: python3 fetch_ws.py [--from 1] [--to 240] [--delay 3]
"""
import json, os, sys, time, urllib.request, urllib.parse, urllib.error

OUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "case_library_src")
os.makedirs(OUT_DIR, exist_ok=True)
UA = {"User-Agent": "thesis-caselib-builder/1.0 (academic; contact: local)"}
API = "https://zh.wikisource.org/w/api.php"


def query_batch(titles):
    """一次query多标题 → {title: wikitext}; 缺页不出现在revisions里。"""
    params = {
        "action": "query", "format": "json", "prop": "revisions",
        "rvprop": "content", "rvslots": "main",
        "titles": "|".join(titles), "redirects": "1",
    }
    url = API + "?" + urllib.parse.urlencode(params)
    for attempt, wait in enumerate((0, 15, 45, 90)):
        if wait:
            print(f"  429退避{wait}s", flush=True)
            time.sleep(wait)
        try:
            req = urllib.request.Request(url, headers=UA)
            with urllib.request.urlopen(req, timeout=40) as r:
                d = json.load(r)
            out = {}
            for p in d.get("query", {}).get("pages", {}).values():
                revs = p.get("revisions") or []
                if revs:
                    wt = revs[0].get("slots", {}).get("main", {}).get("*", "")
                    if wt:
                        out[p.get("title", "")] = wt
            return out
        except urllib.error.HTTPError as e:
            if e.code == 429:
                continue
            print(f"  HTTP{e.code} on batch", flush=True)
            return {}
        except Exception as e:
            print(f"  {type(e).__name__}: {str(e)[:60]}", flush=True)
            time.sleep(3)
    return {}


def main():
    args = sys.argv
    lo = int(args[args.index("--from") + 1]) if "--from" in args else 1
    hi = int(args[args.index("--to") + 1]) if "--to" in args else 240
    delay = float(args[args.index("--delay") + 1]) if "--delay" in args else 3.0
    todo = [n for n in range(lo, hi + 1)
            if not os.path.exists(os.path.join(OUT_DIR, f"ws_{n}.json"))]
    print(f"待抓 {len(todo)} 篇 ({lo}-{hi})", flush=True)
    B = int(args[args.index("--batch") + 1]) if "--batch" in args else 1
    # 注: 40标题/批必429; 6/批短延也429; 稳态=1标题/请求+20s间隔(限流窗滑过后)
    got = 0
    for i in range(0, len(todo), B):
        chunk = todo[i:i + B]
        titles = [f"指导案例{n}号" for n in chunk]
        res = query_batch(titles)
        for n, t in zip(chunk, titles):
            wt = res.get(t) or res.get(t.replace("指导案例", "指導案例"))
            if wt and wt.strip():
                json.dump({"no": n, "title": t, "wikitext": wt},
                          open(os.path.join(OUT_DIR, f"ws_{n}.json"), "w", encoding="utf-8"),
                          ensure_ascii=False)
                got += 1
        print(f"[{i + len(chunk)}/{len(todo)}] 本批命中{len(res)} 累计{got}", flush=True)
        if i + B < len(todo):
            time.sleep(delay)
    print(f"完成: 新抓{got}")


if __name__ == "__main__":
    main()

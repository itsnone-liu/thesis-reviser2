#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""最高法官网指导案例爬取: 走 /shenpan/xiangqing/{id}.html ID区间(39号=13223起连续),
按标题"指导案例N号"识别案例页, 非案例页跳过, 存 case_library_src/spc_N.json。
用法: python3 fetch_spc.py --from 13223 --to 13500 [--delay 0.8]
"""
import json, os, re, sys, time, html as H, urllib.request, urllib.error

OUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "case_library_src")
os.makedirs(OUT_DIR, exist_ok=True)
UA = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) thesis-caselib/1.0"}


def fetch_page(pid: int):
    url = f"https://www.court.gov.cn/shenpan/xiangqing/{pid}.html"
    try:
        req = urllib.request.Request(url, headers=UA)
        with urllib.request.urlopen(req, timeout=20) as r:
            return r.read().decode("utf-8", "replace")
    except urllib.error.HTTPError as e:
        return "" if e.code in (404, 410) else None
    except Exception:
        return None


def extract(src: str):
    """返回 (案号int, 名称, 正文text) 或 None。"""
    t = H.unescape(re.sub(r"<script.*?</script>", "", src, flags=re.S))
    t = re.sub(r"<style.*?</style>", "", t, flags=re.S)
    t = re.sub(r"<[^>]+>", "\n", t)
    t = re.sub(r"\n{2,}", "\n", t)
    m = re.search(r"指导案例\s*(\d{1,3})\s*号[：:]\s*([^\n]{4,60})", t)
    if not m:
        return None
    # 正文起点: 名称之后; 终点: 版权/页脚
    i = t.find(m.group(2), m.end())
    body = t[i if i > 0 else m.end():]
    j = body.find("责任编辑")
    if j < 0:
        j = body.find("最高人民法院版权所有")
    if j > 0:
        body = body[:j]
    return int(m.group(1)), m.group(2).strip(), body.strip()


def main():
    a = sys.argv
    lo = int(a[a.index("--from") + 1]) if "--from" in a else 13223
    hi = int(a[a.index("--to") + 1]) if "--to" in a else 13500
    delay = float(a[a.index("--delay") + 1]) if "--delay" in a else 0.8
    got = other = fail = 0
    for pid in range(lo, hi + 1):
        cache = os.path.join(OUT_DIR, f"spc_{pid}.json")
        if os.path.exists(cache):
            got += 1
            continue
        src = fetch_page(pid)
        if src is None:
            fail += 1
            time.sleep(2)
            continue
        if src:
            r = extract(src)
            if r and 1 <= r[0] <= 300:
                json.dump({"page_id": pid, "no": r[0], "name": r[1], "text": r[2]},
                          open(cache, "w", encoding="utf-8"), ensure_ascii=False)
                got += 1
            else:
                other += 1
        if (pid - lo) % 40 == 0:
            print(f"[id {pid}] 案例页{got} 非案例{other} 失败{fail}", flush=True)
        time.sleep(delay)
    print(f"完成: 案例页{got} 非案例{other} 失败{fail} (id {lo}-{hi})")


if __name__ == "__main__":
    main()

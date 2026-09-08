#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""导入最高人民法院公报批次页(一页含多个指导案例)。
用法: python3 import_gazette.py URL --from-case 45 --to-case 52
先下载官方公报HTML，再按“指导案例N号”切分，核心栏目齐备才落缓存。
"""
import html, json, os, re, sys, urllib.request
OUT=os.path.join(os.path.dirname(os.path.abspath(__file__)),"case_library_src")

def main():
    a=sys.argv; url=a[a.index("--url")+1] if "--url" in a else a[1]
    lo=int(a[a.index("--from-case")+1]); hi=int(a[a.index("--to-case")+1])
    req=urllib.request.Request(url,headers={"User-Agent":"Mozilla/5.0 thesis-caselib/1.0"})
    try:
        src=urllib.request.urlopen(req,timeout=40).read().decode("utf-8","replace")
    except Exception:
        # 调试/断点：允许传本地缓存以免重复请求公报站
        local=a[a.index("--local")+1] if "--local" in a else ""
        if not local: raise
        src=open(local,encoding="utf8",errors="replace").read()
    t=html.unescape(re.sub(r"<script.*?</script>","",src,flags=re.S))
    t=re.sub(r"<style.*?</style>","",t,flags=re.S)
    t=re.sub(r"<[^>]+>","\n",t); t=re.sub(r"\n{2,}","\n",t)
    starts=list(re.finditer(r"指导案例\s*(\d{1,3})\s*号",t))
    found=0
    for i,m in enumerate(starts):
        n=int(m.group(1))
        if not lo<=n<=hi: continue
        seg=t[m.start():starts[i+1].start() if i+1<len(starts) else len(t)]
        # 标题取案号后第一条长文本，去掉日期/页脚
        lines=[x.strip() for x in seg.splitlines() if x.strip()]
        name=""
        for line in lines[1:8]:
            if len(line)>=8 and "最高人民法院" not in line and "发布" not in line and not re.match(r"^\d{4}年",line):
                name=line; break
        if not name or not all(k in seg for k in ("基本案情","裁判要点")):
            # 批次介绍段也会出现“指导案例45-52号”;继续找该号的正文起点
            continue
        json.dump({"page_id":"gazette:"+url,"no":n,"name":name,"text":seg},
                  open(os.path.join(OUT,f"gb_{n}.json"),"w",encoding="utf8"),ensure_ascii=False)
        found+=1
    print(f"公报导入缓存 {found} 篇 ({lo}-{hi})")
if __name__=='__main__': main()

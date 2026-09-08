#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""最高法“指导案例”栏目分页→详情页批量入缓存。
比盲扫article ID可靠: /shenpan/gengduo/77_N.html列出真实详情链接。
用法: python3 fetch_spc_list.py --pages 1 30 --delay 0.8
"""
import html, json, os, re, sys, time, urllib.request
from fetch_spc import extract, OUT_DIR, UA
BASE="https://www.court.gov.cn"

def get(url):
    req=urllib.request.Request(url,headers=UA)
    with urllib.request.urlopen(req,timeout=30) as r: return r.read().decode('utf8','replace')

def main():
    a=sys.argv; lo=int(a[a.index('--pages')+1]); hi=int(a[a.index('--pages')+2]); delay=float(a[a.index('--delay')+1]) if '--delay' in a else .8
    links={}
    for p in range(lo,hi+1):
        try:
            src=get(f"{BASE}/shenpan/gengduo/77_{p}.html")
        except Exception as e:
            print(f'列表页{p}跳过:{type(e).__name__}',flush=True)
            continue
        for href,title in re.findall(r'href="([^"]*?/shenpan/xiangqing/\d+\.html)"[^>]*>([^<]{5,160})',src):
            title=html.unescape(title).strip()
            m=re.search(r'指导案例\s*(\d{1,3})\s*号[：:]',title)
            if m: links[int(m.group(1))]=(BASE+href if href.startswith('/') else href,title)
        print(f'列表页{p}: 累计案例链接{len(links)}',flush=True)
    got=0
    for no,(url,title) in sorted(links.items()):
        # 已有同案号缓存跳过
        if any(os.path.basename(f).startswith('spc_') and json.load(open(f,encoding='utf8')).get('no')==no for f in __import__('glob').glob(OUT_DIR+'/spc_*.json')):
            continue
        try:
            src=get(url); r=extract(src)
            if r:
                pid=int(re.search(r'/(\d+)\.html',url).group(1))
                json.dump({'page_id':pid,'no':r[0],'name':r[1],'text':r[2]},open(f'{OUT_DIR}/spc_{pid}.json','w',encoding='utf8'),ensure_ascii=False)
                got+=1
        except Exception as e: print('跳过',no,url,type(e).__name__,flush=True)
        time.sleep(delay)
    print(f'完成: 新增{got}, 链接总数{len(links)}')
if __name__=='__main__': main()

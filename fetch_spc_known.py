#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""抓取最高法栏目已枚举的详情URL；用于 200-259 等栏目分页。"""
import html,json,os,re,sys,time,urllib.request
OUT='case_library_src'; UA={'User-Agent':'Mozilla/5.0 thesis-caselib/1.0'}
def main():
 urls=[]
 for path in sys.argv[1:]:
  s=open(path,encoding='utf8',errors='replace').read()
  for href,pid0,title in re.findall(r'href="([^"]*?/shenpan/xiangqing/(\d+)\.html)"[^>]*>([^<]{5,180})',s):
   title=html.unescape(title);m=re.search(r'指导(?:性)?案例\s*(\d+)\s*号',title)
   if m: urls.append((int(m.group(1)),int(href.rsplit('/',1)[1][:-5])))
 done=set()
 for f in __import__('glob').glob(OUT+'/spc_*.json'):
  try: done.add(json.load(open(f,encoding='utf8')).get('no'))
  except: pass
 for no,pid in sorted(set(urls)):
  if no in done: continue
  try:
   s=urllib.request.urlopen(urllib.request.Request(f'https://www.court.gov.cn/shenpan/xiangqing/{pid}.html',headers=UA),timeout=35).read().decode('utf8','replace')
   t=html.unescape(re.sub(r'<script.*?</script>','',s,flags=re.S));t=re.sub(r'<style.*?</style>','',t,flags=re.S);t=re.sub(r'<[^>]+>','\n',t);t=re.sub(r'\n{2,}','\n',t)
   m=re.search(r'指导(?:性)?案例\s*'+str(no)+r'\s*[号：]\s*([^\n]{4,140})',t)
   if m and '基本案情' in t and '裁判要点' in t:
    json.dump({'page_id':pid,'no':no,'name':m.group(1).strip(),'text':t},open(f'{OUT}/spc_{pid}.json','w'),ensure_ascii=False);print('OK',no,pid)
   else: print('SKIP',no,pid)
  except Exception as e: print('ERR',no,pid,type(e).__name__)
  time.sleep(.7)
 print('links',len(urls))
if __name__=='__main__':main()

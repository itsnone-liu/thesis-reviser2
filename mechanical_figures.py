# -*- coding: utf-8 -*-
"""Deterministic 2-D mechanical figures for the fixture sample."""
from __future__ import annotations
import os
from typing import Any, Dict, Optional
from PIL import Image, ImageDraw, ImageFont

BLUE="#1f4e79"; RED="#b22222"; GRAY="#666"; LIGHT="#eaf2f8"

def F(n=28,b=False):
    p="C:/Windows/Fonts/simhei.ttf" if b else "C:/Windows/Fonts/simfang.ttf"
    return ImageFont.truetype(p,n) if os.path.exists(p) else ImageFont.load_default()

def canvas(title):
    im=Image.new("RGB",(1800,1200),"white"); d=ImageDraw.Draw(im)
    d.text((900,48),title,font=F(44,True),fill="#111",anchor="ma")
    return im,d

def save(im, *args):
    # Accept both save(im, drawing, out) and the legacy four-argument calls.
    drawing, out = args[-2], args[-1]
    os.makedirs(out,exist_ok=True); p=os.path.join(out,f"drawing_{drawing.get('seq') or drawing.get('id') or 'mech'}.png"); im.save(p,dpi=(180,180)); return p

def dim(d,x1,y1,x2,y2,label,vertical=False):
    d.line((x1,y1,x2,y2),fill=GRAY,width=2)
    if vertical: d.text((x1-18,(y1+y2)//2),label,font=F(24),fill=GRAY,anchor="mm")
    else: d.text(((x1+x2)//2,y1-16),label,font=F(24),fill=GRAY,anchor="mm")

def lug(d,x,y,s=1.0):
    # simplified forged lug outline with two functional ear holes
    pts=[(x,y+140*s),(x,y-20*s),(x+35*s,y-80*s),(x+100*s,y-120*s),(x+190*s,y-120*s),(x+250*s,y-80*s),(x+285*s,y-20*s),(x+285*s,y+140*s)]
    d.line(pts,fill="#111",width=5,joint="curve"); d.line((x,y+140*s,x+285*s,y+140*s),fill="#111",width=5)
    for cx in (x+70*s,x+215*s):
        d.ellipse((cx-32*s,y-42*s,cx+32*s,y+22*s),outline=BLUE,width=5); d.ellipse((cx-12*s,y-22*s,cx+12*s,y+2*s),outline=RED,width=4)
    return pts

def fixture_base(d,x,y,w=760,h=260):
    d.rounded_rectangle((x,y,x+w,y+h),radius=18,fill="#d9e2f3",outline="#111",width=5)
    d.rectangle((x+35,y+35,x+w-35,y+100),fill="#f5f5f5",outline=BLUE,width=4)
    for cx in (x+180,x+w-180): d.ellipse((cx-28,y+68,cx+28,y+124),fill="#aaa",outline="#111",width=4)
    d.line((x+40,y+h-38,x+w-40,y+h-38),fill=BLUE,width=8)

def arrow(d,a,b,label):
    d.line((*a,*b),fill=RED,width=5)
    d.polygon([(b[0],b[1]),(b[0]-14,b[1]-8),(b[0]-10,b[1]+12)],fill=RED)
    d.text((a[0]+8,a[1]-25),label,font=F(24,True),fill=RED)

def render_part(drawing,out):
    im,d=canvas(drawing.get("title") or "后钢板弹簧吊耳零件结构图"); lug(d,500,470,2.0)
    dim(d,500,820,1070,820,"180 mm"); dim(d,420,230,420,750,"80 mm",True); dim(d,640,275,930,275,"孔距 120 mm")
    d.text((100,1020),"功能耳孔：Φ32 mm；零件厚度：25 mm；基准面：底面 A",font=F(27),fill="#222")
    d.text((100,1080),"正文零件结构示意图，不替代正式零件图",font=F(22),fill=GRAY); return save(im,d,drawing,out)

def render_locating(drawing,out):
    im,d=canvas(drawing.get("title") or "一面两孔定位原理图"); fixture_base(d,400,650,900,240)
    lug(d,650,390,1.25)
    for cx,label in [(745,"圆柱销 Φ18"),(925,"菱形销 Φ12")]:
        d.ellipse((cx-24,610,cx+24,658),fill=BLUE,outline="#111",width=3); d.text((cx,570),label,font=F(23),fill=BLUE,anchor="ma")
    arrow(d,(450,880),(450,600),"Z向约束"); arrow(d,(1320,700),(1480,700),"X/Y向约束"); d.text((100,1040),"底面限制3个自由度；圆柱销限制2个；菱形销限制1个",font=F(27),fill="#222"); return save(im,d,drawing,out)

def render_force(drawing,out):
    im,d=canvas(drawing.get("title") or "铣削工序受力分析图"); fixture_base(d,430,690,900,220); lug(d,720,470,1.2)
    arrow(d,(860,430),(860,180),"W=1500 N"); arrow(d,(1030,570),(1450,570),"Fc=2328 N"); arrow(d,(760,560),(500,560),"Ff=1304 N"); arrow(d,(1120,620),(1120,900),"Fp=2049 N")
    d.text((100,1040),"切削力、进给力和背向力按第5章计算值绘制",font=F(27),fill="#222"); return save(im,d,drawing,out)

def render_clamp(drawing,out):
    im,d=canvas(drawing.get("title") or "螺旋夹紧机构结构图"); d.rectangle((400,760,1370,850),fill="#c9c9c9",outline="#111",width=5); d.line((760,330,760,780),fill="#111",width=18); d.ellipse((720,270,800,350),fill="#aaa",outline="#111",width=4); d.rectangle((700,430,1120,500),fill="#d9e2f3",outline="#111",width=4); d.ellipse((1040,485,1150,595),fill="#bbb",outline="#111",width=4); d.text((410,940),"M12夹紧螺杆　压板长度120 mm　夹紧行程10 mm",font=F(27),fill="#222"); arrow(d,(760,250),(760,580),"夹紧方向"); return save(im,d,drawing,out)

def render_fixture(drawing,out):
    im,d=canvas(drawing.get("title") or "后钢板弹簧吊耳铣削夹具总体布局图"); fixture_base(d,390,690,1020,250); lug(d,650,470,1.4); lug(d,950,470,1.4); d.rectangle((1310,560,1380,710),fill="#aaa",outline="#111",width=4); d.text((1385,620),"对刀块 10 mm",font=F(24),fill="#222"); dim(d,570,980,1130,980,"夹具体 320×200×180 mm"); return save(im,d,drawing,out)

    return save(im,d,drawing,out)

def render_body(drawing,out):
    im,d=canvas(drawing.get("title") or "HT200夹具体零件图"); d.rectangle((420,360,1280,760),fill="#d9d9d9",outline="#111",width=5); d.rectangle((500,430,1200,690),fill="#f5f5f5",outline=BLUE,width=4)
    for cx in (650,1050): d.ellipse((cx-45,500,cx+45,590),outline="#111",width=6); d.line((cx-70,635,cx+70,635),fill=BLUE,width=12)
    dim(d,420,820,1280,820,"320 mm"); dim(d,350,360,350,760,"180 mm",True); d.text((100,1030),"壁厚18 mm；加强筋25×12 mm；定位孔距120 mm；材料HT200",font=F(27),fill="#222"); return save(im,d,drawing,out)

    return save(im,d,drawing,out)

def render_generic_fixture(drawing,out):
    im,d=canvas(drawing.get("title") or "夹具工程示意图"); fixture_base(d,420,650,900,250); lug(d,700,430,1.2); d.text((100,1040),"一面两孔定位、螺旋夹紧、对刀块导向",font=F(28),fill="#222"); return save(im,d,drawing,out)

    return save(im,d,drawing,out)

def render_mechanical_figure(drawing: Dict[str,Any], out: str) -> Optional[str]:
    t=" ".join(str(drawing.get(k,"")) for k in ("type","title","description"))
    if "零件结构" in t or "吊耳零件" in t: return render_part(drawing,out)
    if "定位原理" in t or "定位元件" in t: return render_locating(drawing,out)
    if "受力" in t or "切削力" in t or "夹紧力" in t: return render_force(drawing,out)
    if "螺旋夹紧" in t: return render_clamp(drawing,out)
    if "夹具体零件" in t: return render_body(drawing,out)
    if "总体布局" in t or "装配" in t or "工作原理" in t or "结构尺寸" in t: return render_fixture(drawing,out)
    if "夹具" in t: return render_generic_fixture(drawing,out)
    return None

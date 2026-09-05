# -*- coding: utf-8 -*-
"""论文终版 docx 机械缺陷修复器(备份→修复→复检)。
R1 双重编号剥离+引用同步  R2 无空格图注补空格  R3 简单式编号缺口重排+引用同步
R4 重复章头清除  R5 语义引用映射(每篇显式配置)
"""
import os, re, sys, shutil, json
from docx import Document

ROOT = '论文终版'
BAK = f'{ROOT}/_修改备份'
os.makedirs(BAK, exist_ok=True)

CH = re.compile(r"^第[一二三四五六七八九十\d]+章")
DBL = re.compile(r"^([图表])(\d{1,2})\s+(\1\d{1,2}-\d{1,2}\s*\S.*)$")
NOSPACE = re.compile(r"^([图表])(\d{1,2})([^\s\d\-–—：:，。；、])")
VERBY = re.compile(r"显示|如下|所示|可知|可以看出|表明|中可")
CAPT = re.compile(r"^([图表])\s*(\d{1,2})\s+\S")
CAPT_NS = re.compile(r"^([图表])(\d{1,2})\s*\S")
REF = re.compile(r"([如见]?)([图表])(\d{1,2})(?![\d\-–—])")


def set_text(p, new):
    """保留首run格式整体替换段落文本"""
    if not p.runs:
        p.text = new
        return
    p.runs[0].text = new
    for r in p.runs[1:]:
        r.text = ""


def fix_paper(path, ref_maps=None, dry=False):
    log = []
    d = Document(path)
    paras = d.paragraphs
    texts = [p.text for p in paras]

    # ---- R4 重复章头 ----
    i = 0
    while i < len(paras):
        t1 = texts[i].strip()
        if CH.match(t1):
            j = i + 1
            while j < len(paras) and j <= i + 4:
                if texts[j].strip() == t1:
                    # 删除 i 与 j 之间的游离行 + 第一个重复头
                    for k in range(i, j):
                        paras[k]._element.getparent().remove(paras[k]._element)
                        log.append(f"R4删除重复段[{texts[k][:24]}]")
                    texts = [p.text for p in d.paragraphs]
                    paras = d.paragraphs
                    break
                j += 1
        i += 1

    # ---- R2 无空格图注补空格(短行/非引用句) ----
    for p in paras:
        s = p.text.strip()
        if len(s) < 50 and NOSPACE.match(s) and not VERBY.search(s[:16]):
            m = NOSPACE.match(s)
            new = f"{m.group(1)}{m.group(2)} {m.group(3)}{s[m.end():]}"
            set_text(p, new)
            log.append(f"R2图注补空格[{s[:20]}]")

    # ---- R1 双重编号剥离 + 裸引用映射 ----
    dbl_map = {}  # (kind, seq) -> (C, M)
    for p in paras:
        s = p.text.strip()
        m = DBL.match(s)
        if m:
            kind, seq, rest = m.group(1), m.group(2), m.group(3)
            mm = re.match(rf"{kind}(\d{{1,2}})-(\d{{1,2}})", rest)
            dbl_map[(kind, int(seq))] = (int(mm.group(1)), int(mm.group(2)))
            set_text(p, rest)
            log.append(f"R1剥离双重编号[{s[:22]}→{rest[:22]}]")
    if dbl_map:
        for p in paras:
            s = p.text
            if not s.strip() or CAPT.match(s.strip()):
                continue
            def sub(m):
                kind, num = m.group(2), int(m.group(3))
                if (kind, num) in dbl_map:
                    C, M = dbl_map[(kind, num)]
                    return f"{m.group(1)}{kind}{C}-{M}"
                return m.group(0)
            new = REF.sub(sub, s)
            if new != s:
                set_text(p, new)
                log.append(f"R1引用同步[{s[:26]}…]")

    # ---- R5 语义引用映射(显式字面替换) ----
    for old, new in (ref_maps or {}).items():
        for p in paras:
            if old in p.text:
                set_text(p, p.text.replace(old, new))
                log.append(f"R5映射[{old[:18]}→{new[:18]}]")

    # ---- R3 简单式编号缺口重排 ----
    paras = d.paragraphs
    caps = []
    for idx, p in enumerate(paras):
        s = p.text.strip()
        m = CAPT.match(s)
        if m and not VERBY.search(s[:16]) and len(s) < 60:
            caps.append((idx, m.group(1), int(m.group(2))))
    for kind in ("图", "表"):
        kc = [(i, k, n) for i, k, n in caps if k == kind]
        nums = [n for _, _, n in kc]
        if not nums or "-" in "".join(str(n) for n in nums):
            continue
        if sorted(nums) == list(range(1, len(nums) + 1)):
            continue  # 已连续
        # 按文档序重排为 1..N
        remap = {}
        for newn, (i, k, oldn) in enumerate(kc, start=1):
            if oldn != newn:
                remap[oldn] = newn
                s = paras[i].text
                set_text(paras[i], re.sub(rf"^{k}\s*{oldn}\b", f"{k}{newn}", s))
        if remap:
            log.append(f"R3{kind}重排{remap}")
            for p in paras:
                s = p.text
                if CAPT.match(s.strip()):
                    continue
                def sub(m):
                    k, n = m.group(2), int(m.group(3))
                    if k == kind and n in remap:
                        return f"{m.group(1)}{k}{remap[n]}"
                    return m.group(0)
                new = REF.sub(sub, s)
                if new != s:
                    set_text(p, new)

    # ---- R6 孤注清除(图注邻近无图) ----
    if '--r6' not in sys.argv:
        pass
    else:
        W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
        def _img(el):
            return bool(el.findall(f'.//{W}drawing')) or bool(el.findall(f'.//{W}pict'))
        paras = d.paragraphs
        texts = [p.text for p in paras]
        # 顺序配对: 逐段扫描, 图注登记待配, 遇内嵌图分配给最近未配图注
        pending, assigned = [], {}
        for i, para in enumerate(paras):
            s = texts[i].strip()
            if re.match(r"^图\s*\d", s) and not VERBY.search(s[:16]) and len(s) < 60:
                pending.append(i)
            if _img(para._element) and pending:
                assigned[pending.pop(0)] = True
        orphans = [i for i in range(len(paras))
                   if re.match(r"^图\s*\d", texts[i].strip())
                   and not VERBY.search(texts[i].strip()[:16]) and len(texts[i].strip()) < 60
                   and i not in assigned]
        if orphans:
            for idx in reversed(orphans):
                log.append(f"R6删除孤注[{texts[idx][:26]}]")
                paras[idx]._element.getparent().remove(paras[idx]._element)
            paras = d.paragraphs
            # 重排: 简单式全量1..N; 章节式按章内N连续
            def _cap(i):
                m = re.match(r"^图\s*(\d+)(?:-(\d+))?", paras[i].text.strip())
                return (int(m.group(1)), int(m.group(2))) if m else None
            idxs = [i for i in range(len(paras)) if _cap(i) and not VERBY.search(paras[i].text.strip()[:16]) and len(paras[i].text.strip()) < 60]
            if any(_cap(i)[1] for i in idxs):  # 章节式
                remap = {}
                seq = {}
                for i in idxs:
                    C, N = _cap(i)
                    if N is None: continue
                    seq[C] = seq.get(C, 0) + 1
                    if N != seq[C]:
                        remap[(C, N)] = (C, seq[C])
                        set_text(paras[i], re.sub(rf"^图\s*{C}-{N}\b", f"图{C}-{seq[C]}", paras[i].text))
                if remap:
                    log.append(f"R6章节式重排{remap}")
                    for p_ in paras:
                        s = p_.text
                        if _cap(paras.index(p_)) : pass
                        def sub7(m):
                            C, N = int(m.group(2)), int(m.group(3))
                            if m.group(1) == "" and (C, N) in remap:
                                c2, n2 = remap[(C, N)]
                                return f"{m.group(0)[:0]}图{c2}-{n2}"
                            return m.group(0)
                        # 章节式引用
                        new = re.sub(r"([如见]?图)(\d{1,2})-(\d{1,2})", lambda m: (f"{m.group(1)}{remap[(int(m.group(2)),int(m.group(3)))][0]}-{remap[(int(m.group(2)),int(m.group(3)))][1]}" if (int(m.group(2)),int(m.group(3))) in remap else m.group(0)), s)
                        if new != s: set_text(p_, new)
            else:  # 简单式
                remap = {}
                for newn, i in enumerate(idxs, start=1):
                    oldn = _cap(i)[0]
                    if oldn != newn:
                        remap[oldn] = newn
                        set_text(paras[i], re.sub(rf"^图\s*{oldn}\b", f"图{newn}", paras[i].text))
                if remap:
                    log.append(f"R6图重排{remap}")
                    for p_ in paras:
                        s = p_.text
                        if re.match(r"^图\s*\d", s.strip()): continue
                        def sub6(m):
                            n = int(m.group(3))
                            if m.group(2) == "图" and n in remap:
                                return f"{m.group(1)}图{remap[n]}"
                            return m.group(0)
                        new = REF.sub(sub6, s)
                        if new != s: set_text(p_, new)
    if not dry:
        d.save(path)
    return log


# R5 语义映射表(证据来自诊断取证)
SEMANTIC = {
    f'{ROOT}/25/经管/包倩倩_029823430102.docx': {
        "图5数据表明": "图2数据表明",
        "图4显示，服装类": "图1显示，服装类",
    },
    f'{ROOT}/25/经管/谭仁杰_029823431219.docx': {
        "从图3-1可以看出": "从图1可以看出",
        "表3-3的数据显示": "表3的数据显示",
        "其经营数据如表3-1所示": "其经营数据如表1所示",
        "具体成本结构数据如表3-2所示": "具体成本结构数据如表2所示",
        "图3-1展示了胖东来主要": "图1展示了胖东来主要",
        "表3-3进一步对比了胖东": "表3进一步对比了胖东",
    },
    f'{ROOT}/26/机械/孙为民_029822430181.docx': {
        "如图5-1所示": "如图10所示",
        "见表5-1。": "见表2。",
    },
    f'{ROOT}/25/设计/李柏雅_029823410227.docx': {
        "图3-1：": "", "图3-2：": "", "图3-3：": "", "图3-4：": "",
    },
    f'{ROOT}/26/设计/何素珍_029823410119.docx': {
        "，如图1-1(项目区位图)与图1-2(建筑外观图)所示": "",
    },
    f'{ROOT}/26/经管/曹梦丹_029822430338.docx': {
        "如表3-1所示": "如表1所示",
        "结果如表3-4所示": "结果如表3所示",
    },
    f'{ROOT}/26/经管/王庆安_029823431150.docx': {
        "从表4-1可以看出": "从表2可以看出",
        "表4-2的结果表明": "表4的结果表明",
        "由图4-1可知": "由图3可知",
        "具体数据如表4-1所示": "具体数据如表2所示",
        "图4-1展示了不同财政分": "图3展示了不同财政分",
        "表4-2进一步分析了财政": "表4进一步分析了财政",
    },
}

if __name__ == '__main__':
    targets = []
    if len(sys.argv) > 1 and sys.argv[1] == '--one':
        targets = [(sys.argv[2], SEMANTIC.get(sys.argv[2], {}))]
        dry = '--dry' in sys.argv
    else:
        dry = False
        import csv
        for r in csv.DictReader(open(f'{ROOT}/审计报告.csv', encoding='utf-8-sig')):
            need = (r['终审'] == '❌' or '字数过短' in r['问题'] or '编号不连续' in r['问题']
                    or '双重编号' in r['问题'] or '引用失配' in r['问题'])
            if '--r6' in sys.argv:
                need = need or ('vs图片' in r['问题'])
            if need:
                targets.append((f"{ROOT}/{r['文件']}", SEMANTIC.get(f"{ROOT}/{r['文件']}", {})))
    allok = True
    for path, maps in targets:
        base = os.path.basename(path)
        shutil.copy2(path, f"{BAK}/{base}")
        try:
            log = fix_paper(path, maps, dry=dry)
            print(f"✓ {base}: {len(log)}项 " + ("; ".join(log[:3]) + ("…" if len(log) > 3 else "")))
        except Exception as e:
            allok = False
            print(f"✗ {base}: {e}")
    print("完成", len(targets), "篇" if allok else "(有失败)")

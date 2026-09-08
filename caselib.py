# -*- coding: utf-8 -*-
"""
caselib.py — 法学案例库(案例先行模式的底座)
===========================================
- 案例为唯一案情来源; 每案带来源元数据与 verified 标记
- 入库字段与 law.py CASE_MATERIAL_BLOCK 对齐
- T1: 取单案; T2: cluster(domain) 取同主题3-5案
- check_consistency: 生成文本 vs 案例(案号存在/来源词/要素一致)
库文件: case_library/*.json (一案一文件)
"""
import json, os, re, glob
from typing import Optional

LIB_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "case_library")

REQUIRED_FIELDS = ["名称", "案号", "当事人", "案由", "基本事实", "诉讼请求", "裁判结果", "争议焦点", "来源"]
OPTIONAL_FIELDS = ["来源链接", "核心争点短语", "共性问题短语", "裁判要点", "tags", "verified", "领域", "批次"]

def load_library() -> list:
    return [json.load(open(p, encoding="utf-8")) for p in sorted(glob.glob(f"{LIB_DIR}/*.json"))]

def save_case(case: dict) -> str:
    os.makedirs(LIB_DIR, exist_ok=True)
    cid = re.sub(r"[^\w]", "_", case.get("案号") or case.get("名称", "case"))[:60]
    path = f"{LIB_DIR}/{cid}.json"
    json.dump(case, open(path, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    return path

def get(name_or_no: str) -> Optional[dict]:
    k = name_or_no.strip()
    for c in load_library():
        if k in c.get("名称", "") or k == c.get("案号") or k in str(c.get("tags", [])):
            return c
    return None

def cluster(domain: str, n: int = 4, exclude=None, rotate=None) -> list:
    """同领域已核案例群(类型二用), 不足3案返回空。

    0908修复(百篇量产前提):
    - exclude: 本批已用案号/名称集合, 排除后再取样 → 同批不撞
    - rotate: 轮换序号(如任务索引)。不传时保持旧行为(取前N, 兼容单篇调试);
      传入时按 案号+rotate 稳定哈希洗牌后取N → 同领域多篇各拿不同案例群,
      且同一任务重跑结果不变(确定性, 断点续跑安全)
    """
    hits = [c for c in load_library()
            if c.get("verified") and (domain in c.get("领域", "") or domain in str(c.get("tags", [])))]
    if exclude:
        hits = [c for c in hits if str(c.get("案号", "")) not in exclude
                and c.get("名称", "") not in exclude]
    if len(hits) < 3:
        return []
    if rotate is None:
        return hits[:max(n, 3)]
    import hashlib
    def _key(c):
        h = hashlib.md5((str(c.get("案号", "")) + "#" + str(rotate)).encode("utf-8")).hexdigest()
        return (h, str(c.get("案号", "")))
    return sorted(hits, key=_key)[:max(n, 3)]

def validate_case(case: dict) -> list:
    """入库体检: 必填字段/案号格式/verified标记"""
    probs = [f"缺必填字段: {f}" for f in REQUIRED_FIELDS if not str(case.get(f, "")).strip()]
    if case.get("案号") and not re.search(r"[（(]\d{4}[）)].{0,12}第?\d+号|^指导案例\d+号$|第42?批指导", str(case["案号"])):
        probs.append(f"案号格式可疑: {case['案号']}")
    if not case.get("verified"):
        probs.append("未标记verified(要素未核对)")
    return probs

def check_consistency(text: str, cases: list, source_words=("裁判文书网", "指导性案例", "指导案例", "公报", "典型案例", "北大法宝")) -> list:
    """生成文本(全篇) vs 案例素材一致性: 案号出现 + 特征词覆盖 + 来源标注"""
    probs = []
    for c in cases:
        no = str(c.get("案号", ""))
        nm = str(c.get("名称", ""))
        if no and no not in text:
            # 名称主体也认可(指导案例编号常以"第42批/237号"形式出现)
            if not (nm and nm[:6] in text):
                probs.append(f"[{nm[:12]}] 案号未在正文出现({no})")
        # 特征词: 当事人名片段 + 案由核心词
        tokens = [nm[2:6]] if len(nm) >= 6 else []
        for key in ("当事人", "案由"):
            for m in re.findall(r"[\u4e00-\u9fa5]{2,6}", str(c.get(key, "")))[:4]:
                tokens.append(m)
        tokens = list(dict.fromkeys(tokens))[:6]
        miss = [t for t in tokens if t not in text]
        if len(miss) >= 3:
            probs.append(f"[{nm[:12]}] 案例特征词大量缺失: {miss[:3]}")
    joined = "".join(str(c.get("来源", "")) for c in cases)
    if not any(w in text for w in source_words) and joined:
        probs.append("正文未标注任何案例来源词")
    return probs

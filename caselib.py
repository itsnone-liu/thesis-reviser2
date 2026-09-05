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

def cluster(domain: str, n: int = 4) -> list:
    """同领域已核案例群(类型二用), 不足3案返回空"""
    hits = [c for c in load_library()
            if c.get("verified") and (domain in c.get("领域", "") or domain in str(c.get("tags", [])))]
    return hits[:max(n, 3)] if len(hits) >= 3 else []

def validate_case(case: dict) -> list:
    """入库体检: 必填字段/案号格式/verified标记"""
    probs = [f"缺必填字段: {f}" for f in REQUIRED_FIELDS if not str(case.get(f, "")).strip()]
    if case.get("案号") and not re.search(r"[（(]\d{4}[）)].{0,12}第?\d+号|^指导案例\d+号$|第42?批指导", str(case["案号"])):
        probs.append(f"案号格式可疑: {case['案号']}")
    if not case.get("verified"):
        probs.append("未标记verified(要素未核对)")
    return probs

def check_consistency(text: str, cases: list, source_words=("裁判文书网", "指导案例", "公报", "典型案例", "北大法宝")) -> list:
    """生成文本 vs 案例素材一致性"""
    probs = []
    for c in cases:
        no = c.get("案号", "")
        if no and no not in text and "指导案例" not in str(c.get("来源", "")):
            probs.append(f"[{c.get('名称','?')}] 案号未在正文出现")
        for key_fact in re.split(r"[。;；]", str(c.get("基本事实", "")))[:3]:
            core = re.sub(r"[的了与及或在对于向从中被是]", "", key_fact)[:18]
            if len(core) >= 8 and core not in re.sub(r"[的了与及或在对于向从中被是]", "", text):
                probs.append(f"[{c.get('名称','?')}] 关键事实疑似缺失: {key_fact[:24]}…")
    joined = "".join(str(c.get("来源", "")) for c in cases)
    if not any(w in text for w in source_words) and joined:
        probs.append("正文未标注任何案例来源词")
    return probs

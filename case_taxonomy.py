# -*- coding: utf-8 -*-
"""案例库受控领域词表(taxonomy) — 400建库配比基准。

用途: ①入库时 领域 字段必须取 AGG 字典键之一(或其别名映射);
      ②cluster(domain) 按此词表匹配, 杜绝"劳动法/劳动关系"分裂成两池;
      ③stats() 按配比报缺口, 建库优先级由此驱动。
"""
from collections import Counter

# 领域 → (目标配比@400, 别名/子领域)
AGG = {
    # 目标合计严格为400；这是配额，不代表已收录数量。
    "合同法":       (55, ["股权转让", "企业借贷", "海上货物运输合同", "买卖合同", "借款合同", "确认合同无效", "居间合同", "赠与合同", "租赁合同", "担保", "保证", "定金", "违约金"]),
    "侵权责任":     (50, ["机动车交通事故", "船舶碰撞", "航空旅客运输", "交通运输", "产品责任", "网络侵权", "银行卡纠纷", "网络盗刷", "安全保障义务", "医疗损害", "生命权", "人身权", "饲养动物", "环境污染", "生态环境", "生态环境损害赔偿", "名誉权", "荣誉权", "故意伤害"]),
    "劳动法":       (40, ["劳动合同", "竞业限制", "劳动关系认定", "新就业形态", "工伤", "劳务派遣", "拖欠劳动报酬", "追索劳动报酬", "平等就业权", "就业歧视"]),
    "公司商事":     (40, ["公司法", "公司解散", "公司决议", "股东", "破产", "保险", "票据", "海事", "市场支配地位", "垄断"]),
    "刑法":         (35, ["故意杀人", "盗窃", "诈骗", "抢劫", "贪污", "受贿", "贪污受贿", "挪用公款", "滥用职权", "危险驾驶", "拒不支付劳动报酬", "非法经营", "破坏计算机信息系统", "刑事案件", "刑事"]),
    "婚姻家事":     (30, ["离婚", "抚养", "监护权", "探望权", "继承", "赡养", "收养", "夫妻共同财产"]),
    "消保与网络":   (30, ["业主共有权", "专项维修资金", "消费者权益", "食品安全", "欺诈", "个人信息", "网络购物", "数据权益"]),
    "知识产权":     (25, ["专利", "商标", "著作权", "集成电路布图设计", "技术秘密", "不正当竞争", "植物新品种"]),
    "民事诉讼法":   (25, ["管辖", "证据", "民事诉讼", "虚假诉讼", "执行", "再审", "诉讼保全", "公益诉讼", "反垄断诉讼程序", "颁发证书", "高等学校"]),
    "行政法":       (25, ["行政处罚", "行政许可", "行政征收", "政府信息公开", "行政复议", "国家赔偿", "行政诉讼", "公安行政登记", "土地使用权"]),
    "建设工程":     (25, ["施工合同", "工程款", "招标投标", "别除权", "优先受偿权", "人防", "房地产开发"]),
    "涉外与其他":   (20, ["涉外", "国际商事", "仲裁", "海难救助", "外国法院民事判决", "其他"]),
}
TARGET_TOTAL = 400

_ALIAS2AGG = {}
for agg, (_t, aliases) in AGG.items():
    _ALIAS2AGG[agg] = agg
    for a in aliases:
        _ALIAS2AGG.setdefault(a, agg)


def normalize(domain: str) -> str:
    """任意领域串 → 受控AGG键; 无法归类返回空串(入库时报警)。"""
    d = (domain or "").strip()
    if not d:
        return ""
    if d in _ALIAS2AGG:
        return _ALIAS2AGG[d]
    # 子串匹配取最长命中(受控优先)
    best = ""
    for alias, agg in _ALIAS2AGG.items():
        if alias in d and len(alias) > len(best):
            best = alias
    return _ALIAS2AGG.get(best, "")


def classify(*texts) -> str:
    """从案由/名称/关键词文本推断领域(规则式, 零LLM)。"""
    blob = " ".join(t for t in texts if t)
    scores = Counter()
    for alias, agg in _ALIAS2AGG.items():
        if alias in blob:
            scores[agg] += min(len(alias), 6)  # 长别名权重高
    if not scores:
        return ""
    # 并列时按目标配比大的优先(常用领域优先)
    top = max(scores.values())
    cands = [a for a, s in scores.items() if s == top]
    return sorted(cands, key=lambda a: -AGG[a][0])[0]


def stats(cases: list) -> dict:
    """库内领域分布 vs 目标配比 → 缺口(建库优先级)。"""
    c = Counter()
    verified = Counter()
    for case in cases:
        agg = normalize(case.get("领域", ""))
        if agg:
            c[agg] += 1
            if case.get("verified"):
                verified[agg] += 1
    return {
        "total": sum(c.values()),
        "verified": sum(verified.values()),
        "target": TARGET_TOTAL,
        "by_agg": {agg: {"have": c.get(agg, 0), "verified": verified.get(agg, 0),
                         "target": t}
                   for agg, (t, _a) in sorted(AGG.items(), key=lambda kv: -kv[1][0])},
        "gap": {agg: max(0, t - c.get(agg, 0))
                for agg, (t, _a) in AGG.items() if c.get(agg, 0) < t},
        "unknown_domain": sum(1 for case in cases if not normalize(case.get("领域", ""))),
    }

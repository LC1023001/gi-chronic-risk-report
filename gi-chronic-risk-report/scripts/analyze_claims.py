# -*- coding: utf-8 -*-
"""
团单赔案明细分析器：读取保险公司导出的赔案明细清单(xls)与赔付统计报表(xlsx)，
执行数据质量诊断 + 全维度画像 + 病种归类 + 慢病队列 + 红黄绿分层，输出 JSON。

用法:
  python analyze_claims.py --files "明细1.xls" "明细2.xls" [--stats "统计报表.xlsx"]
                           --unit "企业全称" --short "简称" --out _analytics.json

依赖: xlrd (pip install xlrd)；统计报表用标准库 zipfile 解析，无需 openpyxl。
"""
import argparse, json, statistics, sys, zipfile, re
from collections import Counter, defaultdict

# ============ 词库（详见 references/disease-catalog.md） ============
CATS = [
 ('恶性肿瘤', ['恶性肿瘤','癌','白血病','淋巴瘤','骨髓瘤']),
 ('良性肿瘤及囊肿', ['良性肿瘤','囊肿','息肉','肌瘤','脂肪瘤','腺瘤']),
 ('循环系统疾病', ['高血压','动脉硬化','冠心病','心梗','心肌梗死','心绞痛','脑血管','脑梗死','脑出血','脑缺血','心律','心悸','房颤','心力衰竭','心功能','静脉曲张','动脉粥样硬化','心脏','动脉瘤','肺动脉','静脉炎']),
 ('内分泌及代谢疾病', ['糖尿病','高脂血症','血脂','甲状腺','痛风','尿酸','肥胖','代谢','低血糖','碘缺乏']),
 ('呼吸系统疾病', ['上呼吸道感染','急性鼻咽炎','支气管','肺炎','气管炎','鼻炎','咽炎','哮喘','咳嗽','肺','扁桃体','感冒','呼吸','胸膜','气胸','喉炎','肺气肿','慢阻肺']),
 ('消化系统疾病', ['胃炎','胃肠炎','肠炎','消化不良','便秘','腹泻','肝','胆','胰腺','痔疮','腹痛','结肠','直肠','胃','肠','牙','口腔','舌','唾液','食管','恶心','呕吐','疝']),
 ('肌肉骨骼疾病', ['关节','颈椎','腰椎','腰痛','坐骨神经','骨质疏松','骨','肌肉','肩周','椎间盘','骨关节','软骨','腱','脊柱','背痛','颈痛']),
 ('泌尿生殖系统疾病', ['肾','泌尿','前列腺','膀胱','尿道','子宫','卵巢','宫颈','阴道','月经','外阴','盆腔','睾丸','附睾','乳腺炎','乳腺增生']),
 ('眼科疾病', ['白内障','青光眼','结膜炎','干眼','眼','视力','视网膜','玻璃体','屈光','睑缘炎','麦粒肿']),
 ('耳鼻咽喉(除呼吸)', ['耳','耳鸣','听力','中耳','外耳']),
 ('神经及精神疾病', ['失眠','睡眠障碍','头晕','眩晕','头痛','神经','焦虑','抑郁','帕金森','癫痫','痴呆','脑神经','面瘫','多汗','自主神经','紧张','精神']),
 ('皮肤及皮下疾病', ['皮炎','湿疹','荨麻疹','皮疹','痤疮','癣','银屑病','皮肤','带状疱疹','疣','毛囊','指甲','瘙痒','蜂窝织炎','烧伤','烫伤']),
 ('损伤中毒及外伤', ['骨折','扭伤','挫伤','外伤','损伤','咬伤','异物','中毒','裂伤','摔','切割']),
 ('血液系统疾病', ['贫血','血小板','白细胞减少','血友病','紫癜']),
 ('传染病及寄生虫', ['结核','水痘','麻疹','腮腺炎','病毒性肝炎','手足口','感染性腹泻']),
 ('症状体征及未特指', ['不适','疲劳','疼痛','发热','症状','衰老','检查','健康查体','无','正常']),
]
CHRONIC = ['高血压','糖尿病','冠心病','动脉硬化性心脏病','心绞痛','心梗','心肌梗死','脑梗死','脑血管','脑出血','高脂血症','痛风','尿酸','骨质疏松','慢性支气管炎','肺气肿','哮喘','甲状腺功能','前列腺增生','肾功能','肾衰竭','透析','肝硬化','房颤','心功能不全','心力衰竭','颈动脉斑块']
CHRONIC_TYPES = {'高血压':['高血压'],'糖尿病':['糖尿病'],'冠心病/心梗':['冠心病','心梗','心肌梗死','心绞痛'],
  '脑血管病':['脑梗死','脑血管','脑出血','脑缺血'],'高脂血症':['高脂血症'],'痛风/高尿酸':['痛风','尿酸'],
  '骨质疏松':['骨质疏松'],'甲状腺疾病':['甲状腺'],'慢支/哮喘':['慢性支气管炎','哮喘','肺气肿'],'前列腺增生':['前列腺增生']}

def cat_of(diag):
    d = str(diag)
    for cat, kws in CATS:
        for kw in kws:
            if kw in d: return cat
    return '其他'
def is_chronic(diag):
    return any(kw in str(diag) for kw in CHRONIC)
def num(x):
    try: return float(x)
    except (TypeError, ValueError): return 0.0
def age_of(x):
    try:
        a = int(float(x))
        return a if 0 < a < 120 else None
    except (TypeError, ValueError): return None

# ============ 读取 ============
# 标准列名 -> 系统导出常见别名（按序探测）
ALIASES = {
    '报案人证件号': ['报案人证件号', '被保人证件号', '被保险人证件号'],
    '性别':        ['性别', '被保人性别', '被保险人性别'],
    '年龄':        ['年龄', '被保人年龄', '被保险人年龄'],
    '投保单位':    ['投保单位', '投保单位名称'],
    '诊断名称':    ['诊断名称', '疾病诊断名称'],
    '案件摘要':    ['案件摘要', '出险摘要'],
    '赔案号':      ['赔案号', '案件编号'],
    '出险类型':    ['出险类型'],
    '核定金额':    ['核定金额'],
    '应给付金额':  ['应给付金额'],
    '出险日期':    ['出险日期'],
    '报案日期':    ['报案日期'],
    '报案方式':    ['报案方式'],
    '发票张数':    ['发票张数'],
}

def load_xls(path):
    import xlrd
    wb = xlrd.open_workbook(path)
    sh = wb.sheet_by_index(0)
    if sh.nrows < 3: return []
    hdr = [str(c.value).strip() for c in sh.row(1)]
    idx = {}
    for std, alts in ALIASES.items():
        for a in alts:
            if a in hdr: idx[std] = hdr.index(a); break
    missing = [h for h in ALIASES if h not in idx]
    if missing:
        raise SystemExit(f"[{path}] 缺少必需列: {missing}；实际表头: {hdr}")
    rows = []
    for r in range(2, sh.nrows):
        if not str(sh.cell_value(r, 0)).strip(): continue
        rows.append({h: sh.cell_value(r, i) for h, i in idx.items()})
    return rows

def load_xlsx_stats(path):
    """赔付统计报表(xlsx, 含图片)——zipfile+XML 重建，返回行列表(每行为单元格值序列)"""
    z = zipfile.ZipFile(path)
    ss = z.read('xl/sharedStrings.xml').decode('utf-8')
    strings = []
    for si in re.findall(r'<si>(.*?)</si>', ss, re.S):
        strings.append(''.join(re.findall(r'<t[^>]*>([^<]*)</t>', si)))
    sheet = z.read('xl/worksheets/sheet1.xml').decode('utf-8')
    out = []
    for rxml in re.findall(r'<row[^>]*>(.*?)</row>', sheet, re.S):
        vals = []
        for t, v in re.findall(r'<c r="[A-Z]+\d+"(?:[^>]*t="(\w+)")?[^>]*>(?:<v>([^<]*)</v>)?', rxml):
            if t == 's' and v: vals.append(strings[int(v)])
            elif v: vals.append(v)
            else: vals.append('')
        if any(str(x).strip() for x in vals): out.append(vals)
    return out

# ============ 质量诊断 ============
def diagnose(rows):
    D = {}
    months = Counter(str(r['出险日期'])[5:7] for r in rows)
    top_month, top_n = months.most_common(1)[0] if months else ('', 0)
    D['某月出险占比'] = {'month': top_month, 'pct': round(top_n / len(rows) * 100, 1) if rows else 0}
    gaps = []
    for r in rows:
        try:
            from datetime import datetime
            d1 = datetime.strptime(str(r['出险日期'])[:10], '%Y-%m-%d')
            d2 = datetime.strptime(str(r['报案日期'])[:10], '%Y-%m-%d')
            gaps.append((d2 - d1).days)
        except Exception: pass
    if gaps:
        D['报案-出险间隔'] = {'median': statistics.median(gaps), 'mean': round(statistics.mean(gaps), 1)}
    invs = [num(r.get('发票张数', 0)) for r in rows if str(r.get('发票张数', '')).strip()]
    D['案均发票张数'] = round(statistics.mean(invs), 1) if invs else 0
    D['批次报销伪像判定'] = bool(
        (D['某月出险占比']['pct'] >= 20) or
        (gaps and statistics.median(gaps) > 90) or
        D['案均发票张数'] > 5)
    years = sorted(set(str(r['出险日期'])[:4] for r in rows))
    D['覆盖年度'] = years
    missing = [y for y in range(int(years[0]), int(years[-1]) + 1) if str(y) not in years] if len(years) > 1 else []
    D['缺失年度'] = missing
    yc = Counter(str(r['出险日期'])[:4] for r in rows)
    D['临界行数疑似截断'] = [y for y, c in yc.items() if c in (2499, 2500, 4999, 5000)]
    return D

# ============ 全维度画像 ============
def analyze(name, rows):
    R = {'name': name, 'total_cases': len(rows), 'total_paid': round(sum(num(r['应给付金额']) for r in rows), 2)}
    R['total_check'] = round(sum(num(r['核定金额']) for r in rows), 2)
    pid = lambda r: str(r['报案人证件号'])
    R['persons'] = len(set(pid(r) for r in rows))
    # 年度
    yr = defaultdict(lambda: {'cases': 0, 'paid': 0.0, 'persons': set(), 'hosp': 0, 'hosp_paid': 0.0, 'opd': 0, 'opd_paid': 0.0})
    for r in rows:
        y = str(r['出险日期'])[:4]; d = yr[y]
        d['cases'] += 1; d['paid'] += num(r['应给付金额']); d['persons'].add(pid(r))
        if '住院' in str(r['出险类型']): d['hosp'] += 1; d['hosp_paid'] += num(r['应给付金额'])
        else: d['opd'] += 1; d['opd_paid'] += num(r['应给付金额'])
    R['years'] = {y: {k: (len(v[k]) if k == 'persons' else round(v[k], 2)) for k in v} for y, v in sorted(yr.items())}
    # 病种
    dg = defaultdict(lambda: {'cases': 0, 'paid': 0.0, 'hosp': 0, 'persons': set()})
    for r in rows:
        c = cat_of(r['诊断名称']); d = dg[c]
        d['cases'] += 1; d['paid'] += num(r['应给付金额']); d['persons'].add(pid(r))
        if '住院' in str(r['出险类型']): d['hosp'] += 1
    R['disease'] = {c: {k: (len(v[k]) if k == 'persons' else round(v[k], 2)) for k in v}
                    for c, v in sorted(dg.items(), key=lambda x: -x[1]['paid'])}
    # 金额分层
    bands = [(0, 300), (300, 500), (500, 1000), (1000, 3000), (3000, 10000), (10000, 30000), (30000, 1e12)]
    bl = [{'lo': lo, 'hi': (None if hi > 1e11 else hi), 'cases': 0, 'paid': 0.0} for lo, hi in bands]
    for r in rows:
        p = num(r['应给付金额'])
        for b in bl:
            if b['lo'] <= p < b['hi'] if b['hi'] else p >= b['lo']: b['cases'] += 1; b['paid'] += p; break
    for b in bl: b['paid'] = round(b['paid'], 2)
    R['bands'] = bl
    # 年龄/性别
    ab = [('<40', 0, 40), ('40-49', 40, 50), ('50-59', 50, 60), ('60-69', 60, 70), ('70+', 70, 200)]
    ag = {b[0]: {'cases': 0, 'paid': 0.0, 'persons': set()} for b in ab}
    sx = defaultdict(lambda: {'cases': 0, 'paid': 0.0, 'persons': set()})
    for r in rows:
        a = age_of(r.get('年龄'))
        if a is not None:
            for n, lo, hi in ab:
                if lo <= a < hi: g = ag[n]; g['cases'] += 1; g['paid'] += num(r['应给付金额']); g['persons'].add(pid(r)); break
        g = sx[(str(r.get('性别', '')).strip() or '未知')]
        g['cases'] += 1; g['paid'] += num(r['应给付金额']); g['persons'].add(pid(r))
    R['age'] = {n: {'cases': v['cases'], 'paid': round(v['paid'], 2), 'persons': len(v['persons'])} for n, v in ag.items()}
    R['sex'] = {k: {'cases': v['cases'], 'paid': round(v['paid'], 2), 'persons': len(v['persons'])} for k, v in sx.items()}
    # 人级：慢病队列/频次/分层
    per = defaultdict(lambda: {'paid': 0.0, 'cases': 0, 'chronic': False, 'hosp': 0, 'sex': '', 'age': None, 'diags': set()})
    for r in rows:
        p = per[pid(r)]
        p['paid'] += num(r['应给付金额']); p['cases'] += 1
        p['sex'] = str(r.get('性别', '')); p['age'] = age_of(r.get('年龄')) or p['age']
        p['diags'].add(str(r['诊断名称'])[:20])
        if '住院' in str(r['出险类型']): p['hosp'] += 1
        if is_chronic(r['诊断名称']): p['chronic'] = True
    cp = {k: v for k, v in per.items() if v['chronic']}
    np_ = {k: v for k, v in per.items() if not v['chronic']}
    R['chronic'] = {
        'persons': len(cp), 'all_persons': len(per),
        'cases': sum(v['cases'] for v in cp.values()), 'paid': round(sum(v['paid'] for v in cp.values()), 2),
        'hosp_persons': sum(1 for v in cp.values() if v['hosp'] > 0),
        'non_persons': len(np_), 'non_cases': sum(v['cases'] for v in np_.values()),
        'non_paid': round(sum(v['paid'] for v in np_.values()), 2),
        'non_hosp_persons': sum(1 for v in np_.values() if v['hosp'] > 0)}
    fr = Counter()
    for v in per.values():
        c = v['cases']
        fr['1次' if c == 1 else '2-3次' if c <= 3 else '4-6次' if c <= 6 else '7-9次' if c <= 9 else '10次以上'] += 1
    R['freq'] = dict(fr)
    red = {k: v for k, v in per.items() if v['chronic'] and (v['cases'] >= 7 or v['paid'] >= 20000)}
    yel = {k: v for k, v in per.items() if v['chronic'] and not (v['cases'] >= 7 or v['paid'] >= 20000)}
    grn = {k: v for k, v in per.items() if not v['chronic']}
    tier = lambda g: {'persons': len(g), 'cases': sum(v['cases'] for v in g.values()), 'paid': round(sum(v['paid'] for v in g.values()), 2)}
    R['tiers'] = {'红': tier(red), '黄': tier(yel), '绿': tier(grn)}
    # 慢病检出细分（人级）
    ct = defaultdict(set)
    for r in rows:
        for c, kws in CHRONIC_TYPES.items():
            if any(k in str(r['诊断名称']) for k in kws): ct[c].add(pid(r))
    R['chronic_types'] = {c: len(s) for c, s in sorted(ct.items(), key=lambda x: -len(x[1]))}
    # 肿瘤队列
    tu = [r for r in rows if any(k in str(r['诊断名称']) for k in ['恶性肿瘤', '癌'])]
    R['tumors'] = {'cases': len(tu), 'paid': round(sum(num(r['应给付金额']) for r in tu), 2),
                   'persons': len(set(pid(r) for r in tu)),
                   'list': [{'sex': str(r.get('性别')), 'age': age_of(r.get('年龄')), 'diag': str(r['诊断名称'])[:22],
                             'paid': round(num(r['应给付金额']), 2), 'year': str(r['出险日期'])[:4]}
                            for r in sorted(tu, key=lambda r: -num(r['应给付金额']))[:10]]}
    # 大额个案 TOP10（同人合并）
    top_persons = sorted(per.items(), key=lambda x: -x[1]['paid'])[:12]
    R['top_persons'] = [{'sex': v['sex'], 'age': v['age'], 'cases': v['cases'], 'hosp': v['hosp'],
                         'paid': round(v['paid'], 2), 'diags': sorted(v['diags'])[:3]} for _, v in top_persons]
    top_cases = sorted(rows, key=lambda r: -num(r['应给付金额']))[:12]
    R['top_cases'] = [{'sex': str(r.get('性别')), 'age': age_of(r.get('年龄')), 'diag': str(r['诊断名称'])[:26],
                       'summary': str(r.get('案件摘要', ''))[:60], 'type': '住院' if '住院' in str(r['出险类型']) else '门急诊',
                       'paid': round(num(r['应给付金额']), 2), 'year': str(r['出险日期'])[:4],
                       'inv': int(num(r.get('发票张数', 0)))} for r in top_cases]
    # 报案月节奏（口径安全的时间维度）
    bm = defaultdict(lambda: [0, 0.0])
    for r in rows:
        m = str(r['报案日期'])[:7]; bm[m][0] += 1; bm[m][1] += num(r['应给付金额'])
    R['report_month'] = {k: [v[0], round(v[1], 2)] for k, v in sorted(bm.items())}
    ch = defaultdict(Counter)
    for r in rows: ch[str(r['出险日期'])[:4]][str(r.get('报案方式', ''))] += 1
    R['channel_by_year'] = {y: dict(c) for y, c in sorted(ch.items())}
    # 症状体征专项（不适和疲劳核验入口）
    bf = [r for r in rows if cat_of(r['诊断名称']) == '症状体征及未特指']
    R['unspecified'] = {'cases': len(bf), 'pct': round(len(bf) / len(rows) * 100, 1) if rows else 0,
                        'paid': round(sum(num(r['应给付金额']) for r in bf), 2)}
    R['diag'] = diagnose(rows)
    return R

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--files', nargs='+', required=True, help='赔案明细清单 xls 文件')
    ap.add_argument('--stats', help='团体保单赔付统计报表 xlsx（可选）')
    ap.add_argument('--unit', required=True); ap.add_argument('--short', required=True)
    ap.add_argument('--out', default='_analytics.json')
    a = ap.parse_args()
    rows = []
    for f in a.files:
        rs = load_xls(f)
        print(f"[load] {f}: {len(rs)} 件")
        rows += rs
    units = set(str(r['投保单位']) for r in rows)
    print(f"[unit] {units}")
    R = analyze(a.unit, rows)
    R['short'] = a.short
    if a.stats:
        R['stats_raw'] = load_xlsx_stats(a.stats)
    json.dump(R, open(a.out, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    # 摘要
    d = R['diag']; c = R['chronic']
    print(f"\n===== 摘要 [{a.short}] =====")
    print(f"总件数 {R['total_cases']} | 总给付 {R['total_paid']:,.0f} | 出险人数 {R['persons']}")
    print(f"覆盖年度 {d['覆盖年度']} | 缺失年度 {d['缺失年度']} | 截断疑似 {d['临界行数疑似截断']}")
    print(f"批次报销伪像: {'⚠ 是——时间分析改用报案日期口径' if d['批次报销伪像判定'] else '否'} (间隔中位{d.get('报案-出险间隔',{}).get('median','?')}天, 案均发票{d['案均发票张数']}张)")
    print(f"慢病 {c['persons']}/{c['all_persons']}人 ({c['persons']/max(c['all_persons'],1)*100:.1f}%) 占给付 {c['paid']/max(R['total_paid'],1)*100:.1f}% | 人均 {c['paid']/max(c['persons'],1):,.0f} vs 非慢病 {c['non_paid']/max(c['non_persons'],1):,.0f}")
    print(f"分层 红{R['tiers']['红']['persons']}人/{R['tiers']['红']['paid']:,.0f} 黄{R['tiers']['黄']['persons']}人 绿{R['tiers']['绿']['persons']}人")
    print(f"病种TOP5: {[(k, v['cases'], round(v['paid'])) for k, v in list(R['disease'].items())[:5]]}")
    print(f"症状未特指占比 {R['unspecified']['pct']}% {'(>15%→执行不适和疲劳专项核验)' if R['unspecified']['pct'] > 15 else ''}")
    print(f"\n[ok] 已输出 {a.out}")

if __name__ == '__main__':
    main()

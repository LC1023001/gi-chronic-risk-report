# -*- coding: utf-8 -*-
"""
团单报告四重校验器：HTML 结构 / 图表对应 / JS 语法 / 敏感词 / 隐私。

用法: python validate_report.py 报告.html [报告2.html ...]
退出码 0=全部通过, 1=存在问题。详细规则见 references/report-perspective.md。
"""
import io, re, subprocess, sys, tempfile, os
from html.parser import HTMLParser

VOID = {'meta', 'link', 'br', 'hr', 'img', 'input'}

# 视角禁词（保险公司内部语汇 + 销售经营语汇 + AI 味表述）
BANNED = ['核保', '定价', '赔付率', '盈亏', '亏损', '费率', '涨价', '精算', '对赌',
          '满期', '超赔', '承保', '免赔额', 'Stop-Loss', '续保率', '加保率',
          '粘性', '离不开', '换不起', '留得住', '客户经营', '一句话：', '风险敞口']

class TagChecker(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.stack, self.problems = [], []
    def handle_starttag(self, tag, attrs):
        if tag not in VOID: self.stack.append((tag, self.getpos()[0]))
    def handle_endtag(self, tag):
        if tag in VOID: return
        if self.stack and self.stack[-1][0] == tag: self.stack.pop()
        else: self.problems.append((self.getpos()[0], tag))

def validate(path):
    fails = []
    html = io.open(path, encoding='utf-8').read()
    # 1) HTML 结构
    p = TagChecker(); p.feed(html)
    if p.problems: fails.append(f"标签错配: {p.problems[:5]}")
    if p.stack: fails.append(f"未闭合标签: {[(t, l) for t, l in p.stack][:5]}")
    # 2) 图表对应
    ids = re.findall(r'id="(chart[A-Za-z0-9]+)"', html)
    mks = re.findall(r"mk\('(chart[A-Za-z0-9]+)'", html)
    if sorted(ids) != sorted(mks):
        fails.append(f"图表不对应: 多余容器{set(ids)-set(mks)} 缺少初始化{set(mks)-set(ids)}")
    # 3) JS 语法（node --check，不可用时跳过并提示）
    scripts = re.findall(r'<script>(.*?)</script>', html, re.S)
    main_js = [s for s in scripts if 'function mk' in s]
    if not main_js:
        fails.append("未找到主图表脚本(function mk)")
    else:
        with tempfile.NamedTemporaryFile('w', suffix='.js', delete=False, encoding='utf-8') as f:
            f.write(main_js[0]); tmp = f.name
        try:
            r = subprocess.run(['node', '--check', tmp], capture_output=True, text=True)
            if r.returncode != 0:
                fails.append(f"JS语法错误: {r.stderr.strip()[:300]}")
        except FileNotFoundError:
            print("  [warn] node 不可用，跳过 JS 语法检查")
        finally:
            os.remove(tmp)
    # 4) 敏感词（剥离 script 后扫描正文，避免 JS 注释误报）
    body = re.sub(r'<script[^>]*>.*?</script>', '', html, flags=re.S)
    body = re.sub(r'<style>.*?</style>', '', body, flags=re.S)
    hits = sorted({w for w in BANNED if w in body})
    if hits: fails.append(f"视角禁词命中: {hits}")
    # 5) 隐私扫描
    privacy = []
    if re.search(r'\d{6}\*+\d', body): privacy.append('证件号掩码残留')
    if re.search(r'(?<!\d)1[3-9]\d{9}(?!\d)', body): privacy.append('手机号')
    if re.search(r'姓名[""：:]', body) and '姓名' in body: privacy.append('姓名字段标签')
    if privacy: fails.append(f"隐私风险: {privacy}")
    # 结果
    tag = 'PASS' if not fails else 'FAIL'
    print(f"[{tag}] {path}")
    print(f"       图表 {len(ids)} 个 | 正文可见字符 {len(re.sub(r'<[^>]+>', '', body).replace(chr(10), '').replace(' ', ''))}")
    for f in fails: print(f"       ✗ {f}")
    if not fails:
        print("       ✓ 标签配对 / 图表对应 / JS 语法 / 视角禁词 / 隐私 全部通过")
    return not fails

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print(__doc__); sys.exit(1)
    ok = all(validate(p) for p in sys.argv[1:])
    sys.exit(0 if ok else 1)

# -*- coding: utf-8 -*-
"""校验生成的 docx：OMML 数量、乱码残留、标题结构、公式语义泄漏"""
import sys, os, re, glob, argparse
sys.stdout.reconfigure(encoding='utf-8')
from docx import Document

ap = argparse.ArgumentParser(add_help=False)
ap.add_argument('--out', default=os.environ.get('OUT_DIR', os.getcwd()))
_a, _ = ap.parse_known_args()
OUT = _a.out
LEAK = ['overline', 'widetilde', 'widehat', 'overrightarrow', 'underline',
        'bar', 'hat', 'tilde', 'ddot', 'vec', 'check', 'breve', 'acute',
        'grave', 'frac', 'begin', 'end', 'pmatrix', 'bmatrix', 'aligned',
        'displaystyle', 'qquad', 'quad', 'mathbb', 'mathcal', 'mathrm', 'partial']

for f in sorted(glob.glob(os.path.join(OUT, '*.docx'))):
    d = Document(f)
    xml = d.element.xml
    noml = xml.count('<m:oMath')
    paras = len(d.paragraphs)
    tables = len(d.tables)
    text = '\n'.join(p.text for p in d.paragraphs)
    bad_bs = text.count('\\')
    bad_slash = len(re.findall(r'\)\s*/\s*\(', text))
    h1 = sum(1 for p in d.paragraphs if p.style.name == 'Heading 1')
    h2 = sum(1 for p in d.paragraphs if p.style.name == 'Heading 2')
    h3 = sum(1 for p in d.paragraphs if p.style.name == 'Heading 3')

    # 公式语义泄漏：只统计"独立出现"的可疑词（排除正文里的合法英文词）
    leaks = []
    for w in LEAK:
        for m in re.finditer(r'(?<![A-Za-z])' + w + r'(?![A-Za-z])', text):
            # 忽略标题/正文中合法使用（如 "bar" 出现在英文里）
            ctx = text[max(0, m.start()-30):m.end()+30].replace('\n', ' ')
            if w in ('begin', 'end', 'frac', 'bar', 'hat', 'check', 'vec',
                     'tilde', 'quad', 'displaystyle', 'partial'):
                leaks.append((w, ctx))
    print(os.path.basename(f))
    print('   段落=%d 表格=%d 公式对象=%d  标题: H1=%d H2=%d H3=%d' % (paras, tables, noml, h1, h2, h3))
    print('   残留反斜杠=%d  (a)/(b)伪分式=%d  字数≈%d' % (bad_bs, bad_slash, len(text)))
    if leaks:
        seen = set()
        print('   [疑似公式泄漏 %d 处]' % len(leaks))
        for w, ctx in leaks[:6]:
            if w in seen:
                continue
            seen.add(w)
            print('      %-12s ... %s ...' % (w, ctx[:70]))

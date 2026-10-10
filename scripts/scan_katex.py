# -*- coding: utf-8 -*-
"""扫描 md 中的 LaTeX 公式，找出 KaTeX 不支持 / 高风险的命令"""
import sys, os, re, argparse
sys.stdout.reconfigure(encoding='utf-8')

ap = argparse.ArgumentParser(add_help=False)
ap.add_argument('--build', default=os.environ.get('BUILD_DIR', os.getcwd()))
_a, _ = ap.parse_known_args()
BUILD = _a.build

# KaTeX 不支持或需注意的命令
UNSUPPORTED = [
    r'\\text\{[^}]*[\u4e00-\u9fff]',   # \text{中文} —— KaTeX 可渲染但要检查
    r'\\mathdefault',
    r'\\substack',
    r'\\xrightarrow',
    r'\\xleftarrow',
    r'\\bm\b',
    r'\\begin\{aligned\}',
    r'\\begin\{array\}',
    r'\\begin\{cases\}',
    r'\\begin\{pmatrix\}',
    r'\\begin\{bmatrix\}',
    r'\\begin\{matrix\}',
    r'\\begin\{split\}',
    r'\\over\b',            # \over 老式分数
    r'\\raisebox',
    r'\\phantom',
    r'\\vphantom',
    r'\\hphantom',
    r'\\ooalign',
    r'\\stackrel',
    r'\\overset',
    r'\\underset',
    r'\\underbrace',
    r'\\overbrace',
    r'\\tag\b',
    r'\\label\b',
    r'\\nonumber',
    r'\\notag',
    r'\\hspace',
    r'\\vspace',
    r'\\displaystyle',
    r'\\textstyle',
    r'\\limits',
    r'\\nolimits',
    r'\\qquad',
    r'\\quad',
]

TEX_PATTERNS = [r'\$\$(.+?)\$\$', r'(?<!\$)\$([^$\n]+)\$(?!\$)']

stats = {}
# 例：python scan_katex.py a.md b.md  （只扫描指定文件）
#     不带参数时扫描下方示例文件列表，请换成你自己的 .md
files = [x for x in sys.argv[1:] if x.endswith('.md') and not x.startswith('-')]
if not files:
    files = ['micro_all.md', 'macro.md', 'econ.md']
for f in files:
    p = os.path.join(BUILD, f)
    if not os.path.exists(p): continue
    txt = open(p, encoding='utf-8').read()
    total = 0
    hits = {}
    for pat in TEX_PATTERNS:
        for m in re.finditer(pat, txt, re.S):
            src = m.group(1); total += 1
            for u in UNSUPPORTED:
                if re.search(u, src):
                    hits.setdefault(u, []).append(src[:100].replace('\n', ' '))
    print('===', f, ' 公式数:', total, ' 命中命令种类:', len(hits))
    for u, exs in sorted(hits.items(), key=lambda x: -len(x[1])):
        print('   %-28s x%-4d  例: %s' % (u, len(exs), exs[0][:80]))

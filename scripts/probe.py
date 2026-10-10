# -*- coding: utf-8 -*-
"""扫描 md 中所有公式，找出：
   (a) 转换失败 / 畸形 XML 的条目；
   (b) 『静默语义损坏』—— preprocess 之后出现裸 LaTeX 命令名残片
       （如 \\bar r -> 'bar'、\\overline{X} -> 'overline'）。这类问题不会抛异常，
       但会在 Word 里显示成 'bar' 这样的乱码，必须专门拦截。
"""
import sys, os, re, argparse
sys.stdout.reconfigure(encoding='utf-8')
ap = argparse.ArgumentParser(add_help=False)
ap.add_argument('--build', default=os.environ.get('BUILD_DIR', os.getcwd()))
_a, _ = ap.parse_known_args()
BUILD = _a.build
sys.path.insert(0, BUILD)
from make_docx import latex_to_omml, preprocess, split_aligned
from lxml import etree

# 出现在 preprocess 结果里即视为损坏的裸命令名（区别于合法内容）
LEAK_WORDS = ['overline', 'underline', 'widetilde', 'widehat', 'overrightarrow',
              'bar', 'hat', 'tilde', 'ddot', 'vec', 'check', 'breve', 'acute',
              'grave', 'frac', 'begin', 'end', 'left', 'right', 'text', 'mathrm',
              'mathbb', 'mathcal', 'displaystyle', 'quad', 'qquad']


def leaks(tex):
    """返回 preprocess 后泄露的裸命令名列表；只检查不含反斜杠的独立单词。"""
    pre = preprocess(tex)
    found = []
    for w in LEAK_WORDS:
        # 匹配"独立的"该单词：前不为字母/反斜杠，后不为字母
        if re.search(r'(?<![\\A-Za-z])' + w + r'(?![A-Za-z])', pre):
            found.append(w)
    return pre, found


bad = []
leak = []
total = 0
# 例：python probe.py a.md b.md  （只检查指定文件）
#     不带参数时检查下方示例文件列表，请换成你自己的 .md
files = [x for x in sys.argv[1:] if x.endswith('.md') and not x.startswith('-')]
if not files:
    files = ['micro_all.md', 'macro.md', 'econ.md']
for f in files:
    p = os.path.join(BUILD, f)
    if not os.path.exists(p):
        continue
    txt = open(p, encoding='utf-8').read()
    pats = [r'\$\$(.+?)\$\$', r'(?<!\$)\$([^$\n]+)\$(?!\$)']
    for pat in pats:
        for m in re.finditer(pat, txt, re.S):
            src = m.group(1)
            total += 1
            # (a) XML 合法性
            for ox in latex_to_omml(src):
                if ox is None:
                    continue
                try:
                    etree.fromstring(ox.encode('utf-8'))
                except Exception as e:
                    bad.append((f, src[:120], str(e)[:90]))
                    break
            # (b) 语义泄漏
            pre, found = leaks(src)
            if found:
                leak.append((f, src[:110], pre[:110], found))

print('total formulas:', total)
print('xml-bad :', len(bad))
print('leak-bad:', len(leak))
for b in bad[:40]:
    print('--- XML', b[0])
    print('   TEX:', b[1].replace('\n', ' '))
    print('   ERR:', b[2])
for b in leak[:40]:
    print('--- LEAK', b[0], b[3])
    print('   TEX :', b[1].replace('\n', ' '))
    print('   PRE :', b[2].replace('\n', ' '))

# -*- coding: utf-8 -*-
"""Markdown -> Word (.docx)  |  v3  精美排版 + 原生 OMML 公式 + 封面 + 目录

支持扩展语法（在标准 Markdown 之上）：
  ::: box 标题        ... :::      -> 彩色批注框（直觉/易错/必背）
  ::: example 标题    ... :::      -> 例题框（浅黄底 + 左侧粗边）
  ::: answer 标题     ... :::      -> 解答框（浅绿底）
  ::: warn 标题       ... :::      -> 警示框（浅红底）
  > 行首引用                        -> 楷体批注
"""
import sys, os, re, argparse

sys.stdout.reconfigure(encoding='utf-8')


def _resolve_paths():
    """BUILD / OUT 可配置，解析优先级：
       命令行 --build / --out  >  环境变量 BUILD_DIR / OUT_DIR  >  当前工作目录
    """
    ap = argparse.ArgumentParser(add_help=False)
    ap.add_argument('--build', default=os.environ.get('BUILD_DIR', os.getcwd()))
    ap.add_argument('--out',   default=os.environ.get('OUT_DIR',   os.getcwd()))
    a, _ = ap.parse_known_args()
    return a.build, a.out
BUILD, OUT = _resolve_paths()
os.makedirs(OUT, exist_ok=True)

import latex2mathml.converter as lc
import mathml2omml
from lxml import etree
from docx import Document
from docx.shared import Pt, Cm, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import parse_xml, OxmlElement

MNS = 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"'

# ======================= 配色 =======================
C_NAVY = '1F3864'
C_BLUE = '2E5496'
C_GRAY = '404040'
C_INK  = '1A1A1A'
C_BROWN = '9C5700'
BOX = {
    'box':     ('EAF1FB', 'C5D8F1', '1F3864'),   # 直觉/说明：淡蓝
    'example': ('FFF8E6', 'F0D9A0', '9C6500'),   # 例题：淡黄
    'answer':  ('EDF7EE', 'B7DFBC', '1E6B34'),   # 解答：淡绿
    'warn':    ('FDEEEE', 'F2C0C0', 'A32020'),   # 易错：淡红
    'note':    ('F5F2FA', 'D6CCEC', '5B3E96'),   # 必背：淡紫
}


# ======================= LaTeX 预处理 =======================
DROP_ARG = ['underbrace', 'overbrace', 'underset', 'overset', 'stackrel',
            'text', 'textrm', 'mathrm', 'mbox', 'textbf', 'textit', 'emph',
            'operatorname', 'label', 'tag', 'nonumber', 'notag', 'bm',
            'boldsymbol', 'mathbb', 'mathcal', 'widehat', 'widetilde',
            'textnormal', 'textup', 'substack']


def _match_brace(t, i):
    depth, j = 0, i
    while j < len(t):
        if t[j] == '{':
            depth += 1
        elif t[j] == '}':
            depth -= 1
            if depth == 0:
                return j
        j += 1
    return -1


def _strip_cmds(t):
    changed, rounds = True, 0
    while changed and rounds < 8:
        changed, rounds = False, rounds + 1
        for cmd in DROP_ARG:
            pat = '\\' + cmd
            pos = t.find(pat)
            while pos >= 0:
                after = pos + len(pat)
                if after < len(t) and t[after].isalpha():
                    pos = t.find(pat, after); continue
                k = after
                while k < len(t) and t[k] == ' ':
                    k += 1
                if k < len(t) and t[k] == '{':
                    e = _match_brace(t, k)
                    if e < 0:
                        pos = t.find(pat, after); continue
                    inner = t[k + 1:e]
                    rest_i = e + 1
                    m2 = re.match(r'\s*[_^]\s*\{', t[rest_i:])
                    if m2:
                        s2 = rest_i + m2.end() - 1
                        e2 = _match_brace(t, s2)
                        if e2 > 0:
                            rest_i = e2 + 1
                    t = t[:pos] + inner + t[rest_i:]
                    changed = True
                else:
                    t = t[:pos] + t[after:]
                    changed = True
                pos = t.find(pat, pos)
    return t


def preprocess(tex):
    t = tex.strip()
    t = re.sub(r'\\\[|\\\]|\\\(|\\\)', '', t)
    t = _strip_cmds(t)
    # 所有 accent 命令统一处理：
    #   \overline{X}_i / \hat{Y}_{it} / \bar Y_i  ->  X_i / Y_{it} / Y_i
    #   （mathml2omml 对"带上下标的重音符号"会生成畸形 groupChr XML）
    # 注意：ACC 本身是第 1 个捕获组，因此"内容"组从 \2 开始编号；
    # 早期版本误用 \1 输出 ACC 匹配到的命令名（如 'bar'），造成裸文本残片。
    ACC = r'(widetilde|widehat|overrightarrow|overline|underline|check|breve|acute|grave|tilde|ddot|dot|vec|hat|bar)'
    # (1) accent{...} 后紧跟 _{} 或 ^{} -> 直接丢弃 accent，保留 {内容} 与脚本
    t = re.sub(r'\\' + ACC + r'\s*\{([^{}]*)\}\s*([_^]\s*\{[^{}]*\})', r'\2\3', t)
    # (2) accent{...} 其余情形 -> 丢弃 accent（保留内容）
    t = re.sub(r'\\' + ACC + r'\s*\{([^{}]*)\}', r'\2', t)
    # (3) accent X 后紧跟 _{} / ^{}   （X 为单字母）
    t = re.sub(r'\\' + ACC + r'\s+([A-Za-z])\s*([_^]\s*\{[^{}]*\})', r'\2\3', t)
    # (3b) accent X 后紧跟 _c / ^c     （X 为单字母，脚本不带花括号）
    t = re.sub(r'\\' + ACC + r'\s+([A-Za-z])\s*([_^]\s*[A-Za-z0-9])', r'\2\3', t)
    # (4) accent X 其余情形（X 为单字母，后面不得再跟字母，避免吃掉命令名）
    t = re.sub(r'\\' + ACC + r'\s+([A-Za-z])(?![A-Za-z])', r'\2', t)
    # (5) 兜底：任何残留的 accent 命令（含后接非 ASCII 字母如 \bar\mu）一律删除，
    #     只保留其后的内容，绝不留下 'bar' 这类裸文本残片
    t = re.sub(r'\\' + ACC + r'(?![a-zA-Z])', '', t)

    # 安全网：清除任何以反斜杠开头的 accent 残片
    t = re.sub(r'\\(overline|underline|widetilde|widehat|overrightarrow|tilde|ddot|dot|vec|hat|bar|check|breve|acute|grave)(?![a-zA-Z])', '', t)
    # 剔除 mathml2omml 无法处理的符号
    for bad in (r'\checkmark', r'\varnothing', r'\blacksquare', r'\square',
                r'\bigstar', r'\degree', r'\triangle', r'\blacktriangle',
                r'\S', r'\P'):
        t = t.replace(bad, '')
    t = t.replace('\\%', r'\%').replace('%', r'\%')
    # 先把行分隔符 \\ 占位保护起来，避免后续"细空格 \ "替换把它吃掉
    t = t.replace('\\\\', '\x00ROWBREAK\x00')
    for a in (r'\quad', r'\qquad', r'\,', r'\;', r'\:', r'\!', r'\ ', r'\thinspace'):
        t = t.replace(a, ' ')
    for a in (r'\displaystyle', r'\limits', r'\nonumber', r'\notag',
              r'\medspace', r'\centering', r'\small', r'\footnotesize'):
        t = t.replace(a, '')
    t = re.sub(r'\\(big|Big|bigg|Bigg|left|right|bigl|bigr|Bigl|Bigr)(?![a-zA-Z])', '', t)
    t = re.sub(r'\\(math)?rm(?![a-zA-Z])', '', t)
    # 说明：mathml2omml 原生支持 array/matrix/pmatrix/bmatrix/cases，且能保留
    #       行列结构；唯独 aligned 会生成畸形 XML。
    #       另有一个硬限制：matrix 节点**不能直接挂上下标**（报
    #       "Too large number of children for msub"）。解决办法是把
    #       pmatrix/bmatrix 降级为 matrix 并用 \left...\right 包裹，
    #       这样下标就挂在"单一定界符组"上，可以正常渲染。
    t = t.replace(r'\begin{split}', r'\begin{array}{c}').replace(r'\end{split}', r'\end{array}')
    t = t.replace(r'\begin{gather}', r'\begin{array}{c}').replace(r'\end{gather}', r'\end{array}')
    t = re.sub(r'\\begin\{aligned\}', r'\\begin{array}{c}', t)
    t = re.sub(r'\\end\{aligned\}', r'\\end{array}', t)
    if 'array' in t:
        t = re.sub(r'\\begin\{array\}\{[clr|]*\}', r'\\begin{array}{c}', t)
    # 关键修复：只对"后面紧跟 _ 或 ^ 脚本"的 pmatrix/bmatrix 做定界符改写
    def _fix_matrix_scripts(s):
        out, i = s, 0
        pat = re.compile(r'\\begin\{(p|b)matrix\}(.*?)\\end\{\1matrix\}(\s*[_^])', re.S)
        while True:
            m = pat.search(out)
            if not m:
                break
            kind, inner, tail = m.group(1), m.group(2), m.group(3)
            lo, hi = ('(', ')') if kind == 'p' else ('[', ']')
            repl = (r'\left' + lo + r'\begin{matrix}' + inner + r'\end{matrix}'
                    + r'\right' + hi + tail)
            out = out[:m.start()] + repl + out[m.end():]
        return out
    t = _fix_matrix_scripts(t)
    t = t.replace(r'\R', r'\mathbb{R}').replace(r'\E', r'\mathbb{E}')
    t = t.replace(r'\N', r'\mathbb{N}')
    t = re.sub(r'\\1\{([^{}]*)\}', r'\1', t)
    t = re.sub(r'\\(d)?frac\s+(\w)\s+(\w)', r'\\frac{\2}{\3}', t)
    # 行距修饰 \\[6pt] / \\[-2pt] 等：mathml2omml 不支持，一律清除
    t = re.sub(r'\\\\\s*\[\s*-?\d+(\.\d+)?\s*(pt|ex|em|mu|mm|cm)\s*\]', r'\\\\', t)
    # 还原被保护的行分隔符
    t = t.replace('\x00ROWBREAK\x00', '\\\\')
    # 修复 "\partialc" "\nablak" 这类：命令名与紧跟的单字母粘连，补空格
    t = re.sub(r'(\\(?:partial|nabla|mathrm|mathbb|mathcal|operatorname))(?=[A-Za-z])',
               r'\1 ', t)
    t = t.replace('\\\\', ' \\\\ ')
    t = t.replace('\n', ' ').strip()
    t = re.sub(r'\s+', ' ', t)
    return t


def split_aligned(tex):
    """多行环境处理。

    早期版本把 aligned 拆成多个单行公式分别渲染（因为 mathml2omml 对 aligned
    生成畸形 XML）。现在 preprocess 已把 aligned 统一改写为 array，而
    mathml2omml 能原生、正确地渲染 array（保留换行与分列），因此这里**不再拆分**，
    直接整段交给渲染器。仅当整段渲染失败时，才退化为按行拆分。
    """
    return [tex]


_omml_cache = {}


def _validate(xml):
    try:
        etree.fromstring(xml.encode('utf-8'))
        return True
    except Exception:
        return False


def latex_to_omml(tex):
    try:
        tex = preprocess(tex)
    except Exception:
        return []
    if not tex:
        return []
    out = []
    for ln in split_aligned(tex):
        if ln in _omml_cache:
            out.append(_omml_cache[ln]); continue
        val = None
        try:
            omml = mathml2omml.convert(lc.convert(ln))
            omml = omml.replace('<m:oMathPara>', '').replace('</m:oMathPara>', '')
            omml = omml.replace('<m:oMath>', '<m:oMath ' + MNS + '>')
            if _validate(omml):
                val = omml
        except Exception:
            val = None
        _omml_cache[ln] = val
        out.append(val)
    return out


def add_omml(paragraph, omml_xml):
    paragraph._p.append(parse_xml(omml_xml))


# ======================= 降级纯文本 =======================
GREEK = {'alpha': 'α', 'beta': 'β', 'gamma': 'γ', 'delta': 'δ', 'epsilon': 'ε',
         'varepsilon': 'ε', 'zeta': 'ζ', 'eta': 'η', 'theta': 'θ', 'iota': 'ι',
         'kappa': 'κ', 'lambda': 'λ', 'mu': 'μ', 'nu': 'ν', 'xi': 'ξ', 'pi': 'π',
         'rho': 'ρ', 'sigma': 'σ', 'tau': 'τ', 'phi': 'φ', 'varphi': 'φ',
         'chi': 'χ', 'psi': 'ψ', 'omega': 'ω', 'Gamma': 'Γ', 'Delta': 'Δ',
         'Theta': 'Θ', 'Lambda': 'Λ', 'Xi': 'Ξ', 'Pi': 'Π', 'Sigma': 'Σ',
         'Phi': 'Φ', 'Psi': 'Ψ', 'Omega': 'Ω'}
SYMBOLS = [('\\sum', 'Σ'), ('\\prod', 'Π'), ('\\int', '∫'), ('\\infty', '∞'),
           ('\\partial', '∂'), ('\\geq', '≥'), ('\\ge', '≥'), ('\\leq', '≤'),
           ('\\le', '≤'), ('\\neq', '≠'), ('\\ne', '≠'), ('\\approx', '≈'),
           ('\\equiv', '≡'), ('\\cdot', '·'), ('\\times', '×'), ('\\div', '÷'),
           ('\\Rightarrow', '⇒'), ('\\rightarrow', '→'), ('\\to', '→'),
           ('\\leftarrow', '←'), ('\\in', '∈'), ('\\subset', '⊂'), ('\\cup', '∪'),
           ('\\cap', '∩'), ('\\forall', '∀'), ('\\exists', '∃'), ('\\dots', '…'),
           ('\\ldots', '…'), ('\\cdots', '⋯'), ('\\max', 'max'), ('\\min', 'min'),
           ('\\lim', 'lim'), ('\\ln', 'ln'), ('\\log', 'log'), ('\\exp', 'exp'),
           ('\\prime', '′'), ('\\top', '⊤'), ('\\perp', '⊥'), ('\\pm', '±'),
           ('\\propto', '∝'), ('\\sim', '∼')]


def clean_math(s):
    s = _strip_cmds(s)
    for _ in range(6):
        s = re.sub(r'\\d?frac\s*\{([^{}]*)\}\s*\{([^{}]*)\}', r'(\1)/(\2)', s)
    # accent 命令：先剥掉，再把 \overline{X} 之类化为 X
    for acc in ('overline', 'widehat', 'widetilde', 'overrightarrow', 'underline'):
        s = re.sub(r'\\' + acc + r'\s*\{([^{}]*)\}', r'\1', s)
    for g, v in GREEK.items():
        s = s.replace('\\' + g, v)
    for a, b in SYMBOLS:
        s = s.replace(a, b)
    s = s.replace('\\{', '{').replace('\\}', '}')
    s = re.sub(r'\\(operatorname|mathrm|text|mbox|textbf|textit)\s*\{([^{}]*)\}', r'\2', s)
    s = s.replace('\\', '')
    s = re.sub(r'_\{([^{}]*)\}', r'_\1', s)
    s = re.sub(r'\^\{([^{}]*)\}', r'^\1', s)
    return s.strip()


def _sanitize_outside_math(text):
    """清除 $...$ 之外的裸 LaTeX 命令（作者笔误，如 '\\quad⇒\\quad'）。

    这类残片若不清除，会以源码形式出现在文档正文里。做法是把所有数学片段
    （$$...$$ 与 $...$）替换成占位符，清理剩余文本里的 \\命令，再还原。
    """
    store = []

    def _stash(m):
        store.append(m.group(0))
        return '\x01M%d\x01' % (len(store) - 1)

    t = re.sub(r'\$\$.+?\$\$', _stash, text, flags=re.S)
    t = re.sub(r'\$[^$\n]+\$', _stash, t)
    # Markdown 转义残留：\* \_ \$ \& \# \[ \] 等，去掉反斜杠保留字符本身
    t = re.sub(r'\\([*_$&#\[\]()~`>+\-=|{}.!])', r'\1', t)
    # 常见裸命令：\quad \qquad \, \; \: \! \ 以及 \text{...} 之类
    t = re.sub(r'\\(quad|qquad|thinspace|medspace|enspace|,|;|:|!)\s?', ' ', t)
    t = re.sub(r'\\[A-Za-z]+\s*', '', t)
    t = re.sub(r'\x01M(\d+)\x01', lambda m: store[int(m.group(1))], t)
    return t


def strip_inline(text):
    text = _sanitize_outside_math(text)
    t = re.sub(r'\*\*(.+?)\*\*', r'\1', text)
    t = re.sub(r'\*\*(.+?)\*\*', r'\1', t)
    t = re.sub(r'`([^`]+)`', r'\1', t)
    t = re.sub(r'\$\$(.+?)\$\$', lambda m: clean_math(m.group(1)), t, flags=re.S)
    t = re.sub(r'\$([^$]+)\$', lambda m: clean_math(m.group(1)), t)
    return t


INLINE_RE = re.compile(r'(\*\*.+?\*\*|`[^`]+`|\$\$[^$]+\$\$|\$[^$]+\$)')


def add_inline(par, text, size=10.5, cn=None, color=None, bold=False):
    """加入 runs，处理加粗、行内代码、行内公式(OMML)。

    关键点：粗体片段 **...** 内部可能**同时包含**行内公式与中文（例如
    「**技巧 1　$x_t\\approx \\bar x$ 的定义**」）。早期版本把整个粗体片段
    当成一个纯文本 run 直接输出，导致其中的 $...$ 原样显示为 LaTeX 源码。
    现改为：进入粗体片段时递归调用自身（bold=True），使内部的公式照常
    渲染成 OMML、中文照常套用字体。

    另外对"数学片段之外的裸 LaTeX 命令"（如正文里的 \\quad）做一次清理，
    避免 LaTeX 源码出现在 Word 正文中。
    """
    if not bold:
        text = _sanitize_outside_math(text)
    pos = 0
    for m in INLINE_RE.finditer(text):
        if m.start() > pos:
            r = par.add_run(text[pos:m.start()]); r.font.size = Pt(size)
            r.bold = bold
            if cn: set_cn(r, cn)
            if color: r.font.color.rgb = RGBColor.from_string(color)
        tok = m.group(0)
        if tok.startswith('**'):
            # 递归处理粗体片段内部内容（可能含公式与中文）
            add_inline(par, tok[2:-2], size=size, cn=cn, color=color, bold=True)
        elif tok.startswith('`'):
            r = par.add_run(tok[1:-1]); r.font.name = 'Consolas'
            r.font.size = Pt(size - 1)
            r.font.color.rgb = RGBColor.from_string('A32020')
        else:
            inner = tok[2:-2] if tok.startswith('$$') else tok[1:-1]
            oml = latex_to_omml(inner)
            ok = False
            for ox in oml:
                if ox:
                    add_omml(par, ox); ok = True
            if not ok:
                r = par.add_run(clean_math(inner)); r.italic = True
                r.font.size = Pt(size)
                r.font.color.rgb = RGBColor.from_string(C_BROWN)
        pos = m.end()
    if pos < len(text):
        r = par.add_run(text[pos:]); r.font.size = Pt(size)
        r.bold = bold
        if cn: set_cn(r, cn)
        if color: r.font.color.rgb = RGBColor.from_string(color)


def set_cn(run, cn='宋体', en='Times New Roman'):
    run.font.name = en
    try:
        run._element.rPr.rFonts.set(qn('w:eastAsia'), cn)
    except Exception:
        pass


def shade(el, fill):
    """给段落或单元格加底纹"""
    pr = el._p.get_or_add_pPr() if hasattr(el, '_p') else el._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear'); shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill)
    pr.append(shd)


def border(el, left=None, top=None, bottom=None, right=None):
    pr = el._p.get_or_add_pPr() if hasattr(el, '_p') else el._tc.get_or_add_tcPr()
    bd = OxmlElement('w:pBdr')
    for tag, spec in (('top', top), ('left', left), ('bottom', bottom), ('right', right)):
        if not spec:
            continue
        e = OxmlElement('w:' + tag)
        e.set(qn('w:val'), 'single')
        e.set(qn('w:sz'), str(spec[0]))
        e.set(qn('w:space'), '4')
        e.set(qn('w:color'), spec[1])
        bd.append(e)
    pr.append(bd)


# ======================= 文档骨架 =======================
def new_doc():
    doc = Document()
    st = doc.styles['Normal']
    st.font.size = Pt(10.5); st.font.name = 'Times New Roman'
    st.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    st.paragraph_format.space_after = Pt(4)
    st.paragraph_format.line_spacing = 1.45

    hspec = [('Heading 1', 17, C_NAVY, '微软雅黑', 20, 10, True),
             ('Heading 2', 14, C_BLUE, '微软雅黑', 15, 7, False),
             ('Heading 3', 12, C_BLUE, '微软雅黑', 12, 5, False),
             ('Heading 4', 11, C_GRAY, '微软雅黑', 9, 4, False)]
    for nm, sz, col, fnt, sb, sa, pb in hspec:
        s = doc.styles[nm]
        s.font.size = Pt(sz); s.font.bold = True
        s.font.color.rgb = RGBColor.from_string(col); s.font.name = 'Times New Roman'
        s.element.rPr.rFonts.set(qn('w:eastAsia'), fnt)
        s.paragraph_format.space_before = Pt(sb)
        s.paragraph_format.space_after = Pt(sa)
        s.paragraph_format.keep_with_next = True
        if pb:
            s.paragraph_format.page_break_before = True

    for sec in doc.sections:
        sec.left_margin = Cm(2.2); sec.right_margin = Cm(2.2)
        sec.top_margin = Cm(2.1); sec.bottom_margin = Cm(1.9)
    return doc


def cover_page(doc, title, subtitle, meta):
    """封面页：色块 + 标题 + 信息表"""
    for _ in range(3):
        doc.add_paragraph()
    t = doc.add_table(rows=1, cols=1); t.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = t.rows[0].cells[0]
    shade(cell, C_NAVY)
    cell.width = Cm(15)
    p = cell.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(24); p.paragraph_format.space_after = Pt(6)
    r = p.add_run(title); r.font.size = Pt(34); r.bold = True
    r.font.color.rgb = RGBColor.from_string('FFFFFF'); set_cn(r, '微软雅黑')
    p2 = cell.add_paragraph(); p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.paragraph_format.space_after = Pt(24)
    r2 = p2.add_run(subtitle); r2.font.size = Pt(13)
    r2.font.color.rgb = RGBColor.from_string('C9D8F0'); set_cn(r2, '微软雅黑')
    doc.add_paragraph()
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run('博 士 资 格 考 试 （ 博 资 考 ）'); r.font.size = Pt(15)
    r.bold = True; r.font.color.rgb = RGBColor.from_string(C_BLUE); set_cn(r, '微软雅黑')
    doc.add_paragraph()
    for line in meta:
        p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run(line); r.font.size = Pt(11.5)
        r.font.color.rgb = RGBColor.from_string('555555'); set_cn(r, '微软雅黑')


def add_toc(doc):
    doc.add_paragraph()
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run('目　　录'); r.font.size = Pt(18); r.bold = True
    r.font.color.rgb = RGBColor.from_string(C_NAVY); set_cn(r, '微软雅黑')
    p = doc.add_paragraph()
    run = p.add_run()
    f1 = OxmlElement('w:fldChar'); f1.set(qn('w:fldCharType'), 'begin')
    it = OxmlElement('w:instrText'); it.set(qn('xml:space'), 'preserve')
    it.text = r'TOC \o "1-2" \h \z \u'
    f2 = OxmlElement('w:fldChar'); f2.set(qn('w:fldCharType'), 'separate')
    t = OxmlElement('w:t'); t.text = '【 请按 Ctrl+A 全选，再按 F9 选择“更新整个目录” 】'
    f3 = OxmlElement('w:fldChar'); f3.set(qn('w:fldCharType'), 'end')
    for x in (f1, it, f2, t, f3):
        run._r.append(x)
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run('（目录为 Word 域，首次打开需更新一次）')
    r.font.size = Pt(9); r.italic = True
    r.font.color.rgb = RGBColor.from_string('999999'); set_cn(r, '楷体')


def add_footer(doc, text):
    for sec in doc.sections:
        p = sec.footer.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r0 = p.add_run(text + '　·　第 ')
        r0.font.size = Pt(8.5); r0.font.color.rgb = RGBColor.from_string('888888')
        set_cn(r0, '微软雅黑')
        run = p.add_run()
        f1 = OxmlElement('w:fldChar'); f1.set(qn('w:fldCharType'), 'begin')
        it = OxmlElement('w:instrText'); it.set(qn('xml:space'), 'preserve'); it.text = 'PAGE'
        f2 = OxmlElement('w:fldChar'); f2.set(qn('w:fldCharType'), 'end')
        for x in (f1, it, f2):
            run._r.append(x)
        r1 = p.add_run(' 页'); r1.font.size = Pt(8.5)
        r1.font.color.rgb = RGBColor.from_string('888888'); set_cn(r1, '微软雅黑')


# ======================= 构建 =======================
def render_table(doc, rows):
    raw = [[c.strip() for c in r.strip('|').split('|')] for r in rows]
    ncol = max(len(r) for r in raw)
    tb = doc.add_table(rows=0, cols=ncol)
    tb.style = 'Table Grid'; tb.alignment = WD_TABLE_ALIGNMENT.CENTER
    tb.autofit = True
    for ri, r in enumerate(raw):
        row = tb.add_row()
        for ci in range(ncol):
            txt = r[ci] if ci < len(r) else ''
            cell = row.cells[ci]
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.line_spacing = 1.12
            add_inline(p, txt, size=8.8)
            for run in p.runs:
                run.bold = (ri == 0)
                if ri == 0:
                    run.font.color.rgb = RGBColor.from_string(C_NAVY)
            if ri == 0:
                shade(cell, C_NAVY)
                for run in p.runs:
                    run.font.color.rgb = RGBColor.from_string('FFFFFF')
            elif ri % 2 == 0:
                shade(cell, 'F6F8FC')
    doc.add_paragraph().paragraph_format.space_after = Pt(2)


def render_box(doc, kind, title, body_lines):
    """彩色批注框：单列表格 + 底纹 + 左边粗条"""
    fill, edge, dark = BOX.get(kind, BOX['box'])
    tb = doc.add_table(rows=1, cols=1)
    tb.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = tb.rows[0].cells[0]
    shade(cell, fill)
    try:
        tcPr = cell._tc.get_or_add_tcPr()
        borders = OxmlElement('w:tcBorders')
        for tag, sz in (('left', 24), ('top', 4), ('bottom', 4), ('right', 4)):
            e = OxmlElement('w:' + tag)
            e.set(qn('w:val'), 'single'); e.set(qn('w:sz'), str(sz))
            e.set(qn('w:space'), '0')
            e.set(qn('w:color'), edge if tag == 'left' else edge)
            borders.append(e)
        tcPr.append(borders)
    except Exception:
        pass
    first = cell.paragraphs[0]
    first.paragraph_format.space_before = Pt(3)
    first.paragraph_format.space_after = Pt(2)
    if title:
        r = first.add_run(title); r.bold = True; r.font.size = Pt(10.5)
        r.font.color.rgb = RGBColor.from_string(dark); set_cn(r, '微软雅黑')
    for i, ln in enumerate(body_lines):
        p = cell.paragraphs[0] if (i == 0 and not title) else cell.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = 1.35
        add_inline(p, ln, size=10, color=dark)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)


def build(md_path, docx_path, title, subtitle, meta, footer):
    doc = new_doc()
    cover_page(doc, title, subtitle, meta)
    doc.add_page_break()
    add_toc(doc)

    lines = open(md_path, encoding='utf-8').read().split('\n')
    i, n = 0, len(lines)
    fails = []
    first_h1 = True

    def flush_para(buf):
        if not buf:
            return
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(5)
        add_inline(p, ' '.join(buf))

    buf = []
    while i < n:
        ln = lines[i]; s = ln.strip()

        if not s:
            flush_para(buf); buf = []
            i += 1; continue

        # ---- ::: 框 ----
        m = re.match(r'^:::\s*(\w+)\s*(.*)$', s)
        if m:
            flush_para(buf); buf = []
            kind = m.group(1).lower(); btitle = m.group(2).strip()
            body = []
            j = i + 1
            while j < n and not lines[j].strip().startswith(':::'):
                body.append(lines[j].rstrip())
                j += 1
            render_box(doc, kind, btitle, [b for b in body])
            i = j + 1; continue

        # ---- 表格 ----
        if s.startswith('|') and i + 1 < n and re.match(r'^\s*\|[\s:\-|]+\|\s*$', lines[i + 1]):
            flush_para(buf); buf = []
            rows = [s]; j = i + 2
            while j < n and lines[j].strip().startswith('|'):
                rows.append(lines[j].strip()); j += 1
            render_table(doc, rows)
            i = j; continue

        # ---- 代码块 ----
        if s.startswith('```'):
            flush_para(buf); buf = []
            j = i + 1; cb = []
            while j < n and not lines[j].strip().startswith('```'):
                cb.append(lines[j]); j += 1
            for b in cb:
                p = doc.add_paragraph()
                p.paragraph_format.left_indent = Cm(0.7)
                p.paragraph_format.space_after = Pt(0)
                p.paragraph_format.line_spacing = 1.1
                r = p.add_run(b); r.font.name = 'Consolas'; r.font.size = Pt(9)
                r.font.color.rgb = RGBColor.from_string('2F4F4F')
            doc.add_paragraph().paragraph_format.space_after = Pt(2)
            i = j + 1; continue

        # ---- 独立公式 ----
        if s.startswith('$$'):
            flush_para(buf); buf = []
            body = s
            if s.count('$$') < 2:
                j = i + 1
                while j < n and '$$' not in lines[j]:
                    body += ' ' + lines[j].strip(); j += 1
                if j < n:
                    body += ' ' + lines[j].strip(); j += 1
                i = j
            else:
                i += 1
            inner = body.replace('$$', ' ')
            oml = latex_to_omml(inner)
            if not any(oml):
                fails.append(inner[:80])
            for ox in oml:
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.space_before = Pt(5)
                p.paragraph_format.space_after = Pt(6)
                shade(p, 'FBFCFE')
                if ox:
                    add_omml(p, ox)
                else:
                    r = p.add_run(clean_math(inner)); r.italic = True
                    r.font.size = Pt(10.5)
                    r.font.color.rgb = RGBColor.from_string(C_BROWN)
            continue

        # ---- 标题 ----
        m = re.match(r'^(#{1,4})\s+(.*)$', s)
        if m:
            flush_para(buf); buf = []
            lvl = min(len(m.group(1)), 4); txt = m.group(2).strip()
            if lvl == 1:
                if first_h1:
                    first_h1 = False
                    doc.paragraphs[-1].paragraph_format.page_break_before = False
                else:
                    doc.add_page_break()
            p = doc.add_heading(level=lvl)
            size = {1: 17, 2: 14, 3: 12, 4: 11}[lvl]
            add_inline(p, txt, size=size)
            col = {1: C_NAVY, 2: C_BLUE, 3: C_BLUE, 4: C_GRAY}[lvl]
            for r_ in p.runs:
                r_.bold = True
                r_.font.size = Pt(size)
                r_.font.color.rgb = RGBColor.from_string(col)
                set_cn(r_, '微软雅黑')
            if lvl in (1, 2):
                border(p, bottom=(6 if lvl == 1 else 4, 'C5D8F1'))
            i += 1; continue

        if s in ('---', '***', '___'):
            flush_para(buf); buf = []
            p = doc.add_paragraph(); p.paragraph_format.space_after = Pt(3)
            border(p, bottom=(4, 'D9D9D9'))
            i += 1; continue

        # ---- 引用框 ----
        if s.startswith('>'):
            flush_para(buf); buf = []
            txt = s.lstrip('> ').strip()
            if txt.startswith('【') or txt.startswith('💡') or txt.startswith('⚠'):
                kind = 'warn' if '⚠' in txt[:4] else ('note' if '📌' in txt[:4] else 'box')
                render_box(doc, kind, '', [txt])
            else:
                p = doc.add_paragraph()
                p.paragraph_format.left_indent = Cm(0.5)
                p.paragraph_format.space_before = Pt(3)
                p.paragraph_format.space_after = Pt(5)
                add_inline(p, txt, size=10, cn='楷体', color=C_BROWN)
                border(p, left=(18, 'D9B36C'))
            i += 1; continue

        # ---- 列表 ----
        m = re.match(r'^(\s*)([-*+]|\d+[.)])\s+(.*)$', ln)
        if m:
            flush_para(buf); buf = []
            indent = len(m.group(1))
            ordered = not re.match(r'[-*+]$', m.group(2))
            style = 'List Number' if ordered else 'List Bullet'
            p = doc.add_paragraph(style=style)
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.line_spacing = 1.35
            if indent >= 2:
                p.paragraph_format.left_indent = Cm(1.5 + 0.6)
            add_inline(p, m.group(3))
            i += 1; continue

        buf.append(s)
        i += 1

    flush_para(buf)
    add_footer(doc, footer)
    doc.save(docx_path)
    if fails:
        print('   [降级纯文本公式 %d 条]' % len(fails))
        for f in fails[:8]:
            print('     -', f)
    return docx_path


if __name__ == '__main__':
    # 单文件模式：python make_docx.py input.md output.docx [标题]
    #   不带位置参数时，跑下方 JOBS 示例（把文件名换成你自己的 .md）
    pos = [x for x in sys.argv[1:] if not x.startswith('--')]
    if len(pos) >= 2:
        src, out = pos[0], pos[1]
        title = pos[2] if len(pos) > 2 else os.path.splitext(os.path.basename(src))[0]
        print('saved:', build(src, out, title, title, ['公式排版助手 生成'], title))
    else:
        jobs = [
            ('micro_all.md', '高级微观经济学_博资考复习全书.docx',
             '高级微观经济学', 'MWG《微观经济理论》第 1–14 章 · 考点分级详解',
             ['对外经济贸易大学 · 金融学院', '博士生资格考试 · 2026 年 12 月'],
             '高级微观经济学'),
            ('macro.md', '高级宏观经济学_博资考复习全书.docx',
             '高级宏观经济学', 'Romer《高级宏观经济学》· 动态优化 / 消费 / 投资 / DSGE',
             ['对外经济贸易大学 · 金融学院', '博士生资格考试 · 2026 年 12 月'],
             '高级宏观经济学'),
            ('econ.md', '高级计量经济学_博资考复习全书.docx',
             '高级计量经济学', 'OLS 与标准误 · 因果推断 · 离散选择 · 样本选择 · 面板',
             ['对外经济贸易大学 · 金融学院', '博士生资格考试 · 2026 年 12 月'],
             '高级计量经济学'),
        ]
        for src, name, t1, t2, meta, foot in jobs:
            p = os.path.join(BUILD, src)
            if not os.path.exists(p):
                print('skip (missing):', src); continue
            print('saved:', build(p, os.path.join(OUT, name), t1, t2, meta, foot))

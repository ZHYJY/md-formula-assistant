# -*- coding: utf-8 -*-
"""Markdown -> KaTeX HTML  (三科复习全书)

特点:
  * 公式用 KaTeX CDN 渲染（\\[...\\] 行间大号居中, \\(...\\) 行内）
  * 左侧自动目录, 右侧正文, 打印友好
  * 代码块 / 表格 / 引用框 / 提示框样式
  * 中文排版优化
"""
import sys, os, re, html as ihtml, argparse

sys.stdout.reconfigure(encoding='utf-8')


def _resolve_paths():
    """BUILD / OUT 可配置，解析优先级同 make_docx.py：
       命令行 --build/--out  >  环境变量 BUILD_DIR/OUT_DIR  >  当前工作目录
    """
    ap = argparse.ArgumentParser(add_help=False)
    ap.add_argument('--build', default=os.environ.get('BUILD_DIR', os.getcwd()))
    ap.add_argument('--out',   default=os.environ.get('OUT_DIR',   os.getcwd()))
    a, _ = ap.parse_known_args()
    return a.build, a.out
BUILD, OUT = _resolve_paths()
os.makedirs(OUT, exist_ok=True)


def esc(t):
    return ihtml.escape(t, quote=False)


# ---------------------------------------------------------------- LaTeX 兼容修正
def fix_tex(tex):
    """把 KaTeX 不支持 / 表现不佳的写法替换为等价且兼容的写法。

    已知需要处理：
      * \\tag{1.1}      —— auto-render 不加载 tag 扩展；改写为末尾的 (1.1)
      * \\xrightarrow{d} —— 保留（KaTeX 0.16 支持）
      * \\notag / \\nonumber —— 直接删除
      * \\limits / \\nolimits —— 删除（KaTeX 默认行为已足够）
    """
    # 1) \tag{...}  →  \qquad\qquad(\text{...})
    tex = re.sub(r'\\tag\s*\{([^{}]*)\}', r'\\qquad\\qquad(\\text{\1})', tex)
    # 2) \notag / \nonumber
    tex = re.sub(r'\\(?:notag|nonumber)\b', '', tex)
    # 3) \limits / \nolimits
    tex = re.sub(r'\\(?:limits|nolimits)\b', '', tex)
    # 4) \substack{a\\b} → \begin{subarray}{c}...（KaTeX 支持 subarray）
    #    此处保持原样，KaTeX 0.16 已支持 \substack
    return tex


# ---------------------------------------------------------------- inline
INLINE_RE = re.compile(r'(\*\*.+?\*\*|`[^`]+`|\$[^$\n]+\$)')


def inline(text):
    """把行内的 **粗体** / `代码` / $公式$ 转成 HTML（公式交给 KaTeX）"""
    out, pos = [], 0
    for m in INLINE_RE.finditer(text):
        if m.start() > pos:
            out.append(esc(text[pos:m.start()]))
        tok = m.group(0)
        if tok.startswith('**'):
            out.append('<strong>' + inline(tok[2:-2]) + '</strong>')
        elif tok.startswith('`'):
            out.append('<code>' + esc(tok[1:-1]) + '</code>')
        else:                                    # $ ... $
            body = fix_tex(tok[1:-1])
            # 保护可能被 HTML 转义的字符
            out.append(r'<span class="math-inline">\(%s\)</span>' % body)
        pos = m.end()
    if pos < len(text):
        out.append(esc(text[pos:]))
    return ''.join(out)


# ---------------------------------------------------------------- blocks
def convert(md_text, title, accent):
    lines = md_text.split('\n')
    body, toc = [], []
    i, n, hid = 0, len(lines), 0
    while i < n:
        raw = lines[i]
        s = raw.strip()

        # ---- 空行
        if not s:
            i += 1
            continue

        # ---- 表格
        if s.startswith('|') and i + 1 < n and re.match(r'^\s*\|[\s:\-|]+\|\s*$', lines[i + 1]):
            rows = [s]
            j = i + 2
            while j < n and lines[j].strip().startswith('|'):
                rows.append(lines[j].strip()); j += 1
            cells = [[c.strip() for c in r.strip('|').split('|')] for r in rows]
            b = ['<table><thead><tr>']
            for c in cells[0]:
                b.append('<th>' + inline(c) + '</th>')
            b.append('</tr></thead><tbody>')
            for r in cells[1:]:
                b.append('<tr>' + ''.join('<td>' + inline(c) + '</td>' for c in r) + '</tr>')
            b.append('</tbody></table>')
            body.append('\n'.join(b)); i = j
            continue

        # ---- 代码块
        if s.startswith('```'):
            j = i + 1; buf = []
            while j < n and not lines[j].strip().startswith('```'):
                buf.append(esc(lines[j])); j += 1
            body.append('<pre class="code">' + '\n'.join(buf) + '</pre>')
            i = j + 1
            continue

        # ---- 独立公式 $$...$$
        if s.startswith('$$'):
            buf = s
            if s.count('$$') < 2:
                j = i + 1
                while j < n and '$$' not in lines[j]:
                    buf += '\n' + lines[j]; j += 1
                if j < n:
                    buf += '\n' + lines[j]; j += 1
                i = j
            else:
                i += 1
            inner = buf.replace('$$', '').strip()
            inner = fix_tex(inner)
            body.append(r'<div class="math-block">\[%s\]</div>' % inner)
            continue

        # ---- 标题
        m = re.match(r'^(#{1,4})\s+(.*)$', s)
        if m:
            lvl = len(m.group(1)); txt = m.group(2).strip()
            hid += 1; anchor = 'h%d' % hid
            toc.append((lvl, re.sub(r'[*`]', '', txt), anchor))
            body.append('<h%d id="%s">%s</h%d>' % (lvl, anchor, inline(txt), lvl))
            i += 1
            continue

        # ---- 水平线
        if s in ('---', '***', '___'):
            body.append('<hr>')
            i += 1
            continue

        # ---- 引用 / 提示框
        if s.startswith('>'):
            buf = []
            while i < n and lines[i].strip().startswith('>'):
                buf.append(lines[i].strip().lstrip('> ').strip())
                i += 1
            text_join = '\n'.join(buf)
            first = buf[0] if buf else ''
            if any(k in text_join[:60] for k in ('⚠', '💡', '📌', '🔢', '必背', '易错', '直觉', '关键')):
                cls = 'tip'
                if '⚠' in first or '易错' in first:
                    cls = 'tip warn'
                elif '📌' in first or '必背' in first:
                    cls = 'tip note'
                body.append('<div class="%s">%s</div>' % (cls, '<br>'.join(inline(x) for x in buf)))
            else:
                body.append('<blockquote>%s</blockquote>' % '<br>'.join(inline(x) for x in buf))
            continue

        # ---- 列表
        m = re.match(r'^(\s*)([-*+]|\d+[.)])\s+(.*)$', raw)
        if m:
            items, j = [], i
            while j < n:
                mm = re.match(r'^(\s*)([-*+]|\d+[.)])\s+(.*)$', lines[j])
                if not mm:
                    break
                items.append((len(mm.group(1)), mm.group(2) not in ('-', '*', '+'), mm.group(3)))
                j += 1
            out = []
            stack = []          # (tag, indent)
            for indent, ordered, txt in items:
                tag = 'ol' if ordered else 'ul'
                if not stack or indent > stack[-1][1]:
                    out.append('<%s>' % tag); stack.append((tag, indent))
                elif indent < stack[-1][1]:
                    while stack and indent < stack[-1][1]:
                        out.append('</%s>' % stack.pop()[0])
                    if not stack or indent > stack[-1][1]:
                        out.append('<%s>' % tag); stack.append((tag, indent))
                out.append('<li>%s</li>' % inline(txt))
            while stack:
                out.append('</%s>' % stack.pop()[0])
            body.append('\n'.join(out)); i = j
            continue

        # ---- 普通段落
        body.append('<p>%s</p>' % inline(s))
        i += 1

    # ---------- 目录
    toc_html = []
    for lvl, txt, anc in toc:
        toc_html.append('<a class="lv%d" href="#%s">%s</a>' % (min(lvl, 4), anc, esc(txt)))

    page = TEMPLATE.format(
        title=esc(title),
        accent=accent,
        toc='\n'.join(toc_html),
        body='\n'.join(body),
    )
    return page


TEMPLATE = r'''<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{title}</title>

<!-- KaTeX -->
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/katex@0.16.11/dist/katex.min.css">
<script defer src="https://cdn.jsdelivr.net/npm/katex@0.16.11/dist/katex.min.js"></script>
<script defer src="https://cdn.jsdelivr.net/npm/katex@0.16.11/dist/contrib/auto-render.min.js"
  onload="renderMathInElement(document.body,{{
    delimiters:[
      {{left:'\\\\[',right:'\\\\]',display:true}},
      {{left:'\\\\(',right:'\\\\)',display:false}}
    ],
    throwOnError:false,
    strict:false,
    trust:true,
    macros: {{
      '\\RR':'\\mathbb{{R}}',
      '\\E':'\\mathbb{{E}}',
      '\\Var':'\\mathrm{{Var}}',
      '\\Cov':'\\mathrm{{Cov}}',
      '\\plim':'\\operatorname{{plim}}',
      '\\se':'\\operatorname{{se}}',
      '\\argmax':'\\operatorname*{{arg\\,max}}',
      '\\argmin':'\\operatorname*{{arg\\,min}}'
    }}
  }});"></script>

<style>
:root {{
  --accent:{accent};
  --ink:#1b1f24;
  --sub:#5a6572;
  --line:#e2e6ee;
  --bg:#f4f6fa;
  --card:#ffffff;
  --block-bg:#fbfcfe;
}}
* {{ box-sizing:border-box; }}
html,body {{ margin:0; padding:0; }}
body {{
  font-family:"Source Han Serif SC","Noto Serif SC","Songti SC",serif;
  background:var(--bg); color:var(--ink);
  font-size:16px; line-height:1.95;
  -webkit-font-smoothing:antialiased;
}}
#wrap {{ display:flex; min-height:100vh; }}

/* ---------- 侧栏目录 ---------- */
aside {{
  width:300px; min-width:300px; height:100vh; position:sticky; top:0;
  background:var(--card); border-right:1px solid var(--line);
  overflow-y:auto; padding:20px 12px 60px;
}}
aside h3 {{
  margin:4px 8px 14px; font-size:13px; letter-spacing:2px;
  color:var(--sub); font-weight:600;
}}
aside a {{
  display:block; padding:5px 9px; margin:1px 0; border-radius:6px;
  color:#39424e; text-decoration:none; font-size:13px; line-height:1.5;
  border-left:3px solid transparent;
}}
aside a:hover {{ background:#eef2f9; color:var(--accent); }}
aside a.lv1 {{ font-weight:700; color:var(--accent); margin-top:12px; border-left-color:var(--accent); }}
aside a.lv2 {{ padding-left:16px; }}
aside a.lv3 {{ padding-left:28px; color:#5f6a78; font-size:12.5px; }}
aside a.lv4 {{ padding-left:40px; color:#7c8899; font-size:12px; }}
aside::-webkit-scrollbar {{ width:8px; }}
aside::-webkit-scrollbar-thumb {{ background:#d3d9e3; border-radius:4px; }}

/* ---------- 正文 ---------- */
main {{
  flex:1; min-width:0; max-width:940px; margin:0 auto;
  padding:40px 56px 120px; background:var(--card);
  box-shadow:0 0 24px rgba(20,40,80,.05);
}}
h1 {{
  font-size:30px; color:var(--accent); margin:0 0 6px;
  border-bottom:3px solid var(--accent); padding-bottom:12px; line-height:1.4;
}}
h2 {{
  font-size:22px; color:var(--accent); margin:44px 0 14px;
  border-left:6px solid var(--accent); padding-left:12px; line-height:1.5;
}}
h3 {{ font-size:18px; margin:30px 0 10px; color:#243247; }}
h4 {{ font-size:16px; margin:22px 0 8px; color:#3b4657; }}
p  {{ margin:12px 0; text-align:justify; }}
strong {{ color:#0f1a2b; }}

/* ---------- 公式 ---------- */
.math-block {{
  margin:20px 0; padding:14px 10px;
  background:var(--block-bg); border:1px solid var(--line);
  border-left:4px solid var(--accent);
  border-radius:4px; overflow-x:auto;
  text-align:center;
}}
.math-block .katex {{ font-size:1.32em; }}      /* 块级公式大号 */
.math-inline .katex {{ font-size:1.05em; }}     /* 行内公式微放大 */
.katex {{ color:#101828; }}
/* 公式内的 \text{{中文}} 回退到中文字体，避免方块 */
.katex .mord.text, .katex .mord.textord, .katex .textrm {{
  font-family:"Source Han Serif SC","Noto Serif SC","Songti SC","SimSun",serif;
}}
.katex .mathnormal, .katex .mord, .katex .mbin, .katex .mrel {{
  font-feature-settings:"liga" 0;
}}

/* ---------- 表格 ---------- */
table {{
  border-collapse:collapse; width:100%; margin:18px 0;
  font-size:14px; background:#fff;
  box-shadow:0 1px 3px rgba(20,40,80,.06);
}}
th {{
  background:var(--accent); color:#fff; font-weight:600;
  padding:9px 12px; text-align:left; border:1px solid var(--accent);
}}
td {{ padding:8px 12px; border:1px solid #dde2ec; vertical-align:top; line-height:1.7; }}
tbody tr:nth-child(even) {{ background:#f5f7fb; }}

/* ---------- 代码 ---------- */
.code, pre {{
  background:#f6f8fb; border:1px solid var(--line); border-radius:6px;
  padding:12px 14px; overflow-x:auto;
  font-family:Consolas,"Courier New",monospace; font-size:13px; line-height:1.65;
  color:#2b3442;
}}
code {{
  background:#eef1f6; padding:1px 6px; border-radius:4px;
  font-family:Consolas,"Courier New",monospace; font-size:13.5px; color:#a3304a;
}}

/* ---------- 引用 / 提示 ---------- */
blockquote {{
  margin:16px 0; padding:10px 18px;
  background:#fdf8ee; border-left:4px solid #d9a441;
  color:#6b5218; border-radius:0 6px 6px 0;
}}
.tip {{
  margin:16px 0; padding:12px 18px; border-radius:6px;
  background:#eaf2fd; border-left:5px solid #3a6fd8; color:#1d3a6b;
}}
.tip.warn {{ background:#fdeeee; border-left-color:#c0392b; color:#8a2620; }}
.tip.note {{ background:#f3effc; border-left-color:#6a4bc4; color:#422a80; }}

hr {{ border:none; border-top:1px dashed #cbd3e0; margin:34px 0; }}

/* ---------- 打印 ---------- */
@media print {{
  body {{ background:#fff; font-size:11.5pt; }}
  aside {{ display:none; }}
  main {{ box-shadow:none; max-width:none; padding:0 6px; }}
  .math-block {{ break-inside:avoid; }}
  h1,h2,h3 {{ break-after:avoid; }}
  table {{ break-inside:avoid; }}
}}
@media (max-width:1000px) {{
  aside {{ display:none; }}
  main {{ padding:24px 20px 80px; box-shadow:none; }}
}}
</style>
</head>
<body>
<div id="wrap">
  <aside><h3>目 录</h3>{toc}</aside>
  <main>{body}</main>
</div>
</body>
</html>
'''


# ---------------------------------------------------------------- run
JOBS = [
    ('micro_all.md', '高级微观经济学_博士资格考试复习全书.html', '高级微观经济学 · 博士资格考试复习全书', '#1f3864'),
    ('macro.md',     '高级宏观经济学_博士资格考试复习全书.html', '高级宏观经济学 · 博士资格考试复习全书', '#7a2e2e'),
    ('econ.md',      '高级计量经济学_博士资格考试复习全书.html', '高级计量经济学 · 博士资格考试复习全书', '#1d5c4d'),
]

if __name__ == '__main__':
    # 单文件模式：python make_katex_html.py input.md output.html [标题] [主题色]
    #   不带位置参数时，跑下方 JOBS 示例（把文件名换成你自己的 .md）
    pos = [x for x in sys.argv[1:] if not x.startswith('--')]
    if len(pos) >= 2:
        src, out = pos[0], pos[1]
        title = pos[2] if len(pos) > 2 else os.path.splitext(os.path.basename(src))[0]
        accent = pos[3] if len(pos) > 3 else '#1f3864'
        md = open(src, encoding='utf-8').read()
        page = convert(md, title, accent)
        open(out, 'w', encoding='utf-8').write(page)
        print('saved:', out, '(%.1f KB)' % (len(page) / 1024))
    else:
        for src, out, title, accent in JOBS:
            p = os.path.join(BUILD, src)
            if not os.path.exists(p):
                print('skip (missing):', src); continue
            md = open(p, encoding='utf-8').read()
            page = convert(md, title, accent)
            dst = os.path.join(OUT, out)
            open(dst, 'w', encoding='utf-8').write(page)
            print('saved:', dst, '(%.1f KB)' % (len(page) / 1024))
    print('ALL DONE')

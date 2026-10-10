---
name: 公式排版助手
description: 把含数学公式的 Markdown 一键变成精美 Word / PDF / HTML。解决 Word 里公式显示乱码、HTML 里 KaTeX 加载失败、PDF 公式需静态矢量等高频痛点。触发词：公式排版、md 转 Word、md 转 PDF、KaTeX、LaTeX 公式、复习资料、讲义、教材、报告、OMML、Word 公式乱码、公式不显示。
agent_created: true
---

# 公式排版助手

## 这个技能解决什么问题

用 `python-docx` 生成含数学公式的 Word 文档时，最常见的失败是**公式以纯文本或乱码出现**。
本技能用 `latex2mathml` → `mathml2omml` 把 LaTeX 转成 **Word 原生公式对象**（`<m:oMath>`），
公式在 Word 里可正常显示、可编辑。

同时处理五类**已知的库缺陷**（这些缺陷不会抛异常或只抛难懂的错误，极易漏掉）。

## 使用方式

```bash
# 依赖
pip install python-docx lxml latex2mathml mathml2omml

# 1) Markdown -> Word（公式转原生 OMML 对象，可编辑）
python scripts/make_docx.py your.md your.docx          # 单文件
python scripts/make_docx.py --build /src --out /out     # 批量（跑脚本内置示例 JOBS）

# 2) 转换前体检：公式 XML 合法性 + 语义泄漏（裸命令残片）
python scripts/probe.py your.md

# 3) 转换后体检：OMML 数量 / 残留源码 / 标题结构
python scripts/verify.py

# 4) Markdown -> KaTeX HTML（网页/打印友好，块级公式大号居中）
python scripts/make_katex_html.py your.md your.html

# 5) HTML -> PDF（headless Chromium 渲染 KaTeX 后打印）
pip install playwright && python -m playwright install chromium
python scripts/html_to_pdf.py --out /out
```

路径解析优先级（所有脚本统一）：
**命令行 `--build`/`--out`  >  环境变量 `BUILD_DIR`/`OUT_DIR`  >  当前工作目录**。
不带参数时默认在当前目录读写，clone 后开箱即用。
```

Bash 环境若 `python` 不在 PATH，先 `export PATH="/usr/bin:/bin:$PATH"`。

## 扩展 Markdown 书写规范

| 语法 | 效果 |
|---|---|
| `# 章` | 一级标题（**前面自动分页**） |
| `## 节` / `### 小节` | 二/三级标题（带下划线装饰） |
| `$$...$$` | 独立公式（居中，多行用 `\\` 换行） |
| `$...$` | 行内公式 |
| `\begin{aligned}...\end{aligned}` | 多行公式（自动转 `array`） |
| `\begin{pmatrix}...\end{pmatrix}` | 矩阵（**可以直接带下标**，见修复 2） |
| `\|...\|` | 表格（隔行浅色底纹） |
| `> ` | 引用块（左侧竖线） |
| `- ` / `1. ` | 无序/有序列表 |
| ` ```...``` ` | 代码块 |
| `::: box 标题` … `:::` | 蓝色提示框 |
| `::: example 标题` … `:::` | 黄色例题框 |
| `::: answer 标题` … `:::` | 绿色解答框 |
| `::: warn 标题` … `:::` | 红色警示框 |
| `::: note 标题` … `:::` | 紫色注释框 |

**⚠ 最重要的一条**：行内公式可以放在粗体里（`**技巧 1　$x_t\approx\bar x$ 的定义**`），
渲染器会正确处理——但前提是使用了本技能修正版 `add_inline`（见修复 5）。

## 五类已知缺陷与修复（都已内置于 make_docx.py）

### 修复 1　`aligned` 环境生成畸形 XML
`mathml2omml` 处理 `aligned` 会生成不合法 XML（标签不匹配）。
**做法**：`preprocess` 里把 `aligned`/`split`/`gather` 统一改写为 `array`。

### 修复 2　矩阵不能挂上下标 ⭐ 最容易踩
`\begin{pmatrix}...\end{pmatrix}_{n}` 会报
`RuntimeError: Too large number of children for msub`——因为下标要挂到多行 matrix 节点上。
**做法**：只对"后面紧跟 `_` 或 `^`"的 `pmatrix`/`bmatrix`，
降级为 `matrix` 并用 `\left(...\right)` / `\left[...\right]` 包裹，
使下标挂到单一定界符组上。**不要全局替换**（会造成双层括号）。

### 修复 3　重音命令的反向引用编号 ⭐ 静默语义损坏
剥离 `\bar`/`\hat`/`\overline` 等重音命令时，`ACC` 本身就是**第 1 个捕获组**，
因此内容组从 `\2` 开始。若误写成 `\1`，`\bar r`（重音+空格+单字母）会被替换成裸文本 **`bar`**——
不报任何错误，但 Word 里就显示成 "bar"，公式语义全错。
**做法**：替换串用 `\2`/`\2\3`，并在最后加两条兜底规则清除任何残留重音命令。

### 修复 4　行分隔符 `\\` 被细空格替换吃掉
`for a in (r'\quad', r'\ ', ...): t = t.replace(a, ' ')` 这条规则会把
`\\ ` （行分隔符+空格）中的第二个反斜杠+空格当作细空格 `\ ` 替换掉，
导致矩阵/多行公式的换行丢失、内容被静默拼接（数学上错误！）。
**做法**：替换前先把 `\\` 用哨兵 `\x00ROWBREAK\x00` 占位，全部替换完成后再还原。

### 修复 5　粗体片段内部的行内公式不渲染 ⭐ 最高频
`INLINE_RE` 的交替顺序把 `\*\*.+?\*\*` 放在最前，
所以 `**技巧 1　$x_t\approx\bar x$ 的定义**` 会整体匹配为粗体 token，
粗体分支直接把内部文本当成一个 run 输出 → **`$...$` 原样显示为 LaTeX 源码**。
**做法**：`add_inline` 里对粗体片段**递归调用自身**并透传 `bold=True`；
同时在函数入口（非粗体分支）调用 `_sanitize_outside_math()`。

## 两道质量防线（务必都跑）

1. **`probe.py`**：转换**前**扫全部公式，报告
   - `xml-bad`：XML 不合法的公式；
   - `leak-bad`：**语义泄漏**——`preprocess` 后出现裸命令名（`bar`、`overline`、`frac`…）。
     这类错误不抛异常但会显示成乱码，**必须专门拦截**。
2. **`verify.py`**：转换**后**开 docx，统计公式对象数、残留反斜杠、
   伪分式、标题层级，并列出疑似泄漏上下文。
   **注意**：MATLAB 代码里的 `\`（左除）、ASCII 图里的 `/ \` 是合法的，不必处理。

## 其他实用细节

- `_sanitize_outside_math()`：清除 `$...$` **之外**的裸 LaTeX 命令
  （如作者笔误 `\quad⇒\quad`）与 Markdown 转义残留（`Q\*`、`\$`、`\&`）。
  这是必要的，因为这类残片会以源码形式出现在正文里。
- 中文粗体无效果？Word 里中文加粗需设 `w:bCs`，并给东亚字体设 `qn('w:eastAsia')`。
- 目录：插入 `TOC \o "1-2" \h \z \u` 域；页码：`PAGE` 域。
  **首次打开需按 Ctrl+A、F9 更新目录**，否则显示为空白域。
- 判断题/选择题的 `\checkmark`、`\blacksquare`、`\varnothing` 等符号
  `mathml2omml` 不支持，需在 `preprocess` 中直接删除。

## 备选输出：KaTeX HTML（公式最标准美观）

当用户觉得 **Word 里的公式"看着别扭"**，或明确说要 "KaTeX / LaTeX 格式" 时，
改用 `scripts/make_katex_html.py` 生成 HTML —— 公式由浏览器端 KaTeX 引擎渲染，
是数学排版质量最高、且可一键打印成 PDF 的方案。

```bash
# 生成三份 KaTeX HTML（CDN 加载 katex@0.16.11，无需本地包）
python scripts/make_katex_html.py

# 兼容性扫描（找出 KaTeX 不支持 / 高风险的命令）
python scripts/scan_katex.py

# 真实引擎验证（推荐，最可靠）
cd scripts && npm install katex --no-save && node test_katex.js
```

### KaTeX 版的要点

- 定界符：块级 `\[...\]`（**1.32em 居中**）、行内 `\(...\)`（1.05em）。
- **必须做兼容层 `fix_tex()`**，否则部分公式会红字报错：

| 原写法 | KaTeX 问题 | 处理 |
|---|---|---|
| `\tag{1.1}` | auto-render 默认不载 tag 扩展 | → `\qquad\qquad(\text{1.1})` |
| `\notag` / `\nonumber` | 不支持 | 直接删除 |
| `\limits` / `\nolimits` | 不支持 | 直接删除 |
| `\xrightarrow{d}` | ✅ 0.16 起支持 | 保留 |
| `\underbrace` `\begin{cases}` `\begin{pmatrix}` `\displaystyle` `\qquad` | ✅ 支持 | 保留 |

- `\text{中文}` 需在 CSS 里给 `.katex .mord.text` 加 **CJK 字体回退**
  （Source Han Serif / Noto Serif SC / Songti SC / SimSun），否则显示方块。
- 用 `macros` 配置注册常用算子：`\RR` `\E` `\Var` `\Cov` `\plim` `\se` `\argmax` `\argmin`。
- **验证脚本的坑**：统计定界符配对时，LaTeX 的 `\\[4pt]`（换行带间距）里含 `\[`，
  会被误判为块级定界符——用 `re.sub(r'\\\\\[[^\]]*\]', '', h)` 先剔除。
  用真实 katex 引擎测试时，也必须先 `h.replace(/<script[\s\S]*?<\/script>/g, '')`
  剔除 `renderMathInElement` 的配置段（里面的 `delimiters:[{left:'\\['...}]` 会被当成公式）。

## 第三种输出：HTML → PDF（公式最可靠）

当用户反馈"HTML 在浏览器里公式还是不对"或直接要求 PDF 时，
用 `scripts/html_to_pdf.py`（Playwright + headless Chromium）打印 PDF：
公式在 PDF 里是静态 SVG 矢量图，浏览器渲染失败也不会影响。

```bash
# 一次性安装（约 130MB）
python -m pip install playwright pypdf
python -m playwright install chromium

# 生成 PDF
python scripts/html_to_pdf.py
```

### 打印 PDF 的关键坑

1. **必须显式调 `renderMathInElement`**：HTML 里 `<script defer>` 的 auto-render
   在 headless + `wait_until='networkidle'` 下 `onload` 不可靠，KaTeX 节点数会为 0。
   在 `page.evaluate()` 里手动调一次（macros 与 HTML 一致），等 `.katex` 节点 > 50
   再 `page.pdf()`。本项目实测：宏 3214 / 微观 736 / 计量 1344 节点全部成功渲染。

2. **中文字体回退**：Chromium headless 默认无 CJK 字体，PDF 正文回退到日韩繁体
   （如 `卷⽔`、`复习精⼒`）。在 CSS 字体栈首位放 `SimSun`（Windows 自带 `simsun.ttc`）
   或 `Microsoft YaHei`；公式内的 `\text{中文}` 同理。

3. **页脚模板**：`display_header_footer=True` + `footer_template` 放页码，
   `header_template` 留空（避免空白头）。

4. **体积**：KaTeX SVG 精度高，PDF 体积约 3-5 MB / 科（5350 公式）。
   若用户嫌大可改 `katex.options.output = 'html'`（用 HTML+CSS 而非 SVG），
   体积减半但复杂公式排版略差。

## 排查顺序（公式出问题时按此顺序）

1. 跑 `probe.py` 看是 `xml-bad` 还是 `leak-bad`；
2. `leak-bad` → 检查 `preprocess` 的重音剥离与反向引用编号（修复 3）；
3. `xml-bad` 且含 `matrix` → 检查矩阵下标改写（修复 2）；
4. `xml-bad` 且含多行 → 检查 `aligned` 改写与 `\\` 占位（修复 1、4）；
5. **正文里出现 `$...$` 原文** → 检查粗体递归（修复 5）；
6. 都对但仍失败 → 该公式确实不被支持，走 `clean_math()` 的 Unicode 降级路径
   （会显示为斜体纯文本但至少不是乱码）。

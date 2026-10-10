# 公式排版助手

把含数学公式的 Markdown 一键变成精美 **Word / PDF / HTML**。

专治三大痛点：

- **Word 里公式乱码 / 纯文本** → 转成 Word 原生 OMML 公式对象（`<m:oMath>`，可编辑、可复制）
- **网页里 KaTeX 加载失败 / 显示成 `\frac{...}` 源码** → 显式触发渲染，零泄漏
- **PDF 公式需静态矢量** → headless Chromium 渲染 KaTeX 后打印，不会再出现 LaTeX 源码

---

## 目录结构

```
公式排版助手/
├── SKILL.md              # 技能描述（给 AI 看的触发词 + 工作流）
├── README.md             # 本文档
└── scripts/
    ├── make_docx.py        # Markdown -> Word（原生 OMML 公式）
    ├── make_katex_html.py  # Markdown -> KaTeX HTML
    ├── html_to_pdf.py      # HTML -> PDF（Playwright headless Chromium）
    ├── probe.py            # 转换前体检：XML 合法性 + 语义泄漏
    ├── verify.py           # 转换后体检：OMML 数 / 残留源码 / 标题结构
    ├── scan_katex.py       # 扫描 KaTeX 不支持的命令
    └── test_katex.js       # 用真实 KaTeX 引擎逐条验证渲染是否成功
```

---

## 环境依赖

```bash
pip install python-docx lxml latex2mathml mathml2omml

# 生成 HTML / PDF 还需：
npm install katex
pip install playwright && python -m playwright install chromium
```

---

## 快速开始

### 1. 生成 Word（公式可编辑）

```bash
python scripts/make_docx.py your.md your.docx
```

### 2. 生成 KaTeX HTML（网页阅读 / 打印）

```bash
python scripts/make_katex_html.py your.md your.html
```

### 3. 生成 PDF

```bash
python scripts/html_to_pdf.py --out /path/to/html_dir
```

---

## 路径参数

所有脚本统一支持以下配置（不再写死任何本地路径）：

| 参数 | 环境变量 | 默认 |
|---|---|---|
| `--build` | `BUILD_DIR` | 当前工作目录 |
| `--out`   | `OUT_DIR`   | 当前工作目录 |

解析优先级：**命令行参数 > 环境变量 > 当前工作目录**。
不带任何参数时，默认在当前目录读写，`git clone` 后开箱即用。

此外，`probe.py` / `scan_katex.py` 支持把待处理的 `.md` 文件作为**位置参数**传入：

```bash
python scripts/probe.py a.md b.md      # 只检查指定文件
python scripts/scan_katex.py a.md      # 只扫描指定文件
```

---

## 扩展 Markdown 语法

| 语法 | 效果 |
|---|---|
| `# 章` | 一级标题（前面自动分页） |
| `## 节` / `### 小节` | 二 / 三级标题（带下划线装饰） |
| `$$...$$` | 独立公式（居中，多行用 `\\` 换行） |
| `$...$` | 行内公式 |
| `::: box 标题` ... `:::` | 彩色批注框（直觉 / 说明） |
| `::: example 标题` ... `:::` | 例题框（浅黄底 + 左侧粗边） |
| `::: answer 标题` ... `:::` | 解答框（浅绿底） |
| `::: warn 标题` ... `:::` | 警示框（浅红底） |
| `> 行首引用` | 楷体批注 |

---

## 兼容性要点（KaTeX）

KaTeX 与 Word 的 OMML 对 LaTeX 子集支持不同，已做针对性处理：

| 原写法 | 处理 |
|---|---|
| `\tag{1.1}` | 改写为 `\qquad\qquad(\text{1.1})`（auto-render 默认不载 tag 扩展） |
| `\notag` / `\nonumber` | 直接删除 |
| `\limits` / `\nolimits` | 直接删除 |
| `\xrightarrow{d}` | 保留（KaTeX 0.16 起支持） |
| `\underbrace` / `\begin{cases}` / `\begin{pmatrix}` / `\displaystyle` | 保留（支持） |

公式内的 `\text{中文}` 已配置 CJK 字体回退（Source Han Serif / Noto Serif SC / SimSun），不会显示方块。

---

## 许可

MIT License

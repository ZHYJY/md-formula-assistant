# Formula Typeset Assistant

One-click conversion of Markdown with math formulas into polished **Word / PDF / HTML**.

Three pain points it solves:

- **Garbled / plain-text formulas in Word** → converted to native Word OMML formula objects (`<m:oMath>`, editable and copyable)
- **KaTeX load failures / `\frac{...}` source showing in web pages** → explicitly trigger rendering, zero leakage
- **Static vector formulas needed in PDF** → headless Chromium renders KaTeX then prints, no more LaTeX source

> Chinese documentation: see [README.md](./README.md). The Chinese skill name is **公式排版助手**.

---

## Directory Structure

```
Formula Typeset Assistant/
├── SKILL.md              # Skill description (Chinese, name: 公式排版助手) — triggers + workflow for the AI
├── SKILL.en.md           # English version of this skill (name: Formula Typeset Assistant)
├── README.md             # Chinese documentation
├── README.en.md          # This document (English)
└── scripts/
    ├── make_docx.py        # Markdown -> Word (native OMML formulas)
    ├── make_katex_html.py  # Markdown -> KaTeX HTML
    ├── html_to_pdf.py      # HTML -> PDF (Playwright headless Chromium)
    ├── probe.py            # Pre-conversion check: XML validity + semantic leakage
    ├── verify.py           # Post-conversion check: OMML count / residual source / heading structure
    ├── scan_katex.py       # Scan KaTeX-unsupported commands
    └── test_katex.js       # Verify rendering one-by-one with the real KaTeX engine
```

---

## Dependencies

```bash
pip install python-docx lxml latex2mathml mathml2omml

# To generate HTML / PDF additionally:
npm install katex
pip install playwright && python -m playwright install chromium
```

---

## Quick Start

### 1. Generate Word (editable formulas)

```bash
python scripts/make_docx.py your.md your.docx
```

### 2. Generate KaTeX HTML (web reading / printing)

```bash
python scripts/make_katex_html.py your.md your.html
```

### 3. Generate PDF

```bash
python scripts/html_to_pdf.py --out /path/to/html_dir
```

---

## Path Parameters

All scripts uniformly support the following configuration (no hardcoded local paths):

| Parameter | Env Variable | Default |
|---|---|---|
| `--build` | `BUILD_DIR` | current working directory |
| `--out`   | `OUT_DIR`   | current working directory |

Resolution priority: **command-line argument > environment variable > current working directory**.
With no arguments, reads/writes in the current directory by default — works out of the box after `git clone`.

In addition, `probe.py` / `scan_katex.py` accept the `.md` files to process as **positional arguments**:

```bash
python scripts/probe.py a.md b.md      # check only specified files
python scripts/scan_katex.py a.md      # scan only specified files
```

---

## Extended Markdown Syntax

| Syntax | Effect |
|---|---|
| `# Chapter` | Level-1 heading (auto page break before) |
| `## Section` / `### Subsection` | Level-2 / 3 headings (with underline decoration) |
| `$$...$$` | Standalone formula (centered, use `\\` for line breaks) |
| `$...$` | Inline formula |
| `::: box Title` ... `:::` | Colored annotation box (intuition / explanation) |
| `::: example Title` ... `:::` | Example box (light yellow background + left bold border) |
| `::: answer Title` ... `:::` | Answer box (light green background) |
| `::: warn Title` ... `:::` | Warning box (light red background) |
| `> Quote at line start` | Kaiti (楷体) annotation |

---

## Compatibility Notes (KaTeX)

KaTeX and Word's OMML support different LaTeX subsets; targeted handling is in place:

| Original | Handling |
|---|---|
| `\tag{1.1}` | rewritten to `\qquad\qquad(\text{1.1})` (auto-render doesn't load tag extension by default) |
| `\notag` / `\nonumber` | deleted directly |
| `\limits` / `\nolimits` | deleted directly |
| `\xrightarrow{d}` | kept (supported since KaTeX 0.16) |
| `\underbrace` / `\begin{cases}` / `\begin{pmatrix}` / `\displaystyle` | kept (supported) |

`\text{Chinese}` inside formulas is configured with a CJK font fallback (Source Han Serif / Noto Serif SC / SimSun), no boxes.

---

## License

MIT License

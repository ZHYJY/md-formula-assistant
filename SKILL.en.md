---
name: Formula Typeset Assistant
description: One-click conversion of Markdown with math formulas into polished Word / PDF / HTML. Solves common pain points — garbled formulas in Word, KaTeX load failures in HTML, and the need for static vector formulas in PDF. Triggers: formula typesetting, md to Word, md to PDF, KaTeX, LaTeX formulas, study notes, lecture notes, textbooks, reports, OMML, Word formula garbling, formulas not displaying.
agent_created: true
---

# Formula Typeset Assistant

## What problem does this skill solve

When using `python-docx` to generate Word documents containing math formulas, the most common failure is **formulas appearing as plain text or garbled characters**.
This skill uses `latex2mathml` → `mathml2omml` to convert LaTeX into **native Word formula objects** (`<m:oMath>`),
so formulas display correctly and stay editable inside Word.

It also handles five categories of **known library defects** (these defects raise no exception or only cryptic errors, making them extremely easy to miss).

## Usage

```bash
# Dependencies
pip install python-docx lxml latex2mathml mathml2omml

# 1) Markdown -> Word (formulas become native OMML objects, editable)
python scripts/make_docx.py your.md your.docx          # single file
python scripts/make_docx.py --build /src --out /out     # batch (runs the built-in example JOBS)

# 2) Pre-conversion check: formula XML validity + semantic leakage (naked command fragments)
python scripts/probe.py your.md

# 3) Post-conversion check: OMML count / residual source / heading structure
python scripts/verify.py

# 4) Markdown -> KaTeX HTML (web/print friendly, block formulas large and centered)
python scripts/make_katex_html.py your.md your.html

# 5) HTML -> PDF (headless Chromium renders KaTeX then prints)
pip install playwright && python -m playwright install chromium
python scripts/html_to_pdf.py --out /out
```

Path resolution priority (uniform across all scripts):
**command-line `--build`/`--out`  >  environment variable `BUILD_DIR`/`OUT_DIR`  >  current working directory**.
With no arguments it reads/writes in the current directory by default — works out of the box after clone.

If `python` is not on PATH in your shell, first run `export PATH="/usr/bin:/bin:$PATH"`.

## Extended Markdown Syntax

| Syntax | Effect |
|---|---|
| `# Chapter` | Level-1 heading (**auto page break before**) |
| `## Section` / `### Subsection` | Level-2/3 headings (with underline decoration) |
| `$$...$$` | Standalone formula (centered, use `\\` for line breaks in multi-line) |
| `$...$` | Inline formula |
| `\begin{aligned}...\end{aligned}` | Multi-line formula (auto-converted to `array`) |
| `\begin{pmatrix}...\end{pmatrix}` | Matrix (**can directly carry subscripts**, see Fix 2) |
| `\|...\|` | Table (alternating light row shading) |
| `> ` | Blockquote (left vertical bar) |
| `- ` / `1. ` | Unordered/ordered list |
| ` ```...``` ` | Code block |
| `::: box Title` … `:::` | Blue tip box |
| `::: example Title` … `:::` | Yellow example box |
| `::: answer Title` … `:::` | Green answer box |
| `::: warn Title` … `:::` | Red warning box |
| `::: note Title` … `:::` | Purple note box |

**⚠ The most important rule**: inline formulas can be placed inside bold text (`**Tip 1　$x_t\approx\bar x$ definition**`),
the renderer handles it correctly — but only if using this skill's patched `add_inline` (see Fix 5).

## Five Known Defects and Fixes (all built into make_docx.py)

### Fix 1　`aligned` environment generates malformed XML
`mathml2omml` produces invalid XML when processing `aligned` (mismatched tags).
**Approach**: in `preprocess`, rewrite `aligned`/`split`/`gather` uniformly to `array`.

### Fix 2　Matrix cannot carry sub/superscripts ⭐ easiest to trip over
`\begin{pmatrix}...\end{pmatrix}_{n}` raises
`RuntimeError: Too large number of children for msub` — because the subscript must attach to the multi-row matrix node.
**Approach**: only for `pmatrix`/`bmatrix` **immediately followed by `_` or `^`**,
downgrade to `matrix` and wrap with `\left(...\right)` / `\left[...\right]`,
so the subscript attaches to a single delimiter group. **Do not do a global replace** (would cause double brackets).

### Fix 3　Accent commands' backreference numbering ⭐ silent semantic corruption
When stripping accent commands like `\bar`/`\hat`/`\overline`, `ACC` itself is **capture group 1**,
so the content group starts from `\2`. If mistakenly written as `\1`, `\bar r` (accent+space+single letter) would be replaced with the naked text **`bar`** —
no error is raised, but Word displays "bar", completely breaking the formula semantics.
**Approach**: use `\2`/`\2\3` in the replacement string, and add two fallback rules at the end to clear any residual accent commands.

### Fix 4　Line separator `\\` eaten by thin-space replacement
The rule `for a in (r'\quad', r'\ ', ...): t = t.replace(a, ' ')` replaces
the `\ ` (second backslash + space) in `\\ ` (line separator + space) as if it were a thin space `\ `,
causing line breaks in matrices/multi-line formulas to be lost and content silently concatenated (mathematically wrong!).
**Approach**: before replacement, placeholder `\\` with sentinel `\x00ROWBREAK\x00`, restore after all replacements are done.

### Fix 5　Inline formulas inside bold spans don't render ⭐ most frequent
`INLINE_RE`'s alternation order puts `\*\*.+?\*\*` first,
so `**Tip 1　$x_t\approx\bar x$ definition**` gets matched entirely as a bold token,
and the bold branch outputs the inner text as a single run → **`$...$` displays as raw LaTeX source**.
**Approach**: in `add_inline`, recursively call itself on bold spans and pass through `bold=True`;
also call `_sanitize_outside_math()` at the function entry (non-bold branch).

## Two Quality Guardrails (run both)

1. **`probe.py`**: scans all formulas **before** conversion, reporting
   - `xml-bad`: formulas with invalid XML;
   - `leak-bad`: **semantic leakage** — bare command names (`bar`, `overline`, `frac`…) appearing after `preprocess`.
     These errors raise no exception but display as garbled text, **must be specifically intercepted**.
2. **`verify.py`**: opens the docx **after** conversion, counts formula objects, residual backslashes,
   pseudo-fractions, heading levels, and lists suspected-leak contexts.
   **Note**: `\` (left-divide) in MATLAB code and `/ \` in ASCII art are legitimate and need no handling.

## Other Practical Details

- `_sanitize_outside_math()`: removes bare LaTeX commands **outside** `$...$`
  (e.g., author typos `\quad⇒\quad`) and Markdown escape residues (`Q\*`、`\$`、`\&`).
  This is necessary because such fragments appear as source text in the body.
- Chinese bold has no visual effect? In Word, bolding Chinese requires setting `w:bCs` and assigning `qn('w:eastAsia')` to the East Asian font.
- Table of contents: insert `TOC \o "1-2" \h \z \u` field; page numbers: `PAGE` field.
  **On first open, press Ctrl+A then F9 to update the TOC**, otherwise it shows as empty fields.
- Symbols like `\checkmark`, `\blacksquare`, `\varnothing` in true/false or multiple-choice questions
  are unsupported by `mathml2omml` and must be deleted directly in `preprocess`.

## Alternative Output: KaTeX HTML (highest-quality, most standard formulas)

When the user feels **formulas in Word "look awkward"**, or explicitly asks for "KaTeX / LaTeX format",
switch to `scripts/make_katex_html.py` to generate HTML — formulas are rendered by the browser-side KaTeX engine,
the highest-quality math typesetting solution that can also be printed to PDF with one click.

```bash
# Generate KaTeX HTML (loads katex@0.16.11 via CDN, no local package needed)
python scripts/make_katex_html.py

# Compatibility scan (finds KaTeX-unsupported / high-risk commands)
python scripts/scan_katex.py

# Real-engine verification (recommended, most reliable)
cd scripts && npm install katex --no-save && node test_katex.js
```

### KaTeX Version Key Points

- Delimiters: block `\[...\]` (**1.32em centered**), inline `\(...\)` (1.05em).
- **A compatibility layer `fix_tex()` is required**, otherwise some formulas will error in red:

| Original | KaTeX problem | Handling |
|---|---|---|
| `\tag{1.1}` | auto-render doesn't load tag extension by default | → `\qquad\qquad(\text{1.1})` |
| `\notag` / `\nonumber` | unsupported | delete directly |
| `\limits` / `\nolimits` | unsupported | delete directly |
| `\xrightarrow{d}` | ✅ supported since 0.16 | keep |
| `\underbrace` `\begin{cases}` `\begin{pmatrix}` `\displaystyle` `\qquad` | ✅ supported | keep |

- `\text{Chinese}` requires adding a **CJK font fallback** to `.katex .mord.text` in CSS
  (Source Han Serif / Noto Serif SC / Songti SC / SimSun), otherwise boxes appear.
- Register common operators via `macros` config: `\RR` `\E` `\Var` `\Cov` `\plim` `\se` `\argmax` `\argmin`.
- **Pitfall in verification scripts**: when counting delimiter pairings, LaTeX's `\\[4pt]` (line break with spacing) contains `\[`,
  which is misjudged as a block delimiter — use `re.sub(r'\\\\\[[^\]]*\]', '', h)` to remove it first.
  When testing with the real katex engine, you must also `h.replace(/<script[\s\S]*?<\/script>/g, '')`
  to remove the `renderMathInElement` config block (the `delimiters:[{left:'\\['...}]` inside would be treated as a formula).

## Third Output: HTML → PDF (most reliable formulas)

When the user reports "formulas are still wrong in the browser HTML" or directly requests PDF,
use `scripts/html_to_pdf.py` (Playwright + headless Chromium) to print PDF:
formulas in the PDF are static SVG vector graphics — even if browser rendering fails, it won't affect the output.

```bash
# One-time install (~130MB)
python -m pip install playwright pypdf
python -m playwright install chromium

# Generate PDF
python scripts/html_to_pdf.py
```

### Key Pitfalls in Printing PDF

1. **Must explicitly call `renderMathInElement`**: the `<script defer>` auto-render in HTML
   has unreliable `onload` under headless + `wait_until='networkidle'`, resulting in 0 KaTeX nodes.
   Call it manually once inside `page.evaluate()` (macros consistent with HTML), wait until `.katex` nodes > 50
   then `page.pdf()`. Measured in this project: macros 3214 / micro 736 / econometrics 1344 nodes all rendered successfully.

2. **Chinese font fallback**: Chromium headless has no CJK font by default, PDF body falls back to Japanese/Korean traditional forms
   (e.g., `卷⽔`, `复习精⼒`). Put `SimSun` (Windows built-in `simsun.ttc`) or `Microsoft YaHei` first in the CSS font stack;
   `\text{Chinese}` inside formulas is the same.

3. **Footer template**: `display_header_footer=True` + `footer_template` for page numbers,
   `header_template` left empty (to avoid blank header).

4. **Size**: KaTeX SVG is high-precision, PDF size ~3-5 MB / subject (5350 formulas).
   If the user finds it too large, change `katex.options.output = 'html'` (use HTML+CSS instead of SVG),
   halving the size but slightly worse typesetting for complex formulas.

## Troubleshooting Order (when formulas fail, follow this order)

1. Run `probe.py` to see if it's `xml-bad` or `leak-bad`;
2. `leak-bad` → check accent stripping and backreference numbering in `preprocess` (Fix 3);
3. `xml-bad` and contains `matrix` → check matrix subscript rewriting (Fix 2);
4. `xml-bad` and contains multi-line → check `aligned` rewriting and `\\` placeholder (Fix 1, 4);
5. **`$...$` source appears in body** → check bold recursion (Fix 5);
6. All correct but still fails → the formula is genuinely unsupported, fall back to `clean_math()`'s Unicode degradation path
   (displays as italic plain text but at least not garbled).

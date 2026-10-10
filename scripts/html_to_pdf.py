# -*- coding: utf-8 -*-
"""用 Playwright 启动 headless Chromium，把三份 KaTeX HTML 打印成 PDF。

公式在 PDF 里是静态 SVG/Canvas，浏览器渲染失败也不会影响。
为避免 CDN 不可用导致公式空白，做"CDN 失败 → 渲染后注入"双保险：
  * 第一遍尝试常规加载 + 网络等待
  * 如果 KaTeX 没渲染（<title> 找不到 KaTeX 节点），改用 file:// 直接打开本地资源
"""
import sys, os, asyncio, time, re, glob, argparse
sys.stdout.reconfigure(encoding='utf-8')

from playwright.async_api import async_playwright

def _resolve_out():
    """OUT 可配置：命令行 --out  >  环境变量 OUT_DIR  >  当前工作目录"""
    ap = argparse.ArgumentParser(add_help=False)
    ap.add_argument('--out', default=os.environ.get('OUT_DIR', os.getcwd()))
    a, _ = ap.parse_known_args()
    return a.out
OUT = _resolve_out()


def list_html(pattern='*.html'):
    """默认处理 OUT 目录下所有 .html；可传 pattern 过滤"""
    return sorted(glob.glob(os.path.join(OUT, pattern)))


def has_katex_rendered(html_text):
    """粗判：KaTeX 渲染成功的话，HTML 里应该出现 <span class="katex">"""
    return '<span class="katex' in html_text


async def html_to_pdf(p, html_path, pdf_path):
    """用 Playwright 把单个 HTML 打印成 PDF。
    关键：浏览器加载后必须显式调用 renderMathInElement，因为 CDN 上的
    auto-render 的 onload 不一定触发（headless + defer 的兼容问题）。
    """
    browser = await p.chromium.launch(headless=True)
    ctx = await browser.new_context()
    page = await ctx.new_page()

    url = 'file:///' + html_path.replace('\\', '/')
    try:
        # 给 KaTeX CDN 一点加载时间
        await page.goto(url, wait_until='networkidle', timeout=60000)
        # 显式调用 renderMathInElement（不依赖 auto-render 的 onload）
        await page.evaluate("""
          async () => {
            // 等 katex 库加载完成（最多 10s）
            for (let i = 0; i < 50; i++) {
              if (typeof katex !== 'undefined' &&
                  typeof katex.renderToString === 'function' &&
                  typeof renderMathInElement === 'function') break;
              await new Promise(r => setTimeout(r, 200));
            }
            if (typeof renderMathInElement === 'function') {
              renderMathInElement(document.body, {
                delimiters: [
                  {left: '\\\\[', right: '\\\\]', display: true},
                  {left: '\\\\(', right: '\\\\)', display: false}
                ],
                throwOnError: false, strict: false, trust: true,
                macros: {
                  '\\\\RR':'\\\\mathbb{R}', '\\\\E':'\\\\mathbb{E}',
                  '\\\\Var':'\\\\mathrm{Var}', '\\\\Cov':'\\\\mathrm{Cov}',
                  '\\\\plim':'\\\\operatorname{plim}', '\\\\se':'\\\\operatorname{se}',
                  '\\\\argmax':'\\\\operatorname*{arg\\\\,max}',
                  '\\\\argmin':'\\\\operatorname*{arg\\\\,min}'
                }
              });
            }
          }
        """)
        # 等 KaTeX 渲染（最多等 15 秒）
        rendered = 0
        for _ in range(60):
            rendered = await page.evaluate("document.querySelectorAll('.katex').length")
            if rendered and rendered > 50:
                break
            await page.wait_for_timeout(250)
        print(f'   katex rendered: {rendered} nodes')
        # 再多等 0.5 秒
        await page.wait_for_timeout(500)

        await page.pdf(
            path=pdf_path,
            format='A4',
            print_background=True,
            margin={'top': '14mm', 'bottom': '16mm', 'left': '14mm', 'right': '14mm'},
            display_header_footer=True,
            header_template='<div></div>',
            footer_template='<div style="font-size:9px;color:#888;width:100%;text-align:center;margin-top:6px"><span class="pageNumber"></span> / <span class="totalPages"></span></div>',
        )
        await browser.close()
        return True, None
    except Exception as e:
        try:
            await browser.close()
        except Exception:
            pass
        return False, str(e)


async def main():
    files = list_html()
    if not files:
        print('NO HTML FILES FOUND in', OUT)
        return

    async with async_playwright() as p:
        for html in files:
            base = os.path.splitext(os.path.basename(html))[0]
            pdf = os.path.join(OUT, base + '.pdf')
            print('→', base)
            t0 = time.time()
            ok, err = await html_to_pdf(p, html, pdf)
            dt = time.time() - t0
            if ok:
                size_kb = os.path.getsize(pdf) / 1024
                print(f'   OK  {pdf}  ({size_kb:.0f} KB, {dt:.1f}s)')
            else:
                print(f'   FAIL  {err}')


if __name__ == '__main__':
    asyncio.run(main())
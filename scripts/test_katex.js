// 用真实 KaTeX 引擎解析 HTML 中所有公式，统计失败项
const fs = require('fs');
const path = require('path');
const katex = require('katex');

// OUT 可配置：环境变量 OUT_DIR > 命令行第一个参数 > 当前目录
const OUT = process.env.OUT_DIR || process.argv[2] || process.cwd();
const files = fs.readdirSync(OUT).filter(f => f.endsWith('.html'));

const AC = '\\[';
const AC2 = '\\]';
const PC = '\\(';
const PC2 = '\\)';

for (const fn of files) {
  let h = fs.readFileSync(path.join(OUT, fn), 'utf8');
  // 剔除 <script>...</script> 段（里面是 renderMathInElement 配置，不是公式）
  h = h.replace(/<script[\s\S]*?<\/script>/g, '');
  let fail = 0, ok = 0;
  const errs = [];

  // 块级 \[ ... \]
  let idx = 0;
  while (true) {
    const s = h.indexOf(AC, idx);
    if (s < 0) break;
    const e = h.indexOf(AC2, s + 2);
    if (e < 0) break;
    let body = h.slice(s + 2, e);
    idx = e + 2;
    try {
      katex.renderToString(body, { displayMode: true, throwOnError: true, strict: false, trust: true });
      ok++;
    } catch (err) {
      fail++;
      if (errs.length < 8) errs.push('BLOCK: ' + body.slice(0, 90).replace(/\n/g, ' '));
    }
  }

  // 行内 \( ... \)
  idx = 0;
  while (true) {
    const s = h.indexOf(PC, idx);
    if (s < 0) break;
    const e = h.indexOf(PC2, s + 2);
    if (e < 0) break;
    let body = h.slice(s + 2, e);
    idx = e + 2;
    try {
      katex.renderToString(body, { displayMode: false, throwOnError: true, strict: false, trust: true });
      ok++;
    } catch (err) {
      fail++;
      if (errs.length < 8) errs.push('INLINE: ' + body.slice(0, 90).replace(/\n/g, ' '));
    }
  }

  console.log('===', fn);
  console.log('   渲染成功:', ok, ' 失败:', fail);
  errs.forEach(e => console.log('     ', e));
}

/**
 * PDF生成模块
 * 使用Electron的printToPDF功能生成真正的PDF文件
 * 使用KaTeX渲染LaTeX公式（完整CSS支持）
 */

import { marked } from 'marked'
import katex from 'katex'

/**
 * 页面尺寸配置
 */
const PAGE_SIZES = {
  'A4': 'A4',
  'A5': 'A5',
  'Letter': 'Letter'
}

/**
 * KaTeX配置
 */
const KATEX_OPTIONS = {
  throwOnError: false,
  strict: false,
  trust: true,
  output: 'html',
  fleqn: false,
  leqno: false
}

/**
 * LaTeX符号到Unicode的转换表（用于公式外的符号）
 */
const LATEX_TO_UNICODE = {
  // 大小比较符号
  '\\lt': '<', '\\gt': '>', '\\le': '≤', '\\ge': '≥',
  '\\leq': '≤', '\\geq': '≥', '\\leqq': '≦', '\\geqq': '≧',
  '\\lneq': '≨', '\\gneq': '≩', '\\nless': '≮', '\\ngtr': '≯',
  '\\nleq': '≰', '\\ngeq': '≱', '\\lesssim': '≲', '\\gtrsim': '≳',
  '\\lessapprox': '⪅', '\\gtrapprox': '⪆', '\\ll': '≪', '\\gg': '≫',
  '\\lessgtr': '≶', '\\gtrless': '≷', '\\prec': '≺', '\\succ': '≻',
  '\\preceq': '≼', '\\succeq': '≽',
  '\\greater': '>', '\\less': '<',

  // 运算符
  '\\pm': '±', '\\mp': '∓', '\\times': '×', '\\div': '÷',
  '\\cdot': '·', '\\ast': '∗', '\\star': '★',
  '\\neq': '≠', '\\ne': '≠', '\\approx': '≈', '\\equiv': '≡',
  '\\sim': '∼', '\\propto': '∝', '\\cong': '≅',

  // 集合符号
  '\\in': '∈', '\\notin': '∉', '\\subset': '⊂', '\\supset': '⊃',
  '\\subseteq': '⊆', '\\supseteq': '⊇', '\\cup': '∪', '\\cap': '∩',
  '\\emptyset': '∅', '\\varnothing': '∅',

  // 逻辑符号
  '\\forall': '∀', '\\exists': '∃', '\\neg': '¬', '\\land': '∧', '\\lor': '∨',

  // 箭头
  '\\rightarrow': '→', '\\leftarrow': '←', '\\to': '→', '\\gets': '←',
  '\\Rightarrow': '⇒', '\\Leftarrow': '⇐', '\\leftrightarrow': '↔', '\\Leftrightarrow': '⇔',

  // 其他符号
  '\\infty': '∞', '\\partial': '∂', '\\nabla': '∇', '\\hbar': 'ℏ',
  '\\angle': '∠', '\\triangle': '△', '\\square': '□', '\\circ': '○',
  '\\bullet': '•', '\\dots': '…', '\\ldots': '…', '\\cdots': '⋯',
  '\\vdots': '⋮', '\\ddots': '⋱', '\\parallel': '∥', '\\perp': '⊥',
  '\\vert': '|', '\\Vert': '‖',

  // 希腊字母
  '\\alpha': 'α', '\\beta': 'β', '\\gamma': 'γ', '\\delta': 'δ',
  '\\epsilon': 'ε', '\\zeta': 'ζ', '\\eta': 'η', '\\theta': 'θ',
  '\\iota': 'ι', '\\kappa': 'κ', '\\lambda': 'λ', '\\mu': 'μ',
  '\\nu': 'ν', '\\xi': 'ξ', '\\pi': 'π', '\\rho': 'ρ',
  '\\sigma': 'σ', '\\tau': 'τ', '\\upsilon': 'υ', '\\phi': 'φ',
  '\\chi': 'χ', '\\psi': 'ψ', '\\omega': 'ω',
  '\\varepsilon': 'ε', '\\varphi': 'φ', '\\vartheta': 'θ',
  '\\Gamma': 'Γ', '\\Delta': 'Δ', '\\Theta': 'Θ', '\\Lambda': 'Λ',
  '\\Xi': 'Ξ', '\\Pi': 'Π', '\\Sigma': 'Σ', '\\Phi': 'Φ',
  '\\Psi': 'Ψ', '\\Omega': 'Ω',

  // 求和积分等
  '\\sum': '∑', '\\int': '∫', '\\prod': '∏', '\\coprod': '∐',
  '\\lim': 'lim', '\\log': 'log', '\\sin': 'sin', '\\cos': 'cos',
  '\\tan': 'tan', '\\exp': 'exp',
}

/**
 * 字体配置
 */
const FONT_FAMILY_CSS = {
  'SimSun': 'SimSun, "宋体", serif',
  'Microsoft YaHei': '"Microsoft YaHei", "微软雅黑", sans-serif',
  'Arial': 'Arial, sans-serif',
  'Times New Roman': '"Times New Roman", serif'
}

/**
 * 完整的KaTeX CSS（使用系统字体替代）
 */
const KATEX_FULL_CSS = `
/* KaTeX基础样式 - 使用系统字体 */
.katex {
  font: normal 1.21em 'Cambria Math', 'Times New Roman', KaTeX_Main, serif;
  line-height: 1.2;
  position: relative;
  text-indent: 0;
  text-rendering: auto;
}
.katex * {
  -ms-high-contrast-adjust: none !important;
  border-color: currentColor;
}
.katex .katex-mathml {
  clip: rect(1px, 1px, 1px, 1px);
  border: 0;
  height: 1px;
  overflow: hidden;
  padding: 0;
  position: absolute;
  width: 1px;
}
.katex .katex-html > .newline { display: block; }
.katex .base {
  position: relative;
  white-space: nowrap;
  width: min-content;
  display: inline-block;
}
.katex .strut { display: inline-block; }
.katex .textbf { font-weight: 700; }
.katex .textit { font-style: italic; }
.katex .textrm { font-family: 'Cambria Math', 'Times New Roman', serif; }
.katex .textsf { font-family: 'Cambria Math', sans-serif; }
.katex .texttt { font-family: 'Cambria Math', monospace; }
.katex .mathnormal { font-family: 'Cambria Math', serif; font-style: italic; }
.katex .mathit { font-family: 'Cambria Math', serif; font-style: italic; }
.katex .mathrm { font-style: normal; }
.katex .mathbf { font-family: 'Cambria Math', serif; font-weight: 700; }
.katex .boldsymbol { font-family: 'Cambria Math', serif; font-style: italic; font-weight: 700; }
.katex .amsrm, .katex .mathbb, .katex .textbb { font-family: 'Cambria Math', serif; }
.katex .mathcal { font-family: 'Cambria Math', serif; }
.katex .mathfrak, .katex .textfrak { font-family: 'Cambria Math', serif; }
.katex .mathboldfrak, .katex .textboldfrak { font-family: 'Cambria Math', serif; font-weight: 700; }
.katex .mathtt { font-family: 'Cambria Math', monospace; }
.katex .mathscr, .katex .textscr { font-family: 'Cambria Math', serif; }
.katex .mathsf, .katex .textsf { font-family: 'Cambria Math', sans-serif; }
.katex .mathboldsf, .katex .textboldsf { font-family: 'Cambria Math', sans-serif; font-weight: 700; }
.katex .mathitsf, .katex .mathsfit, .katex .textitsf { font-family: 'Cambria Math', sans-serif; font-style: italic; }
.katex .mainrm { font-family: 'Cambria Math', serif; font-style: normal; }

/* 下标上标关键样式 */
.katex .vlist-t {
  border-collapse: collapse;
  display: inline-table;
  table-layout: fixed;
}
.katex .vlist-r { display: table-row; }
.katex .vlist {
  display: table-cell;
  position: relative;
  vertical-align: bottom;
}
.katex .vlist > span {
  display: block;
  height: 0;
  position: relative;
}
.katex .vlist > span > span { display: inline-block; }
.katex .vlist > span > .pstrut {
  overflow: hidden;
  width: 0;
}
.katex .vlist-t2 { margin-right: -2px; }
.katex .vlist-s {
  display: table-cell;
  font-size: 1px;
  min-width: 2px;
  vertical-align: bottom;
  width: 2px;
}
.katex .vbox {
  align-items: baseline;
  display: inline-flex;
  flex-direction: column;
}
.katex .hbox { width: 100%; display: inline-flex; flex-direction: row; }
.katex .thinbox { display: inline-flex; flex-direction: row; max-width: 0; width: 0; }
.katex .msupsub { text-align: left; }

/* 分数样式 */
.katex .mfrac > span > span { text-align: center; }
.katex .mfrac .frac-line {
  border-bottom-style: solid;
  display: inline-block;
  width: 100%;
  min-height: 1px;
}

/* 其他样式 */
.katex .mspace { display: inline-block; }
.katex .smash { display: inline; line-height: 0; }
.katex .clap, .katex .llap, .katex .rlap { position: relative; width: 0; }
.katex .clap > .inner, .katex .llap > .inner, .katex .rlap > .inner { position: absolute; }
.katex .clap > .fix, .katex .llap > .fix, .katex .rlap > .fix { display: inline-block; }
.katex .llap > .inner { right: 0; }
.katex .clap > .inner, .katex .rlap > .inner { left: 0; }
.katex .clap > .inner > span { margin-left: -50%; margin-right: 50%; }
.katex .rule { border: 0 solid; display: inline-block; position: relative; }
.katex .hline, .katex .overline .overline-line, .katex .underline .underline-line {
  border-bottom-style: solid;
  display: inline-block;
  width: 100%;
  min-height: 1px;
}
.katex .hdashline { border-bottom-style: dashed; display: inline-block; width: 100%; min-height: 1px; }
.katex .sqrt > .root { margin-left: .2777777778em; margin-right: -.5555555556em; }

/* 尺寸调整 - 下标上标需要这些 */
.katex .sizing.reset-size6.size1 { font-size: 0.5em; }
.katex .sizing.reset-size6.size2 { font-size: 0.6em; }
.katex .sizing.reset-size6.size3 { font-size: 0.7em; }
.katex .sizing.reset-size6.size4 { font-size: 0.8em; }
.katex .sizing.reset-size6.size5 { font-size: 0.9em; }
.katex .sizing.reset-size6.size6 { font-size: 1em; }
.katex .sizing.reset-size6.size7 { font-size: 1.2em; }
.katex .sizing.reset-size6.size8 { font-size: 1.44em; }
.katex .sizing.reset-size6.size9 { font-size: 1.728em; }
.katex .sizing.reset-size6.size10 { font-size: 2.074em; }
.katex .sizing.reset-size6.size11 { font-size: 2.488em; }
.katex .fontsize-ensurer.reset-size6.size1 { font-size: 0.5em; }
.katex .fontsize-ensurer.reset-size6.size2 { font-size: 0.6em; }
.katex .fontsize-ensurer.reset-size6.size3 { font-size: 0.7em; }
.katex .fontsize-ensurer.reset-size6.size4 { font-size: 0.8em; }
.katex .fontsize-ensurer.reset-size6.size5 { font-size: 0.9em; }
.katex .fontsize-ensurer.reset-size6.size6 { font-size: 1em; }
.katex .fontsize-ensurer.reset-size6.size7 { font-size: 1.2em; }
.katex .fontsize-ensurer.reset-size6.size8 { font-size: 1.44em; }
.katex .fontsize-ensurer.reset-size6.size9 { font-size: 1.728em; }
.katex .fontsize-ensurer.reset-size6.size10 { font-size: 2.074em; }
.katex .fontsize-ensurer.reset-size6.size11 { font-size: 2.488em; }

/* 更多尺寸 */
.katex .sizing.reset-size1.size1 { font-size: 1em; }
.katex .sizing.reset-size1.size2 { font-size: 1.2em; }
.katex .sizing.reset-size1.size3 { font-size: 1.4em; }
.katex .sizing.reset-size1.size4 { font-size: 1.6em; }
.katex .sizing.reset-size1.size5 { font-size: 1.8em; }
.katex .sizing.reset-size1.size6 { font-size: 2em; }
.katex .sizing.reset-size1.size7 { font-size: 2.4em; }
.katex .sizing.reset-size1.size8 { font-size: 2.88em; }
.katex .sizing.reset-size1.size9 { font-size: 3.456em; }
.katex .sizing.reset-size1.size10 { font-size: 4.148em; }
.katex .sizing.reset-size1.size11 { font-size: 4.976em; }

/* mtight 样式 - 用于下标上标 */
.katex .mtight { font-family: 'Cambria Math', serif; font-size: 0.7em; }

/* 其他元素样式 */
.katex .delimsizing.size1 { font-family: 'Cambria Math', serif; }
.katex .delimsizing.size2 { font-family: 'Cambria Math', serif; }
.katex .delimsizing.size3 { font-family: 'Cambria Math', serif; }
.katex .delimsizing.size4 { font-family: 'Cambria Math', serif; }
.katex .nulldelimiter { display: inline-block; width: .12em; }
.katex .delimcenter, .katex .op-symbol { position: relative; }
.katex .op-symbol.small-op { font-family: 'Cambria Math', serif; }
.katex .op-symbol.large-op { font-family: 'Cambria Math', serif; }
.katex .accent > .vlist-t, .katex .op-limits > .vlist-t { text-align: center; }
.katex .accent .accent-body { position: relative; }
.katex .accent .accent-body:not(.accent-full) { width: 0; }
.katex .overlay { display: block; }
.katex .mtable .vertical-separator { display: inline-block; min-width: 1px; }
.katex .mtable .arraycolsep { display: inline-block; }
.katex .mtable .col-align-c > .vlist-t { text-align: center; }
.katex .mtable .col-align-l > .vlist-t { text-align: left; }
.katex .mtable .col-align-r > .vlist-t { text-align: right; }
.katex .svg-align { text-align: left; }
.katex svg {
  fill: currentColor;
  stroke: currentColor;
  display: block;
  height: inherit;
  position: absolute;
  width: 100%;
}
.katex svg path { stroke: none; }
.katex img { border-style: none; max-height: none; max-width: none; min-height: 0; min-width: 0; }
.katex .stretchy { display: block; overflow: hidden; position: relative; width: 100%; }
.katex .stretchy:before, .katex .stretchy:after { content: ""; }
.katex .hide-tail { overflow: hidden; position: relative; width: 100%; }
.katex .halfarrow-left { left: 0; overflow: hidden; position: absolute; width: 50.2%; }
.katex .halfarrow-right { overflow: hidden; position: absolute; right: 0; width: 50.2%; }
.katex .brace-left { left: 0; overflow: hidden; position: absolute; width: 25.1%; }
.katex .brace-center { left: 25%; overflow: hidden; position: absolute; width: 50%; }
.katex .brace-right { overflow: hidden; position: absolute; right: 0; width: 25.1%; }
.katex .x-arrow-pad { padding: 0 .5em; }
.katex .cd-arrow-pad { padding: 0 .55556em 0 .27778em; }
.katex .mover, .katex .munder, .katex .x-arrow { text-align: center; }
.katex .boxpad { padding: 0 .3em; }
.katex .fbox, .katex .fcolorbox { border: .04em solid; box-sizing: border-box; }
.katex .cancel-pad { padding: 0 .2em; }
.katex .cancel-lap { margin-left: -.2em; margin-right: -.2em; }
.katex .sout { border-bottom-style: solid; border-bottom-width: .08em; }
.katex .angl { border-right: .049em solid; border-top: .049em solid; box-sizing: border-box; margin-right: .03889em; }
.katex .anglpad { padding: 0 .03889em; }

/* 块级公式样式 */
.katex-display { display: block; margin: 1em 0; text-align: center; }
.katex-display > .katex { display: block; text-align: center; white-space: nowrap; }
.katex-display > .katex > .katex-html { display: block; position: relative; }
.katex-display > .katex > .katex-html > .tag { position: absolute; right: 0; }
.katex-display.leqno > .katex > .katex-html > .tag { left: 0; right: auto; }
.katex-display.fleqn > .katex { padding-left: 2em; text-align: left; }

/* mord, mbin, mrel 等基础元素 */
.katex .mord, .katex .mbin, .katex .mrel, .katex .mopen, .katex .mclose, .katex .mpunct {
  display: inline-block;
  font-family: 'Cambria Math', serif;
}

/* 打印样式 */
@media print {
  .katex { print-color-adjust: exact; -webkit-print-color-adjust: exact; }
  .word-formula-block { page-break-inside: avoid; }
}
`

/**
 * 渲染单个LaTeX公式为HTML
 */
function renderLatexFormula(latex, isBlock = false) {
  try {
    const html = katex.renderToString(latex.trim(), {
      ...KATEX_OPTIONS,
      displayMode: isBlock
    })
    return html
  } catch (error) {
    console.warn('KaTeX渲染失败:', error.message, '公式:', latex)
    const escaped = latex.replace(/[<>&]/g, c => ({ '<': '&lt;', '>': '&gt;', '&': '&amp;' }[c]))
    return `<span class="formula-error" style="color: #c00; font-family: 'Cambria Math', 'Times New Roman', serif; font-style: italic;">${escaped}</span>`
  }
}

/**
 * 提取并渲染LaTeX公式
 */
function extractAndRenderFormulas(content) {
  if (!content) return ''

  let processed = content
  const formulaPlaceholders = []

  // 1. 处理 $$...$$ 块级公式
  processed = processed.replace(/\$\$(.*?)\$\$/gs, (match, latex) => {
    const placeholder = `%%FORMULA_BLOCK_${formulaPlaceholders.length}%%`
    const rendered = renderLatexFormula(latex.trim(), true)
    formulaPlaceholders.push({ placeholder, rendered, isBlock: true })
    return placeholder
  })

  // 2. 处理 \[...\] 块级公式
  processed = processed.replace(/\\\[(.*?)\\\]/gs, (match, latex) => {
    const placeholder = `%%FORMULA_BLOCK_${formulaPlaceholders.length}%%`
    const rendered = renderLatexFormula(latex.trim(), true)
    formulaPlaceholders.push({ placeholder, rendered, isBlock: true })
    return placeholder
  })

  // 3. 处理 $...$ 行内公式
  processed = processed.replace(/\$(?!\$)([^\$\n]+?)\$(?!\$)/g, (match, latex) => {
    const placeholder = `%%FORMULA_INLINE_${formulaPlaceholders.length}%%`
    const rendered = renderLatexFormula(latex.trim(), false)
    formulaPlaceholders.push({ placeholder, rendered, isBlock: false })
    return placeholder
  })

  // 4. 处理 \(...\) 行内公式
  processed = processed.replace(/\\\((.*?)\\\)/g, (match, latex) => {
    const placeholder = `%%FORMULA_INLINE_${formulaPlaceholders.length}%%`
    const rendered = renderLatexFormula(latex.trim(), false)
    formulaPlaceholders.push({ placeholder, rendered, isBlock: false })
    return placeholder
  })

  // 5. 转换公式外的LaTeX符号为Unicode
  const sortedLatexCommands = Object.entries(LATEX_TO_UNICODE)
    .sort((a, b) => b[0].length - a[0].length)

  for (const [latex, unicode] of sortedLatexCommands) {
    const escapedLatex = latex.replace(/\\/g, '\\\\')
    const pattern = new RegExp(escapedLatex + '(?![a-zA-Z])', 'g')
    processed = processed.replace(pattern, unicode)
  }

  // 6. 恢复公式占位符
  for (const { placeholder, rendered, isBlock } of formulaPlaceholders) {
    if (isBlock) {
      // 块级公式使用 katex-display 样式
      processed = processed.replace(placeholder, `<div class="katex-display">${rendered}</div>`)
    } else {
      processed = processed.replace(placeholder, rendered)
    }
  }

  return processed
}

/**
 * 生成PDF HTML模板
 */
function generatePdfHtml(content, styleConfig) {
  const fontFamily = FONT_FAMILY_CSS[styleConfig.fontFamily] || 'SimSun, serif'
  const fontSize = styleConfig.fontSize || 12

  const marginTop = (styleConfig.marginTop || 25) / 10
  const marginBottom = (styleConfig.marginBottom || 25) / 10
  const marginLeft = (styleConfig.marginLeft || 20) / 10
  const marginRight = (styleConfig.marginRight || 20) / 10

  // 提取并渲染LaTeX公式
  const processedContent = extractAndRenderFormulas(content)

  // 解析Markdown
  marked.setOptions({
    gfm: true,
    breaks: true
  })
  const htmlContent = marked.parse(processedContent)

  return `
<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8">
  <title>Document</title>
  <style>
    @page {
      size: ${styleConfig.pageSize || 'A4'};
      margin: ${marginTop}cm ${marginRight}cm ${marginBottom}cm ${marginLeft}cm;
    }

    @media print {
      body { print-color-adjust: exact; -webkit-print-color-adjust: exact; }
      .katex-display { page-break-inside: avoid; }
    }

    * { margin: 0; padding: 0; box-sizing: border-box; }

    body {
      font-family: ${fontFamily};
      font-size: ${fontSize}pt;
      line-height: 1.6;
      color: #333;
      background: white;
      padding: 20px;
    }

    .header { text-align: center; font-size: ${(fontSize * 0.9)}pt; color: #666; border-bottom: 1px solid #ddd; padding-bottom: 10px; margin-bottom: 20px; }
    .footer { text-align: center; font-size: ${(fontSize * 0.9)}pt; color: #666; border-top: 1px solid #ddd; padding-top: 10px; margin-top: 20px; }

    h1 { font-size: ${(fontSize * 2.33)}pt; font-weight: bold; margin-bottom: 16px; page-break-after: avoid; }
    h2 { font-size: ${(fontSize * 2)}pt; font-weight: bold; margin-bottom: 14px; page-break-after: avoid; }
    h3 { font-size: ${(fontSize * 1.67)}pt; font-weight: bold; margin-bottom: 12px; page-break-after: avoid; }
    h4 { font-size: ${(fontSize * 1.5)}pt; font-weight: bold; margin-bottom: 10px; page-break-after: avoid; }

    p { margin-bottom: 12px; line-height: 1.6; orphans: 3; widows: 3; }

    ul, ol { margin-bottom: 12px; padding-left: 24px; }
    li { margin-bottom: 4px; }

    pre { background: #f6f8fa; padding: 16px; border-radius: 6px; overflow-x: auto; font-family: 'Consolas', 'Monaco', monospace; font-size: ${(fontSize * 0.9)}pt; page-break-inside: avoid; }
    code { background: #f6f8fa; padding: 2px 6px; border-radius: 4px; font-family: 'Consolas', 'Monaco', monospace; }
    pre code { background: transparent; padding: 0; }

    table { border-collapse: collapse; width: 100%; margin-bottom: 16px; page-break-inside: avoid; }
    th, td { border: 1px solid #dcdfe6; padding: 8px 12px; text-align: left; }
    th { background: #f5f7fa; font-weight: 600; text-align: center; }

    blockquote { border-left: 4px solid #ddd; margin: 16px 0; padding: 8px 16px; color: #666; font-style: italic; background-color: #f9f9f9; }
    img { max-width: 100%; height: auto; display: block; margin: 16px auto; }
    a { color: #0366d6; text-decoration: none; }
    a:hover { text-decoration: underline; }
    hr { border: none; border-top: 1px solid #ddd; margin: 20px 0; }

    /* 嵌入完整KaTeX CSS */
    ${KATEX_FULL_CSS}
  </style>
</head>
<body>
  ${styleConfig.header ? `<div class="header">${styleConfig.header}</div>` : ''}
  <div class="content">${htmlContent}</div>
  ${styleConfig.footer ? `<div class="footer">${styleConfig.footer}</div>` : ''}
</body>
</html>`
}

/**
 * 生成PDF文件
 */
export async function generatePdf(markdownContent, styleConfig = {}) {
  try {
    const htmlContent = generatePdfHtml(markdownContent, styleConfig)

    if (window.electronAPI && window.electronAPI.generatePdf) {
      const pdfOptions = {
        pageSize: PAGE_SIZES[styleConfig.pageSize] || 'A4',
        marginTop: (styleConfig.marginTop || 25) / 10,
        marginBottom: (styleConfig.marginBottom || 25) / 10,
        marginLeft: (styleConfig.marginLeft || 20) / 10,
        marginRight: (styleConfig.marginRight || 20) / 10
      }

      const pdfArray = await window.electronAPI.generatePdf(htmlContent, pdfOptions)
      const pdfBuffer = new Uint8Array(pdfArray)
      return new Blob([pdfBuffer], { type: 'application/pdf' })
    } else {
      console.warn('Electron API不可用')
      const printWindow = window.open('', '_blank')
      if (printWindow) {
        printWindow.document.write(htmlContent)
        printWindow.document.close()
        alert('请使用浏览器打印功能保存为PDF (Ctrl+P)')
      }
      return new Blob([htmlContent], { type: 'text/html' })
    }
  } catch (error) {
    console.error('PDF生成错误:', error)
    throw error
  }
}

/**
 * 获取PDF HTML内容
 */
export function getPdfHtml(markdownContent, styleConfig = {}) {
  return generatePdfHtml(markdownContent, styleConfig)
}
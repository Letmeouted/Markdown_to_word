/**
 * Markdown解析器
 * 支持LaTeX公式提取和转换
 */

import { marked } from 'marked'
import katex from 'katex'

// KaTeX配置 - 模拟Word公式样式
const KATEX_OPTIONS = {
  throwOnError: false,
  strict: false,
  trust: true,
  output: 'html',
  // 使用更接近Word的字体设置
  fleqn: false,  // 公式居中（Word默认）
  leqno: false,  // 公式编号在右侧
  macros: {
    // 定义常用宏
    "\\R": "\\mathbb{R}",
    "\\N": "\\mathbb{N}",
    "\\Z": "\\mathbb{Z}",
    "\\Q": "\\mathbb{Q}",
    "\\C": "\\mathbb{C}",
  }
}

/**
 * 提取所有LaTeX公式
 * @param {string} content Markdown内容
 * @returns {Object} 包含处理后的内容和公式信息
 */
function extractAndProtectFormulas(content) {
  if (!content) return { processed: '', blockFormulas: [], inlineFormulas: [] }

  let processed = content
  const blockFormulas = []
  const inlineFormulas = []

  // ===== 第一步：处理块级公式 =====

  // 处理 $$...$$ 格式的块级公式
  // 使用更精确的匹配：$$ 后面跟着任意内容（直到遇到结束的 $$）
  let match
  const dollarBlockPattern = /\$\$(.*?)\$\$/gs  // s 标志允许跨行匹配
  while ((match = dollarBlockPattern.exec(processed)) !== null) {
    blockFormulas.push({
      placeholder: `%%BLOCK_FORMULA_${blockFormulas.length}%%`,
      formula: match[1].trim(),
      original: match[0]
    })
  }

  // 处理 \[...\] 格式的块级公式
  // 匹配 \[ 到 \] 之间的所有内容
  const bracketBlockPattern = /\\\[([\s\S]*?)\\\]/g
  while ((match = bracketBlockPattern.exec(processed)) !== null) {
    blockFormulas.push({
      placeholder: `%%BLOCK_FORMULA_${blockFormulas.length}%%`,
      formula: match[1].trim(),
      original: match[0]
    })
  }

  // 替换所有块级公式为占位符
  for (const item of blockFormulas) {
    // 使用 split + join 方式替换，避免正则替换的问题
    processed = processed.split(item.original).join(item.placeholder)
  }

  // ===== 第二步：处理行内公式 =====

  // 处理 $...$ 格式的行内公式
  // 重要：只匹配单个 $ 包围的内容，避免误匹配 $$ 或表格中的 |
  // 使用贪婪匹配的变体，匹配直到下一个 $ 符号（但不能是 $$）
  const dollarInlinePattern = /\$(?!\$)([^\$\n\\]+?|\S.*?\S)\$(?!\$)/g
  while ((match = dollarInlinePattern.exec(processed)) !== null) {
    inlineFormulas.push({
      placeholder: `%%INLINE_FORMULA_${inlineFormulas.length}%%`,
      formula: match[1].trim(),
      original: match[0]
    })
  }

  // 处理 \(...\) 格式的行内公式
  // 匹配 \( 到 \) 之间的内容，支持括号嵌套（通过非贪婪匹配）
  const bracketInlinePattern = /\\\((.*?)\\\)/g
  while ((match = bracketInlinePattern.exec(processed)) !== null) {
    inlineFormulas.push({
      placeholder: `%%INLINE_FORMULA_${inlineFormulas.length}%%`,
      formula: match[1].trim(),
      original: match[0]
    })
  }

  // 替换所有行内公式为占位符
  for (const item of inlineFormulas) {
    processed = processed.split(item.original).join(item.placeholder)
  }

  return { processed, blockFormulas, inlineFormulas }
}

/**
 * 使用KaTeX渲染公式（Word样式）
 * @param {string} formula LaTeX公式
 * @param {boolean} isBlock 是否为块级公式
 * @returns {string} 渲染后的HTML
 */
function renderFormulaWithKatex(formula, isBlock = false) {
  try {
    const html = katex.renderToString(formula, {
      ...KATEX_OPTIONS,
      displayMode: isBlock
    })
    return html
  } catch (error) {
    console.warn('KaTeX render error:', error.message, 'Formula:', formula)
    // 显示原始LaTeX代码作为备用
    return `<span style="color: #c00; font-family: 'Cambria Math', 'Times New Roman', serif; font-style: italic;">${escapeHtml(formula)}</span>`
  }
}

/**
 * HTML转义
 */
function escapeHtml(text) {
  const map = {
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#039;'
  }
  return String(text).replace(/[&<>"']/g, m => map[m])
}

/**
 * 恢复并渲染公式（Word样式）
 * @param {string} html 包含占位符的HTML
 * @param {Array} blockFormulas 块级公式列表
 * @param {Array} inlineFormulas 行内公式列表
 * @returns {string} 渲染后的HTML
 */
function restoreAndRenderFormulas(html, blockFormulas, inlineFormulas) {
  let result = html

  // 恢复并渲染块级公式 - Word样式：居中、专业字体、无明显背景
  result = result.replace(/%%BLOCK_FORMULA_(\d+)%%/g, (match, index) => {
    const formula = blockFormulas[parseInt(index)]
    if (formula) {
      const rendered = renderFormulaWithKatex(formula.formula, true)
      // Word块级公式样式：居中、上下间距、无背景色
      return `<div class="word-formula-block" style="display: block; text-align: center; margin: 12pt 0; padding: 6pt 0; font-family: 'Cambria Math', 'Times New Roman', serif;">${rendered}</div>`
    }
    return match
  })

  // 恢复并渲染行内公式 - Word样式：专业数学字体
  result = result.replace(/%%INLINE_FORMULA_(\d+)%%/g, (match, index) => {
    const formula = inlineFormulas[parseInt(index)]
    if (formula) {
      const rendered = renderFormulaWithKatex(formula.formula, false)
      // Word行内公式样式：斜体数学字体
      return `<span class="word-formula-inline" style="font-family: 'Cambria Math', 'Times New Roman', serif; font-style: italic;">${rendered}</span>`
    }
    return match
  })

  return result
}

/**
 * 解析Markdown内容
 * @param {string} content Markdown内容
 * @param {string} mode 模式: 'preview' | 'docx' | 'pdf'
 * @returns {string} 解析结果
 */
export function parseMarkdown(content, mode = 'preview') {
  if (!content) return ''

  // 提取并保护公式
  const { processed, blockFormulas, inlineFormulas } = extractAndProtectFormulas(content)

  // 配置marked
  marked.setOptions({
    gfm: true,
    breaks: true
  })

  // 使用marked解析（公式已被保护，不会被破坏）
  const html = marked.parse(processed)

  // 后处理：恢复并渲染公式
  if (mode === 'preview') {
    return restoreAndRenderFormulas(html, blockFormulas, inlineFormulas)
  }

  // 其他模式返回原始解析结果
  return html
}

/**
 * 提取所有LaTeX公式（用于其他模块）
 * @param {string} content Markdown内容
 * @returns {Array} 公式列表
 */
export function extractLatexFormulas(content) {
  const formulas = []

  // 匹配块级公式 $$...$$
  const dollarBlockPattern = /\$\$(.*?)\$\$/gs
  let match
  while ((match = dollarBlockPattern.exec(content)) !== null) {
    formulas.push({
      type: 'block',
      content: match[1].trim(),
      start: match.index,
      end: match.index + match[0].length,
      raw: match[0]
    })
  }

  // 匹配 \[...\] 格式的块级公式
  const bracketBlockPattern = /\\\[([\s\S]*?)\\\]/g
  while ((match = bracketBlockPattern.exec(content)) !== null) {
    formulas.push({
      type: 'block',
      content: match[1].trim(),
      start: match.index,
      end: match.index + match[0].length,
      raw: match[0]
    })
  }

  // 匹配行内公式 $...$
  const dollarInlinePattern = /\$(?!\$)([^\$\n\\]+?|\S.*?\S)\$(?!\$)/g
  while ((match = dollarInlinePattern.exec(content)) !== null) {
    formulas.push({
      type: 'inline',
      content: match[1].trim(),
      start: match.index,
      end: match.index + match[0].length,
      raw: match[0]
    })
  }

  // 匹配 \(...\) 格式的行内公式
  const bracketInlinePattern = /\\\((.*?)\\\)/g
  while ((match = bracketInlinePattern.exec(content)) !== null) {
    formulas.push({
      type: 'inline',
      content: match[1].trim(),
      start: match.index,
      end: match.index + match[0].length,
      raw: match[0]
    })
  }

  return formulas.sort((a, b) => a.start - b.start)
}

/**
 * 将Markdown内容转换为AST结构（用于Word生成）
 * @param {string} content Markdown内容
 * @returns {Array} AST节点数组
 */
export function parseToAST(content) {
  const formulas = extractLatexFormulas(content)
  const tokens = marked.lexer(content)

  const ast = []

  for (const token of tokens) {
    if (token.type === 'paragraph' || token.type === 'text') {
      const text = token.text || token.raw || ''
      ast.push({
        type: 'paragraph',
        content: text,
        formulas: formulas.filter(f => text.includes(f.raw))
      })
    } else if (token.type === 'heading') {
      ast.push({
        type: 'heading',
        level: token.depth,
        content: token.text
      })
    } else if (token.type === 'list') {
      ast.push({
        type: 'list',
        ordered: token.start !== undefined,
        items: token.items.map(item => ({
          content: item.text
        }))
      })
    } else if (token.type === 'table') {
      ast.push({
        type: 'table',
        header: token.header,
        rows: token.rows
      })
    } else if (token.type === 'code') {
      ast.push({
        type: 'code',
        lang: token.lang || '',
        content: token.text
      })
    } else if (token.type === 'blockquote') {
      ast.push({
        type: 'blockquote',
        content: token.text
      })
    }
  }

  return ast
}
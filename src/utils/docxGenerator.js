/**
 * Word文档生成模块
 * 使用docx库生成.docx文件
 */

import {
  Document,
  Paragraph,
  TextRun,
  HeadingLevel,
  AlignmentType,
  Packer,
  Table,
  TableRow,
  TableCell,
  WidthType,
  BorderStyle,
  LevelFormat,
  ExternalHyperlink,
  Math as DocxMath,
  MathRun,
  MathFraction,
  MathSubScript,
  MathSuperScript,
  MathSubSuperScript,
  MathRadical,
  MathSum,
  MathIntegral,
  MathAngledBrackets,
  MathRoundBrackets,
  MathSquareBrackets,
  MathCurlyBrackets,
  BorderStyle as HrBorderStyle
} from 'docx'

import { marked } from 'marked'

/**
 * 页面大小映射
 */
const PAGE_SIZES = {
  'A4': { width: 11906, height: 16838 },
  'A5': { width: 8419, height: 11906 },
  'Letter': { width: 12240, height: 15840 }
}

/**
 * 字体映射
 */
const FONT_MAP = {
  'SimSun': '宋体',
  'Microsoft YaHei': '微软雅黑',
  'Arial': 'Arial',
  'Times New Roman': 'Times New Roman'
}

/**
 * 根据样式配置创建文档选项
 */
function createDocumentOptions(styleConfig) {
  const pageSize = PAGE_SIZES[styleConfig.pageSize] || PAGE_SIZES['A4']
  // 确保边距值是数字类型
  const marginConvert = (mm) => Math.round(Number(mm || 25) * 56.692) // mm转twips

  return {
    sections: [{
      properties: {
        page: {
          size: {
            width: pageSize.width,
            height: pageSize.height
          },
          margins: {
            top: marginConvert(styleConfig.marginTop),
            bottom: marginConvert(styleConfig.marginBottom),
            left: marginConvert(styleConfig.marginLeft),
            right: marginConvert(styleConfig.marginRight)
          }
        }
      },
      children: []
    }],
    styles: {
      default: {
        document: {
          run: {
            font: FONT_MAP[styleConfig.fontFamily] || '宋体',
            size: Math.round(Number(styleConfig.fontSize || 12) * 2) // pt转half-pt
          }
        }
      },
      paragraphStyles: [
        {
          id: 'Heading1',
          name: 'Heading 1',
          basedOn: 'Normal',
          next: 'Normal',
          run: {
            size: 32,
            bold: true,
            font: FONT_MAP[styleConfig.fontFamily] || '宋体'
          }
        },
        {
          id: 'Heading2',
          name: 'Heading 2',
          basedOn: 'Normal',
          next: 'Normal',
          run: {
            size: 28,
            bold: true,
            font: FONT_MAP[styleConfig.fontFamily] || '宋体'
          }
        },
        {
          id: 'Heading3',
          name: 'Heading 3',
          basedOn: 'Normal',
          next: 'Normal',
          run: {
            size: 24,
            bold: true,
            font: FONT_MAP[styleConfig.fontFamily] || '宋体'
          }
        }
      ]
    }
  }
}

/**
 * 将Markdown内容转换为docx段落
 * 在解析前预处理公式，确保公式标记被正确移除
 */
function convertToDocxParagraphs(content, styleConfig) {
  const paragraphs = []

  // 预处理：将公式转换为占位符，防止markdown解析器干扰
  const { processed, blockFormulas, inlineFormulas } = preprocessLatexForDocx(content)

  // 使用 marked 解析预处理后的内容
  marked.setOptions({ gfm: true, breaks: true })
  const tokens = marked.lexer(processed)

  // 页眉
  if (styleConfig.header) {
    paragraphs.push(
      new Paragraph({
        children: [
          new TextRun({
            text: styleConfig.header,
            alignment: AlignmentType.CENTER
          })
        ]
      })
    )
    // 添加分隔
    paragraphs.push(new Paragraph({ children: [] }))
  }

  for (const token of tokens) {
    switch (token.type) {
      case 'heading':
        paragraphs.push(createHeading(token, blockFormulas, inlineFormulas))
        // 标题后添加空段落作为分隔
        paragraphs.push(new Paragraph({ spacing: { after: 120 }, children: [] }))
        break

      case 'paragraph':
        const pParagraphs = createParagraphFromToken(token, blockFormulas, inlineFormulas)
        paragraphs.push(...pParagraphs)
        // 段落之间添加适当间距
        if (pParagraphs.length > 0) {
          paragraphs.push(new Paragraph({ spacing: { after: 100 }, children: [] }))
        }
        break

      case 'list':
        paragraphs.push(...createNestedList(token, blockFormulas, inlineFormulas))
        // 列表后添加分隔
        paragraphs.push(new Paragraph({ spacing: { after: 120 }, children: [] }))
        break

      case 'table':
        paragraphs.push(...createTable(token, blockFormulas, inlineFormulas))
        // 表格后添加分隔
        paragraphs.push(new Paragraph({ spacing: { after: 120 }, children: [] }))
        break

      case 'code':
        paragraphs.push(createCodeBlock(token))
        // 代码块后添加分隔
        paragraphs.push(new Paragraph({ spacing: { after: 120 }, children: [] }))
        break

      case 'blockquote':
        paragraphs.push(createBlockquote(token, blockFormulas, inlineFormulas))
        // 引用块后添加分隔
        paragraphs.push(new Paragraph({ spacing: { after: 120 }, children: [] }))
        break

      case 'hr':
      case 'thematicBreak':
        // 分割线处理
        paragraphs.push(createHorizontalLine())
        break

      case 'space':
        // 空行 - 添加空段落作为分隔符
        paragraphs.push(new Paragraph({ spacing: { after: 60 }, children: [] }))
        break

      default:
        if (token.raw) {
          // 处理其他类型的token
          const runs = createTextRunsWithNativeFormulas(token.raw, inlineFormulas)
          paragraphs.push(new Paragraph({
            children: runs
          }))
        }
    }
  }

  // 页脚
  if (styleConfig.footer) {
    paragraphs.push(new Paragraph({ children: [] }))
    paragraphs.push(
      new Paragraph({
        children: [
          new TextRun({
            text: styleConfig.footer,
            alignment: AlignmentType.CENTER
          })
        ]
      })
    )
  }

  return paragraphs
}

/**
 * 预处理LaTeX公式，转换为占位符
 * 支持 $...$, $$...$$, \(...\), \[...\] 四种格式
 * 使用精确的正则表达式避免与markdown表格格式冲突
 */
function preprocessLatexForDocx(content) {
  if (!content) return { processed: '', blockFormulas: [], inlineFormulas: [] }

  let processed = content
  const blockFormulas = []
  const inlineFormulas = []

  // ===== 第一步：处理块级公式 =====

  // 处理 $$...$$ 格式的块级公式
  let match
  const dollarBlockPattern = /\$\$(.*?)\$\$/gs
  while ((match = dollarBlockPattern.exec(processed)) !== null) {
    blockFormulas.push({
      placeholder: `%%BLOCK_FORMULA_${blockFormulas.length}%%`,
      content: match[0],
      formula: match[1].trim()
    })
  }

  // 处理 \[...\] 格式的块级公式
  const bracketBlockPattern = /\\\[([\s\S]*?)\\\]/g
  while ((match = bracketBlockPattern.exec(processed)) !== null) {
    blockFormulas.push({
      placeholder: `%%BLOCK_FORMULA_${blockFormulas.length}%%`,
      content: match[0],
      formula: match[1].trim()
    })
  }

  // 替换所有块级公式为占位符
  for (const item of blockFormulas) {
    processed = processed.split(item.content).join(item.placeholder)
  }

  // ===== 第二步：处理行内公式 =====

  // 处理 $...$ 格式的行内公式
  const dollarInlinePattern = /\$(?!\$)([^\$\n\\]+?|\S.*?\S)\$(?!\$)/g
  while ((match = dollarInlinePattern.exec(processed)) !== null) {
    inlineFormulas.push({
      placeholder: `%%INLINE_FORMULA_${inlineFormulas.length}%%`,
      content: match[0],
      formula: match[1].trim()
    })
  }

  // 处理 \(...\) 格式的行内公式
  const bracketInlinePattern = /\\\((.*?)\\\)/g
  while ((match = bracketInlinePattern.exec(processed)) !== null) {
    inlineFormulas.push({
      placeholder: `%%INLINE_FORMULA_${inlineFormulas.length}%%`,
      content: match[0],
      formula: match[1].trim()
    })
  }

  // 替换所有行内公式为占位符
  for (const item of inlineFormulas) {
    processed = processed.split(item.content).join(item.placeholder)
  }

  return { processed, blockFormulas, inlineFormulas }
}

/**
 * 在文本中恢复公式（转换为Unicode显示）
 */
function restoreFormulasInText(text, blockFormulas, inlineFormulas) {
  // 确保 text 是字符串
  if (text === null || text === undefined) {
    return ''
  }
  let result = String(text)

  // 恢复块级公式
  result = result.replace(/%%BLOCK_FORMULA_(\d+)%%/g, (match, index) => {
    const formula = blockFormulas[parseInt(index)]
    if (formula) {
      return formatLatexForDisplay(formula.formula)
    }
    return match
  })

  // 恢复行内公式
  result = result.replace(/%%INLINE_FORMULA_(\d+)%%/g, (match, index) => {
    const formula = inlineFormulas[parseInt(index)]
    if (formula) {
      return formatLatexForDisplay(formula.formula)
    }
    return match
  })

  return result
}

/**
 * 解析内联样式（加粗、斜体、代码、链接、删除线等）
 * 返回TextRun数组，正确移除Markdown标记符号
 * 支持换行符处理（将\n转换为Word的break）
 */
function parseInlineStyles(text) {
  // 确保 text 是字符串
  if (text === null || text === undefined) {
    return []
  }
  const runs = []
  let remaining = String(text)

  // 首先处理换行符 - 将文本按换行符分割
  // 在Word中，换行符需要用特殊方式处理

  // 匹配模式顺序很重要，先匹配复杂模式
  const patterns = [
    // 加粗斜体 ***text*** 或 ___text___
    { regex: /\*\*\*(.+?)\*\*\*|___(.+?)___/g, style: { bold: true, italics: true }, groupIndex: 1 },
    // 加粗 **text** 或 __text__
    { regex: /\*\*(.+?)\*\*|__(.+?)__/g, style: { bold: true }, groupIndex: 1 },
    // 行内代码 `code`
    { regex: /`([^`]+)`/g, style: { font: 'Consolas', size: 20 }, groupIndex: 1 },
    // 删除线 ~~text~~
    { regex: /~~(.+?)~~/g, style: { strike: true }, groupIndex: 1 },
    // 链接 [text](url)
    { regex: /\[([^\]]+)\]\(([^)]+)\)/g, style: { }, groupIndex: 1, isLink: true },
    // 斜体 *text* 或 _text_ (注意：需要在加粗之后处理，避免误匹配)
    { regex: /\*([^*]+)\*|_([^_]+)_/g, style: { italics: true }, groupIndex: 1 },
  ]

  // 按位置排序所有匹配结果
  const allMatches = []

  for (const pattern of patterns) {
    pattern.regex.lastIndex = 0
    let m
    while ((m = pattern.regex.exec(remaining)) !== null) {
      // 获取内容 - 根据不同的正则使用不同的捕获组
      let content = ''
      let url = ''
      if (pattern.isLink) {
        content = m[1] || ''
        url = m[2] || ''
      } else {
        // 尝试多个可能的捕获组
        content = m[pattern.groupIndex] || m[1] || ''
      }

      allMatches.push({
        start: m.index,
        end: m.index + m[0].length,
        fullMatch: m[0],
        content: content,
        url: url,
        style: pattern.style,
        isLink: pattern.isLink || false
      })
    }
  }

  // 按起始位置排序，去除重叠的匹配
  allMatches.sort((a, b) => a.start - b.start)

  // 去重：只保留最早的非重叠匹配
  const validMatches = []
  let lastEnd = 0
  for (const m of allMatches) {
    if (m.start >= lastEnd) {
      validMatches.push(m)
      lastEnd = m.end
    }
  }

  // 根据匹配结果分割文本
  let pos = 0
  for (const m of validMatches) {
    // 添加匹配前的普通文本（处理换行符）
    if (m.start > pos) {
      const plainText = remaining.substring(pos, m.start)
      // 处理换行符 - 将文本按\n分割，每个片段后添加break
      addTextWithLineBreaks(plainText, runs)
    }

    // 添加带样式的文本（不处理换行符，样式文本通常不含换行）
    if (m.isLink && m.url) {
      // 链接：使用 ExternalHyperlink
      runs.push(new ExternalHyperlink({
        children: [new TextRun({ text: m.content, style: 'Hyperlink' })],
        link: m.url
      }))
    } else {
      runs.push(new TextRun({
        text: m.content,
        ...m.style
      }))
    }
    pos = m.end
  }

  // 添加剩余的普通文本（处理换行符）
  if (pos < remaining.length) {
    const plainText = remaining.substring(pos)
    addTextWithLineBreaks(plainText, runs)
  }

  // 如果没有匹配任何样式，返回处理换行后的原始文本
  if (runs.length === 0 && text) {
    addTextWithLineBreaks(text, runs)
  }

  return runs
}

/**
 * 处理文本中的换行符，将\n转换为Word的break
 * @param {string} text 要处理的文本
 * @param {Array} runs TextRun数组，会添加新的TextRun元素
 */
function addTextWithLineBreaks(text, runs) {
  if (!text) return

  // 清理残留的Markdown标记符号
  const cleanedText = cleanMarkdownSymbols(text)

  // 检查是否有换行符
  if (!cleanedText.includes('\n')) {
    // 无换行，直接添加文本
    runs.push(new TextRun({ text: cleanedText }))
    return
  }

  // 有换行符，按换行符分割
  const parts = cleanedText.split('\n')

  for (let i = 0; i < parts.length; i++) {
    const part = parts[i]
    if (part) {
      runs.push(new TextRun({ text: part }))
    }
    // 在每个换行后添加break（除了最后一个）
    if (i < parts.length - 1) {
      // 使用带break的空TextRun来创建换行
      runs.push(new TextRun({ text: '', break: 1 }))
    }
  }
}

/**
 * 清理残留的Markdown标记符号
 * 用于处理未完全匹配的标记或边缘情况
 */
function cleanMarkdownSymbols(text) {
  if (!text) return ''

  let result = text

  // 清理残留的加粗标记
  result = result.replace(/\*\*/g, '')
  result = result.replace(/__/g, '')

  // 清理残留的斜体标记（单个*或_，但不影响普通用法）
  // 注意：这里要小心，不要误删正常的星号或下划线
  // 只删除明显是Markdown标记的情况（前后没有其他字符的孤立标记）
  result = result.replace(/\*{1}(?![*])/g, '') // 单个孤立星号
  result = result.replace(/_{1}(?![_])/g, '') // 单个孤立下划线

  // 清理残留的删除线标记
  result = result.replace(/~~/g, '')

  // 清理残留的代码标记
  result = result.replace(/`/g, '')

  // 清理残留的链接标记
  result = result.replace(/\[([^\]]*)\]/g, '$1') // 保留方括号内的内容
  result = result.replace(/\(([^)]*)\)/g, '$1') // 保留圆括号内的内容（如果不是URL）

  return result
}

/**
 * 创建标题段落（处理公式）
 */
function createHeading(token, blockFormulas, inlineFormulas) {
  const headingLevelMap = {
    1: HeadingLevel.HEADING_1,
    2: HeadingLevel.HEADING_2,
    3: HeadingLevel.HEADING_3,
    4: HeadingLevel.HEADING_4,
    5: HeadingLevel.HEADING_5,
    6: HeadingLevel.HEADING_6
  }

  // 使用 token.text，标题通常不需要内联样式
  // 但仍需处理公式占位符
  const text = token.text || ''
  // 不调用 restoreFormulasInText，直接处理占位符
  // 这样可以正确创建Word原生公式
  const runs = createTextRunsWithNativeFormulas(text, inlineFormulas)

  return new Paragraph({
    heading: headingLevelMap[token.depth] || HeadingLevel.HEADING_1,
    children: runs
  })
}

/**
 * 创建普通段落（从token处理）
 * 使用 token.raw 保留原始 markdown 格式（包括加粗、斜体等）
 * 支持Word原生公式，并正确处理段落换行
 */
function createParagraphFromToken(token, blockFormulas, inlineFormulas) {
  const paragraphs = []

  // 使用 token.raw 获取原始 markdown 文本（包含样式标记）
  let rawText = token.raw || ''
  rawText = rawText.trim()

  // 检查是否包含块级公式占位符
  const blockPattern = /%%BLOCK_FORMULA_(\d+)%%/g
  let match
  let lastIndex = 0
  let currentRuns = []

  while ((match = blockPattern.exec(rawText)) !== null) {
    // 添加公式前的文本
    if (match.index > lastIndex) {
      const beforeText = rawText.substring(lastIndex, match.index)
      const beforeRuns = createTextRunsWithNativeFormulas(beforeText, inlineFormulas)
      beforeRuns.forEach(r => currentRuns.push(r))
    }

    // 如果有累积的文本，先创建段落
    if (currentRuns.length > 0) {
      paragraphs.push(new Paragraph({
        children: currentRuns,
        alignment: AlignmentType.JUSTIFIED
      }))
      currentRuns = []
    }

    // 创建独立的块级公式段落
    const formulaIndex = parseInt(match[1])
    if (formulaIndex < blockFormulas.length) {
      const formula = blockFormulas[formulaIndex].formula
      const mathParagraph = createBlockFormulaParagraph(formula)
      paragraphs.push(mathParagraph)
    }

    lastIndex = match.index + match[0].length
  }

  // 处理剩余文本
  if (lastIndex < rawText.length) {
    const remainingText = rawText.substring(lastIndex)
    const remainingRuns = createTextRunsWithNativeFormulas(remainingText, inlineFormulas)
    remainingRuns.forEach(r => currentRuns.push(r))
  }

  // 创建最后的段落
  if (currentRuns.length > 0) {
    paragraphs.push(new Paragraph({
      children: currentRuns,
      alignment: AlignmentType.JUSTIFIED
    }))
  }

  // 如果没有处理出内容，使用原始方法
  if (paragraphs.length === 0) {
    const runs = createTextRunsWithNativeFormulas(rawText, inlineFormulas)
    if (runs.length > 0) {
      paragraphs.push(new Paragraph({
        children: runs,
        alignment: AlignmentType.JUSTIFIED
      }))
    }
  }

  return paragraphs
}

/**
 * 创建块级公式段落（使用Word原生公式）
 * 公式居中显示，使用专业数学字体
 */
function createBlockFormulaParagraph(latexFormula) {
  // 尝试创建Word原生公式
  const mathRun = createWordNativeFormula(latexFormula)

  if (mathRun) {
    // 使用Word原生Math元素
    return new Paragraph({
      children: [mathRun],
      alignment: AlignmentType.CENTER,
      spacing: { before: 240, after: 240 }
    })
  } else {
    // 使用改进的Unicode转换，配合专业数学字体
    const displayText = formatLatexForDisplayImproved(latexFormula)
    return new Paragraph({
      children: [
        new TextRun({
          text: displayText,
          font: 'Cambria Math',
          size: 28, // 略大的字体使公式更清晰
          italics: true // 数学公式通常使用斜体
        })
      ],
      alignment: AlignmentType.CENTER,
      spacing: { before: 240, after: 240 }
    })
  }
}

/**
 * 格式化LaTeX公式用于显示（改进版）
 * 使用更美观的Unicode格式化，接近Word公式编辑器效果
 */
function formatLatexForDisplayImproved(latex) {
  if (!latex) return ''

  let result = String(latex)

  // 使用改进的符号转换表
  result = convertLatexSymbolsImproved(result)

  // 处理分数 - 使用更美观的格式
  result = result.replace(/\\frac\s*\{([^}]*)\}\s*\{([^}]*)\}/g, '$1/$2')

  // 处理根号
  result = result.replace(/\\sqrt\s*\[([^\]]*)\]\s*\{([^}]*)\}/g, '$1√($2)')
  result = result.replace(/\\sqrt\s*\{([^}]*)\}/g, '√($1)')

  // 处理上下标 - 使用Unicode上下标字符
  result = convertSuperscripts(result)
  result = convertSubscripts(result)

  // 处理求和、积分
  result = result.replace(/\\sum\s*_\s*\{([^}]*)\}\s*\^\s*\{([^}]*)\}/g, '∑($1→$2)')
  result = result.replace(/\\sum/g, '∑')
  result = result.replace(/\\int\s*_\s*\{([^}]*)\}\s*\^\s*\{([^}]*)\}/g, '∫($1→$2)')
  result = result.replace(/\\int/g, '∫')

  // 清理剩余的LaTeX命令
  result = result.replace(/\\[a-zA-Z]+/g, '')
  result = result.replace(/\\left|\\right/g, '')
  result = result.replace(/\{([^}]*)\}/g, '$1')

  // 清理多余空格
  result = result.replace(/\s+/g, ' ').trim()

  return result
}

/**
 * 改进的符号转换
 */
function convertLatexSymbolsImproved(text) {
  const symbols = {
    '\\alpha': 'α', '\\beta': 'β', '\\gamma': 'γ', '\\delta': 'δ',
    '\\epsilon': 'ε', '\\zeta': 'ζ', '\\eta': 'η', '\\theta': 'θ',
    '\\iota': 'ι', '\\kappa': 'κ', '\\lambda': 'λ', '\\mu': 'μ',
    '\\nu': 'ν', '\\xi': 'ξ', '\\pi': 'π', '\\rho': 'ρ',
    '\\sigma': 'σ', '\\tau': 'τ', '\\upsilon': 'υ', '\\phi': 'φ',
    '\\chi': 'χ', '\\psi': 'ψ', '\\omega': 'ω',
    '\\Gamma': 'Γ', '\\Delta': 'Δ', 'Theta': 'Θ', '\\Lambda': 'Λ',
    '\\Xi': 'Ξ', '\\Pi': 'Π', '\\Sigma': 'Σ', '\\Phi': 'Φ',
    '\\Psi': 'Ψ', '\\Omega': 'Ω',
    '\\infty': '∞', '\\partial': '∂', '\\nabla': '∇',
    '\\pm': '±', '\\mp': '∓', '\\times': '×', '\\div': '÷',
    '\\cdot': '·', '\\leq': '≤', '\\geq': '≥', '\\neq': '≠',
    '\\approx': '≈', '\\equiv': '≡', '\\sim': '∼',
    '\\in': '∈', '\\notin': '∉', '\\subset': '⊂', '\\supset': '⊃',
    '\\cup': '∪', '\\cap': '∩', '\\emptyset': '∅',
    '\\rightarrow': '→', '\\leftarrow': '←', '\\to': '→',
  }

  for (const [cmd, symbol] of Object.entries(symbols)) {
    text = text.replace(new RegExp(cmd.replace(/\\/g, '\\\\\\\\'), 'g'), symbol)
  }

  return text
}

/**
 * 转换上标为Unicode字符
 */
function convertSuperscripts(text) {
  const superscripts = {
    '0': '⁰', '1': '¹', '2': '²', '3': '³', '4': '⁴',
    '5': '⁵', '6': '⁶', '7': '⁷', '8': '⁸', '9': '⁹',
    '+': '⁺', '-': '⁻', '=': '⁼', 'n': 'ⁿ', 'i': 'ⁱ',
    'a': 'ᵃ', 'b': 'ᵇ', 'c': 'ᶜ', 'd': 'ᵈ', 'e': 'ᵉ',
    'f': 'ᶠ', 'g': 'ᵍ', 'h': 'ʰ', 'j': 'ʲ', 'k': 'ᵏ',
    'l': 'ˡ', 'm': 'ᵐ', 'o': 'ᵒ', 'p': 'ᵖ', 'r': 'ʳ',
    's': 'ˢ', 't': 'ᵗ', 'u': 'ᵘ', 'v': 'ᵛ', 'w': 'ᵡ',
    'x': 'ˣ', 'y': 'ʸ', 'z': 'ᶻ',
  }

  // 处理单字符上标 x^2
  text = text.replace(/\^([a-zA-Z0-9+\-=])/g, (m, c) => {
    return superscripts[c] || '^' + c
  })

  // 处理多字符上标 x^{abc} -> x^(abc)
  text = text.replace(/\^\{([^}]*)\}/g, '^($1)')

  return text
}

/**
 * 转换下标为Unicode字符
 */
function convertSubscripts(text) {
  const subscripts = {
    '0': '₀', '1': '₁', '2': '₂', '3': '₃', '4': '₄',
    '5': '₅', '6': '₆', '7': '₇', '8': '₈', '9': '₉',
    '+': '₊', '-': '₋', '=': '₌',
    'a': 'ₐ', 'e': 'ₑ', 'h': 'ₕ', 'i': 'ᵢ', 'j': 'ᵣ',
    'k': 'ₖ', 'l': 'ₗ', 'm': 'ₘ', 'n': 'ₙ', 'o': 'ₒ',
    'p': 'ₚ', 'r': 'ₚ', 's': 'ₛ', 't': 'ₜ', 'u': 'ₘ',
    'v': 'ₓ', 'x': 'ₓ',
  }

  // 处理单字符下标 x_i
  text = text.replace(/_([a-zA-Z0-9+\-=])/g, (m, c) => {
    return subscripts[c] || '_' + c
  })

  // 处理多字符下标 x_{abc} -> x_(abc)
  text = text.replace(/_\{([^}]*)\}/g, '_($1)')

  return text
}

/**
 * 格式化LaTeX公式用于显示
 * 改进版本，更准确地转换LaTeX符号
 */
function formatLatexForDisplay(latex) {
  // 确保 latex 是字符串
  if (latex === null || latex === undefined) {
    return ''
  }
  let result = String(latex)
  const symbolReplacements = {
    // 希腊字母
    '\\alpha': 'α',
    '\\beta': 'β',
    '\\gamma': 'γ',
    '\\delta': 'δ',
    '\\epsilon': 'ε',
    '\\zeta': 'ζ',
    '\\eta': 'η',
    '\\theta': 'θ',
    '\\iota': 'ι',
    '\\kappa': 'κ',
    '\\lambda': 'λ',
    '\\mu': 'μ',
    '\\nu': 'ν',
    '\\xi': 'ξ',
    '\\pi': 'π',
    '\\rho': 'ρ',
    '\\sigma': 'σ',
    '\\tau': 'τ',
    '\\upsilon': 'υ',
    '\\phi': 'φ',
    '\\chi': 'χ',
    '\\psi': 'ψ',
    '\\omega': 'ω',
    '\\varepsilon': 'ɛ',
    '\\varphi': 'ɸ',
    '\\vartheta': 'ϑ',
    '\\varpi': 'ϖ',
    '\\varrho': 'ϱ',
    '\\varsigma': 'ς',
    '\\Gamma': 'Γ',
    '\\Delta': 'Δ',
    '\\Theta': 'Θ',
    '\\Lambda': 'Λ',
    '\\Xi': 'Ξ',
    '\\Pi': 'Π',
    '\\Sigma': 'Σ',
    '\\Phi': 'Φ',
    '\\Psi': 'Ψ',
    '\\Omega': 'Ω',
    // 大小比较符号
    '\\lt': '<',
    '\\gt': '>',
    '\\le': '≤',
    '\\ge': '≥',
    '\\leq': '≤',
    '\\geq': '≥',
    '\\leqq': '≦',
    '\\geqq': '≧',
    '\\lneq': '≨',
    '\\gneq': '≨',
    '\\nless': '≮',
    '\\ngtr': '≯',
    '\\nleq': '≰',
    '\\ngeq': '≱',
    '\\lesssim': '≲',
    '\\gtrsim': '≳',
    '\\lessapprox': '⪅',
    '\\gtrapprox': '⪆',
    '\\ll': '≪',
    '\\gg': '≫',
    '\\lessgtr': '≶',
    '\\gtrless': '≷',
    '\\prec': '≺',
    '\\succ': '≻',
    '\\preceq': '≼',
    '\\succeq': '≽',
    // 其他数学符号
    '\\infty': '∞',
    '\\partial': '∂',
    '\\sum': '∑',
    '\\int': '∫',
    '\\prod': '∏',
    '\\coprod': '∐',
    '\\sqrt': '√',
    '\\pm': '±',
    '\\mp': '∓',
    '\\times': '×',
    '\\div': '÷',
    '\\cdot': '·',
    '\\neq': '≠',
    '\\approx': '≈',
    '\\sim': '∼',
    '\\equiv': '≡',
    '\\propto': '∝',
    '\\parallel': '∥',
    '\\perp': '⊥',
    '\\in': '∈',
    '\\notin': '∉',
    '\\subset': '⊂',
    '\\supset': '⊃',
    '\\subseteq': '⊆',
    '\\supseteq': '⊇',
    '\\cup': '∪',
    '\\cap': '∩',
    '\\emptyset': '∅',
    '\\forall': '∀',
    '\\exists': '∃',
    '\\neg': '¬',
    '\\land': '∧',
    '\\lor': '∨',
    '\\rightarrow': '→',
    '\\leftarrow': '←',
    '\\Rightarrow': '⇒',
    '\\Leftarrow': '⇐',
    '\\leftrightarrow': '↔',
    '\\Leftrightarrow': '⇔',
    '\\hbar': 'ℏ',
    '\\nabla': '∇',
    '\\ell': 'ℓ',
    '\\Re': 'ℜ',
    '\\Im': 'ℑ',
    '\\angle': '∠',
    '\\triangle': '△',
    '\\square': '□',
    '\\circ': '○',
    '\\bullet': '•',
    '\\star': '★',
    '\\dag': '†',
    '\\ddag': '‡',
    '\\S': '§',
    '\\P': '¶',
    '\\copyright': '©',
    '\\registered': '®',
    '\\trademark': '™',
  }

  // 替换符号
  for (const [cmd, symbol] of Object.entries(symbolReplacements)) {
    result = result.replace(new RegExp(cmd.replace(/\\/g, '\\\\\\\\'), 'g'), symbol)
  }

  // 处理分数 \frac{a}{b} -> (a)/(b)
  result = result.replace(/\\frac\s*\{([^}]*)\}\s*\{([^}]*)\}/g, '($1)/($2)')
  // 处理嵌套分数
  result = result.replace(/\\frac\s*\{([^}]*)\}\s*\{([^}]*)\}/g, '($1)/($2)')

  // 处理根号 \sqrt{x} -> √(x), \sqrt[n]{x} -> n√(x)
  result = result.replace(/\\sqrt\s*\[([^\]]*)\]\s*\{([^}]*)\}/g, '$1√($2)')
  result = result.replace(/\\sqrt\s*\{([^}]*)\}/g, '√($1)')
  result = result.replace(/\\sqrt\s+([a-zA-Z0-9])/g, '√$1')

  // 处理上下标 x^2 -> x², x_i -> x_i (下标保持原样因为没有好的Unicode表示)
  const superscripts = {
    '0': '⁰', '1': '¹', '2': '²', '3': '³', '4': '⁴',
    '5': '⁵', '6': '⁶', '7': '⁷', '8': '⁸', '9': '⁹',
    '+': '⁺', '-': '⁻', '=': '⁼', 'n': 'ⁿ', 'i': 'ⁱ',
    'a': 'ᵃ', 'b': 'ᵇ', 'c': 'ᶜ', 'd': 'ᵈ', 'e': 'ᵉ',
    'f': 'ᶠ', 'g': 'ᵍ', 'h': 'ʰ', 'j': 'ʲ', 'k': 'ᵏ',
    'l': 'ˡ', 'm': 'ᵐ', 'o': 'ᵒ', 'p': 'ᵖ', 'r': 'ʳ',
    's': 'ˢ', 't': 'ᵗ', 'u': 'ᵘ', 'v': 'ᵛ', 'w': 'ᵡ',
  }

  // 单字符上标 x^2
  result = result.replace(/\^([a-zA-Z0-9])/g, (m, c) => {
    return superscripts[c] || '^' + c
  })

  // 多字符上标 x^{abc} -> x^(abc)
  result = result.replace(/\^\{([^}]*)\}/g, '^($1)')

  // 下标保持带括号的形式 x_{i} -> x_(i)
  result = result.replace(/_\{([^}]*)\}/g, '_($1)')
  result = result.replace(/_([a-zA-Z0-9])/g, '_$1')

  // 处理求和上下限 \sum_{i=1}^{n} -> ∑_(i=1)^(n)
  result = result.replace(/\\sum\s*_\s*\{([^}]*)\}\s*\^\s*\{([^}]*)\}/g, '∑_($1)^($2)')
  result = result.replace(/\\sum\s*_\s*([a-zA-Z0-9=]+)\s*\^\s*([a-zA-Z0-9]+)/g, '∑_$1^$2')
  result = result.replace(/\\sum/g, '∑')

  // 处理积分上下限
  result = result.replace(/\\int\s*_\s*\{([^}]*)\}\s*\^\s*\{([^}]*)\}/g, '∫_($1)^($2)')
  result = result.replace(/\\int/g, '∫')

  // 处理极限 \lim_{x\to a} -> lim_(x→a)
  result = result.replace(/\\lim\s*_\s*\{([^}]*)\}/g, 'lim_($1)')
  result = result.replace(/\\lim/g, 'lim')

  // 清理剩余的LaTeX命令和括号
  result = result.replace(/\\left\s*/g, '')
  result = result.replace(/\\right\s*/g, '')
  result = result.replace(/\\quad/g, '    ')
  result = result.replace(/\\qquad/g, '        ')
  result = result.replace(/\\,/g, ' ')
  result = result.replace(/\\;/g, '  ')
  result = result.replace(/\\!/g, '')
  result = result.replace(/\\ /g, ' ')

  // 处理 \text{内容} - 支持嵌套花括号和中文
  // 使用更健壮的匹配方式
  result = result.replace(/\\text\s*\{/g, '')
  result = result.replace(/\\textrm\s*\{/g, '')
  result = result.replace(/\\mathrm\s*\{/g, '')
  result = result.replace(/\\mathbf\s*\{/g, '')
  result = result.replace(/\\mathit\s*\{/g, '')
  result = result.replace(/\\mathbb\s*\{/g, '')
  result = result.replace(/\\mathcal\s*\{/g, '')
  result = result.replace(/\\boldsymbol\s*\{/g, '')
  result = result.replace(/\\bm\s*\{/g, '')

  // 清理剩余的反斜杠命令
  result = result.replace(/\\[a-zA-Z]+/g, '')

  // 清理多余的花括号（保留内容）
  result = result.replace(/\{([^}]*)\}/g, '$1')
  // 处理可能剩余的嵌套花括号
  while (result.includes('{') && result.includes('}')) {
    const newResult = result.replace(/\{([^}]*)\}/g, '$1')
    if (newResult === result) break
    result = newResult
  }

  // 清理多余的空格
  result = result.replace(/\s+/g, ' ').trim()

  return result
}

/**
 * 创建列表段落（支持嵌套列表和列表项中的复杂内容）
 * marked 库的 list token 中，每个 item 可能包含嵌套的子列表
 */
function createNestedList(token, blockFormulas, inlineFormulas) {
  const paragraphs = []

  for (const item of token.items) {
    // 获取当前项的层级深度
    const level = item.depth || 0

    // 使用 item.text 作为列表项的主要内容（marked已经正确提取）
    // 这样可以避免从 raw 中提取时重复处理嵌套内容
    let itemText = item.text || ''

    // 检查是否有嵌套子列表（通过 item.tokens 检查）
    const nestedListTokens = item.tokens ? item.tokens.filter(t => t.type === 'list') : []

    if (nestedListTokens.length > 0) {
      // 有嵌套子列表的情况
      // item.text 已经包含了列表项的主要内容（不包含子列表）
      if (itemText.trim()) {
        const runs = createTextRunsWithNativeFormulas(itemText.trim(), inlineFormulas)
        paragraphs.push(
          new Paragraph({
            children: runs,
            bullet: {
              level: Math.min(level, 8),
              format: token.ordered ? LevelFormat.DECIMAL : LevelFormat.BULLET
            }
          })
        )
      }

      // 处理嵌套子列表（递归调用）
      for (const subToken of nestedListTokens) {
        paragraphs.push(...createNestedList(subToken, blockFormulas, inlineFormulas))
      }
    } else {
      // 简单列表项（无嵌套）
      if (itemText.trim()) {
        const runs = createTextRunsWithNativeFormulas(itemText.trim(), inlineFormulas)
        paragraphs.push(
          new Paragraph({
            children: runs,
            bullet: {
              level: Math.min(level, 8),
              format: token.ordered ? LevelFormat.DECIMAL : LevelFormat.BULLET
            }
          })
        )
      }
    }
  }

  return paragraphs
}

/**
 * 创建分割线
 */
function createHorizontalLine() {
  // Word 中的分割线：使用带底部边框的空段落
  return new Paragraph({
    border: {
      bottom: {
        color: 'auto',
        space: 1,
        style: BorderStyle.SINGLE,
        size: 6
      }
    },
    spacing: {
      after: 200,
      before: 200
    }
  })
}

/**
 * 创建包含Word原生公式的TextRun数组
 * 同时处理Markdown内联样式（加粗、斜体、删除线等）
 */
function createTextRunsWithNativeFormulas(text, inlineFormulas) {
  if (!text) {
    return [new TextRun({ text: '' })]
  }

  const runs = []
  let remaining = String(text)

  // 处理行内公式占位符
  const inlinePattern = /%%INLINE_FORMULA_(\d+)%%/g
  let match

  while ((match = inlinePattern.exec(remaining)) !== null) {
    // 添加公式前的文本（解析内联样式）
    if (match.index > 0) {
      const beforeText = remaining.substring(0, match.index)
      runs.push(...parseInlineStyles(beforeText))
    }

    // 添加公式 - 使用Word原生Math或改进的Unicode转换
    const formulaIndex = parseInt(match[1])
    if (formulaIndex < inlineFormulas.length) {
      const formula = inlineFormulas[formulaIndex].formula
      // 尝试创建Word原生公式
      const mathRun = createWordNativeFormula(formula)
      if (mathRun) {
        runs.push(mathRun)
      } else {
        // 使用改进的Unicode转换，配合专业数学字体和斜体
        runs.push(new TextRun({
          text: formatLatexForDisplayImproved(formula),
          font: 'Cambria Math',
          italics: true  // 数学公式使用斜体更美观
        }))
      }
    }

    remaining = remaining.substring(match.index + match[0].length)
    inlinePattern.lastIndex = 0
  }

  // 添加剩余文本（解析内联样式）
  if (remaining) {
    runs.push(...parseInlineStyles(remaining))
  }

  // 如果没有内容，返回原始文本（解析样式）
  if (runs.length === 0 && text) {
    runs.push(...parseInlineStyles(text))
  }

  return runs
}

/**
 * 创建Word原生公式（可在Word公式编辑器中编辑）
 * 使用docx库的Math模块，将LaTeX转换为Word OMML格式的数学公式
 */
function createWordNativeFormula(latexFormula) {
  try {
    // 解析LaTeX并创建Word Math元素
    const mathElements = parseLatexToMathElements(latexFormula)

    if (mathElements && mathElements.length > 0) {
      // 创建Word OMML公式 - 用户可以在Word中编辑
      return new DocxMath({
        children: mathElements
      })
    }

    // 如果无法解析复杂结构，创建简单公式（包含原始LaTeX）
    // 这样用户至少可以在Word中看到并编辑公式
    const simpleElements = createSimpleMathElements(latexFormula)
    if (simpleElements && simpleElements.length > 0) {
      return new DocxMath({
        children: simpleElements
      })
    }

    return null
  } catch (e) {
    console.warn('createWordNativeFormula error:', e.message)
    return null
  }
}

/**
 * 创建简单数学元素（将LaTeX转换为Unicode符号）
 */
function createSimpleMathElements(latex) {
  if (!latex) return null

  const elements = []
  let text = latex

  // 转换常见的LaTeX符号为Unicode
  text = convertLatexToUnicodeSymbols(text)

  // 创建MathRun
  elements.push(new MathRun(text))

  return elements
}

/**
 * 将LaTeX转换为Unicode数学符号（用于简单公式）
 */
function convertLatexToUnicodeSymbols(latex) {
  if (!latex) return ''

  let result = String(latex)

  // 希腊字母
  const greekLetters = {
    '\\alpha': 'α', '\\beta': 'β', '\\gamma': 'γ', '\\delta': 'δ',
    '\\epsilon': 'ε', '\\zeta': 'ζ', '\\eta': 'η', '\\theta': 'θ',
    '\\iota': 'ι', '\\kappa': 'κ', '\\lambda': 'λ', '\\mu': 'μ',
    '\\nu': 'ν', '\\xi': 'ξ', '\\pi': 'π', '\\rho': 'ρ',
    '\\sigma': 'σ', '\\tau': 'τ', '\\upsilon': 'υ', '\\phi': 'φ',
    '\\chi': 'χ', '\\psi': 'ψ', '\\omega': 'ω',
    '\\Gamma': 'Γ', '\\Delta': 'Δ', '\\Theta': 'Θ', '\\Lambda': 'Λ',
    '\\Xi': 'Ξ', '\\Pi': 'Π', '\\Sigma': 'Σ', '\\Phi': 'Φ',
    '\\Psi': 'Ψ', '\\Omega': 'Ω',
  }

  for (const [cmd, symbol] of Object.entries(greekLetters)) {
    result = result.replace(new RegExp(cmd.replace(/\\/g, '\\\\\\\\'), 'g'), symbol)
  }

  // 数学运算符和比较符号
  const operators = {
    // 大小比较符号
    '\\lt': '<', '\\gt': '>',  // 小于、大于
    '\\le': '≤', '\\ge': '≥',  // 小于等于、大于等于
    '\\leq': '≤', '\\geq': '≥',  // 小于等于、大于等于
    '\\leqq': '≦', '\\geqq': '≧',  // 双线小于等于、大于等于
    '\\lneq': '≨', '\\gneq': '≩',  // 小于但不等于、大于但不等于
    '\\nless': '≮', '\\ngtr': '≯',  // 不小于、不大于
    '\\nleq': '≰', '\\ngeq': '≱',  // 不小于等于、不大于等于
    '\\lesssim': '≲', '\\gtrsim': '≳',  // 小于相似、大于相似
    '\\lessapprox': '⪅', '\\gtrapprox': '⪆',  // 小于约等于、大于约等于
    '\\ll': '≪', '\\gg': '≫',  // 远小于、远大于
    '\\lessgtr': '≶', '\\gtrless': '≷',  // 小于大于、大于小于
    '\\prec': '≺', '\\succ': '≻',  // 先于、后于
    '\\preceq': '≼', '\\succeq': '≽',  // 先于等于、后于等于

    // 基本运算符
    '\\pm': '±', '\\mp': '∓', '\\times': '×', '\\div': '÷',
    '\\cdot': '·', '\\neq': '≠', '\\approx': '≈',
    '\\equiv': '≡', '\\sim': '∼', '\\propto': '∝',
    '\\in': '∈', '\\notin': '∉', '\\subset': '⊂', '\\supset': '⊃',
    '\\subseteq': '⊆', '\\supseteq': '⊇',
    '\\cup': '∪', '\\cap': '∩', '\\emptyset': '∅',
    '\\infty': '∞', '\\partial': '∂', '\\nabla': '∇',
    '\\forall': '∀', '\\exists': '∃', '\\neg': '¬',
    '\\land': '∧', '\\lor': '∨',
    '\\rightarrow': '→', '\\leftarrow': '←', '\\to': '→',
    '\\Rightarrow': '⇒', '\\Leftarrow': '⇐',
    '\\sqrt': '√', '\\sum': '∑', '\\int': '∫', '\\prod': '∏',
  }

  for (const [cmd, symbol] of Object.entries(operators)) {
    result = result.replace(new RegExp(cmd.replace(/\\/g, '\\\\\\\\'), 'g'), symbol)
  }

  // 清理花括号和剩余命令
  result = result.replace(/\\left|\\right/g, '')
  result = result.replace(/\\text\s*\{|\\mathrm\s*\{|\\mathbf\s*\{/g, '')
  result = result.replace(/\\[a-zA-Z]+/g, '')
  result = result.replace(/\{([^}]*)\}/g, '$1')

  // 清理空格
  result = result.replace(/\s+/g, ' ').trim()

  return result
}

/**
 * 提取花括号内的内容（支持嵌套）
 * @param {string} str 从花括号开始的位置
 * @returns {Object} { content: 提取的内容, length: 匹配的总长度 }
 */
function extractBracedContent(str) {
  if (!str || str[0] !== '{') {
    return { content: null, length: 0 }
  }

  let depth = 0
  let content = ''
  let i = 0

  while (i < str.length) {
    const char = str[i]
    if (char === '{') {
      depth++
      if (depth > 1) {
        content += char
      }
    } else if (char === '}') {
      depth--
      if (depth === 0) {
        return { content: content, length: i + 1 }
      }
      content += char
    } else if (char === '\\' && i + 1 < str.length) {
      // 处理转义字符
      content += char + str[i + 1]
      i++
    } else {
      if (depth > 0) {
        content += char
      }
    }
    i++
  }

  // 未匹配到完整的花括号
  return { content: null, length: 0 }
}

/**
 * 解析LaTeX并创建Word Math元素数组
 * 支持分数、上下标、根号、文本、显示样式等常见结构
 * 使用正确的docx Math API参数
 */
function parseLatexToMathElements(latex) {
  if (!latex) return null

  const elements = []
  let remaining = latex.trim()

  while (remaining.length > 0) {
    let matched = false

    // 0. 处理显示样式 \displaystyle（跳过，Word自动处理）
    if (remaining.match(/^\\displaystyle\b/) && !matched) {
      remaining = remaining.replace(/^\\displaystyle\s*/, '')
      matched = true
      continue
    }

    // 0.1 处理文本样式 \textstyle（跳过）
    if (remaining.match(/^\\textstyle\b/) && !matched) {
      remaining = remaining.replace(/^\\textstyle\s*/, '')
      matched = true
      continue
    }

    // 0.2 处理 \limits 和 \nolimits（跳过，Word自动处理）
    if (remaining.match(/^\\(limits|nolimits)\b/) && !matched) {
      remaining = remaining.replace(/^\\(limits|nolimits)\s*/, '')
      matched = true
      continue
    }

    // 1. 处理分数 \frac{a}{b}、\dfrac{a}{b}、\tfrac{a}{b}（支持嵌套花括号）
    const fracCmdMatch = remaining.match(/^\\(?:dfrac|frac|tfrac)\s*/)
    if (fracCmdMatch && !matched) {
      let pos = fracCmdMatch[0].length

      // 提取分子（第一个花括号）
      const numeratorPart = extractBracedContent(remaining.substring(pos))
      if (numeratorPart.content !== null) {
        pos += numeratorPart.length

        // 提取分母（第二个花括号）
        const denominatorPart = extractBracedContent(remaining.substring(pos))
        if (denominatorPart.content !== null) {
          pos += denominatorPart.length

          // 递归解析分子和分母
          const numeratorChildren = parseLatexToMathElements(numeratorPart.content) ||
            [new MathRun(convertLatexToUnicodeSymbols(numeratorPart.content))]
          const denominatorChildren = parseLatexToMathElements(denominatorPart.content) ||
            [new MathRun(convertLatexToUnicodeSymbols(denominatorPart.content))]

          elements.push(new MathFraction({
            numerator: numeratorChildren,
            denominator: denominatorChildren
          }))
          remaining = remaining.substring(pos)
          matched = true
        }
      }
    }

    // 1.1 处理简写分数 \atop（如 a\atop b）
    const atopMatch = remaining.match(/^([a-zA-Z0-9]+)\\atop([a-zA-Z0-9]+)/)
    if (atopMatch && !matched) {
      elements.push(new MathFraction({
        numerator: [new MathRun(atopMatch[1])],
        denominator: [new MathRun(atopMatch[2])]
      }))
      remaining = remaining.substring(atopMatch[0].length)
      matched = true
    }

    // 1.2 处理 \over 分数（如 {a \over b}）
    const overMatch = remaining.match(/^\{([^{}]*?)\\over\s*([^{}]*?)\}/)
    if (overMatch && !matched) {
      const numeratorChildren = parseLatexToMathElements(overMatch[1]) || [new MathRun(overMatch[1])]
      const denominatorChildren = parseLatexToMathElements(overMatch[2]) || [new MathRun(overMatch[2])]
      elements.push(new MathFraction({
        numerator: numeratorChildren,
        denominator: denominatorChildren
      }))
      remaining = remaining.substring(overMatch[0].length)
      matched = true
    }

    // 1.3 处理二项式系数 \binom{n}{k}（支持嵌套）
    const binomCmdMatch = remaining.match(/^\\binom\s*/)
    if (binomCmdMatch && !matched) {
      let pos = binomCmdMatch[0].length

      const topPart = extractBracedContent(remaining.substring(pos))
      if (topPart.content !== null) {
        pos += topPart.length

        const bottomPart = extractBracedContent(remaining.substring(pos))
        if (bottomPart.content !== null) {
          pos += bottomPart.length

          const topChildren = parseLatexToMathElements(topPart.content) ||
            [new MathRun(convertLatexToUnicodeSymbols(topPart.content))]
          const bottomChildren = parseLatexToMathElements(bottomPart.content) ||
            [new MathRun(convertLatexToUnicodeSymbols(bottomPart.content))]

          // 使用括号包裹分数来模拟二项式系数
          elements.push(new MathRoundBrackets({
            children: [new MathFraction({
              numerator: topChildren,
              denominator: bottomChildren
            })]
          }))
          remaining = remaining.substring(pos)
          matched = true
        }
      }
    }

    // 2. 处理根号 \sqrt{x} 或 \sqrt[n]{x}
    const sqrtMatch = remaining.match(/^\\sqrt(?:\s*\[([^\]]*)\])?\s*\{([^}]*)\}/)
    if (sqrtMatch && !matched) {
      const baseChildren = parseLatexToMathElements(sqrtMatch[2]) || [new MathRun(convertLatexToUnicodeSymbols(sqrtMatch[2]))]
      const radOptions = { children: baseChildren }
      if (sqrtMatch[1]) {
        radOptions.degree = [new MathRun(convertLatexToUnicodeSymbols(sqrtMatch[1]))]
      }
      elements.push(new MathRadical(radOptions))
      remaining = remaining.substring(sqrtMatch[0].length)
      matched = true
    }

    // 3. 处理文本 \text{内容}、\textrm{内容}、\mathrm{内容} 等
    const textMatch = remaining.match(/^\\(?:text|textrm|mathrm|mathbf|mathit|mathbb|mathcal|boldsymbol|bm)\s*\{([^}]*)\}/)
    if (textMatch && !matched) {
      // 文本内容直接显示，不进行数学格式化
      elements.push(new MathRun(textMatch[1]))
      remaining = remaining.substring(textMatch[0].length)
      matched = true
    }

    // 4. 处理同时有上下标 x_{i}^{n}
    const subSupMatch = remaining.match(/^([a-zA-Z0-9αβγδεζηθικλμνξπρστυφχψωΓΔΘΛΞΠΣΦΨΩ])\s*_\s*\{([^}]*)\}\s*\^\s*\{([^}]*)\}/)
    if (subSupMatch && !matched) {
      elements.push(new MathSubSuperScript({
        children: [new MathRun(subSupMatch[1])],
        subScript: [new MathRun(convertLatexToUnicodeSymbols(subSupMatch[2]))],
        superScript: [new MathRun(convertLatexToUnicodeSymbols(subSupMatch[3]))]
      }))
      remaining = remaining.substring(subSupMatch[0].length)
      matched = true
    }

    // 5. 处理上标 x^{n} 或 x^n
    const supMatch = remaining.match(/^([a-zA-Z0-9αβγδεζηθικλμνξπρστυφχψωΓΔΘΛΞΠΣΦΨΩ])\s*\^\s*(?:\{([^}]*)\}|([a-zA-Z0-9]))/)
    if (supMatch && !matched) {
      const supContent = convertLatexToUnicodeSymbols(supMatch[2] || supMatch[3] || '')
      elements.push(new MathSuperScript({
        children: [new MathRun(supMatch[1])],
        superScript: [new MathRun(supContent)]
      }))
      remaining = remaining.substring(supMatch[0].length)
      matched = true
    }

    // 6. 处理下标 x_{i} 或 x_i
    const subMatch = remaining.match(/^([a-zA-Z0-9αβγδεζηθικλμνξπρστυφχψωΓΔΘΛΞΠΣΦΨΩ])\s*_\s*(?:\{([^}]*)\}|([a-zA-Z0-9]))/)
    if (subMatch && !matched) {
      const subContent = convertLatexToUnicodeSymbols(subMatch[2] || subMatch[3] || '')
      elements.push(new MathSubScript({
        children: [new MathRun(subMatch[1])],
        subScript: [new MathRun(subContent)]
      }))
      remaining = remaining.substring(subMatch[0].length)
      matched = true
    }

    // 7. 处理求和 \sum
    const sumMatch = remaining.match(/^\\sum(?:\s*_\s*(?:\{([^}]*)\}|([a-zA-Z0-9=]+)))?(?:\s*\^\s*(?:\{([^}]*)\}|([a-zA-Z0-9]+)))?/)
    if (sumMatch && !matched) {
      const subContent = convertLatexToUnicodeSymbols(sumMatch[2] || sumMatch[1] || '')
      const supContent = convertLatexToUnicodeSymbols(sumMatch[4] || sumMatch[3] || '')
      elements.push(new MathSum({
        children: [new MathRun('∑')],
        subScript: subContent ? [new MathRun(subContent)] : undefined,
        superScript: supContent ? [new MathRun(supContent)] : undefined
      }))
      remaining = remaining.substring(sumMatch[0].length)
      matched = true
    }

    // 8. 处理积分 \int
    const intMatch = remaining.match(/^\\int(?:\s*_\s*(?:\{([^}]*)\}|([a-zA-Z0-9=]+)))?(?:\s*\^\s*(?:\{([^}]*)\}|([a-zA-Z0-9]+)))?/)
    if (intMatch && !matched) {
      const subContent = convertLatexToUnicodeSymbols(intMatch[2] || intMatch[1] || '')
      const supContent = convertLatexToUnicodeSymbols(intMatch[4] || intMatch[3] || '')
      elements.push(new MathIntegral({
        children: [new MathRun('∫')],
        subScript: subContent ? [new MathRun(subContent)] : undefined,
        superScript: supContent ? [new MathRun(supContent)] : undefined
      }))
      remaining = remaining.substring(intMatch[0].length)
      matched = true
    }

    // 9. 处理极限 \lim
    const limMatch = remaining.match(/^\\lim(?:\s*_\s*\{([^}]*)\})?/)
    if (limMatch && !matched) {
      elements.push(new MathRun('lim'))
      if (limMatch[1]) {
        elements.push(new MathSubScript({
          children: [new MathRun(' ')],
          subScript: [new MathRun(convertLatexToUnicodeSymbols(limMatch[1]))]
        }))
      }
      remaining = remaining.substring(limMatch[0].length)
      matched = true
    }

    // 10. 处理乘积 \prod
    const prodMatch = remaining.match(/^\\prod(?:\s*_\s*(?:\{([^}]*)\}|([a-zA-Z0-9=]+)))?(?:\s*\^\s*(?:\{([^}]*)\}|([a-zA-Z0-9]+)))?/)
    if (prodMatch && !matched) {
      const subContent = convertLatexToUnicodeSymbols(prodMatch[2] || prodMatch[1] || '')
      const supContent = convertLatexToUnicodeSymbols(prodMatch[4] || prodMatch[3] || '')
      elements.push(new MathSum({
        children: [new MathRun('∏')],
        subScript: subContent ? [new MathRun(subContent)] : undefined,
        superScript: supContent ? [new MathRun(supContent)] : undefined
      }))
      remaining = remaining.substring(prodMatch[0].length)
      matched = true
    }

    // 11. 处理括号 (内容)
    const parenMatch = remaining.match(/^\(([^)]+)\)/)
    if (parenMatch && !matched) {
      const innerElements = parseLatexToMathElements(parenMatch[1]) || [new MathRun(parenMatch[1])]
      elements.push(new MathRoundBrackets({
        children: innerElements
      }))
      remaining = remaining.substring(parenMatch[0].length)
      matched = true
    }

    // 12. 处理方括号 [内容]
    const squareBracketMatch = remaining.match(/^\[([^\]]+)\]/)
    if (squareBracketMatch && !matched) {
      const innerElements = parseLatexToMathElements(squareBracketMatch[1]) || [new MathRun(squareBracketMatch[1])]
      elements.push(new MathSquareBrackets({
        children: innerElements
      }))
      remaining = remaining.substring(squareBracketMatch[0].length)
      matched = true
    }

    // 13. 处理花括号 \{内容\}
    const curlyBracketMatch = remaining.match(/^\\\{([^}]+)\\\}/)
    if (curlyBracketMatch && !matched) {
      const innerElements = parseLatexToMathElements(curlyBracketMatch[1]) || [new MathRun(curlyBracketMatch[1])]
      elements.push(new MathCurlyBrackets({
        children: innerElements
      }))
      remaining = remaining.substring(curlyBracketMatch[0].length)
      matched = true
    }

    // 14. 处理尖括号 \langle 和 \rangle
    const angleBracketMatch = remaining.match(/^\\langle\s*(.+?)\s*\\rangle/)
    if (angleBracketMatch && !matched) {
      const innerElements = parseLatexToMathElements(angleBracketMatch[1]) || [new MathRun(angleBracketMatch[1])]
      elements.push(new MathAngledBrackets({
        children: innerElements
      }))
      remaining = remaining.substring(angleBracketMatch[0].length)
      matched = true
    }

    // 15. 处理LaTeX符号
    const symbolMatch = remaining.match(/^\\([a-zA-Z]+)/)
    if (symbolMatch && !matched) {
      const symbol = convertLatexSymbol(symbolMatch[1])
      if (symbol !== null) {
        if (symbol !== '') {
          elements.push(new MathRun(symbol))
        }
        remaining = remaining.substring(symbolMatch[0].length)
        matched = true
      }
    }

    // 16. 处理自适应括号 \left( ... \right)
    const leftRightMatch = remaining.match(/^\\left\s*([(\[{|])\s*(.+?)\s*\\right\s*([)\]}|])/)
    if (leftRightMatch && !matched) {
      const innerElements = parseLatexToMathElements(leftRightMatch[2]) || [new MathRun(leftRightMatch[2])]
      const leftChar = leftRightMatch[1]
      const rightChar = leftRightMatch[3]

      if (leftChar === '(' && rightChar === ')') {
        elements.push(new MathRoundBrackets({ children: innerElements }))
      } else if (leftChar === '[' && rightChar === ']') {
        elements.push(new MathSquareBrackets({ children: innerElements }))
      } else if (leftChar === '{' && rightChar === '}') {
        elements.push(new MathCurlyBrackets({ children: innerElements }))
      } else if (leftChar === '|' && rightChar === '|') {
        // 绝对值符号
        elements.push(new MathRun('|'))
        elements.push(...innerElements)
        elements.push(new MathRun('|'))
      } else {
        elements.push(new MathRun(leftChar))
        elements.push(...innerElements)
        elements.push(new MathRun(rightChar))
      }
      remaining = remaining.substring(leftRightMatch[0].length)
      matched = true
    }

    // 17. 处理花括号 {内容}
    const braceMatch = remaining.match(/^\{([^}]*)\}/)
    if (braceMatch && !matched) {
      const innerElements = parseLatexToMathElements(braceMatch[1])
      if (innerElements) {
        elements.push(...innerElements)
      } else {
        elements.push(new MathRun(braceMatch[1]))
      }
      remaining = remaining.substring(braceMatch[0].length)
      matched = true
    }

    // 18. 处理空格 \quad, \qquad, \,, \;, \:, \
    const spaceMatch = remaining.match(/^\\(?:quad|qquad|,|;|:| )/)
    if (spaceMatch && !matched) {
      let space = ' '
      if (spaceMatch[0] === '\\quad') space = '    '
      if (spaceMatch[0] === '\\qquad') space = '        '
      if (spaceMatch[0] === '\\,') space = ' '
      if (spaceMatch[0] === '\\;') space = '  '
      elements.push(new MathRun(space))
      remaining = remaining.substring(spaceMatch[0].length)
      matched = true
    }

    // 19. 处理普通字符
    if (!matched) {
      const char = remaining[0]
      // 跳过特殊字符
      if (char === '\\' || char === '$') {
        remaining = remaining.substring(1)
        continue
      }
      elements.push(new MathRun(char))
      remaining = remaining.substring(1)
    }
  }

  return elements.length > 0 ? elements : null
}

/**
 * LaTeX 符号到 Unicode 的转换表
 */
const LATEX_SYMBOLS = {
  // 希腊字母
  'alpha': 'α', 'beta': 'β', 'gamma': 'γ', 'delta': 'δ',
  'epsilon': 'ε', 'zeta': 'ζ', 'eta': 'η', 'theta': 'θ',
  'iota': 'ι', 'kappa': 'κ', 'lambda': 'λ', 'mu': 'μ',
  'nu': 'ν', 'xi': 'ξ', 'pi': 'π', 'rho': 'ρ',
  'sigma': 'σ', 'tau': 'τ', 'upsilon': 'υ', 'phi': 'φ',
  'chi': 'χ', 'psi': 'ψ', 'omega': 'ω',
  'Gamma': 'Γ', 'Delta': 'Δ', 'Theta': 'Θ', 'Lambda': 'Λ',
  'Xi': 'Ξ', 'Pi': 'Π', 'Sigma': 'Σ', 'Phi': 'Φ',
  'Psi': 'Ψ', 'Omega': 'Ω',
  'varepsilon': 'ɛ', 'varphi': 'ɸ', 'vartheta': 'ϑ',

  // 大小比较符号（完整支持）
  'lt': '<', 'gt': '>',  // 小于、大于
  'le': '≤', 'ge': '≥',  // 小于等于、大于等于（别名）
  'leq': '≤', 'geq': '≥',  // 小于等于、大于等于
  'leqq': '≦', 'geqq': '≧',  // 双线小于等于、大于等于
  'lneq': '≨', 'gneq': '≩',  // 小于但不等于、大于但不等于
  'lneqq': '≨', 'gneqq': '≩',  // 小于但不等于、大于但不等于
  'nless': '≮', 'ngtr': '≯',  // 不小于、不大于
  'nleq': '≰', 'ngeq': '≱',  // 不小于等于、不大于等于
  'nleqq': '≰', 'ngeqq': '≱',  // 不小于等于、不大于等于
  'lesssim': '≲', 'gtrsim': '≳',  // 小于相似、大于相似
  'lessapprox': '⪅', 'gtrapprox': '⪆',  // 小于约等于、大于约等于
  'll': '≪', 'gg': '≫',  // 远小于、远大于
  'lll': '⋘', 'ggg': '⋙',  // 三重远小于、远大于
  'lessgtr': '≶', 'gtrless': '≷',  // 小于大于、大于小于
  'lesseqgtr': '⋚', 'gtreqless': '⋛',  // 小于等于大于、大于等于小于
  'prec': '≺', 'succ': '≻',  // 先于、后于
  'preceq': '≼', 'succeq': '≽',  // 先于等于、后于等于
  'nprec': '⊀', 'nsucc': '⊁',  // 不先于、不后于
  'precsim': '≾', 'succsim': '≿',  // 先于相似、后于相似
  'preccurlyeq': '≼', 'succcurlyeq': '≽',  // 先于等于、后于等于

  // 基本运算符
  'pm': '±', 'mp': '∓', 'times': '×', 'div': '÷',
  'cdot': '·', 'ast': '∗', 'star': '★',
  'neq': '≠', 'approx': '≈',
  'equiv': '≡', 'sim': '∼', 'propto': '∝',
  'parallel': '∥', 'perp': '⊥', 'in': '∈', 'notin': '∉',
  'subset': '⊂', 'supset': '⊃', 'subseteq': '⊆', 'supseteq': '⊇',
  'cup': '∪', 'cap': '∩', 'emptyset': '∅',
  'forall': '∀', 'exists': '∃', 'neg': '¬',
  'land': '∧', 'lor': '∨', 'rightarrow': '→', 'leftarrow': '←',
  'Rightarrow': '⇒', 'Leftarrow': '⇐', 'leftrightarrow': '↔',
  'Leftrightarrow': '⇔',

  // 数学运算符
  'sum': '∑', 'int': '∫', 'prod': '∏', 'lim': 'lim',
  'iint': '∬', 'iiint': '∭', 'oint': '∮',

  // 其他符号
  'infty': '∞', 'partial': '∂', 'nabla': '∇',
  'hbar': 'ℏ', 'angle': '∠', 'triangle': '△',
  'square': '□', 'circ': '○', 'bullet': '•',
  'text': '', 'textrm': '', 'mathrm': '', 'mathbf': '',
  'mathit': '', 'mathbb': '', 'mathcal': '', 'boldsymbol': '', 'bm': '',
  'quad': '    ', 'qquad': '        ', ' ': ' ', ',': ' ', ';': '  ',
  'left': '', 'right': '', 'to': '→',
}

/**
 * 转换 LaTeX 符号名称为 Unicode 字符
 */
function convertLatexSymbol(name) {
  return LATEX_SYMBOLS[name] || null
}

/**
 * 创建表格（处理公式和Markdown内联样式）
 * 确保表格内容正确显示，与Markdown预览一致
 */
function createTable(token, blockFormulas, inlineFormulas) {
  const rows = []

  // 表头 - marked 的 table token.header 可能是多种格式
  const headerCells = (token.header || []).map((cell, index) => {
    // 提取单元格文本
    let cellText = ''
    if (cell && typeof cell === 'object') {
      cellText = cell.text || (cell.tokens && cell.tokens.map(t => t.text || t.raw || '').join('')) || cell.raw || String(cell)
    } else if (typeof cell === 'string') {
      cellText = cell
    } else if (cell !== null && cell !== undefined) {
      cellText = String(cell)
    }

    // 对于表头，处理公式占位符，并全部加粗
    const runs = []
    let remaining = String(cellText)

    // 处理行内公式占位符
    const inlinePattern = /%%INLINE_FORMULA_(\d+)%%/g
    let match
    let lastEnd = 0

    while ((match = inlinePattern.exec(remaining)) !== null) {
      // 添加公式前的文本（直接创建加粗 TextRun）
      if (match.index > lastEnd) {
        const beforeText = remaining.substring(lastEnd, match.index)
        runs.push(new TextRun({ text: beforeText, bold: true }))
      }

      // 添加公式
      const formulaIndex = parseInt(match[1])
      if (formulaIndex < inlineFormulas.length) {
        const formula = inlineFormulas[formulaIndex].formula
        const mathRun = createWordNativeFormula(formula)
        if (mathRun) {
          runs.push(mathRun)
        } else {
          runs.push(new TextRun({
            text: formatLatexForDisplay(formula),
            font: 'Cambria Math',
            bold: true
          }))
        }
      }

      lastEnd = match.index + match[0].length
    }

    // 添加剩余文本（直接创建加粗 TextRun）
    if (lastEnd < remaining.length) {
      const restText = remaining.substring(lastEnd)
      runs.push(new TextRun({ text: restText, bold: true }))
    }

    // 如果没有内容，直接使用文本创建加粗 TextRun
    if (runs.length === 0 && cellText) {
      runs.push(new TextRun({ text: cellText, bold: true }))
    }

    return new TableCell({
      children: [new Paragraph({ children: runs })],
      width: { size: 100 / (token.header.length || 1), type: WidthType.PERCENTAGE }
    })
  })

  if (headerCells.length > 0) {
    rows.push(new TableRow({ children: headerCells }))
  }

  // 数据行 - 同样处理多种可能的数据结构
  for (const row of (token.rows || [])) {
    const cells = (row || []).map((cell, cellIndex) => {
      // 提取单元格文本 - 处理多种可能的数据结构
      let cellText = ''

      if (cell && typeof cell === 'object') {
        // 可能是 {text: 'xxx'} 或 {tokens: [...]} 结构
        if (cell.text) {
          cellText = cell.text
        } else if (cell.tokens && Array.isArray(cell.tokens)) {
          // 从 tokens 数组中提取文本
          cellText = cell.tokens.map(t => t.text || t.raw || '').join('')
        } else if (cell.raw) {
          cellText = cell.raw
        } else {
          // 尝试转换为字符串
          cellText = String(cell)
        }
      } else if (typeof cell === 'string') {
        cellText = cell
      } else if (cell !== null && cell !== undefined) {
        cellText = String(cell)
      }

      // 直接处理占位符，创建Word原生公式或Unicode文本
      const runs = createTextRunsWithNativeFormulas(cellText, inlineFormulas)

      // 如果 runs 为空，使用文本
      const finalRuns = runs.length > 0 ? runs : [new TextRun({ text: cellText })]

      return new TableCell({
        children: [new Paragraph({ children: finalRuns })],
        width: { size: 100 / (row.length || 1), type: WidthType.PERCENTAGE }
      })
    })

    if (cells.length > 0) {
      rows.push(new TableRow({ children: cells }))
    }
  }

  return [new Table({
    rows: rows,
    width: { size: 100, type: WidthType.PERCENTAGE }
  })]
}

/**
 * 创建代码块
 */
function createCodeBlock(token) {
  return new Paragraph({
    children: [
      new TextRun({
        text: token.text || '',
        font: 'Consolas',
        size: 20
      })
    ],
    shading: {
      fill: 'F5F5F5'
    }
  })
}

/**
 * 创建引用块（处理公式和内联样式）
 */
function createBlockquote(token, blockFormulas, inlineFormulas) {
  // 使用 token.raw 保留原始 markdown 格式
  let rawText = token.raw || token.text || ''
  rawText = rawText.trim()

  // 去掉引用标记 >
  rawText = rawText.replace(/^>\s*/, '')

  // 不调用 restoreFormulasInText，直接处理占位符
  // 这样可以正确创建Word原生公式
  const runs = createTextRunsWithNativeFormulas(rawText, inlineFormulas)

  // 引用块文本默认斜体，但不覆盖已有的样式
  const italicRuns = runs.map(r => {
    if (r instanceof DocxMath) {
      return r // Math元素不支持斜体
    }
    // 获取现有的样式并添加斜体
    if (r && r.options) {
      return new TextRun({
        ...r.options,
        italics: true  // 引用块默认斜体
      })
    }
    const textContent = (r && r.text) || ''
    return new TextRun({ text: textContent, italics: true })
  })

  return new Paragraph({
    children: italicRuns,
    indent: {
      left: 720 // 0.5 inch
    }
  })
}

/**
 * 生成Word文档
 * @param {string} markdownContent Markdown内容
 * @param {Object} styleConfig 样式配置
 * @returns {Promise<Blob>} Word文档Blob
 */
export async function generateDocx(markdownContent, styleConfig = {}) {
  // 创建文档选项
  const options = createDocumentOptions(styleConfig)

  // 转换Markdown内容为段落
  const paragraphs = convertToDocxParagraphs(markdownContent, styleConfig)

  // 添加段落到文档
  options.sections[0].children = paragraphs

  // 创建文档
  const doc = new Document(options)

  // 打包为Blob
  const blob = await Packer.toBlob(doc)

  return blob
}

/**
 * 生成包含OMML公式的Word文档（高级版本）
 * 使用docx库的Math功能
 * @param {string} markdownContent Markdown内容
 * @param {Object} styleConfig 样式配置
 * @returns {Promise<Blob>} Word文档Blob
 */
export async function generateDocxWithOmml(markdownContent, styleConfig = {}) {
  // 这个函数演示如何使用docx库的Math模块
  // 但由于docx库的Math功能限制，实际公式显示可能有限

  const options = createDocumentOptions(styleConfig)
  const paragraphs = []

  // 页眉
  if (styleConfig.header) {
    paragraphs.push(new Paragraph({
      children: [new TextRun({ text: styleConfig.header })]
    }))
  }

  // 简单示例：创建一个数学公式段落
  // 实际使用时需要更复杂的解析和转换
  paragraphs.push(new Paragraph({
    children: [
      new DocxMath({
        children: [
          new MathRun("E = mc²")
        ]
      })
    ]
  }))

  // 处理Markdown内容
  const tokens = marked.lexer(markdownContent)
  for (const token of tokens) {
    if (token.type === 'heading') {
      paragraphs.push(createHeading(token))
    } else if (token.type === 'paragraph') {
      paragraphs.push(createParagraph(token, []))
    }
  }

  // 页脚
  if (styleConfig.footer) {
    paragraphs.push(new Paragraph({
      children: [new TextRun({ text: styleConfig.footer })]
    }))
  }

  options.sections[0].children = paragraphs

  const doc = new Document(options)
  const blob = await Packer.toBlob(doc)

  return blob
}
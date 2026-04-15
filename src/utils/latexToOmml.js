/**
 * LaTeX 到 OMML 转换模块
 * OMML (Office Math Markup Language) 是 Word 的公式格式
 */

/**
 * LaTeX 基础元素到 OMML 映射表
 */
const LATEX_TO_OMML_MAP = {
  // 基础运算符
  '+': '<m:t>+</m:t>',
  '-': '<m:t>-</m:t>',
  '*': '<m:t>×</m:t>',
  '=': '<m:t>=</m:t>',
  '\\pm': '<m:t>±</m:t>',
  '\\mp': '<m:t>∓</m:t>',
  '\\times': '<m:t>×</m:t>',
  '\\div': '<m:t>÷</m:t>',
  '\\cdot': '<m:t>·</m:t>',
  '\\leq': '<m:t>≤</m:t>',
  '\\geq': '<m:t>≥</m:t>',
  '\\neq': '<m:t>≠</m:t>',
  '\\approx': '<m:t>≈</m:t>',
  '\\sim': '<m:t>∼</m:t>',
  '\\equiv': '<m:t>≡</m:t>',
  '\\propto': '<m:t>∝</m:t>',
  '\\parallel': '<m:t>∥</m:t>',
  '\\perp': '<m:t>⊥</m:t>',
  '\\in': '<m:t>∈</m:t>',
  '\\notin': '<m:t>∉</m:t>',
  '\\subset': '<m:t>⊂</m:t>',
  '\\supset': '<m:t>⊃</m:t>',
  '\\subseteq': '<m:t>⊆</m:t>',
  '\\supseteq': '<m:t>⊇</m:t>',
  '\\cup': '<m:t>∪</m:t>',
  '\\cap': '<m:t>∩</m:t>',
  '\\emptyset': '<m:t>∅</m:t>',
  '\\forall': '<m:t>∀</m:t>',
  '\\exists': '<m:t>∃</m:t>',
  '\\neg': '<m:t>¬</m:t>',
  '\\land': '<m:t>∧</m:t>',
  '\\lor': '<m:t>∨</m:t>',
  '\\rightarrow': '<m:t>→</m:t>',
  '\\leftarrow': '<m:t>←</m:t>',
  '\\Rightarrow': '<m:t>⇒</m:t>',
  '\\Leftarrow': '<m:t>⇐</m:t>',
  '\\leftrightarrow': '<m:t>↔</m:t>',
  '\\Leftrightarrow': '<m:t>⇔</m:t>',
  '\\infty': '<m:t>∞</m:t>',
  '\\partial': '<m:t>∂</m:t>',
  '\\nabla': '<m:t>∇</m:t>',
  '\\hbar': '<m:t>ℏ</m:t>',
  '\\ell': '<m:t>ℓ</m:t>',
  '\\Re': '<m:t>ℜ</m:t>',
  '\\Im': '<m:t>ℑ</m:t>',

  // 希腊字母
  '\\alpha': '<m:t>α</m:t>',
  '\\beta': '<m:t>β</m:t>',
  '\\gamma': '<m:t>γ</m:t>',
  '\\delta': '<m:t>δ</m:t>',
  '\\epsilon': '<m:t>ε</m:t>',
  '\\zeta': '<m:t>ζ</m:t>',
  '\\eta': '<m:t>η</m:t>',
  '\\theta': '<m:t>θ</m:t>',
  '\\iota': '<m:t>ι</m:t>',
  '\\kappa': '<m:t>κ</m:t>',
  '\\lambda': '<m:t>λ</m:t>',
  '\\mu': '<m:t>μ</m:t>',
  '\\nu': '<m:t>ν</m:t>',
  '\\xi': '<m:t>ξ</m:t>',
  '\\pi': '<m:t>π</m:t>',
  '\\rho': '<m:t>ρ</m:t>',
  '\\sigma': '<m:t>σ</m:t>',
  '\\tau': '<m:t>τ</m:t>',
  '\\upsilon': '<m:t>υ</m:t>',
  '\\phi': '<m:t>φ</m:t>',
  '\\chi': '<m:t>χ</m:t>',
  '\\psi': '<m:t>ψ</m:t>',
  '\\omega': '<m:t>ω</m:t>',
  '\\Gamma': '<m:t>Γ</m:t>',
  '\\Delta': '<m:t>Δ</m:t>',
  '\\Theta': '<m:t>Θ</m:t>',
  '\\Lambda': '<m:t>Λ</m:t>',
  '\\Xi': '<m:t>Ξ</m:t>',
  '\\Pi': '<m:t>Π</m:t>',
  '\\Sigma': '<m:t>Σ</m:t>',
  '\\Phi': '<m:t>Φ</m:t>',
  '\\Psi': '<m:t>Ψ</m:t>',
  '\\Omega': '<m:t>Ω</m:t>',
  '\\varepsilon': '<m:t>ɛ</m:t>',
  '\\varphi': '<m:t>ɸ</m:t>',
  '\\vartheta': '<m:t>ϑ</m:t>',
  '\\varpi': '<m:t>ϖ</m:t>',
  '\\varrho': '<m:t>ϱ</m:t>',
  '\\varsigma': '<m:t>ς</m:t>',

  // 括号
  '(' : '(',
  ')' : ')',
  '[' : '[',
  ']' : ']',
  '\\{': '{',
  '\\}': '}',
  '\\left(' : '(',
  '\\right)': ')',
  '\\left[' : '[',
  '\\right]': ']',
  '\\left\\{': '{',
  '\\right\\}': '}',
  '\\langle': '<m:t>⟨</m:t>',
  '\\rangle': '<m:t>⟩</m:t>',
  '\\lfloor': '<m:t>⌊</m:t>',
  '\\rfloor': '<m:t>⌋</m:t>',
  '\\lceil': '<m:t>⌈</m:t>',
  '\\rceil': '<m:t>⌉</m:t>',

  // 空格
  '\\,': '<m:t> </m:t>',
  '\\;': '<m:t>  </m:t>',
  '\\ ': '<m:t> </m:t>',
  '\\quad': '<m:t>    </m:t>',
  '\\qquad': '<m:t>        </m:t>',
}

/**
 * 解析LaTeX分数
 * \frac{a}{b} -> OMML fraction
 */
function parseFraction(latex) {
  const match = latex.match(/\\frac\s*\{([^}]*)\}\s*\{([^}]*)\}/)
  if (!match) return null

  const numerator = latexToOmml(match[1])
  const denominator = latexToOmml(match[2])

  return `<m:f>
    <m:fPr><m:type m:val="bar"/></m:fPr>
    <m:num>${numerator}</m:num>
    <m:den>${denominator}</m:den>
  </m:f>`
}

/**
 * 解析LaTeX根号
 * \sqrt{x} 或 \sqrt[n]{x} -> OMML radical
 */
function parseRadical(latex) {
  // 带索引的根号
  const matchWithIndex = latex.match(/\\sqrt\s*\[([^\]]*)\]\s*\{([^}]*)\}/)
  if (matchWithIndex) {
    const index = latexToOmml(matchWithIndex[1])
    const base = latexToOmml(matchWithIndex[2])
    return `<m:rad>
      <m:radPr><m:degHide m:val="0"/></m:radPr>
      <m:deg>${index}</m:deg>
      <m:e>${base}</m:e>
    </m:rad>`
  }

  // 普通根号
  const match = latex.match(/\\sqrt\s*\{([^}]*)\}/)
  if (match) {
    const base = latexToOmml(match[1])
    return `<m:rad>
      <m:radPr><m:degHide m:val="1"/></m:radPr>
      <m:deg/>
      <m:e>${base}</m:e>
    </m:rad>`
  }

  return null
}

/**
 * 解析LaTeX上下标
 * x^2, x_i, x_i^2 -> OMML sSup/sSub
 */
function parseScripts(latex) {
  // 匹配 x^{upper}_{lower} 或 x_{lower}^{upper}
  const patterns = [
    /([a-zA-Z0-9])\s*\^\s*\{([^}]*)\}\s*_\s*\{([^}]*)\}/,
    /([a-zA-Z0-9])\s*_\s*\{([^}]*)\}\s*\^\s*\{([^}]*)\}/,
    /([a-zA-Z0-9])\s*\^\s*\{([^}]*)\}/,
    /([a-zA-Z0-9])\s*_\s*\{([^}]*)\}/,
    /([a-zA-Z0-9])\s*\^\s*([a-zA-Z0-9])/,
    /([a-zA-Z0-9])\s*_\s*([a-zA-Z0-9])/,
  ]

  for (const pattern of patterns) {
    const match = latex.match(pattern)
    if (match) {
      const base = match[1]
      let upper, lower

      if (pattern === patterns[0]) {
        upper = match[2]
        lower = match[3]
      } else if (pattern === patterns[1]) {
        lower = match[2]
        upper = match[3]
      } else if (pattern === patterns[2] || pattern === patterns[4]) {
        upper = match[2]
      } else if (pattern === patterns[3] || pattern === patterns[5]) {
        lower = match[2]
      }

      let omml = `<m:sSup>`
      if (upper && lower) {
        omml = `<m:sSupSup>`
      } else if (lower && !upper) {
        omml = `<m:sSub>`
      }

      omml += `<m:e><m:r><m:t>${base}</m:t></m:r></m:e>`

      if (upper) {
        omml += `<m:sup>${latexToOmml(upper)}</m:sup>`
      }
      if (lower) {
        omml += `<m:sub>${latexToOmml(lower)}</m:sub>`
      }

      if (upper && lower) {
        omml += `</m:sSupSup>`
      } else if (lower && !upper) {
        omml += `</m:sSub>`
      } else {
        omml += `</m:sSup>`
      }

      return omml
    }
  }

  return null
}

/**
 * 解析LaTeX积分
 * \int_a^b f(x) dx -> OMML nary
 */
function parseIntegral(latex) {
  const patterns = [
    /\\int\s*_\s*\{([^}]*)\}\s*\^\s*\{([^}]*)\}/,
    /\\int\s*_\s*([a-zA-Z0-9])\s*\^\s*([a-zA-Z0-9])/,
    /\\int\s*\^\s*\{([^}]*)\}/,
    /\\int\s*_\s*\{([^}]*)\}/,
    /\\int/,
  ]

  for (const pattern of patterns) {
    const match = latex.match(pattern)
    if (match) {
      let lower, upper

      if (pattern === patterns[0]) {
        lower = match[1]
        upper = match[2]
      } else if (pattern === patterns[1]) {
        lower = match[1]
        upper = match[2]
      } else if (pattern === patterns[2]) {
        upper = match[1]
      } else if (pattern === patterns[3]) {
        lower = match[1]
      }

      let omml = `<m:nary>
        <m:naryPr><m:limLoc m:val="undOvr"/><m:chr m:val="∫"/></m:naryPr>`
      if (lower) {
        omml += `<m:sub>${latexToOmml(lower)}</m:sub>`
      }
      if (upper) {
        omml += `<m:sup>${latexToOmml(upper)}</m:sup>`
      }
      omml += `</m:nary>`

      return omml
    }
  }

  return null
}

/**
 * 解析LaTeX求和/求积
 * \sum_{i=1}^{n} -> OMML nary
 */
function parseSumProd(latex) {
  const ops = {
    '\\sum': '∑',
    '\\prod': '∏',
    '\\coprod': '∐',
    '\\bigcup': '⋃',
    '\\bigcap': '⋂',
    '\\bigoplus': '⊕',
    '\\bigotimes': '⊗',
    '\\bigodot': '⊙',
    '\\bigvee': '⋁',
    '\\bigwedge': '⋀',
  }

  for (const [op, chr] of Object.entries(ops)) {
    const patterns = [
      `${op}\\s*_\\s*\\{([^}]*)\\}\\s*\\^\\s*\\{([^}]*)\\}`,
      `${op}\\s*_\\s*([a-zA-Z0-9=]+)\\s*\\^\\s*([a-zA-Z0-9]+)`,
      `${op}\\s*\\^\\s*\\{([^}]*)\\}`,
      `${op}\\s*_\\s*\\{([^}]*)\\}`,
      `${op}`,
    ]

    for (const pattern of patterns) {
      const regex = new RegExp(pattern)
      const match = latex.match(regex)
      if (match) {
        let lower, upper

        if (match.length === 3 && pattern.includes('_') && pattern.includes('^')) {
          lower = match[1]
          upper = match[2]
        } else if (match.length === 2) {
          if (pattern.includes('_')) {
            lower = match[1]
          } else if (pattern.includes('^')) {
            upper = match[1]
          }
        }

        let omml = `<m:nary>
          <m:naryPr><m:limLoc m:val="undOvr"/><m:chr m:val="${chr}"/></m:naryPr>`
        if (lower) {
          omml += `<m:sub>${latexToOmml(lower)}</m:sub>`
        }
        if (upper) {
          omml += `<m:sup>${latexToOmml(upper)}</m:sup>`
        }
        omml += `</m:nary>`

        return omml
      }
    }
  }

  return null
}

/**
 * 解析LaTeX矩阵
 * \begin{matrix} ... \end{matrix} -> OMML m:m
 */
function parseMatrix(latex) {
  const matrixTypes = ['matrix', 'pmatrix', 'bmatrix', 'vmatrix', 'Vmatrix']
  let matrixType = null

  for (const type of matrixTypes) {
    if (latex.includes(`\\begin{${type}}`) && latex.includes(`\\end{${type}}`)) {
      matrixType = type
      break
    }
  }

  if (!matrixType) return null

  // 提取矩阵内容
  const contentMatch = latex.match(new RegExp(`\\\\begin\\{${matrixType}\\}([\\s\\S]+?)\\\\end\\{${matrixType}\\}`))
  if (!contentMatch) return null

  const content = contentMatch[1].trim()
  // 按行分割
  const rows = content.split('\\\\\\\\').map(row => row.trim())
  // 每行按 & 分割
  const cells = rows.map(row => row.split('&').map(cell => cell.trim()))

  // 确定括号类型
  let leftDelim = '', rightDelim = ''
  switch (matrixType) {
    case 'pmatrix': leftDelim = '('; rightDelim = ')'; break
    case 'bmatrix': leftDelim = '['; rightDelim = ']'; break
    case 'vmatrix': leftDelim = '|'; rightDelim = '|'; break
    case 'Vmatrix': leftDelim = '‖'; rightDelim = '‖'; break
  }

  let omml = '<m:m>'
  if (leftDelim) {
    omml += `<m:mPr><m:mcs><m:mc><m:mcPr><m:count m:val="${cells[0].length}"/></m:mcPr></m:mc></m:mcs></m:mPr>`
  }

  omml += '<m:mr>'
  for (const row of cells) {
    for (const cell of row) {
      omml += `<m:e>${latexToOmml(cell)}</m:e>`
    }
    omml += '</m:mr><m:mr>'
  }
  // 移除最后一个多余的 mr
  omml = omml.replace(/<m:mr><\/m:mr>$/, '')

  omml += '</m:m>'

  // 如果需要括号，用 m:d 包裹
  if (leftDelim) {
    omml = `<m:d>
      <m:dPr><m:begChr m:val="${leftDelim}"/><m:endChr m:val="${rightDelim}"/></m:dPr>
      <m:e>${omml}</m:e>
    </m:d>`
  }

  return omml
}

/**
 * 解析LaTeX括号
 * \left( ... \right) -> OMML m:d
 */
function parseDelimiters(latex) {
  const delimPairs = [
    ['\\left(', '\\right)', '(', ')'],
    ['\\left[', '\\right]', '[', ']'],
    ['\\left\\{', '\\right\\}', '{', '}'],
    ['\\left|', '\\right|', '|', '|'],
    ['\\left\\|', '\\right\\|', '‖', '‖'],
    ['\\left\\langle', '\\right\\rangle', '⟨', '⟩'],
  ]

  for (const [leftLatex, rightLatex, leftChar, rightChar] of delimPairs) {
    if (latex.includes(leftLatex) && latex.includes(rightLatex)) {
      const contentMatch = latex.match(new RegExp(`${leftLatex.replace(/\\/g, '\\\\\\\\')}(.+?)${rightLatex.replace(/\\/g, '\\\\\\\\')}`))
      if (contentMatch) {
        const content = latexToOmml(contentMatch[1])
        return `<m:d>
          <m:dPr><m:begChr m:val="${leftChar}"/><m:endChr m:val="${rightChar}"/></m:dPr>
          <m:e>${content}</m:e>
        </m:d>`
      }
    }
  }

  return null
}

/**
 * 主转换函数：将LaTeX转换为OMML
 * @param {string} latex LaTeX公式字符串
 * @returns {string} OMML XML字符串
 */
export function latexToOmml(latex) {
  if (!latex) return '<m:r><m:t></m:t></m:r>'

  let result = ''
  let remaining = latex.trim()

  // 处理复合结构（优先级高）
  // 1. 矩阵
  const matrixResult = parseMatrix(remaining)
  if (matrixResult) {
    return matrixResult
  }

  // 2. 括号
  const delimResult = parseDelimiters(remaining)
  if (delimResult) {
    return delimResult
  }

  // 3. 分数
  const fracMatch = remaining.match(/\\frac\s*\{[^}]*\}\s*\{[^}]*\}/)
  if (fracMatch) {
    const fracResult = parseFraction(fracMatch[0])
    if (fracResult) {
      remaining = remaining.replace(fracMatch[0], '')
      result += fracResult
    }
  }

  // 4. 根号
  const sqrtMatch = remaining.match(/\\sqrt(\s*\[[^\]]*\])?\s*\{[^}]*\}/)
  if (sqrtMatch) {
    const sqrtResult = parseRadical(sqrtMatch[0])
    if (sqrtResult) {
      remaining = remaining.replace(sqrtMatch[0], '')
      result += sqrtResult
    }
  }

  // 5. 积分
  const intMatch = remaining.match(/\\int(\s*_\s*(\{[^}]*\}|[a-zA-Z0-9]))?(\s*\^\s*(\{[^}]*\}|[a-zA-Z0-9]))?/)
  if (intMatch) {
    const intResult = parseIntegral(intMatch[0])
    if (intResult) {
      remaining = remaining.replace(intMatch[0], '')
      result += intResult
    }
  }

  // 6. 求和求积
  const sumMatch = remaining.match(/\\(sum|prod|coprod|bigcup|bigcap|bigoplus|bigotimes|bigodot|bigvee|bigwedge)(\s*_\s*(\{[^}]*\}|[a-zA-Z0-9=]+))?(\s*\^\s*(\{[^}]*\}|[a-zA-Z0-9]+))?/)
  if (sumMatch) {
    const sumResult = parseSumProd(sumMatch[0])
    if (sumResult) {
      remaining = remaining.replace(sumMatch[0], '')
      result += sumResult
    }
  }

  // 7. 上下标
  const scriptMatch = remaining.match(/[a-zA-Z0-9]\s*[\^_]\s*(\{[^}]*\}|[a-zA-Z0-9])/)
  if (scriptMatch) {
    const scriptResult = parseScripts(scriptMatch[0])
    if (scriptResult) {
      remaining = remaining.replace(scriptMatch[0], '')
      result += scriptResult
    }
  }

  // 8. 处理剩余的简单符号和文本
  for (const [latexCmd, omml] of Object.entries(LATEX_TO_OMML_MAP)) {
    if (remaining.includes(latexCmd)) {
      remaining = remaining.replace(latexCmd, '')
      result += omml
    }
  }

  // 9. 处理剩余的普通文本
  remaining = remaining.replace(/[{}\s]/g, '').replace(/\\\\/g, '')
  if (remaining) {
    result += `<m:r><m:t>${remaining}</m:t></m:r>`
  }

  // 如果结果为空，返回空文本
  if (!result) {
    result = `<m:r><m:t>${latex}</m:t></m:r>`
  }

  return result
}

/**
 * 创建完整的OMML公式元素
 * @param {string} ommlContent OMML内容
 * @param {boolean} isBlock 是否为块级公式
 * @returns {string} 完整的OMML XML
 */
export function createOmmlElement(ommlContent, isBlock = false) {
  const omml = `<m:oMathPara>
    <m:oMath>${ommlContent}</m:oMath>
  </m:oMathPara>`

  if (isBlock) {
    return `<w:p><w:pPr><w:jc w:val="center"/></w:pPr>${omml}</w:p>`
  }

  return omml
}

/**
 * 将LaTeX公式转换为Word可嵌入的OMML结构
 * 用于直接嵌入到docx库生成的文档中
 * @param {string} latex LaTeX公式
 * @param {boolean} isBlock 是否为块级公式
 * @returns {Object} 包含OMML的对象
 */
export function convertLatexForWord(latex, isBlock = false) {
  const ommlContent = latexToOmml(latex)
  return {
    type: 'formula',
    isBlock,
    omml: createOmmlElement(ommlContent, isBlock),
    rawContent: latex
  }
}
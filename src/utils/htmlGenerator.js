/**
 * HTML生成模块
 * 导出带有完整样式的HTML文件
 * 复用PDF生成模块的HTML模板逻辑
 */

import { getPdfHtml } from './pdfGenerator'

/**
 * 生成HTML文件
 * @param {string} markdownContent Markdown内容
 * @param {Object} styleConfig 样式配置
 * @returns {Promise<Blob>} HTML文件Blob
 */
export async function generateHtml(markdownContent, styleConfig = {}) {
  try {
    // 获取完整的HTML内容（包含KaTeX CSS和样式）
    const htmlContent = getPdfHtml(markdownContent, styleConfig)

    // 创建HTML Blob
    return new Blob([htmlContent], { type: 'text/html;charset=utf-8' })
  } catch (error) {
    console.error('HTML生成错误:', error)
    throw error
  }
}

/**
 * 获取HTML内容（用于预览或直接下载）
 * @param {string} markdownContent Markdown内容
 * @param {Object} styleConfig 样式配置
 * @returns {string} HTML字符串
 */
export function getHtmlContent(markdownContent, styleConfig = {}) {
  return getPdfHtml(markdownContent, styleConfig)
}
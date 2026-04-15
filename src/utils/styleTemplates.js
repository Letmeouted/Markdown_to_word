/**
 * 样式模板管理模块
 */

/**
 * 默认样式模板
 */
export const defaultTemplates = {
  default: {
    name: '默认模板',
    pageSize: 'A4',
    marginTop: 25,
    marginBottom: 25,
    marginLeft: 20,
    marginRight: 20,
    fontFamily: 'SimSun',
    fontSize: 12,
    header: '',
    footer: ''
  },
  academic: {
    name: '学术论文模板',
    pageSize: 'A4',
    marginTop: 25,
    marginBottom: 25,
    marginLeft: 30,
    marginRight: 30,
    fontFamily: 'Times New Roman',
    fontSize: 12,
    header: '',
    footer: '页码'
  },
  report: {
    name: '报告模板',
    pageSize: 'A4',
    marginTop: 20,
    marginBottom: 20,
    marginLeft: 25,
    marginRight: 25,
    fontFamily: 'Microsoft YaHei',
    fontSize: 14,
    header: '',
    footer: ''
  },
  compact: {
    name: '紧凑模板',
    pageSize: 'A5',
    marginTop: 15,
    marginBottom: 15,
    marginLeft: 15,
    marginRight: 15,
    fontFamily: 'SimSun',
    fontSize: 10.5,
    header: '',
    footer: ''
  }
}

/**
 * 加载用户自定义模板
 * @param {string} appPath 应用数据路径
 * @returns {Promise<Object>} 用户模板
 */
export async function loadUserTemplates(appPath) {
  try {
    const templatePath = `${appPath}/templates/`
    const templates = {}

    // 在实际应用中，这里需要使用Electron的文件系统API读取
    // const files = await fs.readdir(templatePath)
    // for (const file of files) {
    //   if (file.endsWith('.json')) {
    //     const content = await fs.readFile(`${templatePath}/${file}`, 'utf-8')
    //     templates[file.replace('.json', '')] = JSON.parse(content)
    //   }
    // }

    return templates
  } catch (error) {
    console.error('加载用户模板失败:', error)
    return {}
  }
}

/**
 * 保存用户模板
 * @param {string} appPath 应用数据路径
 * @param {string} name 模板名称
 * @param {Object} config 模板配置
 * @returns {Promise<boolean>} 是否成功
 */
export async function saveUserTemplate(appPath, name, config) {
  try {
    const templatePath = `${appPath}/templates/${name}.json`

    // 在实际应用中，这里需要使用Electron的文件系统API写入
    // await fs.writeFile(templatePath, JSON.stringify(config, null, 2))

    return true
  } catch (error) {
    console.error('保存用户模板失败:', error)
    return false
  }
}

/**
 * 删除用户模板
 * @param {string} appPath 应用数据路径
 * @param {string} name 模板名称
 * @returns {Promise<boolean>} 是否成功
 */
export async function deleteUserTemplate(appPath, name) {
  try {
    const templatePath = `${appPath}/templates/${name}.json`

    // 在实际应用中，这里需要使用Electron的文件系统API删除
    // await fs.unlink(templatePath)

    return true
  } catch (error) {
    console.error('删除用户模板失败:', error)
    return false
  }
}

/**
 * 合并模板配置
 * @param {Object} base 基础配置
 * @param {Object} override 覆盖配置
 * @returns {Object} 合合后的配置
 */
export function mergeConfig(base, override) {
  return {
    ...defaultTemplates.default,
    ...base,
    ...override
  }
}

/**
 * 验证样式配置
 * @param {Object} config 样式配置
 * @returns {Object} 验证后的配置
 */
export function validateConfig(config) {
  const validated = { ...config }

  // 页面大小
  const validPageSizes = ['A4', 'A5', 'Letter']
  if (!validPageSizes.includes(validated.pageSize)) {
    validated.pageSize = 'A4'
  }

  // 边距范围验证 (0-100mm)
  validated.marginTop = Math.max(0, Math.min(100, Number(validated.marginTop) || 25))
  validated.marginBottom = Math.max(0, Math.min(100, Number(validated.marginBottom) || 25))
  validated.marginLeft = Math.max(0, Math.min(100, Number(validated.marginLeft) || 20))
  validated.marginRight = Math.max(0, Math.min(100, Number(validated.marginRight) || 20))

  // 字体
  const validFonts = ['SimSun', 'Microsoft YaHei', 'Arial', 'Times New Roman']
  if (!validFonts.includes(validated.fontFamily)) {
    validated.fontFamily = 'SimSun'
  }

  // 字号范围验证 (8-24pt)
  validated.fontSize = Math.max(8, Math.min(24, Number(validated.fontSize) || 12))

  return validated
}
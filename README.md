# Markdown Converter

一个功能强大的 Markdown 转 Word/PDF 桌面应用程序，支持 LaTeX 数学公式的完美渲染。

📄 [英文版本](README_EN.md)

## 功能特点

- **Markdown 转 Word (.docx)**：将 Markdown 文档转换为专业的 Word 文档，LaTeX 公式以可编辑的公式对象形式呈现
- **Markdown 转 PDF**：将 Markdown 文档转换为 PDF 文件，LaTeX 公式以专业数学字体渲染
- **LaTeX 公式支持**：
  - 行内公式：`$...$` 和 `\(...\)`
  - 块级公式：`$$...$$` 和 `\[...\]`
  - 下标上标：`x_i`, `x^2`
  - 分数：`\frac{a}{b}`
  - 常用符号：`\ge`, `\le`, `\alpha`, `\sum`, `\int` 等
- **实时预览**：左侧编辑，右侧实时预览渲染效果
- **样式配置**：支持页面大小、边距、字体、字号等自定义设置
- **模板管理**：支持保存和加载样式模板

## 技术栈

- **前端框架**：Vue 3 + Vite
- **桌面框架**：Electron 28
- **UI组件**：Element Plus
- **Markdown解析**：marked
- **公式渲染**：KaTeX
- **Word生成**：docx (JavaScript Office Open XML 库)
- **PDF生成**：Electron printToPDF API

## 项目结构

```
markdown-converter/
├── electron/                  # Electron 主进程
│   ├── main.js               # 主进程入口
│   └── preload.js           # 预加载脚本（IPC通信）
├── src/                      # 前端源码
│   ├── App.vue              # 主应用组件
│   ├── main.js              # Vue 入口
│   ├── components/          # Vue 组件
│   ├── utils/               # 工具模块
│   │   ├── docxGenerator.js # Word文档生成器
│   │   ├── pdfGenerator.js  # PDF生成器
│   │   ├── markdownParser.js# Markdown解析器
│   │   ├── latexToOmml.js   # LaTeX转OMML（Word公式）
│   │   └── styleTemplates.js# 样式模板
│   ├── styles/              # 样式文件
│   └── templates/           # 模板配置
├── build/                    # 构建资源（图标等）
├── dist/                     # 前端构建输出
├── package.json              # 项目配置
├── vite.config.js           # Vite配置
└── electron-builder.yml     # Electron打包配置
```

## 安装与运行

### 环境要求

- Node.js 18+
- npm 或 yarn

### 安装依赖

```bash
npm install
```

### 开发模式

```bash
npm run electron:dev
```

启动后，Vite 开发服务器和 Electron 应用将同时运行，支持热重载。

### 构建打包

```bash
# 构建前端
npm run build

# 打包 Windows 应用
npm run electron:build:win

# 打包 macOS 应用
npm run electron:build:mac
```

打包后的应用位于 `dist_electron/` 目录。

## 使用说明

### 基本操作

1. **打开文件**：点击"选择文件"按钮，选择 Markdown 文件（.md, .markdown, .txt）
2. **编辑内容**：左侧文本框支持直接编辑 Markdown 内容
3. **实时预览**：右侧实时显示渲染后的效果
4. **样式设置**：点击"样式设置"按钮，自定义页面样式
5. **导出文档**：
   - 点击"导出 Word"生成 Word 文档
   - 点击"导出 PDF"生成 PDF 文件

### 样式配置

| 配置项 | 说明 | 默认值 |
|--------|------|--------|
| 页面大小 | A4, A5, Letter | A4 |
| 边距（上/下/左/右） | 页面边距（mm） | 25/25/20/20 |
| 字体 | 宋体, 微软雅黑, Arial, Times New Roman | 宋体 |
| 字号 | 10.5pt, 12pt, 14pt, 16pt | 12pt |
| 页眉 | 可选文字 | 无 |
| 页脚 | 可选文字 | 无 |

### LaTeX 公式示例

```markdown
# 行内公式
能量公式：$E = mc^2$

# 块级公式
$$
\frac{1}{1 + e^{-x}}
$$

# 复杂公式
$$
\int_0^\infty e^{-x^2} dx = \frac{\sqrt{\pi}}{2}
$$

# 下标上标
温度 $T_1$ 和 $T_2$ 的关系：$T_2 \ge T_1$
```

## 核心模块说明

### Word 生成器 (docxGenerator.js)

- 使用 docx 库生成标准 Office Open XML 格式
- LaTeX 公式转换为 OMML（Office Math Markup Language）
- 支持标题、段落、列表、表格、代码块等所有 Markdown 元素

### PDF 生成器 (pdfGenerator.js)

- 使用 KaTeX 服务端渲染公式
- 嵌入完整 KaTeX CSS 确保正确显示
- 使用 Cambria Math 系统字体确保跨平台兼容
- 通过 Electron printToPDF API 生成真正的 PDF 文件

### Markdown 解析器 (markdownParser.js)

- 基于 marked 库解析 Markdown
- 提取和保护 LaTeX 公式避免解析冲突
- 支持 GitHub Flavored Markdown (GFM)

## 常见问题

### PDF 中公式显示不正确？

确保系统安装了 Cambria Math 或 Times New Roman 字体。这些是 Windows 系统默认字体，通常不需要额外安装。

### Word 中公式无法编辑？

Word 文档中的公式以 OMML 格式存储，需要 Microsoft Word 2007+ 或 WPS Office 才能正确显示和编辑。

### 应用启动失败？

检查 Node.js 版本是否满足要求（18+），并确保所有依赖已正确安装：
```bash
npm install
```

## 开发计划

- [ ] 支持登录注册
- [ ] 支持更多 LaTeX 命令
- [ ] 支持导出html文件
- [ ] 图片支持
- [ ] 批量转换
- [ ] 深色模式
- [ ] 国际化支持

## 许可证

MIT License

## 作者

Markdown转换工具团队: Wang Huafeng

图标资源提供者: Yan Lijuan
# Markdown Converter

A powerful Markdown to Word/PDF desktop application with excellent LaTeX math formula rendering support.

📄 [Chinese Version](README.md)

## Features

- **Markdown to Word (.docx)**: Convert Markdown documents to professional Word documents with LaTeX formulas rendered as editable formula objects
- **Markdown to PDF**: Convert Markdown documents to PDF files with professional math font rendering
- **LaTeX Formula Support**:
  - Inline formulas: `$...$` and `\(...\)`
  - Block formulas: `$$...$$` and `\[...\]`
  - Subscripts and superscripts: `x_i`, `x^2`
  - Fractions: `\frac{a}{b}`
  - Common symbols: `\ge`, `\le`, `\alpha`, `\sum`, `\int`, etc.
- **Real-time Preview**: Edit on the left, preview rendered output on the right in real-time
- **Style Configuration**: Customize page size, margins, fonts, font sizes, and more
- **Template Management**: Save and load style templates

## Tech Stack

- **Frontend Framework**: Vue 3 + Vite
- **Desktop Framework**: Electron 28
- **UI Components**: Element Plus
- **Markdown Parser**: marked
- **Formula Rendering**: KaTeX
- **Word Generation**: docx (JavaScript Office Open XML Library)
- **PDF Generation**: Electron printToPDF API

## Project Structure

```
markdown-converter/
├── electron/                  # Electron main process
│   ├── main.js               # Main process entry
│   └── preload.js           # Preload script (IPC communication)
├── src/                      # Frontend source code
│   ├── App.vue              # Main application component
│   ├── main.js              # Vue entry point
│   ├── components/          # Vue components
│   ├── utils/               # Utility modules
│   │   ├── docxGenerator.js # Word document generator
│   │   ├── pdfGenerator.js  # PDF generator
│   │   ├── markdownParser.js# Markdown parser
│   │   ├── latexToOmml.js   # LaTeX to OMML (Word formulas)
│   │   └── styleTemplates.js# Style templates
│   ├── styles/              # Style files
│   └── templates/           # Template configurations
├── build/                    # Build resources (icons, etc.)
├── dist/                     # Frontend build output
├── package.json              # Project configuration
├── vite.config.js           # Vite configuration
└── electron-builder.yml     # Electron build configuration
```

## Installation & Usage

### Requirements

- Node.js 18+
- npm or yarn

### Install Dependencies

```bash
npm install
```

### Development Mode

```bash
npm run electron:dev
```

After startup, both the Vite development server and Electron application will run simultaneously with hot reload support.

### Build & Package

```bash
# Build frontend
npm run build

# Package Windows app
npm run electron:build:win

# Package macOS app
npm run electron:build:mac
```

The packaged application will be located in the `dist_electron/` directory.

## User Guide

### Basic Operations

1. **Open File**: Click the "Select File" button to choose a Markdown file (.md, .markdown, .txt)
2. **Edit Content**: The left text box supports direct editing of Markdown content
3. **Real-time Preview**: The right side displays the rendered output in real-time
4. **Style Settings**: Click the "Style Settings" button to customize page styles
5. **Export Document**:
   - Click "Export Word" to generate a Word document
   - Click "Export PDF" to generate a PDF file

### Style Configuration

| Setting | Description | Default |
|---------|-------------|---------|
| Page Size | A4, A5, Letter | A4 |
| Margins (Top/Bottom/Left/Right) | Page margins (mm) | 25/25/20/20 |
| Font | Songti, Microsoft YaHei, Arial, Times New Roman | Songti |
| Font Size | 10.5pt, 12pt, 14pt, 16pt | 12pt |
| Header | Optional text | None |
| Footer | Optional text | None |

### LaTeX Formula Examples

```markdown
# Inline Formulas
Energy formula: $E = mc^2$

# Block Formulas
$$
\frac{1}{1 + e^{-x}}
$$

# Complex Formulas
$$
\int_0^\infty e^{-x^2} dx = \frac{\sqrt{\pi}}{2}
$$

# Subscripts and Superscripts
Relationship between temperatures $T_1$ and $T_2$: $T_2 \ge T_1$
```

## Core Modules

### Word Generator (docxGenerator.js)

- Uses the docx library to generate standard Office Open XML format
- Converts LaTeX formulas to OMML (Office Math Markup Language)
- Supports all Markdown elements: headings, paragraphs, lists, tables, code blocks, etc.

### PDF Generator (pdfGenerator.js)

- Uses KaTeX server-side rendering for formulas
- Embeds complete KaTeX CSS for correct display
- Uses Cambria Math system font for cross-platform compatibility
- Generates true PDF files via Electron printToPDF API

### Markdown Parser (markdownParser.js)

- Parses Markdown based on the marked library
- Extracts and protects LaTeX formulas to avoid parsing conflicts
- Supports GitHub Flavored Markdown (GFM)

## FAQ

### Formulas not displaying correctly in PDF?

Ensure that Cambria Math or Times New Roman fonts are installed on your system. These are default Windows system fonts and usually don't require additional installation.

### Cannot edit formulas in Word?

Formulas in Word documents are stored in OMML format and require Microsoft Word 2007+ or WPS Office to correctly display and edit them.

### Application fails to start?

Check if the Node.js version meets the requirements (18+) and ensure all dependencies are properly installed:
```bash
npm install
```

## Development Roadmap

- [ ] Support user login/registration
- [ ] Support more LaTeX commands
- [x] Support HTML export
- [ ] Image support
- [ ] Batch conversion
- [ ] Dark mode
- [ ] Internationalization support

## License

MIT License

## Authors

Markdown Converter Team: Wang Huafeng

Icon resources provided by: Yan Lijuan
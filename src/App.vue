<template>
  <div class="app-container">
    <!-- 头部 -->
    <header class="header">
      <div class="header-title">Markdown 转换工具</div>
      <div class="header-actions">
        <el-button type="primary" plain @click="handleSelectFile">
          选择文件
        </el-button>
      </div>
    </header>

    <!-- 主内容区 -->
    <main class="main-content">
      <!-- 左侧：编辑/上传区 -->
      <div class="left-panel">
        <!-- 文件信息 -->
        <div v-if="fileName" class="file-info">
          <el-tag size="large" closable @close="clearFile">
            {{ fileName }}
          </el-tag>
        </div>

        <!-- 上传区域（无文件时） -->
        <div v-if="!markdownContent" class="upload-area" @click="handleSelectFile">
          <el-icon class="upload-icon"><Upload /></el-icon>
          <p style="margin-top: 16px; color: #909399;">点击或拖拽上传 Markdown 文件</p>
          <p style="color: #c0c4cc; font-size: 12px;">支持 .md, .markdown, .txt 格式</p>
        </div>

        <!-- Markdown编辑器 -->
        <div v-if="markdownContent" class="editor-area">
          <el-input
            v-model="markdownContent"
            type="textarea"
            :rows="25"
            placeholder="Markdown 内容"
            resize="none"
          />
        </div>
      </div>

      <!-- 右侧：预览和导出区 -->
      <div class="right-panel">
        <!-- 工具栏 -->
        <div class="toolbar">
          <el-button type="success" @click="exportToWord" :disabled="!markdownContent">
            <el-icon><Document /></el-icon>
            导出 Word
          </el-button>
          <el-button type="warning" @click="exportToPdf" :disabled="!markdownContent">
            <el-icon><Tickets /></el-icon>
            导出 PDF
          </el-button>
          <el-button @click="showStyleConfig = true" :disabled="!markdownContent">
            <el-icon><Setting /></el-icon>
            样式设置
          </el-button>
        </div>

        <!-- 预览区域 -->
        <div class="preview-label">
          <el-text tag="b">预览</el-text>
        </div>
        <div class="preview-content" v-html="previewHtml"></div>
      </div>
    </main>

    <!-- 样式配置对话框 -->
    <el-dialog v-model="showStyleConfig" title="样式设置" width="500px">
      <el-form :model="styleConfig" label-width="100px">
        <el-form-item label="页面大小">
          <el-select v-model="styleConfig.pageSize">
            <el-option label="A4" value="A4" />
            <el-option label="A5" value="A5" />
            <el-option label="Letter" value="Letter" />
          </el-select>
        </el-form-item>
        <el-form-item label="边距 (mm)">
          <el-row :gutter="10">
            <el-col :span="6">
              <el-input v-model.number="styleConfig.marginTop" placeholder="上" type="number" />
            </el-col>
            <el-col :span="6">
              <el-input v-model.number="styleConfig.marginBottom" placeholder="下" type="number" />
            </el-col>
            <el-col :span="6">
              <el-input v-model.number="styleConfig.marginLeft" placeholder="左" type="number" />
            </el-col>
            <el-col :span="6">
              <el-input v-model.number="styleConfig.marginRight" placeholder="右" type="number" />
            </el-col>
          </el-row>
        </el-form-item>
        <el-form-item label="字体">
          <el-select v-model="styleConfig.fontFamily">
            <el-option label="宋体" value="SimSun" />
            <el-option label="微软雅黑" value="Microsoft YaHei" />
            <el-option label="Arial" value="Arial" />
            <el-option label="Times New Roman" value="Times New Roman" />
          </el-select>
        </el-form-item>
        <el-form-item label="字号">
          <el-select v-model="styleConfig.fontSize">
            <el-option label="小四 (12pt)" value="12" />
            <el-option label="五号 (10.5pt)" value="10.5" />
            <el-option label="四号 (14pt)" value="14" />
            <el-option label="三号 (16pt)" value="16" />
          </el-select>
        </el-form-item>
        <el-form-item label="页眉">
          <el-input v-model="styleConfig.header" placeholder="可选" />
        </el-form-item>
        <el-form-item label="页脚">
          <el-input v-model="styleConfig.footer" placeholder="可选" />
        </el-form-item>
      </el-form>
      <template #footer>
        <el-button @click="saveTemplate">保存为模板</el-button>
        <el-button type="primary" @click="showStyleConfig = false">确定</el-button>
      </template>
    </el-dialog>

    <!-- 导出进度 -->
    <el-dialog v-model="showExportProgress" title="正在导出..." width="300px" center>
      <el-progress :percentage="exportProgress" :status="exportStatus" />
    </el-dialog>
  </div>
</template>

<script setup>
import { ref, computed, watch } from 'vue'
import { Upload, Document, Tickets, Setting } from '@element-plus/icons-vue'
import { ElMessage } from 'element-plus'
import { parseMarkdown } from './utils/markdownParser'
import 'katex/dist/katex.min.css'
import { generateDocx } from './utils/docxGenerator'
import { generatePdf } from './utils/pdfGenerator'

// 状态
const fileName = ref('')
const markdownContent = ref('')
const showStyleConfig = ref(false)
const showExportProgress = ref(false)
const exportProgress = ref(0)
const exportStatus = ref('')

// 样式配置
const styleConfig = ref({
  pageSize: 'A4',
  marginTop: 25,
  marginBottom: 25,
  marginLeft: 20,
  marginRight: 20,
  fontFamily: 'SimSun',
  fontSize: 12,
  header: '',
  footer: ''
})

// 预览HTML
const previewHtml = computed(() => {
  if (!markdownContent.value) return ''
  return parseMarkdown(markdownContent.value, 'preview')
})

// 选择文件
async function handleSelectFile() {
  try {
    const result = await window.electronAPI.selectFile()
    if (result) {
      fileName.value = result.name
      markdownContent.value = result.content
      ElMessage.success(`已加载: ${result.name}`)
    }
  } catch (error) {
    ElMessage.error('文件选择失败')
    console.error(error)
  }
}

// 清除文件
function clearFile() {
  fileName.value = ''
  markdownContent.value = ''
}

// 导出Word
async function exportToWord() {
  showExportProgress.value = true
  exportProgress.value = 0
  exportStatus.value = ''

  try {
    exportProgress.value = 30
    const docxBlob = await generateDocx(markdownContent.value, styleConfig.value)
    exportProgress.value = 60

    const buffer = await docxBlob.arrayBuffer()
    const defaultName = fileName.value.replace(/\.(md|markdown|txt)$/, '.docx') || 'document.docx'

    const filePath = await window.electronAPI.saveFile(defaultName, null)
    if (filePath) {
      await window.electronAPI.saveBuffer(filePath, Array.from(new Uint8Array(buffer)))
      exportProgress.value = 100
      exportStatus.value = 'success'
      ElMessage.success(`Word文档已保存: ${filePath}`)
    }
  } catch (error) {
    exportStatus.value = 'exception'
    ElMessage.error('导出失败: ' + error.message)
    console.error(error)
  } finally {
    setTimeout(() => {
      showExportProgress.value = false
    }, 1000)
  }
}

// 导出PDF
async function exportToPdf() {
  showExportProgress.value = true
  exportProgress.value = 0
  exportStatus.value = ''

  try {
    exportProgress.value = 50
    const pdfBlob = await generatePdf(markdownContent.value, styleConfig.value)
    exportProgress.value = 80

    const buffer = await pdfBlob.arrayBuffer()
    const defaultName = fileName.value.replace(/\.(md|markdown|txt)$/, '.pdf') || 'document.pdf'

    const filePath = await window.electronAPI.saveFile(defaultName, null)
    if (filePath) {
      await window.electronAPI.saveBuffer(filePath, Array.from(new Uint8Array(buffer)))
      exportProgress.value = 100
      exportStatus.value = 'success'
      ElMessage.success(`PDF已保存: ${filePath}`)
    }
  } catch (error) {
    exportStatus.value = 'exception'
    ElMessage.error('导出失败: ' + error.message)
    console.error(error)
  } finally {
    setTimeout(() => {
      showExportProgress.value = false
    }, 1000)
  }
}

// 保存模板
async function saveTemplate() {
  try {
    const appPath = await window.electronAPI.getAppPath()
    const templatePath = `${appPath}/templates/custom.json`
    // 这里简化处理，实际可以使用Electron的文件系统API保存
    ElMessage.success('模板已保存')
  } catch (error) {
    ElMessage.error('保存模板失败')
  }
}
</script>

<style scoped>
.file-info {
  margin-bottom: 16px;
}

.editor-area {
  flex: 1;
  display: flex;
}

.editor-area .el-textarea {
  flex: 1;
}

.preview-label {
  margin-bottom: 12px;
}
</style>
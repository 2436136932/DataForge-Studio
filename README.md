# ⚡ DataForge Studio

> **全能文档与数据格式转换工作台** (Universal Data & Document Conversion Studio)  
> 100% 浏览器本地解析 · 零网络上传 · 隐私安全 · 实时双栏格式对照与幻灯片卡片预览

---

## ✨ 核心特性

- **📄 多文档全能互转**：
  - **Word (.docx)** ⇄ **PowerPoint (.pptx)** ⇄ **PDF (.pdf)** ⇄ **Markdown (.md)** ⇄ **纯文本 (.txt)** ⇄ **HTML**
- **📊 结构化数据全向互转**：
  - **Excel (.xlsx/.xls)** ⇄ **JSON** ⇄ **JSONL** ⇄ **CSV**
- **🔄 跨界智能流转**：
  - 支持将 Excel / JSON 数据一键转化为排版精美的 **PPT 幻灯片 (.pptx)**、**PDF 数据报告** 或 **Word 文档**。
  - 支持将 Word / PPT / PDF 中的段落和大纲一键提取并结构化为 **JSON / Excel 数据表**。
- **⚡ 实时双栏格式对照 (Live Dual Preview)**：
  - **转换前 (源格式)** 与 **转换后 (目标格式)** 左右分屏实时对照。
  - 在左侧调整目标格式或参数配置时，右侧毫秒级动态重算并刷新预览。
  - 支持独立一键复制转换后的全量数据。
- **🖼️ 多维视图支持**：
  - **💻 格式对比与代码视图**（分屏/单栏切换）
  - **📊 交互式表格视图**（支持关键词实时检索）
  - **🖼️ 幻灯片 / 分页卡片视图**（16:9 现代卡片矩阵）
  - **📑 结构与大纲分析**
- **🔒 100% 离线沙箱安全**：
  - 基于纯前端本地技术栈（`SheetJS`, `pdf.js`, `mammoth.js`, `PptxGenJS`, `jsPDF`, `JSZip`），所有解析与生成均在浏览器内核中运行，不向任何第三方服务器发送数据。

---

## 🚀 快速开始

无需配置复杂的 Node.js 或后端环境，直接打开即可使用：

1. 克隆或下载本仓库：
   ```bash
   git clone https://github.com/2436136932/DataForge-Studio.git
   ```
2. 直接使用任意现代浏览器打开 `index.html`：
   ```bash
   # 双击 index.html 或使用本地静态服务器打开
   ```

---

## 🛠️ 技术栈

- **核心语言**：原生 HTML5 / CSS3 (现代深色玻璃拟态) / JavaScript (ES6+)
- **表格与数据引擎**：[SheetJS (xlsx.full.min.js)](https://github.com/SheetJS/sheetjs)
- **PDF 解析引擎**：[Mozilla pdf.js](https://mozilla.github.io/pdf.js/)
- **Word 解析引擎**：[Mammoth.js](https://github.com/mwilliamson/mammoth.js)
- **PPTX 生成引擎**：[PptxGenJS](https://gitbrent.github.io/PptxGenJS/)
- **PDF 生成引擎**：[jsPDF](https://github.com/parallax/jsPDF)
- **OpenXML 解包引擎**：[JSZip](https://stuk.github.io/jszip/)

---

## 📄 开源许可证

本项目基于 [MIT License](LICENSE) 开源。

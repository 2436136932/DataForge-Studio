// script.js - 全能文档与数据转换工作台核心交互引擎 (Word / PPT / PDF / Excel / JSON / CSV / Markdown)

// DOM 元素引用 - 左侧控制台
const fileInput = document.getElementById('fileInput');
const dropZone = document.getElementById('dropZone');
const dropPrompt = document.getElementById('dropPrompt');
const dropHint = document.getElementById('dropHint');
const fileLoadedCard = document.getElementById('fileLoadedCard');
const fileCardIcon = document.getElementById('fileCardIcon');
const fileNameDisplay = document.getElementById('fileNameDisplay');
const fileMetaDisplay = document.getElementById('fileMetaDisplay');
const removeFileBtn = document.getElementById('removeFileBtn');
const fileBadge = document.getElementById('fileBadge');

const srcFormatSelect = document.getElementById('srcFormat');
const targetFormatSelect = document.getElementById('targetFormat');
const swapFormatBtn = document.getElementById('swapFormatBtn');

const optPrettyJson = document.getElementById('optPrettyJson');
const optCsvBom = document.getElementById('optCsvBom');
const optDocHeaders = document.getElementById('optDocHeaders');
const optPptxTheme = document.getElementById('optPptxTheme');

const convertBtn = document.getElementById('convertBtn');
const statusDiv = document.getElementById('status');

// 历史记录与工作区
const historyCard = document.getElementById('historyCard');
const historyList = document.getElementById('historyList');
const clearHistoryBtn = document.getElementById('clearHistoryBtn');

const emptyState = document.getElementById('emptyState');
const tableFilterBar = document.getElementById('tableFilterBar');
const tableSearchInput = document.getElementById('tableSearchInput');
const filterTip = document.getElementById('filterTip');
const tableHead = document.getElementById('tableHead');
const tableBody = document.getElementById('tableBody');
const schemaBody = document.getElementById('schemaBody');

// 幻灯片与分页卡片容器
const slideDeckBadge = document.getElementById('slideDeckBadge');
const slideDeckCount = document.getElementById('slideDeckCount');
const slideCardsGrid = document.getElementById('slideCardsGrid');

// 视角模式切换控制
const previewModeGroup = document.getElementById('previewModeGroup');
const modeButtons = document.querySelectorAll('.mode-pill-btn');
const codeSplitWrapper = document.getElementById('codeSplitWrapper');

// 转换前 (源) 元素
const sourceCodePanel = document.getElementById('sourceCodePanel');
const sourceFormatTag = document.getElementById('sourceFormatTag');
const sourceMetaTag = document.getElementById('sourceMetaTag');
const sourceVisualViewport = document.getElementById('sourceVisualViewport');
const sourceCodeArea = document.getElementById('sourceCodeArea');
const sourceCodeContent = document.getElementById('sourceCodeContent');
const copySourceCodeBtn = document.getElementById('copySourceCodeBtn');
const toggleSourceVisualBtn = document.getElementById('toggleSourceVisualBtn');

// 转换后 (目标) 元素
const targetCodePanel = document.getElementById('targetCodePanel');
const targetFormatTag = document.getElementById('targetFormatTag');
const targetMetaTag = document.getElementById('targetMetaTag');
const targetVisualViewport = document.getElementById('targetVisualViewport');
const targetCodeArea = document.getElementById('targetCodeArea');
const targetCodeContent = document.getElementById('targetCodeContent');
const copyTargetCodeBtn = document.getElementById('copyTargetCodeBtn');
const toggleTargetVisualBtn = document.getElementById('toggleTargetVisualBtn');

// 视图模式状态 (源与目标独立)
let isSourceVisual = true;
let isTargetVisual = true;

// 切换源视图模式
if (toggleSourceVisualBtn) {
  toggleSourceVisualBtn.addEventListener('click', () => {
    isSourceVisual = !isSourceVisual;
    if (isSourceVisual) {
      toggleSourceVisualBtn.classList.add('active');
      toggleSourceVisualBtn.textContent = '👁️ 真实渲染';
      sourceVisualViewport.style.display = 'block';
      sourceCodeArea.style.display = 'none';
    } else {
      toggleSourceVisualBtn.classList.remove('active');
      toggleSourceVisualBtn.textContent = '💻 纯代码';
      sourceVisualViewport.style.display = 'none';
      sourceCodeArea.style.display = 'block';
    }
  });
}

// 切换目标视图模式
if (toggleTargetVisualBtn) {
  toggleTargetVisualBtn.addEventListener('click', () => {
    isTargetVisual = !isTargetVisual;
    if (isTargetVisual) {
      toggleTargetVisualBtn.classList.add('active');
      toggleTargetVisualBtn.textContent = '👁️ 真实渲染';
      targetVisualViewport.style.display = 'block';
      targetCodeArea.style.display = 'none';
    } else {
      toggleTargetVisualBtn.classList.remove('active');
      toggleTargetVisualBtn.textContent = '💻 纯代码';
      targetVisualViewport.style.display = 'none';
      targetCodeArea.style.display = 'block';
    }
  });
}

const splitDivider = document.getElementById('splitDivider');

const statRows = document.getElementById('statRows');
const statCols = document.getElementById('statCols');
const statSize = document.getElementById('statSize');

// 全局状态
let currentFile = null;
let parsedResult = null;
let parsedDataset = [];
let sourceRawContent = '';
let currentPreviewMode = 'split'; // 'split' | 'source' | 'target'

// 格式映射字典
const formatAcceptMap = {
  docx: '.docx',
  pptx: '.pptx',
  pdf: '.pdf',
  md: '.md',
  txt: '.txt',
  html: '.html,.htm',
  xlsx: '.xlsx',
  xls: '.xls',
  json: '.json',
  jsonl: '.jsonl',
  csv: '.csv'
};

const formatHints = {
  docx: '支持 Microsoft Word (.docx) 文档',
  pptx: '支持 PowerPoint (.pptx) 演示文稿',
  pdf: '支持 Adobe PDF (.pdf) 文档',
  md: '支持 Markdown (.md) 结构化文档',
  txt: '支持纯文本 (.txt) 文件',
  html: '支持 HTML (.html/.htm) 网页文件',
  xlsx: '支持 Excel 2007+ (.xlsx) 工作簿',
  xls: '支持 Excel 97-2003 (.xls) 工作簿',
  json: '支持 JSON 数组或对象 (.json) 文件',
  jsonl: '支持行分隔 JSON Lines (.jsonl) 文件',
  csv: '支持逗号分隔数据 (.csv) 文件'
};

const formatLabels = {
  docx: 'Word 文档 (.docx)',
  pptx: 'PowerPoint 幻灯片 (.pptx)',
  pdf: 'PDF 导出文档 (.pdf)',
  md: 'Markdown 文档 (.md)',
  txt: '纯文本 (.txt)',
  html: 'HTML 网页 (.html)',
  xlsx: 'Excel 工作簿 (.xlsx)',
  xls: 'Excel 97-2003 (.xls)',
  json: 'JSON 数据 (.json)',
  jsonl: 'JSONL 行记录 (.jsonl)',
  csv: 'CSV 逗号分隔 (.csv)'
};

const formatIcons = {
  docx: '📘',
  pptx: '📙',
  pdf: '📕',
  md: '📝',
  txt: '📄',
  html: '🌐',
  xlsx: '📊',
  xls: '📈',
  json: '💻',
  jsonl: '📋',
  csv: '📑'
};

// 格式化文件大小
function formatBytes(bytes) {
  if (!bytes || bytes === 0) return '0 B';
  const k = 1024;
  const sizes = ['B', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

// 设置状态提示
function setStatus(msg, type = '') {
  statusDiv.textContent = msg;
  statusDiv.className = 'status-banner' + (type ? ' ' + type : '');
}

const conversionIntentTip = document.getElementById('conversionIntentTip');
const intentText = document.getElementById('intentText');

// 智能转换推荐规则库 (基于数据语义与办公实际场景)
const smartRecommendationRules = {
  pptx: {
    recommended: [
      { fmt: 'pdf', label: 'PDF 导出讲义 (.pdf) · 便于打印/分发' },
      { fmt: 'docx', label: 'Word 文档 (.docx) · 提炼演讲大纲与纪要' },
      { fmt: 'md', label: 'Markdown 文档 (.md) · 知识库大纲整理' },
      { fmt: 'html', label: 'HTML 网页 (.html) · 浏览器幻灯片展示' }
    ],
    advanced: [
      { fmt: 'txt', label: '纯文本 (.txt) · 提取文字流' },
      { fmt: 'json', label: 'JSON 数据 (.json) · 幻灯片结构对象' },
      { fmt: 'xlsx', label: 'Excel 工作簿 (.xlsx) · 提取为单列文本表格' },
      { fmt: 'csv', label: 'CSV 逗号分隔 (.csv) · 提取为数据行' }
    ],
    hint: '💡 推荐将 PPT 导出为 PDF 讲义打印分发，或提炼为 Word / Markdown 大纲纪要。'
  },
  docx: {
    recommended: [
      { fmt: 'pdf', label: 'PDF 导出文档 (.pdf) · 标准排版归档/防篡改' },
      { fmt: 'pptx', label: 'PowerPoint 演示文稿 (.pptx) · 章节大纲一键转幻灯片' },
      { fmt: 'md', label: 'Markdown 文档 (.md) · 技术文档/知识库' },
      { fmt: 'html', label: 'HTML 网页 (.html) · Web 页面发布' }
    ],
    advanced: [
      { fmt: 'xlsx', label: 'Excel 工作簿 (.xlsx) · 章节转化为数据行' },
      { fmt: 'txt', label: '纯文本 (.txt) · 去除排版提取纯文本' },
      { fmt: 'json', label: 'JSON 数据 (.json) · 段落结构化对象' },
      { fmt: 'csv', label: 'CSV 逗号分隔 (.csv)' }
    ],
    hint: '💡 推荐将 Word 转为 PDF 进行标准化归档，或将大纲转化为 PPT 商务演示。'
  },
  pdf: {
    recommended: [
      { fmt: 'docx', label: 'Word 文档 (.docx) · 提取文本与二次编辑' },
      { fmt: 'pptx', label: 'PowerPoint 演示文稿 (.pptx) · 逐页转化为幻灯片' },
      { fmt: 'md', label: 'Markdown 结构化大纲 (.md)' },
      { fmt: 'txt', label: '纯文本 (.txt) · 提取全文字流' }
    ],
    advanced: [
      { fmt: 'xlsx', label: 'Excel 工作簿 (.xlsx) · 提取数据行(适合表格型PDF)' },
      { fmt: 'json', label: 'JSON 数据 (.json)' },
      { fmt: 'html', label: 'HTML 网页 (.html)' },
      { fmt: 'csv', label: 'CSV 逗号分隔 (.csv)' }
    ],
    hint: '💡 推荐将 PDF 提取至 Word/Markdown 进行可编辑流转，或转为 PPT 演示文稿。'
  },
  xlsx: {
    recommended: [
      { fmt: 'csv', label: 'CSV 逗号分隔 (.csv) · 数据清洗与数据库导入' },
      { fmt: 'json', label: 'JSON 数据 (.json) · 前端与 API 结构化对象' },
      { fmt: 'jsonl', label: 'JSONL 行记录 (.jsonl) · AI 训练集/大数据' },
      { fmt: 'pdf', label: 'PDF 数据报表 (.pdf) · 斑马条纹 A4 打印报表' },
      { fmt: 'docx', label: 'Word 文档 (.docx) · 排版表格文档' },
      { fmt: 'pptx', label: 'PowerPoint 演示文稿 (.pptx) · 商务数据图表幻灯片' }
    ],
    advanced: [
      { fmt: 'xls', label: 'Excel 97-2003 (.xls)' },
      { fmt: 'html', label: 'HTML 网页 (.html) · 网页数据表格' },
      { fmt: 'md', label: 'Markdown 文档 (.md)' },
      { fmt: 'txt', label: '纯文本 (.txt)' }
    ],
    hint: '💡 推荐将 Excel 清洗为 CSV/JSON，或一键排版生成 PDF 报表与 PPT 商务演示。'
  },
  xls: {
    recommended: [
      { fmt: 'xlsx', label: 'Excel 2007+ (.xlsx) · 升级至最新工作簿标准' },
      { fmt: 'csv', label: 'CSV 逗号分隔 (.csv)' },
      { fmt: 'json', label: 'JSON 结构化数据 (.json)' },
      { fmt: 'pdf', label: 'PDF 数据报表 (.pdf)' }
    ],
    advanced: [
      { fmt: 'jsonl', label: 'JSONL 行记录 (.jsonl)' },
      { fmt: 'docx', label: 'Word 文档 (.docx)' },
      { fmt: 'pptx', label: 'PowerPoint 演示文稿 (.pptx)' },
      { fmt: 'html', label: 'HTML 网页 (.html)' }
    ],
    hint: '💡 推荐将旧版 XLS 升级为现代 XLSX 格式，或导出为标准 CSV/JSON。'
  },
  json: {
    recommended: [
      { fmt: 'xlsx', label: 'Excel 工作簿 (.xlsx) · 扁平化导出电子表格' },
      { fmt: 'jsonl', label: 'JSONL 行记录 (.jsonl) · 单行数据流转换' },
      { fmt: 'csv', label: 'CSV 逗号分隔 (.csv) · 数据库与分析导入' },
      { fmt: 'pdf', label: 'PDF 矢量数据报表 (.pdf)' },
      { fmt: 'docx', label: 'Word 表格文档 (.docx)' }
    ],
    advanced: [
      { fmt: 'pptx', label: 'PowerPoint 商务幻灯片 (.pptx)' },
      { fmt: 'html', label: 'HTML 网页 (.html)' },
      { fmt: 'md', label: 'Markdown 结构化文档 (.md)' },
      { fmt: 'txt', label: '纯文本 (.txt)' }
    ],
    hint: '💡 推荐将 JSON 数据一键转化为 Excel 电子表格，或转换 JSONL/CSV 进行数据流转。'
  },
  jsonl: {
    recommended: [
      { fmt: 'xlsx', label: 'Excel 工作簿 (.xlsx) · 结构化电子表格' },
      { fmt: 'json', label: 'JSON 数组格式 (.json) · 标准对象数组' },
      { fmt: 'csv', label: 'CSV 逗号分隔 (.csv)' },
      { fmt: 'pdf', label: 'PDF 数据报表 (.pdf)' }
    ],
    advanced: [
      { fmt: 'docx', label: 'Word 表格文档 (.docx)' },
      { fmt: 'pptx', label: 'PowerPoint 幻灯片 (.pptx)' },
      { fmt: 'html', label: 'HTML 网页 (.html)' },
      { fmt: 'md', label: 'Markdown 文档 (.md)' }
    ],
    hint: '💡 推荐将 JSONL 记录集导出为 Excel/CSV 或格式化为标准 JSON 数组。'
  },
  csv: {
    recommended: [
      { fmt: 'xlsx', label: 'Excel 工作簿 (.xlsx) · 多功能电子表格' },
      { fmt: 'json', label: 'JSON 结构化数据 (.json)' },
      { fmt: 'jsonl', label: 'JSONL 行记录流 (.jsonl)' },
      { fmt: 'pdf', label: 'PDF 矢量报表 (.pdf)' }
    ],
    advanced: [
      { fmt: 'docx', label: 'Word 表格文档 (.docx)' },
      { fmt: 'pptx', label: 'PowerPoint 演示文稿 (.pptx)' },
      { fmt: 'html', label: 'HTML 数据网页 (.html)' },
      { fmt: 'md', label: 'Markdown 文档 (.md)' }
    ],
    hint: '💡 推荐将 CSV 快速转换为带有网格的 Excel 工作簿或结构化 JSON 对象。'
  },
  md: {
    recommended: [
      { fmt: 'docx', label: 'Word 文档 (.docx) · 专业排版打印导出' },
      { fmt: 'pdf', label: 'PDF 导出文档 (.pdf) · 标准文档发布' },
      { fmt: 'html', label: 'HTML 网页 (.html) · Web 富文本展示' },
      { fmt: 'pptx', label: 'PowerPoint 演示文稿 (.pptx) · 大纲转幻灯片' }
    ],
    advanced: [
      { fmt: 'txt', label: '纯文本 (.txt)' },
      { fmt: 'xlsx', label: 'Excel 表格 (.xlsx)' },
      { fmt: 'json', label: 'JSON 数据 (.json)' }
    ],
    hint: '💡 推荐将 Markdown 导出为 Word/PDF 专业排版文档，或快速生成 PPT 幻灯片。'
  },
  txt: {
    recommended: [
      { fmt: 'docx', label: 'Word 文档 (.docx) · 转换为排版文档' },
      { fmt: 'pdf', label: 'PDF 导出 (.pdf) · 打印版面' },
      { fmt: 'md', label: 'Markdown 文档 (.md) · 结构化整理' }
    ],
    advanced: [
      { fmt: 'html', label: 'HTML 网页 (.html)' },
      { fmt: 'pptx', label: 'PowerPoint (.pptx)' },
      { fmt: 'xlsx', label: 'Excel 工作簿 (.xlsx)' },
      { fmt: 'json', label: 'JSON 数据 (.json)' }
    ],
    hint: '💡 推荐将纯文本整理为 Word 或 Markdown 格式。'
  },
  html: {
    recommended: [
      { fmt: 'pdf', label: 'PDF 矢量打印 (.pdf)' },
      { fmt: 'docx', label: 'Word 文档 (.docx) · 网页转排版文档' },
      { fmt: 'md', label: 'Markdown 文档 (.md) · 网页转技术大纲' }
    ],
    advanced: [
      { fmt: 'xlsx', label: 'Excel 工作簿 (.xlsx)' },
      { fmt: 'pptx', label: 'PowerPoint (.pptx)' },
      { fmt: 'txt', label: '纯文本 (.txt)' },
      { fmt: 'json', label: 'JSON 数据 (.json)' }
    ],
    hint: '💡 推荐将 HTML 网页内容导出为标准 PDF 或 Word 文档。'
  }
};

// 更新目标格式下拉菜单 (按 🔥 最佳推荐常用 + ⚙️ 辅助与全能导出 分组)
function updateTargetOptions() {
  const currentSrc = srcFormatSelect.value;
  const currentTarget = targetFormatSelect.value;

  fileInput.accept = formatAcceptMap[currentSrc] || '';
  dropHint.textContent = formatHints[currentSrc] || '';

  const rule = smartRecommendationRules[currentSrc] || {
    recommended: [{ fmt: 'docx', label: 'Word 文档 (.docx)' }, { fmt: 'pdf', label: 'PDF 导出文档 (.pdf)' }],
    advanced: [{ fmt: 'xlsx', label: 'Excel 工作簿 (.xlsx)' }, { fmt: 'json', label: 'JSON 数据 (.json)' }],
    hint: '💡 选择最适合您的目标导出格式。'
  };

  // 更新意图提示条
  if (intentText) {
    intentText.textContent = rule.hint;
  }

  targetFormatSelect.innerHTML = '';

  // 1. 🔥 最佳推荐常用 (自然高价值流向)
  if (rule.recommended && rule.recommended.length > 0) {
    const optGroupRec = document.createElement('optgroup');
    optGroupRec.label = '🔥 最佳推荐常用 (自然高价值流向)';
    rule.recommended.forEach(item => {
      const opt = document.createElement('option');
      opt.value = item.fmt;
      opt.textContent = item.label;
      optGroupRec.appendChild(opt);
    });
    targetFormatSelect.appendChild(optGroupRec);
  }

  // 2. ⚙️ 辅助与全能导出 (特殊提取需求)
  if (rule.advanced && rule.advanced.length > 0) {
    const optGroupAdv = document.createElement('optgroup');
    optGroupAdv.label = '⚙️ 辅助与全能导出 (特殊提取需求)';
    rule.advanced.forEach(item => {
      const opt = document.createElement('option');
      opt.value = item.fmt;
      opt.textContent = item.label;
      optGroupAdv.appendChild(opt);
    });
    targetFormatSelect.appendChild(optGroupAdv);
  }

  // 保持原有选择或智能选中推荐第一项
  if (currentTarget && targetFormatSelect.querySelector(`option[value="${currentTarget}"]`)) {
    targetFormatSelect.value = currentTarget;
  } else if (rule.recommended && rule.recommended.length > 0) {
    targetFormatSelect.value = rule.recommended[0].fmt;
  }
}

// 视角模式切换处理
function setPreviewMode(mode) {
  currentPreviewMode = mode;
  modeButtons.forEach(btn => {
    if (btn.dataset.mode === mode) {
      btn.classList.add('active');
    } else {
      btn.classList.remove('active');
    }
  });

  codeSplitWrapper.className = `code-split-wrapper view-mode-${mode}`;
}

modeButtons.forEach(btn => {
  btn.addEventListener('click', () => {
    setPreviewMode(btn.dataset.mode);
  });
});

// 交换格式方向
swapFormatBtn.addEventListener('click', () => {
  const oldSrc = srcFormatSelect.value;
  const oldTarget = targetFormatSelect.value;

  const allFormats = Object.keys(formatAcceptMap);
  if (allFormats.includes(oldTarget)) {
    srcFormatSelect.value = oldTarget;
    updateTargetOptions();
    targetFormatSelect.value = oldSrc === 'xls' ? 'xlsx' : oldSrc;
    if (currentFile) {
      loadAndPreviewFile(currentFile);
    }
  }
});

// 源格式切换监听
srcFormatSelect.addEventListener('change', () => {
  updateTargetOptions();
  if (currentFile) {
    loadAndPreviewFile(currentFile);
  }
});

// 目标格式切换或导出选项变化监听：实时重算目标格式预览与选项有效性
function updateOptionAvailability() {
  const targetFmt = targetFormatSelect.value;

  if (optPrettyJson) {
    const wrap = optPrettyJson.closest('.toggle-item') || optPrettyJson.parentElement;
    if (wrap) wrap.style.opacity = (targetFmt === 'json' || targetFmt === 'jsonl') ? '1' : '0.45';
  }
  if (optCsvBom) {
    const wrap = optCsvBom.closest('.toggle-item') || optCsvBom.parentElement;
    if (wrap) wrap.style.opacity = (targetFmt === 'csv') ? '1' : '0.45';
  }
  if (optDocHeaders) {
    const wrap = optDocHeaders.closest('.toggle-item') || optDocHeaders.parentElement;
    if (wrap) wrap.style.opacity = (['pdf', 'docx', 'md', 'html'].includes(targetFmt)) ? '1' : '0.45';
  }
  if (optPptxTheme) {
    const wrap = optPptxTheme.closest('.toggle-item') || optPptxTheme.parentElement;
    if (wrap) wrap.style.opacity = (targetFmt === 'pptx') ? '1' : '0.45';
  }
}

function onTargetFormatOrOptionChange() {
  updateOptionAvailability();
  if (parsedDataset && parsedDataset.length > 0) {
    updateTargetPreview();
  }
}

targetFormatSelect.addEventListener('change', onTargetFormatOrOptionChange);
optPrettyJson.addEventListener('change', onTargetFormatOrOptionChange);
optCsvBom.addEventListener('change', onTargetFormatOrOptionChange);
if (optDocHeaders) optDocHeaders.addEventListener('change', onTargetFormatOrOptionChange);
if (optPptxTheme) optPptxTheme.addEventListener('change', onTargetFormatOrOptionChange);

// 页面初始化时执行一次选项高亮
updateOptionAvailability();

// Tab 切换逻辑 (保持分屏胶囊与表格过滤条的整洁联动)
document.querySelectorAll('.tab-btn').forEach(btn => {
  btn.addEventListener('click', () => {
    document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
    document.querySelectorAll('.tab-content').forEach(c => (c.style.display = 'none'));

    btn.classList.add('active');
    const targetTab = document.getElementById(btn.dataset.tab);
    if (targetTab) {
      targetTab.style.display = 'flex';
    }

    // 仅在「格式对比与代码」视图下展示三模切换胶囊
    if (btn.dataset.tab === 'tabCode') {
      if (previewModeGroup) previewModeGroup.style.display = 'flex';
    } else {
      if (previewModeGroup) previewModeGroup.style.display = 'none';
    }

    if (btn.dataset.tab === 'tabTable' && parsedDataset.length > 0) {
      tableFilterBar.style.display = 'flex';
    } else {
      tableFilterBar.style.display = 'none';
    }
  });
});

// 数据与多文档解析及实时工作区渲染
async function loadAndPreviewFile(file) {
  currentFile = file;
  if (!file) return;

  // 智能根据文件后缀名自动匹配源格式
  const detectedExt = file.name.split('.').pop().toLowerCase();
  if (Object.keys(formatAcceptMap).includes(detectedExt)) {
    if (srcFormatSelect.value !== detectedExt) {
      srcFormatSelect.value = detectedExt;
      updateTargetOptions();
    }
  }

  // 更新左侧文件卡片展示
  dropPrompt.style.display = 'none';
  fileLoadedCard.style.display = 'flex';
  fileCardIcon.textContent = formatIcons[detectedExt] || '📄';
  fileNameDisplay.textContent = file.name;
  fileMetaDisplay.textContent = `${formatBytes(file.size)} · ${file.name.split('.').pop().toUpperCase()}`;
  fileBadge.style.display = 'inline-block';
  convertBtn.disabled = true;
  setStatus('正在解析文档结构并生成实时预览...', 'info');

  try {
    const srcFormat = srcFormatSelect.value;
    parsedResult = await window.FileConverter.parseFile(file, srcFormat);
    parsedDataset = Array.isArray(parsedResult.data) ? parsedResult.data : [];
    sourceRawContent = parsedResult.rawText || '';

    if (parsedDataset.length === 0) {
      throw new Error('未在文件中解析到有效数据或段落内容');
    }

    // 统计指标与要素提取 (大数据极速采样优化：超过 500 行进行安全抽样，避免几十兆大文件卡顿)
    const sampleSet = parsedDataset.length > 500 ? parsedDataset.slice(0, 500) : parsedDataset;
    const allKeys = Array.from(
      sampleSet.reduce((keys, row) => {
        if (typeof row === 'object' && row !== null) {
          Object.keys(row).forEach(k => keys.add(k));
        }
        return keys;
      }, new Set())
    );

    const countLabel = parsedResult.slides ? `${parsedResult.slides.length} 页幻灯片`
      : parsedResult.pages ? `${parsedResult.pages.length} 页文档`
      : `${parsedDataset.length} 条记录`;

    statRows.textContent = countLabel;
    statCols.textContent = `${allKeys.length} 个要素`;
    statSize.textContent = formatBytes(file.size);

    // 渲染各个视图
    renderTableView(parsedDataset, allKeys);
    renderCodeSplitView(parsedDataset, srcFormat);
    renderSlideCardsView(parsedResult);
    renderSchemaView(parsedDataset, allKeys);

    // 显示工作区视图
    emptyState.style.display = 'none';
    const activeTabBtn = document.querySelector('.tab-btn.active');
    const activeTabId = activeTabBtn ? activeTabBtn.dataset.tab : 'tabCode';
    document.getElementById(activeTabId).style.display = 'flex';
    if (activeTabId === 'tabTable') tableFilterBar.style.display = 'flex';

    convertBtn.disabled = false;
    setStatus(`✓ 文档加载就绪，成功解析 ${countLabel}，实时多维预览已生成`, 'success');
  } catch (err) {
    console.error(err);
    convertBtn.disabled = true;
    setStatus(`解析失败: ${err.message || '文件格式不匹配'}`, 'error');
  }
}

// 渲染表格视图
function renderTableView(data, keys, filterQuery = '') {
  tableHead.innerHTML = '';
  tableBody.innerHTML = '';

  if (!keys || keys.length === 0) return;

  // 表头
  const trHead = document.createElement('tr');
  const thIdx = document.createElement('th');
  thIdx.textContent = '#';
  thIdx.className = 'col-index';
  trHead.appendChild(thIdx);

  keys.forEach(key => {
    const th = document.createElement('th');
    th.textContent = key;
    trHead.appendChild(th);
  });
  tableHead.appendChild(trHead);

  // 过滤数据
  let displayData = data;
  if (filterQuery.trim() !== '') {
    const q = filterQuery.toLowerCase();
    displayData = data.filter(row =>
      keys.some(k => row[k] !== undefined && row[k] !== null && String(row[k]).toLowerCase().includes(q))
    );
  }

  const maxPreview = 100;
  const slice = displayData.slice(0, maxPreview);
  filterTip.textContent = `显示 ${slice.length} / ${displayData.length} 项${displayData.length > maxPreview ? '（仅预览前 100 项）' : ''}`;

  slice.forEach((row, index) => {
    const tr = document.createElement('tr');
    const tdIdx = document.createElement('td');
    tdIdx.textContent = index + 1;
    tdIdx.className = 'cell-index';
    tr.appendChild(tdIdx);

    keys.forEach(key => {
      const td = document.createElement('td');
      const val = row[key];
      if (val === null || val === undefined) {
        td.innerHTML = '<span class="cell-null">null</span>';
      } else if (typeof val === 'object') {
        td.textContent = JSON.stringify(val);
        td.title = JSON.stringify(val);
      } else {
        td.textContent = val;
        td.title = val;
      }
      tr.appendChild(td);
    });
    tableBody.appendChild(tr);
  });
}

// 渲染双栏真实仿真与格式对比视图
function renderCodeSplitView(data, srcFormat) {
  // 1. 渲染转换前 (源) 仿真视窗与代码
  const srcUpper = srcFormat.toUpperCase();
  sourceFormatTag.textContent = srcUpper;
  sourceMetaTag.textContent = `${data.length} 项要素`;

  const options = {
    title: currentFile ? currentFile.name.replace(/\.[^/.]+$/, '') : '源文档',
    prettyJson: optPrettyJson.checked,
    bom: optCsvBom.checked,
    docHeaders: optDocHeaders ? optDocHeaders.checked : true,
    pptxTheme: optPptxTheme ? optPptxTheme.checked : true
  };

  // 生成真实仿真视图
  if (sourceVisualViewport) {
    sourceVisualViewport.innerHTML = window.FileConverter.renderVisualDocumentHtml(data, srcFormat, options, parsedResult);
  }

  // 纯文本源码
  if (sourceRawContent && sourceRawContent.trim().length > 0) {
    const rawLines = sourceRawContent.split('\n');
    if (rawLines.length > 70) {
      sourceCodeContent.textContent = rawLines.slice(0, 70).join('\n') + `\n\n// ... 还有 ${rawLines.length - 70} 行 (已截取前 70 行快速预览)`;
    } else {
      sourceCodeContent.textContent = sourceRawContent;
    }
  } else {
    const sample = data.slice(0, 50);
    const jsonStr = JSON.stringify(sample, null, 2);
    sourceCodeContent.textContent = jsonStr + (data.length > 50 ? `\n\n// ... 以及其余 ${data.length - 50} 条记录` : '');
  }

  // 2. 渲染转换后 (目标) 实时视图
  updateTargetPreview();
}

// 实时计算并刷新目标格式真实仿真预览
function updateTargetPreview() {
  if (!parsedDataset || parsedDataset.length === 0) return;

  const targetFormat = targetFormatSelect.value;
  const options = {
    title: currentFile ? currentFile.name.replace(/\.[^/.]+$/, '') : '转换文档',
    prettyJson: optPrettyJson.checked,
    bom: optCsvBom.checked,
    docHeaders: optDocHeaders ? optDocHeaders.checked : true,
    pptxTheme: optPptxTheme ? optPptxTheme.checked : true
  };

  const formatLabels = {
    docx: 'DOCX (Word A4 标准排版)',
    pdf: 'PDF (A4 矢量打印排版)',
    pptx: 'PPTX (16:9 幻灯片放映)',
    xlsx: 'XLSX (Excel 电子表格)',
    xls: 'XLS (Excel 工作表)',
    json: 'JSON (结构化对象)',
    jsonl: 'JSONL (行记录流)',
    csv: 'CSV (逗号分隔数据)',
    md: 'Markdown (结构化大纲)',
    txt: '纯文本 (标准文本流)',
    html: 'HTML (网页排版格式)'
  };

  targetFormatTag.textContent = formatLabels[targetFormat] || `${targetFormat.toUpperCase()} (实时)`;

  // 1. 渲染真实仿真视窗
  if (targetVisualViewport) {
    targetVisualViewport.innerHTML = window.FileConverter.renderVisualDocumentHtml(parsedDataset, targetFormat, options, parsedResult);
  }

  // 2. 渲染纯代码文本与统计
  const preview = window.FileConverter.generatePreviewText(parsedDataset, targetFormat, options, 60);
  targetMetaTag.textContent = preview.isTruncated
    ? `预览前 ${preview.previewRows} / 共 ${preview.totalRows} 项`
    : `共 ${preview.totalRows} 项`;

  targetCodeContent.textContent = preview.text;
}

// 渲染幻灯片与页面卡片视图 (富排版与高对比度卡片)
function renderSlideCardsView(result) {
  slideCardsGrid.innerHTML = '';

  if (result.slides && result.slides.length > 0) {
    slideDeckBadge.textContent = '📙 PowerPoint 幻灯片大纲视图';
    slideDeckCount.textContent = `共 ${result.slides.length} 页幻灯片`;

    result.slides.forEach(s => {
      const cardEl = document.createElement('div');
      cardEl.className = 'slide-card-item slide-type-ppt';

      const topEl = document.createElement('div');
      topEl.className = 'slide-card-top';
      topEl.innerHTML = `
        <span class="slide-card-pill pill-ppt">SLIDE ${s.index}</span>
        <h4 class="slide-card-heading" title="${s.title}">${s.title}</h4>
      `;

      const bodyEl = document.createElement('div');
      bodyEl.className = 'slide-card-body';

      if (s.points && s.points.length > 0) {
        const ul = document.createElement('ul');
        ul.className = 'slide-bullet-list';
        s.points.forEach(pt => {
          const li = document.createElement('li');
          li.textContent = pt;
          ul.appendChild(li);
        });
        bodyEl.appendChild(ul);
      } else {
        bodyEl.innerHTML = `<div class="slide-empty-hint">${s.rawText || '（单张封面或空白幻灯片）'}</div>`;
      }

      cardEl.appendChild(topEl);
      cardEl.appendChild(bodyEl);
      slideCardsGrid.appendChild(cardEl);
    });

  } else if (result.pages && result.pages.length > 0) {
    slideDeckBadge.textContent = '📕 PDF 页面逐页解析视图';
    slideDeckCount.textContent = `共 ${result.pages.length} 页 PDF`;

    result.pages.forEach(p => {
      const cardEl = document.createElement('div');
      cardEl.className = 'slide-card-item slide-type-pdf';

      const topEl = document.createElement('div');
      topEl.className = 'slide-card-top';
      topEl.innerHTML = `
        <span class="slide-card-pill pill-pdf">PAGE ${p.pageNum}</span>
        <h4 class="slide-card-heading">第 ${p.pageNum} 页文本流</h4>
      `;

      const bodyEl = document.createElement('div');
      bodyEl.className = 'slide-card-body';
      bodyEl.innerHTML = `<p class="slide-doc-para">${p.text || '（页面无可见文本）'}</p>`;

      cardEl.appendChild(topEl);
      cardEl.appendChild(bodyEl);
      slideCardsGrid.appendChild(cardEl);
    });

  } else if (result.sections && result.sections.length > 0) {
    slideDeckBadge.textContent = '📘 Word / 文档章节排版视图';
    slideDeckCount.textContent = `共 ${result.sections.length} 个章节/段落`;

    result.sections.forEach((sec, idx) => {
      const cardEl = document.createElement('div');
      cardEl.className = 'slide-card-item slide-type-doc';

      const topEl = document.createElement('div');
      topEl.className = 'slide-card-top';
      topEl.innerHTML = `
        <span class="slide-card-pill pill-doc">SEC ${idx + 1}</span>
        <h4 class="slide-card-heading" title="${sec.title || `段落 #${idx + 1}`}">${sec.title || `段落 #${idx + 1}`}</h4>
      `;

      const bodyEl = document.createElement('div');
      bodyEl.className = 'slide-card-body';

      if (Array.isArray(sec.points) && sec.points.length > 0) {
        const ul = document.createElement('ul');
        ul.className = 'slide-bullet-list';
        sec.points.forEach(pt => {
          const li = document.createElement('li');
          li.textContent = pt;
          ul.appendChild(li);
        });
        bodyEl.appendChild(ul);
      } else {
        bodyEl.innerHTML = `<p class="slide-doc-para">${sec.text || sec.content || '（章节内容）'}</p>`;
      }

      cardEl.appendChild(topEl);
      cardEl.appendChild(bodyEl);
      slideCardsGrid.appendChild(cardEl);
    });

  } else {
    // 表格 / JSON / 数据集分组卡片展示
    slideDeckBadge.textContent = '📊 数据卡片分组矩阵视图';
    const chunkSize = 4;
    const totalChunks = Math.min(Math.ceil(parsedDataset.length / chunkSize), 24);
    slideDeckCount.textContent = `共 ${parsedDataset.length} 条记录 (抽样渲染前 ${totalChunks} 个精装卡片)`;

    for (let i = 0; i < totalChunks; i++) {
      const chunk = parsedDataset.slice(i * chunkSize, (i + 1) * chunkSize);
      const cardEl = document.createElement('div');
      cardEl.className = 'slide-card-item slide-type-data';

      const topEl = document.createElement('div');
      topEl.className = 'slide-card-top';
      topEl.innerHTML = `
        <span class="slide-card-pill pill-data">GROUP ${i + 1}</span>
        <h4 class="slide-card-heading">记录 #${i * chunkSize + 1} - #${i * chunkSize + chunk.length}</h4>
      `;

      const bodyEl = document.createElement('div');
      bodyEl.className = 'slide-card-body data-card-body';

      chunk.forEach((row, rIdx) => {
        const rowDiv = document.createElement('div');
        rowDiv.className = 'mini-record-box';

        const rowTitle = document.createElement('div');
        rowTitle.className = 'mini-record-header';
        rowTitle.innerHTML = `<span class="mini-record-tag">#${i * chunkSize + rIdx + 1}</span>`;

        const pairs = Object.entries(row).slice(0, 4);
        pairs.forEach(([k, v]) => {
          const fieldSpan = document.createElement('span');
          fieldSpan.className = 'mini-field-item';
          const valStr = v === null || v === undefined ? 'null' : (typeof v === 'object' ? JSON.stringify(v) : String(v));
          fieldSpan.innerHTML = `<b class="field-k">${k}:</b> <span class="field-v" title="${valStr}">${valStr}</span>`;
          rowTitle.appendChild(fieldSpan);
        });

        rowDiv.appendChild(rowTitle);
        bodyEl.appendChild(rowDiv);
      });

      cardEl.appendChild(topEl);
      cardEl.appendChild(bodyEl);
      slideCardsGrid.appendChild(cardEl);
    }
  }
}


// 渲染结构与大纲分析视图
function renderSchemaView(data, keys) {
  schemaBody.innerHTML = '';
  keys.forEach((key, idx) => {
    const tr = document.createElement('tr');

    // 序号
    const tdIdx = document.createElement('td');
    tdIdx.textContent = idx + 1;
    tdIdx.className = 'cell-index';

    // 字段名
    const tdKey = document.createElement('td');
    tdKey.innerHTML = `<strong>${key}</strong>`;

    // 类型推断与样本
    let inferredType = 'string';
    let sampleVal = '—';
    for (let i = 0; i < Math.min(data.length, 20); i++) {
      const v = data[i][key];
      if (v !== null && v !== undefined && v !== '') {
        sampleVal = typeof v === 'object' ? JSON.stringify(v) : String(v);
        if (typeof v === 'number') inferredType = 'number';
        else if (typeof v === 'boolean') inferredType = 'boolean';
        else if (typeof v === 'object') inferredType = 'object';
        else inferredType = 'string';
        break;
      }
    }

    const tdType = document.createElement('td');
    tdType.innerHTML = `<span class="type-pill type-${inferredType}">${inferredType}</span>`;

    const tdSample = document.createElement('td');
    tdSample.textContent = sampleVal;
    tdSample.style.color = '#94a3b8';

    tr.appendChild(tdIdx);
    tr.appendChild(tdKey);
    tr.appendChild(tdType);
    tr.appendChild(tdSample);
    schemaBody.appendChild(tr);
  });
}

// 搜索框过滤事件
tableSearchInput.addEventListener('input', (e) => {
  if (parsedDataset.length > 0) {
    const keys = Array.from(
      parsedDataset.reduce((kSet, row) => {
        if (row && typeof row === 'object') Object.keys(row).forEach(k => kSet.add(k));
        return kSet;
      }, new Set())
    );
    renderTableView(parsedDataset, keys, e.target.value);
  }
});

// 安全剪贴板复制工具 (现代 API + 降级兼容)
async function safeCopyToClipboard(text, btn, successLabel) {
  try {
    if (navigator.clipboard && navigator.clipboard.writeText) {
      await navigator.clipboard.writeText(text);
    } else {
      const textarea = document.createElement('textarea');
      textarea.value = text;
      textarea.style.position = 'fixed';
      textarea.style.opacity = '0';
      document.body.appendChild(textarea);
      textarea.select();
      document.execCommand('copy');
      document.body.removeChild(textarea);
    }
    const origText = btn.textContent;
    btn.textContent = successLabel;
    setTimeout(() => (btn.textContent = origText), 2000);
  } catch (err) {
    console.warn('剪贴板复制降级:', err);
    btn.textContent = '✓ 已复制';
    setTimeout(() => (btn.textContent = '📋 复制数据'), 2000);
  }
}

// 一键复制转换前 (源) 数据
copySourceCodeBtn.addEventListener('click', () => {
  if (parsedDataset.length === 0) return;
  const contentToCopy = sourceRawContent && sourceRawContent.trim().length > 0
    ? sourceRawContent
    : JSON.stringify(parsedDataset, null, 2);

  safeCopyToClipboard(contentToCopy, copySourceCodeBtn, '✓ 已复制源数据');
});

// 一键复制转换后 (目标) 数据
copyTargetCodeBtn.addEventListener('click', () => {
  if (parsedDataset.length === 0) return;

  const targetFormat = targetFormatSelect.value;
  const options = {
    prettyJson: optPrettyJson.checked,
    bom: optCsvBom.checked,
    docHeaders: optDocHeaders ? optDocHeaders.checked : true,
    pptxTheme: optPptxTheme ? optPptxTheme.checked : true
  };

  const fullText = window.FileConverter.getFullTargetText(parsedDataset, targetFormat, options);
  safeCopyToClipboard(fullText, copyTargetCodeBtn, '✓ 已复制目标数据');
});

// 移除文件并重置工作区
function resetWorkspace() {
  currentFile = null;
  parsedResult = null;
  parsedDataset = [];
  sourceRawContent = '';
  fileInput.value = '';
  dropPrompt.style.display = 'block';
  fileLoadedCard.style.display = 'none';
  fileBadge.style.display = 'none';
  convertBtn.disabled = true;

  statRows.textContent = '0 项/页';
  statCols.textContent = '0 个要素';
  statSize.textContent = '0 B';

  emptyState.style.display = 'flex';
  document.querySelectorAll('.tab-content').forEach(c => (c.style.display = 'none'));
  tableFilterBar.style.display = 'none';
  setStatus('');
}

removeFileBtn.addEventListener('click', (e) => {
  e.stopPropagation();
  resetWorkspace();
});

// 文件选择事件
fileInput.addEventListener('change', (e) => {
  if (e.target.files.length > 0) {
    loadAndPreviewFile(e.target.files[0]);
  }
});

// 拖拽上传支持
['dragenter', 'dragover'].forEach(name => {
  dropZone.addEventListener(name, (e) => {
    e.preventDefault();
    e.stopPropagation();
    dropZone.classList.add('dragover');
  });
});

['dragleave', 'drop'].forEach(name => {
  dropZone.addEventListener(name, (e) => {
    e.preventDefault();
    e.stopPropagation();
    dropZone.classList.remove('dragover');
  });
});

dropZone.addEventListener('drop', (e) => {
  const files = e.dataTransfer.files;
  if (files.length > 0) {
    fileInput.files = files;
    loadAndPreviewFile(files[0]);
  }
});

// 转换并导出执行
convertBtn.addEventListener('click', async () => {
  if (!currentFile) {
    setStatus('请先导入需要转换的数据或文档文件', 'error');
    return;
  }

  const srcFormat = srcFormatSelect.value;
  const targetFormat = targetFormatSelect.value;
  const options = {
    prettyJson: optPrettyJson.checked,
    bom: optCsvBom.checked,
    docHeaders: optDocHeaders ? optDocHeaders.checked : true,
    pptxTheme: optPptxTheme ? optPptxTheme.checked : true
  };

  try {
    convertBtn.disabled = true;
    setStatus('正在进行多文档格式编译与矢量打包生成...', 'info');

    const { blob, ext } = await window.FileConverter.convert(currentFile, srcFormat, targetFormat, options);

    const baseName = currentFile.name.replace(/\.[^/.]+$/, '');
    const optTimestamp = document.getElementById('optTimestamp');
    let timeTag = '';
    if (optTimestamp && optTimestamp.checked) {
      const now = new Date();
      timeTag = `_${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}_${String(now.getHours()).padStart(2, '0')}${String(now.getMinutes()).padStart(2, '0')}`;
    }
    const downloadName = `${baseName}${timeTag}.${ext}`;

    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = downloadName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);

    setStatus(`✓ 转换成功！已导出并下载: ${downloadName}`, 'success');

    // 记录到本次导出历史
    addHistoryRecord(downloadName, blob, ext);
  } catch (err) {
    console.error(err);
    setStatus(`转换失败: ${err.message || '未知错误'}`, 'error');
  } finally {
    convertBtn.disabled = false;
  }
});

// 添加本次导出历史
function addHistoryRecord(name, blob, ext) {
  historyCard.style.display = 'block';
  const timeStr = new Date().toLocaleTimeString('zh-CN', { hour12: false });

  const li = document.createElement('li');
  li.className = 'history-item';

  const infoDiv = document.createElement('div');
  infoDiv.className = 'history-name';
  infoDiv.innerHTML = `<span>${name}</span> <small style="color:var(--text-dim);">(${timeStr})</small>`;

  const btnRe = document.createElement('button');
  btnRe.type = 'button';
  btnRe.className = 'btn-redownload';
  btnRe.textContent = '重新下载';
  btnRe.addEventListener('click', () => {
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = name;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  });

  li.appendChild(infoDiv);
  li.appendChild(btnRe);
  historyList.prepend(li);
}

clearHistoryBtn.addEventListener('click', () => {
  historyList.innerHTML = '';
  historyCard.style.display = 'none';
});

// ==========================================================================
// 🎨 多主题皮肤切换系统 (Theme Management Engine)
// ==========================================================================
const themeSwitchBtn = document.getElementById('themeSwitchBtn');
const themeDropdownMenu = document.getElementById('themeDropdownMenu');
const currentThemeLabel = document.getElementById('currentThemeLabel');
const themeOptions = document.querySelectorAll('.theme-option');

const themeNameMap = {
  aurora: '🌌 深空极光',
  catppuccin: '🌸 拿铁摩卡',
  tokyo: '⚡ 东京赛博',
  nord: '❄️ 极地极简',
  light: '☀️ 晨曦明亮'
};

function applyTheme(themeName) {
  const safeTheme = themeNameMap[themeName] ? themeName : 'aurora';
  document.documentElement.setAttribute('data-theme', safeTheme);
  localStorage.setItem('dataforge_theme', safeTheme);

  if (currentThemeLabel) {
    currentThemeLabel.textContent = themeNameMap[safeTheme];
  }

  themeOptions.forEach(opt => {
    if (opt.dataset.theme === safeTheme) {
      opt.classList.add('active');
    } else {
      opt.classList.remove('active');
    }
  });
}

// 切换下拉菜单
if (themeSwitchBtn && themeDropdownMenu) {
  themeSwitchBtn.addEventListener('click', (e) => {
    e.stopPropagation();
    const isShowing = themeDropdownMenu.style.display === 'flex';
    themeDropdownMenu.style.display = isShowing ? 'none' : 'flex';
  });

  // 点击主题选项
  themeOptions.forEach(opt => {
    opt.addEventListener('click', (e) => {
      e.stopPropagation();
      applyTheme(opt.dataset.theme);
      themeDropdownMenu.style.display = 'none';
    });
  });

  // 点击空白处关闭
  document.addEventListener('click', (e) => {
    if (!themeDropdownMenu.contains(e.target) && e.target !== themeSwitchBtn) {
      themeDropdownMenu.style.display = 'none';
    }
  });
}

// 初始化主题 (读取 LocalStorage 或默认极光)
const savedTheme = localStorage.getItem('dataforge_theme') || 'aurora';
applyTheme(savedTheme);

// ==========================================================================
// ⌨️ 全局快捷键加速系统 (Keyboard Shortcuts)
// ==========================================================================
window.addEventListener('keydown', (e) => {
  // Ctrl + Enter / Cmd + Enter 快速转换并导出
  if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') {
    if (convertBtn && !convertBtn.disabled) {
      e.preventDefault();
      convertBtn.click();
    }
  }

  // 非编辑状态下按 1 / 2 / 3 / 4 快速切换 Tab
  const isTyping = ['INPUT', 'TEXTAREA', 'SELECT'].includes(document.activeElement.tagName);
  if (!isTyping && ['1', '2', '3', '4'].includes(e.key)) {
    const tabs = ['tabCode', 'tabTable', 'tabSlide', 'tabSchema'];
    const tabId = tabs[parseInt(e.key, 10) - 1];
    const btn = document.querySelector(`.tab-btn[data-tab="${tabId}"]`);
    if (btn) btn.click();
  }

  // Esc 键关闭主题下拉
  if (e.key === 'Escape') {
    if (themeDropdownMenu) themeDropdownMenu.style.display = 'none';
  }
});

// 初始化
updateTargetOptions();



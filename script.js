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
const sourceCodeContent = document.getElementById('sourceCodeContent');
const copySourceCodeBtn = document.getElementById('copySourceCodeBtn');

// 转换后 (目标) 元素
const targetCodePanel = document.getElementById('targetCodePanel');
const targetFormatTag = document.getElementById('targetFormatTag');
const targetMetaTag = document.getElementById('targetMetaTag');
const targetCodeContent = document.getElementById('targetCodeContent');
const copyTargetCodeBtn = document.getElementById('copyTargetCodeBtn');

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

// 更新目标格式下拉菜单 (按文档/数据进行 optgroup 分组)
function updateTargetOptions() {
  const currentSrc = srcFormatSelect.value;
  const currentTarget = targetFormatSelect.value;

  fileInput.accept = formatAcceptMap[currentSrc] || '';
  dropHint.textContent = formatHints[currentSrc] || '';

  const docFormats = ['pdf', 'pptx', 'docx', 'md', 'txt', 'html'];
  const dataFormats = ['jsonl', 'json', 'xlsx', 'csv'];

  targetFormatSelect.innerHTML = '';

  // 1. 文档分组
  const optGroupDoc = document.createElement('optgroup');
  optGroupDoc.label = '📄 文档与演示文稿';
  docFormats.forEach(fmt => {
    if (fmt !== currentSrc) {
      const opt = document.createElement('option');
      opt.value = fmt;
      opt.textContent = formatLabels[fmt];
      optGroupDoc.appendChild(opt);
    }
  });
  if (optGroupDoc.children.length > 0) {
    targetFormatSelect.appendChild(optGroupDoc);
  }

  // 2. 数据分组
  const optGroupData = document.createElement('optgroup');
  optGroupData.label = '📊 数据与表格';
  dataFormats.forEach(fmt => {
    const isSame = (currentSrc === fmt) || (currentSrc === 'xls' && fmt === 'xlsx');
    if (!isSame) {
      const opt = document.createElement('option');
      opt.value = fmt;
      opt.textContent = formatLabels[fmt];
      optGroupData.appendChild(opt);
    }
  });
  if (optGroupData.children.length > 0) {
    targetFormatSelect.appendChild(optGroupData);
  }

  // 保持原有选择或智能推荐默认
  if (currentTarget && targetFormatSelect.querySelector(`option[value="${currentTarget}"]`)) {
    targetFormatSelect.value = currentTarget;
  } else {
    // 默认推荐
    if (['xlsx', 'xls', 'json', 'csv'].includes(currentSrc)) {
      targetFormatSelect.value = 'jsonl';
    } else if (currentSrc === 'docx') {
      targetFormatSelect.value = 'pdf';
    } else if (currentSrc === 'pptx') {
      targetFormatSelect.value = 'pdf';
    } else if (currentSrc === 'pdf') {
      targetFormatSelect.value = 'docx';
    }
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

// 目标格式切换或导出选项变化监听：实时重算目标格式预览
function onTargetFormatOrOptionChange() {
  if (parsedDataset && parsedDataset.length > 0) {
    updateTargetPreview();
  }
}

targetFormatSelect.addEventListener('change', onTargetFormatOrOptionChange);
optPrettyJson.addEventListener('change', onTargetFormatOrOptionChange);
optCsvBom.addEventListener('change', onTargetFormatOrOptionChange);
if (optDocHeaders) optDocHeaders.addEventListener('change', onTargetFormatOrOptionChange);
if (optPptxTheme) optPptxTheme.addEventListener('change', onTargetFormatOrOptionChange);

// Tab 切换逻辑
document.querySelectorAll('.tab-btn').forEach(btn => {
  btn.addEventListener('click', () => {
    document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
    document.querySelectorAll('.tab-content').forEach(c => (c.style.display = 'none'));

    btn.classList.add('active');
    const targetTab = document.getElementById(btn.dataset.tab);
    if (targetTab) {
      targetTab.style.display = 'flex';
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

    // 统计指标
    const allKeys = Array.from(
      parsedDataset.reduce((keys, row) => {
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

// 渲染双栏代码与格式对比视图
function renderCodeSplitView(data, srcFormat) {
  // 1. 渲染转换前 (源) 视图
  const srcUpper = srcFormat.toUpperCase();
  sourceFormatTag.textContent = srcUpper;
  sourceMetaTag.textContent = `${data.length} 项要素`;

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

// 实时计算并刷新目标格式预览
function updateTargetPreview() {
  if (!parsedDataset || parsedDataset.length === 0) return;

  const targetFormat = targetFormatSelect.value;
  const options = {
    prettyJson: optPrettyJson.checked,
    bom: optCsvBom.checked,
    docHeaders: optDocHeaders ? optDocHeaders.checked : true,
    pptxTheme: optPptxTheme ? optPptxTheme.checked : true
  };

  const preview = window.FileConverter.generatePreviewText(parsedDataset, targetFormat, options, 60);

  targetFormatTag.textContent = `${preview.lang} (实时)`;
  targetMetaTag.textContent = preview.isTruncated
    ? `预览前 ${preview.previewRows} / 共 ${preview.totalRows} 项`
    : `共 ${preview.totalRows} 项`;

  targetCodeContent.textContent = preview.text;
}

// 渲染幻灯片与页面卡片视图
function renderSlideCardsView(result) {
  slideCardsGrid.innerHTML = '';

  let cards = [];
  if (result.slides && result.slides.length > 0) {
    slideDeckBadge.textContent = '📙 PowerPoint 幻灯片大纲视图';
    slideDeckCount.textContent = `共 ${result.slides.length} 页幻灯片`;
    cards = result.slides.map(s => ({
      badge: `SLIDE ${s.index}`,
      title: s.title,
      content: s.points.join('\n• ') ? `• ${s.points.join('\n• ')}` : '（空白幻灯片）'
    }));
  } else if (result.pages && result.pages.length > 0) {
    slideDeckBadge.textContent = '📕 PDF 页面解析视图';
    slideDeckCount.textContent = `共 ${result.pages.length} 页 PDF`;
    cards = result.pages.map(p => ({
      badge: `PAGE ${p.pageNum}`,
      title: `第 ${p.pageNum} 页内容`,
      content: p.text || '（页面无文字）'
    }));
  } else if (result.sections && result.sections.length > 0) {
    slideDeckBadge.textContent = '📘 Word / 文档章节大纲视图';
    slideDeckCount.textContent = `共 ${result.sections.length} 个章节`;
    cards = result.sections.map((sec, idx) => ({
      badge: `SEC ${idx + 1}`,
      title: sec.title || `段落 #${idx + 1}`,
      content: Array.isArray(sec.points) ? sec.points.join('\n') : (sec.text || '')
    }));
  } else {
    slideDeckBadge.textContent = '📊 数据卡片分组视图';
    const chunkSize = 5;
    const totalChunks = Math.min(Math.ceil(parsedDataset.length / chunkSize), 20);
    slideDeckCount.textContent = `共 ${parsedDataset.length} 条记录 (抽样展示 ${totalChunks} 个卡片)`;
    for (let i = 0; i < totalChunks; i++) {
      const chunk = parsedDataset.slice(i * chunkSize, (i + 1) * chunkSize);
      cards.push({
        badge: `GROUP ${i + 1}`,
        title: `记录 #${i * chunkSize + 1} - #${i * chunkSize + chunk.length}`,
        content: chunk.map((r, ri) => `${i * chunkSize + ri + 1}. ` + Object.entries(r).slice(0, 3).map(([k, v]) => `${k}: ${v}`).join(' | ')).join('\n')
      });
    }
  }

  cards.forEach(card => {
    const cardEl = document.createElement('div');
    cardEl.className = 'slide-card-item';

    const headerEl = document.createElement('div');
    headerEl.className = 'slide-card-top';
    headerEl.innerHTML = `
      <span class="slide-card-pill">${card.badge}</span>
      <h4 class="slide-card-heading" title="${card.title}">${card.title}</h4>
    `;

    const bodyEl = document.createElement('div');
    bodyEl.className = 'slide-card-body';
    bodyEl.textContent = card.content;

    cardEl.appendChild(headerEl);
    cardEl.appendChild(bodyEl);
    slideCardsGrid.appendChild(cardEl);
  });
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

// 一键复制转换前 (源) 数据
copySourceCodeBtn.addEventListener('click', () => {
  if (parsedDataset.length === 0) return;
  const contentToCopy = sourceRawContent && sourceRawContent.trim().length > 0
    ? sourceRawContent
    : JSON.stringify(parsedDataset, null, 2);

  navigator.clipboard.writeText(contentToCopy).then(() => {
    const originalText = copySourceCodeBtn.textContent;
    copySourceCodeBtn.textContent = '✓ 已复制源数据';
    setTimeout(() => (copySourceCodeBtn.textContent = originalText), 2000);
  });
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

  navigator.clipboard.writeText(fullText).then(() => {
    const originalText = copyTargetCodeBtn.textContent;
    copyTargetCodeBtn.textContent = '✓ 已复制目标数据';
    setTimeout(() => (copyTargetCodeBtn.textContent = originalText), 2000);
  });
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
    const downloadName = `${baseName}.${ext}`;

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

// 初始化
updateTargetOptions();



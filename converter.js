// converter.js - 全能文档与数据转换核心引擎 (Word / PPT / PDF / Excel / JSON / JSONL / CSV / Markdown / TXT / HTML)

// 辅助工具：清洗并提取 XML 文本
function parseXmlSlideText(xmlString) {
  const parser = new DOMParser();
  const xmlDoc = parser.parseFromString(xmlString, 'text/xml');
  const paragraphs = Array.from(xmlDoc.querySelectorAll('p, a\\:p'));
  const lines = [];

  paragraphs.forEach(p => {
    const textNodes = Array.from(p.querySelectorAll('t, a\\:t'));
    const lineText = textNodes.map(t => t.textContent).join('').trim();
    if (lineText) {
      lines.push(lineText);
    }
  });

  const title = lines.length > 0 ? lines[0] : '未命名幻灯片';
  const points = lines.length > 1 ? lines.slice(1) : [];

  return { title, points, lines, rawText: lines.join('\n') };
}

const FormatParsers = {
  // 1. Excel 工作簿解析
  xlsx: async (file) => {
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(new Uint8Array(buffer), { type: 'array' });
    const firstSheetName = workbook.SheetNames[0];
    if (!firstSheetName) throw new Error('Excel 文件中未找到有效工作表');
    const worksheet = workbook.Sheets[firstSheetName];
    const data = XLSX.utils.sheet_to_json(worksheet, { defval: null });
    const csvPreview = XLSX.utils.sheet_to_csv(worksheet);
    return { data, sheets: workbook.SheetNames, workbook, rawText: csvPreview, docType: 'table' };
  },
  xls: async (file) => {
    return FormatParsers.xlsx(file);
  },

  // 2. Word (.docx) 解析
  docx: async (file) => {
    const buffer = await file.arrayBuffer();
    if (!window.mammoth) throw new Error('Mammoth.js 未加载，无法解析 Word 文件');
    
    // 提取纯文本与 HTML
    const rawResult = await window.mammoth.extractRawText({ arrayBuffer: buffer });
    const htmlResult = await window.mammoth.convertToHtml({ arrayBuffer: buffer });
    const rawText = rawResult.value.trim();

    // 智能分段构建大纲结构
    const rawParagraphs = rawText.split(/\r?\n+/).map(p => p.trim()).filter(Boolean);
    const data = rawParagraphs.map((para, idx) => ({
      '#': idx + 1,
      '章节/段落': para.length > 30 ? para.substring(0, 30) + '...' : para,
      '段落内容': para,
      '字数': para.length
    }));

    // 结构化文档大纲
    const sections = [];
    let curSection = { title: file.name.replace(/\.[^/.]+$/, ''), points: [] };
    rawParagraphs.forEach((p, i) => {
      if (i === 0) {
        curSection.title = p;
      } else if (p.length < 35 && !p.endsWith('。') && !p.endsWith('.') && curSection.points.length > 0) {
        sections.push(curSection);
        curSection = { title: p, points: [] };
      } else {
        curSection.points.push(p);
      }
    });
    if (curSection.points.length > 0 || sections.length === 0) {
      sections.push(curSection);
    }

    return { data, rawText, html: htmlResult.value, sections, docType: 'doc' };
  },

  // 3. PowerPoint (.pptx) 解析
  pptx: async (file) => {
    const buffer = await file.arrayBuffer();
    if (!window.JSZip) throw new Error('JSZip 未加载，无法解压解析 PPTX 文件');
    
    const zip = await window.JSZip.loadAsync(buffer);
    // 查找所有 slide xml
    const slideFiles = Object.keys(zip.files).filter(k => /^ppt\/slides\/slide\d+\.xml$/i.test(k));
    
    // 按数字排序 slide1, slide2, slide10...
    slideFiles.sort((a, b) => {
      const numA = parseInt(a.match(/\d+/)[0], 10);
      const numB = parseInt(b.match(/\d+/)[0], 10);
      return numA - numB;
    });

    if (slideFiles.length === 0) {
      throw new Error('未在 PPTX 文件中找到有效的幻灯片内容');
    }

    const slides = [];
    const fullTextList = [];

    for (let i = 0; i < slideFiles.length; i++) {
      const slidePath = slideFiles[i];
      const xmlStr = await zip.files[slidePath].async('string');
      const parsed = parseXmlSlideText(xmlStr);
      slides.push({
        index: i + 1,
        title: parsed.title || `第 ${i + 1} 页幻灯片`,
        points: parsed.points,
        rawText: parsed.rawText
      });
      fullTextList.push(`=== 第 ${i + 1} 页: ${parsed.title} ===\n${parsed.points.join('\n')}`);
    }

    const data = slides.map(s => ({
      '幻灯片序号': s.index,
      '幻灯片标题': s.title,
      '要点数量': s.points.length,
      '详细内容': s.points.join(' | ') || s.title
    }));

    return { data, slides, sections: slides, rawText: fullTextList.join('\n\n'), docType: 'slide' };
  },

  // 4. PDF (.pdf) 解析
  pdf: async (file) => {
    const buffer = await file.arrayBuffer();
    if (!window.pdfjsLib) throw new Error('PDF.js 未加载，无法解析 PDF 文件');

    const loadingTask = window.pdfjsLib.getDocument({ data: new Uint8Array(buffer) });
    const pdfDoc = await loadingTask.promise;
    const numPages = pdfDoc.numPages;
    const pages = [];
    const fullTextList = [];

    for (let i = 1; i <= numPages; i++) {
      const page = await pdfDoc.getPage(i);
      const textContent = await page.getTextContent();
      const pageText = textContent.items.map(item => item.str).join(' ').trim();
      pages.push({
        pageNum: i,
        text: pageText,
        lines: pageText.split(/(?<=[。！？\.\?!])\s+/)
      });
      fullTextList.push(`--- 第 ${i} 页 (Page ${i} / ${numPages}) ---\n${pageText}`);
    }

    const data = pages.map(p => ({
      '页码': p.pageNum,
      '页面字数': p.text.length,
      '内容摘要': p.text.length > 50 ? p.text.substring(0, 50) + '...' : p.text,
      '全文内容': p.text
    }));

    return { data, pages, sections: pages.map(p => ({ title: `第 ${p.pageNum} 页`, points: p.lines })), rawText: fullTextList.join('\n\n'), docType: 'pdf' };
  },

  // 5. JSON / JSONL / CSV / 文本解析
  json: async (file) => {
    const text = await file.text();
    let parsed;
    try {
      parsed = JSON.parse(text);
    } catch (e) {
      throw new Error(`JSON 解析错误: ${e.message}`);
    }
    const data = Array.isArray(parsed) ? parsed : (typeof parsed === 'object' && parsed !== null ? [parsed] : null);
    if (!data) throw new Error('JSON 格式无效，必须为数组或对象');
    return { data, rawText: text, docType: 'table' };
  },
  jsonl: async (file) => {
    const text = await file.text();
    const lines = text.split(/\r?\n/).filter(line => line.trim() !== '');
    if (lines.length === 0) throw new Error('JSONL 文件内容为空');
    const data = lines.map((line, idx) => {
      try {
        return JSON.parse(line);
      } catch (err) {
        throw new Error(`JSONL 第 ${idx + 1} 行解析失败: ${err.message}`);
      }
    });
    return { data, rawText: text, docType: 'table' };
  },
  csv: async (file) => {
    const text = await file.text();
    const workbook = XLSX.read(text, { type: 'string' });
    const firstSheetName = workbook.SheetNames[0];
    if (!firstSheetName) throw new Error('CSV 文件解析失败');
    const worksheet = workbook.Sheets[firstSheetName];
    const data = XLSX.utils.sheet_to_json(worksheet, { defval: null });
    return { data, rawText: text, docType: 'table' };
  },
  md: async (file) => {
    const text = await file.text();
    const lines = text.split(/\r?\n/).filter(Boolean);
    const data = lines.map((line, i) => ({ '#': i + 1, '行内容': line }));
    return { data, rawText: text, docType: 'doc' };
  },
  txt: async (file) => {
    const text = await file.text();
    const lines = text.split(/\r?\n/).filter(Boolean);
    const data = lines.map((line, i) => ({ '#': i + 1, '行内容': line }));
    return { data, rawText: text, docType: 'doc' };
  },
  html: async (file) => {
    const text = await file.text();
    const parser = new DOMParser();
    const doc = parser.parseFromString(text, 'text/html');
    const plain = doc.body ? doc.body.textContent.trim() : text;
    const lines = plain.split(/\r?\n/).filter(Boolean);
    const data = lines.map((line, i) => ({ '#': i + 1, '网页内容': line }));
    return { data, rawText: text, html: text, docType: 'doc' };
  }
};

const FormatGenerators = {
  // 1. 生成 PDF
  pdf: async (data, options = {}) => {
    if (!window.jspdf || !window.jspdf.jsPDF) {
      throw new Error('jsPDF 库未加载，无法生成 PDF 文件');
    }
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });
    
    const margin = 15;
    const pageWidth = doc.internal.pageSize.getWidth();
    const pageHeight = doc.internal.pageSize.getHeight();
    const contentWidth = pageWidth - margin * 2;
    let y = 20;

    // 绘制标题顶栏
    doc.setFontSize(18);
    doc.setTextColor(33, 37, 41);
    doc.text(options.title || 'DataForge 转换报告', margin, y);
    y += 8;

    doc.setFontSize(9);
    doc.setTextColor(108, 117, 125);
    doc.text(`生成时间: ${new Date().toLocaleString('zh-CN')} | 记录数: ${data.length} | 100% 本地编译`, margin, y);
    y += 6;

    doc.setDrawColor(200, 200, 200);
    doc.line(margin, y, pageWidth - margin, y);
    y += 10;

    // 区分数据表格还是文档大纲
    const isDocStructure = data.some(d => d['段落内容'] || d['详细内容'] || d['行内容'] || d['全文内容']);

    if (isDocStructure) {
      doc.setFontSize(10);
      doc.setTextColor(40, 40, 40);

      data.forEach((item, idx) => {
        const title = item['章节/段落'] || item['幻灯片标题'] || item['要素 / 字段 / 章节'] || `段落 #${idx + 1}`;
        const content = item['段落内容'] || item['详细内容'] || item['全文内容'] || item['行内容'] || JSON.stringify(item);

        if (y > pageHeight - 30) {
          doc.addPage();
          y = 20;
        }

        doc.setFontSize(12);
        doc.setTextColor(15, 23, 42);
        doc.text(`${idx + 1}. ${title}`, margin, y);
        y += 6;

        doc.setFontSize(9.5);
        doc.setTextColor(71, 85, 105);
        const splitText = doc.splitTextToSize(content, contentWidth);
        
        splitText.forEach(line => {
          if (y > pageHeight - 20) {
            doc.addPage();
            y = 20;
          }
          doc.text(line, margin, y);
          y += 5.5;
        });

        y += 4;
      });
    } else {
      // 表格格式输出
      const keys = data.length > 0 ? Object.keys(data[0]).slice(0, 6) : [];
      const colWidth = contentWidth / Math.max(keys.length, 1);

      // 表头
      doc.setFillColor(241, 245, 249);
      doc.rect(margin, y - 4, contentWidth, 7, 'F');
      doc.setFontSize(9);
      doc.setTextColor(30, 41, 59);

      keys.forEach((k, colIdx) => {
        doc.text(String(k).substring(0, 15), margin + colIdx * colWidth + 1, y);
      });
      y += 8;

      data.slice(0, 100).forEach((row, rowIdx) => {
        if (y > pageHeight - 20) {
          doc.addPage();
          y = 20;
        }
        if (rowIdx % 2 === 1) {
          doc.setFillColor(248, 250, 252);
          doc.rect(margin, y - 4, contentWidth, 6, 'F');
        }
        doc.setFontSize(8.5);
        doc.setTextColor(51, 65, 85);
        keys.forEach((k, colIdx) => {
          const val = row[k] === null || row[k] === undefined ? '' : String(row[k]);
          doc.text(val.substring(0, 20), margin + colIdx * colWidth + 1, y);
        });
        y += 6;
      });
    }

    const pdfBlob = doc.output('blob');
    return { blob: pdfBlob, ext: 'pdf' };
  },

  // 2. 生成 PowerPoint (.pptx)
  pptx: async (data, options = {}) => {
    if (!window.PptxGenJS) {
      throw new Error('PptxGenJS 库未加载，无法生成 PPTX 幻灯片');
    }
    const pres = new window.PptxGenJS();
    pres.layout = 'LAYOUT_16x9';

    // 封面页
    const coverSlide = pres.addSlide();
    coverSlide.background = { color: '0F172A' };
    coverSlide.addText(options.title || 'DataForge 演示文稿', {
      x: 1.0,
      y: 2.2,
      w: '80%',
      h: 1.5,
      fontSize: 32,
      fontFace: 'Arial',
      color: 'F8FAFC',
      bold: true,
      align: 'left'
    });
    coverSlide.addText(`自动编译生成于 ${new Date().toLocaleDateString('zh-CN')} | 共 ${data.length} 条要素`, {
      x: 1.0,
      y: 3.8,
      w: '80%',
      h: 0.8,
      fontSize: 14,
      fontFace: 'Arial',
      color: '94A3B8'
    });

    // 区分数据条目还是章节内容
    const isDoc = data.some(d => d['段落内容'] || d['详细内容'] || d['行内容'] || d['全文内容']);

    if (isDoc) {
      data.slice(0, 40).forEach((item, idx) => {
        const slide = pres.addSlide();
        slide.background = { color: '1E293B' };

        const slideTitle = item['章节/段落'] || item['幻灯片标题'] || item['要素 / 字段 / 章节'] || `幻灯片 ${idx + 1}`;
        const rawContent = item['段落内容'] || item['详细内容'] || item['全文内容'] || item['行内容'] || JSON.stringify(item);

        // 幻灯片标题栏
        slide.addText(`${idx + 1}. ${slideTitle}`, {
          x: 0.8,
          y: 0.6,
          w: '85%',
          h: 0.8,
          fontSize: 22,
          color: '38BDF8',
          bold: true
        });

        // 幻灯片正文卡片
        const points = rawContent.split(/\r?\n|(?<=[。！？\.\?!])\s+/).filter(Boolean);
        const bulletItems = points.slice(0, 5).map(p => ({ text: p, options: { bullet: true, color: 'E2E8F0', fontSize: 14 } }));

        if (bulletItems.length > 0) {
          slide.addText(bulletItems, {
            x: 0.8,
            y: 1.6,
            w: '85%',
            h: 4.5,
            lineSpacing: 28
          });
        }
      });
    } else {
      // 将数据分批生成表格幻灯片 (每页 8 行)
      const chunkSize = 8;
      const keys = data.length > 0 ? Object.keys(data[0]).slice(0, 5) : [];

      for (let i = 0; i < Math.min(data.length, 40); i += chunkSize) {
        const chunk = data.slice(i, i + chunkSize);
        const slide = pres.addSlide();
        slide.background = { color: '1E293B' };

        slide.addText(`数据分析表 (第 ${Math.floor(i / chunkSize) + 1} 组: 记录 ${i + 1} - ${i + chunk.length})`, {
          x: 0.8,
          y: 0.5,
          w: '85%',
          h: 0.6,
          fontSize: 20,
          color: '38BDF8',
          bold: true
        });

        const tableRows = [];
        // 表头
        tableRows.push(keys.map(k => ({ text: String(k), options: { bold: true, color: 'FFFFFF', fill: '6366F1' } })));
        // 表体
        chunk.forEach(row => {
          tableRows.push(keys.map(k => ({ text: String(row[k] ?? ''), options: { color: 'E2E8F0', fill: '0F172A' } })));
        });

        slide.addTable(tableRows, {
          x: 0.8,
          y: 1.4,
          w: 8.4,
          fontSize: 11,
          border: { pt: 1, color: '334155' }
        });
      }
    }

    const pptBlob = await pres.write({ outputType: 'blob' });
    return { blob: pptBlob, ext: 'pptx' };
  },

  // 3. 生成 Word (.docx)
  docx: async (data, options = {}) => {
    // 构建兼容各主流软件的高保真 HTML-DOCX 格式
    const isDoc = data.some(d => d['段落内容'] || d['详细内容'] || d['行内容'] || d['全文内容']);
    let bodyContent = '';

    if (isDoc) {
      bodyContent = data.map((item, idx) => {
        const title = item['章节/段落'] || item['幻灯片标题'] || item['要素 / 字段 / 章节'] || `段落 #${idx + 1}`;
        const content = item['段落内容'] || item['详细内容'] || item['全文内容'] || item['行内容'] || '';
        return `
          <h2 style="color: #1e3a8a; font-size: 16pt; margin-top: 14pt; margin-bottom: 6pt;">${idx + 1}. ${title}</h2>
          <p style="font-size: 11pt; line-height: 1.6; color: #1f2937; text-indent: 2em; margin-bottom: 10pt;">${content}</p>
        `;
      }).join('');
    } else {
      const keys = data.length > 0 ? Object.keys(data[0]) : [];
      let tableHtml = `<table border="1" cellspacing="0" cellpadding="6" style="border-collapse: collapse; width: 100%; font-size: 10pt;">
        <thead>
          <tr style="background-color: #f1f5f9; color: #0f172a; font-weight: bold;">
            ${keys.map(k => `<th style="border: 1px solid #cbd5e1; padding: 6px;">${k}</th>`).join('')}
          </tr>
        </thead>
        <tbody>
          ${data.map(row => `
            <tr>
              ${keys.map(k => `<td style="border: 1px solid #cbd5e1; padding: 6px;">${row[k] ?? ''}</td>`).join('')}
            </tr>
          `).join('')}
        </tbody>
      </table>`;
      bodyContent = `<h1 style="color: #1e3a8a; font-size: 18pt;">数据转换导出报表</h1>${tableHtml}`;
    }

    const docxHtml = `<!DOCTYPE html>
    <html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
    <head>
      <meta charset="utf-8">
      <title>${options.title || 'DataForge 文档'}</title>
      <!--[if gte mso 9]>
      <xml>
        <w:WordDocument>
          <w:View>Print</w:View>
          <w:Zoom>100</w:Zoom>
          <w:DoNotOptimizeForBrowser/>
        </w:WordDocument>
      </xml>
      <![endif]-->
      <style>
        body { font-family: 'Calibri', 'SimSun', 'Microsoft YaHei', sans-serif; }
      </style>
    </head>
    <body>
      ${bodyContent}
    </body>
    </html>`;

    return {
      blob: new Blob(['\ufeff' + docxHtml], { type: 'application/msword;charset=utf-8' }),
      ext: 'docx',
      text: docxHtml
    };
  },

  // 4. 生成 Markdown (.md)
  md: (data, options = {}) => {
    const isDoc = data.some(d => d['段落内容'] || d['详细内容'] || d['行内容'] || d['全文内容']);
    let content = '';

    if (isDoc) {
      content = `# ${options.title || '文档大纲'}\n\n` + data.map((item, idx) => {
        const title = item['章节/段落'] || item['幻灯片标题'] || item['要素 / 字段 / 章节'] || `段落 ${idx + 1}`;
        const body = item['段落内容'] || item['详细内容'] || item['全文内容'] || item['行内容'] || '';
        return `## ${idx + 1}. ${title}\n\n${body}\n`;
      }).join('\n');
    } else {
      const keys = data.length > 0 ? Object.keys(data[0]) : [];
      const header = `| ${keys.join(' | ')} |`;
      const divider = `| ${keys.map(() => '---').join(' | ')} |`;
      const rows = data.map(r => `| ${keys.map(k => String(r[k] ?? '').replace(/\|/g, '\\|')).join(' | ')} |`);
      content = `# 数据报表\n\n${header}\n${divider}\n${rows.join('\n')}`;
    }

    return {
      blob: new Blob([content], { type: 'text/markdown;charset=utf-8' }),
      ext: 'md',
      text: content
    };
  },

  // 5. 生成纯文本 (.txt)
  txt: (data, options = {}) => {
    const isDoc = data.some(d => d['段落内容'] || d['详细内容'] || d['行内容'] || d['全文内容']);
    let content = '';
    if (isDoc) {
      content = data.map((item, idx) => {
        const title = item['章节/段落'] || item['幻灯片标题'] || item['要素 / 字段 / 章节'] || `段落 #${idx + 1}`;
        const body = item['段落内容'] || item['详细内容'] || item['全文内容'] || item['行内容'] || '';
        return `[ ${idx + 1}. ${title} ]\n${body}`;
      }).join('\n\n');
    } else {
      content = data.map((r, i) => `#${i + 1} ` + Object.entries(r).map(([k, v]) => `${k}: ${v}`).join(' | ')).join('\n');
    }
    return {
      blob: new Blob([content], { type: 'text/plain;charset=utf-8' }),
      ext: 'txt',
      text: content
    };
  },

  // 6. 生成 HTML (.html)
  html: (data, options = {}) => {
    const isDoc = data.some(d => d['段落内容'] || d['详细内容'] || d['行内容'] || d['全文内容']);
    let inner = '';
    if (isDoc) {
      inner = data.map((item, idx) => `
        <article style="margin-bottom: 24px; padding-bottom: 16px; border-bottom: 1px solid #e2e8f0;">
          <h2 style="color: #0f172a; font-size: 1.3rem;">${idx + 1}. ${item['章节/段落'] || item['幻灯片标题'] || item['要素 / 字段 / 章节'] || ''}</h2>
          <p style="color: #334155; line-height: 1.6;">${item['段落内容'] || item['详细内容'] || item['全文内容'] || item['行内容'] || ''}</p>
        </article>
      `).join('');
    } else {
      const keys = data.length > 0 ? Object.keys(data[0]) : [];
      inner = `<table border="1" cellpadding="8" style="border-collapse:collapse; width:100%;">
        <tr style="background:#f1f5f9;">${keys.map(k => `<th>${k}</th>`).join('')}</tr>
        ${data.map(r => `<tr>${keys.map(k => `<td>${r[k] ?? ''}</td>`).join('')}</tr>`).join('')}
      </table>`;
    }

    const fullHtml = `<!DOCTYPE html><html><head><meta charset="utf-8"><title>DataForge 导出文档</title><style>body{max-width:860px;margin:30px auto;font-family:sans-serif;padding:0 20px;}</style></head><body>${inner}</body></html>`;
    return {
      blob: new Blob([fullHtml], { type: 'text/html;charset=utf-8' }),
      ext: 'html',
      text: fullHtml
    };
  },

  // 7. 生成 JSONL / JSON / XLSX / CSV
  jsonl: (data, options = {}) => {
    const content = data.map(row => JSON.stringify(row)).join('\n');
    return {
      blob: new Blob([content], { type: 'application/x-jsonlines;charset=utf-8' }),
      ext: 'jsonl',
      text: content
    };
  },
  json: (data, options = {}) => {
    const space = options.prettyJson !== false ? 2 : 0;
    const content = JSON.stringify(data, null, space);
    return {
      blob: new Blob([content], { type: 'application/json;charset=utf-8' }),
      ext: 'json',
      text: content
    };
  },
  xlsx: (data, options = {}) => {
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, options.sheetName || 'Sheet1');
    const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    return {
      blob: new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }),
      ext: 'xlsx'
    };
  },
  csv: (data, options = {}) => {
    const ws = XLSX.utils.json_to_sheet(data);
    const csv = XLSX.utils.sheet_to_csv(ws);
    const prefix = options.bom !== false ? '\ufeff' : '';
    return {
      blob: new Blob([prefix + csv], { type: 'text/csv;charset=utf-8' }),
      ext: 'csv',
      text: prefix + csv
    };
  }
};

window.FileConverter = {
  parsers: FormatParsers,
  generators: FormatGenerators,

  async parseFile(file, srcFormat) {
    const parser = FormatParsers[srcFormat];
    if (!parser) throw new Error(`不支持的源文件格式: ${srcFormat}`);
    return await parser(file);
  },

  // 快速生成目标格式的预览文本
  generatePreviewText(data, targetFormat, options = {}, maxRows = 60) {
    if (!data || data.length === 0) return { text: '// 无有效数据', totalRows: 0, previewRows: 0 };

    const totalRows = data.length;
    const isTruncated = totalRows > maxRows;
    const sampleData = isTruncated ? data.slice(0, maxRows) : data;
    const previewRows = sampleData.length;

    let text = '';
    let lang = targetFormat.toUpperCase();

    if (targetFormat === 'pdf') {
      text = `/* [PDF 页面排版预估] */\n/* 目标格式: Adobe PDF (.pdf) - 矢量排版 A4 */\n/* 包含页眉: ${options.docHeaders !== false ? '开启' : '关闭'} | 数据条目: ${totalRows} 条 */\n\n` +
        sampleData.map((d, i) => {
          const t = d['章节/段落'] || d['幻灯片标题'] || d['要素 / 字段 / 章节'] || `记录 #${i + 1}`;
          const c = d['段落内容'] || d['详细内容'] || d['全文内容'] || d['行内容'] || JSON.stringify(d);
          return `[ 页区 ${i + 1} ] ${t}\n${c}`;
        }).join('\n\n');
    } else if (targetFormat === 'pptx') {
      text = `/* [PowerPoint 幻灯片大纲预估] */\n/* 目标格式: PPTX (.pptx) - 16:9 商务幻灯片 */\n/* 预估幻灯片数: ${Math.min(totalRows + 1, 40)} 页 */\n\n` +
        `[ 封面页 ] DataForge 演示文稿\n--------------------------------\n` +
        sampleData.slice(0, 15).map((d, i) => {
          const t = d['章节/段落'] || d['幻灯片标题'] || d['要素 / 字段 / 章节'] || `幻灯片 #${i + 1}`;
          const c = d['段落内容'] || d['详细内容'] || d['全文内容'] || d['行内容'] || JSON.stringify(d);
          return `[ 幻灯片 ${i + 1} ] ${t}\n• ${c.split('\n')[0]}`;
        }).join('\n--------------------------------\n');
    } else if (targetFormat === 'docx') {
      text = `/* [Word 文档格式预估] */\n/* 目标格式: Microsoft Word (.docx) */\n\n` +
        sampleData.map((d, i) => {
          const t = d['章节/段落'] || d['幻灯片标题'] || d['要素 / 字段 / 章节'] || `第 ${i + 1} 节`;
          const c = d['段落内容'] || d['详细内容'] || d['全文内容'] || d['行内容'] || JSON.stringify(d);
          return `## ${i + 1}. ${t}\n${c}\n`;
        }).join('\n');
    } else if (targetFormat === 'md') {
      const gen = FormatGenerators.md(sampleData, options);
      text = gen.text;
    } else if (targetFormat === 'txt') {
      const gen = FormatGenerators.txt(sampleData, options);
      text = gen.text;
    } else if (targetFormat === 'html') {
      const gen = FormatGenerators.html(sampleData, options);
      text = gen.text;
    } else if (targetFormat === 'jsonl') {
      const lines = sampleData.map(row => JSON.stringify(row));
      text = lines.join('\n');
    } else if (targetFormat === 'json') {
      const space = options.prettyJson !== false ? 2 : 0;
      text = JSON.stringify(sampleData, null, space);
    } else if (targetFormat === 'csv') {
      const ws = XLSX.utils.json_to_sheet(sampleData);
      text = XLSX.utils.sheet_to_csv(ws);
    } else if (targetFormat === 'xlsx' || targetFormat === 'xls') {
      const keys = sampleData.length > 0 ? Object.keys(sampleData[0]) : [];
      text = `/* 目标工作簿: Excel (.${targetFormat}) */\n/* 字段清单 (${keys.length} 个): ${keys.join(', ')} */\n/* 数据行数: ${totalRows} 行 */\n\n` +
        XLSX.utils.sheet_to_csv(XLSX.utils.json_to_sheet(sampleData));
    }

    if (isTruncated) {
      text += `\n\n// ... 还有 ${totalRows - maxRows} 条记录 (已自动截取前 ${maxRows} 条用于实时快速预览)`;
    }

    return {
      text,
      lang,
      totalRows,
      previewRows,
      isTruncated
    };
  },

  // 生成目标格式的完整文本内容（用于一键复制）
  getFullTargetText(data, targetFormat, options = {}) {
    if (!data || data.length === 0) return '';
    if (FormatGenerators[targetFormat]) {
      const res = FormatGenerators[targetFormat](data, options);
      if (res && res.text) return res.text;
    }
    return JSON.stringify(data, null, 2);
  },

  async convert(file, srcFormat, targetFormat, options = {}) {
    const parsed = await this.parseFile(file, srcFormat);
    const data = parsed.data;
    if (!data || data.length === 0) {
      throw new Error('解析后的数据或文档为空，无法生成目标文件');
    }
    const generator = FormatGenerators[targetFormat];
    if (!generator) throw new Error(`不支持的转换为格式: ${targetFormat}`);

    const baseTitle = file.name.replace(/\.[^/.]+$/, '');
    return await generator(data, { ...options, title: baseTitle, parsedRaw: parsed });
  }
};



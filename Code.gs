// ============================================================
// نظام إدارة الطلبات - جمعية صناع الحياة المنوفية
// Google Apps Script Backend
// ============================================================
// BEFORE USING:
// 1. Go to Resources > Cloud Platform Project > enable Drive API
// 2. In the GCP Console, enable the "Google Drive API"
// 3. Deploy as Web App: Execute as "Me", Access "Anyone with link"

// ---------- CONFIGURATION ----------
const CONFIG = {
  rootFolderId: '1ZRdAIAjLBHv8ytI7mqUfH9wlyhZf8C_U',
  sheetId: '1x7UhPnFWM8RHYacFoa9tHG5yoWA7cRDHZg5xFkv-tW4',
  sheetName: 'all',
  templatesFolderName: 'تمبلت الطلبات',

  templates: {
    purchase: {
      fileName: 'تمبلت الشراء.docx',
      name: 'طلب شراء (سلعة/خدمة)',
      folderPrefix: 'طلبات شراء',
      sheetBand: 'مشتريات',
      fields: [
        {id: 'requestDate', label: 'تاريخ الطلب', type: 'date', required: true},
        {id: 'requestingDept', label: 'الإدارة الطالبة', type: 'text', required: true},
        {id: 'project', label: 'المشروع', type: 'text'},
        {id: 'subject', label: 'موضوع الطلب', type: 'text', required: true},
        {id: 'jobTitle', label: 'الوظيفة', type: 'text'},
        {id: 'requester', label: 'مقدم الطلب', type: 'text', required: true},
        {id: 'details', label: 'تفاصيل الطلب', type: 'textarea'},
        {id: 'requesterSignature', label: 'توقيع الجهة الطالبة', type: 'text'},
        {id: 'warehouseReview', label: 'مراجعة المخازن', type: 'text'},
        {id: 'financialApproval', label: 'الإدارة المالية', type: 'text'},
        {id: 'executiveApproval', label: 'المدير التنفيذي', type: 'text'}
      ],
      tableColumns: ['بند رقم', 'الصنف والمواصفات', 'الوحدة', 'الكمية', 'سعر الوحدة التقديرية', 'القيمة التقديرية'],
      tableField: 'items'
    },
    disbursement: {
      fileName: 'تمبلت طلبات الصرف.docx',
      name: 'إيصال استلام منحة عينية',
      folderPrefix: 'صرف منح',
      sheetBand: 'استلامات',
      fields: [
        {id: 'requestDate', label: 'تاريخ الاستلام', type: 'date', required: true},
        {id: 'recipientName', label: 'اسم المستلم', type: 'text', required: true},
        {id: 'nationalId', label: 'رقم قومي', type: 'text'},
        {id: 'acknowledgment', label: 'نص الإقرار', type: 'textarea'},
        {id: 'signatureName', label: 'الاسم (التوقيع)', type: 'text'},
        {id: 'signature', label: 'التوقيع', type: 'text'}
      ],
      tableFields: [
        {id: 'itemDesc', label: 'الصنف المستلم'}
      ],
      tableField: 'receivedItems'
    },
    cases: {
      fileName: 'كشف الحالات.docx',
      name: 'كشف توزيع',
      folderPrefix: 'كشوف توزيع',
      sheetBand: 'توزيع',
      fields: [
        {id: 'requestDate', label: 'تاريخ التوزيع', type: 'date', required: true},
        {id: 'title', label: 'عنوان الكشف', type: 'text', required: true}
      ],
      tableColumns: ['الاسم', 'العدد', 'التوقيع'],
      tableField: 'beneficiaries'
    }
  }
};

// ============================================================
// INCLUDE HELPER
// ============================================================
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ============================================================
// WEB APP
// ============================================================
function doGet(e) {
  const output = HtmlService.createHtmlOutput(
    '<html><body style="font-family:sans-serif;text-align:center;padding:40px;direction:rtl">' +
    '<h1>نظام إدارة الطلبات</h1>' +
    '<p>API Backend شغال ✅</p>' +
    '<p>استخدم <a href="https://kareemshair2.github.io/RequestFlow" target="_blank">GitHub Pages</a> للفورم</p>' +
    '<p>أو شغّل <code>setupSheet()</code> من المحرر</p>' +
    '</body></html>'
  ).setTitle('نظام إدارة الطلبات - API Backend');
  return output;
}

// POST handler for standalone HTML form (GitHub Pages)
function doPost(e) {
  try {
    let formData;
    if (e && e.postData && e.postData.contents) {
      formData = JSON.parse(e.postData.contents);
    } else {
      formData = e.parameter || {};
    }

    const result = submitRequest(formData);

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({success: false, error: err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ============================================================
// GET CONFIG for Frontend
// ============================================================
function getConfig() {
  const c = { templates: {} };
  for (const [key, tmpl] of Object.entries(CONFIG.templates)) {
    c.templates[key] = {
      name: tmpl.name,
      fields: tmpl.fields,
      tableColumns: tmpl.tableColumns,
      tableField: tmpl.tableField,
      tableFields: tmpl.tableFields
    };
  }
  return c;
}

// ============================================================
// MAIN SUBMISSION
// ============================================================
function submitRequest(formData) {
  try {
    const tmpl = CONFIG.templates[formData.templateKey];
    if (!tmpl) throw new Error('قالب غير معروف');

    const band = formData.band || tmpl.sheetBand || 'عام';
    const requestDate = formData.requestDate || Utilities.formatDate(new Date(), 'GMT+2', 'yyyy-MM-dd');
    const title = formData.title || formData.subject || 'طلب';

    // Create folder structure
    const folderInfo = createFolderStructure(band, requestDate, title);

    // Generate document
    const docInfo = generateDocument(tmpl, formData, folderInfo.folder, title);

    // Save to sheet
    saveToSheet(formData, docInfo.url, band);

    return {
      success: true,
      docUrl: docInfo.url,
      folderUrl: folderInfo.folderUrl,
      message: 'تم إنشاء الطلب بنجاح'
    };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// ============================================================
// FOLDER STRUCTURE
// ============================================================
function createFolderStructure(band, dateStr, title) {
  const rootFolder = DriveApp.getFolderById(CONFIG.rootFolderId);

  // Get or create band folder
  let bandFolder;
  const folders = rootFolder.getFoldersByName(band);
  bandFolder = folders.hasNext() ? folders.next() : rootFolder.createFolder(band);

  // Format date
  let dateObj = new Date(dateStr);
  if (isNaN(dateObj.getTime())) dateObj = new Date();
  const datePrefix = Utilities.formatDate(dateObj, 'GMT+2', 'yyyy-MM-dd');

  // Create dated subfolder
  const folderName = `${datePrefix} - ${title}`;
  const subFolders = bandFolder.getFoldersByName(folderName);
  const requestFolder = subFolders.hasNext() ? subFolders.next() : bandFolder.createFolder(folderName);

  return {
    folder: requestFolder,
    folderUrl: requestFolder.getUrl()
  };
}

// ============================================================
// DOCUMENT GENERATION
// ============================================================
function generateDocument(tmpl, formData, targetFolder, title) {
  const templatesFolder = findFolder(CONFIG.rootFolderId, CONFIG.templatesFolderName);

  // Find template file
  const templateFile = findTemplateFile(templatesFolder, tmpl.fileName);
  if (!templateFile) throw new Error('لم يتم العثور على ملف القالب: ' + tmpl.fileName);

  const docName = `${tmpl.name} - ${title}`;

  // Copy and convert to Google Doc
  const docId = Drive.Files.copy(
    {
      title: docName,
      parents: [{id: targetFolder.getId()}],
      mimeType: MimeType.GOOGLE_DOCS
    },
    templateFile.getId()
  ).id;

  // Fill content
  const doc = DocumentApp.openById(docId);
  const body = doc.getBody();

  fillDocumentFields(body, tmpl, formData);
  fillDocumentTables(body, tmpl, formData);

  doc.saveAndClose();

  return {
    url: `https://docs.google.com/document/d/${docId}/edit`,
    id: docId
  };
}

function findTemplateFile(folder, fileName) {
  const files = folder.getFilesByName(fileName);
  if (files.hasNext()) return files.next();

  // Search by partial name
  const allFiles = folder.getFiles();
  const baseName = fileName.replace('.docx', '').replace('.doc', '');
  while (allFiles.hasNext()) {
    const f = allFiles.next();
    if (f.getName().includes(baseName)) return f;
  }
  return null;
}

function fillDocumentFields(body, tmpl, formData) {
  for (const field of tmpl.fields) {
    const value = formData[field.id] || '';
    const placeholder = '{{' + field.id + '}}';
    if (body.findText(placeholder)) {
      body.replaceText(escapeRegex(placeholder), value);
    }
  }

  // Direct text replacement for known patterns
  const replacements = {
    'تاريخ الطلب:': { key: 'requestDate', position: 'after' },
    'تاريخ الاستلام:': { key: 'requestDate', position: 'after' },
    'تاريخ التوزيع:': { key: 'requestDate', position: 'after' },
    'الإدارة الطالبة:': { key: 'requestingDept', position: 'after' },
    'المشروع:': { key: 'project', position: 'after' },
    'موضوع الطلب:': { key: 'subject', position: 'after' },
    'الوظيفة:': { key: 'jobTitle', position: 'after' },
    'مقدم الطلب:': { key: 'requester', position: 'after' },
    'اقر انا /': { key: 'recipientName', position: 'replace' },
    'رقم قومي /': { key: 'nationalId', position: 'replace' },
  };

  for (const [pattern, rule] of Object.entries(replacements)) {
    const value = formData[rule.key];
    if (!value) continue;

    if (rule.position === 'after') {
      const found = body.findText(pattern);
      if (found) {
        const elem = found.getElement();
        const text = elem.asText();
        const start = found.getStartOffset();
        const matchLen = pattern.length;
        // Insert value after the label with a space
        if (start + matchLen < text.getText().length) {
          // Try to set the text after the label
          const fullText = text.getText();
          const before = fullText.substring(0, start + matchLen);
          const after = fullText.substring(start + matchLen);
          // If there's already content after the colon, clear it first
          const contentAfter = after.trim();
          if (contentAfter) {
            text.deleteText(start + matchLen, start + matchLen + contentAfter.length);
          }
          text.insertText(start + matchLen, ' ' + value);
        }
      }
    } else if (rule.position === 'replace') {
      const found = body.findText(pattern);
      if (found) {
        const elem = found.getElement();
        const text = elem.asText();
        const start = found.getStartOffset();
        const end = found.getEndOffsetInclusive();
        text.replaceText(pattern, pattern + ' ' + value);
      }
    }
  }
}

function fillDocumentTables(body, tmpl, formData) {
  if (!tmpl.tableField) return;
  const items = formData[tmpl.tableField];
  if (!items || !items.length) return;

  const tables = body.getTables();
  if (!tables.length) return;

  const table = tables[0];
  const numRows = table.getNumRows();

  // Find data rows (rows starting with a number or empty after header)
  let dataStartRow = -1;
  let totalRow = -1;

  // First pass: find header, data rows, and total row
  for (let r = 0; r < numRows; r++) {
    const row = table.getRow(r);
    const cell = row.getCell(0);
    const text = cell.getText().trim();
    if (/^\d+$/.test(text) && parseInt(text) >= 1 && parseInt(text) <= items.length) {
      if (dataStartRow === -1) dataStartRow = r;
    }
    if (text === 'الاجمالي' || text === 'الإجمالي' || text === 'المجموع') {
      totalRow = r;
      break;
    }
  }

  if (dataStartRow === -1) {
    // Try: after the first row that has column headers
    for (let r = 1; r < numRows; r++) {
      const row = table.getRow(r);
      const cell = row.getCell(0);
      const text = cell.getText().trim();
      if (text === '1') { dataStartRow = r; break; }
    }
  }

  if (dataStartRow === -1) {
    // Fallback: after header row (row 1 = header)
    dataStartRow = 1;
  }

  // Determine end row (before total or last row)
  let endRow = totalRow > -1 ? totalRow : numRows;
  const availableRows = endRow - dataStartRow;
  const fillCount = Math.min(items.length, availableRows);

  // Map column headers to field keys
  const colMap = {
    'بند رقم': null,  // auto-number
    'م': null,
    'الصنف والمواصفات': 'itemName',
    'الصنف': 'itemName',
    'اسم الصنف': 'itemName',
    'الوحدة': 'unit',
    'الكمية': 'quantity',
    'العدد': 'count',
    'سعر الوحدة التقديرية': 'unitPrice',
    'القيمة التقديرية': 'totalPrice',
    'الاسم': 'name',
    'التوقيع': 'signature'
  };

  // Read header to map columns
  const headerRow = table.getRow(dataStartRow > 0 ? dataStartRow - 1 : 0);
  const headerCells = headerRow.getNumCells();
  const columnMapping = [];
  for (let c = 0; c < headerCells; c++) {
    const headerText = headerRow.getCell(c).getText().trim();
    columnMapping.push(colMap[headerText] || null);
  }

  for (let i = 0; i < fillCount; i++) {
    const row = table.getRow(dataStartRow + i);
    const item = items[i] || {};
    const numCells = row.getNumCells();

    for (let c = 0; c < numCells && c < columnMapping.length; c++) {
      const fieldKey = columnMapping[c];
      let value = '';

      if (fieldKey === null) {
        value = String(i + 1);  // auto-number for بند رقم
      } else if (fieldKey) {
        value = item[fieldKey] || '';
      } else if (c === 0) {
        value = String(i + 1);
      }

      row.getCell(c).setText(value);
    }
  }

  // Calculate total if totalRow exists
  if (totalRow > -1) {
    let grandTotal = 0;
    for (const item of items) {
      const price = parseFloat(item.totalPrice) || 0;
      grandTotal += price;
    }
    if (grandTotal > 0) {
      const totalCells = table.getRow(totalRow).getNumCells();
      table.getRow(totalRow).getCell(totalCells - 1).setText(String(grandTotal));
    }
  }
}

// ============================================================
// SHEET OPERATIONS
// ============================================================
const SHEET_HEADERS = [
  'كود', 'تاريخ التقديم', 'البند', 'القالب', 'عنوان الطلب',
  'قيمته', 'رابط المستند', 'حالة', 'ملاحظات', 'تاريخ التسوية',
  'مقدم الطلب', 'تاريخ الإنشاء'
];

function getOrCreateBandSheet(band) {
  const ss = SpreadsheetApp.openById(CONFIG.sheetId);
  
  // Standardize band name
  const sheetName = band.trim() || 'عام';
  
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(SHEET_HEADERS);
    sheet.getRange(1, 1, 1, SHEET_HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function saveToSheet(formData, docUrl, band) {
  const dateStr = formData.requestDate
    ? Utilities.formatDate(new Date(formData.requestDate), 'GMT+2', 'dd/MM/yyyy')
    : Utilities.formatDate(new Date(), 'GMT+2', 'dd/MM/yyyy');

  let totalValue = formData.totalCost || '';
  if (!totalValue) {
    const items = formData.items || formData.receivedItems || formData.beneficiaries || [];
    let sum = 0;
    for (const item of items) {
      sum += parseFloat(item.totalPrice) || 0;
    }
    if (sum > 0) totalValue = String(sum);
  }

  const tmplName = CONFIG.templates[formData.templateKey]?.name || '';
  const requester = formData.requester || formData.recipientName || formData.title || '';

  const row = [
    dateStr,
    band,
    tmplName,
    formData.title || formData.subject || '',
    totalValue,
    docUrl,
    'جديد',
    formData.notes || '',
    '',
    requester,
    Utilities.formatDate(new Date(), 'GMT+2', 'dd/MM/yyyy HH:mm')
  ];

  // Save to band-specific tab
  const bandSheet = getOrCreateBandSheet(band);
  const bandCode = bandSheet.getLastRow();
  const bandRow = [bandCode].concat(row);
  bandSheet.appendRow(bandRow);

  // Save to master "all" sheet
  const ss = SpreadsheetApp.openById(CONFIG.sheetId);
  let allSheet = ss.getSheetByName('all');
  if (!allSheet) {
    allSheet = ss.insertSheet('all');
    allSheet.appendRow(SHEET_HEADERS);
    allSheet.getRange(1, 1, 1, SHEET_HEADERS.length).setFontWeight('bold');
    allSheet.setFrozenRows(1);
  }
  const allCode = allSheet.getLastRow();
  const allRow = [allCode].concat(row);
  allSheet.appendRow(allRow);
}

// Run this once from the Apps Script editor to set up the "all" sheet
function setupSheet() {
  const ss = SpreadsheetApp.openById(CONFIG.sheetId);
  
  // Create "all" sheet with headers
  let allSheet = ss.getSheetByName('all');
  if (!allSheet) {
    allSheet = ss.insertSheet('all');
  }
  allSheet.clear();
  allSheet.appendRow(SHEET_HEADERS);
  allSheet.getRange(1, 1, 1, SHEET_HEADERS.length).setFontWeight('bold');
  allSheet.setFrozenRows(1);
  
  // Create default band sheets
  const defaultBands = ['مشتريات', 'استلامات', 'توزيع', 'مخالفات', 'كفالات', 'نقل', 'إعانات', 'أخرى'];
  for (const band of defaultBands) {
    getOrCreateBandSheet(band);
  }
  
  return 'تم تجهيز الشيت بنجاح!';
}

// ============================================================
// UTILITY
// ============================================================
function findFolder(parentId, folderName) {
  const parent = DriveApp.getFolderById(parentId);
  const folders = parent.getFoldersByName(folderName);
  return folders.hasNext() ? folders.next() : parent.createFolder(folderName);
}

function escapeRegex(str) {
  return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

// ============================================================
// Code.gs — Google Apps Script for CSVC Survey (v5.1)
// Checklist đọc từ Sheet tab "HANG_MUC" → sửa Sheet là form tự update
// ============================================================

// ── CONFIG ──
const SHEET_DATA   = 'CSVC_Tracking';      // Tab lưu kết quả khảo sát
const SHEET_ITEMS  = 'HANG_MUC';           // Tab lưu danh sách hạng mục (Khu vực, Hạng mục)
const DRIVE_FOLDER = 'CSVC_Khao_Sat_Anh';  // Folder lưu ảnh trên Drive

// ── WEB APP ──
function doGet(e) {
  // ?action=checklist → trả về JSON danh sách hạng mục từ Sheet
  if (e && e.parameter && e.parameter.action === 'checklist') {
    const items = getChecklist();
    return ContentService.createTextOutput(JSON.stringify(items))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  // Mặc định → trả về form HTML
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Khao sat CSVC - WinRural v5')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const result = saveSubmission(data);
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({
      success: false,
      error: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ── GET CHECKLIST từ Sheet ──
function getChecklist() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ITEMS);
    if (!sheet) return null;
    
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return null;
    
    // Header: Khu vực, Hạng mục, Mã HM
    const areas = {};
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const khuVuc = String(row[0] || '').trim();
      const hangMuc = String(row[1] || '').trim();
      if (!khuVuc || !hangMuc) continue;
      
      if (!areas[khuVuc]) areas[khuVuc] = [];
      areas[khuVuc].push(hangMuc);
    }
    
    // Convert to array format
    const result = [];
    for (const [name, items] of Object.entries(areas)) {
      result.push({ name, items });
    }
    return result;
  } catch(e) {
    return null;
  }
}

// ── SAVE SUBMISSION ──
function saveSubmission(data) {
  const sheet = getOrCreateSheet(SHEET_DATA);
  const folder = getOrCreateFolder(DRIVE_FOLDER);
  
  const results = [];
  let savedPhotos = 0;
  
  for (const item of data.items) {
    // Save photos to Drive
    const photoUrls = [];
    if (item.photos && item.photos.length > 0) {
      for (let i = 0; i < item.photos.length; i++) {
        const photo = item.photos[i];
        if (photo.dataUrl && photo.dataUrl.startsWith('data:')) {
          try {
            const blob = dataUrlToBlob(photo.dataUrl);
            const safeFileName = sanitizeFileName(
              `${data.store}_${String(item.item).replace(/[^a-z0-9]/gi,'_')}_${i+1}.jpg`
            );
            const file = folder.createFile(blob.setName(safeFileName));
            photoUrls.push(file.getUrl());
            savedPhotos++;
          } catch(pe) {
            photoUrls.push('Lỗi: ' + pe.toString().substring(0, 50));
          }
        }
      }
    }
    
    const row = [
      new Date().toLocaleString('vi-VN'),
      data.store,
      data.storeName,
      data.tinh,
      data.gdv,
      data.qlkv,
      item.area,
      item.item,
      item.result,
      item.note || '',
      photoUrls.join('\n'),
      data.generalNote || ''
    ];
    sheet.appendRow(row);
    
    results.push({ item: item.item, result: item.result, photos: photoUrls.length });
  }
  
  return {
    success: true,
    timestamp: new Date().toISOString(),
    store: data.store,
    totalItems: data.totalItems,
    passCount: data.passCount,
    failCount: data.failCount,
    savedPhotos: savedPhotos,
    rowsSaved: results.length
  };
}

// ── HELPERS ──
function getOrCreateSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow([
      'Ngày báo', 'Mã CH', 'Tên CH', 'Tỉnh', 'GĐV', 'QLKV',
      'Khu vực', 'Hạng mục', 'Đánh giá', 'Diễn giải',
      'Ảnh Drive', 'Ghi chú chung'
    ]);
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, 12);
  }
  return sheet;
}

function getOrCreateFolder(name) {
  const folders = DriveApp.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(name);
}

function dataUrlToBlob(dataUrl) {
  const match = dataUrl.match(/^data:(.+);base64,(.+)$/);
  if (!match) throw new Error('Invalid data URL');
  return Utilities.newBlob(Utilities.base64Decode(match[2]), match[1]);
}

function sanitizeFileName(name) {
  return name.replace(/[\\/:*?"<>|]/g, '_').substring(0, 200);
}
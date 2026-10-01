// ============================================================
// Code.gs — Google Apps Script for CSVC Survey (v5)
// Deploy: Publish → Deploy as web app → Execute as "Me" → Anyone
// Copy URL & paste into index.html GAS_URL variable
// ============================================================

// ── CONFIG ──
const SHEET_NAME = 'CSVC_Tracking';   // Sheet tab name for submissions
const DRIVE_FOLDER_NAME = 'CSVC_Khao_Sat_Anh';  // Drive folder for photos

function doGet(e) {
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

// ── SAVE SUBMISSION ──
function saveSubmission(data) {
  const sheet = getOrCreateSheet(SHEET_NAME);
  const folder = getOrCreateFolder(DRIVE_FOLDER_NAME);
  const timestamp = new Date().toISOString();
  
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
              `${data.store}_${item.item.replace(/[^a-z0-9]/gi,'_')}_${i+1}.jpg`
            );
            const file = folder.createFile(blob.setName(safeFileName));
            photoUrls.push(file.getUrl());
            savedPhotos++;
          } catch(pe) {
            photoUrls.push('Lưu lỗi: ' + pe.toString().substring(0, 50));
          }
        }
      }
    }
    
    // Append row to sheet
    const row = [
      new Date().toLocaleString('vi-VN'),
      data.store,           // Mã CH
      data.storeName,       // Tên CH
      data.tinh,            // Tỉnh
      data.gdv,             // GĐV
      data.qlkv,            // QLKV
      item.area,            // Khu vực
      item.item,            // Hạng mục
      item.result,          // Đánh giá (Đạt/Không đạt)
      item.note || '',      // Diễn giải
      photoUrls.join('\n'), // Ảnh Drive URLs
      data.generalNote || '' // Ghi chú chung
    ];
    sheet.appendRow(row);
    
    results.push({
      item: item.item,
      result: item.result,
      photos: photoUrls.length
    });
  }
  
  return {
    success: true,
    timestamp: timestamp,
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
    // Create headers
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
  const mimeType = match[1];
  const base64 = match[2];
  const bytes = Utilities.base64Decode(base64);
  return Utilities.newBlob(bytes, mimeType);
}

function sanitizeFileName(name) {
  return name.replace(/[\\/:*?"<>|]/g, '_').substring(0, 200);
}
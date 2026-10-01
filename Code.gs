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

// ── GET CHECKLIST từ Sheet (auto-create nếu chưa có) ──
function getChecklist() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_ITEMS);
    
    // Auto-create HANG_MUC tab with default data if missing or empty
    if (!sheet || sheet.getLastRow() < 2) {
      sheet = createDefaultChecklist(ss);
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return null;
    
    const areas = {};
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const khuVuc = String(row[0] || '').trim();
      const hangMuc = String(row[1] || '').trim();
      if (!khuVuc || !hangMuc) continue;
      
      if (!areas[khuVuc]) areas[khuVuc] = [];
      areas[khuVuc].push(hangMuc);
    }
    
    const result = [];
    for (const [name, items] of Object.entries(areas)) {
      result.push({ name, items });
    }
    return result;
  } catch(e) {
    return null;
  }
}

function createDefaultChecklist(ss) {
  let sheet = ss.getSheetByName(SHEET_ITEMS);
  if (!sheet) sheet = ss.insertSheet(SHEET_ITEMS);
  sheet.clear();
  
  // Header
  sheet.appendRow(['Khu vực', 'Hạng mục', 'Mã HM']);
  
  // 33 hạng mục gốc từ phiếu khảo sát
  const items = [
    ['Trước CH', 'Biển hiệu WM', 'HM01'],
    ['Trước CH', 'Mặt tiền/Sơn tường ngoài', 'HM02'],
    ['Trước CH', 'Cửa kính/Cửa cuốn', 'HM03'],
    ['Trước CH', 'Mái hiên/Mái che', 'HM04'],
    ['Trước CH', 'Đèn chiếu sáng ngoài', 'HM05'],
    ['Trước CH', 'Bãi đỗ xe/Vỉa hè', 'HM06'],
    ['Trước CH', 'Biển quảng cáo/Poster', 'HM07'],
    ['Trước CH', 'Cây xanh/Cảnh quan', 'HM08'],
    ['Trước CH', 'Thùng rác ngoài', 'HM09'],
    ['Trước CH', 'Camera an ninh ngoài', 'HM10'],
    ['Trước CH', 'Lối vào/Cửa thoát hiểm', 'HM11'],
    ['Trong CH', 'Sàn nhà/Gạch lát', 'HM12'],
    ['Trong CH', 'Trần nhà/Thạch cao', 'HM13'],
    ['Trong CH', 'Tường trong/Sơn nội thất', 'HM14'],
    ['Trong CH', 'Đèn chiếu sáng trong', 'HM15'],
    ['Trong CH', 'Điều hòa/Quạt/Thông gió', 'HM16'],
    ['Trong CH', 'Kệ trưng bày/Sạp hàng', 'HM17'],
    ['Trong CH', 'Tủ mát/Tủ đông/Tủ lạnh', 'HM18'],
    ['Trong CH', 'Bảng giá/Tag giá điện tử', 'HM19'],
    ['Trong CH', 'Giỏ hàng/Xe đẩy', 'HM20'],
    ['Trong CH', 'Quầy thu ngân/POS', 'HM21'],
    ['Trong CH', 'Cân điện tử', 'HM22'],
    ['Trong CH', 'Wifi/Mạng/Thiết bị CNTT', 'HM23'],
    ['Trong CH', 'Hệ thống âm thanh', 'HM24'],
    ['Trong CH', 'Bình chữa cháy/PCCC', 'HM25'],
    ['Trong CH', 'Biển báo lối thoát hiểm', 'HM26'],
    ['Trong CH', 'Ổ điện/Hệ thống điện', 'HM27'],
    ['Trong CH', 'Camera an ninh trong', 'HM28'],
    ['Kho & Khu vực Nhân viên', 'Cửa kho', 'HM29'],
    ['Kho & Khu vực Nhân viên', 'Kệ kho/Pallet', 'HM30'],
    ['Kho & Khu vực Nhân viên', 'Đèn kho', 'HM31'],
    ['Kho & Khu vực Nhân viên', 'Nhà vệ sinh nhân viên', 'HM32'],
    ['Kho & Khu vực Nhân viên', 'Khu vực nghỉ nhân viên', 'HM33']
  ];
  
  for (const row of items) {
    sheet.appendRow(row);
  }
  
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, 3);
  return sheet;
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
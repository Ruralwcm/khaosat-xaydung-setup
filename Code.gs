// ============================================================
// Code.gs — Google Apps Script for CSVC Survey (v5.3)
// Checklist cập nhật từ Demo_Rà_soát_CH_xuống_cấp (01/10/2026)
// 34 hạng mục · 3 khu vực: TRƯỚC CH (10) + TRONG CH (16) + KHO - NVS (8)
// ============================================================

// ── CONFIG ──
const SHEET_DATA   = 'CSVC_Tracking';          // Tab lưu kết quả khảo sát
const SHEET_ITEMS  = 'HANG_MUC';               // Tab lưu danh sách hạng mục (Khu vực, Hạng mục, Mã HM, Nội dung rà soát)
const DRIVE_FOLDER = 'CSVC_Khao_Sat_Anh';      // Folder lưu ảnh trên Drive

// ── WEB APP ──
function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'checklist') {
    const items = getChecklist();
    return ContentService.createTextOutput(JSON.stringify(items))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  if (e && e.parameter && e.parameter.action === 'storeData') {
    const data = getStoreData();
    return ContentService.createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Khảo sát CSVC - WinRural v5.3')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── GET STORE DATA (GĐV → QLKV → CH cascade từ Storeslist) ──
function getStoreData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Storeslist');
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];
    
    // Tìm header row: GĐV, QLKV, Mã CH, Tên CH, Tỉnh, Tình trạng
    const headers = data[0].map(h => String(h || '').trim());
    const idxGDV = headers.findIndex(h => h === 'GĐV' || h === 'GDV');
    const idxQLKV = headers.findIndex(h => h === 'QLKV');
    const idxMa = headers.findIndex(h => h === 'Mã CH' || h === 'Mã SAP' || h === 'SAP');
    const idxTen = headers.findIndex(h => h === 'Tên cửa hàng' || h === 'Tên CH' || h === 'Store Name');
    const idxTinh = headers.findIndex(h => h === 'Tỉnh/TP' || h === 'Tỉnh' || h === 'Province');
    const idxTT = headers.findIndex(h => h === 'Tình trạng' || h === 'Tình Trạng' || h === 'Status');
    
    if (idxGDV < 0 || idxQLKV < 0 || idxMa < 0) return [];
    
    const result = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      // Filter active stores only
      if (idxTT >= 0) {
        const tt = String(row[idxTT] || '').trim();
        if (tt !== 'Đang hoạt động') continue;
      }
      const gdv = String(row[idxGDV] || '').trim();
      const qlkv = String(row[idxQLKV] || '').trim();
      const ma = String(row[idxMa] || '').trim();
      if (!gdv || !qlkv || !ma) continue;
      
      result.push({
        gdv: gdv,
        qlkv: qlkv,
        ma: ma,
        ten: String(row[idxTen] || '').trim(),
        tinh: String(row[idxTinh] || '').trim()
      });
    }
    return result;
  } catch(e) {
    return [];
  }
}

// ── SUBMIT SURVEY (nhận data từ form → save) ──
function submitSurvey(data) {
  return saveSubmission(data);
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

// ── GET CHECKLIST từ Sheet (auto-create nếu chưa có hoặc khác version) ──
function getChecklist() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_ITEMS);
    
    // Auto-create HANG_MUC tab with v5.2 data if missing or empty
    if (!sheet || sheet.getLastRow() < 2) {
      sheet = createChecklistV52(ss);
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

// ── CREATE CHECKLIST v5.2: 34 hạng mục từ Demo_Rà_soát_CH_xuống_cấp ──
function createChecklistV52(ss) {
  let sheet = ss.getSheetByName(SHEET_ITEMS);
  if (!sheet) sheet = ss.insertSheet(SHEET_ITEMS);
  sheet.clear();
  
  // Header: Khu vực | Hạng mục kiểm tra | Mã HM | Nội dung rà soát
  sheet.appendRow(['Khu vực', 'Hạng mục kiểm tra', 'Mã HM', 'Nội dung rà soát']);
  
  // ── 34 hạng mục (chuẩn từ checklist Demo 01/10/2026) ──
  const items = [
    // TRƯỚC CH (10 mục)
    ['TRƯỚC CH', 'Biển hiệu WM+', 'CSVC01', 'Kiểm tra biển hiệu có sáng đầy đủ không; bóng LED có cháy/nhấp nháy không; mặt biển có bạc màu, nứt, bong tróc không; chữ/logo có thiếu hoặc hỏng không'],
    ['TRƯỚC CH', 'Mặt tiền/ sơn tường ngoài', 'CSVC02', 'Kiểm tra sơn mặt tiền có bong tróc, nứt, ố mốc, bạc màu không; tường có thấm nước hoặc rêu mốc không'],
    ['TRƯỚC CH', 'Cửa cuốn/ cửa kính', 'CSVC03', 'Kiểm tra cửa cuốn có đóng/mở bình thường không; cửa có kẹt, lệch, han gỉ, móp méo không; cửa kính có nứt/vỡ/xước không'],
    ['TRƯỚC CH', 'Mái hiên/ mái che', 'CSVC04', 'Kiểm tra mái hiên/mái che có dột, thấm, võng, nứt hoặc han gỉ không; khung mái có chắc chắn không'],
    ['TRƯỚC CH', 'Đèn chiếu sáng ngoài', 'CSVC05', 'Kiểm tra toàn bộ đèn ngoài CH: đèn có cháy, yếu, chập chờn không; khu vực mặt tiền, lối vào, biển hiệu có đủ sáng không'],
    ['TRƯỚC CH', 'Vỉa hè', 'CSVC06', 'Kiểm tra vỉa hè có nứt, vỡ, sụt lún, kênh hoặc đọng nước không; có vật cản gây khó khăn cho KH không'],
    ['TRƯỚC CH', 'Biển quảng cáo/Poster', 'CSVC07', 'Kiểm tra poster/decal/biển quảng cáo có cũ, rách, bong mép, bạc màu không; nội dung đã hết chương trình chưa'],
    ['TRƯỚC CH', 'Cây xanh', 'CSVC08', 'Kiểm tra cây xanh có khô/héo, chết, mọc um tùm hoặc che biển hiệu/lối đi không; chậu cây có vỡ, bẩn hoặc mất vệ sinh không'],
    ['TRƯỚC CH', 'Thùng rác ngoài', 'CSVC09', 'Kiểm tra thùng rác có đủ số lượng, nắp/túi, bánh xe không; có nứt vỡ, rò rỉ, bẩn hoặc có mùi không'],
    ['TRƯỚC CH', 'Camera an ninh ngoài', 'CSVC10', 'Kiểm tra camera ngoài CH có hoạt động không; góc quay có bị che khuất không; camera có lỏng, nghiêng, hỏng hoặc mất tín hiệu không'],
    
    // TRONG CH (16 mục)
    ['TRONG CH', 'Trần nhà/ thạch cao', 'CSVC11', 'Kiểm tra trần có dột/thấm, ố vàng, mốc, nứt, bong hoặc võng không; tấm thạch cao có vỡ/mất không'],
    ['TRONG CH', 'Tường', 'CSVC12', 'Kiểm tra tường trong CH có bong sơn, nứt, ố, mốc, thấm nước hoặc va đập không; chân tường có hư hỏng không'],
    ['TRONG CH', 'Đèn chiếu sáng', 'CSVC13', 'Kiểm tra toàn bộ đèn trong CH: số lượng đèn cháy/yếu/chập chờn; khu vực bán hàng có bị tối không'],
    ['TRONG CH', 'Điều hòa/ quạt/ thông gió', 'CSVC14', 'Kiểm tra điều hòa/quạt/thông gió có hoạt động bình thường không; khả năng làm mát có đảm bảo không'],
    ['TRONG CH', 'Kệ trưng bày', 'CSVC15', 'Kiểm tra kệ có cong, gãy, han gỉ, bong sơn hoặc thiếu chi tiết không; chân kệ có chắc chắn không'],
    ['TRONG CH', 'Tủ mát/ tủ đông/ tủ lạnh', 'CSVC16', 'Kiểm tra tủ mát/tủ đông/tủ lạnh có hoạt động ổn định không; nhiệt độ có đảm bảo không'],
    ['TRONG CH', 'Bảng giá', 'CSVC17', 'Kiểm tra bảng giá có đầy đủ và đúng vị trí không; bảng có cũ, cong, bẩn hoặc hư hỏng không'],
    ['TRONG CH', 'Giỏ hàng/ xe đẩy', 'CSVC18', 'Kiểm tra giỏ hàng/xe đẩy có đủ số lượng không; bánh xe có kẹt/lệch không; tay cầm có hỏng không'],
    ['TRONG CH', 'Quầy thu ngân', 'CSVC19', 'Kiểm tra quầy thu ngân có hư hỏng, bong tróc, nứt hoặc xuống cấp không; mặt bàn, ngăn kéo có chắc chắn không'],
    ['TRONG CH', 'Cân điện tử', 'CSVC20', 'Kiểm tra cân điện tử có hoạt động và hiển thị chính xác không; màn hình/phím bấm có lỗi không'],
    ['TRONG CH', 'Wifi/ Mạng/ Thiết bị IT', 'CSVC21', 'Kiểm tra Wifi/mạng có ổn định không; modem/router/switch có hoạt động không'],
    ['TRONG CH', 'Hạng mục khác', 'CSVC22', 'Rà soát các hạng mục cơ sở vật chất khác chưa có trong checklist: biển chỉ dẫn, giá treo, tay vịn, phụ kiện, vật dụng cố định'],
    ['TRONG CH', 'Hệ thống âm thanh', 'CSVC23', 'Kiểm tra loa và hệ thống âm thanh có hoạt động không; âm lượng có rõ và đồng đều không'],
    ['TRONG CH', 'Bình chữa cháy/ PCCC', 'CSVC24', 'Kiểm tra bình chữa cháy có đủ số lượng và đúng vị trí không; bình có còn hạn/tem kiểm định không'],
    ['TRONG CH', 'Ổ điện, hệ thống điện', 'CSVC25', 'Kiểm tra ổ cắm, công tắc, tủ điện và dây điện có cháy xém, nứt vỡ, lỏng hoặc quá tải không'],
    ['TRONG CH', 'Camera an ninh trong', 'CSVC26', 'Kiểm tra camera trong CH có hoạt động và ghi hình không; hình ảnh có rõ không'],
    
    // KHO - NVS (8 mục)
    ['KHO - NVS', 'Cửa kho / cửa NVS', 'CSVC27', 'Kiểm tra cửa kho/cửa NVS có đóng mở bình thường không; khóa, bản lề, tay nắm có hỏng không'],
    ['KHO - NVS', 'Kệ kho/ Palet', 'CSVC28', 'Kiểm tra kệ kho/palet có đủ chắc chắn không; có cong, gãy, han gỉ hoặc mối mọt không'],
    ['KHO - NVS', 'Sàn kho / NVS', 'CSVC29', 'Kiểm tra sàn kho/NVS có nứt, vỡ, sụt lún, bong gạch, đọng nước hoặc trơn trượt không'],
    ['KHO - NVS', 'Đường điện', 'CSVC30', 'Kiểm tra đường điện trong kho/NVS có dây hở, dây chắp nối, ổ cắm lỏng hoặc đi dây mất an toàn không'],
    ['KHO - NVS', 'Thiết bị vệ sinh', 'CSVC31', 'Kiểm tra vòi nước, lavabo, bồn rửa, van khóa, vòi xịt và các thiết bị liên quan có hoạt động không'],
    ['KHO - NVS', 'Đường nước', 'CSVC32', 'Kiểm tra đường cấp/thoát nước có rò rỉ, tắc nghẽn, nứt vỡ hoặc mùi bất thường không'],
    ['KHO - NVS', 'Thiết bị nhà vệ sinh', 'CSVC33', 'Kiểm tra bồn cầu, lavabo, vòi xịt, gương, móc treo, đèn, quạt thông gió và phụ kiện NVS'],
    ['KHO - NVS', 'Sàn kho / NVS (tổng thể)', 'CSVC34', 'Kiểm tra tổng thể sàn kho/NVS tại các vị trí có nguy cơ xuống cấp: nứt, vỡ, bong, lún, thấm, đọng nước']
  ];
  
  for (const row of items) {
    sheet.appendRow(row);
  }
  
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, 4);
  return sheet;
}

// ── SAVE SUBMISSION ──
function saveSubmission(data) {
  const sheet = getOrCreateSheet(SHEET_DATA);
  const folder = getOrCreateFolder(DRIVE_FOLDER);
  
  const results = [];
  let savedPhotos = 0;
  
  for (const item of data.items) {
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

// ── FORCE RELOAD CHECKLIST (gọi manually để update HANG_MUC từ file mới) ──
function reloadChecklist() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = createChecklistV52(ss);
  return { success: true, message: 'Đã cập nhật HANG_MUC lên v5.2 (34 hạng mục)', rowCount: sheet.getLastRow() - 1 };
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
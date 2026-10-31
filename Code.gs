// ============================================================
// Code.gs — Google Apps Script for Khảo sát hỏng hóc Xây dựng - Kỹ Thuật (v14.1)
// v12 NEW: action=data endpoint for full export (all CSVC_Tracking rows)
// v13 FIX: đọc Storeslist từ sheet nguồn Storelist Rural (openById) + normalize NFC header (fix 'Mã SAP' dấu ngã tổ hợp)
// v14 NEW: action=unsurveyed → danh sách CH chưa khảo sát (cho "Cùng nhắc đồng đội")
// v14.1 FIX: getStoreData trả về 'mien' (Miền) từ Storelist → cascade Miền hoạt động
// ============================================================

// ── CONFIG ──
const SHEET_DATA   = 'CSVC_Tracking';          // Tab lưu kết quả khảo sát
const SHEET_ITEMS  = 'HANG_MUC';               // Tab danh sách hạng mục
const DRIVE_FOLDER = 'CSVC_Khao_Sat_Anh';      // Folder lưu ảnh Drive
const STORESLIST_ID = '1UMtB9MOxGOaZvxN5VAUGAS1bPaAL2tVVFEViGANd4Tk'; // Storelist Rural (nguồn danh sách CH)

// Lấy tab Storeslist từ sheet nguồn Storelist Rural (không copy vào sheet survey)
function getStoreslistSheet() {
  try {
    return SpreadsheetApp.openById(STORESLIST_ID).getSheetByName('Storeslist');
  } catch (e) { return null; }
}

// ── WEB APP ──
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || '';

  if (action === 'checklist') {
    const items = getChecklist();
    return ContentService.createTextOutput(JSON.stringify(items))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (action === 'storeData') {
    const data = getStoreData();
    return ContentService.createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (action === 'history') {
    const store = (e && e.parameter && e.parameter.store) || '';
    const data = getHistory(store);
    return ContentService.createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (action === 'dashboard') {
    const data = getDashboardData();
    return ContentService.createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // v12 NEW: action=data → trả về TOÀN BỘ dữ liệu khảo sát (cho export)
  if (action === 'data') {
    const data = getAllSurveyData();
    return ContentService.createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // v14 NEW: action=unsurveyed → danh sách CH chưa khảo sát (cho "Cùng nhắc đồng đội")
  if (action === 'unsurveyed') {
    const data = getUnsurveyedStores();
    return ContentService.createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Default: serve Index.html
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Khảo sát hỏng hóc Xây dựng - Kỹ Thuật')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── GET STORE DATA ──
function getStoreData() {
  try {
    const sheet = getStoreslistSheet();
    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];

    const headers = data[0].map(h => String(h || '').normalize('NFC').trim());
    const idxGDV = headers.findIndex(h => h === 'GĐV' || h === 'GDV');
    const idxQLKV = headers.findIndex(h => h === 'QLKV');
    const idxMa = headers.findIndex(h => h === 'Mã CH' || h === 'Mã SAP' || h === 'SAP');
    const idxTen = headers.findIndex(h => h === 'Tên cửa hàng' || h === 'Tên CH' || h === 'Store Name');
    const idxTinh = headers.findIndex(h => h === 'Tỉnh/TP' || h === 'Tỉnh' || h === 'Province');
    const idxTT = headers.findIndex(h => h === 'Tình trạng' || h === 'Tình Trạng' || h === 'Status');
    const idxMien = headers.findIndex(h => h === 'Miền');

    if (idxGDV < 0 || idxQLKV < 0 || idxMa < 0) return [];

    const result = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (idxTT >= 0) {
        const tt = String(row[idxTT] || '').trim();
        if (tt !== 'Đang hoạt động') continue;
      }
      const gdv = String(row[idxGDV] || '').trim();
      const qlkv = String(row[idxQLKV] || '').trim();
      const ma = String(row[idxMa] || '').trim();
      if (!gdv || !qlkv || !ma) continue;

      result.push({
        gdv: gdv, qlkv: qlkv, ma: ma,
        ten: String(row[idxTen] || '').trim(),
        tinh: String(row[idxTinh] || '').trim(),
        mien: idxMien >= 0 ? String(row[idxMien] || '').trim() : ''
      });
    }
    return result;
  } catch(e) { return []; }
}

// ── v12 NEW: GET ALL SURVEY DATA (for export) ──
function getAllSurveyData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_DATA);
    if (!sheet || sheet.getLastRow() < 2) return { ok: true, rows: [], headers: [] };

    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h || '').trim());

    const rows = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const obj = {};
      headers.forEach((h, j) => { obj[h] = row[j] == null ? '' : String(row[j]); });
      rows.push(obj);
    }

    return { ok: true, headers: headers, rows: rows, total: rows.length };
  } catch(e) {
    return { ok: false, error: e.toString() };
  }
}

// ── v14 NEW: GET UNSURVEYED STORES (CH chưa khảo sát) ──
function getUnsurveyedStores() {
  try {
    const slSheet = getStoreslistSheet();
    const activeStores = [];
    if (slSheet && slSheet.getLastRow() > 1) {
      const slData = slSheet.getDataRange().getValues();
      const slHeaders = slData[0].map(h => String(h || '').normalize('NFC').trim());
      const iG = slHeaders.findIndex(h => h === 'GĐV' || h === 'GDV');
      const iQ = slHeaders.findIndex(h => h === 'QLKV');
      const iM = slHeaders.findIndex(h => h === 'Mã CH' || h === 'Mã SAP' || h === 'SAP');
      const iT = slHeaders.findIndex(h => h === 'Tên cửa hàng' || h === 'Tên CH');
      const iP = slHeaders.findIndex(h => h === 'Tỉnh/TP' || h === 'Tỉnh');
      const iTT = slHeaders.findIndex(h => h === 'Tình trạng' || h === 'Tình Trạng');
      for (let i = 1; i < slData.length; i++) {
        const r = slData[i];
        if (iTT >= 0 && String(r[iTT] || '').trim() !== 'Đang hoạt động') continue;
        const gdv = String(r[iG] || '').trim();
        const qlkv = String(r[iQ] || '').trim();
        const ma = String(r[iM] || '').trim();
        if (!gdv || !qlkv || !ma) continue;
        activeStores.push({ ma: ma, ten: String(r[iT] || '').trim(), tinh: String(r[iP] || '').trim(), gdv: gdv, qlkv: qlkv });
      }
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const trSheet = ss.getSheetByName(SHEET_DATA);
    const surveyed = {};
    if (trSheet && trSheet.getLastRow() > 1) {
      const trData = trSheet.getDataRange().getValues();
      const trHeaders = trData[0].map(h => String(h || '').trim());
      const iMa = trHeaders.findIndex(h => h === 'Mã CH' || h === 'Mã SAP');
      if (iMa >= 0) {
        for (let i = 1; i < trData.length; i++) {
          const ma = String(trData[i][iMa] || '').trim();
          if (ma) surveyed[ma] = true;
        }
      }
    }

    const uns = activeStores.filter(s => !surveyed[s.ma]);
    const byGdv = {};
    uns.forEach(s => { if (!byGdv[s.gdv]) byGdv[s.gdv] = []; byGdv[s.gdv].push(s); });
    return { ok: true, total: uns.length, totalActive: activeStores.length, byGdv: byGdv, stores: uns };
  } catch(e) {
    return { ok: false, error: e.toString() };
  }
}

// ── GET HISTORY ──
function getHistory(storeCode) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_DATA);
    if (!sheet || sheet.getLastRow() < 2) return [];

    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h || '').trim());
    const idx = {
      date: headers.findIndex(h => h.includes('Ngày báo')),
      maCH: headers.findIndex(h => h === 'Mã CH' || h === 'Mã SAP'),
      tenCH: headers.findIndex(h => h === 'Tên CH'),
      tinh: headers.findIndex(h => h === 'Tỉnh'),
      gdv: headers.findIndex(h => h === 'GĐV'),
      qlkv: headers.findIndex(h => h === 'QLKV'),
      area: headers.findIndex(h => h === 'Khu vực'),
      item: headers.findIndex(h => h === 'Hạng mục'),
      result: headers.findIndex(h => h === 'Đánh giá'),
      note: headers.findIndex(h => h === 'Diễn giải'),
      photos: headers.findIndex(h => h === 'Ảnh Drive'),
      code: headers.findIndex(h => h === 'Mã theo dõi'),
      psDate: headers.findIndex(h => h === 'Ngày PS'),
      su: headers.findIndex(h => h === 'Đã báo SU'),
      suDate: headers.findIndex(h => h === 'Ngày báo SU')
    };
    if (idx.maCH < 0) return [];

    const submissionsMap = {};
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const ma = String(row[idx.maCH] || '').trim();
      if (ma !== storeCode) continue;
      const submitDate = String(row[idx.date] || '').trim();
      const code = String(row[idx.code] || '').trim();
      const key = submitDate + '|' + code;
      if (!submissionsMap[key]) {
        submissionsMap[key] = {
          date: submitDate, store: ma,
          storeName: String(row[idx.tenCH] || '').trim(),
          tinh: String(row[idx.tinh] || '').trim(),
          gdv: String(row[idx.gdv] || '').trim(),
          qlkv: String(row[idx.qlkv] || '').trim(),
          items: []
        };
      }
      let photoList = [];
      const photoStr = String(row[idx.photos] || '').trim();
      if (photoStr) photoList = photoStr.split('\n').filter(u => u && !u.startsWith('Lỗi:'));
      submissionsMap[key].items.push({
        area: String(row[idx.area] || '').trim(),
        item: String(row[idx.item] || '').trim(),
        result: String(row[idx.result] || '').trim(),
        note: String(row[idx.note] || '').trim(),
        photos: photoList, code: code,
        date: String(row[idx.psDate] || '').trim(),
        su: String(row[idx.su] || '').trim(),
        sudate: String(row[idx.suDate] || '').trim()
      });
    }
    const result = Object.values(submissionsMap);
    result.sort((a, b) => b.date.localeCompare(a.date));
    return result;
  } catch(e) { return []; }
}

// ── DASHBOARD ──
function getDashboardData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const slSheet = getStoreslistSheet();
    const storeMap = {};
    const gdvStats = {};
    const qlkvStats = {};

    if (slSheet && slSheet.getLastRow() > 1) {
      const slData = slSheet.getDataRange().getValues();
      const slHeaders = slData[0].map(h => String(h || '').normalize('NFC').trim());
      const iG = slHeaders.findIndex(h => h === 'GĐV' || h === 'GDV');
      const iQ = slHeaders.findIndex(h => h === 'QLKV');
      const iM = slHeaders.findIndex(h => h === 'Mã CH' || h === 'Mã SAP' || h === 'SAP');
      const iT = slHeaders.findIndex(h => h === 'Tên cửa hàng' || h === 'Tên CH');
      const iP = slHeaders.findIndex(h => h === 'Tỉnh/TP' || h === 'Tỉnh');
      const iTT = slHeaders.findIndex(h => h === 'Tình trạng' || h === 'Tình Trạng');
      for (let i = 1; i < slData.length; i++) {
        const r = slData[i];
        if (iTT >= 0 && String(r[iTT] || '').trim() !== 'Đang hoạt động') continue;
        const gdv = String(r[iG] || '').trim();
        const qlkv = String(r[iQ] || '').trim();
        const ma = String(r[iM] || '').trim();
        if (!gdv || !qlkv || !ma) continue;
        storeMap[ma] = { gdv, qlkv, tinh: String(r[iP] || '').trim(), ten: String(r[iT] || '').trim() };
        if (!gdvStats[gdv]) gdvStats[gdv] = { totalCH: 0, qlkvMap: {} };
        gdvStats[gdv].totalCH++;
        if (!gdvStats[gdv].qlkvMap[qlkv]) gdvStats[gdv].qlkvMap[qlkv] = 0;
        gdvStats[gdv].qlkvMap[qlkv]++;
        const qk = gdv + '|' + qlkv;
        if (!qlkvStats[qk]) qlkvStats[qk] = { gdv, qlkv, totalCH: 0 };
        qlkvStats[qk].totalCH++;
      }
    }

    const trSheet = ss.getSheetByName(SHEET_DATA);
    const surveyedSet = {}, issueByGdv = {}, issueByQlkv = {}, weekData = {};
    if (trSheet && trSheet.getLastRow() > 1) {
      const trData = trSheet.getDataRange().getValues();
      const trHeaders = trData[0].map(h => String(h || '').trim());
      const iMa = trHeaders.findIndex(h => h === 'Mã CH' || h === 'Mã SAP');
      const iRes = trHeaders.findIndex(h => h === 'Đánh giá');
      const iDate = trHeaders.findIndex(h => h.includes('Ngày báo'));
      for (let i = 1; i < trData.length; i++) {
        const r = trData[i];
        const ma = String(r[iMa] || '').trim();
        if (!ma) continue;
        const store = storeMap[ma];
        if (!store) continue;
        const gdv = store.gdv, qlkv = store.qlkv;
        const result = String(r[iRes] || '').trim();
        const dateStr = String(r[iDate] || '').trim();
        surveyedSet[ma] = true;
        const w = weekFromDate(dateStr);
        if (w) { if (!weekData[w]) weekData[w] = { surveyed: {}, issues: 0 }; weekData[w].surveyed[ma] = true; weekData[w].issues++; }
        if (!issueByGdv[gdv]) issueByGdv[gdv] = { total: 0, resolved: 0 };
        issueByGdv[gdv].total++;
        if (result === 'Đạt') issueByGdv[gdv].resolved++;
        const qk = gdv + '|' + qlkv;
        if (!issueByQlkv[qk]) issueByQlkv[qk] = { total: 0, resolved: 0 };
        issueByQlkv[qk].total++;
        if (result === 'Đạt') issueByQlkv[qk].resolved++;
      }
    }

    let totalCH = 0, surveyedCH = 0;
    const byGdv = [], byQlkv = [];
    for (const [gdv, stats] of Object.entries(gdvStats).sort()) {
      totalCH += stats.totalCH;
      let gSurv = 0, gIss = 0, gRes = 0;
      const ibg = issueByGdv[gdv];
      if (ibg) { gIss = ibg.total; gRes = ibg.resolved; }
      for (const ma of Object.keys(surveyedSet)) { if (storeMap[ma] && storeMap[ma].gdv === gdv) gSurv++; }
      byGdv.push({ gdv, totalCH: stats.totalCH, surveyed: gSurv, notSurveyed: stats.totalCH - gSurv, pctSurveyed: stats.totalCH > 0 ? Math.round(gSurv / stats.totalCH * 100) : 0, totalIssues: gIss, resolved: gRes, unresolved: gIss - gRes, pctResolved: gIss > 0 ? Math.round(gRes / gIss * 100) : 0 });
    }
    surveyedCH = Object.keys(surveyedSet).length;
    for (const [qk, qs] of Object.entries(qlkvStats).sort()) {
      let qSurv = 0;
      for (const ma of Object.keys(surveyedSet)) { if (storeMap[ma] && storeMap[ma].gdv + '|' + storeMap[ma].qlkv === qk) qSurv++; }
      const ibq = issueByQlkv[qk] || { total: 0, resolved: 0 };
      byQlkv.push({ gdv: qs.gdv, qlkv: qs.qlkv, totalCH: qs.totalCH, surveyed: qSurv, notSurveyed: qs.totalCH - qSurv, pctSurveyed: qs.totalCH > 0 ? Math.round(qSurv / qs.totalCH * 100) : 0, totalIssues: ibq.total, resolved: ibq.resolved, unresolved: ibq.total - ibq.resolved, pctResolved: ibq.total > 0 ? Math.round(ibq.resolved / ibq.total * 100) : 0 });
    }
    const weeks = Object.entries(weekData).sort((a, b) => b[0].localeCompare(a[0])).slice(0, 12).map(([w, d]) => ({ week: w, surveyed: Object.keys(d.surveyed).length, issues: d.issues })).reverse();
    let allIss = 0, allRes = 0;
    for (const [_, ib] of Object.entries(issueByGdv)) { allIss += ib.total; allRes += ib.resolved; }

    return { ok: true, summary: { totalCH, surveyedCH, notSurveyed: totalCH - surveyedCH, pctSurveyed: totalCH > 0 ? Math.round(surveyedCH / totalCH * 100) : 0, totalIssues: allIss, resolved: allRes, unresolved: allIss - allRes, pctResolved: allIss > 0 ? Math.round(allRes / allIss * 100) : 0 }, byGdv, byQlkv, byWeek: weeks, updated: new Date().toISOString() };
  } catch(e) { return { ok: false, error: e.toString() }; }
}

function weekFromDate(dateStr) {
  const m = String(dateStr).match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (!m) return null;
  const d = new Date(+m[3], +m[2] - 1, +m[1]);
  const start = new Date(d.getFullYear(), 0, 1);
  const weekNum = Math.ceil(((d - start) / 86400000 + start.getDay() + 1) / 7);
  return d.getFullYear() + '-W' + ('0' + weekNum).slice(-2);
}

// ── POST ──
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const result = saveSubmission(data);
    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.toString() })).setMimeType(ContentService.MimeType.JSON);
  }
}

function submitSurvey(data) { return saveSubmission(data); }

// ── CHECKLIST ──
function getChecklist() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_ITEMS);
    if (!sheet || sheet.getLastRow() < 2) sheet = createChecklistV11(ss);
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];
    const areas = {};
    for (let i = 1; i < data.length; i++) {
      const row = data[i], khuVuc = String(row[0] || '').trim(), hangMuc = String(row[1] || '').trim();
      if (!khuVuc || !hangMuc) continue;
      if (!areas[khuVuc]) areas[khuVuc] = [];
      areas[khuVuc].push(hangMuc);
    }
    return Object.entries(areas).map(([name, items]) => ({ name, items }));
  } catch(e) { return []; }
}

function createChecklistV11(ss) {
  let sheet = ss.getSheetByName(SHEET_ITEMS);
  if (!sheet) sheet = ss.insertSheet(SHEET_ITEMS);
  sheet.clear();
  sheet.appendRow(['Khu vực', 'Hạng mục kiểm tra', 'Mã HM', 'Nội dung rà soát']);
  const items = [
    ['TRƯỚC CH', 'Biển hiệu WM+', 'CSVC01', 'Kiểm tra biển hiệu có sáng đầy đủ; bóng LED có cháy/nhấp nháy; mặt biển có bạc màu, nứt, bong tróc; chữ/logo có thiếu hoặc hỏng'],
    ['TRƯỚC CH', 'Mặt tiền/ sơn tường ngoài', 'CSVC02', 'Kiểm tra sơn mặt tiền có bong tróc, nứt, ố mốc, bạc màu; tường có thấm nước hoặc rêu mốc'],
    ['TRƯỚC CH', 'Cửa cuốn/ cửa kính', 'CSVC03', 'Kiểm tra cửa cuốn đóng/mở bình thường; cửa có kẹt, lệch, han gỉ, móp méo; cửa kính có nứt/vỡ/xước'],
    ['TRƯỚC CH', 'Mái hiên/ mái che', 'CSVC04', 'Kiểm tra mái hiên/mái che có dột, thấm, võng, nứt hoặc han gỉ; khung mái có chắc chắn'],
    ['TRƯỚC CH', 'Đèn chiếu sáng ngoài', 'CSVC05', 'Kiểm tra đèn ngoài CH: đèn có cháy, yếu, chập chờn; mặt tiền, lối vào, biển hiệu có đủ sáng'],
    ['TRƯỚC CH', 'Vỉa hè', 'CSVC06', 'Kiểm tra vỉa hè có nứt, vỡ, sụt lún, kênh hoặc đọng nước; có vật cản gây khó khăn cho KH'],
    ['TRƯỚC CH', 'Biển quảng cáo/Poster', 'CSVC07', 'Kiểm tra poster/decal/biển quảng cáo có cũ, rách, bong mép, bạc màu; nội dung đã hết CTKM chưa'],
    ['TRƯỚC CH', 'Cây xanh', 'CSVC08', 'Kiểm tra cây xanh có khô/héo, chết, mọc um tùm hoặc che biển hiệu/lối đi; chậu cây có vỡ, bẩn'],
    ['TRƯỚC CH', 'Thùng rác ngoài', 'CSVC09', 'Kiểm tra thùng rác có đủ số lượng, nắp/túi, bánh xe; có nứt vỡ, rò rỉ, bẩn hoặc có mùi'],
    ['TRƯỚC CH', 'Camera an ninh ngoài', 'CSVC10', 'Kiểm tra camera ngoài CH có hoạt động; góc quay có bị che khuất; camera có lỏng, nghiêng, hỏng, mất tín hiệu'],
    ['TRONG CH', 'Trần nhà/ thạch cao', 'CSVC11', 'Kiểm tra trần có dột/thấm, ố vàng, mốc, nứt, bong hoặc võng; tấm thạch cao có vỡ/mất'],
    ['TRONG CH', 'Tường', 'CSVC12', 'Kiểm tra tường trong CH có bong sơn, nứt, ố, mốc, thấm nước hoặc va đập; chân tường có hư hỏng'],
    ['TRONG CH', 'Đèn chiếu sáng', 'CSVC13', 'Kiểm tra đèn trong CH: số lượng đèn cháy/yếu/chập chờn; khu vực bán hàng có bị tối'],
    ['TRONG CH', 'Điều hòa/ quạt/ thông gió', 'CSVC14', 'Kiểm tra điều hòa/quạt/thông gió có hoạt động bình thường; khả năng làm mát có đảm bảo'],
    ['TRONG CH', 'Kệ trưng bày', 'CSVC15', 'Kiểm tra kệ có cong, gãy, han gỉ, bong sơn hoặc thiếu chi tiết; chân kệ có chắc chắn'],
    ['TRONG CH', 'Tủ mát/ tủ đông/ tủ lạnh', 'CSVC16', 'Kiểm tra tủ mát/tủ đông/tủ lạnh có hoạt động ổn định; nhiệt độ có đảm bảo'],
    ['TRONG CH', 'Bảng giá', 'CSVC17', 'Kiểm tra bảng giá có đầy đủ và đúng vị trí; bảng có cũ, cong, bẩn hoặc hư hỏng'],
    ['TRONG CH', 'Giỏ hàng/ xe đẩy', 'CSVC18', 'Kiểm tra giỏ hàng/xe đẩy có đủ số lượng; bánh xe có kẹt/lệch; tay cầm có hỏng'],
    ['TRONG CH', 'Quầy thu ngân', 'CSVC19', 'Kiểm tra quầy thu ngân có hư hỏng, bong tróc, nứt hoặc xuống cấp; mặt bàn, ngăn kéo có chắc chắn'],
    ['TRONG CH', 'Cân điện tử', 'CSVC20', 'Kiểm tra cân điện tử có hoạt động và hiển thị chính xác; màn hình/phím bấm có lỗi'],
    ['TRONG CH', 'Wifi/ Mạng/ Thiết bị IT', 'CSVC21', 'Kiểm tra Wifi/mạng có ổn định; modem/router/switch có hoạt động'],
    ['TRONG CH', 'Hạng mục khác', 'CSVC22', 'Rà soát các hạng mục CSVC khác chưa có trong checklist: biển chỉ dẫn, giá treo, tay vịn, phụ kiện, vật dụng cố định'],
    ['TRONG CH', 'Hệ thống âm thanh', 'CSVC23', 'Kiểm tra loa và hệ thống âm thanh có hoạt động; âm lượng có rõ và đồng đều'],
    ['TRONG CH', 'Bình chữa cháy/ PCCC', 'CSVC24', 'Kiểm tra bình chữa cháy có đủ số lượng và đúng vị trí; bình có còn hạn/tem kiểm định'],
    ['TRONG CH', 'Ổ điện, hệ thống điện', 'CSVC25', 'Kiểm tra ổ cắm, công tắc, tủ điện và dây điện có cháy xém, nứt vỡ, lỏng hoặc quá tải'],
    ['TRONG CH', 'Camera an ninh trong', 'CSVC26', 'Kiểm tra camera trong CH có hoạt động và ghi hình; hình ảnh có rõ'],
    ['KHO - NVS', 'Cửa kho / cửa NVS', 'CSVC27', 'Kiểm tra cửa kho/cửa NVS có đóng mở bình thường; khóa, bản lề, tay nắm có hỏng'],
    ['KHO - NVS', 'Kệ kho/ Palet', 'CSVC28', 'Kiểm tra kệ kho/palet có đủ chắc chắn; có cong, gãy, han gỉ hoặc mối mọt'],
    ['KHO - NVS', 'Sàn kho / NVS', 'CSVC29', 'Kiểm tra sàn kho/NVS có nứt, vỡ, sụt lún, bong gạch, đọng nước hoặc trơn trượt'],
    ['KHO - NVS', 'Đường điện', 'CSVC30', 'Kiểm tra đường điện trong kho/NVS có dây hở, dây chắp nối, ổ cắm lỏng hoặc đi dây mất an toàn'],
    ['KHO - NVS', 'Thiết bị vệ sinh', 'CSVC31', 'Kiểm tra vòi nước, lavabo, bồn rửa, van khóa, vòi xịt và các thiết bị liên quan có hoạt động'],
    ['KHO - NVS', 'Đường nước', 'CSVC32', 'Kiểm tra đường cấp/thoát nước có rò rỉ, tắc nghẽn, nứt vỡ hoặc mùi bất thường'],
    ['KHO - NVS', 'Thiết bị nhà vệ sinh', 'CSVC33', 'Kiểm tra bồn cầu, lavabo, vòi xịt, gương, móc treo, đèn, quạt thông gió và phụ kiện NVS'],
    ['KHO - NVS', 'Sàn kho / NVS (tổng thể)', 'CSVC34', 'Kiểm tra tổng thể sàn kho/NVS tại các vị trí có nguy cơ xuống cấp: nứt, vỡ, bong, lún, thấm, đọng nước']
  ];
  for (const row of items) sheet.appendRow(row);
  sheet.setFrozenRows(1); sheet.autoResizeColumns(1, 4);
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
        if (photo && photo.startsWith && photo.startsWith('data:')) {
          try {
            const blob = dataUrlToBlob(photo);
            const safeFileName = sanitizeFileName(`${data.store}_${String(item.item || '').replace(/[^a-z0-9]/gi, '_')}_${i + 1}.jpg`);
            const file = folder.createFile(blob.setName(safeFileName));
            photoUrls.push(file.getUrl()); savedPhotos++;
          } catch(pe) { photoUrls.push('Lỗi: ' + pe.toString().substring(0, 50)); }
        } else if (photo && typeof photo === 'string') { photoUrls.push(photo); }
      }
    }
    const row = [
      new Date().toLocaleString('vi-VN'), data.store, data.storeName || '', data.tinh || '', data.gdv || '', data.qlkv || '',
      item.area || '', item.item || '', item.result || '', item.note || '', photoUrls.join('\n'),
      item.code || '', item.date || '', item.su || '', item.sudate || ''
    ];
    sheet.appendRow(row);
    results.push({ item: item.item, result: item.result, photos: photoUrls.length });
  }
  return { ok: true, timestamp: new Date().toISOString(), store: data.store, totalItems: data.totalItems, passCount: data.passCount, failCount: data.failCount, savedPhotos: savedPhotos, rowsSaved: results.length };
}

function reloadChecklist() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = createChecklistV11(ss);
  return { ok: true, message: 'Đã cập nhật HANG_MUC v11 (34 hạng mục)', rowCount: sheet.getLastRow() - 1 };
}

// ── HELPERS ──
function getOrCreateSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(['Ngày báo', 'Mã CH', 'Tên CH', 'Tỉnh', 'GĐV', 'QLKV', 'Khu vực', 'Hạng mục', 'Đánh giá', 'Diễn giải', 'Ảnh Drive', 'Mã theo dõi', 'Ngày PS', 'Đã báo SU', 'Ngày báo SU']);
    sheet.setFrozenRows(1); sheet.autoResizeColumns(1, 15);
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

function sanitizeFileName(name) { return name.replace(/[\\/:*?"<>|]/g, '_').substring(0, 200); }
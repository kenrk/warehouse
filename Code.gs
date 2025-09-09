/**
 * Script Otomatis untuk Manajemen Inventaris dan QC
 * Fitur: Autofill Supplies, Validasi Produksi, QC, Sinkronisasi Data + Dedupe
 * Timezone: Asia/Jakarta
 * @version 4.0 (Installable Triggers + Strong Validation + Dedupe)
 */

// =========================================================================
//                          ⭐ KONFIGURASI GLOBAL ⭐
// =========================================================================
const TZ = 'Asia/Jakarta';
const START_ROW = 2;

// Nama Sheet
const SHEET_SUPPLIES   = 'Supplies';
const SHEET_PN         = 'PN';
const SHEET_PRODUCTION = 'Production';
const SHEET_QC         = 'QC';        // UBAH ke 'NG' bila tab kamu bernama NG
const SHEET_NG_PARAM   = 'NG_Param';
const SHEET_DASHBOARD  = 'Dashboard';

// Definisi Kolom (1-based)
const COL_SUP  = { QR: 1, PARTNO: 2, PARTNAME: 3, BATCH: 4, DATE: 5, TIME: 6, OP: 7, QR_CODE: 8 };
const COL_PROD = { QR: 1, DATE: 2, TIME: 3, OP: 4, LINE: 5, FIRST_SEEN: 6, SCAN_COUNT: 7 };
const COL_QC   = { QR: 1, DATE: 2, TIME: 3, OP: 4, STATUS: 5, JENIS_NG: 6, DETAIL_NG: 7 };

// =========================================================================
//                        ⭐ FUNGSI PEMICU (INSTALLABLE TRIGGERS) ⭐
// =========================================================================

/** Trigger installable untuk edit sel dengan logging lengkap */
function onEditTrigger(e) {
  try {
    // Log 1: Memastikan event object diterima dengan benar
    console.log(`onEditTrigger dimulai. Informasi event: ${e ? 'Ada' : 'Tidak Ada'}. Informasi range: ${e && e.range ? e.range.getA1Notation() : 'Tidak Ada'}`);

    if (!e || !e.range) {
      console.log("PROSES BERHENTI: Event object atau range tidak lengkap.");
      return; // Keluar dari fungsi
    }

    const sh = e.range.getSheet();
    const sheetName = sh.getName();
    const col = e.range.getColumn();
    const row = e.range.getRow();

    // Log 2: Mencatat detail editan
    console.log(`Edit terdeteksi di sheet: "${sheetName}", Sel: ${e.range.getA1Notation()}, Kolom: ${col}`);

    if (sheetName === SHEET_SUPPLIES) {
      console.log('Logika dijalankan untuk sheet SUPPLIES.');
      if (col === COL_SUP.QR) {
        console.log('Memanggil fungsi handleSuppliesEdit_...');
        const valid = handleSuppliesEdit_(e, sh);
        // Logika sync dari script asli Anda
        if (valid === true || String(e.value || '').trim() === '') {
          syncAllSheets();
        }
      } else {
        console.log(`PROSES DIABAIKAN: Edit di sheet Supplies, tetapi bukan di kolom QR (Kolom ke-${col}).`);
      }
    } else if (sheetName === SHEET_PRODUCTION) {
      console.log('Logika dijalankan untuk sheet PRODUCTION.');
      handleProductionEdit_(e, sh);
      syncQCWithProduction();
    } else if (sheetName === SHEET_QC) {
      console.log('Logika dijalankan untuk sheet QC.');
      handleQCEdit_(e, sh);
    } else {
      console.log(`PROSES DIABAIKAN: Edit terjadi di sheet "${sheetName}" yang tidak dipantau.`);
    }

  } catch (err) {
    // Log 5: Menangkap error tak terduga
    console.error(`ERROR KRITIS di onEditTrigger: ${err.message}`);
  }
}

/** Trigger installable untuk perubahan struktur (hapus baris, dll.) */
function onChangeTrigger(e) {
  try {
    if (!e || !e.changeType) return;
    const t = String(e.changeType);
    if (t === 'REMOVE_ROW' || t === 'INSERT_ROW' || t === 'PASTE') {
      console.log(`${t} terdeteksi → syncAllSheets()`);
      syncAllSheets();
    }
  } catch (err) {
    console.error(`Error in onChangeTrigger: ${err.message}\nStack: ${err.stack}`);
  }
}

/** Menu kustom */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚙️ Custom Menu')
    .addItem('Update Dashboard', 'generateDashboard')
    .addSeparator()
    .addItem('▶️ Sinkronisasi Manual Lengkap', 'forceSyncAllSheets')
    .addToUi();
}

// =========================================================================
//                     ⭐ FUNGSI SINKRONISASI DATA ⭐
// =========================================================================

/** Sinkronisasi lengkap: Supplies → Production → QC (dedupe + orphan cleanup) */
function syncAllSheets() {
  try {
    const ss = SpreadsheetApp.getActive();
    const suppliesSheet   = ss.getSheetByName(SHEET_SUPPLIES);
    const productionSheet = ss.getSheetByName(SHEET_PRODUCTION);
    const qcSheet         = ss.getSheetByName(SHEET_QC);

    if (!suppliesSheet || !productionSheet || !qcSheet) {
      console.warn('syncAllSheets: Pastikan Supplies/Production/QC ada semua.');
      return;
    }

    // 1) Bersihkan duplikat di Supplies (sisakan kemunculan pertama)
    const cleaned = dedupeSheetByQR_(suppliesSheet, COL_SUP.QR);
    if (cleaned) console.log(`🧹 Supplies: ${cleaned} baris duplikat dihapus`);

    // 2) Production mengikuti Supplies
    const suppliesQRs = getValidQRsFromSheet_(suppliesSheet, COL_SUP.QR);
    const productionRemoved = removeInvalidDataFromSheet_(productionSheet, COL_PROD.QR, suppliesQRs);

    // 3) QC mengikuti Production
    const productionQRs = getValidQRsFromSheet_(productionSheet, COL_PROD.QR);
    const qcRemoved = removeInvalidDataFromSheet_(qcSheet, COL_QC.QR, productionQRs);

    console.log(`✅ Sync: Supplies(dedupe ${cleaned}), Production(-${productionRemoved}), QC(-${qcRemoved})`);
    if (cleaned || productionRemoved || qcRemoved) {
      SpreadsheetApp.getActiveSpreadsheet()
        .toast(`Sync: Supplies(dedupe ${cleaned}), Production(-${productionRemoved}), QC(-${qcRemoved})`);
    }
  } catch (e) {
    console.error(`syncAllSheets error: ${e.message}\n${e.stack}`);
  }
}

/** Sinkronisasi Production dari Supplies (dipakai juga di onEditTrigger) */
function syncProductionWithSupplies() {
  try {
    const ss = SpreadsheetApp.getActive();
    const suppliesSheet   = ss.getSheetByName(SHEET_SUPPLIES);
    const productionSheet = ss.getSheetByName(SHEET_PRODUCTION);
    if (!suppliesSheet || !productionSheet) return 0;

    const suppliesQRs = getValidQRsFromSheet_(suppliesSheet, COL_SUP.QR);
    const productionRemoved = removeInvalidDataFromSheet_(productionSheet, COL_PROD.QR, suppliesQRs);
    if (productionRemoved > 0) console.log(`✅ ${productionRemoved} data Production disinkronkan`);
    return productionRemoved;
  } catch (e) {
    console.error(`Error sync Production: ${e.message}`); 
    return 0;
  }
}

/** Sinkronisasi QC dari Production */
function syncQCWithProduction() {
  try {
    const ss = SpreadsheetApp.getActive();
    const productionSheet = ss.getSheetByName(SHEET_PRODUCTION);
    const qcSheet         = ss.getSheetByName(SHEET_QC);
    if (!productionSheet || !qcSheet) return 0;

    const productionQRs = getValidQRsFromSheet_(productionSheet, COL_PROD.QR);
    const qcRemoved = removeInvalidDataFromSheet_(qcSheet, COL_QC.QR, productionQRs);
    if (qcRemoved > 0) console.log(`✅ ${qcRemoved} data QC disinkronkan`);
    return qcRemoved;
  } catch (e) {
    console.error(`Error sync QC: ${e.message}`); 
    return 0;
  }
}

/** Ambil QR valid dari sheet (Set) */
function getValidQRsFromSheet_(sheet, qrColumn) {
  const lastRow = sheet.getLastRow();
  if (lastRow < START_ROW) return new Set();
  return new Set(
    sheet.getRange(START_ROW, qrColumn, lastRow - START_ROW + 1, 1)
         .getValues()
         .flat()
         .map(qr => String(qr).trim())
         .filter(Boolean)
  );
}

/** Hapus data invalid dari sheet (baris yang QR-nya tidak ada di Set valid) */
function removeInvalidDataFromSheet_(sheet, qrColumn, validQRs) {
  const lastRow = sheet.getLastRow();
  if (lastRow < START_ROW) return 0;

  const qrValues = sheet.getRange(START_ROW, qrColumn, lastRow - START_ROW + 1, 1).getValues();
  const rowsToRemove = [];

  qrValues.forEach((row, index) => {
    const qr = String(row[0]).trim();
    if (qr && !validQRs.has(qr)) rowsToRemove.push(START_ROW + index);
  });

  if (rowsToRemove.length > 0) {
    rowsToRemove.sort((a, b) => b - a).forEach(rowIndex => sheet.deleteRow(rowIndex));
  }
  return rowsToRemove.length;
}

/** Hapus duplikat QR pada sheet (sisakan kemunculan pertama) */
function dedupeSheetByQR_(sheet, qrColumn) {
  const lastRow = sheet.getLastRow();
  if (lastRow < START_ROW) return 0;

  const values = sheet.getRange(START_ROW, qrColumn, lastRow - START_ROW + 1, 1).getValues();
  const seen = new Set();
  const toDelete = [];

  values.forEach((r, i) => {
    const qr = String(r[0]).trim();
    if (!qr) return;
    const rowNum = START_ROW + i;
    if (seen.has(qr)) toDelete.push(rowNum);
    else seen.add(qr);
  });

  if (toDelete.length) {
    toDelete.sort((a,b)=>b-a).forEach(r => sheet.deleteRow(r));
  }
  return toDelete.length;
}

// =========================================================================
//                   ⭐ FUNGSI HANDLER UNTUK SETIAP SHEET ⭐
// =========================================================================

/** Handler untuk sheet Supplies */
function handleSuppliesEdit_(e, sh) {
  try {
    // --- PERBAIKAN UTAMA DI SINI ---
    // Perintah ini memaksa Google untuk menyelesaikan semua operasi tulis
    // sebelum script melanjutkan. Ini menyelesaikan masalah waktu (timing issue).
    SpreadsheetApp.flush(); 
    
    const row = e.range.getRow();
    const col = e.range.getColumn();
    if (row < START_ROW || col !== COL_SUP.QR) return false;

    // Sekarang aman untuk membaca nilai langsung dari sel karena sudah di-flush
    const qr = String(e.range.getValue() || '').trim();
    
    console.log(`(1) Memproses edit di baris ${row} untuk QR: "${qr}"`);

    if (!qr) {
      console.log("(2) QR kosong, membersihkan baris.");
      clearRowContent_(sh, row, sh.getLastColumn());
      return false;
    }

    if (isDuplicateQR_(sh, row, qr, COL_SUP.QR)) {
      console.log("(E) Gagal: QR duplikat.");
      e.range.clearContent();
      SpreadsheetApp.getActiveSpreadsheet().toast(`QR "${qr}" duplikat, dibersihkan`);
      return false;
    }

    const parsed = parseQR_(qr);
    console.log(`(2) Hasil parsing QR: ${JSON.stringify(parsed)}`);
    if (!parsed.sn || !parsed.partNo || !parsed.batch) {
      console.log("(E) Gagal: Format QR tidak valid.");
      clearRowContent_(sh, row, sh.getLastColumn());
      SpreadsheetApp.getActiveSpreadsheet().toast(`Format QR tidak valid: ${qr}`);
      return false;
    }

    const pnData = lookupPN_(parsed.sn);
    console.log(`(3) Hasil lookup dari PN untuk SN "${parsed.sn}": ${JSON.stringify(pnData)}`);
    if (!pnData || !pnData.partName) {
      console.log("(E) Gagal: SN tidak ditemukan di sheet PN.");
      clearRowContent_(sh, row, sh.getLastColumn());
      SpreadsheetApp.getActiveSpreadsheet().toast(`SN "${parsed.sn}" tidak ditemukan di PN`);
      return false;
    }

    console.log("(4) Semua validasi berhasil. Memulai penulisan data baris...");
    writeSuppliesRow_(sh, row, pnData, parsed);
    console.log("(5) Penulisan data baris selesai.");
    return true; // valid
  } catch (err) {
    console.error(`Terjadi error tak terduga di handleSuppliesEdit_: ${err.message}`);
    SpreadsheetApp.getActiveSpreadsheet().toast(`Error: ${err.message}`);
  }
}

/** Handler untuk sheet Production */
function handleProductionEdit_(e, sh) {
  const row = e.range.getRow();
  const col = e.range.getColumn();
  if (row < START_ROW || col !== COL_PROD.QR) return;

  const qr = String(e.range.getValue() || '').trim();
  if (!qr) {
    clearRowContent_(sh, row, sh.getLastColumn());
    return;
  }

  if (isDuplicateQR_(sh, row, qr, COL_PROD.QR)) {
    e.range.clearContent();
    SpreadsheetApp.getActiveSpreadsheet().toast(`QR "${qr}" sudah discan di Production`);
    return;
  }

  if (!isQRExistsInSupplies_(qr)) {
    e.range.clearContent();
    SpreadsheetApp.getActiveSpreadsheet().toast(`QR "${qr}" tidak terdaftar di Supplies`);
    return;
  }

  const { date, time } = now_();
  sh.getRange(row, COL_PROD.DATE, 1, 3).setValues([[date, time, getEmail_()]]);

  const scanCountCell = sh.getRange(row, COL_PROD.SCAN_COUNT);
  const scanCount = scanCountCell.getValue() || 0;
  if (scanCount === 0) {
    sh.getRange(row, COL_PROD.FIRST_SEEN).setValue(date);
  }
  scanCountCell.setValue(scanCount + 1);
}

/** Handler untuk sheet QC */
function handleQCEdit_(e, sh) {
  const row = e.range.getRow();
  const col = e.range.getColumn();
  if (row < START_ROW) return;

  if (col === COL_QC.QR) {
    const qr = String(e.range.getValue() || '').trim();
    if (!qr) {
      clearRowContent_(sh, row, sh.getLastColumn(), true);
      return;
    }

    if (isDuplicateQR_(sh, row, qr, COL_QC.QR)) {
      e.range.clearContent();
      SpreadsheetApp.getActiveSpreadsheet().toast(`QR "${qr}" sudah ada di QC`);
      return;
    }

    if (!isQRExistsInProduction_(qr)) {
      e.range.clearContent();
      SpreadsheetApp.getActiveSpreadsheet().toast(`QR "${qr}" tidak ditemukan di Production`);
      return;
    }

    const { date, time } = now_();
    sh.getRange(row, COL_QC.DATE, 1, 3).setValues([[date, time, getEmail_()]]);
    setStatusValidation_(sh, row);
  }

  if (col === COL_QC.STATUS) {
    const status = String(e.range.getValue() || '').trim().toUpperCase();
    const jenisNgCell  = sh.getRange(row, COL_QC.JENIS_NG);
    const detailNgCell = sh.getRange(row, COL_QC.DETAIL_NG);

    if (status === 'NG') {
      setJenisNGValidation_(sh, row);
      detailNgCell.clearContent().clearDataValidations();
    } else {
      jenisNgCell.clearContent().clearDataValidations();
      detailNgCell.clearContent().clearDataValidations();
    }
  }

  if (col === COL_QC.JENIS_NG) {
    const jenisNG = String(e.range.getValue() || '').trim();
    const detailNgCell = sh.getRange(row, COL_QC.DETAIL_NG);
    detailNgCell.clearContent();

    if (jenisNG) setDetailNGValidation_(sh, row, jenisNG);
    else detailNgCell.clearDataValidations();
  }
}

// =========================================================================
//                        ⭐ FUNGSI PEMBANTU (HELPERS) ⭐
// =========================================================================

/** True bila baris saat ini duplikat; sekaligus hapus duplikat lain */
function isDuplicateQR_(sheet, currentRow, qr, qrColumn) {
  const lastRow = sheet.getLastRow();
  if (lastRow < START_ROW) return false;

  const rng = sheet.getRange(START_ROW, qrColumn, lastRow - START_ROW + 1, 1).getValues();
  const dupRows = [];
  let firstRow = null;

  rng.forEach((r, i) => {
    const rowNum = START_ROW + i;
    if (String(r[0]).trim() === qr) {
      if (firstRow === null) firstRow = rowNum;
      else dupRows.push(rowNum);
    }
  });

  if (dupRows.length) dupRows.sort((a,b)=>b-a).forEach(r => sheet.deleteRow(r));
  return firstRow !== null && currentRow !== firstRow;
}

function isQRExistsInSupplies_(qr) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_SUPPLIES);
  if (!sh || sh.getLastRow() < START_ROW) return false;
  const qrValues = sh.getRange(START_ROW, COL_SUP.QR, sh.getLastRow() - START_ROW + 1, 1).getValues();
  return qrValues.some(row => String(row[0]).trim() === qr);
}

function isQRExistsInProduction_(qr) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_PRODUCTION);
  if (!sheet || sheet.getLastRow() < START_ROW) return false;
  const qrValues = sheet.getRange(START_ROW, COL_PROD.QR, sheet.getLastRow() - START_ROW + 1, 1).getValues();
  return qrValues.some(row => String(row[0]).trim() === qr);
}

/** Parse QR: 4 char SN + 4 char PartNo + 2 char Batch (+ optional suffix) */
function parseQR_(qr) {
  const s = String(qr).trim().toUpperCase();
  const m = s.match(/^([A-Z0-9]{4})([A-Z0-9]{4})([A-Z0-9]{2})([A-Z0-9]+)?$/);
  if (!m) return {};
  return { sn: m[1], partNo: m[2], batch: m[3] };
}

/** Wajib match di PN (tanpa fallback) */
function lookupPN_(sn) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_PN);
  if (!sh || sh.getLastRow() < 2) return null;
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues(); // A:D = SN, PartNo, PartName, Batch?
  const target = String(sn).toUpperCase().trim();
  for (const row of data) {
    if (String(row[0]).toUpperCase().trim() === target) {
      return {
        partNo:  String(row[1] || ''),
        partName:String(row[2] || ''),
        batch:   String(row[3] || '')
      };
    }
  }
  return null; // tidak ada di PN = tolak
}

function writeSuppliesRow_(sheet, row, pnData, parsedQR) {
  const { date, time } = now_();
  const values = [ parsedQR.partNo, pnData.partName, parsedQR.batch, date, time, getEmail_() ];
  sheet.getRange(row, COL_SUP.PARTNO, 1, 6).setValues([values]);
  generateQRCodeForRow_(sheet, row);
}

function setStatusValidation_(sheet, row) {
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['OK', 'NG'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(row, COL_QC.STATUS).setDataValidation(rule);
}

function setJenisNGValidation_(sheet, row) {
  const ngParamSheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NG_PARAM);
  if (!ngParamSheet) return;
  const headers = ngParamSheet.getRange(1, 1, 1, ngParamSheet.getLastColumn()).getValues()[0].filter(String);
  if (headers.length > 0) {
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(headers, true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(row, COL_QC.JENIS_NG).setDataValidation(rule);
  }
}

function setDetailNGValidation_(sheet, row, jenisNG) {
  const ngParamSheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NG_PARAM);
  if (!ngParamSheet) return;
  const colIndex = getColumnIndexByHeader_(ngParamSheet, jenisNG);
  if (colIndex <= 0) return;
  const detailValues = ngParamSheet.getRange(2, colIndex, ngParamSheet.getLastRow()).getValues().flat().filter(String);
  if (detailValues.length > 0) {
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(detailValues, true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(row, COL_QC.DETAIL_NG).setDataValidation(rule);
  }
}

function clearRowContent_(sheet, row, numColumns, clearValidations = false) {
  const range = sheet.getRange(row, 1, 1, numColumns);
  range.clearContent();
  if (clearValidations) range.clearDataValidations();
}

function generateQRCodeForRow_(sheet, row, qrSize = 150) {
  try {
    const qrId = sheet.getRange(row, COL_SUP.QR).getValue();
    if (qrId) {
      const qrCodeUrl = `https://quickchart.io/qr?text=${encodeURIComponent(qrId)}&size=${qrSize}`;
      sheet.getRange(row, COL_SUP.QR_CODE).setFormula(`=IMAGE("${qrCodeUrl}")`);
    }
  } catch (e) {
    console.error(`Error generating QR code: ${e.message}`);
  }
}

function getColumnIndexByHeader_(sheet, headerText) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return headers.indexOf(headerText) + 1;
}

function now_() {
  const d = new Date();
  return {
    date: Utilities.formatDate(d, TZ, 'M/d/yyyy'),
    time: Utilities.formatDate(d, TZ, 'h:mm:ss a')
  };
}

function getEmail_() {
  try {
    // Menggunakan getEffectiveUser() lebih andal untuk web app
    // Ini akan mengembalikan email pemilik script (Anda)
    return Session.getEffectiveUser().getEmail() || '';
  }
  catch (e) {
    return '';
  }
}

// =========================================================================
//                      ⭐ FUNGSI DASHBOARD ⭐
// =========================================================================

function generateDashboard() {
  console.log('=== START GENERATE DASHBOARD ===');
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dashboardSheet = ss.getSheetByName(SHEET_DASHBOARD);
    const suppliesSheet  = ss.getSheetByName(SHEET_SUPPLIES);
    if (!dashboardSheet || !suppliesSheet) return;

    dashboardSheet.clear();
    const headers = [['Date', 'Part_Name', 'Supplies', 'Production', 'QC_OK', 'QC_NG']];
    dashboardSheet.getRange(1, 1, 1, 6).setValues(headers).setFontWeight('bold');

    const suppliesData = suppliesSheet.getDataRange().getValues();
    if (suppliesData.length <= 1) return;

    const suppliesValues = suppliesData.slice(1);
    const resultArray = [];

    for (let i = 0; i < suppliesValues.length; i++) {
      const row = suppliesValues[i];
      const qr = String(row[0] || '').trim();
      const partName = String(row[2] || '').trim();
      const dateCell = row[4];
      if (!qr || !partName || !dateCell) continue;

      let date;
      if (dateCell instanceof Date) date = Utilities.formatDate(dateCell, TZ, 'M/d/yyyy');
      else date = String(dateCell);

      const suppliesCount   = 1;
      const productionCount = countInSheet(SHEET_PRODUCTION, qr, 0);
      const qcOKCount       = countInSheet(SHEET_QC, qr, 4, 'OK');
      const qcNGCount       = countInSheet(SHEET_QC, qr, 4, 'NG');

      resultArray.push([date, partName, suppliesCount, productionCount, qcOKCount, qcNGCount]);
    }

    if (resultArray.length > 0) {
      dashboardSheet.getRange(2, 1, resultArray.length, 6).setValues(resultArray);
      dashboardSheet.autoResizeColumns(1, 6);
      console.log('✅ Dashboard berhasil di-generate dengan', resultArray.length, 'baris data');
    }
  } catch (error) {
    console.error('Error in generateDashboard:', error.toString());
  }
  console.log('=== END GENERATE DASHBOARD ===');
}

/** Helper hitung pada sheet lain */
function countInSheet(sheetName, targetQR, statusColumn = 0, statusValue = '') {
  try {
    const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
    if (!sheet) return 0;
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return 0;

    let count = 0;
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const qr = String(row[0] || '').trim();
      if (qr !== targetQR) continue;
      if (statusColumn === 0) count++;
      else if (statusColumn > 0 && statusValue) {
        const status = String(row[statusColumn] || '').trim().toUpperCase();
        if (status === statusValue.toUpperCase()) count++;
      }
    }
    return count;
  } catch (error) {
    console.error('Error in countInSheet:', error.toString());
    return 0;
  }
}

// =========================================================================
// ⭐ FUNGSI MANUAL SYNC (menu) & SETUP TRIGGERS
// =========================================================================

function forceSyncAllSheets() {
  syncAllSheets();
  SpreadsheetApp.getActiveSpreadsheet().toast('Forced sync dijalankan');
}

function setupTriggersOnce() {
  const ssId = SpreadsheetApp.getActive().getId();
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('onEditTrigger').forSpreadsheet(ssId).onEdit().create();
  ScriptApp.newTrigger('onChangeTrigger').forSpreadsheet(ssId).onChange().create();
}

// =========================================================================
//                  ⭐ FUNGSI UTAMA WEB APP (Versi 2.1 - Clean) ⭐
// =========================================================================

/**
 * Fungsi utama untuk menjalankan Web App.
 * Ini adalah SATU-SATUNYA doGet yang harus ada di proyek Anda.
 */
function doGet(e) {
  // Fungsi ini memanggil file WebApp.html sebagai template
  return HtmlService.createTemplateFromFile('WebApp')
      .evaluate()
      .setTitle('Warehouse Scanner')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, user-scalable=no');
}

/**
 * Helper untuk menyisipkan konten file CSS/JS ke dalam HTML utama.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Helper untuk menyisipkan konten CSS/JS (jika kita memisahkannya nanti).
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Mengambil data awal saat Web App pertama kali dimuat.
 * Tujuannya adalah untuk mendapatkan jumlah baris terakhir di sheet Production.
 * @returns {Object} - Objek berisi jumlah baris terakhir.
 */
function getInitialData() {
  try {
    const productionSheet = SpreadsheetApp.getActive().getSheetByName(SHEET_PRODUCTION);
    return {
      lastRow: productionSheet.getLastRow()
    };
  } catch (err) {
    return {
      lastRow: 0,
      error: err.message
    };
  }
}

/**
 * Fungsi yang akan dipanggil secara berkala oleh Web App (polling).
 * Mengecek apakah ada baris baru YANG VALID di sheet Production.
 * @param {number} lastKnownRowCount - Jumlah baris yang terakhir diketahui oleh client.
 * @returns {Object} - Objek JSON yang berisi status scan baru dan datanya.
 */
function getLatestProductionScan(lastKnownRowCount) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_PRODUCTION);
  const currentLastRow = sheet.getLastRow();

  // Jika jumlah baris sekarang lebih besar dari yang diketahui sebelumnya, berarti ada potensi scan baru
  if (currentLastRow > lastKnownRowCount) {
    const newQrValue = String(sheet.getRange(currentLastRow, COL_PROD.QR).getValue() || '').trim();

    // --- VALIDASI TAMBAHAN ---
    // Sebelum mengirim sinyal "Production Done", kita cek dulu apakah QR yang baru
    // masuk ini valid (terdaftar di sheet Supplies).
    // Kita gunakan lagi fungsi yang sudah ada: isQRExistsInSupplies_()
    if (newQrValue && isQRExistsInSupplies_(newQrValue)) {
      // HANYA JIKA VALID, kirim sinyal scan baru ke web app
      return {
        newScan: true,
        qr: newQrValue,
        newRowCount: currentLastRow
      };
    }
    // Jika tidak valid, kita tidak melakukan apa-apa dan akan lanjut ke return default di bawah.
    // Ini mencegah web app menampilkan "Production Done" untuk data yang salah.
  }

  // Jika tidak ada baris baru ATAU baris baru ternyata tidak valid, kirim status false.
  return {
    newScan: false,
    newRowCount: sheet.getLastRow() // Selalu kirim jumlah baris terkini untuk sinkronisasi
  };
}

// =========================================================================
//                  ⭐ FUNGSI UNTUK WEB APP INTERFACE QC ⭐
// =========================================================================

/**
 * Fungsi untuk menampilkan halaman khusus QC.
 * Kita akan memanggilnya dengan parameter URL, contoh: .../exec?page=qc
 */
// function doGet(e) {
//   if (e.parameter.page == 'qc') {
//     return HtmlService.createTemplateFromFile('Index_QC')
//       .evaluate()
//       .setTitle('QC Scanner Interface')
//       .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
//   }
//   // Jika tidak ada parameter 'page', tampilkan halaman Produksi default
//   return HtmlService.createTemplateFromFile('Index')
//     .evaluate()
//     .setTitle('Production Scanner Interface')
//     .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
// }

/**
 * Mengecek apakah ada baris baru di sheet QC.
 * @param {number} lastKnownRowCount - Jumlah baris yang terakhir diketahui oleh client.
 * @returns {Object} - Objek JSON yang berisi status scan baru dan datanya.
 */
function getLatestQCScan(lastKnownRowCount) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_QC);
  const currentLastRow = sheet.getLastRow();

  if (currentLastRow > lastKnownRowCount) {
    const newQrValue = sheet.getRange(currentLastRow, COL_QC.QR).getValue();
    // Validasi: pastikan QR ada di Production sebelum ditampilkan di UI QC
    if (newQrValue && isQRExistsInProduction_(newQrValue)) {
       return {
        newScan: true,
        qr: newQrValue,
        row: currentLastRow, // Kirim nomor baris untuk diupdate nanti
        newRowCount: currentLastRow
      };
    }
  }
  return {
    newScan: false,
    newRowCount: sheet.getLastRow()
  };
}

/**
 * Mengupdate status di sheet QC berdasarkan input dari Web App.
 * @param {number} row - Nomor baris yang akan diupdate.
 * @param {string} status - Status yang diinput ('OK' atau 'NG').
 * @returns {Object} - Objek berisi status keberhasilan.
 */
function setQCStatus(row, status) {
  try {
    const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_QC);
    // Validasi input status
    if (status !== 'OK' && status !== 'NG') {
      throw new Error("Invalid status value.");
    }
    sheet.getRange(row, COL_QC.STATUS).setValue(status);
    
    // Kita juga perlu mengisi tanggal, waktu, dan operator seperti di handleQCEdit_
    const { date, time } = now_();
    sheet.getRange(row, COL_QC.DATE, 1, 3).setValues([[date, time, getEmail_()]]);

    return { success: true, message: `Status for row ${row} updated to ${status}.` };
  } catch (e) {
    console.error(`Error in setQCStatus: ${e.message}`);
    return { success: false, message: e.message };
  }
}

// =========================================================================
//         ⭐ FUNGSI UNTUK MENERIMA DATA DARI WEB APP (Versi 2.0) ⭐
// =========================================================================

/**
 * Memproses scan untuk ditambahkan ke sheet Production.
 * @param {string} qrCode - QR Code yang di-scan.
 * @returns {Object} - Objek berisi status keberhasilan dan pesan.
 */
function addProductionRecord(qrCode) {
  try {
    if (!qrCode) throw new Error("QR Code tidak boleh kosong.");

    // Validasi 1: QR harus terdaftar di Supplies
    if (!isQRExistsInSupplies_(qrCode)) {
      throw new Error(`QR "${qrCode}" tidak terdaftar di Supplies.`);
    }

    const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_PRODUCTION);
    const { date, time } = now_();
    const email = getEmail_();

    // --- PERBAIKAN: Logika untuk First_Seen dan Scan_Count ---
    const dataRange = sheet.getRange(START_ROW, COL_PROD.QR, sheet.getLastRow() - START_ROW + 1, sheet.getLastColumn());
    const values = dataRange.getValues();
    let existingRowIndex = -1;

    for (let i = 0; i < values.length; i++) {
      if (values[i][COL_PROD.QR - 1] === qrCode) {
        existingRowIndex = i;
        break;
      }
    }

    if (existingRowIndex !== -1) {
      // Jika QR sudah ada, update Scan_Count
      const currentCount = values[existingRowIndex][COL_PROD.SCAN_COUNT - 1] || 0;
      sheet.getRange(START_ROW + existingRowIndex, COL_PROD.SCAN_COUNT).setValue(currentCount + 1);
       return { success: true, message: `✅ Scan count untuk ${qrCode} diupdate!` };
    } else {
      // Jika QR baru, tambahkan baris baru dengan First_Seen dan Scan_Count = 1
      sheet.appendRow([qrCode, date, time, email, '', date, 1]); // LINE (kolom 5) kosong, First_Seen (6), Scan_Count (7)
      return { success: true, message: `✅ Berhasil ditambahkan ke Production!` };
    }

  } catch (e) {
    console.error(`Error in addProductionRecord: ${e.message}`);
    return { success: false, message: `❌ Gagal: ${e.message}` };
  }
}

/**
 * Memproses scan untuk ditambahkan ke sheet QC.
 * @param {string} qrCode - QR Code yang di-scan.
 * @param {string} status - Status yang dipilih ('OK' atau 'NG').
 * @returns {Object} - Objek berisi status keberhasilan dan pesan.
 */
function addQCRecord(qrCode, status) {
   try {
    if (!qrCode) throw new Error("QR Code tidak boleh kosong.");
    if (status !== 'OK' && status !== 'NG') throw new Error("Status tidak valid.");

    // Validasi 1: QR harus ada di Production
    if (!isQRExistsInProduction_(qrCode)) {
      throw new Error(`QR "${qrCode}" tidak ditemukan di Production.`);
    }

    // Validasi 2: QR tidak boleh duplikat di QC
    const qcSheet = SpreadsheetApp.getActive().getSheetByName(SHEET_QC);
    const qcQRs = getValidQRsFromSheet_(qcSheet, COL_QC.QR);
    if (qcQRs.has(qrCode)) {
      throw new Error(`QR "${qrCode}" sudah pernah di-scan di QC.`);
    }

    const { date, time } = now_();
    const email = getEmail_();

    // Menambahkan baris baru
    qcSheet.appendRow([qrCode, date, time, email, status, '', '']); // Kolom NG dikosongkan

    return { success: true, message: `✅ Status "${status}" berhasil disimpan!` };

  } catch (e) {
    console.error(`Error in addQCRecord: ${e.message}`);
    return { success: false, message: `❌ Gagal: ${e.message}` };
  }
}
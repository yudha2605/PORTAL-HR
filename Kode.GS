/**
 * Portal HR - Configuration & Backend Logic
 */

const SPREADSHEET_ID = '1NdxWvdzYjTlNb7uYlymNpsO0JRUIBkqjig4Rvw261ak'; // Database Karyawan
const SHEET_NAME = 'Form Responses 1';

const ATTENDANCE_SS_ID = '1fnUKoNq4vppI3RlQEeUewY1QCsJghcia1Y2FrkN6WQQ'; // Database Absensi
const ATTENDANCE_SHEET_NAME = 'Data Absensi';

function doGet(e) {
  var page = e.parameter.page || 'Login';
  var fileName = (page.toLowerCase() === 'dashboard') ? 'index' : page;

  var htmlContent = "";
  try {
    htmlContent = HtmlService.createHtmlOutputFromFile(fileName).getContent();
  } catch (err) {
    return HtmlService.createHtmlOutput(
      "<div style='font-family:sans-serif; padding:20px; color:#721c24; background:#f8d7da; border-radius:8px;'>" +
      "<h3>Error: File '" + fileName + "' tidak ditemukan.</h3>" +
      "</div>"
    );
  }

  htmlContent = htmlContent.replace(/{{include:(.*?)}}/g, function(match, p1) {
    return include(p1.trim());
  });

  var template = HtmlService.createTemplate(htmlContent);

  return template.evaluate()
    .setTitle('Portal HR')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  try {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  } catch (e) {
    return "<!-- Error Include: " + filename + " -->";
  }
}

function getEmployeeData() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error("Sheet '" + SHEET_NAME + "' tidak ditemukan!");
    const data = sheet.getDataRange().getValues();
    const result = [];
    const val = (row, idx) => {
      let v = row[idx];
      if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), "dd/MM/yyyy");
      return v === undefined || v === null ? "" : String(v);
    };
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[2]) continue; 
      result.push({
        row: i + 1,
        nip: val(row, 1),
        nama: val(row, 2),
        cabang: val(row, 3),
        jabatan: val(row, 4),
        tglMasuk: val(row, 5),
        status: val(row, 9),
        tglResign: val(row, 10),
        jenisKelamin: val(row, 16),
        wa: val(row, 17),
        email: val(row, 18),
        nik: val(row, 19),
        alamat: val(row, 20),
        pengalaman: val(row, 21),
        tglLahir: val(row, 23),
        goldar: val(row, 25),
        disc: val(row, 26),
        fotoKtp: val(row, 27),
        pasFoto: val(row, 28),
        fotoKk: val(row, 31)
      });
    }
    return result;
  } catch (e) { throw new Error(e.message); }
}

function updateEmployeeData(rowData) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    const rowNum = parseInt(rowData.row);
    let tglResignValue = rowData.status === 'Resign' ? rowData.tglResign : "";
    sheet.getRange(rowNum, 2).setValue(rowData.nip);
    sheet.getRange(rowNum, 3).setValue(rowData.nama);
    sheet.getRange(rowNum, 4).setValue(rowData.cabang);
    sheet.getRange(rowNum, 5).setValue(rowData.jabatan);
    sheet.getRange(rowNum, 6).setValue(rowData.tglMasuk);
    sheet.getRange(rowNum, 10).setValue(rowData.status);
    sheet.getRange(rowNum, 11).setValue(tglResignValue);
    sheet.getRange(rowNum, 17).setValue(rowData.jenisKelamin);
    sheet.getRange(rowNum, 18).setValue(rowData.wa);
    sheet.getRange(rowNum, 19).setValue(rowData.email);
    sheet.getRange(rowNum, 20).setValue(rowData.nik);
    sheet.getRange(rowNum, 21).setValue(rowData.alamat);
    sheet.getRange(rowNum, 24).setValue(rowData.tglLahir);
    sheet.getRange(rowNum, 26).setValue(rowData.goldar);
    return true;
  } catch (e) { throw new Error("Gagal update data: " + e.message); }
}

function deleteEmployeeData(rowNum) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    sheet.deleteRow(parseInt(rowNum));
    return true;
  } catch (e) { throw new Error("Gagal menghapus data: " + e.message); }
}

function getAttendanceFormData() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    const employees = [];
    const branches = new Set();
    for (let i = 1; i < data.length; i++) {
      const nama = data[i][2]; 
      const cabang = data[i][3]; 
      const status = data[i][9]; 
      if (nama && status === 'Bekerja') {
        employees.push({ nama: nama, cabang: cabang });
        if (cabang) branches.add(cabang);
      }
    }
    return { employees: employees, branches: Array.from(branches).sort() };
  } catch (e) { throw new Error("Gagal mengambil data dropdown: " + e.message); }
}

function saveAttendance(formData) {
  try {
    const ss = SpreadsheetApp.openById(ATTENDANCE_SS_ID);
    let sheet = ss.getSheetByName(ATTENDANCE_SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(ATTENDANCE_SHEET_NAME);
      sheet.appendRow(["Tanggal", "Empty", "Nama Karyawan", "Cabang", "Waktu masuk"]);
    }
    const data = sheet.getDataRange().getValues();
    const inputDate = new Date(formData.tanggal);
    const inputDateStr = Utilities.formatDate(inputDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
    for (let i = 1; i < data.length; i++) {
      let rowDateStr = "";
      if (data[i][0] instanceof Date) rowDateStr = Utilities.formatDate(data[i][0], Session.getScriptTimeZone(), "yyyy-MM-dd");
      else if (data[i][0]) rowDateStr = Utilities.formatDate(new Date(data[i][0]), Session.getScriptTimeZone(), "yyyy-MM-dd");
      if (rowDateStr === inputDateStr && data[i][2] === formData.nama) return { success: false, message: "Data sudah ada!" };
    }
    sheet.appendRow([formData.tanggal, "", formData.nama, formData.cabang, formData.waktu]);
    return { success: true };
  } catch (e) { throw new Error("Gagal menyimpan absensi: " + e.message); }
}

/**
 * Mengambil data absensi dengan dukungan Pencarian (Search)
 */
function getAttendanceData(limit, offset, search) {
  limit = limit || 50;
  offset = offset || 0;
  search = search ? search.toLowerCase() : "";
  
  try {
    const ss = SpreadsheetApp.openById(ATTENDANCE_SS_ID);
    const sheet = ss.getSheetByName(ATTENDANCE_SHEET_NAME);
    if (!sheet) return [];
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];

    // Jika ada pencarian, kita terpaksa mengambil lebih banyak data untuk difilter
    // Namun untuk efisiensi, kita ambil data teks (DisplayValues)
    if (search !== "") {
      const allData = sheet.getRange(2, 1, lastRow - 1, 5).getDisplayValues();
      const filtered = [];
      for (let i = allData.length - 1; i >= 0; i--) {
        const row = allData[i];
        if (row[2].toLowerCase().includes(search) || row[3].toLowerCase().includes(search)) {
          filtered.push({
            row: i + 2,
            tanggal: row[0],
            nama: row[2],
            cabang: row[3],
            waktu: formatTimeDisplay(row[4])
          });
          if (filtered.length >= limit + offset) break;
        }
      }
      return filtered.slice(offset, offset + limit);
    }

    // Jika tidak ada pencarian, gunakan logika Pagination standar (cepat)
    let endRow = lastRow - offset;
    let startRow = Math.max(2, endRow - limit + 1);
    let rowsToGet = endRow - startRow + 1;
    if (rowsToGet <= 0) return [];

    const displayData = sheet.getRange(startRow, 1, rowsToGet, 5).getDisplayValues();
    const result = [];
    for (let i = displayData.length - 1; i >= 0; i--) {
      const row = displayData[i];
      if (!row[0]) continue;
      result.push({
        row: startRow + i,
        tanggal: row[0], 
        nama: row[2],
        cabang: row[3],
        waktu: formatTimeDisplay(row[4])
      });
    }
    return result;
  } catch (e) { throw new Error("Gagal mengambil data absensi: " + e.message); }
}

function formatTimeDisplay(waktuRaw) {
  if (!waktuRaw || !waktuRaw.includes(':')) return waktuRaw;
  let parts = waktuRaw.split(':');
  if (parts.length >= 2) return parts[0].padStart(2, '0') + ':' + parts[1].padStart(2, '0');
  return waktuRaw;
}

function updateAttendanceData(rowData) {
  try {
    const ss = SpreadsheetApp.openById(ATTENDANCE_SS_ID);
    const sheet = ss.getSheetByName(ATTENDANCE_SHEET_NAME);
    const rowNum = parseInt(rowData.row);
    sheet.getRange(rowNum, 1).setValue(rowData.tanggal);
    sheet.getRange(rowNum, 3).setValue(rowData.nama);
    sheet.getRange(rowNum, 4).setValue(rowData.cabang);
    sheet.getRange(rowNum, 5).setValue(rowData.waktu);
    return true;
  } catch (e) { throw new Error("Gagal update absensi: " + e.message); }
}

function deleteAttendanceData(rowNum) {
  try {
    const ss = SpreadsheetApp.openById(ATTENDANCE_SS_ID);
    const sheet = ss.getSheetByName(ATTENDANCE_SHEET_NAME);
    sheet.deleteRow(parseInt(rowNum));
    return true;
  } catch (e) { throw new Error("Gagal menghapus absensi: " + e.message); }
}

function getPageContent(pageName) {
  try {
    var template = HtmlService.createTemplateFromFile(pageName);
    var content = template.evaluate().getContent();
    content = content.replace(/{{include:(.*?)}}/g, function(match, p1) { return include(p1.trim()); });
    return content;
  } catch (e) { return "<div class='alert alert-danger'>Halaman " + pageName + " tidak ditemukan.</div>"; }
}

function checkLogin(username, password) { return (username === 'hrd' && password === '0'); }
function getScriptUrl() { return ScriptApp.getService().getUrl(); }

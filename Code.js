/**
 * EduSecure CBT Enterprise - Google Apps Script Backend (Code.gs)
 * Fungsi: Menayangkan Web App CBT & Menghubungkan Data ke Google Sheets
 */

function doGet(e) {
  var template = HtmlService.createTemplateFromFile('index');
  return template.evaluate()
    .setTitle('EduSecure CBT Enterprise v5.0')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getOrCreateSheet(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    // Jika dijalankan standalone, gunakan Spreadsheet aktif atau buat baru
    return null;
  }
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  return sheet;
}

function apiSaveResult(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { success: false, message: "Spreadsheet tidak ditemukan." };
    
    var sheet = getOrCreateSheet("Hasil_Ujian");
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["Timestamp", "NIS", "Nama Siswa", "Kelas", "Mata Pelajaran", "Paket", "Nilai", "Status Kejujuran", "Jumlah Pelanggaran"]);
    }
    
    sheet.appendRow([
      new Date(),
      data.nis,
      data.name,
      data.class,
      data.subject,
      data.paket || "Paket A",
      data.score,
      data.status,
      data.violations || 0
    ]);
    
    return { success: true, message: "Hasil ujian berhasil tersimpan di Google Sheets!" };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function apiGetQuestions() {
  try {
    var sheet = getOrCreateSheet("Bank_Soal");
    if (!sheet || sheet.getLastRow() <= 1) {
      return { success: true, data: [] };
    }
    
    var rows = sheet.getDataRange().getValues();
    var questions = [];
    
    for (var i = 1; i < rows.length; i++) {
      var row = rows[i];
      questions.push({
        id: row[0],
        paket: row[1],
        type: row[2],
        question: row[3],
        options: row[4] ? JSON.parse(row[4]) : [],
        key: row[5],
        guruName: row[6],
        subject: row[7],
        image: row[8] || null,
        audio: row[9] || null
      });
    }
    
    return { success: true, data: questions };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function apiSaveQuestion(qData) {
  try {
    var sheet = getOrCreateSheet("Bank_Soal");
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["ID", "Paket", "Tipe", "Pertanyaan", "Opsi (JSON)", "Kunci", "Nama Guru", "Mapel", "URL Gambar", "URL Audio"]);
    }
    
    sheet.appendRow([
      qData.id || Date.now(),
      qData.paket,
      qData.type,
      qData.question,
      JSON.stringify(qData.options || []),
      JSON.stringify(qData.key || ""),
      qData.guruName,
      qData.subject,
      qData.image || "",
      qData.audio || ""
    ]);
    
    return { success: true, message: "Soal berhasil disimpan ke Google Sheets!" };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}
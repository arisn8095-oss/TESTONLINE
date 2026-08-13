/**
 * EduSecure CBT Enterprise v5.0 - Google Apps Script Backend (Code.gs)
 * Restriksi Akun Guru: 1 Guru = 1 Mata Pelajaran
 * Menghubungkan Bank Soal, Jadwal Ujian, Data Guru, Murid, & Hasil Ke Google Sheets.
 */

function doGet(e) {
  var template = HtmlService.createTemplateFromFile('index');
  return template.evaluate()
    .setTitle('EduSecure CBT Enterprise v5.0 - Anti Cheat System')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getOrCreateSheet(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return null;
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  return sheet;
}

function apiLogActivity(actorId, role, action, details) {
  try {
    var sheet = getOrCreateSheet("Log_Aktivitas");
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["Timestamp", "ID / NIS / NIP", "Role", "Tindakan / Aktivitas", "Detail Informasi"]);
    }
    sheet.appendRow([
      new Date().toLocaleString("id-ID"),
      String(actorId || "SYSTEM"),
      String(role || "System"),
      String(action || "-"),
      String(details || "-")
    ]);
    return { success: true };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function apiGetLogs() {
  try {
    var sheet = getOrCreateSheet("Log_Aktivitas");
    if (!sheet || sheet.getLastRow() <= 1) {
      return { success: true, data: [] };
    }
    var rows = sheet.getDataRange().getValues();
    var logs = [];
    for (var i = 1; i < rows.length; i++) {
      logs.push({
        time: rows[i][0],
        actor: rows[i][1],
        role: rows[i][2],
        action: rows[i][3],
        details: rows[i][4]
      });
    }
    return { success: true, data: logs };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function apiGetTeachers() {
  try {
    var sheet = getOrCreateSheet("Data_Guru");
    if (!sheet || sheet.getLastRow() <= 1) {
      return { success: true, data: [] };
    }
    var rows = sheet.getDataRange().getValues();
    var teachers = [];
    for (var i = 1; i < rows.length; i++) {
      teachers.push({
        nip: String(rows[i][0]),
        name: String(rows[i][1]),
        subject: String(rows[i][2]), // Enforce 1 Guru = 1 Mata Pelajaran
        pakets: rows[i][3] ? JSON.parse(rows[i][3]) : ['Paket A', 'Paket B', 'Paket C']
      });
    }
    return { success: true, data: teachers };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function apiSaveTeacher(tData) {
  try {
    var sheet = getOrCreateSheet("Data_Guru");
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["NIP", "Nama Guru", "Mata Pelajaran (Restriksi 1 Mapel)", "Paket Soal"]);
    }
    sheet.appendRow([
      String(tData.nip),
      String(tData.name),
      String(tData.subject),
      JSON.stringify(tData.pakets || ['Paket A', 'Paket B', 'Paket C'])
    ]);
    apiLogActivity(tData.nip, "Admin/Guru", "Tambah Guru", "Menambahkan Guru " + tData.name + " (" + tData.subject + ")");
    return { success: true, message: "Data Guru (1 Mapel) berhasil disimpan ke Google Sheets!" };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function apiGetStudents() {
  try {
    var sheet = getOrCreateSheet("Data_Murid");
    if (!sheet || sheet.getLastRow() <= 1) {
      return { success: true, data: [] };
    }
    var rows = sheet.getDataRange().getValues();
    var students = [];
    for (var i = 1; i < rows.length; i++) {
      students.push({
        id: String(rows[i][0]),
        name: String(rows[i][1]),
        class: String(rows[i][2]),
        status: String(rows[i][3] || "Normal"),
        violations: Number(rows[i][4] || 0)
      });
    }
    return { success: true, data: students };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function apiUpdateStudentStatus(nis, status, violations) {
  try {
    var sheet = getOrCreateSheet("Data_Murid");
    if (!sheet || sheet.getLastRow() <= 1) return { success: false, message: "Data Murid kosong" };
    
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(nis)) {
        sheet.getRange(i + 1, 4).setValue(status);
        sheet.getRange(i + 1, 5).setValue(violations);
        apiLogActivity(nis, "Murid", "Update Security Lockdown", "Status: " + status + " | Pelanggaran: " + violations + "x");
        return { success: true, message: "Status NIS " + nis + " diperbarui di Google Sheets!" };
      }
    }
    return { success: false, message: "NIS tidak ditemukan di Google Sheets" };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function apiGetExams() {
  try {
    var sheet = getOrCreateSheet("Jadwal_Ujian");
    if (!sheet || sheet.getLastRow() <= 1) {
      return { success: true, data: [] };
    }
    var rows = sheet.getDataRange().getValues();
    var exams = [];
    for (var i = 1; i < rows.length; i++) {
      exams.push({
        id: String(rows[i][0]),
        subject: String(rows[i][1]),
        teacher: String(rows[i][2]),
        paket: String(rows[i][3]),
        token: String(rows[i][4]),
        durationMinutes: Number(rows[i][5]),
        docLinkSource: String(rows[i][6] || "-"),
        targetClass: String(rows[i][7] || "Semua Kelas")
      });
    }
    return { success: true, data: exams };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function apiPublishExam(examData) {
  try {
    var sheet = getOrCreateSheet("Jadwal_Ujian");
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["Exam ID", "Mata Pelajaran", "Guru Pengampu", "Paket Soal", "Token Keamanan", "Durasi (Menit)", "Link/Sumber Doc Soal", "Target Kelas"]);
    }
    
    var rows = sheet.getDataRange().getValues();
    var updated = false;
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(examData.id) || (String(rows[i][1]) === String(examData.subject) && String(rows[i][3]) === String(examData.paket))) {
        sheet.getRange(i + 1, 4).setValue(examData.paket);
        sheet.getRange(i + 1, 5).setValue(examData.token);
        sheet.getRange(i + 1, 6).setValue(examData.durationMinutes);
        if (examData.docLinkSource) sheet.getRange(i + 1, 7).setValue(examData.docLinkSource);
        updated = true;
        break;
      }
    }
    
    if (!updated) {
      sheet.appendRow([
        examData.id || ("EX-" + Date.now()),
        examData.subject,
        examData.teacher,
        examData.paket || "Paket A",
        examData.token,
        examData.durationMinutes || 90,
        examData.docLinkSource || "Input Manual Guru",
        examData.targetClass || "Semua Kelas"
      ]);
    }

    apiLogActivity(examData.teacher, "Guru", "Terbitkan Ujian", "Mapel: " + examData.subject + " | Paket: " + examData.paket + " | Token: " + examData.token);
    return { success: true, message: "Jadwal Ujian berhasil tersimpan di Google Sheets!" };
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
      questions.push({
        id: Number(rows[i][0]),
        paket: String(rows[i][1]),
        type: String(rows[i][2]),
        question: String(rows[i][3]),
        options: rows[i][4] ? JSON.parse(rows[i][4]) : [],
        key: rows[i][5] ? JSON.parse(rows[i][5]) : "",
        guruName: String(rows[i][6]),
        subject: String(rows[i][7]),
        image: rows[i][8] || null,
        audio: rows[i][9] || null,
        docLinkSource: String(rows[i][10] || "-")
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
      sheet.appendRow(["ID", "Paket", "Tipe", "Pertanyaan", "Opsi (JSON)", "Kunci", "Nama Guru", "Mapel", "URL Gambar", "URL Audio", "Sumber Link Doc"]);
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
      qData.audio || "",
      qData.docLinkSource || "Input Manual Guru"
    ]);
    
    apiLogActivity(qData.guruName, "Guru", "Input Soal Manual", "Mapel: " + qData.subject + " | Paket: " + qData.paket);
    return { success: true, message: "Soal Mapel " + qData.subject + " berhasil disimpan ke Google Sheets!" };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function apiBatchSaveQuestions(questionsList, docLinkSource) {
  try {
    var sheet = getOrCreateSheet("Bank_Soal");
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["ID", "Paket", "Tipe", "Pertanyaan", "Opsi (JSON)", "Kunci", "Nama Guru", "Mapel", "URL Gambar", "URL Audio", "Sumber Link Doc"]);
    }
    
    var rowsToAdd = questionsList.map(function(q) {
      return [
        q.id || Date.now(),
        q.paket || "Paket A",
        q.type || "PG",
        q.question,
        JSON.stringify(q.options || []),
        JSON.stringify(q.key || "A"),
        q.guruName || "Guru",
        q.subject || "Informatika",
        q.image || "",
        q.audio || "",
        docLinkSource || q.docLinkSource || "File/Link Import Doc"
      ];
    });
    
    if (rowsToAdd.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAdd.length, 11).setValues(rowsToAdd);
    }
    
    apiLogActivity(questionsList[0] ? questionsList[0].guruName : "Guru", "Guru", "Import Massal Soal Doc", "Mengimpor " + questionsList.length + " soal mapel " + (questionsList[0] ? questionsList[0].subject : ""));
    return { success: true, message: questionsList.length + " Soal berhasil tersimpan di Google Sheets!" };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function apiSaveResult(data) {
  try {
    var sheet = getOrCreateSheet("Hasil_Ujian");
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["Timestamp", "NIS", "Nama Siswa", "Kelas", "Mata Pelajaran", "Paket", "Nilai", "Status Kejujuran", "Jumlah Pelanggaran"]);
    }
    
    sheet.appendRow([
      new Date().toLocaleString("id-ID"),
      data.nis,
      data.name,
      data.class,
      data.subject,
      data.paket || "Paket A",
      data.score,
      data.status,
      data.violations || 0
    ]);
    
    apiLogActivity(data.nis, "Murid", "Selesaikan Ujian", "Mapel: " + data.subject + " | Nilai: " + data.score + " | Status: " + data.status);
    return { success: true, message: "Hasil ujian berhasil tersimpan di Google Sheets!" };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// Menampilkan halaman HTML saat webapp diakses
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Aplikasi Audit Internal')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Fungsi untuk memproses data dari frontend
function processAuditForm(formData) {
  try {
    // 1. ID Folder Google Drive Anda (Tempat menyimpan gambar TTD)
    const folderId = "1xjVi1fxtIcZm28BoPKSyFoKJA_8-Elei"; 
    const folder = DriveApp.getFolderById(folderId);
    
    // Proses Tanda Tangan Auditor
    let auditorSignUrl = "";
    if (formData.auditorSignature) {
      const auditorSignData = formData.auditorSignature.split(',')[1];
      const auditorBlob = Utilities.newBlob(Utilities.base64Decode(auditorSignData), 'image/png', formData.auditorName + '_auditor_sign.png');
      const auditorFile = folder.createFile(auditorBlob);
      auditorSignUrl = auditorFile.getUrl(); // Mengambil link gambar TTD di Drive
    }

    // Proses Tanda Tangan Auditee
    let auditeeSignUrl = "";
    if (formData.auditeeSignature) {
      const auditeeSignData = formData.auditeeSignature.split(',')[1];
      const auditeeBlob = Utilities.newBlob(Utilities.base64Decode(auditeeSignData), 'image/png', formData.auditeeName + '_auditee_sign.png');
      const auditeeFile = folder.createFile(auditeeBlob);
      auditeeSignUrl = auditeeFile.getUrl(); // Mengambil link gambar TTD di Drive
    }

    // 2. ID Google Sheet Anda (Sebagai Database)
    const sheetId = "1Cmccso6hkwHPJNX1jZy5arK5OAOLnhMBCp89JwJBG50";
    const ss = SpreadsheetApp.openById(sheetId);
    const sheet = ss.getSheets()[0]; // Otomatis menggunakan sheet pertama (paling kiri)
    
    sheet.appendRow([
      new Date(),                     // Kolom A: Timestamp
      formData.auditorName,           // Kolom B: Nama Auditor
      formData.auditeeName,           // Kolom C: Nama Auditee
      formData.checklistData,         // Kolom D: Hasil Checklist
      formData.pengamatan,            // Kolom E: Form Pengamatan
      formData.lks,                   // Kolom F: Form LKS
      auditorSignUrl,                 // Kolom G: Link TTD Auditor
      auditeeSignUrl                  // Kolom H: Link TTD Auditee
    ]);
    
    return { success: true, message: "Data Audit Berhasil Disimpan ke Google Sheet!" };
  } catch (error) {
    return { success: false, message: "Terjadi kesalahan: " + error.message };
  }
}

// Fungsi untuk mengecek login menggunakan NIK dan Password
function checkLogin(nik, password) {
  try {
    const sheetId = "1Cmccso6hkwHPJNX1jZy5arK5OAOLnhMBCp89JwJBG50";
    const ss = SpreadsheetApp.openById(sheetId);
    const userSheet = ss.getSheetByName("user");
    
    if (!userSheet) return { success: false, message: "Sheet 'user' tidak ditemukan di database." };
    
    const data = userSheet.getDataRange().getValues();
    
    // Looping data mulai dari baris 2 (indeks 1) karena baris 1 adalah Header (NIK | Password | Nama)
    for (let i = 1; i < data.length; i++) {
      let sheetNik = String(data[i][0]).trim();
      let sheetPass = String(data[i][1]).trim();
      let sheetNama = String(data[i][2]).trim();
      
      if (sheetNik === String(nik).trim() && sheetPass === String(password).trim()) {
        return { success: true, nama: sheetNama, message: "Login Berhasil!" };
      }
    }
    return { success: false, message: "NIK atau Password salah!" };
  } catch (error) {
    return { success: false, message: "Terjadi kesalahan sistem: " + error.message };
  }
}

// Fungsi ini menangani request POST (fetch) dari GitHub Pages
function doPost(e) {
  try {
    const postData = JSON.parse(e.postData.contents);
    
    if (postData.action === "login") {
      const result = checkLogin(postData.nik, postData.password);
      return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
    } else if (postData.action === "submitAudit") {
      const result = processAuditForm(postData.formData);
      return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
    }
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: error.message })).setMimeType(ContentService.MimeType.JSON);
  }
}
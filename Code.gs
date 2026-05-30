// Menampilkan halaman HTML saat webapp diakses
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Aplikasi Audit Internal Matrika')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// Konfigurasi ID (PENTING: Sesuaikan dengan ID Google Drive & Sheet Anda)
const FOLDER_ID = "1xjVi1fxtIcZm28BoPKSyFoKJA_8-Elei";
const SHEET_ID = "1Cmccso6hkwHPJNX1jZy5arK5OAOLnhMBCp89JwJBG50";

// --- Fungsi Utilitas ---
function getSheet(sheetName) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(sheetName);
  
  // Buat sheet jika belum ada (untuk kemudahan setup pertama kali)
  if (!sheet) {
    if (sheetName === 'checklist_template') {
       sheet = ss.insertSheet(sheetName);
       sheet.appendRow(['Kategori', 'No', 'Pertanyaan']);
    } else if (sheetName === 'data_pengamatan') {
       sheet = ss.insertSheet(sheetName);
       sheet.appendRow(['ID', 'Timestamp', 'Auditor', 'Auditee', 'Kategori Menu', 'Data Checklist (JSON)', 'Status', 'Tanda Tangan Auditor', 'Tanda Tangan Auditee']);
    }
  }
  return sheet;
}

function saveImageToDrive(base64Data, filename) {
  if (!base64Data) return "";
  try {
    const folder = DriveApp.getFolderById(FOLDER_ID);
    const data = base64Data.split(',')[1];
    const blob = Utilities.newBlob(Utilities.base64Decode(data), 'image/png', filename);
    const file = folder.createFile(blob);
    return file.getUrl();
  } catch(e) {
    return "Error Upload: " + e.message;
  }
}

// --- Handler API ---

function handleLogin(nik, password) {
  try {
    const userSheet = getSheet("user");
    if (!userSheet) return { success: false, message: "Sheet 'user' tidak ditemukan." };
    
    const data = userSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(nik).trim() && String(data[i][1]).trim() === String(password).trim()) {
        return { success: true, nama: String(data[i][2]).trim(), message: "Login Berhasil!" };
      }
    }
    return { success: false, message: "NIK atau Password salah!" };
  } catch (error) {
    return { success: false, message: "Error sistem: " + error.message };
  }
}

function handleGetChecklist(kategori) {
  try {
    const sheet = getSheet("checklist_template");
    const data = sheet.getDataRange().getValues();
    let questions = [];
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim().toUpperCase() === kategori.toUpperCase()) {
        questions.push({
          no: data[i][1],
          pertanyaan: data[i][2]
        });
      }
    }
    
    // Fallback Mock Data jika sheet kosong (untuk demo)
    if (questions.length === 0) {
      questions = [
        { no: 1, pertanyaan: "Apakah manajemen puncak telah membuat Kebijakan Halal tertulis untuk hanya memproduksi produk halal secara konsisten dan berkesinambungan?" },
        { no: 2, pertanyaan: "Apakah ada kegiatan sosialisasi Kebijakan Halal ? Sebutkan bentuk sosialisasinya." },
        { no: 3, pertanyaan: "Apakah bangunan bebas dari bahan najis/haram?" }
      ];
    }
    
    return { success: true, data: questions };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function handleSaveData(formData, status) {
  try {
    const sheet = getSheet("data_pengamatan");
    const id = Utilities.getUuid();
    
    // Handle Signatures only if Submit Final
    let urlAuditor = "";
    let urlAuditee = "";
    if (status === 'Submitted') {
       urlAuditor = saveImageToDrive(formData.auditorSignature, id + "_auditor.png");
       urlAuditee = saveImageToDrive(formData.auditeeSignature, id + "_auditee.png");
    }

    sheet.appendRow([
      id,
      new Date(),
      formData.auditorName,
      formData.auditeeName,
      formData.kategori,
      JSON.stringify(formData.checklistData), // Simpan array objek sebagai JSON
      status, // 'Draft' atau 'Submitted'
      urlAuditor,
      urlAuditee
    ]);
    
    return { success: true, message: `Data berhasil disimpan sebagai ${status}!` };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function handleGetList(auditorName) {
  try {
    const sheet = getSheet("data_pengamatan");
    const data = sheet.getDataRange().getValues();
    let list = [];
    
    for (let i = 1; i < data.length; i++) {
      // Filter by auditor if needed, for now return all or just match name
      // If we want auditor to see only their data: if(data[i][2] === auditorName)
      list.push({
        id: data[i][0],
        tanggal: data[i][1],
        auditor: data[i][2],
        auditee: data[i][3],
        kategori: data[i][4],
        status: data[i][6]
      });
    }
    // Balik urutan agar yang terbaru di atas
    list.reverse();
    return { success: true, data: list };
  } catch(e) {
    return { success: false, message: e.message };
  }
}


// --- POST Entry Point ---
function doPost(e) {
  try {
    const postData = JSON.parse(e.postData.contents);
    let result = {};
    
    switch (postData.action) {
      case "login":
        result = handleLogin(postData.nik, postData.password);
        break;
      case "getChecklist":
        result = handleGetChecklist(postData.kategori);
        break;
      case "saveDraft":
        result = handleSaveData(postData.formData, "Draft");
        break;
      case "submitFinal":
        result = handleSaveData(postData.formData, "Submitted");
        break;
      case "getList":
        result = handleGetList(postData.auditorName);
        break;
      default:
        result = { success: false, message: "Action tidak dikenal." };
    }
    
    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: error.message })).setMimeType(ContentService.MimeType.JSON);
  }
}
/**
 * INTERNAL AUDIT MANAGEMENT SYSTEM (IAMS)
 * Core Backend - Google Apps Script (Code.gs)
 * 
 * UPGRADES & OPTIMIZATIONS:
 * 1. Enterprise Password Hashing (SHA-256) untuk perlindungan kredensial staf.
 * 2. High-Performance Batch Write (setValues) pada Input Audit & Temuan untuk memangkas waktu loading.
 * 3. Mekanisme Auto-Overdue Engine terpadu di sisi server.
 */

const SPREADSHEET = SpreadsheetApp.getActiveSpreadsheet();

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('IAMS - Internal Audit Management System')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Fungsi Enkripsi Keamanan SHA-256
 */
function hashPassword(input) {
  if (!input) return "";
  const rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(input), Utilities.Charset.UTF_8);
  let output = "";
  for (let i = 0; i < rawHash.length; i++) {
    let v = rawHash[i] & 0xff;
    output += (v < 16 ? "0" : "") + v.toString(16);
  }
  return output;
}

/**
 * Setup Database & Inisialisasi Akun Terenkripsi
 */
function setupDatabase() {
  const sheets = {
    "Users": [["ID", "Email", "Password", "Role", "Nama"]],
    "Audits": [["Nomor Audit", "Tanggal Audit", "Auditor ID", "Outlet ID", "Kesimpulan Audit", "Status Audit"]],
    "Findings": [["ID Temuan", "Nomor Audit", "Tanggal", "Outlet", "Auditor", "Jenis Audit", "Deskripsi Temuan", "Kategori Risiko", "PIC", "Deadline", "Status", "Photos"]],
    "Followups": [["ID Temuan", "Action Plan", "Progress", "Upload Bukti Perbaikan", "Catatan Auditor", "Status Verifikasi"]],
    "MasterAuditors": [["ID", "Nama Auditor"]],
    "MasterOutlets": [["ID", "Nama Outlet"]],
    "MasterJenisAudit": [["ID", "Nama Jenis Audit"]]
  };

  for (let sheetName in sheets) {
    let sheet = SPREADSHEET.getSheetByName(sheetName);
    if (!sheet) {
      sheet = SPREADSHEET.insertSheet(sheetName);
      sheet.getRange(1, 1, 1, sheets[sheetName][0].length).setValues(sheets[sheetName]);
      sheet.getRange(1, 1, 1, sheets[sheetName][0].length).setFontWeight("bold").setBackground("#001f3f").setFontColor("#ffffff");
    }
  }

  const userSheet = SPREADSHEET.getSheetByName("Users");
  if (userSheet.getLastRow() <= 1) {
    userSheet.appendRow(["USR-01", "admin@iams.com", hashPassword("admin123"), "Admin", "Super Admin"]);
    userSheet.appendRow(["USR-02", "trianto@iams.com", hashPassword("auditor123"), "Auditor", "Trianto"]);
    userSheet.appendRow(["USR-03", "ilyas@iams.com", hashPassword("auditor123"), "Auditor", "Ilyas"]);
  }

  const auditorSheet = SPREADSHEET.getSheetByName("MasterAuditors");
  if (auditorSheet.getLastRow() <= 1) {
    auditorSheet.appendRow(["ADT-01", "Trianto"]);
    auditorSheet.appendRow(["ADT-02", "Ilyas"]);
  }

  const outletSheet = SPREADSHEET.getSheetByName("MasterOutlets");
  if (outletSheet.getLastRow() <= 1) {
    const outlets = [
      "JCHICKEN CILEDUG", "MOMOYO", "LUUCA", "MAXI CILEDUG", "ULON CILEDUG",
      "JCHICKEN CIREBON", "ULON SIGNATURE", "PENTA", "MAXI CIREBON", "PADELNIS"
    ];
    outlets.forEach((o, idx) => outletSheet.appendRow(["OUT-" + (idx + 1), o]));
  }

  const jenisSheet = SPREADSHEET.getSheetByName("MasterJenisAudit");
  if (jenisSheet.getLastRow() <= 1) {
    const types = [
      "Cash Opname", "Audit Aktiva", "Audit Kesesuaian Transaksi dan Pembatalan",
      "Audit Pelayanan", "Validasi Formula Menu", "Audit Pembelian Aset Bahan Baku dan Biaya",
      "Audit Mutasi Bank"
    ];
    types.forEach((t, idx) => jenisSheet.appendRow(["TP-" + (idx + 1), t]));
  }
}

/**
 * Autentikasi Login Pengguna Aman (SHA-256 Match Verification)
 */
function loginUser(email, password) {
  const sheet = SPREADSHEET.getSheetByName("Users");
  const data = sheet.getDataRange().getValues();
  const incomingHash = hashPassword(password);

  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === email && data[i][2] === incomingHash) {
      return {
        success: true,
        user: { id: data[i][0], email: data[i][1], role: data[i][3], name: data[i][4] }
      };
    }
  }
  return { success: false, message: "Kredensial salah. Silakan coba kembali." };
}

/**
 * Cloud Media Storage - Upload File ke Google Drive Publik
 */
function uploadToDrive(base64Data, fileName) {
  try {
    const contentType = base64Data.substring(5, base64Data.indexOf(';base64,'));
    const bytes = Utilities.base64Decode(base64Data.split(',')[1]);
    const blob = Utilities.newBlob(bytes, contentType, fileName);
    
    const file = DriveApp.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    return { success: true, url: file.getUrl() };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

/**
 * Penarikan Data Komprehensif dengan Sinkronisasi Auto-Overdue
 */
function getAppData() {
  const data = {};
  const sheets = ["Users", "Audits", "Findings", "Followups", "MasterAuditors", "MasterOutlets", "MasterJenisAudit"];
  
  sheets.forEach(sheetName => {
    const sheet = SPREADSHEET.getSheetByName(sheetName);
    if (sheet) {
      const values = sheet.getDataRange().getValues();
      const headers = values[0];
      const rows = [];
      for (let i = 1; i < values.length; i++) {
        let rowObj = {};
        headers.forEach((h, idx) => {
          let val = values[i][idx];
          if (val instanceof Date) {
            rowObj[h] = Utilities.formatDate(val, "GMT+7", "yyyy-MM-dd");
          } else {
            rowObj[h] = val;
          }
        });
        rows.push(rowObj);
      }
      data[sheetName] = rows;
    }
  });
  
  // Real-time Auto Overdue Engine
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const findingSheet = SPREADSHEET.getSheetByName("Findings");
  const findings = data["Findings"] || [];
  
  findings.forEach((finding, idx) => {
    if (finding["Deadline"]) {
      const deadline = new Date(finding["Deadline"]);
      deadline.setHours(0, 0, 0, 0);
      
      if (finding["Status"] !== "Closed" && today > deadline && finding["Status"] !== "Overdue") {
        findingSheet.getRange(idx + 2, 11).setValue("Overdue");
        finding["Status"] = "Overdue";
      }
    }
  });

  return data;
}

/**
 * Simpan Audit & Temuan Menggunakan Metode Batch-Processing (setValues)
 */
function saveAuditWithFindings(auditForm, findingsList) {
  try {
    const auditSheet = SPREADSHEET.getSheetByName("Audits");
    const findingsSheet = SPREADSHEET.getSheetByName("Findings");
    const followupSheet = SPREADSHEET.getSheetByName("Followups");
    
    const today = new Date();
    const dateStr = Utilities.formatDate(today, "GMT+7", "yyyyMM");
    const nextNum = auditSheet.getLastRow();
    const nomorAudit = `AUD-${dateStr}-${nextNum.toString().padStart(4, '0')}`;
    
    auditSheet.appendRow([
      nomorAudit, auditForm.tanggal, auditForm.auditor, auditForm.outlet, auditForm.kesimpulan, auditForm.status
    ]);
    
    if (findingsList && findingsList.length > 0) {
      const initialFindingsLastRow = findingsSheet.getLastRow();
      const initialFollowupsLastRow = followupSheet.getLastRow();
      
      let bulkFindingsData = [];
      let bulkFollowupsData = [];
      
      findingsList.forEach((finding, idx) => {
        const fNum = initialFindingsLastRow + idx;
        const idTemuan = `TMN-${dateStr}-${fNum.toString().padStart(4, '0')}`;
        
        bulkFindingsData.push([
          idTemuan, nomorAudit, auditForm.tanggal, auditForm.outlet, auditForm.auditor,
          finding.jenis, finding.deskripsi, finding.risiko, finding.pic, finding.deadline,
          "Open", finding.photos || ""
        ]);
        
        bulkFollowupsData.push([idTemuan, "", 0, "", "", "Pending"]);
      });
      
      findingsSheet.getRange(initialFindingsLastRow + 1, 1, bulkFindingsData.length, bulkFindingsData[0].length).setValues(bulkFindingsData);
      followupSheet.getRange(initialFollowupsLastRow + 1, 1, bulkFollowupsData.length, bulkFollowupsData[0].length).setValues(bulkFollowupsData);
    }
    
    return { success: true, nomorAudit: nomorAudit };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

/**
 * Modifikasi Status Temuan Lapangan
 */
function updateFindingStatus(findingId, newStatus) {
  try {
    const sheet = SPREADSHEET.getSheetByName("Findings");
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === findingId) {
        sheet.getRange(i + 1, 11).setValue(newStatus);
        return { success: true };
      }
    }
    return { success: false, message: "ID tidak ditemukan" };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

/**
 * Intervensi Hasil Tindak Lanjut Auditee (Follow-up Action)
 */
function submitFollowup(followupData) {
  try {
    const fSheet = SPREADSHEET.getSheetByName("Followups");
    const findingsSheet = SPREADSHEET.getSheetByName("Findings");
    const fData = fSheet.getDataRange().getValues();
    
    for (let i = 1; i < fData.length; i++) {
      if (fData[i][0] === followupData.idTemuan) {
        fSheet.getRange(i + 1, 2).setValue(followupData.actionPlan);
        fSheet.getRange(i + 1, 3).setValue(followupData.progress);
        fSheet.getRange(i + 1, 4).setValue(followupData.buktiPerbaikan);
        fSheet.getRange(i + 1, 5).setValue(followupData.catatanAuditor);
        fSheet.getRange(i + 1, 6).setValue(followupData.statusVerifikasi);
        
        const findingsData = findingsSheet.getDataRange().getValues();
        for (let j = 1; j < findingsData.length; j++) {
          if (findingsData[j][0] === followupData.idTemuan) {
            let statusTemuan = "Progress";
            if (followupData.statusVerifikasi === "Verified") {
              statusTemuan = "Closed";
              fSheet.getRange(i + 1, 3).setValue(100);
            } else if (followupData.statusVerifikasi === "Rejected") {
              statusTemuan = "Open";
            }
            findingsSheet.getRange(j + 1, 11).setValue(statusTemuan);
            break;
          }
        }
        return { success: true };
      }
    }
    return { success: false, message: "ID Tidak valid" };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

/**
 * Pendaftaran Master Pengguna Baru (SHA-256 Automated)
 */
function saveUser(userData) {
  try {
    const sheet = SPREADSHEET.getSheetByName("Users");
    const securePassword = hashPassword(userData.password);
    sheet.appendRow([userData.id, userData.email, securePassword, userData.role, userData.nama]);
    return { success: true };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function addNewOutletToMaster(id, name) {
  try { SPREADSHEET.getSheetByName("MasterOutlets").appendRow([id, name]); return { success: true }; } catch(e) { return { success: false }; }
}

function addNewScopeToMaster(id, name) {
  try { SPREADSHEET.getSheetByName("MasterJenisAudit").appendRow([id, name]); return { success: true }; } catch(e) { return { success: false }; }
}

/**
 * Penghapusan Temuan dengan Proteksi Row Cascade
 */
function deleteFinding(findingId) {
  try {
    const sheet = SPREADSHEET.getSheetByName("Findings");
    const data = sheet.getDataRange().getValues();
    let deleted = false;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === findingId) { sheet.deleteRow(i + 1); deleted = true; break; }
    }
    const fuSheet = SPREADSHEET.getSheetByName("Followups");
    if (fuSheet) {
      const fuData = fuSheet.getDataRange().getValues();
      for (let i = 1; i < fuData.length; i++) {
        if (fuData[i][0] === findingId) { fuSheet.deleteRow(i + 1); break; }
      }
    }
    return { success: deleted };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}
// Utils.gs
// Helper functions yang dipakai di mana-mana

// Tulis error ke tab Errors
function logError(affectedRow, sheetName, errorDescription) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const errorsSheet = ss.getSheetByName("Errors");
  
  errorsSheet.appendRow([
    new Date(),        // Timestamp
    affectedRow,       // Row yang error
    sheetName,         // Dari sheet mana
    errorDescription   // Keterangan error
  ]);
  
  Logger.log(`ERROR logged: Row ${affectedRow} - ${errorDescription}`);
}

// Ambil semua email di Do Not Contact list
function getDNCList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dncSheet = ss.getSheetByName("Do Not Contact");
  const data = dncSheet.getDataRange().getValues();
  
  const dncEmails = [];
  
  // Mulai dari baris 2 (skip header)
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      dncEmails.push(data[i][0].toLowerCase().trim());
    }
  }
  
  return dncEmails;
}

// Cek apakah email ada di DNC list
function isInDNC(email) {
  const dncList = getDNCList();
  return dncList.includes(email.toLowerCase().trim());
}

// Hitung selisih hari antara dua tanggal
function daysBetween(date1, date2) {
  const d1 = new Date(date1);
  const d2 = new Date(date2);
  const diffTime = Math.abs(d2 - d1);
  const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
  return diffDays;
}

// Update status di kolom V (kolom ke-22) di Calls for Proposals
function updateStatus(row, status) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Calls for Proposals");
  sheet.getRange(row, 22).setValue(status); // Kolom V = kolom 22
}

// Update error log di kolom W (kolom ke-23) di Calls for Proposals
function updateErrorLog(row, message) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Calls for Proposals");
  sheet.getRange(row, 23).setValue(message); // Kolom W = kolom 23
}

// Test DNC function
function testDNC() {
  Logger.log(getDNCList());
  Logger.log(isInDNC("blocked@example.com")); // Harus true
  Logger.log(isInDNC("random@email.com"));    // Harus false
}

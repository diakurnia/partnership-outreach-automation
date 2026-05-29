// Config.gs
// Mengambil semua settings dari tab Settings di sheet

function getConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName("Settings");
  const data = settingsSheet.getDataRange().getValues();
  
  const config = {};
  
  // Loop semua baris, baris pertama adalah header jadi mulai dari index 1
  for (let i = 1; i < data.length; i++) {
    const key = data[i][0];   // Kolom A = nama setting
    const value = data[i][1]; // Kolom B = value setting
    if (key) {
      config[key] = value;
    }
  }
  
  return config;
}

// Test function — jalankan ini untuk cek apakah config terbaca
function testConfig() {
  const config = getConfig();
  Logger.log(config);
}

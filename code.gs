// Code.gs
// Main function — memproses baris baru dan generate email drafts

function procesNewOpportunities() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfpSheet = ss.getSheetByName("Calls for Proposals");
  const partnersSheet = ss.getSheetByName("Partners");
  
  const cfpData = cfpSheet.getDataRange().getValues();
  const partnersData = partnersSheet.getDataRange().getValues();
  
  // Loop semua baris di Calls for Proposals, mulai dari baris 3 (skip 2 header)
  for (let i = 2; i < cfpData.length; i++) {
    const row = cfpData[i];
    const rowNumber = i + 1; // Apps Script rows are 1-indexed
    
    const callName     = row[0];  // Kolom A
    const callId       = row[1];  // Kolom B
    const intro        = row[2];  // Kolom C
    const deadline     = row[3];  // Kolom D
    const partnerEmail = row[5];  // Kolom F
    const status       = row[21]; // Kolom V
    
    // Skip kalau bukan status "New" atau baris kosong
    if (!callName || status !== "New") continue;
    
    Logger.log(`Processing row ${rowNumber}: ${callName}`);
    
    try {
      // === CEK 1: Do Not Contact ===
      if (isInDNC(partnerEmail)) {
        updateStatus(rowNumber, "DNC — blocked");
        updateErrorLog(rowNumber, "Partner ada di Do Not Contact list");
        continue;
      }
      
      // === CEK 2: Duplicate Detection Rule 1 ===
      if (isDuplicate(callId, partnerEmail, cfpData, i)) {
        updateStatus(rowNumber, "Duplicate — skipped");
        updateErrorLog(rowNumber, "Kombinasi partner + opportunity ini sudah pernah diproses");
        continue;
      }
      
      // === CEK 3: 30-Day Cooldown ===
      const cooldownResult = isInCooldown(partnerEmail, partnersData);
      if (cooldownResult.inCooldown) {
        updateStatus(rowNumber, "Cooldown active");
        updateErrorLog(rowNumber, `Partner terakhir dihubungi ${cooldownResult.daysAgo} hari lalu. Tunggu ${cooldownResult.daysLeft} hari lagi.`);
        continue;
      }
      
      // === AMBIL DATA PARTNER ===
      const partner = getPartnerByEmail(partnerEmail, partnersData);
      if (!partner) {
        updateStatus(rowNumber, "Failed — retry");
        updateErrorLog(rowNumber, `Partner dengan email ${partnerEmail} tidak ditemukan di tab Partners`);
        logError(rowNumber, "Calls for Proposals", `Partner tidak ditemukan: ${partnerEmail}`);
        continue;
      }
      
      // === GENERATE EMAIL VIA CLAUDE ===
      const opportunityData = { callName, callId, intro, deadline };
      const emailDraft = generateEmailWithClaude(partner, opportunityData);
      
      // === TULIS DRAFT KE SHEET ===
      cfpSheet.getRange(rowNumber, 7).setValue(emailDraft.subject); // Kolom G
      cfpSheet.getRange(rowNumber, 8).setValue(emailDraft.body);    // Kolom H
      updateStatus(rowNumber, "Draft generated");
      
      Logger.log(`✓ Draft generated for row ${rowNumber}`);
      
    } catch(e) {
      updateStatus(rowNumber, "Failed — retry");
      updateErrorLog(rowNumber, e.message);
      logError(rowNumber, "Calls for Proposals", e.message);
      Logger.log(`✗ Error on row ${rowNumber}: ${e.message}`);
    }
  }
}

// Cek duplikat: kombinasi callId + partnerEmail
function isDuplicate(callId, partnerEmail, allData, currentIndex) {
  for (let i = 2; i < allData.length; i++) {
    if (i === currentIndex) continue; // Skip baris saat ini
    if (allData[i][1] === callId && allData[i][5] === partnerEmail) {
      return true;
    }
  }
  return false;
}

// Cek cooldown 30 hari
function isInCooldown(partnerEmail, partnersData) {
  for (let i = 1; i < partnersData.length; i++) {
    if (partnersData[i][2].toLowerCase().trim() === partnerEmail.toLowerCase().trim()) {
      const lastContacted = partnersData[i][4]; // Kolom E = Date Last Contacted
      
      if (!lastContacted) return { inCooldown: false };
      
      const daysAgo = daysBetween(lastContacted, new Date());
      const config = getConfig();
      const cooldownDays = parseInt(config["Cooldown Period (days)"]) || 30;
      
      if (daysAgo < cooldownDays) {
        return {
          inCooldown: true,
          daysAgo: daysAgo,
          daysLeft: cooldownDays - daysAgo
        };
      }
    }
  }
  return { inCooldown: false };
}

// Ambil data partner berdasarkan email
function getPartnerByEmail(email, partnersData) {
  for (let i = 1; i < partnersData.length; i++) {
    if (partnersData[i][2].toLowerCase().trim() === email.toLowerCase().trim()) {
      return {
        orgName: partnersData[i][0],
        contactPerson: partnersData[i][1],
        email: partnersData[i][2],
        description: partnersData[i][3]
      };
    }
  }
  return null;
}

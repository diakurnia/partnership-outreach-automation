// Gmail.gs
// Semua fungsi untuk kirim email, tracking reply, dan follow-up via Gmail API

// =============================================
// BAGIAN 1 — KIRIM EMAIL ORIGINAL
// =============================================

function sendApprovedEmails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfpSheet = ss.getSheetByName("Calls for Proposals");
  const partnersSheet = ss.getSheetByName("Partners");
  const cfpData = cfpSheet.getDataRange().getValues();
  const partnersData = partnersSheet.getDataRange().getValues();

  for (let i = 2; i < cfpData.length; i++) {
    const row = cfpData[i];
    const rowNumber = i + 1;

    const partnerEmail = row[5];  // Kolom F
    const subject      = row[6];  // Kolom G
    const body         = row[7];  // Kolom H
    const approveToSend = row[8]; // Kolom I
    const sentAt       = row[9];  // Kolom J
    const status       = row[21]; // Kolom V

    // Hanya proses kalau:
    // - Approve to Send = "Yes"
    // - Belum pernah dikirim (sentAt kosong)
    // - Status = "Draft generated"
    if (approveToSend !== "Yes" || sentAt || status !== "Draft generated") continue;

    try {
      // === CEK DNC sebelum kirim ===
      if (isInDNC(partnerEmail)) {
        updateStatus(rowNumber, "DNC — blocked");
        updateErrorLog(rowNumber, "Diblokir DNC saat hendak kirim");
        continue;
      }

      // === KIRIM EMAIL ===
      const gmailThread = GmailApp.sendEmail(
        partnerEmail,
        subject,
        body,
        {
          name: getConfig()["Sender Name"] || "Partnership Team"
        }
      );

      // Ambil Thread ID dari email yang baru dikirim
      // Cari thread terbaru ke partner tersebut
      const threads = GmailApp.search(`to:${partnerEmail} subject:"${subject}"`, 0, 1);
      const threadId = threads.length > 0 ? threads[0].getId() : "";

      // === CATAT KE SHEET ===
      cfpSheet.getRange(rowNumber, 10).setValue(new Date()); // Kolom J = Sent At
      cfpSheet.getRange(rowNumber, 11).setValue(threadId);   // Kolom K = Gmail Thread ID
      updateStatus(rowNumber, "Sent");

      // === UPDATE DATE LAST CONTACTED di Partners tab ===
      updatePartnerLastContacted(partnerEmail, partnersData, partnersSheet);

      Logger.log(`✓ Email sent to ${partnerEmail} — Row ${rowNumber}`);

    } catch(e) {
      updateStatus(rowNumber, "Failed — retry");
      updateErrorLog(rowNumber, "Gagal kirim email: " + e.message);
      logError(rowNumber, "Calls for Proposals", "Send failed: " + e.message);
      Logger.log(`✗ Failed to send row ${rowNumber}: ${e.message}`);
    }
  }
}


// =============================================
// BAGIAN 2 — CEK REPLY MASUK
// =============================================

function checkForReplies() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfpSheet = ss.getSheetByName("Calls for Proposals");
  const cfpData = cfpSheet.getDataRange().getValues();

  for (let i = 2; i < cfpData.length; i++) {
    const row = cfpData[i];
    const rowNumber = i + 1;

    const threadId      = row[10]; // Kolom K = Gmail Thread ID
    const status        = row[21]; // Kolom V
    const replyReceived = row[19]; // Kolom T = Reply Received At

    // Hanya cek row yang sudah dikirim dan belum ada reply
    const activeStatuses = ["Sent", "Follow-up 1 sent", "Follow-up 2 sent"];
    if (!threadId || !activeStatuses.includes(status) || replyReceived) continue;

    try {
      const thread = GmailApp.getThreadById(threadId);
      if (!thread) continue;

      const messages = thread.getMessages();

      // Kalau jumlah pesan lebih dari 1, berarti ada balasan
      if (messages.length > 1) {
        // Ambil pesan terakhir (balasan dari partner)
        const lastMessage = messages[messages.length - 1];
        const replySnippet = lastMessage.getPlainBody().substring(0, 200);

        // Catat reply ke sheet
        cfpSheet.getRange(rowNumber, 20).setValue(new Date(lastMessage.getDate())); // Kolom T
        cfpSheet.getRange(rowNumber, 21).setValue(replySnippet);                    // Kolom U
        updateStatus(rowNumber, "Reply received");

        Logger.log(`✓ Reply detected on row ${rowNumber}`);
      }

    } catch(e) {
      logError(rowNumber, "Calls for Proposals", "Reply check failed: " + e.message);
      Logger.log(`✗ Reply check error row ${rowNumber}: ${e.message}`);
    }
  }
}


// =============================================
// BAGIAN 3 — GENERATE & KIRIM FOLLOW-UP 1
// =============================================

function processFollowUp1() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfpSheet = ss.getSheetByName("Calls for Proposals");
  const partnersSheet = ss.getSheetByName("Partners");
  const cfpData = cfpSheet.getDataRange().getValues();
  const partnersData = partnersSheet.getDataRange().getValues();
  const config = getConfig();
  const followUp1Days = parseInt(config["Follow-up 1 Delay (days)"]) || 7;

  for (let i = 2; i < cfpData.length; i++) {
    const row = cfpData[i];
    const rowNumber = i + 1;

    const callName      = row[0];  // Kolom A
    const callId        = row[1];  // Kolom B
    const intro         = row[2];  // Kolom C
    const deadline      = row[3];  // Kolom D
    const partnerEmail  = row[5];  // Kolom F
    const sentAt        = row[9];  // Kolom J
    const fu1Subject    = row[11]; // Kolom L
    const approveFU1    = row[13]; // Kolom N
    const fu1SentAt     = row[14]; // Kolom O
    const status        = row[21]; // Kolom V

    // === STEP A: Generate draft follow-up 1 kalau belum ada ===
    if (status === "Sent" && sentAt && !fu1Subject) {
      const daysSinceSent = daysBetween(sentAt, new Date());

      if (daysSinceSent >= followUp1Days) {
        try {
          const partner = getPartnerByEmail(partnerEmail, partnersData);
          if (!partner) continue;

          const opportunityData = { callName, callId, intro, deadline };
          const fu1Draft = generateFollowUpWithClaude(partner, opportunityData, 1);

          cfpSheet.getRange(rowNumber, 12).setValue(fu1Draft.subject); // Kolom L
          cfpSheet.getRange(rowNumber, 13).setValue(fu1Draft.body);    // Kolom M
          Logger.log(`✓ Follow-up 1 draft generated for row ${rowNumber}`);

        } catch(e) {
          logError(rowNumber, "Calls for Proposals", "FU1 generate failed: " + e.message);
        }
      }
    }

    // === STEP B: Kirim follow-up 1 kalau sudah di-approve ===
    if (approveFU1 === "Yes" && fu1Subject && !fu1SentAt && status === "Sent") {
      try {
        if (isInDNC(partnerEmail)) {
          updateStatus(rowNumber, "DNC — blocked");
          continue;
        }

        // Kirim sebagai reply di thread yang sama
        const threadId = row[10]; // Kolom K
        const thread = GmailApp.getThreadById(threadId);
        thread.reply(row[12], { subject: fu1Subject }); // Kolom M = body

        cfpSheet.getRange(rowNumber, 15).setValue(new Date()); // Kolom O = FU1 Sent At
        updateStatus(rowNumber, "Follow-up 1 sent");
        updatePartnerLastContacted(partnerEmail, partnersData, partnersSheet);

        Logger.log(`✓ Follow-up 1 sent for row ${rowNumber}`);

      } catch(e) {
        updateStatus(rowNumber, "Failed — retry");
        updateErrorLog(rowNumber, "FU1 send failed: " + e.message);
        logError(rowNumber, "Calls for Proposals", "FU1 send failed: " + e.message);
      }
    }
  }
}


// =============================================
// BAGIAN 4 — GENERATE & KIRIM FOLLOW-UP 2
// =============================================

function processFollowUp2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfpSheet = ss.getSheetByName("Calls for Proposals");
  const partnersSheet = ss.getSheetByName("Partners");
  const cfpData = cfpSheet.getDataRange().getValues();
  const partnersData = partnersSheet.getDataRange().getValues();
  const config = getConfig();
  const followUp2Days = parseInt(config["Follow-up 2 Delay (days)"]) || 14;

  for (let i = 2; i < cfpData.length; i++) {
    const row = cfpData[i];
    const rowNumber = i + 1;

    const callName      = row[0];
    const callId        = row[1];
    const intro         = row[2];
    const deadline      = row[3];
    const partnerEmail  = row[5];
    const sentAt        = row[9];  // Kolom J = original sent at
    const fu2Subject    = row[15]; // Kolom P
    const approveFU2    = row[17]; // Kolom R
    const fu2SentAt     = row[18]; // Kolom S
    const status        = row[21]; // Kolom V

    // === STEP A: Generate draft follow-up 2 ===
    if (status === "Follow-up 1 sent" && sentAt && !fu2Subject) {
      const daysSinceSent = daysBetween(sentAt, new Date());

      if (daysSinceSent >= followUp2Days) {
        try {
          const partner = getPartnerByEmail(partnerEmail, partnersData);
          if (!partner) continue;

          const opportunityData = { callName, callId, intro, deadline };
          const fu2Draft = generateFollowUpWithClaude(partner, opportunityData, 2);

          cfpSheet.getRange(rowNumber, 16).setValue(fu2Draft.subject); // Kolom P
          cfpSheet.getRange(rowNumber, 17).setValue(fu2Draft.body);    // Kolom Q
          Logger.log(`✓ Follow-up 2 draft generated for row ${rowNumber}`);

        } catch(e) {
          logError(rowNumber, "Calls for Proposals", "FU2 generate failed: " + e.message);
        }
      }
    }

    // === STEP B: Kirim follow-up 2 kalau sudah di-approve ===
    if (approveFU2 === "Yes" && fu2Subject && !fu2SentAt && status === "Follow-up 1 sent") {
      try {
        if (isInDNC(partnerEmail)) {
          updateStatus(rowNumber, "DNC — blocked");
          continue;
        }

        const threadId = row[10];
        const thread = GmailApp.getThreadById(threadId);
        thread.reply(row[16], { subject: fu2Subject }); // Kolom Q = body

        cfpSheet.getRange(rowNumber, 19).setValue(new Date()); // Kolom S = FU2 Sent At
        updateStatus(rowNumber, "Follow-up 2 sent");
        updatePartnerLastContacted(partnerEmail, partnersData, partnersSheet);

        Logger.log(`✓ Follow-up 2 sent for row ${rowNumber}`);

      } catch(e) {
        updateStatus(rowNumber, "Failed — retry");
        updateErrorLog(rowNumber, "FU2 send failed: " + e.message);
        logError(rowNumber, "Calls for Proposals", "FU2 send failed: " + e.message);
      }
    }

    // === STEP C: Auto-close kalau FU2 sudah dikirim dan tidak ada reply ===
    if (status === "Follow-up 2 sent" && fu2SentAt) {
      const daysSinceFU2 = daysBetween(fu2SentAt, new Date());
      if (daysSinceFU2 >= 7) { // Tunggu 7 hari setelah FU2
        updateStatus(rowNumber, "No response — closed");
        Logger.log(`✓ Row ${rowNumber} auto-closed — no response`);
      }
    }
  }
}


// =============================================
// BAGIAN 5 — HELPER FUNCTIONS
// =============================================

// Generate follow-up email via Claude
function generateFollowUpWithClaude(partnerData, opportunityData, followUpNumber) {
  const config = getConfig();
  const apiKey = config["Claude API Key"];

  const prompt = `You are a partnership outreach specialist writing a follow-up email.
This is follow-up number ${followUpNumber}.
Partner: ${partnerData.contactPerson} from ${partnerData.orgName}
Their focus: ${partnerData.description}
Opportunity: ${opportunityData.callName} (${opportunityData.callId})
Details: ${opportunityData.intro}
Deadline: ${opportunityData.deadline}

Write a brief, friendly follow-up email. Keep it short (3-4 sentences max).
Do not be pushy. Acknowledge they are busy.
Format your response exactly like this:
SUBJECT: [subject line here]
BODY: [email body here]`;

  const url = "https://api.anthropic.com/v1/messages";

  const payload = {
    model: "claude-sonnet-4-5",
    max_tokens: 500,
    messages: [{ role: "user", content: prompt }]
  };

  const options = {
    method: "post",
    headers: {
      "Content-Type": "application/json",
      "x-api-key": apiKey,
      "anthropic-version": "2023-06-01"
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const result = JSON.parse(response.getContentText());
  const generatedText = result.content[0].text;

  return parseEmailFromClaude(generatedText);
}

// Update Date Last Contacted di tab Partners
function updatePartnerLastContacted(email, partnersData, partnersSheet) {
  for (let i = 1; i < partnersData.length; i++) {
    if (partnersData[i][2].toLowerCase().trim() === email.toLowerCase().trim()) {
      partnersSheet.getRange(i + 1, 5).setValue(new Date()); // Kolom E
      Logger.log(`✓ Updated last contacted for ${email}`);
      return;
    }
  }
}


// =============================================
// BAGIAN 6 — TEST FUNCTIONS
// =============================================

function testSendEmail() {
  // Pastikan ada 1 baris dengan status "Draft generated" 
  // dan Approve to Send = "Yes" sebelum test ini
  Logger.log("Running sendApprovedEmails...");
  sendApprovedEmails();
  Logger.log("Done. Cek tab Calls for Proposals.");
}

function testCheckReplies() {
  Logger.log("Checking for replies...");
  checkForReplies();
  Logger.log("Done. Cek kolom T, U, V.");
}

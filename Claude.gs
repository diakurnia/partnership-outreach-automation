// Claude.gs
// Semua komunikasi dengan Claude API

function generateEmailWithClaude(partnerData, opportunityData) {
  const config = getConfig();
  const apiKey = config["Claude API Key"];
  
  if (!apiKey) {
    throw new Error("Claude API Key tidak ditemukan di Settings!");
  }
  
  // Ambil prompt template dari settings
  let promptTemplate = config["Email Prompt Template"];
  
  // Ganti placeholder dengan data sebenarnya
  const prompt = promptTemplate
    .replace("{contact_person}", partnerData.contactPerson)
    .replace("{org_name}", partnerData.orgName)
    .replace("{description}", partnerData.description)
    .replace("{call_name}", opportunityData.callName)
    .replace("{call_id}", opportunityData.callId)
    .replace("{intro}", opportunityData.intro)
    .replace("{deadline}", opportunityData.deadline);
  
  // Kirim request ke Claude API
  const url = "https://api.anthropic.com/v1/messages";
  
  const payload = {
    model: "claude-sonnet-4-5",
    max_tokens: 1000,
    messages: [
      {
        role: "user",
        content: prompt
      }
    ]
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
  const responseCode = response.getResponseCode();
  const responseText = response.getContentText();
  
  if (responseCode !== 200) {
    throw new Error(`Claude API error ${responseCode}: ${responseText}`);
  }
  
  const result = JSON.parse(responseText);
  const generatedText = result.content[0].text;
  
  // Parse subject dan body dari response Claude
  return parseEmailFromClaude(generatedText);
}

// Pisahkan subject dan body dari response Claude
function parseEmailFromClaude(text) {
  let subject = "";
  let body = "";
  
  // Cari baris yang mulai dengan "SUBJECT:"
  const subjectMatch = text.match(/SUBJECT:\s*(.+)/i);
  if (subjectMatch) {
    subject = subjectMatch[1].trim();
  }
  
  // Ambil semua teks setelah "BODY:"
  const bodyMatch = text.match(/BODY:\s*([\s\S]+)/i);
  if (bodyMatch) {
    body = bodyMatch[1].trim();
  }
  
  return { subject, body };
}

// Test Claude API connection
function testClaudeAPI() {
  const partnerData = {
    contactPerson: "Budi Santoso",
    orgName: "WWF Indonesia",
    description: "Organisasi fokus pada lingkungan hidup dan konservasi alam"
  };
  
  const opportunityData = {
    callName: "EU Green Fund 2026",
    callId: "EGF-2026-001",
    intro: "Hibah €500k untuk proyek konservasi dan lingkungan hidup di Asia Tenggara",
    deadline: "30/09/2026"
  };
  
  try {
    const result = generateEmailWithClaude(partnerData, opportunityData);
    Logger.log("Subject: " + result.subject);
    Logger.log("Body: " + result.body);
    Logger.log("SUCCESS!");
  } catch(e) {
    Logger.log("ERROR: " + e.message);
  }
}

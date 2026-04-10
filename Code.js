// ============================================================
// AUTOMATED ZELLE PAYMENT LEDGER
// ============================================================

const CONFIG = {
  
  SHEET_NAME: "History",               
  SUMMARY_SHEET_NAME: "Summary",       
  FIRST_MONTH_COL: 6,                  
  GMAIL_QUERY: "from:no.reply.alerts@chase.com -label:Processed", // could be different depending on how your bank account recieves zelle
  PROCESSED_LABEL: "Zelle-Processed",
  REMINDER_SUBJECT: "Friendly Reminder: Subscription Balance",
  SENDER_NAME: "Subscription Ledger",
  MONTHS_TO_GENERATE: 48,              
  START_MONTH: 0,                      
  START_YEAR: 2025, // this could be subjected to change 
};

// 1. SHEET SETUP
function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let historySheet = ss.getSheetByName(CONFIG.SHEET_NAME) || ss.insertSheet(CONFIG.SHEET_NAME);

  const historyHeaders = ["Name", "Email", "Phone", "Monthly Rate ($)", "Memo Keyword"];
  const startDate = new Date(CONFIG.START_YEAR, CONFIG.START_MONTH, 1);

  for (let i = 0; i < CONFIG.MONTHS_TO_GENERATE; i++) {
    const d = new Date(startDate.getFullYear(), startDate.getMonth() + i, 1);
    const label = Utilities.formatDate(d, Session.getScriptTimeZone(), "MMM yyyy");
    historyHeaders.push(label);
  }

  // Add extra columns if the sheet isn't wide enough for 48 months
  if (historySheet.getMaxColumns() < historyHeaders.length) {
    historySheet.insertColumnsAfter(historySheet.getMaxColumns(), historyHeaders.length - historySheet.getMaxColumns());
  }

  historySheet.getRange(1, 1, 1, historyHeaders.length).setValues([historyHeaders]);
  historySheet.getRange(1, 1, 1, historyHeaders.length).setFontWeight("bold").setBackground("#1a1a2e").setFontColor("#ffffff");

  const monthCols = historyHeaders.length - 5; 
  if (monthCols > 0) {
    historySheet.getRange(2, CONFIG.FIRST_MONTH_COL, 5, monthCols).insertCheckboxes();
  }
  
  historySheet.setFrozenRows(1);
  historySheet.setFrozenColumns(5);

  let summarySheet = ss.getSheetByName(CONFIG.SUMMARY_SHEET_NAME) || ss.insertSheet(CONFIG.SUMMARY_SHEET_NAME, 0);
  const summaryHeaders = ["Name", "Monthly Rate", "Months Paid", "Months Elapsed", "Total Balance", "Current Status"];
  summarySheet.getRange(1, 1, 1, summaryHeaders.length).setValues([summaryHeaders]);

  for (let row = 2; row <= 6; row++) {
    summarySheet.getRange(row, 1).setFormula(`=IF(ISBLANK(History!A${row}), "", History!A${row})`);
    summarySheet.getRange(row, 2).setFormula(`=IF(ISBLANK(History!D${row}), "", History!D${row})`);
    summarySheet.getRange(row, 3).setFormula(`=IF(ISBLANK(A${row}), "", COUNTIF(History!F${row}:ZZ${row}, TRUE))`);
    
    // Formula now pulls the year dynamically from the CONFIG block
    summarySheet.getRange(row, 4).setFormula(`=IF(ISBLANK(A${row}), "", DATEDIF("${CONFIG.START_YEAR}-01-01", TODAY(), "M") + 1)`);
    summarySheet.getRange(row, 5).setFormula(`=IF(ISBLANK(A${row}), "", (D${row} - C${row}) * B${row})`);
    summarySheet.getRange(row, 6).setFormula(`=IF(ISBLANK(A${row}), "", IF(C${row}=D${row}, "✅ Up to Date", IF(C${row}>D${row}, C${row}-D${row} & " Ahead", D${row}-C${row} & " Behind")))`);
  }
  
  // Safe alert handling
  try {
    SpreadsheetApp.getUi().alert("✅ Setup Complete for " + CONFIG.START_YEAR + "!");
  } catch (e) {
    console.log("✅ Setup Complete! Check your spreadsheet tabs.");
  }
}

// 2. GMAIL SCANNER
function scanZelleEmails() {
  console.log("Starting Zelle email scan...");

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    console.log("❌ Error: Could not find sheet [" + CONFIG.SHEET_NAME + "]");
    return;
  }

  const memory = PropertiesService.getScriptProperties();
  
  // Clean up any memory older than 30 days so we never run out of space
  const allMemory = memory.getProperties();
  const now = Date.now();
  const thirtyDaysInMs = 30 * 24 * 60 * 60 * 1000;
  let clearedCount = 0;

  for (const id in allMemory) {
    const timestamp = parseInt(allMemory[id], 10);
    if (now - timestamp > thirtyDaysInMs) {
      memory.deleteProperty(id);
      clearedCount++;
    }
  }
  if (clearedCount > 0) console.log(`🧹 Memory cleanup: Deleted ${clearedCount} old records.`);
  // -----------------------------

  const threads = GmailApp.search(CONFIG.GMAIL_QUERY, 0, 50);
  console.log("🔍 Found " + threads.length + " threads matching your query.");

  let processedCount = 0;

  for (const thread of threads) {
    const messages = thread.getMessages();

    for (const message of messages) {
      const messageId = message.getId();
      
      // Check if this ID is currently in our active memory
      if (memory.getProperty(messageId)) {
        continue; 
      }

      console.log("Processing a brand new payment email...");
      
      const fullText = message.getSubject() + " " + (message.getPlainBody() || message.getBody().replace(/<[^>]+>/g, " "));
      const payment = parseZellePayment(fullText);
      
      if (payment) {
        matchAndRecord(sheet, payment);
      }
      
      // Save the ID with the CURRENT TIMESTAMP instead of just a text string
      memory.setProperty(messageId, Date.now().toString());
      processedCount++;
    }
    
    thread.moveToArchive();
  }

  console.log("🏁 Finished scanning. Processed " + processedCount + " new messages.");
}
function parseZellePayment(text) {
  // 1. Clean the text into one long string
  const cleaned = text.replace(/\s+/g, " ").trim();
  console.log("🔍 Read Email Text: " + cleaned.substring(0, 150) + "..."); 
  
  // 2. Find the Amount
  const amountMatch = cleaned.match(/\$\s?([\d,]+\.?\d{0,2})/);
  if (!amountMatch) {
    console.log("❌ Parsing Failed: Could not find a dollar amount in the text.");
    return null;
  }
  const amount = parseFloat(amountMatch[1].replace(/,/g, ""));

  // 3. Find the Name 
  let name = null;
  
  // Added the hyphen \- to the character list so it catches hyphenated names
  const patterns = [
    /(?:payment\.\s+)?([A-Za-z\s\-]{2,50})\s+sent\s+you/i, 
    /([A-Za-z\s\-]{2,50})\s+sent\s+you/i,                  
    /(?:from|From)\s+([A-Za-z\s\-]{2,50})/i                
  ];

  for (const pattern of patterns) {
    const match = cleaned.match(pattern);
    if (match) { 
      name = match[1].replace(/Zelle®/gi, "").replace(/payment/gi, "").replace(/\./g, "").trim();
      break; 
    }
  }

  // 4. Final Verification
  if (!name) {
    console.log(`❌ Parsing Failed: Found amount ($${amount}) but could not isolate a name.`);
    return null;
  }

  console.log(`✅ Parsing Success: Found payment from [${name}] for [$${amount}]`);
  return { name, amount };
}
function matchAndRecord(sheet, payment) {
  const historyData = sheet.getDataRange().getValues();
  let matchFound = false;

  for (let r = 1; r < historyData.length; r++) {
    const rowName = String(historyData[r][0]).toLowerCase().trim(); // Column A

    if (!rowName) continue; // Skip empty rows

    // Check if the name from the email matches the History row
    if (payment.name.toLowerCase().includes(rowName)) {
      console.log(`Found name [${rowName}] in History. Calculating months paid...`);
      
      const monthlyRate = parseFloat(historyData[r][3]); // Column D
      if (isNaN(monthlyRate) || monthlyRate <= 0) {
        console.log(`❌ Error: Row ${r+1} does not have a valid monthly rate in Column D.`);
        continue;
      }

      // Divide the payment by the monthly rate to get the number of months
      const monthsPaid = Math.round(payment.amount / monthlyRate);
      const expectedTotal = monthsPaid * monthlyRate;

      // Verify the math checks out
      if (Math.abs(payment.amount - expectedTotal) < 0.1 && monthsPaid > 0) {
        console.log(`✅ Math checks out! $${payment.amount} pays for ${monthsPaid} month(s) at $${monthlyRate}/mo.`);
        
        let monthsLeftToFill = monthsPaid;
        let columnsUpdated = 0;

        // Start sweeping across the columns starting at Column F (Index 5)
        for (let c = 5; c < historyData[0].length; c++) {
          let cellValue = historyData[r][c];
          
          // If the cell is empty or unchecked, check it
          if (cellValue !== true) {
            sheet.getRange(r + 1, c + 1).setValue(true);
            monthsLeftToFill--;
            columnsUpdated++;
          }
          
          // Stop once we have filled the required number of empty boxes
          if (monthsLeftToFill <= 0) break;
        }

        SpreadsheetApp.flush(); // Force the checkboxes to appear visually instantly
        console.log(`✅ Success: Filled ${columnsUpdated} empty gap(s) for ${rowName}.`);
        matchFound = true;
        break;
        
      } else {
        console.log(`❌ Amount mismatch. $${payment.amount} is not a clean multiple of the $${monthlyRate} monthly rate.`);
      }
    }
  }

  if (!matchFound) {
    console.log(`❌ Done checking. Could not find a complete match for [${payment.name}] with the correct amount.`);
  }
}

// 3. DISCORD & TRIGGERS
function sendDiscordReminders() {
  const WEBHOOK = PropertiesService.getScriptProperties().getProperty('DISCORD_WEBHOOK_URL'); // go to project settings and add a script property with the value being the webhook url
  if (!WEBHOOK || WEBHOOK === "YOUR_DISCORD_WEBHOOK_URL_HERE") return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const summary = ss.getSheetByName("Summary");
  const data = summary.getRange(2, 1, summary.getLastRow() - 1, 7).getValues();

  data.forEach(row => {
    if (row[4] > 0) { // If Balance > 0
      const mention = row[6] ? `<@${row[6]}>` : row[0];
      const payload = {
        content: `⚠️ ${mention}: Balance of $${row[4].toFixed(2)} is due.`,
        username: "Ledger Bot"
      };
      UrlFetchApp.fetch(WEBHOOK, { method: "post", contentType: "application/json", payload: JSON.stringify(payload) });
    }
  });
}

function createTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => ScriptApp.deleteTrigger(t));

  // Hourly scanner
  ScriptApp.newTrigger("scanZelleEmails").timeBased().everyHours(1).create();
  
  // Monthly Discord
  ScriptApp.newTrigger("sendDiscordReminders").timeBased().onMonthDay(22).atHour(10).create();
  
  console.log("✅ Triggers active.");
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu("💰 Ledger Tools")
    .addItem("📋 Setup Sheet", "setupSheet")
    .addItem("📧 Scan Emails Now", "scanZelleEmails")
    .addItem("💬 Discord Reminders Now", "sendDiscordReminders")
    .addItem("⚙️ Reset Triggers", "createTriggers")
    .addToUi();
}
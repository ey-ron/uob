require('dotenv').config();
const { chromium } = require('playwright');
const fs = require('fs');
const XLSX = require('xlsx');

const scrubDescription = (desc) => {
  return String(desc)
    .replace(/ref\s*no[\.\:\s]*\d*/gi, '')
    .replace(/singapore/gi, '')
    .replace(/\bsg\b/gi, '')
    .replace(/\+65/gi, '')
    .replace(/[-,\s]+$/g, '')
    .replace(/\s{2,}/g, ' ')
    .trim();
};

const express = require('express');
const app = express();
const port = process.env.PORT || 3000;

// This wrapper allows the script to be triggered via an HTTP request
app.post('/trigger-uob', async (req, res) => {
  console.log("Trigger received from iPhone...");
  
  try {
    const data = await downloadStatement();
    res.status(200).json(data);
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

app.listen(port, () => {
  console.log(`Server running on port ${port}`);
});

async function downloadStatement() {
  console.log("Starting Playwright...");
  // headless: false allows you to see the browser window
  // slowMo: 500 adds a slight delay so you can follow the actions
  const browser = await chromium.launch({ 
    headless: true, 
    slowMo: 500 
  });
  const context = await browser.newContext();
  const page = await context.newPage();

  // 1. Go to UOB Personal Internet Banking
  await page.goto('https://pib.uob.com.sg/PIBLogin/Public/processPreCapture.do?keyId=lpc');

  // 2. Login Sequence
  await page.waitForSelector('#userName'); 
  await page.fill('#userName', process.env.UOB_USERNAME);
  
  await page.waitForSelector('#PASSWORD1');
  await page.fill('#PASSWORD1', process.env.UOB_PASSWORD);
  
  // 3. Submit Login & Handle Intermediate "Proceed" steps
  await page.waitForSelector('#btnSubmit');
  await page.click('#btnSubmit');

  try {
    const proceedBtn = page.locator('button#btnsubmit:has-text("Proceed")');
    if (await proceedBtn.isVisible({ timeout: 5000 })) {
      console.log("Found 'Proceed' button. Clicking to trigger 2FA...");
      await proceedBtn.click();
    }
  } catch (e) {
    console.log("No intermediate 'Proceed' button found, moving to 2FA check.");
  }

  console.log("Check your iPhone 15 Pro for the 2FA push notification...");

  // 4. Wait for 2FA Approval and verify Dashboard Access
  console.log("Awaiting 2FA approval on your iPhone...");
  try {
    console.log("Waiting for the dashboard to load...");
    await page.waitForSelector('b:has-text("Welcome to UOB Personal Internet Banking")', { timeout: 60000 });
    
    await page.waitForLoadState('networkidle');
    console.log("Login successful! I see the Welcome message.");
  } catch (e) {
    console.error("2FA Timeout or Denied");
    await browser.close();
    return;
  }

  // 5. Navigate to Card Transactions
  console.log("Navigating to card transactions...");
  await page.click('a:has-text("UOB EVOL")');

  // 6. Configure Statement View (New Transactions)
  await page.waitForSelector('#frequency-account-summary');
  await page.selectOption('#frequency-account-summary', '0'); 
  console.log("Dropdown set to: New Transactions Since Last Statement");

  // Wait for the UI/Table to stabilize after the dropdown selection
  await page.waitForLoadState('networkidle');
  await page.waitForTimeout(2000); 

  // 7. Execution of File Download
  // Targets the icon-style download button specifically via classes
  const downloadBtn = page.locator('button#btnsubmit.btn-default.btn-icon');
  
  console.log("Attempting to click download button...");
  const [download] = await Promise.all([
    page.waitForEvent('download', { timeout: 60000 }), 
    downloadBtn.click()
  ]);

  const localPath = `./${download.suggestedFilename()}`;
  await download.saveAs(localPath);
  console.log(`File saved locally: ${localPath}`);

  // 8. Parsing logic for the downloaded XLS data
  try {
    console.log("\n=== STARTING PARSE ===");
    const buffer = fs.readFileSync(localPath);
    const workbook = XLSX.read(buffer, { type: 'buffer', cellDates: true });
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { header: 1 });
    
    let transactions = [];
    let cardName = "";
    const today = new Date();

    // Extract Header Info (Rows 0-10)
    for (let i = 0; i < Math.min(10, rows.length); i++) {
      const r = rows[i];
      if (r && String(r[0]).includes("Account Type")) cardName = String(r[1]).trim();
    }

    // Extract Transactions (Rows 10+)
    for (let i = 10; i < rows.length; i++) {
      const r = rows[i];
      if (!r || r.length < 3) continue;

      let amt = parseFloat(String(r[6]).replace(/,/g, ''));
      if (isNaN(amt) || amt === 0) continue;

      let finalDesc = scrubDescription(r[2] || "Unknown");
      const descUpper = finalDesc.toUpperCase();

      // Filter out payments and balances
      if (descUpper.includes('PREVIOUS BALANCE') || descUpper.includes('PAYMT')) continue;

      let badge = null;
      if (amt < 0) {
        badge = finalDesc.toLowerCase().includes('cashback') ? 'Cashback' : 'Refund';
        if (badge === 'Cashback') finalDesc = 'Cashback';
      }

      transactions.push({
        date: new Date(r[0] || today).toISOString(),
        description: finalDesc,
        currency: r[5] || "SGD",
        amount: amt,
        badge: badge
      });
    }

    console.log(`\nCard: ${cardName || "UOB EVOL"}`);
    console.log(`Transactions Found: ${transactions.length}`);
    console.table(transactions);
    console.log("\n=== PARSE COMPLETE ===");

  } catch (parseError) {
    console.error("Failed to parse the downloaded file:", parseError);
  }

  await browser.close();

  // Return the data and the file as base64 so the iPhone can save it
  const fileBuffer = fs.readFileSync(localPath);
  const base64File = fileBuffer.toString('base64');

  return {
    success: true,
    card: cardName || "UOB EVOL",
    count: transactions.length,
    transactions: transactions,
    fileName: download.suggestedFilename(),
    fileData: base64File
  };
}
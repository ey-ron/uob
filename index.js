require('dotenv').config();
const express = require('express');
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

const app = express();
app.use(express.json());

// Cloud Run supplies its own PORT env variable dynamically
const port = process.env.PORT || 8080;

app.post('/trigger-uob', async (req, res) => {
  // 🔒 Security Check
  const authHeader = req.headers['x-api-key'];
  if (!process.env.TRIGGER_SECRET || authHeader !== process.env.TRIGGER_SECRET) {
    return res.status(401).json({ error: 'Unauthorized' });
  }

  console.log("Trigger received from iPhone...");
  try {
    const data = await downloadStatement();
    res.status(200).json(data);
  } catch (error) {
    console.error("[CRITICAL ERROR]", error.message);
    res.status(500).json({ error: error.message });
  }
});

app.listen(port, () => {
  console.log(`Server running on port ${port}`);
});

async function downloadStatement() {
  let browser;
  let localPath = null;
  try {
    console.log("[DEBUG 1/7] Launching Chromium browser...");
    browser = await chromium.launch({
      headless: true,
      slowMo: 0,
      args: [
        '--no-sandbox',
        '--disable-setuid-sandbox',
        '--disable-dev-shm-usage',
        '--disable-accelerated-2d-canvas',
        '--no-first-run',
        '--no-zygote'
      ]
    });

    // Custom User-Agent prevents basic headless detection
    const context = await browser.newContext({
      userAgent: 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36'
    });
    const page = await context.newPage();

    // 1. Navigate to UOB PIB
    console.log("[DEBUG 2/7] Navigating to UOB Personal Internet Banking login page...");
    await page.goto('https://pib.uob.com.sg/PIBLogin/Public/processPreCapture.do?keyId=lpc', { waitUntil: 'networkidle', timeout: 30000 });
    console.log(`[DEBUG 2/7] Page landed. URL: ${page.url()} | Title: ${await page.title()}`);

    // 2. Fill Credentials
    console.log("[DEBUG 3/7] Filling username and password...");
    await page.waitForSelector('#userName', { timeout: 15000 });
    await page.fill('#userName', process.env.UOB_USERNAME);

    await page.waitForSelector('#PASSWORD1', { timeout: 15000 });
    await page.fill('#PASSWORD1', process.env.UOB_PASSWORD);
    console.log("[DEBUG 3/7] Credentials filled successfully.");

    await page.waitForTimeout(1000);

    // 3. Submit Login
    console.log("[DEBUG 4/7] Submitting login form...");
    await page.click('#btnSubmit');

    console.log("[DEBUG 4/7] Submitted login. Waiting for post-submit modal or 2FA trigger...");

    // Handle active session popup if present
    try {
      const proceedBtn = page.locator('button#btnsubmit:has-text("Proceed"), input[type="button"][value*="Proceed"], button:has-text("Proceed")');
      console.log("[DEBUG 4.5] Checking for stale active session...");
      await proceedBtn.waitFor({ state: 'visible', timeout: 8000 });
      console.log("[DEBUG 4.5] Stale active session detected! Clicking 'Proceed'...");
      await proceedBtn.click();
      await page.waitForTimeout(1000);
    } catch (e) {
      console.log("[DEBUG 4.5] No active session modal found or Proceed button did not appear, continuing...");
    }

    console.log("--------------------------------------------------");
    console.log("📲 CHECK YOUR PHONE NOW FOR THE UOB 2FA PUSH NOTIFICATION");
    console.log("--------------------------------------------------");

    // 4. Wait for ACTUAL Authenticated Dashboard
    // Strictly checks for post-login selectors (cards or summary container)
    console.log("[DEBUG 6/7] Awaiting 2FA approval on phone (timeout: 60s)...");
    try {
      await page.waitForSelector('a:has-text("UOB EVOL"), #account-summary, td.account-summary-header', { timeout: 60000 });
      console.log("[DEBUG 6/7] 2FA Approved! Authenticated dashboard loaded successfully!");
    } catch (e) {
      console.error("[DEBUG ERROR] 2FA approval timed out or dashboard failed to load. Final URL:", page.url());
      throw new Error(`2FA Timeout or Login Failed. Final URL: ${page.url()}`);
    }

    // 5. Navigate to Credit Card Page
    console.log("[DEBUG 7/7] Locating UOB EVOL card link...");
    await page.waitForSelector('a:has-text("UOB EVOL")', { timeout: 15000 });

    console.log("[DEBUG 7/7] Clicking 'UOB EVOL' via DOM evaluation...");
    await page.evaluate(() => {
      const links = Array.from(document.querySelectorAll('a'));
      const evolLink = links.find(el => el.textContent.includes('UOB EVOL'));
      if (evolLink) evolLink.click();
    });

    // 6. Configure Dropdown
    console.log("[DEBUG 7/7] Waiting for statement frequency dropdown (#frequency-account-summary)...");
    await page.waitForSelector('#frequency-account-summary', { timeout: 20000 });

    console.log("[DEBUG 7/7] Selecting '0' (New Transactions Since Last Statement)...");
    await page.selectOption('#frequency-account-summary', '0');

    await page.waitForLoadState('networkidle').catch(() => { });
    await page.waitForTimeout(2500);

    // 7. Execute File Download
    console.log("[DEBUG 7/7] Locating download button (button#btnsubmit.btn-default.btn-icon)...");
    const downloadBtn = page.locator('button#btnsubmit.btn-default.btn-icon');
    await downloadBtn.waitFor({ state: 'visible', timeout: 15000 });

    console.log("[DEBUG 7/7] Triggering download event...");
    const [download] = await Promise.all([
      page.waitForEvent('download', { timeout: 60000 }),
      downloadBtn.click()
    ]);

    localPath = `./${download.suggestedFilename()}`;
    await download.saveAs(localPath);
    console.log(`[DEBUG 7/7] File downloaded successfully: ${localPath}`);

    // 8. Parse XLS Data
    let transactions = [];
    let cardName = "";

    try {
      const buffer = fs.readFileSync(localPath);
      const workbook = XLSX.read(buffer, { type: 'buffer', cellDates: true });
      const rows = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { header: 1 });
      const today = new Date();

      for (let i = 0; i < Math.min(10, rows.length); i++) {
        const r = rows[i];
        if (r && String(r[0]).includes("Account Type")) cardName = String(r[1]).trim();
      }

      for (let i = 10; i < rows.length; i++) {
        const r = rows[i];
        if (!r || r.length < 3) continue;

        let amt = parseFloat(String(r[6]).replace(/,/g, ''));
        if (isNaN(amt) || amt === 0) continue;

        let finalDesc = scrubDescription(r[2] || "Unknown");
        const descUpper = finalDesc.toUpperCase();

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
    } catch (parseError) {
      console.error("[DEBUG ERROR] Failed parsing downloaded file:", parseError);
    }

    const fileBuffer = fs.readFileSync(localPath);
    const base64File = fileBuffer.toString('base64');
    const fileName = download.suggestedFilename();

    // Parse statement timestamp from file name: e.g., CC_TXN_History_25072026105048.xls
    let statementTimestamp = "";
    const match = fileName.match(/CC_TXN_History_(\d{2})(\d{2})(\d{4})(\d{2})(\d{2})(\d{2})/);
    if (match) {
      const [_, day, month, year, hour, minute, second] = match;
      const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
      const monthIndex = parseInt(month, 10) - 1;
      const monthStr = (monthIndex >= 0 && monthIndex < 12) ? months[monthIndex] : month;
      statementTimestamp = `${hour}:${minute}:${second} ${monthStr} ${day}, ${year}`;
    } else {
      const now = new Date();
      const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
      const pad = (num) => String(num).padStart(2, '0');
      statementTimestamp = `${pad(now.getHours())}:${pad(now.getMinutes())}:${pad(now.getSeconds())} ${months[now.getMonth()]} ${pad(now.getDate())}, ${now.getFullYear()}`;
    }

    return {
      success: true,
      card: cardName || "UOB EVOL",
      count: transactions.length,
      transactions: transactions,
      fileName: fileName,
      fileData: base64File,
      statementTimestamp: statementTimestamp
    };
  } finally {
    if (browser) {
      await browser.close();
    }
    if (localPath && fs.existsSync(localPath)) {
      try {
        fs.unlinkSync(localPath);
      } catch (err) {
        console.error("[DEBUG ERROR] Failed to clean up local file:", err);
      }
    }
  }
}
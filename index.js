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

app.all('/trigger-uob', async (req, res) => {
  // 🔒 Security Check (support headers, body, or query with full unicode/whitespace stripping)
  const rawAuth = (
    req.headers['x-api-key'] ||
    req.headers['authorization'] ||
    req.headers['apikey'] ||
    (req.body && (req.body.apiKey || req.body['x-api-key'])) ||
    req.query.apiKey ||
    ''
  );

  const cleanAuth = String(rawAuth).replace(/[\s\u00A0\u200B\u200C\u200D\uFEFF\r\n\t]/g, '');
  const expectedSecret = String(process.env.TRIGGER_SECRET || '').replace(/[\s\u00A0\u200B\u200C\u200D\uFEFF\r\n\t]/g, '');

  const alphanumericAuth = cleanAuth.replace(/[^a-zA-Z0-9]/g, '').toLowerCase();
  const alphanumericExpected = expectedSecret.replace(/[^a-zA-Z0-9]/g, '').toLowerCase();

  const isMatch = (
    cleanAuth === expectedSecret ||
    cleanAuth.includes(expectedSecret) ||
    (alphanumericExpected && alphanumericAuth === alphanumericExpected) ||
    (alphanumericExpected && alphanumericAuth.includes(alphanumericExpected))
  );

  if (!expectedSecret || !isMatch) {
    console.warn(`[AUTH REJECTED] Unauthorized request received. Provided key length: ${cleanAuth.length}, IP: ${req.ip || req.socket.remoteAddress}`);
    return res.status(401).json({ error: 'Unauthorized. Invalid or missing x-api-key.' });
  }

  // 🔀 Mode / Type Parameter: 'R' (Running Report Excel) or 'S' (e-Statement PDF)
  const rawMode = (
    req.query.type ||
    req.query.mode ||
    (req.body && (req.body.type || req.body.mode)) ||
    req.headers['x-type'] ||
    req.headers['x-mode'] ||
    'R'
  );

  const normalized = String(rawMode).trim().toUpperCase();
  const mode = (normalized === 'S') ? 'S' : 'R';
  console.log(`Trigger received from client. Execution mode: [${mode}]`);

  try {
    const data = await downloadStatement(mode);
    res.status(200).json(data);
  } catch (error) {
    console.error("[CRITICAL ERROR]", error.message);
    res.status(500).json({ error: error.message });
  }
});

app.listen(port, () => {
  console.log(`Server running on port ${port}`);
});

function formatStatementPdfFileName(rawFileName) {
  let formattedFileName = rawFileName;
  const monthMatch = rawFileName.match(/([A-Za-z]{3,4})\s+(\d{4})/i);
  if (monthMatch) {
    const rawMonth = monthMatch[1];
    const month = rawMonth.charAt(0).toUpperCase() + rawMonth.slice(1, 3).toLowerCase();
    const year = monthMatch[2];
    formattedFileName = `${month} ${year} Statement.pdf`;
  } else {
    const now = new Date();
    const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    const prevMonth = (now.getMonth() + 11) % 12;
    const year = prevMonth === 11 ? now.getFullYear() - 1 : now.getFullYear();
    formattedFileName = `${months[prevMonth]} ${year} Statement.pdf`;
  }
  return formattedFileName;
}

function parseExcelReport(localPath, rawFileName) {
  let transactions = [];
  let cardName = "";
  let finalBuffer = fs.readFileSync(localPath);

  try {
    const workbook = XLSX.read(finalBuffer, { type: 'buffer', cellDates: true });
    const sheetName = workbook.SheetNames[0];
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1 });
    const today = new Date();

    for (let i = 0; i < Math.min(10, rows.length); i++) {
      const r = rows[i];
      if (r && String(r[0]).includes("Account Type")) cardName = String(r[1]).trim();
    }

    // Find header row dynamically (e.g. "Transaction Date", "Posting Date", "Description")
    let headerRowIndex = -1;
    for (let i = 0; i < Math.min(20, rows.length); i++) {
      const r = rows[i];
      if (!r) continue;
      const rowStr = r.join(' ').toLowerCase();
      if (rowStr.includes('transaction date') || rowStr.includes('posting date') || (rowStr.includes('date') && rowStr.includes('description'))) {
        headerRowIndex = i;
        console.log(`[DEBUG 8/8] Table header detected at row ${i + 1}`);
        break;
      }
    }

    // Find 'Previous Balance' in Column C (index 2) or Column A (index 0)
    let prevBalanceRowIndex = -1;
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      if (!r) continue;
      const colC = String(r[2] || '').trim().toUpperCase();
      const colA = String(r[0] || '').trim().toUpperCase();
      if (colC.includes('PREVIOUS BALANCE') || colA.includes('PREVIOUS BALANCE')) {
        prevBalanceRowIndex = i;
        console.log(`[DEBUG 8/8] Found 'Previous Balance' at row ${i + 1}. Truncating all older statement rows.`);
        break;
      }
    }

    const rawRowsUpToPrevBalance = prevBalanceRowIndex !== -1 ? rows.slice(0, prevBalanceRowIndex) : rows;
    const startIndex = headerRowIndex !== -1 ? headerRowIndex + 1 : 8;

    // Keep rows up to Previous Balance and strip out prior bill payment rows (e.g. PAYMT)
    const rowsToProcess = rawRowsUpToPrevBalance.filter((r, idx) => {
      if (idx < startIndex) return true; // Keep headers & metadata
      if (!r || r.length < 3) return false;
      const colC = String(r[2] || '').toUpperCase();
      const colA = String(r[0] || '').toUpperCase();
      if (colC.includes('PAYMT') || colC.includes('PAYMENT') || colA.includes('PAYMT') || colC.includes('PREVIOUS BALANCE')) {
        console.log(`[DEBUG 8/8] Stripped payment row: "${r[2]}" (Amount: ${r[6]})`);
        return false;
      }
      return true;
    });

    for (let i = startIndex; i < rowsToProcess.length; i++) {
      const r = rowsToProcess[i];
      if (!r || r.length < 3) continue;

      let amt = parseFloat(String(r[6]).replace(/,/g, ''));
      if (isNaN(amt) || amt === 0) continue;

      let finalDesc = scrubDescription(r[2] || "Unknown");

      let badge = null;
      if (amt < 0) {
        if (finalDesc.toLowerCase().includes('cashback')) {
          badge = 'Cashback';
          finalDesc = 'Cashback';
        } else {
          badge = 'Refund';
        }
      }

      transactions.push({
        date: new Date(r[0] || today).toISOString(),
        description: finalDesc,
        currency: r[5] || "SGD",
        amount: amt,
        badge: badge
      });
    }

    // Format current month string, e.g. "Sep 2026"
    const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    const currentMonthYear = `${months[today.getMonth()]} ${today.getFullYear()}`;

    // Set Cell B6 (Row index 5, Col index 1) to current month
    if (rowsToProcess.length > 5) {
      if (!Array.isArray(rowsToProcess[5])) rowsToProcess[5] = [];
      rowsToProcess[5][1] = currentMonthYear;
      console.log(`[DEBUG 8/8] Set cell B6 (statement date) to "${currentMonthYear}"`);
    }

    // Also ensure any row labeled "Statement Date" has its value set to current month
    for (let i = 0; i < Math.min(10, rowsToProcess.length); i++) {
      const r = rowsToProcess[i];
      if (r && String(r[0]).toLowerCase().includes('statement date')) {
        r[1] = currentMonthYear;
      }
    }

    // Rebuild clean Excel file without Previous Balance or Payment rows
    const trimmedSheet = XLSX.utils.aoa_to_sheet(rowsToProcess);
    const newWb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(newWb, trimmedSheet, sheetName || 'Transactions');
    finalBuffer = XLSX.write(newWb, { type: 'buffer', bookType: 'biff8' });
    fs.writeFileSync(localPath, finalBuffer);
    console.log(`[DEBUG 8/8] Rewrote Excel file with ${rowsToProcess.length} clean running report rows.`);
  } catch (parseError) {
    console.error("[DEBUG ERROR] Failed parsing/truncating downloaded file:", parseError);
  }

  const base64File = finalBuffer.toString('base64');
  const fileName = rawFileName;

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

  const totalAmount = transactions.reduce((sum, t) => sum + (parseFloat(t.amount) || 0), 0);
  const totalExpenses = transactions.filter(t => t.amount > 0).reduce((sum, t) => sum + t.amount, 0);
  const totalRefunds = transactions.filter(t => t.amount < 0).reduce((sum, t) => sum + Math.abs(t.amount), 0);

  return {
    cardName: cardName || "UOB EVOL",
    count: transactions.length,
    totalAmount: parseFloat(totalAmount.toFixed(2)),
    total: parseFloat(totalAmount.toFixed(2)),
    totalAmountFormatted: totalAmount.toFixed(2),
    totalExpenses: parseFloat(totalExpenses.toFixed(2)),
    totalRefunds: parseFloat(totalRefunds.toFixed(2)),
    transactions: transactions,
    fileName: fileName,
    fileData: base64File,
    fileSize: finalBuffer.length,
    statementTimestamp: statementTimestamp
  };
}

async function downloadStatement(mode = 'BOTH') {
  let browser;
  const downloadedFiles = [];
  const reqMode = String(mode || 'BOTH').trim().toUpperCase();

  try {
    const isHeadless = process.env.HEADLESS !== 'false';
    console.log(`[DEBUG 1/7] Launching Chromium browser (headless: ${isHeadless}, mode: ${reqMode})...`);
    browser = await chromium.launch({
      headless: isHeadless,
      slowMo: isHeadless ? 0 : 50,
      ignoreDefaultArgs: ['--enable-automation'],
      args: [
        '--disable-blink-features=AutomationControlled',
        '--no-sandbox',
        '--disable-setuid-sandbox',
        '--disable-dev-shm-usage',
        '--disable-accelerated-2d-canvas',
        '--no-first-run',
        '--no-zygote'
      ]
    });

    // Custom User-Agent and headers prevent headless bot detection
    const context = await browser.newContext({
      userAgent: 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/133.0.0.0 Safari/537.36',
      viewport: { width: 1280, height: 800 },
      locale: 'en-SG',
      timezoneId: 'Asia/Singapore',
      acceptDownloads: true,
      extraHTTPHeaders: {
        'Accept-Language': 'en-SG,en-US;q=0.9,en;q=0.8',
        'sec-ch-ua': '"Not(A:Brand";v="99", "Google Chrome";v="133", "Chromium";v="133"',
        'sec-ch-ua-mobile': '?0',
        'sec-ch-ua-platform': '"macOS"'
      }
    });

    // Mask webdriver flag
    await context.addInitScript(() => {
      Object.defineProperty(navigator, 'webdriver', { get: () => undefined });
      window.chrome = { runtime: {} };
    });

    const page = await context.newPage();

    // 1. Navigate to UOB PIB
    console.log("[DEBUG 2/7] Navigating to UOB Personal Internet Banking login page...");
    await page.goto('https://pib.uob.com.sg/PIBLogin/Public/processPreCapture.do?keyId=lpc', { waitUntil: 'domcontentloaded', timeout: 45000 });
    console.log(`[DEBUG 2/7] Page landed. URL: ${page.url()} | Title: ${await page.title()}`);

    // 2. Fill Credentials
    console.log("[DEBUG 3/7] Filling username and password...");
    const usernameSelector = 'input[placeholder*="USERNAME" i], #userName, input[name="userName"], input[type="text"]';
    await page.waitForSelector(usernameSelector, { timeout: 20000 });
    await page.locator(usernameSelector).first().click();
    await page.locator(usernameSelector).first().pressSequentially(process.env.UOB_USERNAME, { delay: 15 });

    const passwordSelector = 'input[placeholder*="PASSWORD" i], #PASSWORD1, input[name="password"], input[type="password"]';
    await page.waitForSelector(passwordSelector, { timeout: 15000 });
    await page.locator(passwordSelector).first().click();
    await page.locator(passwordSelector).first().pressSequentially(process.env.UOB_PASSWORD, { delay: 15 });
    console.log("[DEBUG 3/7] Credentials filled successfully.");

    await page.waitForTimeout(200);

    // 3. Submit Form
    console.log("[DEBUG 4/7] Submitting login form...");
    const submitBtnSelector = 'button[type="submit"], input[type="submit"], #btnSubmit, button:has-text("LOG IN"), button:has-text("Submit"), button.btn-primary';
    const submitBtn = page.locator(submitBtnSelector).first();
    await submitBtn.click();

    console.log("[DEBUG 4/7] Submitted login. Waiting for post-submit modal or 2FA trigger...");

    // Check for stale active session modal and click Proceed
    console.log("[DEBUG 4.5] Checking for stale active session...");
    try {
      const proceedBtn = page.locator('button:has-text("Proceed"), button:has-text("proceed"), #btnProceed').first();
      await proceedBtn.waitFor({ state: 'visible', timeout: 8000 });
      console.log("[DEBUG 4.5] Active session popup detected. Clicking Proceed...");
      await proceedBtn.click();
    } catch (e) {
      console.log("[DEBUG 4.5] No active session modal found or Proceed button did not appear, continuing...");
    }

    console.log("--------------------------------------------------");
    console.log("📲 CHECK YOUR PHONE NOW FOR THE UOB 2FA PUSH NOTIFICATION");
    console.log("--------------------------------------------------");

    // 4. Wait for ACTUAL Authenticated Dashboard
    console.log("[DEBUG 6/7] Awaiting 2FA approval on phone (timeout: 90s)...");
    try {
      await Promise.race([
        page.waitForURL(/.*accountsDashboard.*/, { timeout: 90000 }),
        page.waitForSelector('text=UOB EVOL, text=Cards, #account-summary, a:has-text("UOB EVOL"), [role="tab"]:has-text("Apply & Services")', { timeout: 90000 })
      ]);
      console.log("[DEBUG 6/7] 2FA Approved! Authenticated dashboard loaded successfully!");
      await page.waitForTimeout(2000); // allow dashboard hydration
      await page.screenshot({ path: 'debug-dashboard.png', fullPage: true }).catch(() => {});
    } catch (e) {
      await page.screenshot({ path: 'debug-login-failed.png', fullPage: true }).catch(() => {});
      console.error("[DEBUG ERROR] 2FA approval timed out or dashboard failed to load. Saved screenshot to debug-login-failed.png. Final URL:", page.url());
      const bodyText = await page.locator('body').innerText().catch(() => '');
      console.error("[DEBUG ERROR] Page text preview:", bodyText.slice(0, 500));
      throw new Error(`2FA Timeout or Login Failed. Final URL: ${page.url()}`);
    }

    // ==========================================
    // 📄 MODE 'S': Statement PDF Download
    // ==========================================
    if (reqMode === 'S') {
      console.log("[DEBUG 7/7] Mode 'S': Navigating to eStatements...");
      await page.evaluate(() => window.scrollTo(0, 0)).catch(() => {});

      // Step 1: Click "Apply & Services" tab
      console.log("[DEBUG S-1] Clicking 'Apply & Services' tab...");
      const applyTab = page.locator('[role="tab"], button').filter({ hasText: 'Apply & Services' }).first();
      await applyTab.waitFor({ state: 'visible', timeout: 20000 });
      await applyTab.click({ force: true });
      await page.waitForURL(/.*servicesHome.*/, { timeout: 15000 });
      await page.waitForTimeout(1000);
      console.log("[DEBUG S-1] On servicesHome. URL:", page.url());

      // Step 2: Click "View eStatements" tile
      console.log("[DEBUG S-2] Clicking 'View eStatements' tile...");
      const viewEstatements = page.locator('div[description="View eStatements"], [description="View eStatements"]').first();
      await viewEstatements.waitFor({ state: 'visible', timeout: 20000 });
      await viewEstatements.click({ force: true });

      try {
        await page.waitForURL(/.*estatements.*/, { timeout: 8000 });
      } catch (err) {
        console.log("[DEBUG S-2] Fallback: clicking leaf text 'View eStatements'...");
        const textTarget = page.locator('text="View eStatements"').first();
        await textTarget.click({ force: true });
        await page.waitForURL(/.*estatements.*/, { timeout: 15000 });
      }
      await page.waitForTimeout(1000);
      console.log("[DEBUG S-2] On estatements. URL:", page.url());

      // Step 3: Click "PERSONAL CREDIT CARD STATEMENT"
      console.log("[DEBUG S-3] Clicking 'PERSONAL CREDIT CARD STATEMENT'...");
      await page.evaluate(() => window.scrollTo(0, document.body.scrollHeight)).catch(() => {});
      const cardStatementBtn = page.locator('button:has-text("PERSONAL CREDIT CARD STATEMENT")').first();
      await cardStatementBtn.waitFor({ state: 'visible', timeout: 30000 });
      await cardStatementBtn.scrollIntoViewIfNeeded().catch(() => {});
      await cardStatementBtn.click({ force: true });
      await page.waitForURL(/.*estatements\/view.*/, { timeout: 15000 });
      await page.waitForTimeout(3000); // Allow statement table and download button to fully populate and enable
      console.log("[DEBUG S-3] On estatements/view. URL:", page.url());

      // Step 4: Click "Download" button
      console.log("[DEBUG S-4] Clicking 'Download' button...");
      await page.evaluate(() => window.scrollTo(0, 0)).catch(() => {});
      
      // Wait for Download button to be enabled (not disabled)
      const statementDownloadBtn = page.locator('button:has-text("Download"):not([disabled])').first();
      await statementDownloadBtn.waitFor({ state: 'visible', timeout: 30000 });
      await page.waitForTimeout(1000);

      console.log("[DEBUG S-4] Download button enabled and ready. Triggering download...");
      const downloadPromise = page.waitForEvent('download', { timeout: 60000 });
      await statementDownloadBtn.click();

      // Ensure download initiated; if not within 4s, re-click as fallback
      const statementDownload = await Promise.race([
        downloadPromise,
        page.waitForTimeout(4000).then(async () => {
          console.log("[DEBUG S-4] Download event pending, re-clicking download button...");
          await statementDownloadBtn.click().catch(() => {});
          return downloadPromise;
        })
      ]);

      const rawStatementFileName = statementDownload.suggestedFilename();
      const localStatementPath = `./${rawStatementFileName}`;
      downloadedFiles.push(localStatementPath);
      await statementDownload.saveAs(localStatementPath);
      console.log(`[DEBUG 7/7] [STATEMENT] PDF downloaded successfully: ${localStatementPath}`);

      const formattedPdfName = formatStatementPdfFileName(rawStatementFileName);
      console.log(`[DEBUG 7/7] [STATEMENT] Output filename formatted: "${formattedPdfName}"`);

      const pdfBuffer = fs.readFileSync(localStatementPath);
      const pdfBase64 = pdfBuffer.toString('base64');

      const now = new Date();
      const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
      const pad = (num) => String(num).padStart(2, '0');
      const statementTimestamp = `${pad(now.getHours())}:${pad(now.getMinutes())}:${pad(now.getSeconds())} ${months[now.getMonth()]} ${pad(now.getDate())}, ${now.getFullYear()}`;

      return {
        success: true,
        mode: 'S',
        type: 'S',
        fileName: formattedPdfName,
        fileData: pdfBase64,
        fileSize: pdfBuffer.length,
        mimeType: 'application/pdf',
        statementFileName: formattedPdfName,
        statementFileData: pdfBase64,
        statementTimestamp: statementTimestamp
      };
    }

    // ==========================================
    // 📊 MODE 'R': Running Report Excel Download
    // ==========================================
    console.log("[DEBUG 7/7] Mode 'R': Locating and clicking UOB EVOL card tile...");
    const evolCard = page.locator('text="UOB EVOL"').first();
    await evolCard.waitFor({ state: 'visible', timeout: 20000 });
    await evolCard.click();
    console.log(`[DEBUG 7/7] Clicked UOB EVOL card tile. Waiting for download button...`);
    await page.waitForTimeout(2000);

    const downloadBtn = page.locator('button:has-text("Download as Excel"), button:has-text("Download"), [aria-label*="Download" i], button#btnsubmit.btn-default.btn-icon').first();
    console.log("[DEBUG 7/7] Locating download button...");
    await downloadBtn.waitFor({ state: 'visible', timeout: 15000 });

    console.log("[DEBUG 7/7] Triggering download event...");
    const [download] = await Promise.all([
      page.waitForEvent('download', { timeout: 60000 }),
      downloadBtn.click()
    ]);

    const reportFileName = download.suggestedFilename();
    const localReportPath = `./${reportFileName}`;
    downloadedFiles.push(localReportPath);
    await download.saveAs(localReportPath);
    console.log(`[DEBUG 7/7] File downloaded successfully: ${localReportPath}`);

    const reportResult = parseExcelReport(localReportPath, reportFileName);

    return {
      success: true,
      mode: 'R',
      type: 'R',
      cardName: reportResult.cardName || "UOB EVOL",
      card: reportResult.cardName || "UOB EVOL",
      count: reportResult.count,
      totalAmount: reportResult.totalAmount,
      total: reportResult.total,
      totalAmountFormatted: reportResult.totalAmountFormatted,
      totalExpenses: reportResult.totalExpenses,
      totalRefunds: reportResult.totalRefunds,
      transactions: reportResult.transactions,
      fileName: reportResult.fileName,
      fileData: reportResult.fileData,
      fileSize: reportResult.fileSize,
      reportFileName: reportResult.fileName,
      reportFileData: reportResult.fileData,
      statementTimestamp: reportResult.statementTimestamp
    };
  } finally {
    if (browser) {
      await browser.close();
    }
    for (const f of downloadedFiles) {
      if (f && fs.existsSync(f)) {
        try {
          fs.unlinkSync(f);
        } catch (err) {
          console.error("[DEBUG ERROR] Failed to clean up file:", f, err);
        }
      }
    }
  }
}
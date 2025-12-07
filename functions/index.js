// === [1] Setup and Authentication ===
const express = require('express');
const cors = require('cors');
const fs = require('fs');
const { google } = require('googleapis');
const path = require('path');
const ExcelJS = require('exceljs'); // NEW: for Excel export
require('dotenv').config();

console.log(
  'Loaded GOOGLE_CREDENTIALS_SHEET_ID:',
  process.env.GOOGLE_CREDENTIALS_SHEET_ID
);

const app = express();
app.use(express.json());
app.use(cors());
app.use(express.static(path.join(__dirname, '..', 'public')));

const SERVICE_ACCOUNT_JSON =
  process.env.GOOGLE_SERVICE_ACCOUNT_JSON || 'service-account.json';

// === [1b] Logger ===
function logDebug(msg, data = null) {
  const entry =
    `[${new Date().toISOString()}] ${msg}` +
    (data !== null ? `: ${JSON.stringify(data)}` : '');
  console.log(entry);
  // fs.appendFileSync('server-debug.log', entry + '\n');
}

// === [2] Google Sheets Helper ===
function getSheetsClient() {
  const creds = JSON.parse(fs.readFileSync(SERVICE_ACCOUNT_JSON, 'utf8'));
  const auth = new google.auth.GoogleAuth({
    credentials: creds,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
  return google.sheets({ version: 'v4', auth });
}

// Helpers
function serialToDate(serial) {
  return new Date(Date.UTC(1899, 11, 30) + serial * 86400000);
}
function isSameMonth(date, now = new Date()) {
  return (
    date &&
    date.getFullYear() === now.getFullYear() &&
    date.getMonth() === now.getMonth()
  );
}

// (kept for compatibility if needed later)
function getYearMonthCairo(date = new Date()) {
  const y = new Intl.DateTimeFormat('sv-SE', {
    timeZone: 'Africa/Cairo',
    year: 'numeric',
  }).format(date);
  const m = new Intl.DateTimeFormat('sv-SE', {
    timeZone: 'Africa/Cairo',
    month: '2-digit',
  }).format(date);
  return `${y}-${m}`;
}
function getYearMonthFromSvSe(ts) {
  return (ts || '').toString().slice(0, 7);
}

// Optional legacy helper – not used now but kept if needed
function getBudgetSheetId(userType) {
  if (userType === 'cash') return process.env.GOOGLE_CASH_BUDGET_SHEET_ID;
  return process.env.GOOGLE_TASA_BUDGET_SHEET_ID;
}

/* ========================================================================== */
/* === [2b] Generic helper: ALWAYS append from A to K                      === */
/* ========================================================================== */
/**
 * Appends rows (each up to 11 columns) into sheetName,
 * ALWAYS starting in column A and ending at column K.
 */
async function appendRowsAtoK(sheets, spreadsheetId, sheetName, values) {
  if (!Array.isArray(values) || values.length === 0) return;

  // Read column A only to find last non-empty row
  const colAResp = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: `${sheetName}!A:A`,
  });
  const colArows = colAResp.data.values || [];

  let lastRow = 0;
  for (let i = colArows.length - 1; i >= 0; i--) {
    const v = colArows[i] && colArows[i][0];
    if (v !== undefined && v !== null && v.toString().trim() !== '') {
      lastRow = i + 1; // 1-based row index
      break;
    }
  }

  const startRow = lastRow + 1;
  const endRow = startRow + values.length - 1;
  const range = `${sheetName}!A${startRow}:K${endRow}`;

  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range,
    valueInputOption: 'USER_ENTERED',
    resource: { values },
  });

  logDebug('appendRowsAtoK wrote rows', {
    sheetName,
    startRow,
    endRow,
    rows: values.length,
  });
}

/* ========================================================================== */
/* === [2c] Master-per-client helpers                                      === */
/* ========================================================================== */

/**
 * Login helper:
 * Master Credentials sheet (GOOGLE_CREDENTIALS_SHEET_ID) contains one tab per client.
 * Columns (starting row 2): A=username, B=password, C=branch, D=restricted(Y/N),
 * E=Level(L1/L2), Z=PaperMode(Y/N). BudgetSheetId is in F2 on the same tab.
 */
async function findClientTabAndSheetIdByUser(
  sheets,
  masterSheetId,
  username,
  password
) {
  const metaResp = await sheets.spreadsheets.get({
    spreadsheetId: masterSheetId,
  });
  const allSheets = metaResp.data.sheets || [];
  const sheetTitles = allSheets
    .map((s) => s.properties?.title)
    .filter(Boolean);

  const exclude = new Set(['Config', 'Readme']);
  const clientTabs = sheetTitles.filter((t) => !exclude.has(t));

  for (const title of clientTabs) {
    const credResp = await sheets.spreadsheets.values.get({
      spreadsheetId: masterSheetId,
      range: `${title}!A2:Z`,
    });
    const rows = credResp.data.values || [];

    for (const row of rows) {
      const u = (row[0] || '').toString().trim();
      const p = (row[1] || '').toString();
      if (!u && !p) continue;
      if (u === username && p === password) {
        const branch = (row[2] || '').toString().trim();
        const restricted =
          (row[3] || '').toString().trim().toUpperCase() === 'Y';
        const level =
          (row[4] || '').toString().trim().toUpperCase() || 'L1';
        const paperMode =
          (row[25] || '').toString().trim().toUpperCase() === 'Y';

        const idCellResp = await sheets.spreadsheets.values.get({
          spreadsheetId: masterSheetId,
          range: `${title}!F2`,
        });
        const budgetSheetId =
          (idCellResp.data.values?.[0]?.[0] || '').toString().trim();
        if (!budgetSheetId)
          throw new Error(`BudgetSheetId missing in ${title}!F2`);

        return {
          tab: title,
          branch,
          restricted,
          level,
          paperMode,
          budgetSheetId,
        };
      }
    }
  }
  return null;
}

/**
 * Generic helper: get user info (tab, branch, restricted, level, paperMode, budgetSheetId)
 * by username only. Used after login in all endpoints that need to know L1/L2 and sheet id.
 */
async function getUserInfoByUsername(sheets, masterSheetId, username) {
  const metaResp = await sheets.spreadsheets.get({
    spreadsheetId: masterSheetId,
  });
  const allSheets = metaResp.data.sheets || [];
  const sheetTitles = allSheets
    .map((s) => s.properties?.title)
    .filter(Boolean);

  const exclude = new Set(['Config', 'Readme']);
  const clientTabs = sheetTitles.filter((t) => !exclude.has(t));

  for (const title of clientTabs) {
    const credResp = await sheets.spreadsheets.values.get({
      spreadsheetId: masterSheetId,
      range: `${title}!A2:Z`,
    });
    const rows = credResp.data.values || [];

    for (const row of rows) {
      const u = (row[0] || '').toString().trim();
      if (!u) continue;
      if (u === username) {
        const branch = (row[2] || '').toString().trim();
        const restricted =
          (row[3] || '').toString().trim().toUpperCase() === 'Y';
        const level =
          (row[4] || '').toString().trim().toUpperCase() || 'L1';
        const paperMode =
          (row[25] || '').toString().trim().toUpperCase() === 'Y';

        const idCellResp = await sheets.spreadsheets.values.get({
          spreadsheetId: masterSheetId,
          range: `${title}!F2`,
        });
        const budgetSheetId =
          (idCellResp.data.values?.[0]?.[0] || '').toString().trim();
        if (!budgetSheetId)
          throw new Error(`BudgetSheetId missing in ${title}!F2`);

        return {
          tab: title,
          branch,
          restricted,
          level,
          paperMode,
          budgetSheetId,
        };
      }
    }
  }
  return null;
}

/**
 * Legacy: resolve only budgetSheetId by username (used in some endpoints).
 */
async function resolveBudgetSheetIdFromF2ByUsername(
  sheets,
  masterSheetId,
  username
) {
  const metaResp = await sheets.spreadsheets.get({
    spreadsheetId: masterSheetId,
  });
  const allSheets = metaResp.data.sheets || [];
  const sheetTitles = allSheets
    .map((s) => s.properties?.title)
    .filter(Boolean);

  const exclude = new Set(['Config', 'Readme']);
  const clientTabs = sheetTitles.filter((t) => !exclude.has(t));

  for (const title of clientTabs) {
    const credResp = await sheets.spreadsheets.values.get({
      spreadsheetId: masterSheetId,
      range: `${title}!A2:A`,
    });
    const usernames = credResp.data.values || [];
    for (const row of usernames) {
      const u = (row[0] || '').toString().trim();
      if (!u) continue;
      if (u === username) {
        const idCellResp = await sheets.spreadsheets.values.get({
          spreadsheetId: masterSheetId,
          range: `${title}!F2`,
        });
        const budgetSheetId =
          (idCellResp.data.values?.[0]?.[0] || '').toString().trim();
        if (!budgetSheetId)
          throw new Error(`BudgetSheetId missing in ${title}!F2`);
        return budgetSheetId;
      }
    }
  }
  throw new Error('User not found in any client tab');
}

/* ========================================================================== */
/* === [2d] Helper: merge items into orders sheet (Final or Waiting)      === */
/* ========================================================================== */
/**
 * LEGACY: kept for compatibility in case needed later.
 * Currently NOT used for new orders / approvals, because we no longer consolidate orders.
 */
async function mergeItemsIntoOrdersSheet(
  sheets,
  spreadsheetId,
  sheetName,
  branchName,
  items,
  defaultUsername
) {
  if (!Array.isArray(items) || items.length === 0) return;

  const ordersResp = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: sheetName,
  });
  const ordersRows = ordersResp.data.values || [];
  const now = new Date();

  const existingIndex = {};
  for (let i = 1; i < ordersRows.length; i++) {
    const row = ordersRows[i];
    const cell = row[0];
    const date = !cell
      ? null
      : !isNaN(cell) && Number(cell) > 30000
      ? serialToDate(Number(cell))
      : new Date(cell);
    if (!date || !isSameMonth(date, now)) continue;
    const rowBranch = row[1];
    if (rowBranch !== branchName) continue;
    const code = row[3];
    const unitPrice = parseFloat(row[5]) || 0;
    const qty = parseInt(row[8]) || 0;
    existingIndex[code] = { rowIndex: i, qty, unitPrice };
  }

  const batchUpdates = [];
  const rowsToAppend = [];

  items.forEach((item) => {
    const code = item.productCode;
    const existing = existingIndex[code];
    const qtyToAdd = item.quantity || 0;

    if (existing) {
      const newQty = (existing.qty || 0) + qtyToAdd;
      const price = existing.unitPrice || item.unitPrice || 0;
      const newSubtotal = price * newQty;
      const r = existing.rowIndex + 1;

      batchUpdates.push(
        { range: `${sheetName}!I${r}`, values: [[newQty]] },
        { range: `${sheetName}!G${r}`, values: [[newSubtotal]] }
      );
    } else {
      rowsToAppend.push(item);
    }
  });

  if (batchUpdates.length > 0) {
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId,
      resource: { valueInputOption: 'USER_ENTERED', data: batchUpdates },
    });
    logDebug(`Updated existing lines in ${sheetName}`, {
      count: batchUpdates.length / 2,
    });
  }

  if (rowsToAppend.length > 0) {
    const now2 = new Date();
    const dateString = now2
      .toLocaleString('sv-SE', { timeZone: 'Africa/Cairo' })
      .replace('T', ' ');

    const appendValues = rowsToAppend.map((item) => {
      const price = Number(item.unitPrice) || 0;
      const qty = Number(item.quantity) || 0;
      const subtotal =
        (typeof item.subtotal === 'number'
          ? item.subtotal
          : price * qty) || 0;
      const cat = item.category || '';
      const user = item.username || defaultUsername || '';
      return [
        dateString,
        branchName,
        user,
        item.productCode,
        item.productName,
        price,
        subtotal,
        cat,
        qty,
        '',
        '',
      ];
    });

    // use new helper so it always starts in A
    await appendRowsAtoK(sheets, spreadsheetId, sheetName, appendValues);
    logDebug(`Appended new rows to ${sheetName}`, {
      count: rowsToAppend.length,
    });
  }
}

/* ========================================================================== */
/* === [2e] NEW Helper: Order Serial Number                               === */
/* ========================================================================== */
/**
 * Reads the last serial from Serial Numbers!B2, increments it, writes back,
 * and returns new serial. Format: "AA<number>" (AA1, AA2, ...).
 */
async function getNextOrderSerial(sheets, spreadsheetId) {
  const prefix = 'AA';
  let currentSerial = '';
  try {
    const resp = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: 'Serial Numbers!B2',
    });
    currentSerial =
      (resp.data.values && resp.data.values[0] && resp.data.values[0][0]) ||
      '';
    currentSerial = currentSerial.toString().trim();
  } catch (err) {
    logDebug('Error reading Serial Numbers!B2, assuming first serial', {
      error: err.message,
    });
    currentSerial = '';
  }

  const upper = currentSerial.toUpperCase();
  let currentNum = 0;
  if (upper && upper.startsWith(prefix)) {
    const numPartStr = upper.slice(prefix.length).trim();
    const parsed = parseInt(numPartStr, 10);
    if (!isNaN(parsed) && parsed >= 0) {
      currentNum = parsed;
    }
  }

  const newNum = currentNum + 1;
  const newSerial = `${prefix}${newNum}`;

  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range: 'Serial Numbers!B2',
    valueInputOption: 'USER_ENTERED',
    resource: { values: [[newSerial]] },
  });

  logDebug('Generated new order serial', {
    previous: currentSerial || null,
    newSerial,
  });

  return newSerial;
}

/* ========================================================================== */
/* === [3] Login Endpoint (restricted + paperMode + level L1/L2)          === */
/* ========================================================================== */
app.post('/api/validateLogin', async (req, res) => {
  try {
    const { username, password } = req.body;
    logDebug('Login attempt', { username });

    const sheets = getSheetsClient();
    const credentialsSheetId = process.env.GOOGLE_CREDENTIALS_SHEET_ID;

    const match = await findClientTabAndSheetIdByUser(
      sheets,
      credentialsSheetId,
      username,
      password
    );
    if (!match) {
      logDebug('Login failed', { username });
      return res.json({
        success: false,
        message: 'اسم المستخدم أو كلمة المرور غير صحيحة',
      });
    }

    const user = {
      username,
      branch: match.branch || '',
      userType: 'tasa', // kept for compatibility
      restricted: !!match.restricted,
      paperMode: !!match.paperMode,
      level: match.level || 'L1',
    };

    logDebug('Login success (client tab via F2)', {
      username,
      tab: match.tab,
      paperMode: user.paperMode,
      level: match.level,
    });
    return res.json({ success: true, user });
  } catch (err) {
    logDebug('Login error', { error: err.message });
    res
      .status(500)
      .json({ success: false, message: 'حدث خطأ في النظام' });
  }
});

/* ========================================================================== */
/* === [3b] For L2: Get list of branches for client                       === */
/* ========================================================================== */
// Original endpoint (not used by current frontend but kept)
app.get('/api/clientBranches', async (req, res) => {
  try {
    const username = (req.query.username || '').trim();
    if (!username) {
      return res
        .status(400)
        .json({ success: false, message: 'Missing username' });
    }

    const sheets = getSheetsClient();
    const credentialsSheetId = process.env.GOOGLE_CREDENTIALS_SHEET_ID;

    const userInfo = await getUserInfoByUsername(
      sheets,
      credentialsSheetId,
      username
    );
    if (!userInfo) {
      return res
        .status(400)
        .json({ success: false, message: 'المستخدم غير موجود' });
    }

    const clientTab = userInfo.tab;
    const resp = await sheets.spreadsheets.values.get({
      spreadsheetId: credentialsSheetId,
      range: `${clientTab}!C2:C`,
    });

    const rows = resp.data.values || [];
    const set = new Set();
    rows.forEach((r) => {
      const b = (r[0] || '').toString().trim();
      if (b) set.add(b);
    });

    res.json({ success: true, branches: Array.from(set) });
  } catch (err) {
    logDebug('Error loading client branches', { error: err.message });
    res
      .status(500)
      .json({ success: false, message: 'حدث خطأ في تحميل الفروع' });
  }
});

// NEW: Endpoint used by frontend: /api/branchesForL2
app.get('/api/branchesForL2', async (req, res) => {
  try {
    const username = (req.query.username || '').trim();
    if (!username) {
      return res.status(400).json({
        success: false,
        message: 'Missing username',
        branches: [],
      });
    }

    const sheets = getSheetsClient();
    const credentialsSheetId = process.env.GOOGLE_CREDENTIALS_SHEET_ID;

    const userInfo = await getUserInfoByUsername(
      sheets,
      credentialsSheetId,
      username
    );
    if (!userInfo) {
      return res.status(400).json({
        success: false,
        message: 'المستخدم غير موجود',
        branches: [],
      });
    }

    const clientTab = userInfo.tab;
    const resp = await sheets.spreadsheets.values.get({
      spreadsheetId: credentialsSheetId,
      range: `${clientTab}!C2:C`,
    });

    const rows = resp.data.values || [];
    const set = new Set();
    rows.forEach((r) => {
      const b = (r[0] || '').toString().trim();
      if (b) set.add(b);
    });

    const branches = Array.from(set);
    res.json({ success: true, branches });
  } catch (err) {
    logDebug('Error in branchesForL2', { error: err.message });
    res.status(500).json({
      success: false,
      message: 'حدث خطأ في تحميل الفروع',
      branches: [],
    });
  }
});

/* ========================================================================== */
/* === [4] Load Order Data Endpoint                                       === */
/* ========================================================================== */
// CLEANED: now only loads products (no budgets, no Monthly Paper Count, no limits).
app.get('/api/loadOrderDataWithSpending', async (req, res) => {
  try {
    const branchName = req.query.branchName;
    const userType = req.query.userType || 'tasa';
    const username = req.query.username || '';
    const sheets = getSheetsClient();

    const credentialsSheetId = process.env.GOOGLE_CREDENTIALS_SHEET_ID;
    const BUDGET_SHEET_ID =
      await resolveBudgetSheetIdFromF2ByUsername(
        sheets,
        credentialsSheetId,
        username
      );

    logDebug('Loading order data (products only)', {
      branchName,
      userType,
      BUDGET_SHEET_ID,
    });

    const productsResp = await sheets.spreadsheets.values.get({
      spreadsheetId: BUDGET_SHEET_ID,
      range: 'Product Catalog',
    });
    const productsRows = productsResp.data.values || [];
    const products = [];
    for (let i = 1; i < productsRows.length; i++) {
      products.push({
        code: productsRows[i][0],
        name: productsRows[i][1],
        category: productsRows[i][2],
        price: parseFloat(productsRows[i][3]) || 0,
        imageUrl: productsRows[i][4] || '',
      });
    }
    logDebug('Loaded products', { count: products.length });

    res.json({ products });
  } catch (err) {
    logDebug('Error loading data', { error: err.message });
    res.status(500).json({
      success: false,
      message: 'حدث خطأ في تحميل البيانات',
    });
  }
});

/* ========================================================================== */
/* === [4b] GET Past Orders For Branch, Current Month                      === */
/* ========================================================================== */
app.get('/api/previousOrders', async (req, res) => {
  try {
    const branchName = req.query.branchName;
    const userType = req.query.userType || 'tasa';
    const username = req.query.username || '';

    const sheets = getSheetsClient();
    const credentialsSheetId = process.env.GOOGLE_CREDENTIALS_SHEET_ID;
    const BUDGET_SHEET_ID =
      await resolveBudgetSheetIdFromF2ByUsername(
        sheets,
        credentialsSheetId,
        username
      );

    logDebug('Extracting previous orders', {
      branchName,
      userType,
      BUDGET_SHEET_ID,
    });

    const productsResp = await sheets.spreadsheets.values.get({
      spreadsheetId: BUDGET_SHEET_ID,
      range: 'Product Catalog',
    });
    const productRows = productsResp.data.values || [];
    const codeToProduct = {};
    for (let i = 1; i < productRows.length; i++) {
      const row = productRows[i];
      codeToProduct[row[0]] = {
        name: row[1],
        imageUrl: row[4] || '',
      };
    }

    const ordersResp = await sheets.spreadsheets.values.get({
      spreadsheetId: BUDGET_SHEET_ID,
      range: 'Final Orders',
    });
    const ordersRows = ordersResp.data.values || [];
    const now = new Date();

    const ordersMap = {};
    for (let i = 1; i < ordersRows.length; i++) {
      const row = ordersRows[i];
      const cell = row[0];
      const date = !cell
        ? null
        : !isNaN(cell) && Number(cell) > 30000
        ? serialToDate(Number(cell))
        : new Date(cell);
      if (!date || !isSameMonth(date, now)) continue;
      const orderBranch = row[1];
      if (orderBranch !== branchName) continue;
      const productCode = row[3];
      const quantity = parseInt(row[8]) || 0;

      ordersMap[productCode] = {
        productCode,
        productName: codeToProduct[productCode]?.name || productCode,
        imageUrl: codeToProduct[productCode]?.imageUrl || '',
        quantity,
        rowIndex: i,
      };
    }

    const ordersList = Object.values(ordersMap);
    logDebug('Extracted previous orders', {
      count: ordersList.length,
    });
    res.json({ orders: ordersList });
  } catch (err) {
    logDebug('Error extracting previous orders', {
      error: err.message,
    });
    res.status(500).json({
      success: false,
      message: 'حدث خطأ أثناء استخراج الطلبيات السابقة',
    });
  }
});

/* ========================================================================== */
/* === [4c] L2 Approvals (original branch-based endpoints)                 === */
/* ========================================================================== */

// Summary of "Waiting for Approval" by branch (current month).
app.get('/api/approvalsSummary', async (req, res) => {
  try {
    const username = (req.query.username || '').trim();
    if (!username) {
      return res
        .status(400)
        .json({ success: false, message: 'Missing username' });
    }

    const sheets = getSheetsClient();
    const credentialsSheetId = process.env.GOOGLE_CREDENTIALS_SHEET_ID;
    const userInfo = await getUserInfoByUsername(
      sheets,
      credentialsSheetId,
      username
    );

    if (!userInfo) {
      return res
        .status(400)
        .json({ success: false, message: 'المستخدم غير موجود' });
    }
    if ((userInfo.level || '').toUpperCase() !== 'L2') {
      return res.status(403).json({
        success: false,
        message: 'غير مسموح بالموافقة على الطلبات',
      });
    }

    const BUDGET_SHEET_ID = userInfo.budgetSheetId;

    const waitingResp = await sheets.spreadsheets.values.get({
      spreadsheetId: BUDGET_SHEET_ID,
      range: 'Waiting for Approval',
    });
    const rows = waitingResp.data.values || [];
    const now = new Date();
    const summary = {};

    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      const cell = row[0];
      const date = !cell
        ? null
        : !isNaN(cell) && Number(cell) > 30000
        ? serialToDate(Number(cell))
        : new Date(cell);
      if (!date || !isSameMonth(date, now)) continue;

      const branch = (row[1] || '').toString().trim();
      if (!branch) continue;
      const unitPrice = parseFloat(row[5]) || 0;
      const qty = parseInt(row[8]) || 0;
      const subtotal = parseFloat(row[6]) || unitPrice * qty;

      if (!summary[branch]) {
        summary[branch] = {
          branchName: branch,
          totalAmount: 0,
          totalQty: 0,
          lines: 0,
        };
      }
      summary[branch].totalAmount += subtotal;
      summary[branch].totalQty += qty;
      summary[branch].lines += 1;
    }

    res.json({ success: true, branches: Object.values(summary) });
  } catch (err) {
    logDebug('Error in approvalsSummary', {
      error: err.message,
    });
    res.status(500).json({
      success: false,
      message: 'حدث خطأ أثناء تحميل طلبات الموافقة',
    });
  }
});

// Details of "Waiting for Approval" for a specific branch (current month).
app.get('/api/approvalDetails', async (req, res) => {
  try {
    const username = (req.query.username || '').trim();
    const branchName = (req.query.branchName || '').trim();

    if (!username || !branchName) {
      return res
        .status(400)
        .json({ success: false, message: 'البيانات غير مكتملة' });
    }

    const sheets = getSheetsClient();
    const credentialsSheetId = process.env.GOOGLE_CREDENTIALS_SHEET_ID;
    const userInfo = await getUserInfoByUsername(
      sheets,
      credentialsSheetId,
      username
    );

    if (!userInfo) {
      return res
        .status(400)
        .json({ success: false, message: 'المستخدم غير موجود' });
    }
    if ((userInfo.level || '').toUpperCase() !== 'L2') {
      return res.status(403).json({
        success: false,
        message: 'غير مسموح بالموافقة على الطلبات',
      });
    }

    const BUDGET_SHEET_ID = userInfo.budgetSheetId;

    const productsResp = await sheets.spreadsheets.values.get({
      spreadsheetId: BUDGET_SHEET_ID,
      range: 'Product Catalog',
    });
    const productRows = productsResp.data.values || [];
    const codeToImage = {};
    for (let i = 1; i < productRows.length; i++) {
      const row = productRows[i];
      codeToImage[row[0]] = row[4] || '';
    }

    const waitingResp = await sheets.spreadsheets.values.get({
      spreadsheetId: BUDGET_SHEET_ID,
      range: 'Waiting for Approval',
    });
    const rows = waitingResp.data.values || [];
    const now = new Date();
    const items = [];

    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      const cell = row[0];
      const date = !cell
        ? null
        : !isNaN(cell) && Number(cell) > 30000
        ? serialToDate(Number(cell))
        : new Date(cell);
      if (!date || !isSameMonth(date, now)) continue;

      const rowBranch = (row[1] || '').toString().trim();
      if (rowBranch !== branchName) continue;

      const productCode = row[3];
      const productName = row[4];
      const unitPrice = parseFloat(row[5]) || 0;
      const qty = parseInt(row[8]) || 0;
      const subtotal = parseFloat(row[6]) || unitPrice * qty;
      const category = row[7] || '';

      items.push({
        productCode,
        productName,
        unitPrice,
        quantity: qty,
        subtotal,
        category,
        imageUrl: codeToImage[productCode] || '',
      });
    }

    res.json({ success: true, branchName, items });
  } catch (err) {
    logDebug('Error in approvalDetails', {
      error: err.message,
    });
    res.status(500).json({
      success: false,
      message: 'حدث خطأ أثناء تحميل تفاصيل الطلب',
    });
  }
});

/* ========================================================================== */
/* === [4d] Approve branch order (L2 moves Waiting -> Final Orders)       === */
/* ========================================================================== */
app.post('/api/approveBranchOrder', async (req, res) => {
  try {
    const { username, branchName } = req.body;

    if (!username || !branchName) {
      return res
        .status(400)
        .json({ success: false, message: 'البيانات غير مكتملة' });
    }

    const sheets = getSheetsClient();
    const credentialsSheetId = process.env.GOOGLE_CREDENTIALS_SHEET_ID;
    const userInfo = await getUserInfoByUsername(
      sheets,
      credentialsSheetId,
      username
    );

    if (!userInfo) {
      return res
        .status(400)
        .json({ success: false, message: 'المستخدم غير موجود' });
    }
    if ((userInfo.level || '').toUpperCase() !== 'L2') {
      return res.status(403).json({
        success: false,
        message: 'غير مسموح بالموافقة على الطلبات',
      });
    }

    const BUDGET_SHEET_ID = userInfo.budgetSheetId;

    const waitingResp = await sheets.spreadsheets.values.get({
      spreadsheetId: BUDGET_SHEET_ID,
      range: 'Waiting for Approval',
    });
    const rows = waitingResp.data.values || [];
    const now = new Date();

    const rowsToClear = [];
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      const cell = row[0];
      const date = !cell
        ? null
        : !isNaN(cell) && Number(cell) > 30000
        ? serialToDate(Number(cell))
        : new Date(cell);
      if (!date || !isSameMonth(date, now)) continue;

      const rowBranch = (row[1] || '').toString().trim();
      if (rowBranch !== branchName) continue;

      rowsToClear.push(i);
    }

    if (rowsToClear.length === 0) {
      return res.status(400).json({
        success: false,
        message: 'لا يوجد طلبات معلقة لهذا الفرع في هذا الشهر',
      });
    }

    const appendValues = rowsToClear.map((idx) => {
      const row = rows[idx] || [];
      const get = (r, col) =>
        r[col] !== undefined && r[col] !== null ? r[col] : '';
      return [
        get(row, 0),
        get(row, 1),
        get(row, 2),
        get(row, 3),
        get(row, 4),
        get(row, 5),
        get(row, 6),
        get(row, 7),
        get(row, 8),
        get(row, 9),
        get(row, 10),
      ];
    });

    // write to Final Orders from column A
    await appendRowsAtoK(sheets, BUDGET_SHEET_ID, 'Final Orders', appendValues);

    const clearData = rowsToClear.map((idx) => ({
      range: `Waiting for Approval!A${idx + 1}:K${idx + 1}`,
      values: [['', '', '', '', '', '', '', '', '', '', '']],
    }));
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: BUDGET_SHEET_ID,
      resource: {
        valueInputOption: 'USER_ENTERED',
        data: clearData,
      },
    });

    logDebug('Approved branch order', {
      branchName,
      items: rowsToClear.length,
    });
    res.json({ success: true });
  } catch (err) {
    logDebug('Error in approveBranchOrder', {
      error: err.message,
      stack: err.stack,
    });
    res.status(500).json({
      success: false,
      message: 'حدث خطأ أثناء تأكيد الطلب',
    });
  }
});

/* ========================================================================== */
/* === [4e] NEW L2 endpoints used by current frontend (per-branch order)  === */
/* ========================================================================== */

// GET /api/pendingOrders?username=...
app.get('/api/pendingOrders', async (req, res) => {
  try {
    const username = (req.query.username || '').trim();
    if (!username) {
      return res
        .status(400)
        .json({ success: false, message: 'Missing username' });
    }

    const sheets = getSheetsClient();
    const credentialsSheetId = process.env.GOOGLE_CREDENTIALS_SHEET_ID;
    const userInfo = await getUserInfoByUsername(
      sheets,
      credentialsSheetId,
      username
    );

    if (!userInfo) {
      return res
        .status(400)
        .json({ success: false, message: 'المستخدم غير موجود' });
    }
    if ((userInfo.level || '').toUpperCase() !== 'L2') {
      return res.status(403).json({
        success: false,
        message: 'غير مسموح بالموافقة على الطلبات',
      });
    }

    const BUDGET_SHEET_ID = userInfo.budgetSheetId;

    const waitingResp = await sheets.spreadsheets.values.get({
      spreadsheetId: BUDGET_SHEET_ID,
      range: 'Waiting for Approval',
    });
    const rows = waitingResp.data.values || [];
    const now = new Date();

    const ordersByBranch = {};

    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      const cell = row[0];
      const date = !cell
        ? null
        : !isNaN(cell) && Number(cell) > 30000
        ? serialToDate(Number(cell))
        : new Date(cell);
      if (!date || !isSameMonth(date, now)) continue;

      const branchName = (row[1] || '').toString().trim();
      if (!branchName) continue;

      const requestedBy = (row[2] || '').toString().trim();
      const productCode = row[3];
      const productName = row[4];
      const unitPrice = parseFloat(row[5]) || 0;
      const qty = parseInt(row[8]) || 0;
      const subtotal = parseFloat(row[6]) || unitPrice * qty;
      const category = row[7] || '';

      if (!ordersByBranch[branchName]) {
        ordersByBranch[branchName] = {
          branchName,
          total: 0,
          createdAt: date,
          requestors: new Set(),
          items: [],
        };
      }

      const entry = ordersByBranch[branchName];
      entry.total += subtotal;
      entry.items.push({
        productCode,
        productName,
        unitPrice,
        quantity: qty,
        subtotal,
        category,
      });
      if (requestedBy) entry.requestors.add(requestedBy);
      if (date && (!entry.createdAt || date < entry.createdAt)) {
        entry.createdAt = date;
      }
    }

    const orders = Object.values(ordersByBranch).map((entry) => {
      const requestorsArr = Array.from(entry.requestors);
      const requestedBy =
        requestorsArr.length === 0
          ? ''
          : requestorsArr.length === 1
          ? requestorsArr[0]
          : 'أكثر من مستخدم';

      const createdAtStr = entry.createdAt
        ? entry.createdAt.toLocaleString('sv-SE', {
            timeZone: 'Africa/Cairo',
          })
        : '';

      return {
        orderId: entry.branchName,
        branchName: entry.branchName,
        requestedBy,
        createdAt: createdAtStr,
        total: entry.total,
        items: entry.items,
      };
    });

    res.json({ success: true, orders });
  } catch (err) {
    logDebug('Error in pendingOrders', {
      error: err.message,
      stack: err.stack,
    });
    res.status(500).json({
      success: false,
      message: 'حدث خطأ في تحميل الطلبات المعلقة',
    });
  }
});

// POST /api/approveOrder
// Body: { username, orderId } where orderId = "AA13__waiting"
// OR:   { username, serial } where serial = "AA13" (treated as waiting)
app.post('/api/approveOrder', async (req, res) => {
  try {
    const { username, orderId, serial } = req.body;

    const composite = (orderId || serial || '').toString();
    const [serialRaw, statusRaw] = composite.split('__');
    const orderSerial = (serialRaw || '').trim();
    let status = (statusRaw || '').trim().toLowerCase();
    if (!status) status = 'waiting';

    if (!username || !orderSerial) {
      return res
        .status(400)
        .json({ success: false, message: 'البيانات غير مكتملة' });
    }

    if (status !== 'waiting') {
      return res.status(400).json({
        success: false,
        message: 'لا يمكن اعتماد طلب بهذه الحالة',
      });
    }

    const sheets = getSheetsClient();
    const credentialsSheetId = process.env.GOOGLE_CREDENTIALS_SHEET_ID;
    const userInfo = await getUserInfoByUsername(
      sheets,
      credentialsSheetId,
      username
    );

    if (!userInfo) {
      return res
        .status(400)
        .json({ success: false, message: 'المستخدم غير موجود' });
    }
    if ((userInfo.level || '').toUpperCase() !== 'L2') {
      return res.status(403).json({
        success: false,
        message: 'غير مسموح بالموافقة على الطلبات',
      });
    }

    const BUDGET_SHEET_ID = userInfo.budgetSheetId;

    const waitingResp = await sheets.spreadsheets.values.get({
      spreadsheetId: BUDGET_SHEET_ID,
      range: 'Waiting for Approval',
    });
    const rows = waitingResp.data.values || [];
    const now = new Date();

    const rowsToClear = [];
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      const cell = row[0];
      const date = !cell
        ? null
        : !isNaN(cell) && Number(cell) > 30000
        ? serialToDate(Number(cell))
        : new Date(cell);
      if (!date || !isSameMonth(date, now)) continue;

      const rowSerial = (row[10] || '').toString().trim();
      if (rowSerial !== orderSerial) continue;

      rowsToClear.push(i);
    }

    if (rowsToClear.length === 0) {
      return res.status(400).json({
        success: false,
        message: 'لا يوجد طلبات معلقة لهذا الرقم في هذا الشهر',
      });
    }

    const appendValues = rowsToClear.map((idx) => {
      const row = rows[idx] || [];
      const get = (r, col) =>
        r[col] !== undefined && r[col] !== null ? r[col] : '';
      return [
        get(row, 0),
        get(row, 1),
        get(row, 2),
        get(row, 3),
        get(row, 4),
        get(row, 5),
        get(row, 6),
        get(row, 7),
        get(row, 8),
        get(row, 9),
        get(row, 10),
      ];
    });

    await appendRowsAtoK(sheets, BUDGET_SHEET_ID, 'Final Orders', appendValues);

    const clearData = rowsToClear.map((idx) => ({
      range: `Waiting for Approval!A${idx + 1}:K${idx + 1}`,
      values: [['', '', '', '', '', '', '', '', '', '', '']],
    }));
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: BUDGET_SHEET_ID,
      resource: {
        valueInputOption: 'USER_ENTERED',
        data: clearData,
      },
    });

    logDebug('Approved order via /api/approveOrder', {
      orderSerial,
      items: rowsToClear.length,
    });
    res.json({ success: true });
  } catch (err) {
    logDebug('Error in approveOrder', {
      error: err.message,
      stack: err.stack,
    });
    res.status(500).json({
      success: false,
      message: 'حدث خطأ أثناء اعتماد الطلب',
    });
  }
});
/* ========================================================================== */
/* === [4f] NEW L2 endpoints: update & cancel waiting orders               === */
/* ========================================================================== */

/**
 * POST /api/updateWaitingOrder
 * Body: { username, orderId, items }
 * - orderId: "AA13__waiting" OR just "AA13"
 * - items: [{ productCode, quantity }, ...]
 * Only affects rows in "Waiting for Approval" with this serial (current month).
 */
app.post('/api/updateWaitingOrder', async (req, res) => {
  try {
    const { username, orderId, items } = req.body;

    const composite = (orderId || '').toString();
    const [serialRaw, statusRaw] = composite.split('__');
    const orderSerial = (serialRaw || '').trim();
    let status = (statusRaw || '').trim().toLowerCase();
    if (!status) status = 'waiting';

    if (
      !username ||
      !orderSerial ||
      !Array.isArray(items) ||
      items.length === 0
    ) {
      return res
        .status(400)
        .json({ success: false, message: 'البيانات غير مكتملة' });
    }

    if (status !== 'waiting') {
      return res.status(400).json({
        success: false,
        message: 'لا يمكن تعديل طلب بهذه الحالة',
      });
    }

    const sheets = getSheetsClient();
    const credentialsSheetId = process.env.GOOGLE_CREDENTIALS_SHEET_ID;
    const userInfo = await getUserInfoByUsername(
      sheets,
      credentialsSheetId,
      username
    );

    if (!userInfo) {
      return res
        .status(400)
        .json({ success: false, message: 'المستخدم غير موجود' });
    }
    if ((userInfo.level || '').toUpperCase() !== 'L2') {
      return res.status(403).json({
        success: false,
        message: 'غير مسموح بتعديل الطلبات',
      });
    }

    const BUDGET_SHEET_ID = userInfo.budgetSheetId;

    const waitingResp = await sheets.spreadsheets.values.get({
      spreadsheetId: BUDGET_SHEET_ID,
      range: 'Waiting for Approval',
    });
    const rows = waitingResp.data.values || [];
    const now = new Date();

    // index rows by productCode for this serial (current month)
    const index = {};
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      const cell = row[0];
      const date = !cell
        ? null
        : !isNaN(cell) && Number(cell) > 30000
        ? serialToDate(Number(cell))
        : new Date(cell);
      if (!date || !isSameMonth(date, now)) continue;

      const rowSerial = (row[10] || '').toString().trim();
      if (rowSerial !== orderSerial) continue;

      const code = row[3];
      if (!code) continue;
      index[code] = i; // 0-based row index
    }

    const batch = [];
    for (const item of items) {
      const code = item.productCode;
      if (!code || index[code] === undefined) continue;

      const rowIdx = index[code];
      const row = rows[rowIdx];
      const unitPrice = parseFloat(row[5]) || 0;
      const qty = Math.max(0, parseInt(item.quantity) || 0);
      const subtotal = unitPrice * qty;
      const excelRow = rowIdx + 1; // 1-based

      batch.push(
        { range: `Waiting for Approval!I${excelRow}`, values: [[qty]] },
        { range: `Waiting for Approval!G${excelRow}`, values: [[subtotal]] }
      );
    }

    if (batch.length === 0) {
      return res.status(400).json({
        success: false,
        message: 'لم يتم العثور على بنود لتعديلها',
      });
    }

    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: BUDGET_SHEET_ID,
      resource: { valueInputOption: 'USER_ENTERED', data: batch },
    });

    logDebug('Updated waiting order via /api/updateWaitingOrder', {
      orderSerial,
      updatedLines: batch.length / 2,
    });

    res.json({ success: true });
  } catch (err) {
    logDebug('Error in updateWaitingOrder', {
      error: err.message,
      stack: err.stack,
    });
    res.status(500).json({
      success: false,
      message: 'حدث خطأ أثناء حفظ التعديلات',
    });
  }
});

/**
 * POST /api/cancelOrder
 * Body: { username, orderId }
 * - orderId: "AA13__waiting" OR just "AA13"
 * Moves rows from "Waiting for Approval" to "Cancelled Orders".
 */
app.post('/api/cancelOrder', async (req, res) => {
  try {
    const { username, orderId } = req.body;

    const composite = (orderId || '').toString();
    const [serialRaw, statusRaw] = composite.split('__');
    const orderSerial = (serialRaw || '').trim();
    let status = (statusRaw || '').trim().toLowerCase();
    if (!status) status = 'waiting';

    if (!username || !orderSerial) {
      return res
        .status(400)
        .json({ success: false, message: 'البيانات غير مكتملة' });
    }

    if (status !== 'waiting') {
      return res.status(400).json({
        success: false,
        message: 'لا يمكن إلغاء طلب بهذه الحالة',
      });
    }

    const sheets = getSheetsClient();
    const credentialsSheetId = process.env.GOOGLE_CREDENTIALS_SHEET_ID;
    const userInfo = await getUserInfoByUsername(
      sheets,
      credentialsSheetId,
      username
    );

    if (!userInfo) {
      return res
        .status(400)
        .json({ success: false, message: 'المستخدم غير موجود' });
    }
    if ((userInfo.level || '').toUpperCase() !== 'L2') {
      return res.status(403).json({
        success: false,
        message: 'غير مسموح بإلغاء الطلبات',
      });
    }

    const BUDGET_SHEET_ID = userInfo.budgetSheetId;

    const waitingResp = await sheets.spreadsheets.values.get({
      spreadsheetId: BUDGET_SHEET_ID,
      range: 'Waiting for Approval',
    });
    const rows = waitingResp.data.values || [];
    const now = new Date();

    const rowsToClear = [];
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      const cell = row[0];
      const date = !cell
        ? null
        : !isNaN(cell) && Number(cell) > 30000
        ? serialToDate(Number(cell))
        : new Date(cell);
      if (!date || !isSameMonth(date, now)) continue;

      const rowSerial = (row[10] || '').toString().trim();
      if (rowSerial !== orderSerial) continue;

      rowsToClear.push(i);
    }

    if (rowsToClear.length === 0) {
      return res.status(400).json({
        success: false,
        message: 'لا يوجد طلبات معلقة لهذا الرقم في هذا الشهر',
      });
    }

    const appendValues = rowsToClear.map((idx) => {
      const row = rows[idx] || [];
      const get = (r, col) =>
        r[col] !== undefined && r[col] !== null ? r[col] : '';
      return [
        get(row, 0),
        get(row, 1),
        get(row, 2),
        get(row, 3),
        get(row, 4),
        get(row, 5),
        get(row, 6),
        get(row, 7),
        get(row, 8),
        get(row, 9),
        get(row, 10),
      ];
    });

    await appendRowsAtoK(
      sheets,
      BUDGET_SHEET_ID,
      'Cancelled Orders',
      appendValues
    );

    const clearData = rowsToClear.map((idx) => ({
      range: `Waiting for Approval!A${idx + 1}:K${idx + 1}`,
      values: [['', '', '', '', '', '', '', '', '', '', '']],
    }));
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: BUDGET_SHEET_ID,
      resource: {
        valueInputOption: 'USER_ENTERED',
        data: clearData,
      },
    });

    logDebug('Cancelled order via /api/cancelOrder', {
      orderSerial,
      items: rowsToClear.length,
    });

    res.json({ success: true });
  } catch (err) {
    logDebug('Error in cancelOrder', {
      error: err.message,
      stack: err.stack,
    });
    res.status(500).json({
      success: false,
      message: 'حدث خطأ أثناء إلغاء الطلب',
    });
  }
});


/* ========================================================================== */
/* === [5] Submit Order Endpoint (L1 -> Waiting, L2 -> Final)             === */
/* ========================================================================== */
app.post('/api/submitOrder', async (req, res) => {
  try {
    const { branchName, orderItems, userType, username } = req.body;
    logDebug('Received submitOrder', {
      branchName,
      userType,
      username,
      orderItems,
    });

    if (!branchName || !Array.isArray(orderItems) || orderItems.length === 0) {
      logDebug('Order data missing or invalid', {
        branchName,
        orderItems,
      });
      return res.status(400).json({
        success: false,
        message: 'بيانات الطلب غير مكتملة',
      });
    }

    const sheets = getSheetsClient();
    const credentialsSheetId = process.env.GOOGLE_CREDENTIALS_SHEET_ID;

    const userInfo = await getUserInfoByUsername(
      sheets,
      credentialsSheetId,
      username
    );
    if (!userInfo) {
      return res.status(400).json({
        success: false,
        message: 'بيانات المستخدم غير صحيحة',
      });
    }

    const BUDGET_SHEET_ID = userInfo.budgetSheetId;
    const userLevel = (userInfo.level || 'L1').toUpperCase();

    if (userLevel !== 'L2') {
      const userBranch = (userInfo.branch || '').trim();
      if (userBranch && userBranch !== branchName) {
        return res.status(403).json({
          success: false,
          message: 'غير مسموح لك بالطلب لهذا الفرع',
        });
      }
    }

    const targetSheet =
      userLevel === 'L2' ? 'Final Orders' : 'Waiting for Approval';

    const orderSerial = await getNextOrderSerial(sheets, BUDGET_SHEET_ID);

    const now = new Date();
    const dateString = now
      .toLocaleString('sv-SE', { timeZone: 'Africa/Cairo' })
      .replace('T', ' ');

    const appendValues = orderItems.map((i) => {
      const price = Number(i.unitPrice) || 0;
      const qty = Number(i.quantity) || 0;
      const subtotal =
        typeof i.subtotal === 'number' ? i.subtotal : price * qty;
      const cat = i.category || '';
      return [
        dateString,
        branchName,
        username,
        i.productCode,
        i.productName,
        price,
        subtotal,
        cat,
        qty,
        '',
        orderSerial,
      ];
    });

    await appendRowsAtoK(sheets, BUDGET_SHEET_ID, targetSheet, appendValues);

    logDebug('submitOrder saved lines with serial in column K', {
      branchName,
      targetSheet,
      lines: appendValues.length,
      orderSerial,
    });

    res.json({ success: true, orderSerial });
  } catch (err) {
    logDebug('Error submitting order', {
      error: err.message,
      stack: err.stack,
    });
    res.status(500).json({
      success: false,
      message: 'حدث خطأ أثناء إرسال الطلب',
    });
  }
});

/* ========================================================================== */
/* === [6] Update Previous Orders (returns)                                === */
/* ========================================================================== */
app.post('/api/updatePreviousOrders', async (req, res) => {
  try {
    const { branchName, userType, updatedOrders, username } = req.body;
    logDebug('Received updatePreviousOrders', {
      branchName,
      userType,
      username,
      updatedOrders,
    });

    if (
      !branchName ||
      !Array.isArray(updatedOrders) ||
      updatedOrders.length === 0
    ) {
      logDebug('Missing data for updatePreviousOrders', {
        branchName,
        updatedOrders,
      });
      return res.status(400).json({
        success: false,
        message: 'بيانات الطلبات غير مكتملة',
      });
    }

    const sheets = getSheetsClient();
    const credentialsSheetId = process.env.GOOGLE_CREDENTIALS_SHEET_ID;
    const BUDGET_SHEET_ID =
      await resolveBudgetSheetIdFromF2ByUsername(
        sheets,
        credentialsSheetId,
        username
      );

    const ordersResp = await sheets.spreadsheets.values.get({
      spreadsheetId: BUDGET_SHEET_ID,
      range: 'Final Orders',
    });
    const rows = ordersResp.data.values || [];
    const now = new Date();

    const latestIndex = {};
    for (let i = 1; i < rows.length; i++) {
      const cell = rows[i][0];
      const date = !cell
        ? null
        : !isNaN(cell) && Number(cell) > 30000
        ? serialToDate(Number(cell))
        : new Date(cell);
      if (!date || !isSameMonth(date, now)) continue;
      if (rows[i][1] !== branchName) continue;
      const code = rows[i][3];
      latestIndex[code] = i;
    }

    const batch = [];
    updatedOrders.forEach((o) => {
      let idx =
        typeof o.rowIndex === 'number' &&
        o.rowIndex > 0 &&
        o.rowIndex < rows.length
          ? o.rowIndex
          : latestIndex[o.productCode];

      if (idx === undefined) return;

      const unitPrice = parseFloat(rows[idx][5]) || 0;
      const qty = Math.max(0, parseInt(o.quantity) || 0);
      const newSubtotal = unitPrice * qty;
      const r = idx + 1;

      batch.push(
        { range: `Final Orders!I${r}`, values: [[qty]] },
        { range: `Final Orders!G${r}`, values: [[newSubtotal]] }
      );
    });

    if (batch.length === 0) {
      return res.status(400).json({
        success: false,
        message: 'لم يتم العثور على بيانات لتحديثها',
      });
    }

    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: BUDGET_SHEET_ID,
      resource: { valueInputOption: 'USER_ENTERED', data: batch },
    });

    logDebug('Batch update of previous orders complete', {
      updatedLines: batch.length / 2,
    });
    res.json({ success: true });
  } catch (err) {
    logDebug('Error updating previous orders', {
      error: err.message,
      stack: err.stack,
    });
    res.status(500).json({
      success: false,
      message: 'حدث خطأ أثناء تحديث الطلبيات',
    });
  }
});

/* ========================================================================== */
/* === [6c] L2: Orders summary by status + details (NEW)                  === */
/* ========================================================================== */

// Shared handler for orders summary by status (grouped by Serial)
async function handleOrdersSummaryForL2(req, res) {
  try {
    const username = (req.query.username || '').trim();
    if (!username) {
      return res
        .status(400)
        .json({ success: false, message: 'Missing username' });
    }

    const sheets = getSheetsClient();
    const credentialsSheetId = process.env.GOOGLE_CREDENTIALS_SHEET_ID;
    const userInfo = await getUserInfoByUsername(
      sheets,
      credentialsSheetId,
      username
    );

    if (!userInfo) {
      return res
        .status(400)
        .json({ success: false, message: 'المستخدم غير موجود' });
    }
    if ((userInfo.level || '').toUpperCase() !== 'L2') {
      return res.status(403).json({
        success: false,
        message: 'غير مسموح بعرض قائمة الطلبات',
      });
    }

    const BUDGET_SHEET_ID = userInfo.budgetSheetId;

    const productsResp = await sheets.spreadsheets.values.get({
      spreadsheetId: BUDGET_SHEET_ID,
      range: 'Product Catalog',
    });
    const productRows = productsResp.data.values || [];
    const codeToImage = {};
    for (let i = 1; i < productRows.length; i++) {
      const row = productRows[i];
      const code = row[0];
      if (code) {
        codeToImage[code] = row[4] || '';
      }
    }

    const now = new Date();
    const summaryMap = {};

    function processSheet(rows, statusKey) {
      if (!rows || rows.length <= 1) return;
      for (let i = 1; i < rows.length; i++) {
        const row = rows[i];
        const cell = row[0];
        const date = !cell
          ? null
          : !isNaN(cell) && Number(cell) > 30000
          ? serialToDate(Number(cell))
          : new Date(cell);
        if (!date || !isSameMonth(date, now)) continue;

        const branchName = (row[1] || '').toString().trim();
        if (!branchName) continue;

        const requestedByVal = (row[2] || '').toString().trim();
        const productCode = row[3];
        const productName = row[4];
        const unitPrice = parseFloat(row[5]) || 0;
        const qty = parseInt(row[8]) || 0;
        const subtotal = parseFloat(row[6]) || unitPrice * qty;
        const category = row[7] || '';
        const serial = (row[10] || '').toString().trim();
        if (!serial) continue;

        const key = `${serial}__${statusKey}`;
        if (!summaryMap[key]) {
          summaryMap[key] = {
            serial,
            branchName,
            status: statusKey,
            total: 0,
            createdAt: date,
            requestors: new Set(),
            items: [],
          };
        }
        const entry = summaryMap[key];
        entry.total += subtotal;
        if (date && (!entry.createdAt || date < entry.createdAt)) {
          entry.createdAt = date;
        }
        if (requestedByVal) entry.requestors.add(requestedByVal);

        entry.items.push({
          productCode,
          productName,
          unitPrice,
          quantity: qty,
          subtotal,
          category,
          imageUrl: codeToImage[productCode] || '',
        });
      }
    }

    const waitingResp = await sheets.spreadsheets.values.get({
      spreadsheetId: BUDGET_SHEET_ID,
      range: 'Waiting for Approval',
    });
    processSheet(waitingResp.data.values || [], 'Waiting');

    const finalResp = await sheets.spreadsheets.values.get({
      spreadsheetId: BUDGET_SHEET_ID,
      range: 'Final Orders',
    });
    processSheet(finalResp.data.values || [], 'Approved');

    try {
      const cancelledResp = await sheets.spreadsheets.values.get({
        spreadsheetId: BUDGET_SHEET_ID,
        range: 'Cancelled Orders',
      });
      processSheet(cancelledResp.data.values || [], 'Cancelled');
    } catch (e) {
      logDebug('Cancelled Orders sheet read failed (maybe missing)', {
        error: e.message,
      });
    }

    const orders = Object.values(summaryMap).map((entry) => {
      const creators = Array.from(entry.requestors);
      const requestedBy =
        creators.length === 0
          ? ''
          : creators.length === 1
          ? creators[0]
          : 'أكثر من مستخدم';

      const createdAtStr = entry.createdAt
        ? entry.createdAt.toLocaleString('sv-SE', {
            timeZone: 'Africa/Cairo',
          })
        : '';

      return {
        orderId: `${entry.serial}__${entry.status.toLowerCase()}`,
        serial: entry.serial,
        branchName: entry.branchName,
        status: entry.status,
        requestedBy,
        createdAt: createdAtStr,
        total: entry.total,
        items: entry.items,
      };
    });

    res.json({ success: true, orders });
  } catch (err) {
    logDebug('Error in ordersSummaryForL2', {
      error: err.message,
      stack: err.stack,
    });
    res.status(500).json({
      success: false,
      message: 'حدث خطأ في تحميل قائمة الطلبات',
    });
  }
}

// Support multiple possible paths used by the frontend
app.get(
  ['/api/ordersSummaryForL2', '/api/ordersSummary', '/api/ordersSummaryByStatus'],
  handleOrdersSummaryForL2
);

// Shared handler for order details by branch + status OR serial + status
async function handleOrderDetailsForL2(req, res) {
  try {
    const username = (req.query.username || '').trim();
    const branchNameQuery = (req.query.branchName || '').trim();
    const statusRaw = (req.query.status || '').trim().toLowerCase();
    const serialQuery = (req.query.serial || '').trim();

    if (!username || (!branchNameQuery && !serialQuery)) {
      return res
        .status(400)
        .json({ success: false, message: 'البيانات غير مكتملة' });
    }

    let status = 'waiting';
    if (statusRaw === 'approved') status = 'approved';
    else if (statusRaw === 'cancelled') status = 'cancelled';

    const sheetName =
      status === 'waiting'
        ? 'Waiting for Approval'
        : status === 'approved'
        ? 'Final Orders'
        : 'Cancelled Orders';

    const sheets = getSheetsClient();
    const credentialsSheetId = process.env.GOOGLE_CREDENTIALS_SHEET_ID;
    const userInfo = await getUserInfoByUsername(
      sheets,
      credentialsSheetId,
      username
    );

    if (!userInfo) {
      return res
        .status(400)
        .json({ success: false, message: 'المستخدم غير موجود' });
    }
    if ((userInfo.level || '').toUpperCase() !== 'L2') {
      return res.status(403).json({
        success: false,
        message: 'غير مسموح بعرض تفاصيل الطلب',
      });
    }

    const BUDGET_SHEET_ID = userInfo.budgetSheetId;

    const productsResp = await sheets.spreadsheets.values.get({
      spreadsheetId: BUDGET_SHEET_ID,
      range: 'Product Catalog',
    });
    const productRows = productsResp.data.values || [];
    const codeToImage = {};
    for (let i = 1; i < productRows.length; i++) {
      const row = productRows[i];
      codeToImage[row[0]] = row[4] || '';
    }

    const sheetResp = await sheets.spreadsheets.values.get({
      spreadsheetId: BUDGET_SHEET_ID,
      range: sheetName,
    });
    const rows = sheetResp.data.values || [];
    const now = new Date();
    const items = [];
    let effectiveBranchName = branchNameQuery;

    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      const cell = row[0];
      const date = !cell
        ? null
        : !isNaN(cell) && Number(cell) > 30000
        ? serialToDate(Number(cell))
        : new Date(cell);
      if (!date || !isSameMonth(date, now)) continue;

      const rowBranch = (row[1] || '').toString().trim();
      const rowSerial = (row[10] || '').toString().trim();

      if (serialQuery) {
        if (rowSerial !== serialQuery) continue;
        if (!effectiveBranchName && rowBranch) {
          effectiveBranchName = rowBranch;
        }
      } else {
        if (rowBranch !== branchNameQuery) continue;
      }

      const productCode = row[3];
      const productName = row[4];
      const unitPrice = parseFloat(row[5]) || 0;
      const qty = parseInt(row[8]) || 0;
      const subtotal = parseFloat(row[6]) || unitPrice * qty;
      const category = row[7] || '';

      items.push({
        productCode,
        productName,
        unitPrice,
        quantity: qty,
        subtotal,
        category,
        imageUrl: codeToImage[productCode] || '',
      });
    }

    res.json({
      success: true,
      branchName: effectiveBranchName || '',
      serial: serialQuery || null,
      status,
      items,
    });
  } catch (err) {
    logDebug('Error in orderDetailsForL2', {
      error: err.message,
      stack: err.stack,
    });
    res.status(500).json({
      success: false,
      message: 'حدث خطأ أثناء تحميل تفاصيل الطلب',
    });
  }
}

// Support multiple paths to be safe with frontend
app.get(
  ['/api/orderDetailsForL2', '/api/orderDetailsByStatus'],
  handleOrderDetailsForL2
);

/* ========================================================================== */
/* === [6e] NEW: Export orders as Excel (L2) – MULTI ORDERS (GET)          === */
/* ========================================================================== */
/**
 * GET /api/exportOrdersExcel
 * Query: ?username=...&serials=AA10,AA9,AA6
 * Only L2 users. Exports the selected orders (by serial) as .xlsx.
 */
app.get('/api/exportOrdersExcel', async (req, res) => {
  try {
    const username = (req.query.username || '').trim();
    const serialsParam = (req.query.serials || '').trim();

    // serials comes like "AA10,AA9,AA6"
    const serials = serialsParam
      .split(',')
      .map((s) => s.trim())
      .filter((s) => s);

    if (!username || serials.length === 0) {
      return res.status(400).json({
        success: false,
        message: 'البيانات غير مكتملة',
      });
    }

    const sheets = getSheetsClient();
    const credentialsSheetId = process.env.GOOGLE_CREDENTIALS_SHEET_ID;
    const userInfo = await getUserInfoByUsername(
      sheets,
      credentialsSheetId,
      username
    );

    if (!userInfo) {
      return res
        .status(400)
        .json({ success: false, message: 'المستخدم غير موجود' });
    }
    if ((userInfo.level || '').toUpperCase() !== 'L2') {
      return res.status(403).json({
        success: false,
        message: 'غير مسموح بتحميل ملف إكسل للطلبات',
      });
    }

    const BUDGET_SHEET_ID = userInfo.budgetSheetId;
    const now = new Date();
    const selectedSerials = new Set(serials);
    const ordersMap = {}; // keyed ONLY by serial

    function processSheet(rows, statusKey) {
      if (!rows || rows.length <= 1) return;
      for (let i = 1; i < rows.length; i++) {
        const row = rows[i];
        const cell = row[0];
        const date = !cell
          ? null
          : !isNaN(cell) && Number(cell) > 30000
          ? serialToDate(Number(cell))
          : new Date(cell);
        if (!date || !isSameMonth(date, now)) continue;

        const branchName = (row[1] || '').toString().trim();
        if (!branchName) continue;

        const requestedByVal = (row[2] || '').toString().trim();
        const productCode = row[3];
        const productName = row[4];
        const unitPrice = parseFloat(row[5]) || 0;
        const qty = parseInt(row[8]) || 0;
        const subtotal = parseFloat(row[6]) || unitPrice * qty;
        const category = row[7] || '';
        const serial = (row[10] || '').toString().trim();
        if (!serial || !selectedSerials.has(serial)) continue;

        // ONE entry per serial – last status wins if duplicated
        if (!ordersMap[serial]) {
          ordersMap[serial] = {
            serial,
            status: statusKey,
            branchName,
            createdAt: date,
            requestors: new Set(),
            items: [],
          };
        }
        const entry = ordersMap[serial];
        entry.status = statusKey;
        if (date && (!entry.createdAt || date < entry.createdAt)) {
          entry.createdAt = date;
        }
        if (requestedByVal) entry.requestors.add(requestedByVal);

        entry.items.push({
          productCode,
          productName,
          category,
          quantity: qty,
          unitPrice,
          subtotal,
        });
      }
    }

    // Read from all three sheets and keep only selected serials
    const waitingResp = await sheets.spreadsheets.values.get({
      spreadsheetId: BUDGET_SHEET_ID,
      range: 'Waiting for Approval',
    });
    processSheet(waitingResp.data.values || [], 'Waiting');

    const finalResp = await sheets.spreadsheets.values.get({
      spreadsheetId: BUDGET_SHEET_ID,
      range: 'Final Orders',
    });
    processSheet(finalResp.data.values || [], 'Approved');

    try {
      const cancelledResp = await sheets.spreadsheets.values.get({
        spreadsheetId: BUDGET_SHEET_ID,
        range: 'Cancelled Orders',
      });
      processSheet(cancelledResp.data.values || [], 'Cancelled');
    } catch (e) {
      logDebug('Cancelled Orders sheet read failed (maybe missing)', {
        error: e.message,
      });
    }

    const serialKeys = Object.keys(ordersMap);
    if (serialKeys.length === 0) {
      return res.status(400).json({
        success: false,
        message: 'لم يتم العثور على بيانات للطلبات المحددة',
      });
    }

    // Build Excel workbook
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Orders');

    worksheet.columns = [
      { header: 'Order Serial', key: 'serial', width: 15 },
      { header: 'Status', key: 'status', width: 12 },
      { header: 'Branch', key: 'branch', width: 25 },
      { header: 'Requested By', key: 'requestedBy', width: 25 },
      { header: 'Order Date', key: 'createdAt', width: 22 },
      { header: 'Product Code', key: 'productCode', width: 15 },
      { header: 'Product Name', key: 'productName', width: 40 },
      { header: 'Category', key: 'category', width: 20 },
      { header: 'Quantity', key: 'quantity', width: 10 },
      { header: 'Unit Price', key: 'unitPrice', width: 14 },
      { header: 'Subtotal', key: 'subtotal', width: 14 },
    ];

    serialKeys.forEach((serial) => {
      const entry = ordersMap[serial];
      if (!entry) return;

      const creators = Array.from(entry.requestors);
      const requestedBy =
        creators.length === 0
          ? ''
          : creators.length === 1
          ? creators[0]
          : 'أكثر من مستخدم';

      const createdAtStr = entry.createdAt
        ? entry.createdAt.toLocaleString('sv-SE', {
            timeZone: 'Africa/Cairo',
          })
        : '';

      entry.items.forEach((item) => {
        worksheet.addRow({
          serial: entry.serial,
          status: entry.status,
          branch: entry.branchName,
          requestedBy,
          createdAt: createdAtStr,
          productCode: item.productCode,
          productName: item.productName,
          category: item.category,
          quantity: item.quantity,
          unitPrice: item.unitPrice,
          subtotal: item.subtotal,
        });
      });

      // blank separator row between different orders
      worksheet.addRow({});
    });

    const buffer = await workbook.xlsx.writeBuffer();

    res.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    );
    res.setHeader(
      'Content-Disposition',
      'attachment; filename="orders.xlsx"'
    );
    res.send(Buffer.from(buffer));
  } catch (err) {
    logDebug('Error in exportOrdersExcel (GET multi)', {
      error: err.message,
      stack: err.stack,
    });
    res.status(500).json({
      success: false,
      message: 'حدث خطأ أثناء تحميل ملف الإكسل',
    });
  }
});

/**
 * GET /api/exportOrderExcel
 * Query: ?username=...&serial=AA10
 * Only L2 users. Exports ONE order (any status) as .xlsx.
 */
app.get('/api/exportOrderExcel', async (req, res) => {
  try {
    const username = (req.query.username || '').trim();
    const serialQuery = (req.query.serial || '').trim();

    if (!username || !serialQuery) {
      return res.status(400).json({
        success: false,
        message: 'البيانات غير مكتملة',
      });
    }

    const sheets = getSheetsClient();
    const credentialsSheetId = process.env.GOOGLE_CREDENTIALS_SHEET_ID;
    const userInfo = await getUserInfoByUsername(
      sheets,
      credentialsSheetId,
      username
    );

    if (!userInfo) {
      return res
        .status(400)
        .json({ success: false, message: 'المستخدم غير موجود' });
    }
    if ((userInfo.level || '').toUpperCase() !== 'L2') {
      return res.status(403).json({
        success: false,
        message: 'غير مسموح بتحميل ملف إكسل للطلبات',
      });
    }

    const BUDGET_SHEET_ID = userInfo.budgetSheetId;
    const now = new Date();
    const orders = []; // will contain all statuses for this serial

    function processSheet(rows, statusKey) {
      if (!rows || rows.length <= 1) return;
      for (let i = 1; i < rows.length; i++) {
        const row = rows[i];
        const cell = row[0];
        const date = !cell
          ? null
          : !isNaN(cell) && Number(cell) > 30000
          ? serialToDate(Number(cell))
          : new Date(cell);
        if (!date || !isSameMonth(date, now)) continue;

        const branchName = (row[1] || '').toString().trim();
        if (!branchName) continue;

        const requestedByVal = (row[2] || '').toString().trim();
        const productCode = row[3];
        const productName = row[4];
        const unitPrice = parseFloat(row[5]) || 0;
        const qty = parseInt(row[8]) || 0;
        const subtotal = parseFloat(row[6]) || unitPrice * qty;
        const category = row[7] || '';
        const serial = (row[10] || '').toString().trim();
        if (!serial || serial !== serialQuery) continue;

        let entry = orders.find(
          (o) =>
            o.serial === serial && o.status === statusKey && o.branchName === branchName
        );
        if (!entry) {
          entry = {
            serial,
            status: statusKey,
            branchName,
            createdAt: date,
            requestors: new Set(),
            items: [],
          };
          orders.push(entry);
        }

        if (date && (!entry.createdAt || date < entry.createdAt)) {
          entry.createdAt = date;
        }
        if (requestedByVal) entry.requestors.add(requestedByVal);

        entry.items.push({
          productCode,
          productName,
          category,
          quantity: qty,
          unitPrice,
          subtotal,
        });
      }
    }

    // Read sheets and collect this serial from all statuses
    const waitingResp = await sheets.spreadsheets.values.get({
      spreadsheetId: BUDGET_SHEET_ID,
      range: 'Waiting for Approval',
    });
    processSheet(waitingResp.data.values || [], 'Waiting');

    const finalResp = await sheets.spreadsheets.values.get({
      spreadsheetId: BUDGET_SHEET_ID,
      range: 'Final Orders',
    });
    processSheet(finalResp.data.values || [], 'Approved');

    try {
      const cancelledResp = await sheets.spreadsheets.values.get({
        spreadsheetId: BUDGET_SHEET_ID,
        range: 'Cancelled Orders',
      });
      processSheet(cancelledResp.data.values || [], 'Cancelled');
    } catch (e) {
      logDebug('Cancelled Orders sheet read failed (maybe missing)', {
        error: e.message,
      });
    }

    if (orders.length === 0) {
      return res.status(400).json({
        success: false,
        message: 'لم يتم العثور على بيانات لهذا الطلب',
      });
    }

    // Build Excel workbook for this single serial
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Order');

    worksheet.columns = [
      { header: 'Order Serial', key: 'serial', width: 15 },
      { header: 'Status', key: 'status', width: 12 },
      { header: 'Branch', key: 'branch', width: 25 },
      { header: 'Requested By', key: 'requestedBy', width: 25 },
      { header: 'Order Date', key: 'createdAt', width: 22 },
      { header: 'Product Code', key: 'productCode', width: 15 },
      { header: 'Product Name', key: 'productName', width: 40 },
      { header: 'Category', key: 'category', width: 20 },
      { header: 'Quantity', key: 'quantity', width: 10 },
      { header: 'Unit Price', key: 'unitPrice', width: 14 },
      { header: 'Subtotal', key: 'subtotal', width: 14 },
    ];

    orders.forEach((entry) => {
      const creators = Array.from(entry.requestors);
      const requestedBy =
        creators.length === 0
          ? ''
          : creators.length === 1
          ? creators[0]
          : 'أكثر من مستخدم';

      const createdAtStr = entry.createdAt
        ? entry.createdAt.toLocaleString('sv-SE', {
            timeZone: 'Africa/Cairo',
          })
        : '';

      entry.items.forEach((item) => {
        worksheet.addRow({
          serial: entry.serial,
          status: entry.status,
          branch: entry.branchName,
          requestedBy,
          createdAt: createdAtStr,
          productCode: item.productCode,
          productName: item.productName,
          category: item.category,
          quantity: item.quantity,
          unitPrice: item.unitPrice,
          subtotal: item.subtotal,
        });
      });

      // separator between statuses if same serial exists in multiple sheets
      worksheet.addRow({});
    });

    const buffer = await workbook.xlsx.writeBuffer();

    res.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    );
    res.setHeader(
      'Content-Disposition',
      `attachment; filename="order_${serialQuery}.xlsx"`
    );
    res.send(Buffer.from(buffer));
  } catch (err) {
    logDebug('Error in exportOrderExcel', {
      error: err.message,
      stack: err.stack,
    });
    res.status(500).json({
      success: false,
      message: 'حدث خطأ أثناء تحميل ملف الإكسل',
    });
  }
});

/* ========================================================================== */
/* === [7] Serve Frontend (SPA)                                            === */
/* ========================================================================== */

// SPA fallback MUST be last, and must point to ../public/main.html
app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, '..', 'public', 'main.html'));
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () =>
  logDebug(`Server running on http://localhost:${PORT}`)
);

// Apps Script for Uptown Shell - Shift Sheet Sync (Version 10)
// v10: department AMOUNTS alongside quantities. doPost {action:'departments',
// date, shift, quantities:[19], amounts:[19]} also upserts the ex-tax $
// amounts per category into a NEW tab 'Department Amounts' in the same
// Department Sales spreadsheet (same header layout, same upsert by
// Date+Shift). doGet ?action=reports also serves departmentAmounts. The
// dashboard Profit tool derives the realised unit sell price per category
// as amount/qty for that date+shift, falling back to the manual sell price
// when there is no amount row. A departments POST still touches ONLY the
// Department Sales spreadsheet — never the Sync sheet or Daily Closing.
// v6: gas profit (total + per grade 87/89/93) from the website payload,
// saved first into the Sync sheet, then linked into Daily Closing AI-AL.
// v7: delete action {action:'delete', date:'YYYY-MM-DD'} — removes the whole
// day from the Sync sheet and unlinks it from Daily Closing (Z-Report rows
// get their shift-sheet columns cleared, partial rows are deleted).
// v8: doGet ?action=reports — serves the closing-report sheets as JSON
// (Daily Closing, Weekly Summary, Daily Breakdown, Monthly Summary,
// Weekly Breakdown) so the website's "Daily Reports" tab can show the
// same Z-Report numbers that go in the emails, per shift / day / week.
// v9: department sales for the dashboard Profit auto-fill. doGet
// ?action=reports also serves departmentSales (the Department Sales sheet,
// one row per Date+Shift with the quantity sold of each of the 19 profit
// categories, in the same order as the dashboard's PROF_CATS). doPost
// {action:'departments', date, shift, quantities:[19 numbers]} upserts one
// row, creating the tab and the header row on first use. Everything runs as
// Spectrum (the sheet is Spectrum-owned), so no sharing with any other
// account is needed — the report runs POST the numbers here instead of
// writing the sheet directly.
// v9 + lottery cash (Nov 10): the website's Ohio lottery section (Shift 3
// only) now sends d.lotteryCash (Instant + Online + Cash = lotterySales).
// doPost writes it into the Sync sheet col AG (Lottery Cash, header AG1),
// after the Raw JSON in AF — the row is 33 values, A-AG. Older uploads
// (before the website sent lotteryCash) leave AG blank; re-upload from
// the website to fill them.
// The shift-sheet data is KEPT in the Shift Sheet Sync sheet (the store),
// and LINKED into the Daily Closing sheet (the main sheet) so all the
// cash / card / EBT / coupon figures show there too — no need to open
// a second sheet to grab the information.
// Deploy as Web App: Execute as Me, Who has access: Anyone.
// Paste this whole file into script.google.com, then Deploy > New deployment > Web app.

var SYNC_SHEET_ID = '1V24ek1xmu88zBqhUKHNAn-PvIX3jTEkGS3sJ0KGQuco'; // Shift Sheet Sync (the store)
var SYNC_SHEET_NAME = 'Sheet1';
var DAILY_SHEET_ID = '1vU0dC7sjTbC41hVZEvSbLyYhRrLwsQGBCvvLOZf3WK4'; // Spectrum Daily Closing Reports
var DAILY_SHEET_NAME = 'Daily Closing';
var WEEKLY_SHEET_ID = '1rR9HkEhefGSbeenycNTq_FGJbhieGBHbY8gTo5A1_WM'; // Spectrum Weekly Reports
var MONTHLY_SHEET_ID = '1DlayYfUJhpHSMFVk3pWInQaNdhc-ezbR7FjNZF_wABU'; // Spectrum Monthly Reports
var DEPT_SHEET_ID = '1KJrpFDXhSiaC4lRmwdWiiWj7iRL86gehEGnIvkWWkww'; // Spectrum Department Sales (profit auto-fill store)
var DEPT_SHEET_NAME = 'Department Sales';
var DEPT_AMOUNTS_SHEET_NAME = 'Department Amounts'; // v10: ex-tax $ amounts per category, same layout/order as Department Sales
// Quantity columns in the Department Sales sheet, same order as the dashboard's PROF_CATS.
var DEPT_CATS = ['Beer', 'Cigarettes', 'Juice', 'Snacks & Chips', 'Tobacco', 'Water', 'Misc. No Tax', 'Energy Drinks', 'Wine', 'Soda', 'Smokeless Tobacco / Chew', 'Dairy Product', 'Candy', 'Gen Grocery', 'Phone Cards', 'TX Groc', 'Frozen Pro', 'Coffee', 'Cigars'];

// Daily Closing columns (1-based, VERIFIED against the live sheet): A Date, B Shift,
// T Expected Deposit to Safe, U Cash Sales, V Cashier (Shift Sheet), W Cash Counted,
// X Safe Drops, Y Credit, Z Debit, AA Cards Total, AB EBT (Shift Sheet), AC Coupons,
// AD Fees, AE Collected, AF Short / Over, AG Expected Deposit (Shift Sheet), AH Synced At,
// AI Gas Profit 87, AJ Gas Profit 89, AK Gas Profit 93, AL Gas Profit Total, AM Comments.
// (The shift-sheet block starts at V, not W: the block was inserted before the Comments
// column, pushing Comments to AM.)
var SHIFT_LABELS = { 1: '1st Shift (Overnight)', 2: '2nd Shift (Morning)', 3: '3rd Shift (Evening)' };

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || '';
  if (action === 'reports') {
    return ContentService.createTextOutput(JSON.stringify(buildReports_()))
      .setMimeType(ContentService.MimeType.JSON);
  }
  return ContentService.createTextOutput(JSON.stringify({ ok: true, service: 'uptown-shift-sync' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// Read a tab's display values (header + rows) as plain strings, or [] if missing.
function readTab_(spreadsheetId, tabName) {
  try {
    var ss = SpreadsheetApp.openById(spreadsheetId);
    var sh = ss.getSheetByName(tabName);
    if (!sh) return [];
    return sh.getDataRange().getDisplayValues();
  } catch (err) {
    return [];
  }
}

// Serve the closing-report sheets as JSON for the website's "Daily Reports" tab.
// Read-only: never writes anything. Values are display strings (money already $).
function buildReports_() {
  return {
    ok: true,
    service: 'uptown-shift-sync',
    action: 'reports',
    generatedAt: new Date().toISOString(),
    dailyClosing: readTab_(DAILY_SHEET_ID, 'Daily Closing'),
    weeklySummary: readTab_(WEEKLY_SHEET_ID, 'Weekly Summary'),
    dailyBreakdown: readTab_(WEEKLY_SHEET_ID, 'Daily Breakdown'),
    monthlySummary: readTab_(MONTHLY_SHEET_ID, 'Monthly Summary'),
    weeklyBreakdown: readTab_(MONTHLY_SHEET_ID, 'Weekly Breakdown'),
    departmentSales: readTab_(DEPT_SHEET_ID, DEPT_SHEET_NAME),
    departmentAmounts: readTab_(DEPT_SHEET_ID, DEPT_AMOUNTS_SHEET_NAME)
  };
}

function num(v) { var n = parseFloat(String(v == null ? '' : v).replace(/[$,\s]/g, '')); return isFinite(n) ? n : 0; }
function r2(x) { return Math.round((x + 1e-9) * 100) / 100; }

function fmtDate(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(v || '');
}

// Normalise a date as SEEN in the sheet (display value) to yyyy-MM-dd.
// Handles ISO text (2000-01-01) and UK display format (01/01/2000, DD/MM/YYYY).
// This is locale-proof: it never relies on how Apps Script parses a date cell,
// which is what broke the Daily Closing match in v4 (UK locale sheet).
function normDisplayDate(v) {
  var t = String(v || '').trim();
  var m = t.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return m[1] + '-' + m[2] + '-' + m[3];
  m = t.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  return t;
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var d = JSON.parse(e.postData.contents);
    var date = String(d.date || '');
    var shift = Number(d.shift || 1);
    if (!date) throw new Error('missing date');
    if (!SHIFT_LABELS[shift]) shift = 1;
    if (d.action === 'delete') {
      // Delete a whole day (all 3 shifts) from BOTH sheets.
      var ssDel = SpreadsheetApp.openById(SYNC_SHEET_ID);
      var shDel = ssDel.getSheetByName(SYNC_SHEET_NAME) || ssDel.getSheets()[0];
      var delData = shDel.getDataRange().getValues();
      var deletedSync = 0;
      for (var di = delData.length - 1; di >= 1; di--) {
        if (fmtDate(delData[di][0]) === date) { shDel.deleteRow(di + 1); deletedSync++; }
      }
      var ssDelD = SpreadsheetApp.openById(DAILY_SHEET_ID);
      var shDelD = ssDelD.getSheetByName(DAILY_SHEET_NAME) || ssDelD.getSheets()[0];
      var delDData = shDelD.getDataRange().getDisplayValues();
      var clearedDaily = 0, deletedDailyRows = 0;
      for (var dj = delDData.length - 1; dj >= 1; dj--) {
        var a0 = String(delDData[dj][0] || '');
        if (a0.indexOf('CUSTOM') === 0 || a0.indexOf(' to ') > -1) continue;
        if (normDisplayDate(a0) !== date) continue;
        var hasZ = String(delDData[dj][4] || '').trim() !== ''; // col E Total Sales
        if (hasZ) {
          shDelD.getRange(dj + 1, 22, 1, 17).setValues([new Array(17).fill('')]); // V-AL (cashier..gas profit total), AM Comments untouched
          shDelD.getRange(dj + 1, 20, 1, 2).setValues([['Not calculable - Cash $0 in Z-Report', 0]]); // T, U
          clearedDaily++;
        } else {
          shDelD.deleteRow(dj + 1); deletedDailyRows++;
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ ok: true, action: 'delete', date: date, deletedSync: deletedSync, clearedDaily: clearedDaily, deletedDailyRows: deletedDailyRows }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    if (d.action === 'departments') {
      // Department quantities for the dashboard Profit auto-fill: one row per
      // Date+Shift — [date, shift, qty x19 in DEPT_CATS order]. Creates the
      // tab and the header row on first use. Runs as Spectrum, so the report
      // runs POST the numbers here instead of writing the sheet directly.
      // v10: an optional amounts:[19] (ex-tax $ per category, same order) is
      // upserted the same way into the 'Department Amounts' tab.
      var dq = d.quantities || [];
      var deptRow = [date, shift];
      for (var qi = 0; qi < DEPT_CATS.length; qi++) deptRow.push(num(dq[qi]));
      var ssD = SpreadsheetApp.openById(DEPT_SHEET_ID);
      var shD = ssD.getSheetByName(DEPT_SHEET_NAME);
      if (!shD) {
        shD = ssD.insertSheet(DEPT_SHEET_NAME);
        shD.getRange(1, 1, 1, 2 + DEPT_CATS.length).setValues([['Date', 'Shift'].concat(DEPT_CATS)]);
        shD.getRange(1, 1, 1, 2 + DEPT_CATS.length).setFontWeight('bold');
        shD.setFrozenRows(1);
      }
      var dData = shD.getDataRange().getValues();
      var dMatches = [];
      for (var di = 1; di < dData.length; di++) {
        if (fmtDate(dData[di][0]) === date && Number(dData[di][1]) === shift) dMatches.push(di + 1);
      }
      if (dMatches.length > 0) {
        shD.getRange(dMatches[0], 1, 1, deptRow.length).setValues([deptRow]);
        for (var dk = dMatches.length - 1; dk >= 1; dk--) shD.deleteRow(dMatches[dk]);
      } else shD.appendRow(deptRow);
      // v10: optional ex-tax $ amounts per category, same Date+Shift, into a
      // separate 'Department Amounts' tab in the same spreadsheet. Only
      // written when an amounts array is actually sent, so old quantity-only
      // POSTs keep working unchanged. This branch touches ONLY the
      // Department Sales spreadsheet — never the Sync or Daily sheets.
      var amountsWritten = false;
      if (d.amounts && d.amounts.length) {
        var am = d.amounts || [];
        var amtRow = [date, shift];
        for (var ai = 0; ai < DEPT_CATS.length; ai++) amtRow.push(num(am[ai]));
        var shA = ssD.getSheetByName(DEPT_AMOUNTS_SHEET_NAME);
        if (!shA) {
          shA = ssD.insertSheet(DEPT_AMOUNTS_SHEET_NAME);
          shA.getRange(1, 1, 1, 2 + DEPT_CATS.length).setValues([['Date', 'Shift'].concat(DEPT_CATS)]);
          shA.getRange(1, 1, 1, 2 + DEPT_CATS.length).setFontWeight('bold');
          shA.setFrozenRows(1);
        }
        var aData = shA.getDataRange().getValues();
        var aMatches = [];
        for (var aj = 1; aj < aData.length; aj++) {
          if (fmtDate(aData[aj][0]) === date && Number(aData[aj][1]) === shift) aMatches.push(aj + 1);
        }
        if (aMatches.length > 0) {
          shA.getRange(aMatches[0], 1, 1, amtRow.length).setValues([amtRow]);
          for (var ak = aMatches.length - 1; ak >= 1; ak--) shA.deleteRow(aMatches[ak]);
        } else shA.appendRow(amtRow);
        amountsWritten = true;
      }
      return ContentService.createTextOutput(JSON.stringify({ ok: true, action: 'departments', date: date, shift: shift, amountsWritten: amountsWritten }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    var cash = r2(num(d.cash));
    var drops = r2(num(d.dropsTotal));
    var credit = r2(num(d.credit)), debit = r2(num(d.debit));
    var cards = r2(credit + debit);
    var ebt = r2(num(d.ebt)), coupons = r2(num(d.coupons)), fees = r2(num(d.fees));
    var payT = r2(num(d.payOutsTotal));
    var fuel = r2(num(d.fuelSales)), merch = r2(num(d.merchSales)), tax = r2(num(d.tax));
    var lot = r2(num(d.lotterySales));
    var lotCash = r2(num(d.lotteryCash)); // Ohio lottery cash (Shift 3 only) -> Sync col AG
    var sales = r2(num(d.totalSales) || (fuel + merch + tax + lot));
    var collected = r2(num(d.collected) || (cash + credit + debit + ebt + coupons + payT));
    var so = r2(num(d.shortOver) || (collected - sales));
    var g87 = num(d.gas87gal), g89 = num(d.gas89gal), g93 = num(d.gas93gal);
    var gasGal = r2(g87 + g89 + g93);
    var gasSales = r2(num(d.gasSales));
    var gas87profit = r2(num(d.gas87profit)), gas89profit = r2(num(d.gas89profit)),
      gas93profit = r2(num(d.gas93profit)), gasProfit = r2(num(d.gasProfit));
    // Expected Deposit to Safe = Cash Counted - Pay Outs - $300 float ($150 x 2 drawers)
    var expectedDeposit = r2(cash - payT - 300);
    var syncedAt = new Date().toISOString();

    // 1) KEEP the data in the Shift Sheet Sync sheet: exactly ONE row per Date+Shift.
    var syncRow = [date, shift, String(d.cashier || ''), cash, drops, credit, debit, cards,
      ebt, coupons, fees, payT, String(d.payOutsDetail || ''), fuel, merch, tax, lot,
      sales, collected, so, g87, g89, g93, gasGal, gasSales, expectedDeposit,
      syncedAt, gas87profit, gas89profit, gas93profit, gasProfit, JSON.stringify(d), lotCash]; // 33 values, A-AG (AG = Lottery Cash)
    var ssSync = SpreadsheetApp.openById(SYNC_SHEET_ID);
    var shSync = ssSync.getSheetByName(SYNC_SHEET_NAME) || ssSync.getSheets()[0];
    var syncData = shSync.getDataRange().getValues();
    var syncMatches = [];
    for (var i = 1; i < syncData.length; i++) {
      if (fmtDate(syncData[i][0]) === date && Number(syncData[i][1]) === shift) syncMatches.push(i + 1);
    }
    if (syncMatches.length > 0) {
      shSync.getRange(syncMatches[0], 1, 1, syncRow.length).setValues([syncRow]);
      for (var k = syncMatches.length - 1; k >= 1; k--) shSync.deleteRow(syncMatches[k]);
    } else shSync.appendRow(syncRow);

    // 2) LINK it into the Daily Closing sheet (main sheet): same shift row,
    // columns V-AL, plus T Expected Deposit and U Cash Sales. No new calculations:
    // cash, card, EBT and coupons are copied straight from the website.
    var shiftCols = [String(d.cashier || ''), cash, drops, credit, debit, cards, ebt, coupons, fees, collected, so, expectedDeposit, syncedAt, gas87profit, gas89profit, gas93profit, gasProfit];
    var ssDaily = SpreadsheetApp.openById(DAILY_SHEET_ID);
    var shDaily = ssDaily.getSheetByName(DAILY_SHEET_NAME) || ssDaily.getSheets()[0];
    var dailyData = shDaily.getDataRange().getDisplayValues();
    var prefix = SHIFT_LABELS[shift];
    var dailyMatches = [];
    for (var j = 1; j < dailyData.length; j++) {
      if (normDisplayDate(dailyData[j][0]) === date && String(dailyData[j][1] || '').indexOf(prefix) === 0) dailyMatches.push(j + 1);
    }
    var matched = dailyMatches.length > 0;
    if (matched) {
      var r = dailyMatches[0];
      shDaily.getRange(r, 22, 1, shiftCols.length).setValues([shiftCols]); // V-AL: cashier..gas profit total
      shDaily.getRange(r, 20, 1, 2).setValues([[expectedDeposit, cash]]);
      for (var m = dailyMatches.length - 1; m >= 1; m--) shDaily.deleteRow(dailyMatches[m]);
    } else {
      // No Z-Report row for this shift yet: add a partial row with the shift-sheet
      // data filled in. The scheduled report run fills the Z-Report columns (C-S)
      // into THIS row instead of adding a duplicate — it must match on Date + Shift.
      var row = new Array(39).fill('');
      row[0] = "'" + date; // leading apostrophe: keep the date as text, never auto-parsed
      row[1] = prefix;
      row[18] = 300; // S Starting Float
      row[19] = expectedDeposit; // T
      row[20] = cash; // U Cash Sales
      for (var n = 0; n < shiftCols.length; n++) row[21 + n] = shiftCols[n]; // V-AL
      row[38] = 'Shift-sheet data — Z-Report columns pending; report run will fill them in.'; // AM Comments
      // Insert BEFORE the "CUSTOM PERIOD - NOT A SHIFT" label row at the bottom,
      // so that row stays the very last row of the sheet.
      var insertAt = dailyData.length + 1;
      for (var q = 1; q < dailyData.length; q++) {
        if (String(dailyData[q][0] || '').indexOf('CUSTOM PERIOD') === 0) { insertAt = q + 1; break; }
      }
      shDaily.insertRowsBefore(insertAt, 1);
      shDaily.getRange(insertAt, 1, 1, row.length).setValues([row]);
    }
    return ContentService.createTextOutput(JSON.stringify({ ok: true, date: date, shift: shift, matched: matched, expectedDeposit: expectedDeposit }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

/* ============================================

   🌍 EU6 RATE CALCULATOR — PERFORMANCE OPTIMIZED

   Performance improvements:
   - Batch all write operations into single setValues calls
   - Cache sheet references to avoid repeated lookups
   - Optimize totals row removal (single pass)
   - Batch number format operations
   - Reduce redundant reads
   - Cache header indices
   - Minimize SpreadsheetApp.flush() calls

   ============================================ */

/* ========== ON OPEN / MENU ========== */

function onOpen() {

  const ui = SpreadsheetApp.getUi();

  const menu = ui.createMenu("🌍 Planner Tools")

    .addItem("Launch Planner Sidebar", "launchPlanner")

    .addItem("Recalculate Planner", "recalculatePlanner")

    .addItem("Clear Planner Rows", "clearPlannerRows")

    .addItem("Remove Selected Row", "removeSelectedPlannerRow")

    .addItem("Create Campaign Sheet", "createCampaignSheetPrompt")

    .addSeparator()

    .addSubMenu(ui.createMenu("⚙️ Admin")

      .addItem("Unhide All Sheets", "unhideAllSheets")

      .addItem("Hide All Except EU6 Calculator", "hideAllOtherSheets")

    );

  menu.addToUi();

  ensureDisplayCurrencyDropdown();

  ensureBuyingPointDropdown();

  hideAllOtherSheets();

}

/* ========== onEdit auto-recalc for key inputs ========== */

function onEdit(e) {

  try {

    const sh = e && e.range && e.range.getSheet();

    if (!sh || sh.getName() !== 'EU6 Calculator') return;

    const a1 = e.range.getA1Notation();

    const importantCells = ['B2','D2','B3','F2','D3'];

    if (importantCells.includes(a1)) {

      // If D3 (Buying Point) changed, update F3 (Trading Deal) with percentage

      if (a1 === 'D3') {

        updateTradingDealFromBuyingPoint(sh);

        // Small delay to ensure F3 is updated before recalculating

        Utilities.sleep(100);

      }

      recalculatePlanner();

    }

  } catch (err) {

    Logger.log('onEdit error: ' + err);

  }

}

/* ========== SHEET VISIBILITY ========== */

function hideAllOtherSheets() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  ss.getSheets().forEach(sh => {

    if (sh.getName() === 'EU6 Calculator') sh.showSheet();

    else sh.hideSheet();

  });

}

function unhideAllSheets() {

  SpreadsheetApp.getActiveSpreadsheet().getSheets().forEach(sh => sh.showSheet());

  safeAlert("✅ All sheets are now visible.");

}

/* ========== SIDEBAR ========== */

function launchPlanner() {

  const html = HtmlService.createHtmlOutputFromFile('Sidebar')

    .setTitle('EU6 Planner')

    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  SpreadsheetApp.getUi().showSidebar(html);

}

/* ========== DSP FEES (from "DSP Fees") with cache ========== */

function getDSPs() {

  const cache = CacheService.getScriptCache();

  const cached = cache.get('dsps_v1');

  if (cached) return JSON.parse(cached);

  const sh = SpreadsheetApp.getActive().getSheetByName('DSP Fees');

  if (!sh) return [];

  const data = sh.getDataRange().getValues();

  if (data.length < 2) return [];

  const headers = data[0].map(h => String(h).trim().toLowerCase());

  let dspCol = headers.findIndex(h => h === 'dsp' || h.startsWith('dsp') || h.includes('dsp'));

  let feeCol = headers.findIndex(h => h === 'dsp fee' || h === 'fee' || h.includes('fee'));

  if (dspCol === -1 || feeCol === -1) return [];

  const norm = (v) => {

    if (v === '' || v === null || v === undefined) return 0;

    if (typeof v === 'number') return v > 1 ? v / 100 : v;

    let s = String(v).trim();

    if (s.endsWith('%')) s = s.slice(0, -1);

    const n = parseFloat(s);

    return isNaN(n) ? 0 : (n > 1 ? n / 100 : n);

  };

  const result = data.slice(1)

    .filter(r => String(r[dspCol]).trim())

    .map(r => ({ name: String(r[dspCol]).trim(), fee: norm(r[feeCol]) }));

  cache.put('dsps_v1', JSON.stringify(result), 600); // 10 minutes

  return result;

}

/* ========== RATES (from "Combined Channel Rates") with cache ========== */

function getRateData() {

  const cache = CacheService.getScriptCache();

  const cached = cache.get('rates_v1');

  if (cached) return JSON.parse(cached);

  const sh = SpreadsheetApp.getActive().getSheetByName('Combined Channel Rates');

  if (!sh) return [];

  const rows = sh.getDataRange().getValues();

  if (rows.length < 2) return [];

  const h = rows[0].map(x => String(x).trim().toLowerCase());

  const idx = (needle) => h.findIndex(col => col.includes(needle.toLowerCase()));

  const result = rows.slice(1).map(r => ({

    country:   String(r[idx('country')]   || '').trim(),

    channel:   String(r[idx('channel')]   || '').trim(),

    publisher: String(r[idx('publisher')] || '').trim(),

    format:    String(r[idx('format')]    || '').trim(),

    buyCpm:    Number(r[idx('buy cpm')]   || 0),

    currency:  idx('currency') !== -1 ? String(r[idx('currency')] || 'EUR').trim() : 'EUR'

  })).filter(x => x.country && x.channel && x.publisher && x.format);

  cache.put('rates_v1', JSON.stringify(result), 600); // 10 minutes

  return result;

}

/* ========== Rate index for O(1) lookups ========== */

function buildRateIndex_(rateRows) {

  const index = {};

  for (const r of rateRows) {

    const key = [r.country, r.channel, r.publisher, r.format]

      .map(s => String(s || '').toLowerCase())

      .join('||');

    index[key] = r;

  }

  return index;

}

/* ========== Remove ALL totals rows helper — OPTIMIZED ========== */

function removeAllTotalsRows_(sh) {

  const last = sh.getLastRow();

  if (last < 6) return;

  // Single read of all data rows

  const vals = sh.getRange(6, 1, last - 5, 1).getValues();

  const toDelete = [];

  for (let i = 0; i < vals.length; i++) {

    const cellValue = String(vals[i][0] || '').trim();

    // Check for totals row marker (more flexible matching)

    if (cellValue === '📊 Totals:' || cellValue.includes('Totals:') || cellValue === 'Totals') {

      toDelete.push(6 + i);

    }

  }

  // Delete from bottom to top to maintain indices

  if (toDelete.length > 0) {

    // Sort descending to delete from bottom up (prevents index shifting issues)

    toDelete.sort((a, b) => b - a);

    for (let i = 0; i < toDelete.length; i++) {

      sh.deleteRow(toDelete[i]);

    }

  }

}

/* ========== APPEND ROW (from sidebar) — optimized ========== 

   Columns A..H: Country, Channel, Publisher, Format, % Delivery, Currency, Buy CPM, DSP Fee

   Faster: do NOT look up rates here; leave F/G blank for recalc.

*/

function appendPlannerRow(row) {

  const sh = SpreadsheetApp.getActive().getSheetByName('EU6 Calculator');

  if (!sh) throw new Error("❌ EU6 Calculator not found.");

  // ensure no stale totals rows

  removeAllTotalsRows_(sh);

  const dsps = getDSPs();

  const dspMatch = dsps.find(d => d.name === row.dsp);

  const dspFee = dspMatch ? dspMatch.fee : 0;

  const pct = Number(row.weight) || 0;

  const next = sh.getLastRow() + 1;

  // Only write A–E and H now; F/G will be filled by recalculatePlanner via rate lookup

  sh.getRange(next, 1, 1, 8).setValues([[row.country, row.channel, row.publisher, row.format, pct, '', '', dspFee]]);

  sh.getRange(next, 5).setNumberFormat("0.00%");

  sh.getRange(next, 8).setNumberFormat("0.00%");

  recalculatePlanner();

}

/* ========== CLEAR PLANNER ROWS ========== */

function clearPlannerRows() {

  const sh = SpreadsheetApp.getActive().getSheetByName('EU6 Calculator');

  if (!sh) return;

  const last = sh.getLastRow();

  if (last > 5) {

    sh.getRange(6, 1, last - 5, sh.getLastColumn())

      .clearContent()

      .setBorder(false, false, false, false, false, false);

  }

  safeAlert('Planner rows cleared!');

}

/* ========== REMOVE SELECTED ROW ========== */

function removeSelectedPlannerRow() {

  const sh = SpreadsheetApp.getActive().getSheetByName('EU6 Calculator');

  if (!sh) throw new Error('No EU6 Calculator found.');

  const range = sh.getActiveRange();

  if (!range) { safeAlert('Select a cell in the row you want to remove.'); return; }

  const row = range.getRow();

  if (row < 6) { safeAlert('Please select a plan row (row 6 or below).'); return; }

  if (String(sh.getRange(row, 1).getValue()) === '📊 Totals:') {

    safeAlert('Cannot delete the totals row.'); return;

  }

  sh.deleteRow(row);

  recalculatePlanner();

  safeAlert('🗑️ Row removed and plan recalculated.');

}

/* ========== DISPLAY CURRENCY CONTROL (F2) ========== */

function ensureDisplayCurrencyDropdown() {

  const sh = SpreadsheetApp.getActive().getSheetByName('EU6 Calculator');

  if (!sh) throw new Error('Missing sheet.');

  if (!sh.getRange('E2').getValue()) sh.getRange('E2').setValue('Display Currency');

  const cell = sh.getRange('F2');

  const rule = SpreadsheetApp.newDataValidation()

    .requireValueInList(['GBP', 'EUR', 'USD'], true)

    .setAllowInvalid(false)

    .build();

  cell.setDataValidation(rule);

  if (!cell.getValue()) cell.setValue('GBP');

}

function getTargetCurrency() {

  const sh = SpreadsheetApp.getActive().getSheetByName('EU6 Calculator');

  const v = String(sh.getRange('F2').getValue() || 'GBP').trim().toUpperCase();

  return ['GBP', 'EUR', 'USD'].includes(v) ? v : 'GBP';

}

/* ========== BUYING POINT & TRADING DEAL CONTROL ========== */

function ensureBuyingPointDropdown() {

  const sh = SpreadsheetApp.getActive().getSheetByName('EU6 Calculator');

  if (!sh) throw new Error('Missing sheet.');

  // Set up labels if not already set

  if (!sh.getRange('C3').getValue()) sh.getRange('C3').setValue('Buying Point');

  if (!sh.getRange('E3').getValue()) sh.getRange('E3').setValue('Trading Deal');

  // Get buying points from Trading Deals sheet

  const tradingDealsSheet = SpreadsheetApp.getActive().getSheetByName('Trading Deals');

  if (!tradingDealsSheet) return; // Skip if sheet doesn't exist

  const data = tradingDealsSheet.getDataRange().getValues();

  if (data.length < 2) return; // Need at least headers + 1 row

  // Extract buying points from column A (skip header row)

  const buyingPoints = data.slice(1)

    .map(row => String(row[0] || '').trim())

    .filter(val => val !== '');

  if (buyingPoints.length === 0) return;

  // Set up dropdown in D3

  const cell = sh.getRange('D3');

  const rule = SpreadsheetApp.newDataValidation()

    .requireValueInList(buyingPoints, true)

    .setAllowInvalid(false)

    .build();

  cell.setDataValidation(rule);

  // If D3 has a value, update F3 with the corresponding percentage

  const buyingPoint = cell.getValue();

  if (buyingPoint) {

    updateTradingDealFromBuyingPoint(sh);

  }

  // Also set up a change handler to update when dropdown value changes

  // Note: onEdit trigger should handle this, but we'll also call it here for safety

}


function updateTradingDealFromBuyingPoint(sh) {

  try {

    if (!sh) sh = SpreadsheetApp.getActive().getSheetByName('EU6 Calculator');

    if (!sh) {

      Logger.log('updateTradingDealFromBuyingPoint: Sheet not found');

      return;

    }

    const buyingPoint = String(sh.getRange('D3').getValue() || '').trim();

    if (!buyingPoint) {

      sh.getRange('F3').setValue('');

      return;

    }

    // Look up percentage from Trading Deals sheet

    const tradingDealsSheet = SpreadsheetApp.getActive().getSheetByName('Trading Deals');

    if (!tradingDealsSheet) {

      Logger.log('updateTradingDealFromBuyingPoint: Trading Deals sheet not found');

      return;

    }

    const data = tradingDealsSheet.getDataRange().getValues();

    if (data.length < 2) {

      Logger.log('updateTradingDealFromBuyingPoint: Trading Deals sheet has no data');

      return;

    }

    // Find matching buying point (column A) and get percentage (column B)

    let percentage = null;

    for (let i = 1; i < data.length; i++) {

      const rowBuyingPoint = String(data[i][0] || '').trim();

      if (rowBuyingPoint === buyingPoint) {

        percentage = parseFloat(data[i][1]) || 0;

        // Since cells are percentage formatted, values are already decimals (0.15 for 15%)

        // No conversion needed - just ensure it's a valid decimal between 0 and 1

        if (percentage > 1) percentage = percentage / 100; // Fallback for any non-formatted cells

        Logger.log(`updateTradingDealFromBuyingPoint: Found ${buyingPoint} = ${percentage}`);

        break;

      }

    }

    // Set F3 with the percentage (as decimal, e.g., 0.15 for 15%)

    if (percentage !== null) {

      sh.getRange('F3').setValue(percentage);

      Logger.log(`updateTradingDealFromBuyingPoint: Set F3 to ${percentage}`);

    } else {

      Logger.log(`updateTradingDealFromBuyingPoint: No match found for ${buyingPoint}`);

      sh.getRange('F3').setValue(0);

    }

  } catch (e) {

    Logger.log('updateTradingDealFromBuyingPoint error: ' + e.toString());

    safeAlert('Error updating Trading Deal: ' + e.toString());

  }

}

function getTradingDealPercentage() {

  const sh = SpreadsheetApp.getActive().getSheetByName('EU6 Calculator');

  if (!sh) return 0;

  const val = parseFloat(sh.getRange('F3').getValue()) || 0;

  // If > 1, assume it's a percentage and convert to decimal

  return val > 1 ? val / 100 : val;

}


/* ========== FX (cached per day) ========== */

function getCachedExchangeRates() {

  const props = PropertiesService.getScriptProperties();

  const today = new Date().toDateString();

  const cached = JSON.parse(props.getProperty('exchangeRates') || '{}');

  if (cached.date === today && cached.rates) return cached.rates;

  const rates = {

    EURtoGBP: getGoogleRateSilent('EUR','GBP'),

    EURtoUSD: getGoogleRateSilent('EUR','USD'),

    GBPtoUSD: getGoogleRateSilent('GBP','USD')

  };

  rates.GBPtoEUR = 1 / rates.EURtoGBP;

  rates.USDtoEUR = 1 / rates.EURtoUSD;

  rates.USDtoGBP = 1 / rates.GBPtoUSD;

  props.setProperty('exchangeRates', JSON.stringify({ date: today, rates }));

  return rates;

}

function getGoogleRateSilent(from, to) {

  try {

    const s = SpreadsheetApp.getActive().getActiveSheet();

    const c = s.getRange('Z1'); const old = c.getValue();

    c.setFormula(`=GOOGLEFINANCE("CURRENCY:${from}${to}")`);

    SpreadsheetApp.flush(); Utilities.sleep(1000);

    const rate = Number(c.getValue());

    c.clearContent(); if (old) c.setValue(old);

    return rate || 1;

  } catch (e) {

    Logger.log('FX error: ' + e);

    return 1;

  }

}

/* ========== FORMATTING HELPERS ========== */

function columnLetter(n){let s='';while(n>0){let m=(n-1)%26;s=String.fromCharCode(65+m)+s;n=Math.floor((n-m)/26);}return s;}

function planFmt_(plan) { return plan === 'GBP' ? '£#,##0.00' : plan === 'USD' ? '$#,##0.00' : '€#,##0.00'; }

function nativeFmt_(curr) { return curr === 'GBP' ? '£#,##0.00' : curr === 'USD' ? '$#,##0.00' : '€#,##0.00'; }

/* ========== RECALCULATE PLANNER (all outputs) — HEAVILY OPTIMIZED ==========

Expected header row (Row 5 names):

A Country | B Channel | C Publisher | D Format | E Percentage Delivery | F Currency | G Buy CPM | H DSP Fee

I Gross Budget (In Plan Currency) | J Net Budget | K Net Budget (In Plan Currency)

L Impressions | M Publisher Spend | N DSP Fee Value | O Total Media Spend | P Gross Margin | Q Gross Profit | R Gross Margin % | S Gross Profit % | T Margin (or Margin %)

- Buying Point in D3, Trading Deal % in F3 (from Trading Deals sheet)

- Buy-side (G, M, N, O) stays in native currency (per row)

- Plan-side (I, K, P, Q) in display currency (F2)

- Gross Profit (Q) = Net - Media - Hard Costs (before trading deal)

- Gross Margin (P) = Gross Profit - (Gross Profit * Trading Deal %)

- Gross Margin % (R) = Gross Margin / Net

- Gross Profit % (S) = Gross Profit / Net

- Margin (T) = Gross Margin / Net (displayed as %), blended total = SUM(P)/SUM(K)

*/

function recalculatePlanner() {

  const sh = SpreadsheetApp.getActive().getSheetByName('EU6 Calculator');

  if (!sh) throw new Error('No EU6 Calculator found.');

  ensureDisplayCurrencyDropdown();

  const target = getTargetCurrency();

  const fx = getCachedExchangeRates();

  const tradingDealPct = getTradingDealPercentage(); // Get trading deal percentage from F3

  // Cache sheet reference and batch format operations

  sh.getRangeList(['B2','B3','D2']).setNumberFormat(planFmt_(target));

  const rateData = getRateData();

  const rateIndex = buildRateIndex_(rateData);

  const last = sh.getLastRow();

  if (last <= 5) return;

  // Remove ALL existing totals rows

  removeAllTotalsRows_(sh);

  // Read headers once and cache

  let lastCol = sh.getLastColumn();

  let headers = sh.getRange(5, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());

  const idx = (nameLike) => headers.findIndex(h => h.toLowerCase().includes(nameLike.toLowerCase())) + 1;

  // Column indices by header - cache all at once

  const cCountry = idx('country');

  const cChannel = idx('channel');

  const cPublisher = idx('publisher');

  const cFormat = idx('format');

  const cPct = headers.indexOf('Percentage Delivery') + 1 || idx('percentage');

  const cCurr = idx('currency');

  const cBuy  = idx('buy cpm');

  const cDSP  = idx('dsp fee');

  const cGross    = idx('gross budget');

  const cGrossPlan= idx('gross budget (in plan');

  const cNet      = idx('net budget');

  const cNetPlan  = idx('net budget (in plan');

  const cImps     = idx('impressions');

  const cPub      = idx('publisher spend');

  const cDSPV     = idx('dsp fee value');

  const cMedia    = idx('total media spend');

  const cHard     = idx('allocated hard');

  const cProfit   = idx('profit');

  // Create missing columns if they don't exist (add at end in correct order)

  let headersUpdated = false;

  // Order: Gross Margin > Margin > Gross Profit > Profit %

  // 1. Gross Margin column (first)

  let cGrossMargin = idx('gross margin');

  if (cGrossMargin <= 0) { 

    lastCol = sh.getLastColumn();

    cGrossMargin = lastCol + 1; 

    sh.getRange(5, cGrossMargin).setValue('Gross Margin'); 

    lastCol = cGrossMargin; // Update for later use

    headersUpdated = true;

  }

  // 2. Margin column (second)

  let cMargin = headers.indexOf('Margin') + 1;

  if (cMargin <= 0) cMargin = headers.indexOf('Margin %') + 1;

  if (cMargin <= 0) { 

    lastCol = sh.getLastColumn();

    cMargin = lastCol + 1; 

    sh.getRange(5, cMargin).setValue('Margin'); 

    lastCol = cMargin; // Update for later use

    headersUpdated = true;

  }

  // 3. Gross Profit column (third)

  let cGrossProfit = idx('gross profit');

  if (cGrossProfit <= 0) { 

    lastCol = sh.getLastColumn();

    cGrossProfit = lastCol + 1; 

    sh.getRange(5, cGrossProfit).setValue('Gross Profit'); 

    lastCol = cGrossProfit; // Update for later use

    headersUpdated = true;

  }

  // 4. Profit % column (fourth)

  let cProfitPct = idx('profit %');

  if (cProfitPct <= 0) {

    // Try to find it with different variations

    cProfitPct = headers.findIndex(h => h.toLowerCase().includes('profit') && h.toLowerCase().includes('%')) + 1;

    if (cProfitPct <= 0) { 

      lastCol = sh.getLastColumn();

      cProfitPct = lastCol + 1; 

      sh.getRange(5, cProfitPct).setValue('Profit %'); 

      lastCol = cProfitPct; // Update for later use

      headersUpdated = true;

    }

  }

  // Re-read headers if we added any columns

  if (headersUpdated) {

    lastCol = sh.getLastColumn();

    headers = sh.getRange(5, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());

    // Re-find column indices after headers update

    cGrossMargin = idx('gross margin');

    cMargin = headers.indexOf('Margin') + 1;

    if (cMargin <= 0) cMargin = headers.indexOf('Margin %') + 1;

    cGrossProfit = idx('gross profit');

    cProfitPct = idx('profit %');

    if (cProfitPct <= 0) {

      cProfitPct = headers.findIndex(h => h.toLowerCase().includes('profit') && h.toLowerCase().includes('%')) + 1;

    }

  }

  // Read all input values in one batch

  const totalBudget = parseFloat(sh.getRange('B2').getValue());

  const hardTotalPlan = parseFloat(sh.getRange('B3').getValue()) || 0;

  const sellCpm     = parseFloat(sh.getRange('D2').getValue());

  if (isNaN(totalBudget) || isNaN(sellCpm)) { 

    safeAlert('Please set Total Budget (B2) and Sell CPM (D2).'); 

    return; 

  }

  const rowsCount = sh.getLastRow() - 5;

  if (rowsCount <= 0) return;

  // Single read of all data rows

  const base = sh.getRange(6, 1, rowsCount, lastCol).getValues();

  // First pass: compute native & plan numbers

  const lines = base.map(row => {

    const country = String(row[cCountry - 1] || '');

    const channel = String(row[cChannel - 1] || '');

    const publisher = String(row[cPublisher - 1] || '');

    const format = String(row[cFormat - 1] || '');

    const pct = parseFloat(row[cPct - 1]) || 0;

    // currency & buy (lookup refresh)

    let curr = String(row[cCurr - 1] || '').trim() || 'EUR';

    let buy  = parseFloat(row[cBuy - 1]) || 0;

    const r = rateIndex[[country, channel, publisher, format].map(s => s.toLowerCase()).join('||')];

    if (r) { curr = r.currency || curr; buy = r.buyCpm || buy; }

    // dsp

    let dsp = parseFloat(row[cDSP - 1]) || 0; if (dsp > 1) dsp = dsp / 100;

    // core calcs (native)

    const gross = totalBudget * pct;

    const net   = gross * 0.85;

    const imps  = (gross / sellCpm) * 1000;

    const pub   = (imps / 1000) * buy;

    const dspv  = (buy * dsp) * (imps / 1000);

    const media = pub + dspv;

    // helper fx

    const conv = (v, from, to) => {

      const k = `${from}to${to}`;

      return fx[k] ? v * fx[k] : v;

    };

    // plan conversions

    const grossPlan = curr === target ? gross : conv(gross, curr, target);

    const netPlan   = curr === target ? net   : conv(net,   curr, target);

    return { country, channel, publisher, format, pct, curr, buy, dsp,

             gross, net, imps, pub, dspv, media, grossPlan, netPlan };

  });

  // Hard costs: B3 is in plan/display currency (F2).

  // Allocate by each row's share of total native net, store both plan and native.

  const totalNet = lines.reduce((s, r) => s + r.net, 0);

  lines.forEach(r => {

    const share = totalNet ? (r.net / totalNet) : 0;

    // Plan-currency hard cost allocation (for column Q)

    const hcPlan = share * hardTotalPlan;

    r.hardPlan = hcPlan;

    // Convert plan hard cost to row's native for profit math

    const toNative = (v, planCurr, rowCurr) => {

      if (planCurr === rowCurr) return v;

      const k = `${planCurr}to${rowCurr}`;

      return fx[k] ? v * fx[k] : v;

    };

    r.hardNative = toNative(hcPlan, target, r.curr);

    // Gross Margin in native (profit BEFORE trading deal - higher value)

    const grossMarginNative = r.net - r.media - r.hardNative;

    const toPlan = (v, fromCurr, planCurr) => {

      if (fromCurr === planCurr) return v;

      const k = `${fromCurr}to${planCurr}`;

      return fx[k] ? v * fx[k] : v;

    };

    // Convert Gross Margin to plan currency

    r.grossMarginPlan = toPlan(grossMarginNative, r.curr, target);

    // Trading Deal Amount (from Gross Margin)

    const tradingDealNative = grossMarginNative * tradingDealPct;

    // Gross Profit (Gross Margin - Trading Deal - AFTER trading deal, lower value)

    const grossProfitNative = grossMarginNative - tradingDealNative;

    r.grossProfitPlan = toPlan(grossProfitNative, r.curr, target);

    // Margin (before rebates) = Gross Margin / Net

    r.margin = r.net ? (grossMarginNative / r.net) : 0;

    // Profit % (after rebates) = Gross Profit / Net (final margin with trading deal taken out)

    r.profitPct = r.net ? (grossProfitNative / r.net) : 0;

    // Keep profitPlan for backward compatibility (now represents Gross Profit after trading deal)

    r.profitPlan = r.grossProfitPlan;

  });

  // OPTIMIZATION: Build complete output matrix in one pass

  const startRow = 6;

  const numRows = lines.length;

  const maxCol = Math.max(cCurr, cBuy, cGross, cGrossPlan, cNet, cNetPlan, cImps, cPub, cDSPV, cMedia, cHard, cGrossProfit, cGrossMargin, cProfit, cProfitPct, cMargin, lastCol);

  // Initialize output matrix with existing values

  const outputMatrix = [];

  for (let i = 0; i < numRows; i++) {

    const row = new Array(maxCol).fill(null);

    // Preserve existing values for columns we're not updating

    for (let j = 0; j < Math.min(base[i].length, maxCol); j++) {

      row[j] = base[i][j];

    }

    const line = lines[i];

    // Update computed columns

    if (cCurr > 0) row[cCurr - 1] = line.curr;

    if (cBuy > 0) row[cBuy - 1] = line.buy;

    if (cGross > 0) row[cGross - 1] = line.gross;

    if (cGrossPlan > 0) row[cGrossPlan - 1] = line.grossPlan;

    if (cNet > 0) row[cNet - 1] = line.net;

    if (cNetPlan > 0) row[cNetPlan - 1] = line.netPlan;

    if (cImps > 0) row[cImps - 1] = line.imps;

    if (cPub > 0) row[cPub - 1] = line.pub;

    if (cDSPV > 0) row[cDSPV - 1] = line.dspv;

    if (cMedia > 0) row[cMedia - 1] = line.media;

    if (cHard > 0) row[cHard - 1] = line.hardPlan;

    // Write columns in order: Gross Margin > Margin > Gross Profit > Profit %

    // 1. Gross Margin (profit BEFORE trading deal - higher value)

    if (cGrossMargin > 0) row[cGrossMargin - 1] = line.grossMarginPlan;

    // 2. Margin (before rebates) = Gross Margin / Net

    if (cMargin > 0) row[cMargin - 1] = line.margin;

    // 3. Gross Profit (profit AFTER trading deal - lower value)

    if (cGrossProfit > 0) row[cGrossProfit - 1] = line.grossProfitPlan;

    // 4. Profit % (after rebates - final margin)

    if (cProfitPct > 0) row[cProfitPct - 1] = line.profitPct;

    // Keep profit column for backward compatibility (if exists)

    if (cProfit > 0) row[cProfit - 1] = line.profitPlan;

    outputMatrix.push(row);

  }

  // SINGLE BATCH WRITE - much faster than multiple setValues calls

  sh.getRange(startRow, 1, numRows, maxCol).setValues(outputMatrix);

  // Batch number format operations

  const planFmt = planFmt_(target);

  // Build format matrix

  const formatMatrix = [];

  for (let i = 0; i < numRows; i++) {

    const rowFormats = new Array(maxCol).fill('');

    const line = lines[i];

    const nativeFmt = nativeFmt_(line.curr || 'EUR');

    // Apply formats

    if (cPct > 0) rowFormats[cPct - 1] = '0.00%';

    if (cDSP > 0) rowFormats[cDSP - 1] = '0.00%';

    if (cBuy > 0) rowFormats[cBuy - 1] = nativeFmt;

    if (cGrossPlan > 0) rowFormats[cGrossPlan - 1] = planFmt;

    if (cNetPlan > 0) rowFormats[cNetPlan - 1] = planFmt;

    if (cImps > 0) rowFormats[cImps - 1] = '#,##0';

    if (cPub > 0) rowFormats[cPub - 1] = nativeFmt;

    if (cDSPV > 0) rowFormats[cDSPV - 1] = nativeFmt;

    if (cMedia > 0) rowFormats[cMedia - 1] = nativeFmt;

    if (cHard > 0) rowFormats[cHard - 1] = planFmt;

    // Format columns in order: Gross Margin > Margin > Gross Profit > Profit %

    if (cGrossMargin > 0) rowFormats[cGrossMargin - 1] = planFmt;

    if (cMargin > 0) rowFormats[cMargin - 1] = '0.0%';

    if (cGrossProfit > 0) rowFormats[cGrossProfit - 1] = planFmt;

    if (cProfitPct > 0) rowFormats[cProfitPct - 1] = '0.0%';

    if (cProfit > 0) rowFormats[cProfit - 1] = planFmt;

    formatMatrix.push(rowFormats);

  }

  // Single batch format operation

  sh.getRange(startRow, 1, numRows, maxCol).setNumberFormats(formatMatrix);

  // Build Totals row (sum key numeric columns) + blended margin

  const endRow = startRow + numRows - 1;

  const nCols  = Math.max(maxCol, lastCol);

  const totals = new Array(nCols).fill('');

  totals[0] = '📊 Totals:';

  const sumCol = (cIdx) => {

    if (cIdx > 0) {

      const L = columnLetter(cIdx);

      totals[cIdx - 1] = `=IFERROR(SUM(${L}${startRow}:${L}${endRow}),"")`;

    }

  };

  if (cPct > 0) sumCol(cPct);

  [cGrossPlan, cNetPlan, cImps, cPub, cDSPV, cMedia, cHard, cGrossProfit, cGrossMargin, cProfit].forEach(sumCol);

  // Margin total (before rebates) = Gross Margin / Net

  if (cMargin > 0 && cGrossMargin > 0 && cNetPlan > 0) {

    const LM = columnLetter(cGrossMargin);

    const LN = columnLetter(cNetPlan);

    totals[cMargin - 1] = `=IF(SUM(${LN}${startRow}:${LN}${endRow})=0,0,SUM(${LM}${startRow}:${LM}${endRow})/SUM(${LN}${startRow}:${LN}${endRow}))`;

  } else if (cMargin > 0 && cProfit > 0 && cNetPlan > 0) {

    // Fallback to old calculation if Gross Margin column doesn't exist

    const LP = columnLetter(cProfit);

    const LN = columnLetter(cNetPlan);

    totals[cMargin - 1] = `=IF(SUM(${LN}${startRow}:${LN}${endRow})=0,0,SUM(${LP}${startRow}:${LP}${endRow})/SUM(${LN}${startRow}:${LN}${endRow}))`;

  }

  // Profit % total (after rebates) = Gross Profit / Net

  if (cProfitPct > 0 && cGrossProfit > 0 && cNetPlan > 0) {

    const LP = columnLetter(cGrossProfit);

    const LN = columnLetter(cNetPlan);

    totals[cProfitPct - 1] = `=IF(SUM(${LN}${startRow}:${LN}${endRow})=0,0,SUM(${LP}${startRow}:${LP}${endRow})/SUM(${LN}${startRow}:${LN}${endRow}))`;

  }

  // Write totals row

  sh.getRange(endRow + 1, 1, 1, nCols).setValues([totals]);

  // Format totals cells in batch

  const totalsFormats = new Array(nCols).fill('');

  if (cGrossPlan > 0) totalsFormats[cGrossPlan - 1] = planFmt;

  if (cNetPlan > 0) totalsFormats[cNetPlan - 1] = planFmt;

  // Format totals in order: Gross Margin > Margin > Gross Profit > Profit %

  if (cGrossMargin > 0) totalsFormats[cGrossMargin - 1] = planFmt;

  if (cMargin > 0) totalsFormats[cMargin - 1] = '0.0%';

  if (cGrossProfit > 0) totalsFormats[cGrossProfit - 1] = planFmt;

  if (cProfitPct > 0) totalsFormats[cProfitPct - 1] = '0.0%';

  if (cProfit > 0) totalsFormats[cProfit - 1] = planFmt;

  sh.getRange(endRow + 1, 1, 1, nCols).setNumberFormats([totalsFormats]);

  // Style totals row (optimized - direct reference)

  const totalsRowNum = endRow + 1;

  const totalsRange = sh.getRange(totalsRowNum, 1, 1, nCols);

  totalsRange.setFontWeight('bold');

  totalsRange.setBorder(true, null, true, null, null, null, 'white', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // FX label

  const r = getCachedExchangeRates();

  sh.getRange('G2').setValue(

    `📊 Rates today: 1 EUR=${r.EURtoGBP.toFixed(2)} GBP | 1 EUR=${r.EURtoUSD.toFixed(2)} USD | 1 GBP=${r.GBPtoUSD.toFixed(2)} USD`

  );

  safeAlert('✅ Recalculated.');

}

/* ========== DRIVE COPY: CREATE CAMPAIGN SHEET ========== */

function createCampaignSheetPrompt() {

  const ui = SpreadsheetApp.getUi();

  const response = ui.prompt('Create Campaign Sheet', 'Enter the Deal ID for this campaign:', ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() !== ui.Button.OK) return;

  const dealId = response.getResponseText().trim();

  if (!dealId) return ui.alert('Please enter a valid Deal ID.');

  createCampaignSheet(dealId);

}

function createCampaignSheet(dealId) {

  const ui = SpreadsheetApp.getUi();

  const sourceFile = SpreadsheetApp.getActiveSpreadsheet();

  const sourceFileId = sourceFile.getId();

  const campaignInfoFolderId = "1bB4gMH2864mR0bZ53mVKeTm34019aKe2"; // Shared drive ID (update if needed)

  const campaignInfoFolder = DriveApp.getFolderById(campaignInfoFolderId);

  const advertiserFolders = campaignInfoFolder.getFolders();

  let targetFolder = null, scanned = [];

  while (advertiserFolders.hasNext()) {

    const advertiserFolder = advertiserFolders.next();

    const dealFolders = advertiserFolder.getFolders();

    while (dealFolders.hasNext()) {

      const dealFolder = dealFolders.next();

      const folderName = dealFolder.getName();

      scanned.push(folderName);

      if (folderName.includes(dealId)) {

        targetFolder = dealFolder;

        break;

      }

    }

    if (targetFolder) break;

  }

  if (!targetFolder) {

    ui.alert(`❌ No folder found with Deal ID: "${dealId}" inside Campaign Info.\n\nScanned:\n${scanned.join('\n')}`);

    return;

  }

  const targetFolderName = targetFolder.getName();

  const newFileName = `${targetFolderName} - Delivery Schedule`;

  const copiedFile = DriveApp.getFileById(sourceFileId).makeCopy(newFileName, targetFolder);

  ui.alert('✅ Campaign sheet created:\n' + copiedFile.getUrl());

}

/* ========== FX PREWARM TRIGGERS (optional) ========== */

// Run manually once to create a daily trigger that pre-warms FX cache

function installFxPrewarmTrigger() {

  const triggers = ScriptApp.getProjectTriggers();

  const exists = triggers.some(t => t.getHandlerFunction && t.getHandlerFunction() === 'prewarmFxCache');

  if (!exists) {

    ScriptApp.newTrigger('prewarmFxCache')

      .timeBased()

      .everyDays(1)

      .atHour(6) // spreadsheet timezone

      .create();

  }

}

function removeFxPrewarmTrigger() {

  const triggers = ScriptApp.getProjectTriggers();

  triggers.forEach(t => {

    if (t.getHandlerFunction && t.getHandlerFunction() === 'prewarmFxCache') {

      ScriptApp.deleteTrigger(t);

    }

  });

}

function prewarmFxCache() {

  try {

    getCachedExchangeRates();

    Logger.log('FX cache pre-warmed for ' + new Date().toDateString());

  } catch (e) {

    Logger.log('prewarmFxCache error: ' + e);

  }

}

/* ========== GET TOTAL PERCENTAGE (for sidebar hint) ========== */

function getTotalPercentage() {

  try {

    const sh = SpreadsheetApp.getActive().getSheetByName('EU6 Calculator');

    if (!sh) return 0;

    const last = sh.getLastRow();

    if (last <= 5) return 0;

    const headers = sh.getRange(5, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h).trim());

    const cPct = headers.indexOf('Percentage Delivery') + 1 || headers.findIndex(h => h.toLowerCase().includes('percentage')) + 1;

    if (cPct <= 0) return 0;

    const rowsCount = last - 5;

    const values = sh.getRange(6, cPct, rowsCount, 1).getValues();

    const firstColValues = sh.getRange(6, 1, rowsCount, 1).getValues(); // To check for totals rows

    const total = values.reduce((sum, row, idx) => {

      // Skip totals rows (marked with 📊 Totals:)

      if (String(firstColValues[idx][0]) === '📊 Totals:') return sum;

      const val = parseFloat(row[0]) || 0;

      return sum + val;

    }, 0);

    return total;

  } catch (e) {

    Logger.log('getTotalPercentage error: ' + e);

    return 0;

  }

}

/* ========== SAFE ALERT WRAPPER ========== */

function safeAlert(msg){try{SpreadsheetApp.getUi().alert(msg);}catch(e){Logger.log(msg);}}

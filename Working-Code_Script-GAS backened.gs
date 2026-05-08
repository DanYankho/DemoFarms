/**
 * ═══════════════════════════════════════════════════════════════════════════
 *  ELIM GROUP LIMITED — Elim Farms Monthly Operational Report Generator
 *  Google Apps Script  |  Version 2.1  |  Production Year 2026
 * ═══════════════════════════════════════════════════════════════════════════
 *
 *  HOW TO DEPLOY:
 *  1. Open script.google.com → New Project
 *  2. Paste this entire file
 *  3. Fill in CONFIGURATION below
 *  4. Run setupTrigger() ONCE to install the monthly auto-trigger
 *  5. Run generateAllPastReports() ONCE to backfill any missing months
 *
 *  GOOGLE SHEET LAYOUT (Tab: "PY26 Data"):
 *  Row 1  = Headers
 *  Row 2  = Jan  ... Row 13 = Dec  ... Row 14 = Total
 *  Columns match the CSV structure:
 *    A=Month, B=Planned Bunches, C=Actual Bunches, D=Variance,
 *    E=Planned Avg Weight, F=Actual Avg Weight,
 *    G=Planned Bearing Rate, H=Actual Bearing Rate,
 *    I=Plant Population, J=Planned Total Volume, K=Actual Total Volume,
 *    L=Planned Avg Selling Price, M=Actual Avg Selling Price,
 *    N=Planned Revenue, O=Actual Revenue,
 *    P=Planned Monthly Costs, Q=Actual Monthly Costs,
 *    R=Planned Margin, S=Planned Margin%, T=Actual Margin, U=Actual Margin%
 * ═══════════════════════════════════════════════════════════════════════════
 *
 *  CHANGELOG v2.1:
 *  - FIX: Margin % now displays as e.g. "6.9%" instead of raw float 0.06874...
 *  - FIX: Bearing rate now displays as e.g. "3.4%" instead of raw float 0.03445...
 *  - FIX: All narrative text uses formatted % values (executive summary, notes)
 *  - IMPROVEMENT: All numbers use comma-separated thousands (1,234,567)
 *  - IMPROVEMENT: Added formatDecimalPct_() and formatBearingRate_() helpers
 * ═══════════════════════════════════════════════════════════════════════════
 */

// ─── CONFIGURATION ────────────────────────────────────────────────────────────
const CONFIG = {
  // Google Sheet ID (from the URL: docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit)
  SPREADSHEET_ID: '1Ru3j74JStFbaPuyt2MrI33Z9WevxLsFunzXAAwcgrFI',

  // Tab name inside the spreadsheet that holds the PY26 data
  SHEET_TAB_NAME: 'Data for Equip Dashboard',

  // Google Docs Template ID (upload the .docx to Drive → open → note the ID from URL)
  TEMPLATE_DOC_ID: '1Ax9t4sPejGTOGMKvbbus4eLvLqp-oavNsI9aWLmvbLk',

  // Google Drive Folder ID where generated reports should be saved
  OUTPUT_FOLDER_ID: '1DfrF_XqfTDdfgVfcPgEBT4y2OAoAWCV6',

  // Email recipients — up to 5 (all will receive each monthly report)
  EMAIL_RECIPIENTS: [
    'dannyaltothefirst@gmail.com',
    'danford.mponda@equipgroup.co'
  ],

  // Company identity
  COMPANY_NAME: 'Elim Group Limited',
  FARM_NAME: 'Elim Farms',
  REPORT_YEAR: 2026,

  // Set to true to send emails; false for Drive-only (useful during testing)
  SEND_EMAILS: true
};

// ─── MONTH CONSTANTS ──────────────────────────────────────────────────────────
const MONTHS = [
  { short: 'Jan', full: 'January',   row: 1,  num: 1  },
  { short: 'Feb', full: 'February',  row: 2,  num: 2  },
  { short: 'Mar', full: 'March',     row: 3,  num: 3  },
  { short: 'Apr', full: 'April',     row: 4,  num: 4  },
  { short: 'May', full: 'May',       row: 5,  num: 5  },
  { short: 'Jun', full: 'June',      row: 6,  num: 6  },
  { short: 'Jul', full: 'July',      row: 7,  num: 7  },
  { short: 'Aug', full: 'August',    row: 8,  num: 8  },
  { short: 'Sep', full: 'September', row: 9,  num: 9  },
  { short: 'Oct', full: 'October',   row: 10, num: 10 },
  { short: 'Nov', full: 'November',  row: 11, num: 11 },
  { short: 'Dec', full: 'December',  row: 12, num: 12 }
];

// ─── INSTALL TRIGGER (run once) ───────────────────────────────────────────────
/**
 * Sets up a monthly time-driven trigger.
 * Run this function ONCE after deployment.
 * It fires on the 10th of every month at 08:00 to generate the previous month's report.
 */
function setupTrigger() {
  // Remove any existing triggers for this function to avoid duplicates
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'runMonthlyReport') {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('runMonthlyReport')
    .timeBased()
    .onMonthDay(10)
    .atHour(8)
    .create();

  Logger.log('✅ Monthly trigger installed. Reports will auto-generate on the 10th of each month at 08:00.');
}

// ─── ENTRY POINT: MONTHLY AUTO-RUN ────────────────────────────────────────────
/**
 * Called automatically on the 10th of each month.
 * Generates the report for the PREVIOUS month.
 */
function runMonthlyReport() {
  const now = new Date();
  // Go back 1 month (handles Jan → Dec of prev year automatically)
  const reportDate = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  const monthNum = reportDate.getMonth() + 1; // 1-based
  const year = reportDate.getFullYear();

  if (year !== CONFIG.REPORT_YEAR) {
    Logger.log(`Skipping: report year ${year} does not match configured year ${CONFIG.REPORT_YEAR}`);
    return;
  }

  const month = MONTHS.find(m => m.num === monthNum);
  if (!month) {
    Logger.log('Could not resolve month. Aborting.');
    return;
  }

  Logger.log(`Running report for: ${month.full} ${year}`);

  if (!reportAlreadyExists_(month)) {
    generateReport_(month);
  } else {
    Logger.log(`Report for ${month.full} already exists — skipping.`);
  }
}

// ─── BACKFILL: GENERATE ALL PAST MISSING REPORTS ─────────────────────────────
/**
 * Run this ONCE after deployment to generate all missing reports
 * for every completed month in PY2026 up to today.
 */
function generateAllPastReports() {
  const today = new Date();
  const currentYear = today.getFullYear();
  const currentMonthNum = today.getMonth() + 1; // 1-based

  let generated = 0;
  let skipped = 0;

  MONTHS.forEach(month => {
    // Only process months in 2026 that have already passed
    const isCompleted = (currentYear > CONFIG.REPORT_YEAR) ||
                        (currentYear === CONFIG.REPORT_YEAR && month.num < currentMonthNum);

    if (!isCompleted) {
      Logger.log(`⏭  ${month.full} ${CONFIG.REPORT_YEAR}: not yet completed — skipping`);
      return;
    }

    if (reportAlreadyExists_(month)) {
      Logger.log(`✅ ${month.full} ${CONFIG.REPORT_YEAR}: report already exists — skipping`);
      skipped++;
      return;
    }

    Logger.log(`📄 Generating: ${month.full} ${CONFIG.REPORT_YEAR}...`);
    try {
      generateReport_(month);
      generated++;
      Utilities.sleep(2000); // Avoid hitting API rate limits
    } catch (e) {
      Logger.log(`❌ Error generating ${month.full}: ${e.message}`);
    }
  });

  Logger.log(`\nDone. Generated: ${generated}, Skipped (already exist): ${skipped}`);
}

// ─── CORE: GENERATE A SINGLE MONTH REPORT ────────────────────────────────────
function generateReport_(month) {
  const sheet = getDataSheet_();
  const data = getMonthData_(sheet, month);
  const ytdData = getYtdData_(sheet, month);

  const placeholders = buildPlaceholders_(month, data, ytdData);

  // 1. Copy the template
  const templateFile = DriveApp.getFileById(CONFIG.TEMPLATE_DOC_ID);
  const reportName = `Elim Farms Monthly Report — ${month.full} ${CONFIG.REPORT_YEAR}`;
  const outputFolder = DriveApp.getFolderById(CONFIG.OUTPUT_FOLDER_ID);
  const copy = templateFile.makeCopy(reportName, outputFolder);
  const doc = DocumentApp.openById(copy.getId());

  // 2. Replace all placeholders
  replacePlaceholders_(doc, placeholders);

  // 3. Save and close
  doc.saveAndClose();

  // 4. Convert to PDF
  const pdfBlob = DriveApp.getFileById(copy.getId()).getAs('application/pdf');
  pdfBlob.setName(`${reportName}.pdf`);
  const pdfFile = outputFolder.createFile(pdfBlob);

  Logger.log(`✅ PDF saved to Drive: ${pdfFile.getUrl()}`);

  // 5. Send emails
  if (CONFIG.SEND_EMAILS) {
    sendReportEmails_(month, data, pdfBlob, pdfFile.getUrl());
  }

  // 6. Remove the intermediate Google Doc copy (keep only PDF)
  copy.setTrashed(true);

  Logger.log(`✅ Report complete: ${month.full} ${CONFIG.REPORT_YEAR}`);
  return pdfFile;
}

// ─── DATA: FETCH MONTH ROW ────────────────────────────────────────────────────
function getDataSheet_() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_TAB_NAME);
  if (!sheet) throw new Error(`Sheet "${CONFIG.SHEET_TAB_NAME}" not found in spreadsheet.`);
  return sheet;
}

/**
 * Reads a single month's row from the Google Sheet.
 * Returns an object with all column values.
 * Row 2 = Jan, Row 3 = Feb, etc. (Row 1 = headers)
 *
 * FIX v2.1: Bearing rates (cols G, H) and margin % (cols S, U) are decimal
 * fractions in the sheet (e.g. 0.0687). We now convert them to display-ready
 * percentage strings (e.g. "6.9%") via formatDecimalPct_() and
 * formatBearingRate_() so they never appear as raw floats in the report.
 */
function getMonthData_(sheet, month) {
  const rowIndex = month.row + 1; // +1 because row 1 is headers
  const range = sheet.getRange(rowIndex, 1, 1, 21); // columns A–U
  const vals = range.getValues()[0];

  return {
    month:              vals[0]  || month.short,
    plannedBunches:     parseNum_(vals[1]),
    actualBunches:      parseNum_(vals[2]),
    variance:           parseNum_(vals[3]),
    plannedAvgWeight:   parseNum_(vals[4]),
    actualAvgWeight:    parseNum_(vals[5]),
    // ── FIXED: bearing rates formatted as "X.XX%" instead of raw float ──────
    plannedBearingRate: formatBearingRate_(vals[6]),
    actualBearingRate:  formatBearingRate_(vals[7]),
    // ────────────────────────────────────────────────────────────────────────
    plantPopulation:    vals[8]  || '—',
    plannedVolume:      parseNum_(vals[9]),
    actualVolume:       parseNum_(vals[10]),
    plannedPrice:       parseNum_(vals[11]),
    actualPrice:        parseNum_(vals[12]),
    plannedRevenue:     parseNum_(vals[13]),
    actualRevenue:      parseNum_(vals[14]),
    plannedCosts:       parseNum_(vals[15]),
    actualCosts:        parseNum_(vals[16]),
    plannedMargin:      parseNum_(vals[17]),
    // ── FIXED: margin % formatted as "X.X%" instead of raw float ────────────
    plannedMarginPct:   formatDecimalPct_(vals[18]),
    actualMargin:       parseNum_(vals[19]),
    actualMarginPct:    formatDecimalPct_(vals[20]),
    // ────────────────────────────────────────────────────────────────────────
    hasActuals:         (parseNum_(vals[2]) > 0) // true if actual data exists
  };
}

/**
 * Aggregates YTD totals from Jan up to (and including) the given month.
 */
function getYtdData_(sheet, month) {
  const ytd = {
    plannedBunches: 0, actualBunches: 0,
    plannedVolume: 0,  actualVolume: 0,
    plannedRevenue: 0, actualRevenue: 0,
    plannedCosts: 0,   actualCosts: 0,
    plannedMargin: 0,  actualMargin: 0
  };

  for (let m = 1; m <= month.num; m++) {
    const mo = MONTHS.find(x => x.num === m);
    const rowIdx = mo.row + 1;
    const v = sheet.getRange(rowIdx, 1, 1, 21).getValues()[0];
    ytd.plannedBunches  += parseNum_(v[1]);
    ytd.actualBunches   += parseNum_(v[2]);
    ytd.plannedVolume   += parseNum_(v[9]);
    ytd.actualVolume    += parseNum_(v[10]);
    ytd.plannedRevenue  += parseNum_(v[13]);
    ytd.actualRevenue   += parseNum_(v[14]);
    ytd.plannedCosts    += parseNum_(v[15]);
    ytd.actualCosts     += parseNum_(v[16]);
    ytd.plannedMargin   += parseNum_(v[17]);
    ytd.actualMargin    += parseNum_(v[19]);
  }

  // Derived margin % — pct_() takes a whole-number percentage (e.g. 25.2),
  // so we multiply ratio * 100 first. This was already correct in v2.0.
  ytd.plannedMarginPct = ytd.plannedRevenue > 0
    ? pct_((ytd.plannedMargin / ytd.plannedRevenue) * 100)
    : '—';
  ytd.actualMarginPct = ytd.actualRevenue > 0
    ? pct_((ytd.actualMargin / ytd.actualRevenue) * 100)
    : '—';

  return ytd;
}

// ─── BUILD PLACEHOLDER MAP ────────────────────────────────────────────────────
function buildPlaceholders_(month, d, ytd) {
  const today = new Date();
  const reportDate = Utilities.formatDate(today, Session.getScriptTimeZone(), 'dd MMMM yyyy');
  const reportStatus = d.hasActuals ? 'ACTUALS AVAILABLE' : 'PLANNED ONLY';

  // Variances
  const bunchesVar   = d.actualBunches   > 0 ? d.actualBunches   - d.plannedBunches   : null;
  const weightVar    = d.actualAvgWeight > 0 ? d.actualAvgWeight - d.plannedAvgWeight : null;
  const volumeVar    = d.actualVolume    > 0 ? d.actualVolume    - d.plannedVolume    : null;
  const priceVar     = d.actualPrice     > 0 ? d.actualPrice     - d.plannedPrice     : null;
  const revenueVar   = d.actualRevenue   > 0 ? d.actualRevenue   - d.plannedRevenue   : null;
  const costsVar     = d.actualCosts     > 0 ? d.actualCosts     - d.plannedCosts     : null;
  const marginVar    = d.actualMargin    > 0 ? d.actualMargin    - d.plannedMargin    : null;

  const varOrDash = (v, formatter) => v !== null ? formatter(v) : '—';
  const varPct    = (v, base) => (v !== null && base > 0) ? pct_((v / base) * 100) : '—';

  // Cost efficiency — fmt_() uses toLocaleString so commas are included
  const plannedCostPerKg   = d.plannedVolume   > 0 ? fmt_(d.plannedCosts   / d.plannedVolume)   : '—';
  const actualCostPerKg    = d.actualVolume    > 0 ? fmt_(d.actualCosts    / d.actualVolume)    : '—';
  const plannedRevPerKg    = d.plannedVolume   > 0 ? fmt_(d.plannedRevenue / d.plannedVolume)   : '—';
  const actualRevPerKg     = d.actualVolume    > 0 ? fmt_(d.actualRevenue  / d.actualVolume)    : '—';
  const plannedMarginBunch = d.plannedBunches  > 0 ? fmt_(d.plannedMargin  / d.plannedBunches)  : '—';
  const actualMarginBunch  = d.actualBunches   > 0 ? fmt_(d.actualMargin   / d.actualBunches)   : '—';

  // Next month targets
  const nextMonthIdx = month.num < 12 ? month.num : null; // 0-based index into MONTHS for next month
  const nextMonth = nextMonthIdx !== null ? MONTHS[nextMonthIdx] : null;
  let nextPlannedBunches = '—', nextPlannedVolume = '—', nextPlannedRevenue = '—';
  let nextPlannedCosts = '—', nextPlannedMargin = '—';
  if (nextMonth) {
    try {
      const sheet = getDataSheet_();
      const nv = sheet.getRange(nextMonth.row + 1, 1, 1, 21).getValues()[0];
      nextPlannedBunches  = numOrDash_(nv[1]);
      nextPlannedVolume   = numOrDash_(nv[9]);
      nextPlannedRevenue  = fmtOrDash_(nv[13]);
      nextPlannedCosts    = fmtOrDash_(nv[15]);
      nextPlannedMargin   = fmtOrDash_(nv[17]);
    } catch (e) {
      Logger.log('Could not fetch next month data: ' + e.message);
    }
  }

  // YTD variances
  const ytdBunchesVar  = ytd.actualBunches  > 0 ? varSign_(ytd.actualBunches  - ytd.plannedBunches)  : '—';
  const ytdVolumeVar   = ytd.actualVolume   > 0 ? varSign_(ytd.actualVolume   - ytd.plannedVolume)   : '—';
  const ytdRevenueVar  = ytd.actualRevenue  > 0 ? varSign_(ytd.actualRevenue  - ytd.plannedRevenue)  : '—';
  const ytdCostsVar    = ytd.actualCosts    > 0 ? varSign_(ytd.actualCosts    - ytd.plannedCosts)    : '—';
  const ytdMarginVar   = ytd.actualMargin   > 0 ? varSign_(ytd.actualMargin   - ytd.plannedMargin)   : '—';
  // ytd.actualMarginPct and plannedMarginPct are already formatted as "25.2%"
  // parseFloat("25.2%") correctly yields 25.2, so subtraction is safe.
  const ytdMarginPctVar = (ytd.actualMarginPct !== '—' && ytd.plannedMarginPct !== '—')
    ? varSign_(parseFloat(ytd.actualMarginPct) - parseFloat(ytd.plannedMarginPct)) + '%'
    : '—';

  // Auto-commentary
  const execSummary     = buildExecutiveSummary_(month, d, revenueVar, marginVar);
  const prodNotes       = buildProductionNotes_(month, d, bunchesVar, volumeVar);
  const finNotes        = buildFinancialNotes_(month, d, revenueVar, costsVar, marginVar);
  const varianceDrivers = buildVarianceDrivers_(month, d, bunchesVar, weightVar, priceVar, revenueVar, costsVar);
  const ytdNarrative    = buildYtdNarrative_(month, ytd);
  const recommendations = buildRecommendations_(month, d, revenueVar, costsVar);

  return {
    '{{MONTH_YEAR}}':            `${month.full} ${CONFIG.REPORT_YEAR}`,
    '{{MONTH_SHORT}}':           month.short,
    '{{REPORT_DATE}}':           reportDate,
    '{{REPORT_STATUS}}':         reportStatus,
    '{{MONTHS_COMPLETED}}':      String(month.num),
    '{{PLANT_POPULATION}}':      numOrDash_(d.plantPopulation),

    // Production
    '{{PLANNED_BUNCHES}}':       numOrDash_(d.plannedBunches),
    '{{ACTUAL_BUNCHES}}':        d.hasActuals ? numOrDash_(d.actualBunches) : '—',
    '{{BUNCHES_VAR}}':           varOrDash(bunchesVar, varSign_),
    '{{BUNCHES_VAR_PCT}}':       varPct(bunchesVar, d.plannedBunches),
    '{{PLANNED_AVG_WEIGHT}}':    numOrDash_(d.plannedAvgWeight),
    '{{ACTUAL_AVG_WEIGHT}}':     d.hasActuals ? numOrDash_(d.actualAvgWeight) : '—',
    '{{WEIGHT_VAR}}':            varOrDash(weightVar, varSign_),
    '{{WEIGHT_VAR_PCT}}':        varPct(weightVar, d.plannedAvgWeight),
    // FIXED: bearing rates are now pre-formatted "X.XX%" strings from getMonthData_()
    '{{PLANNED_BEARING_RATE}}':  d.plannedBearingRate,
    '{{ACTUAL_BEARING_RATE}}':   d.actualBearingRate,
    '{{BEARING_VAR}}':           '—',
    '{{PLANNED_VOLUME}}':        numOrDash_(d.plannedVolume),
    '{{ACTUAL_VOLUME}}':         d.hasActuals ? numOrDash_(d.actualVolume) : '—',
    '{{VOLUME_VAR}}':            varOrDash(volumeVar, varSign_),
    '{{VOLUME_VAR_PCT}}':        varPct(volumeVar, d.plannedVolume),
    '{{PLANNED_PRICE}}':         numOrDash_(d.plannedPrice),
    '{{ACTUAL_PRICE}}':          d.hasActuals ? numOrDash_(d.actualPrice) : '—',
    '{{PRICE_VAR}}':             varOrDash(priceVar, varSign_),
    '{{PRICE_VAR_PCT}}':         varPct(priceVar, d.plannedPrice),

    // Financials (formatted with commas via fmt_())
    '{{PLANNED_REVENUE_FMT}}':   fmt_(d.plannedRevenue),
    '{{ACTUAL_REVENUE_FMT}}':    d.hasActuals ? fmt_(d.actualRevenue) : '—',
    '{{REVENUE_VAR_FMT}}':       varOrDash(revenueVar, v => fmt_(v)),
    '{{REVENUE_VAR_PCT}}':       varPct(revenueVar, d.plannedRevenue),
    '{{PLANNED_COSTS_FMT}}':     fmt_(d.plannedCosts),
    '{{ACTUAL_COSTS_FMT}}':      d.hasActuals ? fmt_(d.actualCosts) : '—',
    '{{COSTS_VAR_FMT}}':         varOrDash(costsVar, v => fmt_(v)),
    '{{COSTS_VAR_PCT}}':         varPct(costsVar, d.plannedCosts),
    '{{PLANNED_MARGIN_FMT}}':    fmt_(d.plannedMargin),
    '{{ACTUAL_MARGIN_FMT}}':     d.hasActuals ? fmt_(d.actualMargin) : '—',
    '{{MARGIN_VAR_FMT}}':        varOrDash(marginVar, v => fmt_(v)),
    '{{MARGIN_VAR_PCT}}':        varPct(marginVar, d.plannedMargin),
    // FIXED: margin % are now pre-formatted "X.X%" strings from getMonthData_()
    '{{PLANNED_MARGIN_PCT}}':    d.plannedMarginPct,
    '{{ACTUAL_MARGIN_PCT}}':     d.actualMarginPct,
    '{{MARGIN_PCT_VAR}}':        '—',

    // Cost efficiency
    '{{PLANNED_COST_PER_KG}}':   plannedCostPerKg,
    '{{ACTUAL_COST_PER_KG}}':    actualCostPerKg,
    '{{PLANNED_REV_PER_KG}}':    plannedRevPerKg,
    '{{ACTUAL_REV_PER_KG}}':     actualRevPerKg,
    '{{PLANNED_MARGIN_BUNCH}}':  plannedMarginBunch,
    '{{ACTUAL_MARGIN_BUNCH}}':   actualMarginBunch,

    // Variance analysis
    '{{VOL_GAP_FMT}}':           varOrDash(volumeVar, v => fmt_(v)),
    '{{VOL_DIRECTION}}':         varOrDash(volumeVar, v => v >= 0 ? '▲ Above Plan' : '▼ Below Plan'),
    '{{VOL_IMPACT}}':            varOrDash(volumeVar, v => v >= 0 ? 'Positive — more product available' : 'Negative — reduced sellable volume'),
    '{{REV_GAP_FMT}}':           varOrDash(revenueVar, v => fmt_(v)),
    '{{REV_DIRECTION}}':         varOrDash(revenueVar, v => v >= 0 ? '▲ Above Plan' : '▼ Below Plan'),
    '{{REV_IMPACT}}':            varOrDash(revenueVar, v => v >= 0 ? 'Positive — revenue beat target' : 'Negative — revenue shortfall'),
    '{{COST_GAP_FMT}}':          varOrDash(costsVar, v => fmt_(v)),
    '{{COST_DIRECTION}}':        varOrDash(costsVar, v => v <= 0 ? '▲ Under Budget' : '▼ Over Budget'),
    '{{COST_IMPACT}}':           varOrDash(costsVar, v => v <= 0 ? 'Positive — costs controlled' : 'Negative — over-expenditure'),
    '{{MAR_GAP_FMT}}':           varOrDash(marginVar, v => fmt_(v)),
    '{{MAR_DIRECTION}}':         varOrDash(marginVar, v => v >= 0 ? '▲ Above Plan' : '▼ Below Plan'),
    '{{MAR_IMPACT}}':            varOrDash(marginVar, v => v >= 0 ? 'Positive — stronger profitability' : 'Negative — margin compression'),

    // YTD
    '{{YTD_PLANNED_BUNCHES}}':    numOrDash_(ytd.plannedBunches),
    '{{YTD_ACTUAL_BUNCHES}}':     ytd.actualBunches > 0 ? numOrDash_(ytd.actualBunches) : '—',
    '{{YTD_BUNCHES_VAR}}':        ytdBunchesVar,
    '{{YTD_PLANNED_VOLUME}}':     numOrDash_(ytd.plannedVolume),
    '{{YTD_ACTUAL_VOLUME}}':      ytd.actualVolume > 0 ? numOrDash_(ytd.actualVolume) : '—',
    '{{YTD_VOLUME_VAR}}':         ytdVolumeVar,
    '{{YTD_PLANNED_REVENUE}}':    fmt_(ytd.plannedRevenue),
    '{{YTD_ACTUAL_REVENUE}}':     ytd.actualRevenue > 0 ? fmt_(ytd.actualRevenue) : '—',
    '{{YTD_REVENUE_VAR}}':        ytdRevenueVar,
    '{{YTD_PLANNED_COSTS}}':      fmt_(ytd.plannedCosts),
    '{{YTD_ACTUAL_COSTS}}':       ytd.actualCosts > 0 ? fmt_(ytd.actualCosts) : '—',
    '{{YTD_COSTS_VAR}}':          ytdCostsVar,
    '{{YTD_PLANNED_MARGIN}}':     fmt_(ytd.plannedMargin),
    '{{YTD_ACTUAL_MARGIN}}':      ytd.actualMargin > 0 ? fmt_(ytd.actualMargin) : '—',
    '{{YTD_MARGIN_VAR}}':         ytdMarginVar,
    '{{YTD_PLANNED_MARGIN_PCT}}': ytd.plannedMarginPct,
    '{{YTD_ACTUAL_MARGIN_PCT}}':  ytd.actualMarginPct,
    '{{YTD_MARGIN_PCT_VAR}}':     ytdMarginPctVar,

    // Next month targets
    '{{NEXT_PLANNED_BUNCHES}}':  nextPlannedBunches,
    '{{NEXT_PLANNED_VOLUME}}':   nextPlannedVolume,
    '{{NEXT_PLANNED_REVENUE}}':  nextPlannedRevenue,
    '{{NEXT_PLANNED_COSTS}}':    nextPlannedCosts,
    '{{NEXT_PLANNED_MARGIN}}':   nextPlannedMargin,

    // Auto-generated narrative
    '{{EXECUTIVE_SUMMARY}}':     execSummary,
    '{{PRODUCTION_NOTES}}':      prodNotes,
    '{{FINANCIAL_NOTES}}':       finNotes,
    '{{VARIANCE_DRIVERS}}':      varianceDrivers,
    '{{YTD_NARRATIVE}}':         ytdNarrative,
    '{{RECOMMENDATIONS}}':       recommendations,

    // Company name safety net
    'Mayilo Enterprises':        CONFIG.COMPANY_NAME
  };
}

// ─── AUTO-COMMENTARY BUILDERS ─────────────────────────────────────────────────
// All builders use d.actualMarginPct / d.plannedMarginPct which are now
// pre-formatted "X.X%" strings — no more raw floats in narrative text.

function buildExecutiveSummary_(month, d, revenueVar, marginVar) {
  if (!d.hasActuals) {
    return `This report presents the planned operational targets for ${month.full} ${CONFIG.REPORT_YEAR}. ` +
           `Actual performance data has not yet been recorded for this period. ` +
           `The planned harvest of ${numOrDash_(d.plannedBunches)} bunches at an average weight of ` +
           `${numOrDash_(d.plannedAvgWeight)} kg is expected to yield ${numOrDash_(d.plannedVolume)} kg of produce, ` +
           `generating a projected revenue of MWK ${fmt_(d.plannedRevenue)} against planned costs of MWK ${fmt_(d.plannedCosts)}, ` +
           `targeting a gross margin of MWK ${fmt_(d.plannedMargin)} (${d.plannedMarginPct}).`;
  }
  const rvDir = revenueVar >= 0 ? 'exceeded' : 'fell short of';
  const mDir  = marginVar  >= 0 ? 'positive variance' : 'negative variance';
  return `${month.full} ${CONFIG.REPORT_YEAR} saw Elim Farms harvest ${numOrDash_(d.actualBunches)} bunches ` +
         `(plan: ${numOrDash_(d.plannedBunches)}), producing ${numOrDash_(d.actualVolume)} kg of bananas. ` +
         `Revenue of MWK ${fmt_(d.actualRevenue)} ${rvDir} the plan of MWK ${fmt_(d.plannedRevenue)}, ` +
         `while operating costs were MWK ${fmt_(d.actualCosts)}. ` +
         `The farm recorded a gross margin of MWK ${fmt_(d.actualMargin)} (${d.actualMarginPct}), ` +
         `representing a ${mDir} of MWK ${fmt_(Math.abs(marginVar || 0))} against the plan.`;
}

function buildProductionNotes_(month, d, bunchesVar, volumeVar) {
  if (!d.hasActuals) {
    return `Production targets are set based on the seasonal growth model and historical bearing rates. ` +
           `A plant population of ${numOrDash_(d.plantPopulation)} is expected to achieve a bearing rate of ${d.plannedBearingRate}. ` +
           `Field teams should track bunch development and flag any deviations from plan promptly.`;
  }
  const vDir = (bunchesVar || 0) >= 0 ? 'above' : 'below';
  return `Actual harvest of ${numOrDash_(d.actualBunches)} bunches was ${vDir} the planned ${numOrDash_(d.plannedBunches)}. ` +
         `Average bunch weight came in at ${numOrDash_(d.actualAvgWeight)} kg versus the planned ${numOrDash_(d.plannedAvgWeight)} kg. ` +
         `The actual bearing rate of ${d.actualBearingRate} (planned ${d.plannedBearingRate}) reflects current field conditions. ` +
         `Total production volume of ${numOrDash_(d.actualVolume)} kg was recorded for the month.`;
}

function buildFinancialNotes_(month, d, revenueVar, costsVar, marginVar) {
  if (!d.hasActuals) {
    return `Financial projections are derived from planned production volumes at a target selling price of ` +
           `MWK ${numOrDash_(d.plannedPrice)}/kg. Budget allocations for operational costs total MWK ${fmt_(d.plannedCosts)}. ` +
           `Finance team should update actuals upon completion of the month to enable variance analysis.`;
  }
  const costDir = (costsVar || 0) <= 0 ? 'under budget by' : 'over budget by';
  return `Gross revenue of MWK ${fmt_(d.actualRevenue)} was achieved at an average selling price of ` +
         `MWK ${numOrDash_(d.actualPrice)}/kg. Operating costs were ${costDir} MWK ${fmt_(Math.abs(costsVar || 0))} ` +
         `relative to the plan. The resulting gross margin of MWK ${fmt_(d.actualMargin)} represents ` +
         `a margin percentage of ${d.actualMarginPct} of revenue. ` +
         `Finance should review cost allocations to identify savings opportunities for future months.`;
}

function buildVarianceDrivers_(month, d, bunchesVar, weightVar, priceVar, revenueVar, costsVar) {
  if (!d.hasActuals) {
    return `Variance analysis will be available once actual data for ${month.full} is recorded in the system.`;
  }
  const drivers = [];
  if (bunchesVar !== null && Math.abs(bunchesVar) > 10)
    drivers.push(`• Harvest volume: ${(bunchesVar || 0) >= 0 ? '+' : ''}${numOrDash_(bunchesVar)} bunches vs plan — ` +
                 `driven by ${(bunchesVar || 0) >= 0 ? 'favourable field conditions' : 'adverse weather or pest pressure'}.`);
  if (weightVar !== null && Math.abs(weightVar) > 0.5)
    drivers.push(`• Bunch weight: ${(weightVar || 0) >= 0 ? '+' : ''}${weightVar?.toFixed(1)} kg vs plan — ` +
                 `attributable to ${(weightVar || 0) >= 0 ? 'good crop nutrition' : 'early harvesting or nutritional deficit'}.`);
  if (priceVar !== null && Math.abs(priceVar) > 50)
    drivers.push(`• Selling price: MWK ${varSign_(priceVar)}/kg vs plan — ` +
                 `reflecting ${(priceVar || 0) >= 0 ? 'stronger market demand' : 'market price softness or buyer terms'}.`);
  if (costsVar !== null && Math.abs(costsVar) > 100000)
    drivers.push(`• Costs: MWK ${fmt_(Math.abs(costsVar))} ${(costsVar || 0) > 0 ? 'over' : 'under'} budget — ` +
                 `review input cost lines for this period.`);
  return drivers.length > 0 ? drivers.join('\n') :
    'No material variances identified. Performance is broadly in line with plan.';
}

function buildYtdNarrative_(month, ytd) {
  if (ytd.actualRevenue === 0) {
    return `Year-to-date planned revenue of MWK ${fmt_(ytd.plannedRevenue)} has been budgeted through ${month.full}. ` +
           `Actual data will be updated as each month's performance is captured.`;
  }
  const revenueAchieved = ytd.plannedRevenue > 0
    ? ((ytd.actualRevenue / ytd.plannedRevenue) * 100).toFixed(1) : '—';
  return `Through ${month.full} ${CONFIG.REPORT_YEAR}, Elim Farms has achieved MWK ${fmt_(ytd.actualRevenue)} in revenue ` +
         `(${revenueAchieved}% of YTD plan), harvesting ${numOrDash_(ytd.actualBunches)} bunches totalling ` +
         `${numOrDash_(ytd.actualVolume)} kg. YTD gross margin stands at MWK ${fmt_(ytd.actualMargin)} (${ytd.actualMarginPct}).`;
}

function buildRecommendations_(month, d, revenueVar, costsVar) {
  const recs = [];
  if (!d.hasActuals) {
    recs.push('• Ensure field supervisors record all harvest data promptly at end of month.');
    recs.push('• Finance team to capture actual cost invoices and allocations before the 5th of the following month.');
    recs.push('• Agronomist to assess crop conditions and bearing rates mid-month for early-warning signals.');
  } else {
    if ((revenueVar || 0) < 0)
      recs.push('• Revenue shortfall: review buyer offtake agreements and explore additional market channels to close the gap.');
    if ((costsVar || 0) > 0)
      recs.push('• Cost overrun: conduct a line-item cost review and identify areas for renegotiation or efficiency gains.');
    if ((revenueVar || 0) >= 0 && (costsVar || 0) <= 0)
      recs.push('• Strong performance this month — maintain current operational practices and field management routines.');
    recs.push('• Ensure next month\'s field activities are planned and inputs procured ahead of schedule.');
    recs.push('• Update cash flow projections based on this month\'s actuals and revised market price outlook.');
  }
  return recs.join('\n');
}

// ─── PLACEHOLDER REPLACEMENT ───────────────────────────────────────────────────
function replacePlaceholders_(doc, placeholders) {
  const body = doc.getBody();

  // Replace in the main body
  Object.entries(placeholders).forEach(([key, value]) => {
    body.replaceText(escapeRegex_(key), String(value));
  });

  // Replace in headers and footers
  const sections = [
    ...doc.getHeader() ? [doc.getHeader()] : [],
    ...doc.getFooter() ? [doc.getFooter()] : []
  ];
  sections.forEach(section => {
    if (section) {
      Object.entries(placeholders).forEach(([key, value]) => {
        section.replaceText(escapeRegex_(key), String(value));
      });
    }
  });
}

function escapeRegex_(str) {
  // Escape special regex characters in placeholder keys
  return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

// ─── DUPLICATE CHECK ───────────────────────────────────────────────────────────
function reportAlreadyExists_(month) {
  const folder = DriveApp.getFolderById(CONFIG.OUTPUT_FOLDER_ID);
  const expectedName = `Elim Farms Monthly Report — ${month.full} ${CONFIG.REPORT_YEAR}.pdf`;
  const files = folder.getFilesByName(expectedName);
  return files.hasNext();
}

// ─── EMAIL SENDING ─────────────────────────────────────────────────────────────
function sendReportEmails_(month, d, pdfBlob, driveUrl) {
  const subject = `[Elim Group] Elim Farms Monthly Operational Report — ${month.full} ${CONFIG.REPORT_YEAR}`;

  const statusNote = d.hasActuals
    ? `This report contains <strong>actual performance data</strong> for ${month.full} ${CONFIG.REPORT_YEAR}.`
    : `This report presents <strong>planned targets only</strong>. Actual figures will be circulated once captured.`;

  const htmlBody = `
    <div style="font-family:Arial,sans-serif;max-width:680px;margin:0 auto;">
      <div style="background:#0D2137;padding:32px 40px;border-radius:4px 4px 0 0;">
        <h1 style="color:#ffffff;font-size:22px;margin:0 0 6px 0;">Elim Group Limited</h1>
        <p style="color:#A9CCE3;margin:0;font-size:13px;">Elim Farms Monthly Operational Report</p>
      </div>
      <div style="background:#f8f9fa;padding:32px 40px;border:1px solid #dee2e6;">
        <h2 style="color:#1B4F72;font-size:18px;margin:0 0 16px 0;">${month.full} ${CONFIG.REPORT_YEAR}</h2>
        <p style="color:#444;line-height:1.6;">${statusNote}</p>
        <p style="color:#444;line-height:1.6;">
          Please find the full report attached as a PDF. The report covers:
        </p>
        <ul style="color:#444;line-height:1.8;">
          <li>Executive Summary &amp; KPI Highlights</li>
          <li>Production Performance (bunches, weight, bearing rate, volume)</li>
          <li>Financial Performance (revenue, costs, margins)</li>
          <li>Variance Analysis (planned vs. actual)</li>
          <li>Year-to-Date Cumulative Performance</li>
          <li>Outlook &amp; Recommendations for the coming month</li>
        </ul>
        <p style="margin:24px 0 0 0;">
          <a href="${driveUrl}"
             style="background:#1B4F72;color:#fff;padding:10px 20px;border-radius:4px;text-decoration:none;font-weight:bold;">
            View on Google Drive
          </a>
        </p>
      </div>
      <div style="background:#eaecef;padding:16px 40px;border-radius:0 0 4px 4px;font-size:11px;color:#888;">
        <p style="margin:0;">This is an automated report generated by the Elim Farms Reporting System.</p>
        <p style="margin:4px 0 0 0;">Elim Group Limited | Elim Farms Banana Operations | Production Year 2026</p>
        <p style="margin:4px 0 0 0;">CONFIDENTIAL — For internal distribution only.</p>
      </div>
    </div>`;

  CONFIG.EMAIL_RECIPIENTS.forEach(email => {
    if (!email || email.includes('example.com')) return; // skip placeholder emails
    try {
      GmailApp.sendEmail(email, subject, '', {
        htmlBody: htmlBody,
        attachments: [pdfBlob.copyBlob()],
        name: CONFIG.COMPANY_NAME + ' — Farm Reports'
      });
      Logger.log(`📧 Email sent to: ${email}`);
    } catch (e) {
      Logger.log(`❌ Failed to send to ${email}: ${e.message}`);
    }
  });
}

// ─── UTILITY FUNCTIONS ────────────────────────────────────────────────────────

/**
 * Parse a number from a cell value that may be a string with commas.
 */
function parseNum_(val) {
  if (val === null || val === undefined || val === '' || val === '—') return 0;
  if (typeof val === 'number') return val;
  const s = String(val).replace(/,/g, '').replace(/[^0-9.\-]/g, '');
  const n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

/**
 * Format a number as MWK with comma-separated thousands, no decimals.
 * e.g. 3897050 → "3,897,050"
 */
function fmt_(n) {
  if (!n && n !== 0) return '—';
  return Math.round(n).toLocaleString('en-US');
}

/**
 * Format a number (not currency) with commas and up to 1 decimal place.
 * e.g. 4744 → "4,744"  |  10.9 → "10.9"
 */
function numOrDash_(n) {
  if (!n && n !== 0) return '—';
  const rounded = Math.round(n * 10) / 10;
  return rounded.toLocaleString('en-US');
}

function fmtOrDash_(v) {
  const n = parseNum_(v);
  return n > 0 ? fmt_(n) : '—';
}

/**
 * Format a whole-number percentage (e.g. 25.2 → "25.2%").
 * NOTE: callers must multiply decimal fractions by 100 BEFORE passing in.
 */
function pct_(n) {
  return (Math.round(n * 10) / 10).toFixed(1) + '%';
}

/**
 * NEW v2.1 — Convert a decimal-fraction margin % from the sheet to a
 * display-ready percentage string.
 * Sheet stores e.g. 0.06874... (= 6.874%)  →  returns "6.9%"
 * Sheet stores e.g. -0.38012... (= -38.0%) →  returns "-38.0%"
 */
function formatDecimalPct_(val) {
  if (val === null || val === undefined || val === '' || val === '—') return '—';
  const n = typeof val === 'number' ? val : parseFloat(String(val).replace(/,/g, ''));
  if (isNaN(n)) return '—';
  // Values are stored as decimal fractions; multiply by 100 to get %
  return pct_(n * 100);
}

/**
 * NEW v2.1 — Convert a decimal-fraction bearing rate from the sheet to a
 * display-ready percentage string with 2 decimal places.
 * Sheet stores e.g. 0.0613 (= 6.13%) → returns "6.13%"
 * Sheet stores e.g. 0.03445... → returns "3.45%"
 */
function formatBearingRate_(val) {
  if (val === null || val === undefined || val === '' || val === '—') return '—';
  const n = typeof val === 'number' ? val : parseFloat(String(val).replace(/,/g, ''));
  if (isNaN(n)) return '—';
  // Multiply by 100 and keep 2 decimal places for bearing rate precision
  return (n * 100).toFixed(2) + '%';
}

/**
 * Sign-prefixed variant for variance values with comma formatting.
 * e.g. -4798 → "-4,798"  |  7343626 → "+7,343,626"
 */
function varSign_(n) {
  if (n === null || n === undefined || n === 0) return '0';
  const formatted = Math.round(Math.abs(n)).toLocaleString('en-US');
  return (n >= 0 ? '+' : '-') + formatted;
}

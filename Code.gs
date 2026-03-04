/**
 * Belvidere CCAP Monthly Attendance Processor
 * Processes monthly attendance data from RecliqueCore Program Check-Ins exports.
 * Automatically creates monthly tabs, archives raw data, and calculates daily averages.
 *
 * Based on the original Childcare Monthly Attendance Processor structure.
 * Updated to accept RecliqueCore's row-per-check-in format and adds support
 * for Schools Out / Full Day (FD) attendance sessions.
 *
 * ── HOW TO USE ───────────────────────────────────────────────────────────────
 *  1. Export the Program Check-Ins report from RecliqueCore (CSV or copy-paste).
 *  2. Open the "Attendance Report" tab and paste starting at cell A10.
 *     Headers will land at row 14; data begins at row 15.
 *  3. Check ONE box: BFY, Pop.Grv., or Both.
 *  4. Click the "Process Attendance" button (or use Attendance Tools menu).
 *
 * ── DATA FORMAT (RecliqueCore Program Check-Ins Report) ──────────────────────
 *  Column A (0) : Checked In  – date/time of check-in
 *  Column C (2) : Member ID   – unique child identifier (used for de-duplication)
 *  Column D (3) : Participant – display name ("Last, First")
 *  Column E (4) : Program     – only "School Age Child Care" rows are processed
 *  Column F (5) : Divisions   – determines site (see SITE MAPPING below)
 *
 * ── SITE MAPPING ─────────────────────────────────────────────────────────────
 *  Division starts with "North Boone" → Pop.Grv.  (Poplar Grove Elementary)
 *  Division starts with "Schools Out" → BFY, Full Day (FD)
 *  Division starts with "Belvidere"   → BFY        (Belvidere Family Y)
 *
 * ── FULL DAY (FD) ROWS ───────────────────────────────────────────────────────
 *  "Schools Out" sessions are full-day care. They always appear on the BFY tab.
 *  A child with both regular B&A care AND Schools Out attendance at the same
 *  site appears as TWO rows in the monthly summary:
 *    "Smith, John"      → Before & After Care     (references B3: Eligible Days)
 *    "Smith, John (FD)" → Schools Out (Full Day)  (references B4: FD Eligible Days)
 *
 * ── DE-DUPLICATION ───────────────────────────────────────────────────────────
 *  A child checking in multiple times on the same date at the same site
 *  (e.g., separate before-care and after-care entries) counts as ONE day.
 */

// ===========================
// MENU SETUP
// ===========================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Attendance Tools')
    .addItem('Process Attendance', 'processAttendance')
    .addToUi();
}

/**
 * Set up the Attendance Report sheet with checkboxes and formatting.
 * Run this once to initialize the sheet (or to reset it after accidental edits).
 */
function setupAttendanceReportSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Attendance Report');
  if (!sheet) {
    sheet = ss.insertSheet('Attendance Report', 0);
  }
  sheet.clear();
  sheet.clearFormats();

  // Header question label
  sheet.getRange('A2:D2').merge();
  const headerCell = sheet.getRange('A2');
  headerCell.setValue('Which attendance report are you uploading?');
  headerCell.setFontWeight('bold');
  headerCell.setFontSize(10);
  headerCell.setFontFamily('Verdana');
  headerCell.setVerticalAlignment('middle');
  headerCell.setHorizontalAlignment('left');

  // Checkboxes: A3 = BFY, C3 = Pop.Grv., A4 = Both
  sheet.getRange('A3').insertCheckboxes();
  sheet.getRange('C3').insertCheckboxes();
  sheet.getRange('A4').insertCheckboxes();
  ['A3', 'C3', 'A4'].forEach(r => {
    sheet.getRange(r).setHorizontalAlignment('left').setVerticalAlignment('middle');
  });

  // Labels
  sheet.getRange('B3').setValue('BFY');
  sheet.getRange('D3').setValue('Pop.Grv.');
  sheet.getRange('B4').setValue('Both');
  ['B3', 'D3', 'B4'].forEach(r => {
    sheet.getRange(r).setFontSize(9).setFontFamily('Verdana').setVerticalAlignment('middle');
  });

  // Divider borders (matching original)
  sheet.getRange('A8:D8').setBorder(null, null, true, null, null, null, '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange('D1:D8').setBorder(null, null, null, true, null, null, '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID);

  // Paste-here marker at A10
  const pasteCell = sheet.getRange('A10');
  pasteCell.setValue('Paste here');
  pasteCell.setBackground('#FFFACD');
  pasteCell.setFontFamily('Verdana');
  pasteCell.setFontSize(9);

  // Sheet-wide defaults
  sheet.setHiddenGridlines(true);
  sheet.getRange('A1:Z100').setFontFamily('Verdana').setFontSize(9).setFontColor('#333333');
  sheet.autoResizeColumns(1, 4);

  SpreadsheetApp.getActiveSpreadsheet().toast('Attendance Report sheet has been set up!', 'Setup Complete', 3);
}

// ===========================
// MAIN PROCESSING FUNCTION
// ===========================

function processAttendance() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheet = ss.getSheetByName('Attendance Report');

  if (!sheet) {
    ui.alert('Error: "Attendance Report" sheet not found.');
    return;
  }

  // Read checkboxes: A3 = BFY, C3 = Pop.Grv., A4 = Both
  const bfyChecked    = sheet.getRange('A3').getValue();
  const popGrvChecked = sheet.getRange('C3').getValue();
  const bothChecked   = sheet.getRange('A4').getValue();

  const checkedCount = (bfyChecked ? 1 : 0) + (popGrvChecked ? 1 : 0) + (bothChecked ? 1 : 0);

  if (checkedCount === 0) {
    ui.alert('No Selection', 'Please select a report type (BFY, Pop.Grv., or Both).', ui.ButtonSet.OK);
    return;
  }
  if (checkedCount > 1) {
    ui.alert('Multiple Selections', 'Please select only ONE option (BFY, Pop.Grv., or Both).', ui.ButtonSet.OK);
    return;
  }

  let reportType = null;
  if (bfyChecked)      reportType = 'BFY';
  else if (popGrvChecked) reportType = 'Pop.Grv.';
  else if (bothChecked)   reportType = 'Both';

  handleReportTypeSelection(reportType);

  // Uncheck all boxes after processing
  sheet.getRange('A3:A4').uncheck();
  sheet.getRange('C3').uncheck();
}

// ===========================
// CORE HANDLER
// ===========================

function handleReportTypeSelection(reportType) {
  if (!reportType) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  ss.toast(`Processing ${reportType} attendance report...`, 'Processing', -1);

  const rawSheet = ss.getSheetByName('Attendance Report');
  if (!rawSheet) {
    ui.alert('Error: "Attendance Report" sheet not found.');
    return;
  }

  const lastRow = rawSheet.getLastRow();
  const lastCol = rawSheet.getLastColumn();

  if (lastRow < 15 || lastCol < 6) {
    ui.alert('No data found. Please paste the RecliqueCore attendance report starting at cell A10.');
    return;
  }

  const allData = rawSheet.getRange(1, 1, lastRow, lastCol).getValues();

  // Find the first row that contains an actual check-in date in column A
  const dataStartRow = findDataStartRow(allData);
  if (dataStartRow === -1) {
    ui.alert('Error: Could not find attendance data. Make sure you pasted the full report starting at A10.');
    return;
  }

  // actualData[0] = first check-in event row
  const actualData = allData.slice(dataStartRow);

  // Detect month from column A of the first check-in row
  const month = detectMonth(actualData[0]);
  if (!month) {
    ui.alert('Error: Could not detect month from data. Ensure the report contains valid check-in dates in column A.');
    return;
  }

  // Determine which sites to output
  const sitesToProcess = reportType === 'Both' ? ['BFY', 'Pop.Grv.'] : [reportType];

  // Process all check-in rows into an attendance map
  const attendanceMap = processAttendanceData(actualData, reportType);

  // Calculate daily averages per site
  const dailyAverages = calculateDailyAverages(actualData, month, reportType);

  // Update Daily Averages sheet first (so tab ordering stays correct)
  updateDailyAveragesSheet(ss, month, reportType, dailyAverages);

  // Create/update one monthly summary tab per site
  const tabNames = [];
  for (const site of sitesToProcess) {
    createOrUpdateMonthlySheet(ss, month, site, attendanceMap);
    tabNames.push(`${month}-${site}`);
  }

  // Archive raw check-in data (hidden tab)
  createOrUpdateArchiveSheet(ss, month, reportType, actualData);

  // Clear the Attendance Report sheet ready for next upload
  clearAttendanceReportSheet(rawSheet);

  ss.toast('', '', 1);

  const tabList = tabNames.map(t => `"${t}"`).join(' and ');
  ui.alert(
    'Success!',
    `Attendance data for ${month} (${reportType}) has been processed.\n\n` +
    `• Monthly summary: ${tabList} tab(s)\n` +
    `• Raw data archived: "${month}-${reportType} Archive" tab (hidden)\n` +
    `• Daily averages updated\n` +
    `• Attendance Report cleared and ready for next report`,
    ui.ButtonSet.OK
  );
}

// ===========================
// HELPER FUNCTIONS
// ===========================

/**
 * Return the 0-based array index of the first row whose column A contains
 * an actual check-in Date (or date-formatted string).
 * Metadata rows ("Report Title:", URL, "Generated:", blank) and the column
 * header row ("Checked In") contain strings, so they are naturally skipped.
 */
function findDataStartRow(allData) {
  for (let r = 0; r < allData.length; r++) {
    const cell = allData[r][0];
    if (cell instanceof Date && !isNaN(cell)) return r;
    // Fallback for string dates like "2/2/2026 5:40"
    if (typeof cell === 'string' && /^\d{1,2}\/\d{1,2}\/\d{4}/.test(cell)) return r;
  }
  return -1;
}

/**
 * Detect month abbreviation ("Jan", "Feb", …) from the check-in date
 * in column A of the first data row.
 */
function detectMonth(firstDataRow) {
  const monthNames = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  const cell = firstDataRow[0];
  if (cell instanceof Date && !isNaN(cell)) return monthNames[cell.getMonth()];
  if (typeof cell === 'string') {
    const d = new Date(cell);
    if (!isNaN(d.getTime())) return monthNames[d.getMonth()];
  }
  return null;
}

/**
 * Map a Divisions value (column F) to a site identifier.
 * Returns 'BFY', 'Pop.Grv.', or null for unrecognised divisions.
 * NOTE: "Schools Out" is checked before "Belvidere" so that any Schools Out
 * session containing the word "Belvidere" is still flagged as Full Day.
 */
function getSiteFromDivision(division) {
  if (!division) return null;
  const d = division.toString().trim();
  if (d.startsWith('North Boone')) return 'Pop.Grv.';
  if (d.startsWith('Schools Out')) return 'BFY';
  if (d.startsWith('Belvidere'))   return 'BFY';
  return null;
}

/**
 * Returns true when the division represents a Schools Out (Full Day) session.
 */
function isDivisionFullDay(division) {
  return division ? division.toString().trim().startsWith('Schools Out') : false;
}

/**
 * Process raw check-in rows into an attendance map.
 *
 * De-duplication: same Member ID + same site + same calendar date = 1 day,
 * regardless of how many check-in rows exist (before-care + after-care on
 * the same date collapses to a single counted day).
 *
 * Returns:
 *   { memberId: { name, sites: { site: { regular: {dateKey:true}, schoolsOut: {dateKey:true} } } } }
 */
function processAttendanceData(actualData, reportType) {
  const attendance = {};
  const COL = { date: 0, memberId: 2, name: 3, program: 4, division: 5 };

  for (const row of actualData) {
    if (!row[COL.date]) continue;

    // Only process "School Age Child Care" rows
    const program = (row[COL.program] || '').toString().trim();
    if (program !== 'School Age Child Care') continue;

    const checkIn = row[COL.date] instanceof Date ? row[COL.date] : new Date(row[COL.date]);
    if (isNaN(checkIn.getTime())) continue;

    const memberId = (row[COL.memberId] || '').toString().trim();
    const name     = (row[COL.name]     || '').toString().trim();
    const division = (row[COL.division] || '').toString().trim();

    if (!memberId || !name) continue;

    const site = getSiteFromDivision(division);
    if (!site) continue;

    // Skip sites that don't match the selected report type
    if (reportType !== 'Both' && site !== reportType) continue;

    const fullDay = isDivisionFullDay(division);
    const dateKey = Utilities.formatDate(checkIn, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const type    = fullDay ? 'schoolsOut' : 'regular';

    if (!attendance[memberId]) attendance[memberId] = { name, sites: {} };
    const A = attendance[memberId];
    if (!A.sites[site])           A.sites[site] = { regular: {}, schoolsOut: {} };

    // Object-as-set: assigning the same dateKey twice still counts as one day
    A.sites[site][type][dateKey] = true;
  }

  return attendance;
}

/**
 * Calculate daily average and peak attendance per site.
 * Counts unique children (by Member ID) who were present on each calendar date.
 * Average is rounded up (Math.ceil) to match the original script's behaviour.
 */
function calculateDailyAverages(actualData, month, reportType) {
  // siteData[site][dateKey] = { memberId: true, … }
  const siteData = {};
  const COL = { date: 0, memberId: 2, program: 4, division: 5 };

  for (const row of actualData) {
    if (!row[COL.date]) continue;

    const program = (row[COL.program] || '').toString().trim();
    if (program !== 'School Age Child Care') continue;

    const checkIn = row[COL.date] instanceof Date ? row[COL.date] : new Date(row[COL.date]);
    if (isNaN(checkIn.getTime())) continue;

    const memberId = (row[COL.memberId] || '').toString().trim();
    const division = (row[COL.division] || '').toString().trim();
    if (!memberId) continue;

    const site = getSiteFromDivision(division);
    if (!site) continue;
    if (reportType !== 'Both' && site !== reportType) continue;

    const dateKey = Utilities.formatDate(checkIn, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    if (!siteData[site])           siteData[site] = {};
    if (!siteData[site][dateKey])  siteData[site][dateKey] = {};
    siteData[site][dateKey][memberId] = true;
  }

  const averages = [];
  for (const site in siteData) {
    const dateCounts = {};
    for (const dateKey in siteData[site]) {
      dateCounts[dateKey] = Object.keys(siteData[site][dateKey]).length;
    }
    const dateKeys = Object.keys(dateCounts);
    if (!dateKeys.length) continue;

    const sum  = dateKeys.reduce((t, d) => t + dateCounts[d], 0);
    const avg  = Math.ceil(sum / dateKeys.length);
    const peak = Math.max(...Object.values(dateCounts));
    averages.push({ site, month, reportType: site, average: avg, peak });
  }

  return averages;
}

// ===========================
// SHEET CREATION FUNCTIONS
// ===========================

/**
 * Create or update the monthly summary sheet for a single site.
 *
 * Layout (matching original, with row 4 added for FD Eligible Days):
 *   Row 1 : Month: [month] ([site])
 *   Row 2 : Year:          [yellow input]
 *   Row 3 : Eligible Days: [yellow input]  ← B&A care
 *   Row 4 : FD Eligible Days: [yellow input]  ← Schools Out days  (new)
 *   Row 5 : Column headers
 *   Row 6+: Child data rows
 *
 * Children with both B&A care and Schools Out attendance appear as two rows:
 *   "Smith, John"      | <days> | Before & After Care    | =IF($B$3>0, B6/$B$3, 0)
 *   "Smith, John (FD)" | <days> | Schools Out (Full Day) | =IF($B$4>0, B7/$B$4, 0)
 */
function createOrUpdateMonthlySheet(ss, month, site, attendanceMap) {
  const tabName = `${month}-${site}`;
  let sheet = ss.getSheetByName(tabName);
  const dailyAvgSheet = ss.getSheetByName('Daily Averages');

  if (!sheet) {
    sheet = dailyAvgSheet
      ? ss.insertSheet(tabName, dailyAvgSheet.getIndex() + 1)
      : ss.insertSheet(tabName);
  } else {
    sheet.clear();
    if (dailyAvgSheet) {
      sheet.activate();
      ss.moveActiveSheet(dailyAvgSheet.getIndex() + 1);
    }
  }

  // ── Header section (rows 1–4) ───────────────────────────────────────────

  sheet.getRange('A1').setValue('Month:');
  sheet.getRange('B1').setValue(`${month} (${site})`);

  sheet.getRange('A2').setValue('Year:');
  sheet.getRange('B2')
    .setBackground('#FFFACD')
    .setHorizontalAlignment('left')
    .setBorder(true, true, true, true, false, false, '#CCCCCC', SpreadsheetApp.BorderStyle.DOTTED);

  sheet.getRange('A3').setValue('Eligible Days:');
  sheet.getRange('B3')
    .setBackground('#FFFACD')
    .setHorizontalAlignment('left')
    .setBorder(true, true, true, true, false, false, '#CCCCCC', SpreadsheetApp.BorderStyle.DOTTED);

  // FD Eligible Days — enter the number of Schools Out days in the month
  sheet.getRange('A4').setValue('FD Eligible Days:');
  sheet.getRange('B4')
    .setBackground('#FFFACD')
    .setHorizontalAlignment('left')
    .setBorder(true, true, true, true, false, false, '#CCCCCC', SpreadsheetApp.BorderStyle.DOTTED);

  sheet.getRange('A1:A4').setFontWeight('bold');

  // ── Column headers (row 5) ──────────────────────────────────────────────

  const headers = ['Last Name, First Name', 'Attended Days', 'Session Type', 'Attend %'];
  sheet.getRange(5, 1, 1, headers.length).setValues([headers]);

  // ── Build child rows ────────────────────────────────────────────────────

  const childRows = [];
  for (const memberId in attendanceMap) {
    const child = attendanceMap[memberId];
    if (!child.sites[site]) continue;

    const siteData = child.sites[site];
    const regCount  = Object.keys(siteData.regular).length;
    const fdCount   = Object.keys(siteData.schoolsOut).length;

    if (regCount > 0) {
      childRows.push({
        displayName: child.name,
        days:        regCount,
        sessionType: 'Before & After Care',
        type:        'regular',
        sortKey:     child.name
      });
    }
    if (fdCount > 0) {
      // "(FD)" suffix denotes Full Day / Schools Out care
      childRows.push({
        displayName: `${child.name} (FD)`,
        days:        fdCount,
        sessionType: 'Schools Out (Full Day)',
        type:        'schoolsOut',
        sortKey:     child.name
      });
    }
  }

  // Sort alphabetically; for the same child, regular row comes before FD row
  childRows.sort((a, b) => {
    const n = a.sortKey.localeCompare(b.sortKey);
    return n !== 0 ? n : (a.type === 'regular' ? -1 : 1);
  });

  // ── Write data rows starting at row 6 ──────────────────────────────────

  const dataRows = [];
  for (let i = 0; i < childRows.length; i++) {
    const cr     = childRows[i];
    const rowNum = 6 + i;
    // Regular rows use B3 (Eligible Days); FD rows use B4 (FD Eligible Days)
    const eligRef       = cr.type === 'regular' ? '$B$3' : '$B$4';
    const attendFormula = `=IF(${eligRef}>0, B${rowNum}/${eligRef}, 0)`;
    dataRows.push([cr.displayName, cr.days, cr.sessionType, attendFormula]);
  }

  if (dataRows.length > 0) {
    sheet.getRange(6, 1, dataRows.length, 4).setValues(dataRows);
  }

  applyMonthlySheetFormatting(sheet, dataRows.length);
  return sheet;
}

/**
 * Create or update archive sheet (hidden).
 * Stores the raw check-in rows for reference/audit.
 */
function createOrUpdateArchiveSheet(ss, month, reportType, rawData) {
  const archiveName = `${month}-${reportType} Archive`;
  let sheet = ss.getSheetByName(archiveName);
  if (!sheet) {
    sheet = ss.insertSheet(archiveName);
  } else {
    sheet.clear();
  }
  if (rawData.length > 0) {
    sheet.getRange(1, 1, rawData.length, rawData[0].length).setValues(rawData);
  }
  sheet.hideSheet();
  return sheet;
}

/**
 * Update the Daily Averages summary sheet.
 * Removes any existing rows for the current month + reportType, then appends
 * the newly calculated averages. Supports re-processing without duplicates.
 */
function updateDailyAveragesSheet(ss, month, reportType, averages) {
  let sheet = ss.getSheetByName('Daily Averages');
  if (!sheet) {
    const arSheet = ss.getSheetByName('Attendance Report');
    sheet = arSheet
      ? ss.insertSheet('Daily Averages', arSheet.getIndex() + 1)
      : ss.insertSheet('Daily Averages');
    sheet.getRange(4, 1, 1, 5).setValues([['Site', 'Month', 'Report Type', 'Average Attendance', 'Peak Attendance']]);
    applyDailyAveragesFormatting(sheet, 0);
  }

  const lastRow = Math.max(sheet.getLastRow(), 4);
  const numCols = Math.max(sheet.getLastColumn(), 5);

  // Read existing data rows (row 5 onwards)
  const existingData = lastRow > 4
    ? sheet.getRange(5, 1, lastRow - 4, numCols).getValues()
    : [];

  // Remove rows that belong to the month(s) being reprocessed
  const filtered = existingData.filter(row => {
    const rowMonth = row[1];
    const rowSite  = row[2];
    const isMatch  = rowMonth === month && (
      rowSite === reportType ||
      (reportType === 'Both' && (rowSite === 'BFY' || rowSite === 'Pop.Grv.'))
    );
    return !isMatch;
  });

  if (lastRow > 4) {
    sheet.getRange(5, 1, lastRow - 4, numCols).clear();
  }

  const allRows = filtered.filter(r => r[0]); // drop blank rows
  for (const avg of averages) {
    allRows.push([avg.site, avg.month, avg.site, avg.average, avg.peak]);
  }

  if (allRows.length > 0) {
    sheet.getRange(5, 1, allRows.length, 5).setValues(allRows);
  }

  applyDailyAveragesFormatting(sheet, allRows.length);
  return sheet;
}

/**
 * Clear pasted attendance data from the Attendance Report sheet
 * (rows 10 onwards), restore the "Paste here" marker, and uncheck all boxes.
 */
function clearAttendanceReportSheet(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow >= 10) {
    sheet.getRange(10, 1, lastRow - 9, lastCol).clear();
  }
  const pasteCell = sheet.getRange('A10');
  pasteCell.setValue('Paste here');
  pasteCell.setBackground('#FFFACD');
  pasteCell.setFontFamily('Verdana');
  pasteCell.setFontSize(9);
  sheet.getRange('A3:A4').uncheck();
  sheet.getRange('C3').uncheck();
}

// ===========================
// FORMATTING FUNCTIONS
// ===========================

/**
 * Apply formatting to a monthly summary sheet.
 * Matches the original script's style exactly.
 */
function applyMonthlySheetFormatting(sheet, dataRowCount) {
  sheet.setHiddenGridlines(true);
  sheet.getRange('A1:Z1000').setFontFamily('Verdana').setFontSize(9).setFontColor('#333333');

  const headerRange = sheet.getRange(5, 1, 1, 4);
  headerRange.setFontWeight('bold');
  headerRange.setBorder(null, null, true, null, null, null, '#666666', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  if (dataRowCount > 0) {
    // Attend % as percentage
    sheet.getRange(6, 4, dataRowCount, 1).setNumberFormat('0%');

    // Centre "Attended Days" column
    sheet.getRange(6, 2, dataRowCount, 1).setHorizontalAlignment('center');

    // Borders around data area
    const dataRange = sheet.getRange(5, 1, dataRowCount + 1, 4);
    dataRange.setBorder(true, true, true, true, false, false, '#666666', SpreadsheetApp.BorderStyle.SOLID);
    dataRange.setBorder(null, null, null, null, true, true, '#CCCCCC', SpreadsheetApp.BorderStyle.DOTTED);
    // Re-apply medium bottom border on header row
    headerRange.setBorder(null, null, true, null, null, null, '#666666', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

    // Conditional formatting: attend % < 70% → light red background
    const attendRange = sheet.getRange(6, 4, dataRowCount, 1);
    const rule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0.70)
      .setBackground('#FFE5E5')
      .setRanges([attendRange])
      .build();
    const rules = sheet.getConditionalFormatRules();
    rules.push(rule);
    sheet.setConditionalFormatRules(rules);

    // Filter on header row
    const existing = sheet.getFilter();
    if (existing) existing.remove();
    sheet.getRange(5, 1, dataRowCount + 1, 4).createFilter();
  }

  sheet.autoResizeColumns(1, 4);
}

/**
 * Apply formatting to the Daily Averages sheet.
 * Matches the original script's style exactly.
 */
function applyDailyAveragesFormatting(sheet, dataRowCount) {
  sheet.setHiddenGridlines(true);
  sheet.getRange('A1:Z1000').setFontFamily('Verdana').setFontSize(9).setFontColor('#333333');

  const headerRange = sheet.getRange(4, 1, 1, 5);
  headerRange.setFontWeight('bold');
  headerRange.setBorder(null, null, true, null, null, null, '#666666', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  if (dataRowCount > 0) {
    const dataRange = sheet.getRange(4, 1, dataRowCount + 1, 5);
    dataRange.setBorder(true, true, true, true, false, false, '#666666', SpreadsheetApp.BorderStyle.SOLID);
    dataRange.setBorder(null, null, null, null, true, true, '#CCCCCC', SpreadsheetApp.BorderStyle.DOTTED);
    headerRange.setBorder(null, null, true, null, null, null, '#666666', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

    const existing = sheet.getFilter();
    if (existing) existing.remove();
    sheet.getRange(4, 1, dataRowCount + 1, 5).createFilter();
  }

  sheet.autoResizeColumns(1, 5);
}

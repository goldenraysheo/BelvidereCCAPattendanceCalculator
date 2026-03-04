/**
 * Belvidere CCAP Attendance Calculator
 *
 * Reads the Program Check-Ins report exported from RecliqueCore and generates
 * monthly attendance summary tabs, one per site per month.
 *
 * ── DATA FORMAT ────────────────────────────────────────────────────────────
 *  Rows 1–4 : Report metadata (title, URL, generated date, etc.)
 *  Row 5    : Column headers
 *  Row 6+   : Attendance data
 *
 *  Column A (0) : Checked In  – date/time of check-in
 *  Column C (2) : Member ID   – unique child identifier (used for de-duplication)
 *  Column D (3) : Participant – display name ("Last, First")
 *  Column E (4) : Program     – only "School Age Child Care" rows are processed
 *  Column F (5) : Divisions   – determines site assignment (see SITE MAPPING)
 *
 * ── SITE MAPPING ────────────────────────────────────────────────────────────
 *  Division starts with "North Boone" → Pop.Grv.  (Poplar Grove Elementary)
 *  Division starts with "Belvidere"   → BFY        (Belvidere Family Y)
 *  Division starts with "Schools Out" → BFY, Full Day (FD)
 *
 *  Schools Out (Full Day) rules:
 *   • Children from any site may attend Schools Out sessions.
 *   • Schools Out totals always appear on the BFY tab only.
 *   • Children listed on a Schools Out row are marked "(FD)" to denote
 *     Full Day care, distinguished from regular Before & After Care.
 *   • If a child has both regular B&A care AND Schools Out attendance
 *     on the same site (e.g., Belvidere School District children), they
 *     appear as TWO rows on that tab: one for B&A care, one "(FD)" row.
 *
 * ── DE-DUPLICATION ──────────────────────────────────────────────────────────
 *  A child is counted once per calendar date per site, regardless of how
 *  many check-in rows appear for them that day (e.g., separate before/after
 *  care entries on the same date collapse to a single day attended).
 *
 * ── OUTPUT TABS ─────────────────────────────────────────────────────────────
 *  Named "{Mon} - {Site}"  →  "Feb - BFY",  "Feb - Pop.Grv."
 *  Each tab lists enrolled children, their session type, and days attended,
 *  followed by a summary section with enrollment counts and total day counts.
 */

// ── Configuration ────────────────────────────────────────────────────────────

var CONFIG = {
  /** Leave blank to use the first sheet; set a name to target a specific sheet. */
  dataSheetName: '',

  /** Row number (1-based) where column headers appear in the source data. */
  headerRow: 5,

  /** First row of actual attendance data (1-based). */
  dataStartRow: 6,

  /** Only rows matching this value in the Program column are processed. */
  programFilter: 'School Age Child Care',

  /** Zero-based column indices matching the source report layout. */
  cols: {
    date:      0,  // A – Checked In (date/time)
    memberId:  2,  // C – Member ID
    name:      3,  // D – Participant
    program:   4,  // E – Program
    division:  5   // F – Divisions
  },

  /** Short site identifiers used in tab names. */
  sites: {
    BFY:    'BFY',
    POPGRV: 'Pop.Grv.'
  },

  months: ['Jan','Feb','Mar','Apr','May','Jun',
           'Jul','Aug','Sep','Oct','Nov','Dec'],

  /** Hex colors for output tab formatting. */
  colors: {
    headerBg:   '#4A86E8',  // Column header background
    headerFg:   '#FFFFFF',  // Column header text
    fullDayBg:  '#FFF2CC',  // Light yellow – Schools Out (FD) rows
    summaryBg:  '#E8F0FE',  // Light blue   – summary section
    altRowBg:   '#F8F9FA'   // Very light grey – alternate regular rows
  }
};

// ── Internal helpers ─────────────────────────────────────────────────────────

/** Safe prefix check (Apps Script V8 supports startsWith but older runtimes may not). */
function startsWith_(str, prefix) {
  return str.length >= prefix.length && str.slice(0, prefix.length) === prefix;
}

/**
 * Map a Divisions value to a site identifier.
 * Returns null for unrecognised divisions (those rows are skipped).
 * NOTE: "Schools Out" is checked before "Belvidere" so that Schools Out
 * sessions that may contain the word "Belvidere" are still flagged as FD.
 */
function getSite_(division) {
  if (!division) return null;
  var d = division.toString().trim();
  if (startsWith_(d, 'North Boone'))  return CONFIG.sites.POPGRV;
  if (startsWith_(d, 'Schools Out'))  return CONFIG.sites.BFY;
  if (startsWith_(d, 'Belvidere'))    return CONFIG.sites.BFY;
  return null;
}

/** Returns true when the division represents a Schools Out (Full Day) session. */
function isFullDay_(division) {
  if (!division) return false;
  return startsWith_(division.toString().trim(), 'Schools Out');
}

/** Format a Date as YYYY-MM-DD for use as a de-duplication key. */
function toDateKey_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

/** Return the 3-letter month abbreviation for a Date (e.g., "Feb"). */
function getMonthAbbr_(date) {
  return CONFIG.months[date.getMonth()];
}

// ── Main processing ──────────────────────────────────────────────────────────

/**
 * Entry point – run this to process attendance data and generate output tabs.
 * Accessible from the CCAP Attendance menu added by onOpen().
 */
function processAttendance() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var dataSheet = CONFIG.dataSheetName
    ? ss.getSheetByName(CONFIG.dataSheetName)
    : ss.getSheets()[0];

  if (!dataSheet) {
    SpreadsheetApp.getUi().alert(
      'Data sheet not found.\n' +
      'Set CONFIG.dataSheetName to the exact sheet tab name, or leave it blank to use the first sheet.'
    );
    return;
  }

  var allData = dataSheet.getDataRange().getValues();
  var startIdx = CONFIG.dataStartRow - 1; // convert 1-based row to 0-based array index

  if (allData.length <= startIdx) {
    SpreadsheetApp.getUi().alert(
      'No attendance data found. Expected data to begin at row ' + CONFIG.dataStartRow + '.'
    );
    return;
  }

  /**
   * Attendance structure:
   *
   *  attendance[memberId] = {
   *    name: 'Last, First',
   *    sites: {
   *      'BFY': {
   *        'Feb 2026': {
   *          regular:    { 'YYYY-MM-DD': true, … },   // de-duplicated by date
   *          schoolsOut: { 'YYYY-MM-DD': true, … }
   *        }
   *      },
   *      'Pop.Grv.': { … }
   *    }
   *  }
   */
  var attendance = {};
  var cols = CONFIG.cols;

  for (var i = startIdx; i < allData.length; i++) {
    var row = allData[i];

    // Skip completely empty rows
    if (!row[cols.date]) continue;

    // Filter: Program column must equal "School Age Child Care"
    var program = (row[cols.program] || '').toString().trim();
    if (program !== CONFIG.programFilter) continue;

    // Parse check-in date
    var checkIn;
    if (row[cols.date] instanceof Date) {
      checkIn = row[cols.date];
    } else {
      checkIn = new Date(row[cols.date]);
    }
    if (isNaN(checkIn.getTime())) continue;

    var memberId = (row[cols.memberId] || '').toString().trim();
    var name     = (row[cols.name]     || '').toString().trim();
    var division = (row[cols.division] || '').toString().trim();

    if (!memberId || !name) continue;

    var site = getSite_(division);
    if (!site) continue; // Unrecognised division – skip

    var fullDay  = isFullDay_(division);
    var dateKey  = toDateKey_(checkIn);
    var monthKey = getMonthAbbr_(checkIn) + ' ' + checkIn.getFullYear(); // e.g., "Feb 2026"
    var type     = fullDay ? 'schoolsOut' : 'regular';

    // Initialise nested structure as needed
    if (!attendance[memberId]) {
      attendance[memberId] = { name: name, sites: {} };
    }
    var A = attendance[memberId];
    if (!A.sites[site])           A.sites[site] = {};
    if (!A.sites[site][monthKey]) A.sites[site][monthKey] = { regular: {}, schoolsOut: {} };

    // Store date key – object property assignment acts as a Set (de-duplicates)
    A.sites[site][monthKey][type][dateKey] = true;
  }

  generateOutputTabs_(ss, attendance);

  SpreadsheetApp.getUi().alert(
    'Attendance processing complete!\n' +
    'Check the new tabs in this spreadsheet.'
  );
}

// ── Output tab generation ────────────────────────────────────────────────────

/** Create or refresh all output tabs derived from the processed attendance map. */
function generateOutputTabs_(ss, attendance) {
  // Discover every (monthKey, site) pair present in the data
  var needed = {}; // tabName → { monthKey, site }

  for (var id in attendance) {
    var child = attendance[id];
    for (var site in child.sites) {
      for (var monthKey in child.sites[site]) {
        var mon     = monthKey.split(' ')[0];          // "Feb"
        var tabName = mon + ' - ' + site;              // "Feb - BFY"
        if (!needed[tabName]) {
          needed[tabName] = { monthKey: monthKey, site: site };
        }
      }
    }
  }

  // Sort tabs chronologically then alphabetically by site
  var tabNames = Object.keys(needed).sort(function(a, b) {
    var aKey = needed[a].monthKey;
    var bKey = needed[b].monthKey;
    // Compare "Mon YYYY" as sortable strings; year comes first for correct order
    var aSort = aKey.split(' ')[1] + '-' + CONFIG.months.indexOf(aKey.split(' ')[0]);
    var bSort = bKey.split(' ')[1] + '-' + CONFIG.months.indexOf(bKey.split(' ')[0]);
    return aSort.localeCompare(bSort) || a.localeCompare(b);
  });

  for (var t = 0; t < tabNames.length; t++) {
    var tabName = tabNames[t];
    var info    = needed[tabName];

    var sheet = ss.getSheetByName(tabName);
    if (sheet) {
      sheet.clearContents();
      sheet.clearFormats();
    } else {
      sheet = ss.insertSheet(tabName);
    }

    populateTab_(sheet, tabName, info.monthKey, info.site, attendance);
  }
}

/**
 * Write attendance data and formatting to a single output tab.
 *
 * Layout (rows):
 *   1  : Title
 *   2  : Subtitle (month | site)
 *   3  : Blank
 *   4  : Column headers
 *   5+ : One row per child / session type
 *   …  : Blank separator
 *   …  : Summary section (4 rows)
 */
function populateTab_(sheet, tabName, monthKey, site, attendance) {
  var siteFull = site === CONFIG.sites.BFY
    ? 'Belvidere Family Y'
    : 'Poplar Grove Elementary';

  // ── Collect child rows ──────────────────────────────────────────────────

  var childRows = [];

  for (var id in attendance) {
    var child = attendance[id];
    if (!child.sites[site] || !child.sites[site][monthKey]) continue;

    var m        = child.sites[site][monthKey];
    var regCount = Object.keys(m.regular).length;
    var fdCount  = Object.keys(m.schoolsOut).length;

    // Regular Before & After Care row
    if (regCount > 0) {
      childRows.push({
        display: child.name,
        type:    'regular',
        days:    regCount,
        sortKey: child.name
      });
    }

    // Schools Out (Full Day) row – "(FD)" suffix denotes Full Day care
    // Children with both types appear as two rows (e.g., Belvidere School District kids)
    if (fdCount > 0) {
      childRows.push({
        display: child.name + ' (FD)',
        type:    'schoolsOut',
        days:    fdCount,
        sortKey: child.name
      });
    }
  }

  // Sort alphabetically by name; for same child, regular row comes before FD row
  childRows.sort(function(a, b) {
    var n = a.sortKey.localeCompare(b.sortKey);
    return n !== 0 ? n : (a.type === 'regular' ? -1 : 1);
  });

  // ── Calculate summary totals ────────────────────────────────────────────

  var totalReg = 0, totalFD = 0;
  var uniqueReg = {}, uniqueFD = {}, allUnique = {};

  for (var c = 0; c < childRows.length; c++) {
    var cr = childRows[c];
    if (cr.type === 'regular') {
      totalReg += cr.days;
      uniqueReg[cr.sortKey] = true;
    } else {
      totalFD += cr.days;
      uniqueFD[cr.sortKey] = true;
    }
    allUnique[cr.sortKey] = true;
  }

  // ── Build output row array ──────────────────────────────────────────────

  var COLS      = 3;
  var DATA_ROW  = 5; // 1-based row where child data starts (after 3 header rows + 1 blank)
  var output    = [];

  // Rows 1–4
  output.push([tabName + '  \u2014  CCAP Attendance Summary', '', '']);  // 1: Title
  output.push([monthKey + '   |   ' + siteFull, '', '']);                // 2: Subtitle
  output.push(['', '', '']);                                              // 3: Blank
  output.push(['Participant', 'Session Type', 'Days Attended']);          // 4: Headers

  // Rows 5+ : child data
  for (var i = 0; i < childRows.length; i++) {
    var r = childRows[i];
    output.push([
      r.display,
      r.type === 'regular' ? 'Before & After Care' : 'Schools Out (Full Day)',
      r.days
    ]);
  }

  // Blank separator + summary (4 rows)
  var SUMMARY_ROW = DATA_ROW + childRows.length + 1; // 1-based, after blank

  output.push(['', '', '']);                          // blank separator
  output.push(['\u2014 SUMMARY \u2014', '', '']);     // "— SUMMARY —"
  output.push([
    'Before & After Care',
    'Children enrolled: ' + Object.keys(uniqueReg).length,
    'Total days: ' + totalReg
  ]);
  output.push([
    'Schools Out (Full Day)',
    'Children enrolled: ' + Object.keys(uniqueFD).length,
    'Total days: ' + totalFD
  ]);
  output.push([
    'Grand Total',
    'Children: ' + Object.keys(allUnique).length,
    'Total days: ' + (totalReg + totalFD)
  ]);

  // ── Write to sheet ──────────────────────────────────────────────────────

  sheet.getRange(1, 1, output.length, COLS).setValues(output);

  // ── Formatting ──────────────────────────────────────────────────────────

  var colors = CONFIG.colors;

  // Row 1: Title
  sheet.getRange(1, 1, 1, COLS)
    .merge()
    .setFontSize(13)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  // Row 2: Subtitle
  sheet.getRange(2, 1, 1, COLS)
    .merge()
    .setFontStyle('italic')
    .setHorizontalAlignment('center');

  // Row 4: Column headers
  sheet.getRange(4, 1, 1, COLS)
    .setFontWeight('bold')
    .setBackground(colors.headerBg)
    .setFontColor(colors.headerFg)
    .setHorizontalAlignment('center');

  // Child data rows – alternating row shading + Full Day highlight
  for (var j = 0; j < childRows.length; j++) {
    var rowNum = DATA_ROW + j;
    var cr2    = childRows[j];
    if (cr2.type === 'schoolsOut') {
      // Schools Out (FD) rows – distinct yellow background
      sheet.getRange(rowNum, 1, 1, COLS).setBackground(colors.fullDayBg);
    } else if (j % 2 === 1) {
      // Alternate regular rows with very light grey for readability
      sheet.getRange(rowNum, 1, 1, COLS).setBackground(colors.altRowBg);
    }
  }

  // Days Attended column – right-align numbers
  if (childRows.length > 0) {
    sheet.getRange(DATA_ROW, COLS, childRows.length, 1).setHorizontalAlignment('right');
  }

  // Summary section
  sheet.getRange(SUMMARY_ROW, 1, 5, COLS)
    .setFontWeight('bold')
    .setBackground(colors.summaryBg);

  // Grand Total row – slightly larger font
  sheet.getRange(SUMMARY_ROW + 4, 1, 1, COLS).setFontSize(11);

  // Column widths
  sheet.setColumnWidth(1, 230);  // Participant name (inc. " (FD)" suffix)
  sheet.setColumnWidth(2, 210);  // Session type
  sheet.setColumnWidth(3, 130);  // Days attended
}

// ── Custom menu ──────────────────────────────────────────────────────────────

/**
 * Adds a "CCAP Attendance" menu to the spreadsheet toolbar when the file is opened.
 * This is an Apps Script trigger – do not rename or remove it.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('CCAP Attendance')
    .addItem('Process Attendance Data', 'processAttendance')
    .addToUi();
}

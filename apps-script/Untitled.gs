const SMART_CALENDAR_CONFIG = {
  locationTabs: [
    "Eastvale", "La Verne", "Moreno Valley", "Redlands", "Rialto", "So Co", "Winery",
    "Azusa", "Mission Hills", "Northridge", "Porter Ranch", "Valencia", "West Covina",
    "Woodland Hills", "NoHo West", "Burbank", "Van Nuys", "Torrance", "Sunset Galleria",
    "Lakewood", "Marina Del Rey", "Monterey Park"
  ],
  baseDateRows: [4, 7, 10, 13, 16, 19],
  monthHeaderColumn: 3,
  monthHeaderRowsFallback: [1, 24, 46],
  monthSectionHeight: 21,
  calendarScanMaxRows: 65,
  debugTiming: true,
  firstScheduleColumn: 2,
  lastScheduleColumn: 8,
  scheduleRowOffsets: [1, 2],
  weekdayHours: 8,
  weekendHours: 6.5,
  yellowHexes: ["#ffff00", "#fff200", "#ffd966"],
  ignoredEntryPattern: /^(closed|close|dc needed|wc needed|needed|doctor\/dc|wc\/swc|n\/?a|na|tbd|none|-+)$/i
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Smart Calendar')
    .addItem('Refresh All Summaries', 'refreshSmartCalendarSummaries')
    .addSeparator()
    .addItem('Refresh Schedule Lookup', 'refreshScheduleLookup')
    .addItem('Refresh Coverage Summary', 'refreshCovSummary')
    .addItem('Refresh Hours Summary', 'refreshHrsSummary')
    .addItem('Refresh Over-Scheduled List', 'refreshOverSched')
    .addItem('Refresh Coverage Needed', 'refreshCoverageNeeded')
    .addToUi();
}

function refreshSmartCalendarSummaries() {
  const ss = SpreadsheetApp.getActive();
  const started = Date.now();
  Logger.log('Smart Calendar refresh started: ' + new Date(started).toISOString());
  const data = collectCalendarData_();
  logTiming_('Calendar data parsed once. Records=' + data.records.length + ', coverageRequests=' + data.coverageRequests.length, started);
  const steps = [
    {name: '_ScheduleData', fn: () => writeScheduleDataExtract_(data.records)},
    {name: 'ScheduleLookup', fn: () => writeScheduleLookup_(data.records)},
    {name: 'CovSummary', fn: () => writeCovSummary_(data.records)},
    {name: 'HrsSummary', fn: () => writeHrsSummary_(data.records)},
    {name: 'OverSched', fn: () => writeOverSched_(data.records)},
    {name: 'CoverageNeeded', fn: () => writeCoverageNeeded_(data.records, data.coverageRequests)}
  ];
  const errors = [];
  steps.forEach(step => {
    try {
      const stepStarted = Date.now();
      step.fn();
      logTiming_('Refresh completed: ' + step.name + ' in ' + elapsedSeconds_(stepStarted) + 's', started);
    } catch (err) {
      const message = step.name + ': ' + (err && err.message ? err.message : err);
      Logger.log('Refresh failed: ' + message);
      errors.push(message);
    }
  });
  const elapsed = elapsedSeconds_(started);
  if (errors.length) {
    ss.toast('Refresh finished with ' + errors.length + ' issue(s). Check Apps Script logs.', 'Smart Calendar', 8);
    throw new Error('Refresh All Summaries finished with errors: ' + errors.join(' | '));
  }
  ss.toast('All Smart Calendar summaries refreshed in ' + elapsed + 's.', 'Smart Calendar', 5);
}

function refreshScheduleLookup() {
  writeScheduleLookup_(collectScheduleRecords_());
}

function refreshCovSummary() {
  writeCovSummary_(collectScheduleRecords_());
}

function refreshHrsSummary() {
  writeHrsSummary_(collectScheduleRecords_());
}

function refreshOverSched() {
  writeOverSched_(collectScheduleRecords_());
}

function refreshCoverageNeeded() {
  const data = collectCalendarData_();
  writeCoverageNeeded_(data.records, data.coverageRequests);
}

function getAllScheduleData() {
  return collectCalendarData_();
}

function collectScheduleRecords_() {
  return collectCalendarData_().records;
}

function collectCalendarData_() {
  const ss = SpreadsheetApp.getActive();
  const records = [];
  const seenRecords = new Set();
  const coverageRequests = [];
  const parseStarted = Date.now();
  SMART_CALENDAR_CONFIG.locationTabs.forEach(tabName => {
    const tabStarted = Date.now();
    const sheet = ss.getSheetByName(tabName);
    if (!sheet) return;
    const colStart = SMART_CALENDAR_CONFIG.firstScheduleColumn;
    const colCount = SMART_CALENDAR_CONFIG.lastScheduleColumn - colStart + 1;
    const scan = getCalendarScan_(sheet);
    const sections = getCalendarSectionsFromDisplays_(scan.monthDisplays);
    if (!sections.length) return;
    const values = sheet.getRange(scan.rowStart, colStart, scan.rowCount, colCount).getValues();
    const displays = sheet.getRange(scan.rowStart, colStart, scan.rowCount, colCount).getDisplayValues();
    const backgrounds = sheet.getRange(scan.rowStart, colStart, scan.rowCount, colCount).getBackgrounds();

    sections.forEach(section => {
      section.dateRows.forEach(dateRow => {
        const dateIndex = dateRow - scan.rowStart;
        const dateValueRow = values[dateIndex];
        const dateDisplayRow = displays[dateIndex];

        // Defensive guard: prevents undefined row access if a month section is lower
        // than expected or a tab has extra spacing.
        if (!dateValueRow || !dateDisplayRow) {
          Logger.log('Skipped out-of-range date row on ' + tabName + ': row ' + dateRow);
          return;
        }

        for (let c = 0; c < colCount; c++) {
          const dateValue = normalizeDate_(dateValueRow[c], dateDisplayRow[c]);
          if (!dateValue || !dateInSectionMonth_(dateValue, section)) continue;
          SMART_CALENDAR_CONFIG.scheduleRowOffsets.forEach(offset => {
            const rIndex = dateRow + offset - scan.rowStart;
            const valueRow = values[rIndex];
            const displayRow = displays[rIndex];
            const backgroundRow = backgrounds[rIndex];

            // Defensive guard: prevents TypeError: Cannot read properties of undefined.
            if (!valueRow || !displayRow || !backgroundRow) {
              Logger.log('Skipped out-of-range schedule row on ' + tabName + ': row ' + (dateRow + offset));
              return;
            }

            const rawValue = valueRow[c];
            const displayValue = displayRow[c];
            const text = String(displayValue || rawValue || '').trim();
            const coverageType = coverageTypeFromText_(text);
            if (coverageType) coverageRequests.push({
              coverageType: coverageType,
              location: tabName,
              date: dateValue,
              sourceCell: tabName + '!' + columnToLetter_(colStart + c) + String(dateRow + offset),
              originalCellText: text
            });
            const names = splitNames_(displayValue || rawValue);
            if (!names.length) return;
            const a1 = columnToLetter_(colStart + c) + String(dateRow + offset);
            const yellow = isYellow_(backgroundRow[c]);
            names.forEach(name => {
              const key = normalizeName_(name) + '|' + tabName + '|' + dateKey_(dateValue) + '|' + a1;
              if (seenRecords.has(key)) return;
              seenRecords.add(key);
              records.push({
                name: name,
                location: tabName,
                date: dateValue,
                monthSection: monthNames_()[section.monthNumber - 1],
                sourceTab: tabName,
                sourceCell: tabName + '!' + a1,
                originalCellText: String(displayValue || rawValue || ''),
                background: backgroundRow[c],
                yellow: yellow
              });
            });
          });
        }
      });
    });
    logTiming_(tabName + ' parsed. Records so far=' + records.length, tabStarted);
  });
  logTiming_('collectCalendarData_ parsed all location tabs. Total records=' + records.length, parseStarted);
  return {records: records, coverageRequests: coverageRequests};
}

function getCalendarScan_(sheet) {
  const rowStart = 1;
  const maxDateBaseRow = Math.max.apply(null, SMART_CALENDAR_CONFIG.baseDateRows);
  const maxScheduleOffset = Math.max.apply(null, SMART_CALENDAR_CONFIG.scheduleRowOffsets);
  const lastFallbackHeader = Math.max.apply(null, SMART_CALENDAR_CONFIG.monthHeaderRowsFallback);
  const minimumNeededRow = lastFallbackHeader + maxDateBaseRow - 1 + maxScheduleOffset;
  const rowEnd = Math.max(SMART_CALENDAR_CONFIG.calendarScanMaxRows || 0, minimumNeededRow);
  const rowCount = rowEnd - rowStart + 1;

  return {
    rowStart: rowStart,
    rowCount: rowCount,
    monthDisplays: sheet.getRange(rowStart, SMART_CALENDAR_CONFIG.monthHeaderColumn, rowCount, 1).getDisplayValues()
  };
}

function getCalendarSectionsFromDisplays_(monthDisplays) {
  const monthNames = monthNames_();
  const sections = [];
  monthDisplays.forEach((row, index) => {
    const headerRow = index + 1;
    const monthNumber = monthNames.indexOf(String(row[0] || '').trim().toUpperCase()) + 1;
    if (!monthNumber) return;
    const dateRows = SMART_CALENDAR_CONFIG.baseDateRows.map(baseRow => headerRow + baseRow - 1);
    sections.push({headerRow: headerRow, monthNumber: monthNumber, dateRows: dateRows});
  });
  if (sections.length) return sections;
  return SMART_CALENDAR_CONFIG.monthHeaderRowsFallback.map(headerRow => {
    const displayRow = monthDisplays[headerRow - 1] || [''];
    const monthNumber = monthNames.indexOf(String(displayRow[0] || '').trim().toUpperCase()) + 1;
    return monthNumber ? {headerRow: headerRow, monthNumber: monthNumber, dateRows: SMART_CALENDAR_CONFIG.baseDateRows.map(baseRow => headerRow + baseRow - 1)} : null;
  }).filter(Boolean);
}

function dateInSectionMonth_(dateValue, section) {
  return !section.monthNumber || stripTime_(dateValue).getMonth() + 1 === section.monthNumber;
}

function monthNames_() {
  return ["JANUARY", "FEBRUARY", "MARCH", "APRIL", "MAY", "JUNE", "JULY", "AUGUST", "SEPTEMBER", "OCTOBER", "NOVEMBER", "DECEMBER"];
}

function writeScheduleDataExtract_(records) {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName('_ScheduleData');
  if (!sheet) sheet = ss.insertSheet('_ScheduleData');
  resetOutputSheet_(sheet);
  sheet.getRange('A1:F1').merge().setValue('Schedule Data Extract').setBackground('#334155').setFontColor('#ffffff').setFontWeight('bold').setFontSize(14);
  sheet.getRange('A2:F2').merge().setValue('Read-only extract from current workbook copy. Apps Script rebuilds equivalent data in Google Sheets; location tabs are not edited.').setFontColor('#475569').setFontStyle('italic');
  const rows = records
    .slice()
    .sort((a, b) => a.date - b.date || String(a.location).localeCompare(String(b.location)) || String(a.name).localeCompare(String(b.name)))
    .map(r => [r.name, r.location, r.date, r.sourceCell, r.originalCellText, r.yellow ? 'Yes' : 'No']);
  writeTableAt_(sheet, 4, ['Name', 'Location', 'Date', 'Source Cell', 'Original Cell Text', 'Yellow Highlight'], rows, '#334155');
  sheet.getRange(5, 3, Math.max(rows.length, 1), 1).setNumberFormat('yyyy-mm-dd');
  sheet.setFrozenRows(4);
  setSummaryColumnWidths_(sheet, [180, 150, 110, 130, 260, 130]);
}

function writeScheduleLookup_(records) {
  const rows = [];
  const seen = new Set();
  records.forEach(r => {
    const key = normalizeName_(r.name) + '|' + r.location + '|' + dateKey_(r.date);
    if (seen.has(key)) return;
    seen.add(key);
    rows.push([r.name, r.location, r.date]);
  });
  rows.sort((a, b) => a[2] - b[2] || String(a[1]).localeCompare(String(b[1])) || String(a[0]).localeCompare(String(b[0])));
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName('ScheduleLookup');
  if (!sheet) sheet = ss.insertSheet('ScheduleLookup');
  resetOutputSheet_(sheet);
  writeTableAt_(sheet, 1, ['Name', 'Location', 'Date'], rows, '#0f766e');
  sheet.getRange(2, 3, Math.max(rows.length, 1), 1).setNumberFormat('yyyy-mm-dd');
  sheet.setFrozenRows(1);
  setSummaryColumnWidths_(sheet, [180, 150, 110]);
  if (sheet.getFilter()) sheet.getFilter().remove();
  sheet.getRange(1, 1, Math.max(rows.length + 1, 2), 3).createFilter();
}

function writeCovSummary_(records) {
  const rows = [];
  const seen = new Set();
  records.filter(r => r.yellow).forEach(r => {
    const key = normalizeName_(r.name) + '|' + r.location + '|' + dateKey_(r.date);
    if (seen.has(key)) return;
    seen.add(key);
    rows.push([r.name, r.location, r.date]);
  });
  rows.sort((a, b) => a[2] - b[2] || String(a[1]).localeCompare(String(b[1])) || String(a[0]).localeCompare(String(b[0])));
  const sheet = prepareSummarySheet_('CovSummary', '#1d4ed8', 'Coverage Summary', 'Yellow-highlighted schedule cells only.');
  writeTable_(sheet, ['Name', 'Location', 'Date'], rows, '#1d4ed8');
  sheet.setFrozenRows(4);
  setSummaryColumnWidths_(sheet, [180, 150, 110]);
}

function writeHrsSummary_(records) {
  const byNameDate = new Map();
  records.forEach(r => {
    const key = normalizeName_(r.name) + '|' + dateKey_(r.date);
    if (!byNameDate.has(key)) byNameDate.set(key, {name: canonicalEmployeeName_(r.name), date: r.date, locations: new Set()});
    byNameDate.get(key).locations.add(r.location);
  });
  const dates = Array.from(byNameDate.values()).map(x => x.date);
  const anchor = dates.length ? mondayOfWeek_(new Date(Math.min.apply(null, dates.map(d => d.getTime())))) : mondayOfWeek_(new Date());
  const groups = new Map();
  byNameDate.forEach(item => {
    const periodStart = twoWeekPeriodStart_(item.date, anchor);
    const key = normalizeName_(item.name) + '|' + dateKey_(periodStart);
    if (!groups.has(key)) groups.set(key, {name: canonicalEmployeeName_(item.name), start: periodStart, end: addDays_(periodStart, 13), weekday: 0, weekend: 0, locations: new Set()});
    const g = groups.get(key);
    if (item.date.getDay() === 0 || item.date.getDay() === 6) g.weekend += 1; else g.weekday += 1;
    item.locations.forEach(loc => g.locations.add(loc));
  });
  const rows = Array.from(groups.values()).map(g => [g.name, g.start, g.end, g.weekday, g.weekend, g.weekday + g.weekend, (g.weekday * SMART_CALENDAR_CONFIG.weekdayHours) + (g.weekend * SMART_CALENDAR_CONFIG.weekendHours), Array.from(g.locations).sort().join(', ')]);
  rows.sort((a, b) => String(a[0]).localeCompare(String(b[0])) || a[1] - b[1]);
  const sheet = prepareSummarySheet_('HrsSummary', '#1d4ed8', 'Hours Summary', 'Estimated only. Weekday=8 hours, weekend=6.5 hours. Two-week anchor: ' + dateKey_(anchor) + '.');
  writeTable_(sheet, ['Name', '2-Week Period Start', '2-Week Period End', 'Weekday Scheduled Days', 'Weekend Scheduled Days', 'Total Scheduled Days', 'Estimated Hours', 'Locations Scheduled'], rows, '#1d4ed8');
  sheet.setFrozenRows(4);
  setSummaryColumnWidths_(sheet, [180, 130, 130, 150, 150, 140, 120, 320]);
}

function writeOverSched_(records) {
  const appearancesByNameDate = new Map();
  records.forEach(r => {
    const key = normalizeName_(r.name) + '|' + dateKey_(r.date);
    if (!appearancesByNameDate.has(key)) appearancesByNameDate.set(key, {name: canonicalEmployeeName_(r.name), date: r.date, entries: []});
    appearancesByNameDate.get(key).entries.push(r);
  });
  const weeks = new Map();
  appearancesByNameDate.forEach(item => {
    const weekStart = mondayOfWeek_(item.date);
    const key = normalizeName_(item.name) + '|' + dateKey_(weekStart);
    if (!weeks.has(key)) weeks.set(key, {name: canonicalEmployeeName_(item.name), weekStart: weekStart, dates: new Map()});
    const dateKey = dateKey_(item.date);
    if (!weeks.get(key).dates.has(dateKey)) weeks.get(key).dates.set(dateKey, new Set());
    item.entries.forEach(entry => weeks.get(key).dates.get(dateKey).add(entry.location));
  });
  const rows = [];
  weeks.forEach(w => {
    let best = [], current = [];
    for (let i = 0; i < 7; i++) {
      const d = addDays_(w.weekStart, i);
      if (w.dates.has(dateKey_(d))) {
        current.push(d);
        if (current.length > best.length) best = current.slice();
      } else {
        current = [];
      }
    }
    if (best.length >= 6) {
      const locs = new Set();
      best.forEach(d => w.dates.get(dateKey_(d)).forEach(loc => locs.add(loc)));
      rows.push(['6-Day Rule', w.name, w.weekStart, addDays_(w.weekStart, 6), best.length, best.map(dateKey_).join(', '), Array.from(locs).sort().join(', '), 'Scheduled 6 days in a row within the same Monday-Sunday week.']);
    }
  });
  appearancesByNameDate.forEach(item => {
    const locationSet = new Set(item.entries.map(e => e.location));
    if (locationSet.size <= 1) return;
    rows.push([
      'Same-Day Conflict',
      item.name,
      item.date,
      item.date,
      locationSet.size,
      dateKey_(item.date),
      Array.from(locationSet).sort().join(', '),
      sameDayWarning_(locationSet.size)
    ]);
  });
  rows.sort((a, b) => a[2] - b[2] || String(a[1]).localeCompare(String(b[1])) || String(a[0]).localeCompare(String(b[0])));
  const sheet = prepareSummarySheet_('OverSched', '#b91c1c', 'Over-Scheduled List', 'Flags 6-day Monday-Sunday violations and same-day multi-location scheduling conflicts.');
  writeTable_(sheet, ['Issue Type', 'Name', 'Date or Week Start Date', 'Week End Date', 'Count', 'Scheduled Dates', 'Locations Involved', 'Issue / Warning'], rows, '#b91c1c');
  if (rows.length) sheet.getRange(5, 1, rows.length, 8).setBackground('#fee2e2');
  sheet.setFrozenRows(4);
  setSummaryColumnWidths_(sheet, [150, 180, 160, 130, 80, 260, 320, 360]);
}

function writeCoverageNeeded_(records, coverageRequests) {
  const started = Date.now();
  const requests = coverageRequests || [];
  const staff = readStaffConfig_();
  const coverageContext = buildCoverageCandidateContext_(records);
  const rows = requests
    .slice()
    .sort((a, b) => a.date - b.date || String(a.location).localeCompare(String(b.location)) || String(a.coverageType).localeCompare(String(b.coverageType)))
    .map(req => {
      const scheduledNames = coverageContext.scheduledByDate.get(dateKey_(req.date)) || new Set();
      const candidates = staff
        .filter(person => person.role === req.coverageType)
        .filter(person => !scheduledNames.has(normalizeName_(person.name)))
        .filter(person => !wouldCreateSixDayViolationFast_(person.name, req.date, coverageContext.scheduledByNameWeek))
        .map(person => person.name)
        .sort();
      return [req.coverageType, req.location, req.date, candidates.join(', ')];
    });
  logTiming_('CoverageNeeded candidate list built. Requests=' + requests.length + ', staff=' + staff.length, started);
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName('CoverageNeeded');
  if (!sheet) sheet = ss.insertSheet('CoverageNeeded');
  resetOutputSheet_(sheet);
  writeTableAt_(sheet, 1, ['Coverage Type', 'Location', 'Date', 'Potential Candidates'], rows, '#7c3aed');
  sheet.getRange(2, 3, Math.max(rows.length, 1), 1).setNumberFormat('yyyy-mm-dd');
  sheet.setFrozenRows(1);
  setSummaryColumnWidths_(sheet, [120, 150, 110, 600]);
  if (sheet.getFilter()) sheet.getFilter().remove();
  sheet.getRange(1, 1, Math.max(rows.length + 1, 2), 4).createFilter();
}

function buildCoverageCandidateContext_(records) {
  const scheduledByDate = new Map();
  const scheduledByNameWeek = new Map();
  records.forEach(r => {
    const dKey = dateKey_(r.date);
    const normalized = normalizeName_(r.name);
    if (!scheduledByDate.has(dKey)) scheduledByDate.set(dKey, new Set());
    scheduledByDate.get(dKey).add(normalized);

    const weekKey = normalized + '|' + dateKey_(mondayOfWeek_(r.date));
    if (!scheduledByNameWeek.has(weekKey)) scheduledByNameWeek.set(weekKey, new Set());
    scheduledByNameWeek.get(weekKey).add(dKey);
  });
  return {scheduledByDate: scheduledByDate, scheduledByNameWeek: scheduledByNameWeek};
}


function resetOutputSheet_(sheet) {
  // Remove stateful sheet decorations before rebuilding script-controlled output.
  const filter = sheet.getFilter();
  if (filter) filter.remove();

  sheet.getBandings().forEach(banding => banding.remove());
  safeUnmergeSheet_(sheet);

  sheet.clear({contentsOnly: false});
  sheet.clearConditionalFormatRules();
  sheet.setFrozenRows(0);
}

function safeUnmergeSheet_(sheet) {
  try {
    const fullSheet = sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns());
    const mergedRanges = fullSheet.getMergedRanges();
    mergedRanges.forEach(range => {
      try {
        range.breakApart();
      } catch (err) {
        Logger.log('Failed to unmerge ' + sheet.getName() + '!' + range.getA1Notation() + ': ' + err);
      }
    });
    if (mergedRanges.length) Logger.log('Unmerged ' + mergedRanges.length + ' merged range(s) on ' + sheet.getName() + '.');
  } catch (err) {
    Logger.log('Unable to inspect merged ranges on ' + sheet.getName() + ': ' + err);
  }
}

function prepareSummarySheet_(name, tabColor, title, subtitle) {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  resetOutputSheet_(sheet);
  sheet.setTabColor(tabColor);
  sheet.getRange('A1:H1').merge().setValue(title).setBackground(tabColor).setFontColor('#ffffff').setFontWeight('bold').setFontSize(16);
  sheet.getRange('A2:H2').merge().setValue(subtitle).setFontColor('#475569').setFontStyle('italic');
  sheet.setHiddenGridlines(true);
  return sheet;
}

function writeTable_(sheet, headers, rows, color) {
  writeTableAt_(sheet, 4, headers, rows, color);
}

function writeTableAt_(sheet, headerRow, headers, rows, color) {
  if (!rows.length) rows = [['No records found.'].concat(new Array(headers.length - 1).fill(''))];
  sheet.getRange(headerRow, 1, 1, headers.length).setValues([headers]).setBackground(color).setFontColor('#ffffff').setFontWeight('bold').setWrap(true);
  const bodyRange = sheet.getRange(headerRow + 1, 1, rows.length, headers.length);
  bodyRange.setValues(rows).setWrap(true).setBorder(true, true, true, true, true, true, '#cbd5e1', SpreadsheetApp.BorderStyle.SOLID);
  applyStaticRowShading_(bodyRange, rows.length, headers.length);
  sheet.getRange(headerRow, 1, 1, headers.length).setBorder(true, true, true, true, true, true, '#cbd5e1', SpreadsheetApp.BorderStyle.SOLID);
  for (let c = 1; c <= headers.length; c++) if (/date/i.test(headers[c - 1]) || /period/i.test(headers[c - 1])) sheet.getRange(headerRow + 1, c, rows.length, 1).setNumberFormat('yyyy-mm-dd');
}

function applyStaticRowShading_(range, rowCount, colCount) {
  const backgrounds = [];
  for (let r = 0; r < rowCount; r++) {
    const color = r % 2 === 0 ? '#ffffff' : '#f8fafc';
    backgrounds.push(new Array(colCount).fill(color));
  }
  range.setBackgrounds(backgrounds);
}

function setSummaryColumnWidths_(sheet, widths) {
  widths.forEach((width, index) => sheet.setColumnWidth(index + 1, width));
}

function splitNames_(value) {
  if (value === null || value === undefined) return [];
  const text = String(value).trim();
  if (!text || text.charAt(0) === '=') return [];
  return text.split(/\s*(?:\/|,|;|\n|\band\b|&)\s*/i)
    .map(s => s.replace(/\s+/g, ' ').trim())
    .filter(s => s && !SMART_CALENDAR_CONFIG.ignoredEntryPattern.test(s) && !/\bneeded\b|\bclosed\b/i.test(s));
}

function normalizeDate_(value, display) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) return stripTime_(value);
  const parsed = new Date(display || value);
  return isNaN(parsed) ? null : stripTime_(parsed);
}

function isYellow_(hex) {
  const h = String(hex || '').toLowerCase();
  if (SMART_CALENDAR_CONFIG.yellowHexes.indexOf(h) >= 0) return true;
  const m = h.match(/^#?([0-9a-f]{6})$/i);
  if (!m) return false;
  const r = parseInt(m[1].slice(0, 2), 16), g = parseInt(m[1].slice(2, 4), 16), b = parseInt(m[1].slice(4, 6), 16);
  return r >= 230 && g >= 210 && b <= 80;
}

function coverageTypeFromText_(text) {
  const t = String(text || '').toLowerCase();
  if (/\bwc\b.*\bneeded\b|\bneeded\b.*\bwc\b/.test(t)) return 'WC';
  if (/\bdc\b.*\bneeded\b|\bneeded\b.*\bdc\b/.test(t)) return 'DC';
  return '';
}

function readStaffConfig_() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('_StaffConfig');
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, 3).getValues()
    .map(row => ({name: String(row[0] || '').trim(), role: String(row[1] || '').trim().toUpperCase(), active: String(row[2]).toLowerCase() !== 'false'}))
    .filter(person => person.name && person.active && (person.role === 'WC' || person.role === 'DC'));
}

function wouldCreateSixDayViolation_(name, proposedDate, records) {
  const targetName = normalizeName_(name);
  const weekStart = mondayOfWeek_(proposedDate);
  const scheduled = new Set([dateKey_(proposedDate)]);
  records.forEach(r => {
    if (normalizeName_(r.name) !== targetName) return;
    if (dateKey_(mondayOfWeek_(r.date)) !== dateKey_(weekStart)) return;
    scheduled.add(dateKey_(r.date));
  });
  let current = 0;
  for (let i = 0; i < 7; i++) {
    current = scheduled.has(dateKey_(addDays_(weekStart, i))) ? current + 1 : 0;
    if (current >= 6) return true;
  }
  return false;
}

function wouldCreateSixDayViolationFast_(name, proposedDate, scheduledByNameWeek) {
  const targetName = normalizeName_(name);
  const weekStart = mondayOfWeek_(proposedDate);
  const scheduled = new Set(scheduledByNameWeek.get(targetName + '|' + dateKey_(weekStart)) || []);
  scheduled.add(dateKey_(proposedDate));
  let current = 0;
  for (let i = 0; i < 7; i++) {
    current = scheduled.has(dateKey_(addDays_(weekStart, i))) ? current + 1 : 0;
    if (current >= 6) return true;
  }
  return false;
}

function formatLocationCounts_(entries) {
  const counts = new Map();
  entries.forEach(entry => counts.set(entry.location, (counts.get(entry.location) || 0) + 1));
  return Array.from(counts.keys()).sort().map(loc => counts.get(loc) > 1 ? loc + ' (' + counts.get(loc) + ')' : loc).join(', ');
}

function sameDayWarning_(count) {
  if (count === 2) return 'Double scheduled on the same date across multiple locations.';
  if (count === 3) return 'Triple scheduled on the same date across multiple locations.';
  if (count === 4) return 'Scheduled 4 times on the same date across multiple locations.';
  return 'Scheduled ' + count + ' times on the same date across multiple locations.';
}

function canonicalEmployeeName_(name) { return String(name || '').replace(/\([^)]*\)/g, ' ').replace(/\boffice\s*day\b/gi, ' ').replace(/\b(?:morning|afternoon|am|pm)\b/gi, ' ').replace(/\b\d{1,2}(?::\d{2})?\s*(?:-|to)\s*\d{1,2}(?::\d{2})?\b/gi, ' ').replace(/\s+/g, ' ').trim(); }
function normalizeName_(name) { return canonicalEmployeeName_(name).toLowerCase().replace(/\s+/g, ' ').trim(); }
function stripTime_(d) { return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }
function dateKey_(d) { const x = stripTime_(d); return x.getFullYear() + '-' + pad2_(x.getMonth() + 1) + '-' + pad2_(x.getDate()); }
function pad2_(n) { return n < 10 ? '0' + n : String(n); }
function addDays_(d, days) { const x = stripTime_(d); x.setDate(x.getDate() + days); return x; }
function mondayOfWeek_(d) { const x = stripTime_(d); const offset = (x.getDay() + 6) % 7; x.setDate(x.getDate() - offset); return x; }
function twoWeekPeriodStart_(d, anchor) { const days = Math.floor((stripTime_(d) - stripTime_(anchor)) / 86400000); return addDays_(anchor, Math.floor(days / 14) * 14); }
function elapsedSeconds_(started) { return Math.round((Date.now() - started) / 100) / 10; }
function logTiming_(message, started) {
  if (!SMART_CALENDAR_CONFIG.debugTiming) return;
  Logger.log(message + ' at +' + elapsedSeconds_(started) + 's.');
}
function columnToLetter_(column) {
  let letter = '';
  while (column > 0) {
    const remainder = (column - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    column = Math.floor((column - 1) / 26);
  }
  return letter;
}
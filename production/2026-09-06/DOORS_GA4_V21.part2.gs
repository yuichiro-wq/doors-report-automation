      channel.totalSessions, '', DOORS_LOG_CFG_V21.channelDefinitionVersion,
      'GA4_国別アクティブユーザー月次推移', DOORS_LOG_CFG_V21.rawChannelSheet, channelStatus,
      'same-run total canonical; breakdown=' + channel.breakdownSum + '; diff=' + channel.gap +
        ' (' + (channel.gapRatio * 100).toFixed(2) + '%)'
    ]
  ];
  doorsAssertRowsWidthV21_(rows, DOORS_HEADERS_V21.mainAggregate.length, 'main aggregate');
  doorsReplaceRowsAppendThenDeleteByPredicateV21_(
    DOORS_LOG_CFG_V21.mainSpreadsheetId,
    DOORS_LOG_CFG_V21.mainAggregateSheet,
    aggregatePredicate,
    rows
  );
}

function doorsBuildChannelRawRowV21_(x) {
  return [
    x.month, x.channel, x.sessions, DOORS_LOG_CFG_V21.country, DOORS_LOG_CFG_V21.landingFilterLabel,
    x.dimension, 'sessions', DOORS_LOG_CFG_V21.channelDefinitionVersion, x.runId, x.fetchedAt,
    x.dateStart, x.dateEnd, 'GA4 Data API', x.status
  ];
}

function doorsBuildAiRawRowV21_(x) {
  return [
    x.month, x.service, x.source, x.medium, x.sessions, x.totalSessions,
    x.totalSessions ? x.sessions / x.totalSessions : 0,
    DOORS_LOG_CFG_V21.country, DOORS_LOG_CFG_V21.landingFilterLabel,
    DOORS_LOG_CFG_V21.aiDefinitionVersion, x.runId, x.fetchedAt,
    x.dateStart, x.dateEnd, 'GA4 Data API', x.status
  ];
}

function doorsAppendAuditV21_(x) {
  const row = [[
    x.runId,
    x.executedAt || new Date(),
    x.targetMonth || '',
    x.dataset || '',
    x.definitionVersion || DOORS_LOG_CFG_V21.runAuditDefinitionVersion,
    x.dateStart || '',
    x.dateEnd || '',
    Number(x.rowCount || 0),
    x.status || '',
    x.errorType || '',
    x.errorMessage || '',
    x.sourceSheet || '',
    x.destinationSheet || '',
    x.note || ''
  ]];
  doorsAssertRowsWidthV21_(row, DOORS_HEADERS_V21.audit.length, 'audit');
  doorsAppendRowsV21_(DOORS_LOG_CFG_V21.ga4LogSpreadsheetId, DOORS_LOG_CFG_V21.runLogSheet, row);
}

function doorsClassifyAiSourceV21_(source, medium) {
  const s = String(source || '').toLowerCase();
  const m = String(medium || '').toLowerCase();
  const allowedMedia = {'ai-assistant': true, '(not set)': true, 'referral': true};
  if (!allowedMedia[m]) return null;
  const rules = [
    {service: 'ChatGPT / OpenAI', re: /(^openai$|(^|\.)(chatgpt\.com|chat\.openai\.com|openai\.com)$)/},
    {service: 'Gemini / Bard', re: /(^|\.)(gemini\.google\.com|bard\.google\.com)$/},
    {service: 'Perplexity', re: /(^|\.)perplexity\.ai$/},
    {service: 'Claude', re: /(^|\.)claude\.ai$/},
    {service: 'NotebookLM', re: /(^|\.)notebooklm\.google\.com$/},
    {service: 'Copilot', re: /(^|\.)(copilot\.com|copilot\.microsoft\.com|copilot\.cloud\.microsoft)$/}
  ];
  for (var i = 0; i < rules.length; i++) {
    if (rules[i].re.test(s)) return {service: rules[i].service};
  }
  return null;
}

function doorsJapanDoorsFilterV21_() {
  return {
    andGroup: {
      expressions: [
        {filter: {fieldName: 'country', stringFilter: {matchType: 'EXACT', value: DOORS_LOG_CFG_V21.country, caseSensitive: false}}},
        {filter: {fieldName: DOORS_LOG_CFG_V21.landingField, stringFilter: {matchType: 'CONTAINS', value: DOORS_LOG_CFG_V21.landingContains, caseSensitive: false}}}
      ]
    }
  };
}

function doorsPreflightHeadersV21_() {
  doorsAssertSheetHeaderV21_(DOORS_LOG_CFG_V21.ga4LogSpreadsheetId, DOORS_LOG_CFG_V21.rawChannelSheet, DOORS_HEADERS_V21.channel);
  doorsAssertSheetHeaderV21_(DOORS_LOG_CFG_V21.ga4LogSpreadsheetId, DOORS_LOG_CFG_V21.rawAiSheet, DOORS_HEADERS_V21.ai);
  doorsAssertSheetHeaderV21_(DOORS_LOG_CFG_V21.ga4LogSpreadsheetId, DOORS_LOG_CFG_V21.runLogSheet, DOORS_HEADERS_V21.audit);
  doorsAssertSheetHeaderV21_(DOORS_LOG_CFG_V21.ga4LogSpreadsheetId, DOORS_LOG_CFG_V21.gateSheet, DOORS_HEADERS_V21.gate);
  doorsAssertSheetHeaderV21_(DOORS_LOG_CFG_V21.mainSpreadsheetId, DOORS_LOG_CFG_V21.mainChannelSheet, DOORS_HEADERS_V21.mainChannel);
  doorsAssertSheetHeaderV21_(DOORS_LOG_CFG_V21.mainSpreadsheetId, DOORS_LOG_CFG_V21.mainAggregateSheet, DOORS_HEADERS_V21.mainAggregate);
  return {status: 'PASS', version: 'v2.1'};
}

function runDoorsLogPreflightV21() {
  return doorsPreflightHeadersV21_();
}

function doorsAssertSheetHeaderV21_(spreadsheetId, sheetName, expected) {
  const sh = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName);
  if (!sh) throw new Error('Missing sheet: ' + sheetName);
  if (sh.getLastColumn() < expected.length) {
    throw new Error('Header width mismatch on ' + sheetName + ': expected at least ' + expected.length + ', got ' + sh.getLastColumn());
  }
  const actual = sh.getRange(1, 1, 1, expected.length).getDisplayValues()[0];
  for (var i = 0; i < expected.length; i++) {
    if (String(actual[i]) !== String(expected[i])) {
      throw new Error('Header mismatch on ' + sheetName + ' col ' + (i + 1) + ': expected=' + expected[i] + ', actual=' + actual[i]);
    }
  }
}

function doorsHeaderMapV21_(header) {
  const map = {};
  header.forEach(function(v, i) { map[String(v)] = i; });
  return map;
}


function doorsReplaceRowsAppendThenDeleteByFieldsV21_(spreadsheetId, sheetName, criteria, newRows) {
  return doorsReplaceRowsAppendThenDeleteByPredicateV21_(spreadsheetId, sheetName, function(obj) {
    const keys = Object.keys(criteria);
    for (var i = 0; i < keys.length; i++) {
      const k = keys[i];
      const actual = k.indexOf('month') >= 0 ? doorsCellMonthV21_(obj[k]) : String(obj[k] == null ? '' : obj[k]);
      const expected = k.indexOf('month') >= 0 ? doorsCellMonthV21_(criteria[k]) : String(criteria[k] == null ? '' : criteria[k]);
      if (actual !== expected) return false;
    }
    return true;
  }, newRows);
}

function doorsReplaceRowsAppendThenDeleteByPredicateV21_(spreadsheetId, sheetName, predicate, newRows) {
  const sh = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName);
  if (!sh) throw new Error('Missing sheet: ' + sheetName);
  const lastRowBefore = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  const oldRows = [];
  if (lastRowBefore > 1 && lastCol > 0) {
    const values = sh.getRange(1, 1, lastRowBefore, lastCol).getValues();
    const header = values[0].map(String);
    for (var i = 1; i < values.length; i++) {
      const obj = {};
      for (var c = 0; c < header.length; c++) obj[header[c]] = values[i][c];
      if (predicate(obj)) oldRows.push(i + 1);
    }
  }
  // Append first. If append fails, old rows remain untouched.
  doorsAppendRowsV21_(spreadsheetId, sheetName, newRows);
  // Delete only rows that existed before the append.
  for (var j = oldRows.length - 1; j >= 0; j--) sh.deleteRow(oldRows[j]);
  return {appended: (newRows || []).length, deletedOld: oldRows.length};
}

function doorsDeleteRowsByFieldsV21_(spreadsheetId, sheetName, criteria) {
  doorsDeleteRowsByPredicateV21_(spreadsheetId, sheetName, function(obj) {
    const keys = Object.keys(criteria);
    for (var i = 0; i < keys.length; i++) {
      const k = keys[i];
      const actual = k.indexOf('month') >= 0 ? doorsCellMonthV21_(obj[k]) : String(obj[k] == null ? '' : obj[k]);
      const expected = k.indexOf('month') >= 0 ? doorsCellMonthV21_(criteria[k]) : String(criteria[k] == null ? '' : criteria[k]);
      if (actual !== expected) return false;
    }
    return true;
  });
}

function doorsDeleteRowsByPredicateV21_(spreadsheetId, sheetName, predicate) {
  const sh = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName);
  if (!sh) throw new Error('Missing sheet: ' + sheetName);
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow <= 1 || lastCol <= 0) return 0;
  const values = sh.getRange(1, 1, lastRow, lastCol).getValues();
  const header = values[0].map(String);
  const deleteRows = [];
  for (var i = 1; i < values.length; i++) {
    const obj = {};
    for (var c = 0; c < header.length; c++) obj[header[c]] = values[i][c];
    if (predicate(obj)) deleteRows.push(i + 1);
  }
  for (var j = deleteRows.length - 1; j >= 0; j--) sh.deleteRow(deleteRows[j]);
  return deleteRows.length;
}

function doorsAppendRowsV21_(spreadsheetId, sheetName, rows) {
  if (!rows || !rows.length) return;
  const sh = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName);
  if (!sh) throw new Error('Missing sheet: ' + sheetName);
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
}

function doorsUpsertByHeaderKeyV21_(spreadsheetId, sheetName, keyHeader, keyValue, row, expectedHeader) {
  doorsAssertSheetHeaderV21_(spreadsheetId, sheetName, expectedHeader);
  const sh = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName);
  const header = sh.getRange(1, 1, 1, expectedHeader.length).getDisplayValues()[0];
  const map = doorsHeaderMapV21_(header);
  const keyIndex = map[keyHeader];
  if (keyIndex == null) throw new Error('Missing key header ' + keyHeader + ' on ' + sheetName);
  const lastRow = sh.getLastRow();
  if (lastRow > 1) {
    const vals = sh.getRange(2, keyIndex + 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < vals.length; i++) {
      if (doorsCellMonthV21_(vals[i][0]) === doorsCellMonthV21_(keyValue)) {
        sh.getRange(i + 2, 1, 1, row.length).setValues([row]);
        return;
      }
    }
  }
  sh.getRange(lastRow + 1, 1, 1, row.length).setValues([row]);
}

function doorsKnownTotalSessionsV21_(reportMonth) {
  const sh = SpreadsheetApp.openById(DOORS_LOG_CFG_V21.ga4LogSpreadsheetId).getSheetByName(DOORS_LOG_CFG_V21.knownTotalSheet);
  if (!sh) return null;
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2) return null;
  const values = sh.getRange(1, 1, lastRow, Math.min(lastCol, 10)).getValues();
  var headerRow = -1;
  var monthCol = -1;
  var sessionsCol = -1;
  for (var r = 0; r < Math.min(values.length, 20); r++) {
    for (var c = 0; c < values[r].length; c++) {
      const s = String(values[r][c] || '').trim();
      if (s === '年月' || s === 'month' || s === 'report_month') {
        headerRow = r;
        monthCol = c;
      }
      if (headerRow === r && (s === 'セッション' || s === 'sessions')) sessionsCol = c;
    }
    if (headerRow >= 0 && monthCol >= 0 && sessionsCol >= 0) break;
  }
  if (headerRow < 0 || monthCol < 0 || sessionsCol < 0) return null;
  for (var i = headerRow + 1; i < values.length; i++) {
    if (doorsCellMonthV21_(values[i][monthCol]) === reportMonth) {
      const n = Number(String(values[i][sessionsCol]).replace(/,/g, ''));
      return isNaN(n) ? null : n;
    }
  }
  return null;
}

function doorsExtractMonthMetricV21_(rows, reportMonth, monthDimIndex, metricIndex) {
  for (var i = 0; i < rows.length; i++) {
    const ym = doorsYearMonthToHyphenV21_(doorsDimV21_(rows[i], monthDimIndex));
    if (ym === reportMonth) return doorsMetricNumberV21_(rows[i], metricIndex);
  }
  return null;
}

function doorsAssertRowsWidthV21_(rows, expectedWidth, label) {
  (rows || []).forEach(function(row, i) {
    if (!Array.isArray(row) || row.length !== expectedWidth) {
      throw new Error(label + ' row width mismatch at index ' + i + ': expected=' + expectedWidth + ', actual=' + (row && row.length));
    }
  });
}

function doorsMonthRangeV21_(month) {
  const p = month.split('-');
  const y = Number(p[0]);
  const m = Number(p[1]);
  const start = new Date(y, m - 1, 1);
  const end = new Date(y, m, 0);
  return {
    startDate: Utilities.formatDate(start, 'Asia/Tokyo', 'yyyy-MM-dd'),
    endDate: Utilities.formatDate(end, 'Asia/Tokyo', 'yyyy-MM-dd')
  };
}

function doorsPrevMonthV21_() {
  const now = new Date();
  const d = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM');
}

function doorsAssertMonthV21_(month) {
  if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(String(month))) {
    throw new Error('reportMonth must be YYYY-MM: ' + month);
  }
}

function doorsYearMonthToHyphenV21_(v) {
  const s = String(v || '');
  return s.length === 6 ? s.slice(0, 4) + '-' + s.slice(4, 6) : s;
}

function doorsCellMonthV21_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy-MM');
  const s = String(v || '');
  const m = s.match(/^(\d{4})[-\/]?(\d{1,2})/);
  if (!m) return s;
  return m[1] + '-' + ('0' + m[2]).slice(-2);
}

function doorsDimV21_(row, index) {
  return row && row.dimensionValues && row.dimensionValues[index] ? row.dimensionValues[index].value : '';
}

function doorsMetricNumberV21_(row, index) {
  const v = row && row.metricValues && row.metricValues[index] ? row.metricValues[index].value : '0';
  return Number(v || 0);
}

/** Re-runnable backfill. Not called automatically. */
function backfillDoorsLogs2025To202607V21() {
  const months = [];
  var d = new Date(2025, 0, 1);
  const end = new Date(2026, 6, 1);
  while (d <= end) {
    months.push(Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM'));
    d = new Date(d.getFullYear(), d.getMonth() + 1, 1);
  }
  return months.map(function(month) { return runDoorsLogAcquisitionV21(month); });
}


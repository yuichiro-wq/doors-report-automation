/**
 * DOORS Production Scheduler 1970 Hotfix - v1.4 REVIEWED
 *
 * Exact Apps Script filename:
 *   DOORS_MONTHLY_SCHEDULER_HOTFIX_V14_REVIEWED.gs
 *
 * Goal:
 * - Fix the Production 1970-01 scheduled-month incident without touching
 *   renderer logic, article population logic, slide IDs/order, or LLMO logic.
 *
 * Safety model:
 * 1) READ-ONLY diagnostic
 * 2) READ-ONLY deterministic previous-month test in Asia/Tokyo
 * 3) WRITE only B3 repair
 * 4) trigger migration = create candidate first -> verify -> remove old -> verify
 * 5) scheduled handler computes month independently in Asia/Tokyo and passes the
 *    explicit validated month into runDoorsMonthlyReportV21(month), then runs
 *    insights/keyword sync explicitly. It does NOT depend on the live V12 wrapper
 *    to resolve the report month.
 *
 * Production IDs only.
 */

var DOORS_SCHEDULER_V14 = Object.freeze({
  spreadsheetId: '1q1B0JGdLaBOCkDWa4yLNlSFozzf8eaky2sTvTL9BAgo',
  panelSheet: '運用パネル',
  timezone: 'Asia/Tokyo',

  newHandler: 'doorsScheduledMonthlyRunV14Reviewed',

  legacyHandlers: Object.freeze([
    'doorsScheduledMonthlyRunV1',
    'doorsScheduledMonthlyRunWithInsightsV1',
    'doorsScheduledMonthlyRunWithInsightsV12'
  ]),

  panel: Object.freeze({
    targetMonth: 'B3',
    status: 'B6',
    executedAt: 'B7',
    keywordPublished: 'B8',
    day: 'B10',
    hour: 'B11',
    triggerState: 'B12'
  })
});


/**
 * READ ONLY.
 * Diagnoses current Production scheduler prerequisites.
 */
function diagnoseDoorsMonthlySchedulerV14Reviewed() {
  var cfg = DOORS_SCHEDULER_V14;
  var ss = SpreadsheetApp.openById(cfg.spreadsheetId);
  var panel = ss.getSheetByName(cfg.panelSheet);

  if (!panel) {
    throw new Error('運用パネルが見つかりません。');
  }

  var nowTokyo = doorsSchedulerTokyoPartsV14_(new Date());
  var computedMonth = doorsSchedulerPreviousCompleteMonthFromPartsV14_(
    nowTokyo.year,
    nowTokyo.month
  );

  doorsSchedulerAssertMonthV14_(computedMonth);

  var day = Number(panel.getRange(cfg.panel.day).getValue());
  var hour = Number(panel.getRange(cfg.panel.hour).getValue());

  if (!Number.isInteger(day) || day < 1 || day > 28) {
    throw new Error('自動実行日は1〜28で指定してください: ' + day);
  }

  if (!Number.isInteger(hour) || hour < 0 || hour > 23) {
    throw new Error('自動実行時刻は0〜23で指定してください: ' + hour);
  }

  var checks = {
    runDoorsMonthlyReportV21:
      typeof runDoorsMonthlyReportV21 === 'function',
    syncDoorsMonthlyInsightsV12:
      typeof syncDoorsMonthlyInsightsV12 === 'function',
    syncDoorsKeywordProposalsToSlidesV1:
      typeof syncDoorsKeywordProposalsToSlidesV1 === 'function',
    finalizeDoorsMonthlyProductionV16:
      typeof finalizeDoorsMonthlyProductionV16 === 'function'
  };

  var missing = Object.keys(checks).filter(function(k) {
    return !checks[k];
  });

  var triggers = ScriptApp.getProjectTriggers().map(function(t) {
    return {
      handler: t.getHandlerFunction(),
      eventType: String(t.getEventType()),
      triggerSource: String(t.getTriggerSource()),
      uniqueId: String(t.getUniqueId ? t.getUniqueId() : '')
    };
  });

  var monthlyTriggers = triggers.filter(function(t) {
    return (
      cfg.legacyHandlers.indexOf(t.handler) >= 0 ||
      t.handler === cfg.newHandler
    );
  });

  return {
    status: missing.length ? 'BLOCK' : 'PASS',
    version: '1.4-reviewed',
    currentTokyo: nowTokyo,
    computedTargetMonth: computedMonth,
    currentPanelTargetMonth: String(
      panel.getRange(cfg.panel.targetMonth).getDisplayValue() || ''
    ).trim(),
    currentPanelStatus: String(
      panel.getRange(cfg.panel.status).getDisplayValue() || ''
    ).trim(),
    day: day,
    hour: hour,
    timezone: cfg.timezone,
    newHandler: cfg.newHandler,
    checks: checks,
    missing: missing,
    monthlyTriggers: monthlyTriggers
  };
}


/**
 * READ ONLY.
 * Deterministic month arithmetic unit tests.
 * No dependency on current wall-clock date.
 */
function testDoorsPreviousMonthMathV14Reviewed() {
  var cases = [
    {year: 2026, month: 9, expected: '2026-08'},
    {year: 2026, month: 1, expected: '2025-12'},
    {year: 2025, month: 12, expected: '2025-11'},
    {year: 2030, month: 3, expected: '2030-02'}
  ];

  var results = cases.map(function(c) {
    var actual = doorsSchedulerPreviousCompleteMonthFromPartsV14_(
      c.year,
      c.month
    );
    return {
      input: c.year + '-' + ('0' + c.month).slice(-2),
      expected: c.expected,
      actual: actual,
      pass: actual === c.expected
    };
  });

  var failed = results.filter(function(r) {
    return !r.pass;
  });

  if (failed.length) {
    throw new Error(
      'Previous-month arithmetic test failed: ' +
      JSON.stringify(failed)
    );
  }

  return {
    status: 'PASS',
    results: results
  };
}


/**
 * READ ONLY.
 * Tests the live current date through explicit Asia/Tokyo extraction.
 */
function testDoorsCurrentTokyoPreviousMonthV14Reviewed() {
  var parts = doorsSchedulerTokyoPartsV14_(new Date());
  var month = doorsSchedulerPreviousCompleteMonthFromPartsV14_(
    parts.year,
    parts.month
  );
  doorsSchedulerAssertMonthV14_(month);

  return {
    status: 'PASS',
    tokyoYear: parts.year,
    tokyoMonth: parts.month,
    previousCompleteMonth: month,
    timezone: DOORS_SCHEDULER_V14.timezone
  };
}


/**
 * WRITE: repairs only B3.
 * Does not run monthly report. Does not touch triggers.
 */
function repairDoorsPanelTargetMonthV14Reviewed() {
  var diag = diagnoseDoorsMonthlySchedulerV14Reviewed();

  if (diag.status !== 'PASS') {
    throw new Error(
      'Diagnostic BLOCK: ' + JSON.stringify(diag)
    );
  }

  var cfg = DOORS_SCHEDULER_V14;
  var ss = SpreadsheetApp.openById(cfg.spreadsheetId);
  var panel = ss.getSheetByName(cfg.panelSheet);

  panel.getRange(cfg.panel.targetMonth)
    .setValue(diag.computedTargetMonth);
  SpreadsheetApp.flush();

  var readback = String(
    panel.getRange(cfg.panel.targetMonth).getDisplayValue() || ''
  ).trim();

  if (readback !== diag.computedTargetMonth) {
    throw new Error(
      'B3 readback mismatch. expected=' +
      diag.computedTargetMonth +
      ', actual=' + readback
    );
  }

  return {
    status: 'PASS',
    repairedTargetMonth: readback,
    reportExecuted: false,
    triggerChanged: false
  };
}


/**
 * WRITE: safe trigger migration.
 *
 * Algorithm:
 * A) create new candidate trigger first
 * B) verify at least one new trigger exists
 * C) delete only legacy DOORS monthly triggers
 * D) if duplicates of the new handler exist, keep one and delete extras
 * E) verify exactly one new handler remains
 *
 * This prevents the "zero monthly trigger" failure mode.
 */
function migrateDoorsMonthlyTriggerV14Reviewed() {
  var cfg = DOORS_SCHEDULER_V14;
  var diag = diagnoseDoorsMonthlySchedulerV14Reviewed();

  if (diag.status !== 'PASS') {
    throw new Error(
      'Diagnostic BLOCK: ' + JSON.stringify(diag)
    );
  }

  var ss = SpreadsheetApp.openById(cfg.spreadsheetId);
  var panel = ss.getSheetByName(cfg.panelSheet);

  var b3 = String(
    panel.getRange(cfg.panel.targetMonth).getDisplayValue() || ''
  ).trim();

  doorsSchedulerAssertMonthV14_(b3);

  if (b3 !== diag.computedTargetMonth) {
    throw new Error(
      'Refusing trigger migration until B3 is repaired. expected=' +
      diag.computedTargetMonth +
      ', actual=' + b3
    );
  }

  // A) Create candidate first.
  ScriptApp.newTrigger(cfg.newHandler)
    .timeBased()
    .onMonthDay(diag.day)
    .atHour(diag.hour)
    .inTimezone(cfg.timezone)
    .create();

  // B) Verify candidate exists before touching legacy triggers.
  var afterCreate = ScriptApp.getProjectTriggers();
  var candidates = afterCreate.filter(function(t) {
    return t.getHandlerFunction() === cfg.newHandler;
  });

  if (candidates.length < 1) {
    throw new Error(
      'New v1.4 trigger was not created; legacy triggers left untouched.'
    );
  }

  // C) Delete only known legacy handlers.
  var removedLegacy = [];
  afterCreate.forEach(function(t) {
    var h = t.getHandlerFunction();
    if (cfg.legacyHandlers.indexOf(h) >= 0) {
      ScriptApp.deleteTrigger(t);
      removedLegacy.push(h);
    }
  });

  // D) De-duplicate new handler, keeping exactly one.
  var afterLegacyRemoval = ScriptApp.getProjectTriggers()
    .filter(function(t) {
      return t.getHandlerFunction() === cfg.newHandler;
    });

  for (var i = 1; i < afterLegacyRemoval.length; i++) {
    ScriptApp.deleteTrigger(afterLegacyRemoval[i]);
  }

  // E) Final readback.
  var finalNew = ScriptApp.getProjectTriggers()
    .filter(function(t) {
      return t.getHandlerFunction() === cfg.newHandler;
    });

  var finalLegacy = ScriptApp.getProjectTriggers()
    .filter(function(t) {
      return cfg.legacyHandlers.indexOf(
        t.getHandlerFunction()
      ) >= 0;
    });

  if (finalNew.length !== 1) {
    throw new Error(
      'Expected exactly one v1.4 trigger after migration, got ' +
      finalNew.length
    );
  }

  if (finalLegacy.length !== 0) {
    throw new Error(
      'Legacy DOORS monthly triggers remain after migration: ' +
      finalLegacy.map(function(t) {
        return t.getHandlerFunction();
      }).join(',')
    );
  }

  panel.getRange(cfg.panel.triggerState).setValue(
    '設定済み｜毎月' +
    diag.day +
    '日 ' +
    diag.hour +
    '時台｜月次自動更新 v1.4 REVIEWED'
  );
  SpreadsheetApp.flush();

  return {
    status: 'PASS',
    installedHandler: cfg.newHandler,
    activeNewTriggerCount: finalNew.length,
    remainingLegacyTriggerCount: finalLegacy.length,
    removedLegacyHandlers: removedLegacy,
    day: diag.day,
    hour: diag.hour,
    timezone: cfg.timezone
  };
}


/**
 * NEW Production scheduled handler.
 *
 * Critical design change:
 * - Month is resolved here, independently.
 * - Explicit `month` is passed into runDoorsMonthlyReportV21(month).
 * - Insights and keyword proposal syncs also receive the same explicit month.
 * - Does NOT call runDoorsMonthlyReportWithInsightsV12(), so the live V12
 *   wrapper cannot re-resolve the target month.
 */
function doorsScheduledMonthlyRunV14Reviewed() {
  var cfg = DOORS_SCHEDULER_V14;
  var ss = SpreadsheetApp.openById(cfg.spreadsheetId);
  var panel = ss.getSheetByName(cfg.panelSheet);

  if (!panel) {
    throw new Error('運用パネルが見つかりません。');
  }

  var tokyo = doorsSchedulerTokyoPartsV14_(new Date());
  var month = doorsSchedulerPreviousCompleteMonthFromPartsV14_(
    tokyo.year,
    tokyo.month
  );

  doorsSchedulerAssertMonthV14_(month);

  panel.getRange(cfg.panel.targetMonth).setValue(month);
  panel.getRange(cfg.panel.status)
    .setValue('AUTO_RUNNING_V14_REVIEWED');
  panel.getRange(cfg.panel.executedAt)
    .setValue(doorsSchedulerNowV14_());
  SpreadsheetApp.flush();

  var b3 = String(
    panel.getRange(cfg.panel.targetMonth).getDisplayValue() || ''
  ).trim();

  if (b3 !== month) {
    throw new Error(
      'Refusing run because B3 readback failed. expected=' +
      month + ', actual=' + b3
    );
  }

  try {
    var monthly = runDoorsMonthlyReportV21(month);

    if (!monthly || monthly.status !== 'SUCCESS') {
      throw new Error(
        'runDoorsMonthlyReportV21 did not return SUCCESS: ' +
        JSON.stringify(monthly)
      );
    }

    var insights = syncDoorsMonthlyInsightsV12(month);

    if (!insights || insights.status !== 'SUCCESS') {
      throw new Error(
        'syncDoorsMonthlyInsightsV12 did not return SUCCESS: ' +
        JSON.stringify(insights)
      );
    }

    var keywords = syncDoorsKeywordProposalsToSlidesV1(month);

    var finalizer = finalizeDoorsMonthlyProductionV16(month);
    if (!finalizer || finalizer.status !== 'SUCCESS') {
      throw new Error(
        'finalizeDoorsMonthlyProductionV16 did not return SUCCESS: ' +
        JSON.stringify(finalizer)
      );
    }

    panel.getRange(cfg.panel.status)
      .setValue('AUTO_SUCCESS_V14_REVIEWED');
    panel.getRange(cfg.panel.executedAt)
      .setValue(doorsSchedulerNowV14_());
    panel.getRange(cfg.panel.keywordPublished)
      .setValue(
        keywords && keywords.publishedRows != null
          ? Number(keywords.publishedRows)
          : 0
      );
    SpreadsheetApp.flush();

    return {
      status: 'SUCCESS',
      version: '1.4-reviewed',
      reportMonth: month,
      monthlyReport: monthly,
      monthlyInsights: insights,
      monthlyFinalizer: finalizer,
      keywordProposals:
        keywords == null ? null : keywords
    };
  } catch (err) {
    panel.getRange(cfg.panel.status)
      .setValue('AUTO_ERROR_V14_REVIEWED');
    panel.getRange(cfg.panel.executedAt)
      .setValue(doorsSchedulerNowV14_());
    SpreadsheetApp.flush();
    throw err;
  }
}


/**
 * Asia/Tokyo calendar extraction without relying on the script runtime timezone.
 */
function doorsSchedulerTokyoPartsV14_(date) {
  var tz = DOORS_SCHEDULER_V14.timezone;
  var year = Number(
    Utilities.formatDate(date, tz, 'yyyy')
  );
  var month = Number(
    Utilities.formatDate(date, tz, 'MM')
  );
  var day = Number(
    Utilities.formatDate(date, tz, 'dd')
  );
  var hour = Number(
    Utilities.formatDate(date, tz, 'HH')
  );
  var minute = Number(
    Utilities.formatDate(date, tz, 'mm')
  );

  if (
    !Number.isInteger(year) ||
    !Number.isInteger(month) ||
    month < 1 ||
    month > 12
  ) {
    throw new Error(
      'Unable to resolve Asia/Tokyo calendar parts.'
    );
  }

  return {
    year: year,
    month: month,
    day: day,
    hour: hour,
    minute: minute
  };
}


/**
 * Pure deterministic month arithmetic.
 * month is 1..12.
 */
function doorsSchedulerPreviousCompleteMonthFromPartsV14_(
  year,
  month
) {
  if (
    !Number.isInteger(year) ||
    !Number.isInteger(month) ||
    month < 1 ||
    month > 12
  ) {
    throw new Error(
      'Invalid year/month parts: ' + year + '/' + month
    );
  }

  var prevYear = year;
  var prevMonth = month - 1;

  if (prevMonth === 0) {
    prevYear = year - 1;
    prevMonth = 12;
  }

  var out =
    String(prevYear) +
    '-' +
    ('0' + prevMonth).slice(-2);

  doorsSchedulerAssertMonthV14_(out);
  return out;
}


function doorsSchedulerAssertMonthV14_(month) {
  var s = String(month || '');

  if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(s)) {
    throw new Error(
      'Invalid report month YYYY-MM: ' + s
    );
  }

  if (s.indexOf('1970-') === 0) {
    throw new Error(
      'Known incident guard: refusing 1970 target month ' + s
    );
  }
}


function doorsSchedulerNowV14_() {
  return Utilities.formatDate(
    new Date(),
    DOORS_SCHEDULER_V14.timezone,
    'yyyy-MM-dd HH:mm:ss'
  );
}

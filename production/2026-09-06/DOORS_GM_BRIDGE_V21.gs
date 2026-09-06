/**
 * DOORS Golden Master Bridge v2.1
 * Safe additive wrapper. Does not modify runDoorsMonthlyReport.
 * Requires the GA4/GSC v2.1 bundle to already exist in the same Apps Script project.
 */

function runDoorsMonthlyReportV21(reportMonth) {
  var month = reportMonth || doorsBridgePrevMonthV21_();
  doorsBridgeAssertMonthV21_(month);
  doorsBridgeWritePanelV21_(month, 'RUNNING', '月次更新を実行中｜検証ゲート確認前', '実行中');

  try {
    var preflight = runDoorsGoldenMasterPreflightV21();
    if (preflight.status !== 'PASS') {
      throw new Error('Golden Master preflight failed: ' + JSON.stringify(preflight));
    }

    // Validate the requested month before the legacy monthly writer is allowed to run.
    // This prevents a failed GA4/GSC gate from occurring after legacy monthly writes.
    var acquisition = runDoorsDataAcquisitionV21(month);
    if (!acquisition || acquisition.status !== 'SUCCESS') {
      throw new Error('Data acquisition failed: ' + JSON.stringify(acquisition));
    }
    if (!acquisition.ga4 || !acquisition.ga4.gate || acquisition.ga4.gate.slidePublishGate !== 'PASS') {
      throw new Error('GA4 validation gate is not PASS: ' + JSON.stringify(acquisition.ga4 && acquisition.ga4.gate));
    }
    if (!acquisition.gscCtr || acquisition.gscCtr.status !== 'SUCCESS' || Number(acquisition.gscCtr.publishedRows) !== 15) {
      throw new Error('GSC CTR validation failed: ' + JSON.stringify(acquisition.gscCtr));
    }

    var monthParts = month.split('-');
    var legacyYear = Number(monthParts[0]);
    var legacyMonth = Number(monthParts[1]);
    var legacyResult = runDoorsMonthlyReport(legacyYear, legacyMonth);

    var seoArticlePerformance = renderDoorsSeoArticlePerformanceProductionV1(month);
    if (!seoArticlePerformance ||
        seoArticlePerformance.status !== 'PASS' ||
        seoArticlePerformance.objectIdSequenceUnchanged !== true) {
      throw new Error(
        'SEO article-performance render verification failed: ' +
        JSON.stringify(seoArticlePerformance)
      );
    }

    doorsBridgeWritePanelV21_(
      month,
      'SUCCESS',
      Number(month.split('-')[1]) + '月実績シート・スライド更新完了｜P.17〜22反映済み',
      '完了｜GA4検証ゲートPASS・P.17〜22更新済み'
    );

    return {
      status: 'SUCCESS',
      reportMonth: month,
      acquisition: acquisition,
      legacyMonthlyReport: legacyResult == null ? null : legacyResult,
      seoArticlePerformance: seoArticlePerformance
    };
  } catch (err) {
    doorsBridgeWritePanelV21_(
      month,
      'BLOCK',
      '月次更新停止｜' + doorsBridgeShortErrorV21_(err),
      '停止｜' + doorsBridgeShortErrorV21_(err)
    );
    throw err;
  }
}

/**
 * Write-free project-level preflight.
 * Checks that required entry points exist before any acquisition/write happens.
 */
function runDoorsGoldenMasterPreflightV21() {
  var checks = {
    runDoorsMonthlyReport: typeof runDoorsMonthlyReport === 'function',
    runDoorsDataAcquisitionV21: typeof runDoorsDataAcquisitionV21 === 'function',
    runDoorsDataAcquisitionPreflightV21: typeof runDoorsDataAcquisitionPreflightV21 === 'function',
    runDoorsLogAcquisitionV21: typeof runDoorsLogAcquisitionV21 === 'function',
    runDoorsGscCtrWorst15V21: typeof runDoorsGscCtrWorst15V21 === 'function',
    renderDoorsSeoArticlePerformanceProductionV1: typeof renderDoorsSeoArticlePerformanceProductionV1 === 'function'
  };

  var missing = Object.keys(checks).filter(function(k) { return !checks[k]; });
  if (missing.length) {
    return {status: 'BLOCK', reason: 'MISSING_FUNCTIONS', missing: missing, checks: checks};
  }

  var dataPreflight = runDoorsDataAcquisitionPreflightV21();
  if (!dataPreflight || dataPreflight.status !== 'PASS') {
    return {status: 'BLOCK', reason: 'DATA_PREFLIGHT_FAILED', checks: checks, dataPreflight: dataPreflight};
  }

  return {status: 'PASS', checks: checks, dataPreflight: dataPreflight};
}

/**
 * Smoke test for 2026-07 known values. This performs live writes through the v2.1 bundle,
 * but does NOT call the legacy runDoorsMonthlyReport.
 */
function smokeTestDoorsDataV21_202607() {
  var preflight = runDoorsGoldenMasterPreflightV21();
  if (preflight.status !== 'PASS') throw new Error('Preflight BLOCK: ' + JSON.stringify(preflight));

  var r = runDoorsDataAcquisitionV21('2026-07');
  var errors = [];
  if (!r || r.status !== 'SUCCESS') errors.push('bundle status != SUCCESS');
  if (!r.ga4 || Number(r.ga4.totalSessions) !== 24743) errors.push('GA4 totalSessions != 24743');
  if (!r.ga4 || Number(r.ga4.channelBreakdownSum) !== 24669) errors.push('channelBreakdownSum != 24669');
  if (!r.ga4 || Number(r.ga4.channelGap) !== 74) errors.push('channelGap != 74');
  if (!r.ga4 || Number(r.ga4.aiSessions) !== 260) errors.push('aiSessions != 260');
  if (!r.ga4 || !r.ga4.gate || r.ga4.gate.slidePublishGate !== 'PASS') errors.push('GA4 gate != PASS');
  if (!r.gscCtr || Number(r.gscCtr.publishedRows) !== 15) errors.push('GSC publishedRows != 15');
  if (!r.gscCtr || String(r.gscCtr.topQuery) !== 'dx化') errors.push('GSC topQuery != dx化');

  return errors.length ? {status: 'FAIL', errors: errors, result: r} : {status: 'PASS', result: r};
}

function doorsBridgePrevMonthV21_() {
  var now = new Date();
  var d = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM');
}

function doorsBridgeAssertMonthV21_(month) {
  if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(String(month))) {
    throw new Error('reportMonth must be YYYY-MM: ' + month);
  }
}


/**
 * Keeps the user-facing operation panel consistent with the actual V2.1 run.
 * Status writes are best-effort and must never mask the original pipeline result.
 */
function doorsBridgeWritePanelV21_(month, state, statusText, slideText) {
  try {
    var ss = SpreadsheetApp.openById('1q1B0JGdLaBOCkDWa4yLNlSFozzf8eaky2sTvTL9BAgo');
    var sh = ss.getSheetByName('運用パネル');
    if (!sh) return;
    sh.getRange('B3').setValue(month);
    sh.getRange('B6').setValue(statusText);
    sh.getRange('B7').setValue(Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss'));
    sh.getRange('B13').setValue(slideText);
  } catch (panelErr) {
    console.warn('Operation panel update failed: ' + panelErr);
  }
}

function doorsBridgeShortErrorV21_(err) {
  var s = String(err && err.message ? err.message : err || 'unknown error');
  return s.length > 180 ? s.slice(0, 177) + '...' : s;
}

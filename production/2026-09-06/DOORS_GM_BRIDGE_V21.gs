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

    // Headline policy: keep operational/data-availability caveats out of slide headlines.
    // Run this after all renderers so legacy writers cannot overwrite the policy.
    var headlinePolicy = doorsApplyHeadlinePolicyV1_(month);
    if (!headlinePolicy || headlinePolicy.status !== 'PASS') {
      throw new Error('Headline policy failed: ' + JSON.stringify(headlinePolicy));
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
      seoArticlePerformance: seoArticlePerformance,
      headlinePolicy: headlinePolicy
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


/**
 * Slide headline policy v1.
 *
 * Principle:
 * - Headline = what the page shows / what happened / the business insight.
 * - Data availability, confirmed month, missing output, validation/gate status,
 *   SSOT state, workaround and cross-page references belong in subtitle/notes.
 *
 * This is intentionally narrow: only known headline-owned slides are touched.
 * It does not reorder slides, add/delete pages, or change data tables.
 */
function doorsApplyHeadlinePolicyV1_(reportMonth) {
  var presentationId = '17qy7JYXMWmjhUbPrmvpURKOkT1TqsF8-jkM74jxw5-4';
  var pres = SlidesApp.openById(presentationId);

  var rules = {
    p8: {
      titleId: 'p8_i2',
      subtitleId: 'p8_i3',
      title: '流入LP｜主要な流入ページの動向',
      subtitle: 'GSCは最新確定値を掲載。記事別の最新変化はP.17〜22で確認できます。'
    },
    p11_channel_202607: {
      titleId: 'p11ch_title',
      subtitleId: 'p11ch_sub',
      title: 'チャネル別流入｜主要チャネルの流入動向',
      subtitle: '総セッションを基準に、主要チャネルとAI / LLM流入の変化を確認します。'
    },
    p13: {
      titleId: 'p13_i2',
      subtitleId: 'p13_i3',
      title: '読了｜カテゴリ別の読了率から改善対象を確認します',
      subtitle: 'カテゴリ間の差を比較し、改善優先度の高いテーマを把握します。計測条件は注記で管理します。'
    },
    p23: {
      titleId: 'p23_i2',
      subtitleId: 'p23_i3',
      title: 'SEO最新ニュース｜検索・生成AI環境の重要トピック',
      subtitle: '検索CTR・生成AI可視性など、次月施策に影響するトピックを整理します。'
    }
  };

  var forbidden = /(未出力|未確定|未登録|直近確定|検証ゲート|validation\s*gate|SSOT|出力待ち|再確定|推測値)/i;
  var slides = pres.getSlides();
  var slideById = {};
  slides.forEach(function(slide) {
    slideById[slide.getObjectId()] = slide;
  });

  var updated = [];
  Object.keys(rules).forEach(function(slideId) {
    var rule = rules[slideId];
    var slide = slideById[slideId];
    if (!slide) throw new Error('Headline policy target slide missing: ' + slideId);

    doorsSetSlideShapeTextByIdV1_(slide, rule.titleId, rule.title);
    doorsSetSlideShapeTextByIdV1_(slide, rule.subtitleId, rule.subtitle);

    var actualTitle = doorsGetSlideShapeTextByIdV1_(slide, rule.titleId);
    if (actualTitle !== rule.title) {
      throw new Error('Headline policy readback mismatch: ' + slideId);
    }
    if (forbidden.test(actualTitle)) {
      throw new Error('Operational wording remains in headline: ' + slideId + ' / ' + actualTitle);
    }
    updated.push(slideId);
  });

  pres.saveAndClose();

  return {
    status: 'PASS',
    reportMonth: reportMonth,
    policyVersion: 'headline-business-first-v1',
    updatedSlides: updated
  };
}

function doorsSetSlideShapeTextByIdV1_(slide, elementId, text) {
  var elements = slide.getPageElements();
  for (var i = 0; i < elements.length; i++) {
    var el = elements[i];
    if (el.getObjectId() !== elementId) continue;
    if (el.getPageElementType() !== SlidesApp.PageElementType.SHAPE) {
      throw new Error('Headline policy target is not a shape: ' + elementId);
    }
    el.asShape().getText().setText(text);
    return;
  }
  throw new Error('Headline policy element missing: ' + elementId);
}

function doorsGetSlideShapeTextByIdV1_(slide, elementId) {
  var elements = slide.getPageElements();
  for (var i = 0; i < elements.length; i++) {
    var el = elements[i];
    if (el.getObjectId() !== elementId) continue;
    if (el.getPageElementType() !== SlidesApp.PageElementType.SHAPE) {
      throw new Error('Headline policy target is not a shape: ' + elementId);
    }
    return String(el.asShape().getText().asString()).replace(/\s+$/, '');
  }
  throw new Error('Headline policy element missing: ' + elementId);
}


function doorsBridgeShortErrorV21_(err) {
  var s = String(err && err.message ? err.message : err || 'unknown error');
  return s.length > 180 ? s.slice(0, 177) + '...' : s;
}

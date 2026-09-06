/**
 * DOORS GA4 log acquisition v2.1
 * Additive/safe module. Does not replace the current Golden Master entry point.
 *
 * Scope:
 *   country = Japan
 *   landingPagePlusQueryString contains /doors/
 *
 * Requires Apps Script Advanced Service:
 *   Google Analytics Data API (AnalyticsData)
 *
 * Raw schemas are treated as contracts and validated by header name before writes.
 */

const DOORS_LOG_CFG_V21 = Object.freeze({
  ga4PropertyId: '260326481',
  ga4LogSpreadsheetId: '13RbANqumJik5iGSRi8mcydUU3gkl0BOPS5udE28p8Pc',
  mainSpreadsheetId: '1q1B0JGdLaBOCkDWa4yLNlSFozzf8eaky2sTvTL9BAgo',
  country: 'Japan',
  landingField: 'landingPagePlusQueryString',
  landingContains: '/doors/',
  landingFilterLabel: 'landingPagePlusQueryString contains /doors/',
  channelDefinitionVersion: 'ga4-channel-japan-doors-lpqs-v2',
  aiDefinitionVersion: 'ai-source-v4-japan-doors-lpqs-copilot',
  runAuditDefinitionVersion: 'run-audit-v1',
  rawChannelSheet: 'LOG_GA4_チャネル月次',
  rawAiSheet: 'LOG_AI_LLM月次',
  runLogSheet: 'LOG_取得実行履歴',
  gateSheet: 'LOG_検証ゲート',
  knownTotalSheet: '日本_DOORS月別セッション',
  mainChannelSheet: 'チャネル別月次',
  mainAggregateSheet: '月次レポート集約',
  maxChannelNonAdditiveRatio: 0.01
});

const DOORS_HEADERS_V21 = Object.freeze({
  channel: [
    'report_month','channel','sessions','country','landing_page_filter','dimension','metric',
    'definition_version','run_id','fetched_at','date_start','date_end','source','status'
  ],
  ai: [
    'report_month','service','session_source','session_medium','sessions','total_sessions','share',
    'country','landing_page_filter','definition_version','run_id','fetched_at','date_start','date_end','source','status'
  ],
  audit: [
    'run_id','executed_at','target_month','dataset','definition_version','date_start','date_end',
    'row_count','status','error_type','error_message','source_sheet','destination_sheet','note'
  ],
  gate: [
    'report_month','known_total_sessions','channel_log_sum','channel_gate','ai_log_sum','ai_share','ai_gate',
    'channel_definition','ai_definition','latest_run_status','slide_publish_gate','note'
  ],
  mainChannel: ['month','channel','sessions','definition_version','source_sheet'],
  mainAggregate: [
    'report_month','section','metric','segment','value','numerator','denominator','definition_version',
    'source_file','source_sheet','status','note'
  ]
});

/**
 * Safe monthly entry point for v2.1.
 * @param {string=} reportMonth YYYY-MM. Defaults to previous complete month.
 */
function runDoorsLogAcquisitionV21(reportMonth) {
  const month = reportMonth || doorsPrevMonthV21_();
  doorsAssertMonthV21_(month);
  doorsPreflightHeadersV21_();

  const baseRunId = 'doors_v21_' + month.replace('-', '') + '_' +
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
  const runIds = {
    channel: baseRunId + ':channel',
    ai: baseRunId + ':ai',
    all: baseRunId + ':all'
  };
  const range = doorsMonthRangeV21_(month);

  try {
    const channel = fetchGa4ChannelMonthlyV21_(month, range, runIds.channel);
    const ai = fetchGa4AiLlmMonthlyV21_(month, range, runIds.ai, channel.totalSessions);
    const gate = validateDoorsGa4MonthV21_(month, channel, ai);

    if (gate.slidePublishGate !== 'PASS') {
      throw new Error('Validation gate blocked writes: ' + JSON.stringify(gate));
    }

    commitDoorsRawLogsV21_(month, channel, ai, runIds, range);
    doorsWriteGateV21_(month, gate);
    syncDoorsValidatedLogsToSsotV21_(month, channel, ai, gate);
    doorsAppendAuditV21_({
      runId: runIds.all,
      executedAt: new Date(),
      targetMonth: month,
      dataset: 'ALL',
      definitionVersion: DOORS_LOG_CFG_V21.channelDefinitionVersion + '|' + DOORS_LOG_CFG_V21.aiDefinitionVersion,
      dateStart: range.startDate,
      dateEnd: range.endDate,
      rowCount: channel.rawRows.length + ai.rawRows.length,
      status: 'SUCCESS',
      sourceSheet: 'Brainpad_ALL_GA4 properties/' + DOORS_LOG_CFG_V21.ga4PropertyId,
      destinationSheet: DOORS_LOG_CFG_V21.rawChannelSheet + '|' + DOORS_LOG_CFG_V21.rawAiSheet,
      note: 'v2.1 validation PASS; targeted SSOT writes only'
    });

    return {
      status: 'SUCCESS',
      runId: baseRunId,
      reportMonth: month,
      totalSessions: channel.totalSessions,
      channelBreakdownSum: channel.breakdownSum,
      channelGap: channel.gap,
      channelGapRatio: channel.gapRatio,
      aiSessions: ai.aiSum,
      aiShare: ai.aiShare,
      gate: gate
    };
  } catch (err) {
    doorsAppendAuditV21_({
      runId: runIds.all,
      executedAt: new Date(),
      targetMonth: month,
      dataset: 'ALL',
      definitionVersion: DOORS_LOG_CFG_V21.channelDefinitionVersion + '|' + DOORS_LOG_CFG_V21.aiDefinitionVersion,
      dateStart: range.startDate,
      dateEnd: range.endDate,
      rowCount: 0,
      status: 'FAILED',
      errorType: 'RUN_FAILED',
      errorMessage: String(err && err.stack || err),
      sourceSheet: 'Brainpad_ALL_GA4 properties/' + DOORS_LOG_CFG_V21.ga4PropertyId,
      destinationSheet: DOORS_LOG_CFG_V21.rawChannelSheet + '|' + DOORS_LOG_CFG_V21.rawAiSheet,
      note: 'SSOT write is refused unless validation PASS'
    });
    throw err;
  }
}

/**
 * Channel acquisition.
 * Saves a same-run __TOTAL__ row and separate channel breakdown rows.
 */
function fetchGa4ChannelMonthlyV21_(reportMonth, range, runId) {
  const totalReq = {
    dateRanges: [{startDate: range.startDate, endDate: range.endDate}],
    dimensions: [{name: 'yearMonth'}],
    metrics: [{name: 'sessions'}],
    dimensionFilter: doorsJapanDoorsFilterV21_(),
    limit: '1000'
  };
  const totalRes = AnalyticsData.Properties.runReport(totalReq, 'properties/' + DOORS_LOG_CFG_V21.ga4PropertyId);
  const totalSessions = doorsExtractMonthMetricV21_(totalRes.rows || [], reportMonth, 0, 0);
  if (totalSessions === null) throw new Error('GA4 total sessions row missing for ' + reportMonth);

  const breakdownReq = {
    dateRanges: [{startDate: range.startDate, endDate: range.endDate}],
    dimensions: [{name: 'yearMonth'}, {name: 'sessionDefaultChannelGroup'}],
    metrics: [{name: 'sessions'}],
    dimensionFilter: doorsJapanDoorsFilterV21_(),
    limit: '10000'
  };
  const breakdownRes = AnalyticsData.Properties.runReport(breakdownReq, 'properties/' + DOORS_LOG_CFG_V21.ga4PropertyId);
  const channelRows = [];
  (breakdownRes.rows || []).forEach(function(row) {
    const ym = doorsYearMonthToHyphenV21_(doorsDimV21_(row, 0));
    if (ym !== reportMonth) return;
    channelRows.push({
      month: ym,
      channel: doorsDimV21_(row, 1) || '(not set)',
      sessions: doorsMetricNumberV21_(row, 0)
    });
  });
  channelRows.sort(function(a, b) { return b.sessions - a.sessions || a.channel.localeCompare(b.channel); });

  const fetchedAt = new Date();
  const rawRows = [doorsBuildChannelRawRowV21_({
    month: reportMonth,
    channel: '__TOTAL__',
    sessions: totalSessions,
    dimension: 'yearMonth',
    runId: runId,
    fetchedAt: fetchedAt,
    dateStart: range.startDate,
    dateEnd: range.endDate,
    status: 'SUCCESS_TOTAL'
  })];
  channelRows.forEach(function(r) {
    rawRows.push(doorsBuildChannelRawRowV21_({
      month: r.month,
      channel: r.channel,
      sessions: r.sessions,
      dimension: 'sessionDefaultChannelGroup',
      runId: runId,
      fetchedAt: fetchedAt,
      dateStart: range.startDate,
      dateEnd: range.endDate,
      status: 'SUCCESS_BREAKDOWN'
    }));
  });
  doorsAssertRowsWidthV21_(rawRows, DOORS_HEADERS_V21.channel.length, 'channel raw');

  const breakdownSum = channelRows.reduce(function(sum, r) { return sum + r.sessions; }, 0);
  const gap = totalSessions - breakdownSum;
  const gapRatio = totalSessions ? Math.abs(gap) / totalSessions : 0;

  return {
    totalSessions: totalSessions,
    channelRows: channelRows,
    rawRows: rawRows,
    breakdownSum: breakdownSum,
    gap: gap,
    gapRatio: gapRatio
  };
}

function fetchGa4AiLlmMonthlyV21_(reportMonth, range, runId, totalSessions) {
  const req = {
    dateRanges: [{startDate: range.startDate, endDate: range.endDate}],
    dimensions: [{name: 'yearMonth'}, {name: 'sessionSource'}, {name: 'sessionMedium'}],
    metrics: [{name: 'sessions'}],
    dimensionFilter: doorsJapanDoorsFilterV21_(),
    limit: '100000'
  };
  const res = AnalyticsData.Properties.runReport(req, 'properties/' + DOORS_LOG_CFG_V21.ga4PropertyId);
  const items = [];
  (res.rows || []).forEach(function(row) {
    const ym = doorsYearMonthToHyphenV21_(doorsDimV21_(row, 0));
    if (ym !== reportMonth) return;
    const source = doorsDimV21_(row, 1) || '(not set)';
    const medium = doorsDimV21_(row, 2) || '(not set)';
    const cls = doorsClassifyAiSourceV21_(source, medium);
    if (!cls) return;
    items.push({
      month: ym,
      service: cls.service,
      source: source,
      medium: medium,
      sessions: doorsMetricNumberV21_(row, 0)
    });
  });
  items.sort(function(a, b) {
    return b.sessions - a.sessions || a.service.localeCompare(b.service) || a.source.localeCompare(b.source);
  });

  const fetchedAt = new Date();
  const rawRows = items.map(function(r) {
    return doorsBuildAiRawRowV21_({
      month: r.month,
      service: r.service,
      source: r.source,
      medium: r.medium,
      sessions: r.sessions,
      totalSessions: totalSessions,
      runId: runId,
      fetchedAt: fetchedAt,
      dateStart: range.startDate,
      dateEnd: range.endDate,
      status: 'SUCCESS'
    });
  });
  doorsAssertRowsWidthV21_(rawRows, DOORS_HEADERS_V21.ai.length, 'AI raw');

  const aiSum = items.reduce(function(sum, r) { return sum + r.sessions; }, 0);
  const aiShare = totalSessions ? aiSum / totalSessions : 0;

  return {items: items, rawRows: rawRows, aiSum: aiSum, aiShare: aiShare};
}

function validateDoorsGa4MonthV21_(reportMonth, channel, ai) {
  const snapshotTotal = doorsKnownTotalSessionsV21_(reportMonth);
  const channelPass = channel.totalSessions > 0 &&
    isFinite(channel.breakdownSum) &&
    channel.breakdownSum >= 0 &&
    channel.gapRatio <= DOORS_LOG_CFG_V21.maxChannelNonAdditiveRatio;
  const aiPass = ai.aiSum >= 0 && ai.aiSum <= channel.totalSessions;
  const slidePass = channelPass && aiPass;

  const channelGate = channelPass ?
    (channel.gap === 0 ? 'PASS' : 'PASS_NONADDITIVE') : 'FAIL';
  const historyNote = snapshotTotal === null ? 'snapshot unavailable' :
    (snapshotTotal === channel.totalSessions ? 'snapshot match' :
      'snapshot ' + snapshotTotal + ' -> same-run total ' + channel.totalSessions);

  const row = [
    reportMonth,
    channel.totalSessions,
    channel.breakdownSum,
    channelGate,
    ai.aiSum,
    ai.aiShare,
    aiPass ? 'PASS' : 'FAIL',
    DOORS_LOG_CFG_V21.channelDefinitionVersion,
    DOORS_LOG_CFG_V21.aiDefinitionVersion,
    slidePass ? 'FETCH_SUCCESS' : 'FAILED',
    slidePass ? 'PASS' : 'BLOCK',
    historyNote + '; channel breakdown=' + channel.breakdownSum + '; diff=' + channel.gap +
      ' (' + (channel.gapRatio * 100).toFixed(2) + '%)'
  ];
  return {
    gateRow: row,
    canonicalTotal: channel.totalSessions,
    snapshotTotal: snapshotTotal,
    channelBreakdownSum: channel.breakdownSum,
    channelGap: channel.gap,
    channelGapRatio: channel.gapRatio,
    channelGate: channelGate,
    aiSum: ai.aiSum,
    aiShare: ai.aiShare,
    aiGate: aiPass ? 'PASS' : 'FAIL',
    slidePublishGate: slidePass ? 'PASS' : 'BLOCK'
  };
}


function commitDoorsRawLogsV21_(reportMonth, channel, ai, runIds, range) {
  doorsReplaceRowsAppendThenDeleteByFieldsV21_(
    DOORS_LOG_CFG_V21.ga4LogSpreadsheetId,
    DOORS_LOG_CFG_V21.rawChannelSheet,
    {report_month: reportMonth, definition_version: DOORS_LOG_CFG_V21.channelDefinitionVersion},
    channel.rawRows
  );
  doorsAppendAuditV21_({
    runId: runIds.channel,
    executedAt: new Date(),
    targetMonth: reportMonth,
    dataset: 'GA4_CHANNEL',
    definitionVersion: DOORS_LOG_CFG_V21.channelDefinitionVersion,
    dateStart: range.startDate,
    dateEnd: range.endDate,
    rowCount: channel.rawRows.length,
    status: channel.gap === 0 ? 'FETCH_OK' : 'FETCH_OK_NONADDITIVE',
    errorType: channel.gap === 0 ? '' : 'CHANNEL_SUM_DIFF',
    errorMessage: channel.gap === 0 ? '' : 'same-run total=' + channel.totalSessions + '; breakdown=' + channel.breakdownSum + '; diff=' + channel.gap,
    sourceSheet: 'Brainpad_ALL_GA4 properties/' + DOORS_LOG_CFG_V21.ga4PropertyId,
    destinationSheet: DOORS_LOG_CFG_V21.rawChannelSheet,
    note: 'append-first replacement; same-run __TOTAL__ is canonical'
  });

  doorsReplaceRowsAppendThenDeleteByFieldsV21_(
    DOORS_LOG_CFG_V21.ga4LogSpreadsheetId,
    DOORS_LOG_CFG_V21.rawAiSheet,
    {report_month: reportMonth, definition_version: DOORS_LOG_CFG_V21.aiDefinitionVersion},
    ai.rawRows
  );
  doorsAppendAuditV21_({
    runId: runIds.ai,
    executedAt: new Date(),
    targetMonth: reportMonth,
    dataset: 'AI_LLM_SOURCE_REGEX',
    definitionVersion: DOORS_LOG_CFG_V21.aiDefinitionVersion,
    dateStart: range.startDate,
    dateEnd: range.endDate,
    rowCount: ai.rawRows.length,
    status: 'FETCH_OK',
    sourceSheet: 'Brainpad_ALL_GA4 properties/' + DOORS_LOG_CFG_V21.ga4PropertyId,
    destinationSheet: DOORS_LOG_CFG_V21.rawAiSheet,
    note: 'append-first replacement; AI sessions=' + ai.aiSum + '; denominator=' + channel.totalSessions + '; share=' + ai.aiShare
  });
}

function doorsWriteGateV21_(reportMonth, gate) {
  doorsUpsertByHeaderKeyV21_(
    DOORS_LOG_CFG_V21.ga4LogSpreadsheetId,
    DOORS_LOG_CFG_V21.gateSheet,
    'report_month',
    reportMonth,
    gate.gateRow,
    DOORS_HEADERS_V21.gate
  );
}

function syncDoorsValidatedLogsToSsotV21_(reportMonth, channel, ai, gate) {
  if (gate.slidePublishGate !== 'PASS') throw new Error('Refusing SSOT write before validation PASS');
  doorsSyncMainChannelV21_(reportMonth, channel.channelRows);
  doorsSyncMainAggregateV21_(reportMonth, channel, ai, gate);
}

function doorsSyncMainChannelV21_(reportMonth, channelRows) {
  doorsAssertSheetHeaderV21_(
    DOORS_LOG_CFG_V21.mainSpreadsheetId,
    DOORS_LOG_CFG_V21.mainChannelSheet,
    DOORS_HEADERS_V21.mainChannel
  );
  const rows = channelRows.map(function(r) {
    return [
      r.month, r.channel, r.sessions, DOORS_LOG_CFG_V21.channelDefinitionVersion,
      'GA4_国別アクティブユーザー月次推移/' + DOORS_LOG_CFG_V21.rawChannelSheet
    ];
  });
  doorsAssertRowsWidthV21_(rows, DOORS_HEADERS_V21.mainChannel.length, 'main channel');
  doorsReplaceRowsAppendThenDeleteByFieldsV21_(
    DOORS_LOG_CFG_V21.mainSpreadsheetId,
    DOORS_LOG_CFG_V21.mainChannelSheet,
    {month: reportMonth, definition_version: DOORS_LOG_CFG_V21.channelDefinitionVersion},
    rows
  );
}

function doorsSyncMainAggregateV21_(reportMonth, channel, ai, gate) {
  doorsAssertSheetHeaderV21_(
    DOORS_LOG_CFG_V21.mainSpreadsheetId,
    DOORS_LOG_CFG_V21.mainAggregateSheet,
    DOORS_HEADERS_V21.mainAggregate
  );

  const aggregatePredicate = function(obj) {
    if (doorsCellMonthV21_(obj.report_month) !== reportMonth) return false;
    return (String(obj.metric) === 'AI経由セッション比率' && String(obj.definition_version) === DOORS_LOG_CFG_V21.aiDefinitionVersion) ||
      (String(obj.metric) === '日本 /doors/ セッション' && String(obj.definition_version) === DOORS_LOG_CFG_V21.channelDefinitionVersion);
  };

  const channelStatus = channel.gap === 0 ? 'READY_LIVE' : 'READY_LIVE_DIMENSION_GAP_ACCEPTED';
  const rows = [
    [
      reportMonth, 'AI経由', 'AI経由セッション比率', 'source partial match + approved medium', ai.aiShare,
      ai.aiSum, channel.totalSessions, DOORS_LOG_CFG_V21.aiDefinitionVersion,
      'GA4_国別アクティブユーザー月次推移', DOORS_LOG_CFG_V21.rawAiSheet, 'READY_LIVE',
      'source partial match + medium whitelist; same-run total denominator'
    ],
    [
      reportMonth, 'チャネル', '日本 /doors/ セッション', '全チャネル合計', channel.totalSessions,

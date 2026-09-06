/**
 * DOORS_MONTHLY_MENU_V1.gs
 *
 * Additive spreadsheet menu adapter for DOORS Production.
 *
 * Menu:
 * - DOORS月次レポート > レポート生成
 * - DOORS月次レポート > キーワードを更新してレポートを再生成
 *
 * Safety:
 * - Does not redefine existing onOpen().
 * - Uses a dedicated installable onOpen trigger.
 * - "レポート生成" delegates to the existing Production authority
 *   runDoorsMonthlyReportV21(reportMonth).
 * - "キーワードを更新してレポートを再生成" does NOT reacquire GA4/GSC/Ahrefs.
 *   It safely re-renders only P.30 from the human-approved rows in 新規KW候補.
 */

var DOORS_MONTHLY_MENU_V1 = Object.freeze({
  version: '1.0.0',
  spreadsheetId: '1q1B0JGdLaBOCkDWa4yLNlSFozzf8eaky2sTvTL9BAgo',
  menuName: 'DOORS月次レポート',
  panelSheet: '運用パネル',
  panelMonthCell: 'B3',
  reportHandler: 'runDoorsMonthlyReportFromMenuV1',
  keywordHandler: 'runDoorsKeywordRefreshAndRegenerateFromMenuV1',
  onOpenHandler: 'doorsMonthlyMenuOnOpenV1'
});

function installDoorsMonthlyMenuV1() {
  var ss = SpreadsheetApp.openById(DOORS_MONTHLY_MENU_V1.spreadsheetId);
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    try {
      if (trigger.getHandlerFunction() === DOORS_MONTHLY_MENU_V1.onOpenHandler) {
        ScriptApp.deleteTrigger(trigger);
      }
    } catch (_) {}
  });
  ScriptApp.newTrigger(DOORS_MONTHLY_MENU_V1.onOpenHandler)
    .forSpreadsheet(ss)
    .onOpen()
    .create();
  try { doorsMonthlyMenuOnOpenV1(); } catch (_) {}
  return {
    status: 'PASS',
    version: DOORS_MONTHLY_MENU_V1.version,
    menu: DOORS_MONTHLY_MENU_V1.menuName,
    triggerHandler: DOORS_MONTHLY_MENU_V1.onOpenHandler
  };
}

function doorsMonthlyMenuOnOpenV1() {
  var active = SpreadsheetApp.getActiveSpreadsheet();
  if (!active || active.getId() !== DOORS_MONTHLY_MENU_V1.spreadsheetId) return;
  SpreadsheetApp.getUi()
    .createMenu(DOORS_MONTHLY_MENU_V1.menuName)
    .addItem('レポート生成', DOORS_MONTHLY_MENU_V1.reportHandler)
    .addItem('キーワードを更新してレポートを再生成', DOORS_MONTHLY_MENU_V1.keywordHandler)
    .addToUi();
}

function runDoorsMonthlyReportFromMenuV1() {
  var ss = doorsMonthlyMenuAssertSpreadsheetV1_();
  var month = doorsMonthlyMenuResolveMonthV1_(ss);
  if (typeof runDoorsMonthlyReportV21 !== 'function') {
    throw new Error('レポート生成停止: runDoorsMonthlyReportV21 が見つかりません。');
  }
  ss.toast(month + ' の月次レポート生成を開始します。', DOORS_MONTHLY_MENU_V1.menuName, 5);
  try {
    var result = runDoorsMonthlyReportV21(month);
    ss.toast(month + ' のレポート生成が完了しました。', DOORS_MONTHLY_MENU_V1.menuName, 8);
    return {status:'PASS', action:'MONTHLY_REPORT_GENERATED', reportMonth:month, result:result};
  } catch (error) {
    ss.toast('レポート生成に失敗しました。実行ログを確認してください。', DOORS_MONTHLY_MENU_V1.menuName, 10);
    throw error;
  }
}

function runDoorsKeywordRefreshAndRegenerateFromMenuV1() {
  var ss = doorsMonthlyMenuAssertSpreadsheetV1_();
  var month = doorsMonthlyMenuResolveMonthV1_(ss);
  if (typeof renderDoorsKeywordProposalP30V1 !== 'function') {
    throw new Error('キーワード再生成停止: renderDoorsKeywordProposalP30V1 が見つかりません。 DOORS_GM_BRIDGE_V21_V24_KEYWORD_HUB_P30 以降を保存してください。');
  }
  ss.toast(month + ' の採用キーワードをP.30へ反映しています。', DOORS_MONTHLY_MENU_V1.menuName, 5);
  try {
    var result = renderDoorsKeywordProposalP30V1(month);
    if (!result || result.status !== 'PASS') {
      throw new Error('P.30キーワード反映に失敗しました: ' + JSON.stringify(result));
    }
    ss.toast('キーワード反映完了｜採用 ' + Number(result.renderedCount || 0) + '件｜P.30を再生成しました。', DOORS_MONTHLY_MENU_V1.menuName, 8);
    return {status:'PASS', action:'KEYWORD_P30_REGENERATED', reportMonth:month, renderedCount:Number(result.renderedCount || 0), result:result};
  } catch (error) {
    ss.toast('キーワード反映に失敗しました。K列「採用状況」と実行ログを確認してください。', DOORS_MONTHLY_MENU_V1.menuName, 10);
    throw error;
  }
}

function doorsMonthlyMenuAssertSpreadsheetV1_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss || ss.getId() !== DOORS_MONTHLY_MENU_V1.spreadsheetId) {
    throw new Error('DOORS Production Spreadsheetから実行してください。');
  }
  return ss;
}

function doorsMonthlyMenuResolveMonthV1_(ss) {
  var panel = ss.getSheetByName(DOORS_MONTHLY_MENU_V1.panelSheet);
  if (!panel) throw new Error('運用パネルが見つかりません。');
  var raw = String(panel.getRange(DOORS_MONTHLY_MENU_V1.panelMonthCell).getDisplayValue() || '').trim();
  if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(raw)) {
    throw new Error('運用パネル!' + DOORS_MONTHLY_MENU_V1.panelMonthCell + ' の対象月が YYYY-MM 形式ではありません: ' + raw);
  }
  return raw;
}

/**
 * DOORS Monthly Production Finalizer v1.6 CANDIDATE
 *
 * Additive finalizer for the current 36-slide Production deck.
 * Scope is intentionally narrow:
 * - Slide 7 (object p8): regenerate LP title/subtitle/insight from GSC page SSOT.
 * - Slide 30 (object SLIDES_API1535510610_0): keep month labels current and repair
 *   the "選定時の見方" body transform only if vertical scale is negative.
 * - No data acquisition. No trigger changes. No LLMO changes.
 *
 * Intended call order: existing monthly core -> insights -> keyword sync -> this finalizer.
 */
var DOORS_MONTHLY_FINALIZER_V16 = Object.freeze({
  version: '1.6.0-candidate',
  spreadsheetId: '1q1B0JGdLaBOCkDWa4yLNlSFozzf8eaky2sTvTL9BAgo',
  presentationId: '17qy7JYXMWmjhUbPrmvpURKOkT1TqsF8-jkM74jxw5-4',
  expectedSlideCount: 36,
  lpSlideId: 'p8',
  lpTitleId: 'p8_i2',
  lpSubtitleId: 'p8_i3',
  lpInsightId: 'p8_i8',
  lpTableId: 'p8_i9',
  keywordSlideId: 'SLIDES_API1535510610_0',
  keywordTitleId: 'SLIDES_API1535510610_1',
  keywordSubtitleId: 'SLIDES_API1535510610_2',
  keywordNoteId: 'SLIDES_API1535510610_7',
  gscPageSheet: 'GSCページ別データ',
  panelSheet: '運用パネル'
});

function finalizeDoorsMonthlyProductionV16(reportMonth) {
  doorsFinalizerAssertMonthV16_(reportMonth);
  var cfg = DOORS_MONTHLY_FINALIZER_V16;
  var ss = SpreadsheetApp.openById(cfg.spreadsheetId);
  var panel = ss.getSheetByName(cfg.panelSheet);
  if (!panel) throw new Error('運用パネルが見つかりません。');
  var panelMonth = String(panel.getRange('B3').getDisplayValue() || '').trim();
  if (panelMonth !== reportMonth) {
    throw new Error('Finalizer month mismatch. panel=' + panelMonth + ', requested=' + reportMonth);
  }

  var pres = SlidesApp.openById(cfg.presentationId);
  var slides = pres.getSlides();
  if (slides.length !== cfg.expectedSlideCount) {
    throw new Error('Finalizer BLOCK: expected 36 slides, actual=' + slides.length);
  }
  var byId = {};
  slides.forEach(function(s){ byId[s.getObjectId()] = s; });
  if (!byId[cfg.lpSlideId]) throw new Error('Missing LP slide: ' + cfg.lpSlideId);
  if (!byId[cfg.keywordSlideId]) throw new Error('Missing keyword slide: ' + cfg.keywordSlideId);

  var lp = doorsFinalizerBuildLpInsightV16_(ss, reportMonth);
  doorsFinalizerSetShapeTextV16_(byId[cfg.lpSlideId], cfg.lpTitleId, lp.title);
  doorsFinalizerSetShapeTextV16_(byId[cfg.lpSlideId], cfg.lpSubtitleId, lp.subtitle);
  doorsFinalizerSetShapeTextV16_(byId[cfg.lpSlideId], cfg.lpInsightId, lp.insight);
  doorsFinalizerWriteLpTableV16_(byId[cfg.lpSlideId], lp);

  var parts = doorsFinalizerMonthPartsV16_(reportMonth);
  doorsFinalizerSetShapeTextV16_(
    byId[cfg.keywordSlideId], cfg.keywordTitleId,
    '新規提案キーワード｜' + parts.year + '年' + parts.month + '月'
  );
  doorsFinalizerSetShapeTextV16_(
    byId[cfg.keywordSlideId], cfg.keywordSubtitleId,
    reportMonth + ' レポート｜人が採用し、シートへ転記した候補のみ掲載'
  );
  var transformRepair = doorsFinalizerRepairVerticalFlipV16_(byId[cfg.keywordSlideId], cfg.keywordNoteId);

  var afterSlides = pres.getSlides();
  if (afterSlides.length !== cfg.expectedSlideCount) {
    throw new Error('Finalizer regression: slide count changed.');
  }
  var afterIds = afterSlides.map(function(s){ return s.getObjectId(); });
  var beforeIds = slides.map(function(s){ return s.getObjectId(); });
  if (JSON.stringify(afterIds) !== JSON.stringify(beforeIds)) {
    throw new Error('Finalizer regression: slide order/objectIds changed.');
  }

  return {
    status: 'SUCCESS', version: cfg.version, reportMonth: reportMonth,
    lp: lp, keywordNoteTransformRepair: transformRepair,
    slideCount: afterSlides.length
  };
}

function diagnoseDoorsMonthlyProductionFinalizerV16(reportMonth) {
  doorsFinalizerAssertMonthV16_(reportMonth);
  var cfg = DOORS_MONTHLY_FINALIZER_V16;
  var ss = SpreadsheetApp.openById(cfg.spreadsheetId);
  var panel = ss.getSheetByName(cfg.panelSheet);
  var pres = SlidesApp.openById(cfg.presentationId);
  var slides = pres.getSlides();
  var ids = slides.map(function(s){ return s.getObjectId(); });
  var lpData;
  var lpError = '';
  try { lpData = doorsFinalizerBuildLpInsightV16_(ss, reportMonth); }
  catch (e) { lpError = String(e && e.message ? e.message : e); }
  var checks = {
    panelPresent: !!panel,
    panelMonth: !!panel && String(panel.getRange('B3').getDisplayValue() || '').trim() === reportMonth,
    slideCount36: slides.length === cfg.expectedSlideCount,
    lpSlidePresent: ids.indexOf(cfg.lpSlideId) >= 0,
    keywordSlidePresent: ids.indexOf(cfg.keywordSlideId) >= 0,
    lpDataReady: !lpError
  };
  var failed = Object.keys(checks).filter(function(k){ return checks[k] !== true; });
  return {status: failed.length ? 'BLOCK' : 'PASS', productionChanged:false,
    version:cfg.version, reportMonth:reportMonth, checks:checks, failed:failed, lpError:lpError, lp:lpData || null};
}

function doorsFinalizerBuildLpInsightV16_(ss, reportMonth) {
  var cfg = DOORS_MONTHLY_FINALIZER_V16;
  var prevMonth = doorsFinalizerPrevMonthV16_(reportMonth);
  var sh = ss.getSheetByName(cfg.gscPageSheet);
  if (!sh) throw new Error('Missing sheet: ' + cfg.gscPageSheet);
  var values = sh.getDataRange().getDisplayValues();
  if (values.length < 2) throw new Error('GSC page SSOT is empty.');
  var h = values[0];
  function ix(name){ var i=h.indexOf(name); if(i<0) throw new Error('Missing GSC header: '+name); return i; }
  var cMonth=ix('month'), cRank=ix('rank'), cPage=ix('page'), cTitle=ix('title'),
      cClicks=ix('clicks'), cImp=ix('impressions'), cCtr=ix('ctr'), cPos=ix('position');
  var current=[], prevByUrl={};
  values.slice(1).forEach(function(r){
    var m=String(r[cMonth]||'').trim(), url=String(r[cPage]||'').trim();
    if (!url) return;
    var obj={month:m,rank:Number(r[cRank]),url:url,title:String(r[cTitle]||'').trim(),
      clicks:Number(r[cClicks]),impressions:Number(r[cImp]),ctr:Number(r[cCtr]),position:Number(r[cPos])};
    if (m===prevMonth) prevByUrl[url]=obj;
    if (m===reportMonth && url.indexOf('/doors/contents/')>=0 && isFinite(obj.clicks)) current.push(obj);
  });
  current.sort(function(a,b){ return a.rank-b.rank; });
  current=current.slice(0,10);
  if (current.length !== 10) throw new Error('Expected 10 current /contents/ GSC pages; got '+current.length);
  current.forEach(function(x){
    var p=prevByUrl[x.url];
    x.prevClicks=p && isFinite(p.clicks) ? p.clicks : 0;
    x.prevPosition=p && isFinite(p.position) ? p.position : NaN;
    x.delta=x.clicks-x.prevClicks;
  });
  var totalDelta=current.reduce(function(s,x){return s+x.delta;},0);
  var declines=current.filter(function(x){return x.delta<0;});
  var gains=current.filter(function(x){return x.delta>0;});
  var pool=(totalDelta<0 ? declines : gains).slice().sort(function(a,b){return Math.abs(b.delta)-Math.abs(a.delta);}).slice(0,3);
  var denom=(totalDelta<0 ? declines : gains).reduce(function(s,x){return s+Math.abs(x.delta);},0);
  var topAbs=pool.reduce(function(s,x){return s+Math.abs(x.delta);},0);
  var share=denom ? topAbs/denom : 0;
  var approxTenth=Math.max(0,Math.min(10,Math.round(share*10)));
  var jpMonth=Number(reportMonth.slice(5,7));
  var jpPrev=Number(prevMonth.slice(5,7));
  var title;
  var insight;
  if (totalDelta < 0 && pool.length) {
    title='流入LP｜クリック減の約'+approxTenth+'割が上位3ページに集中しました';
    insight='主要10ページ中'+declines.length+'ページでクリックが減少。特に'+
      pool.map(function(x){return '「'+x.title+'」'+doorsFinalizerSignedIntV16_(x.delta);}).join('、')+
      'の3ページで減少幅の約'+Math.round(share*100)+'%を占めます。まず3ページの表示回数・CTR・順位をクエリ単位で確認します。';
  } else if (totalDelta > 0 && pool.length) {
    title='流入LP｜クリック増の約'+approxTenth+'割が上位3ページに集中しました';
    insight='主要10ページ中'+gains.length+'ページでクリックが増加。特に'+
      pool.map(function(x){return '「'+x.title+'」'+doorsFinalizerSignedIntV16_(x.delta);}).join('、')+
      'の3ページで増加幅の約'+Math.round(share*100)+'%を占めます。伸長要因をクエリ・CTR・順位の順で確認します。';
  } else {
    title='流入LP｜主要10ページの前月差は概ね横ばいです';
    insight='主要10ページ全体のクリック差は小幅です。個別ページの表示回数・CTR・順位を確認し、変化が大きいページだけを優先して確認します。';
  }
  return {title:title,
    subtitle:'Search Consoleのページ別実績を、'+jpMonth+'月と'+jpPrev+'月の同一条件で比較します。',
    insight:insight,totalDelta:totalDelta,declines:declines.length,gains:gains.length,
    top3:pool.map(function(x){return {title:x.title,delta:x.delta};}), share:share,
    currentMonth:reportMonth,prevMonth:prevMonth,currentMonthLabel:jpMonth+'月',prevMonthLabel:jpPrev+'月',rows:current};
}

function doorsFinalizerWriteLpTableV16_(slide, lp) {
  var cfg=DOORS_MONTHLY_FINALIZER_V16;
  var el=doorsFinalizerGetElementV16_(slide,cfg.lpTableId);
  if (el.getPageElementType() !== SlidesApp.PageElementType.TABLE) {
    throw new Error('LP table target is not TABLE: '+cfg.lpTableId);
  }
  var table=el.asTable();
  if (table.getNumRows() !== 11 || table.getNumColumns() !== 7) {
    throw new Error('LP table contract changed: '+table.getNumRows()+'x'+table.getNumColumns());
  }
  var headers=['ページ',lp.currentMonthLabel+'Click',lp.prevMonthLabel+'Click','差',lp.currentMonthLabel+'Imp','CTR','順位'];
  for (var c=0;c<7;c++) table.getCell(0,c).getText().setText(headers[c]);
  lp.rows.forEach(function(x,i){
    var r=i+1;
    var prevPos = x.prevPosition;
    var posDelta = (isFinite(prevPos) && isFinite(x.position)) ? prevPos-x.position : null;
    var posText = doorsFinalizerFixedV16_(x.position,2)+'位';
    if (posDelta != null && Math.abs(posDelta) >= 0.005) {
      posText += '\n' + (posDelta>0?'↑':'↓') + Math.abs(posDelta).toFixed(2);
    }
    var vals=[x.title,doorsFinalizerCommaV16_(x.clicks),doorsFinalizerCommaV16_(x.prevClicks),doorsFinalizerSignedIntV16_(x.delta),doorsFinalizerCommaV16_(x.impressions),doorsFinalizerPercentV16_(x.ctr),posText];
    for(var cc=0;cc<7;cc++) table.getCell(r,cc).getText().setText(vals[cc]);
    doorsFinalizerSetDeltaCellFillV16_(table.getCell(r,3),x.delta);
    doorsFinalizerSetDeltaCellFillV16_(table.getCell(r,6),posDelta==null?0:posDelta);
  });
}

function doorsFinalizerSetDeltaCellFillV16_(cell,delta){
  var fill=cell.getFill();
  if(delta<0) fill.setSolidFill('#FFEBEB');
  else if(delta>0) fill.setSolidFill('#E8F2FF');
  else fill.setSolidFill('#FFFFFF');
}
function doorsFinalizerCommaV16_(n){ return Math.round(Number(n)||0).toLocaleString('en-US'); }
function doorsFinalizerPercentV16_(v){
  var n=Number(v); if(!isFinite(n)) return '—';
  if(n>1) n=n/100; return (n*100).toFixed(2)+'%';
}
function doorsFinalizerFixedV16_(n,d){ var x=Number(n); return isFinite(x)?x.toFixed(d):'—'; }


function doorsFinalizerRepairVerticalFlipV16_(slide, objectId) {
  var el=doorsFinalizerGetElementV16_(slide, objectId);
  var t=el.getTransform();
  if (!t || Number(t.getScaleY()) >= 0) return false;
  if (typeof SlidesApp.newAffineTransformBuilder !== 'function') {
    throw new Error('Cannot repair negative scaleY: AffineTransformBuilder unavailable.');
  }
  var b=SlidesApp.newAffineTransformBuilder()
    .setScaleX(t.getScaleX()).setScaleY(Math.abs(t.getScaleY()))
    .setShearX(t.getShearX()).setShearY(t.getShearY())
    .setTranslateX(t.getTranslateX()).setTranslateY(t.getTranslateY());
  el.setTransform(b.build());
  return true;
}

function doorsFinalizerSetShapeTextV16_(slide,id,text){
  var el=doorsFinalizerGetElementV16_(slide,id);
  var shape=el.asShape();
  shape.getText().setText(String(text));
}
function doorsFinalizerGetElementV16_(slide,id){
  var found=null;
  slide.getPageElements().some(function(el){ if(el.getObjectId()===id){found=el;return true;} return false; });
  if(!found) throw new Error('Missing page element '+id+' on '+slide.getObjectId());
  return found;
}
function doorsFinalizerPrevMonthV16_(month){
  doorsFinalizerAssertMonthV16_(month);
  var y=Number(month.slice(0,4)),m=Number(month.slice(5,7))-1;
  if(m===0){y--;m=12;} return String(y)+'-'+('0'+m).slice(-2);
}
function doorsFinalizerMonthPartsV16_(month){
  doorsFinalizerAssertMonthV16_(month); return {year:Number(month.slice(0,4)),month:Number(month.slice(5,7))};
}
function doorsFinalizerSignedIntV16_(n){ return (n>0?'+':'')+String(Math.round(n)); }
function doorsFinalizerAssertMonthV16_(month){
  if(!/^20\d{2}-(0[1-9]|1[0-2])$/.test(String(month||''))) throw new Error('Invalid reportMonth: '+month);
}

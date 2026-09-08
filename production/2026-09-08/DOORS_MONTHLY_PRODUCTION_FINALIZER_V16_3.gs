/**
 * DOORS Monthly Production Finalizer v1.6.3
 *
 * Production finalizer for the current 36-slide deck.
 * Adds chart-shape synchronization for P.6 / P.12 / P.14.
 * No data acquisition. No trigger changes. No LLMO changes.
 */
var DOORS_MONTHLY_FINALIZER_V16 = Object.freeze({
  version: '1.6.3',
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
  sessionsSlideId: 'p7',
  aiTrendSlideId: 'p14',
  readingSlideId: 'p13',
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

  [cfg.lpSlideId,cfg.keywordSlideId,cfg.sessionsSlideId,cfg.aiTrendSlideId,cfg.readingSlideId].forEach(function(id){
    if (!byId[id]) throw new Error('Missing required slide: ' + id);
  });

  var lp = doorsFinalizerBuildLpInsightV16_(ss, reportMonth);
  doorsFinalizerSetShapeTextV16_(byId[cfg.lpSlideId], cfg.lpTitleId, lp.title);
  doorsFinalizerSetShapeTextV16_(byId[cfg.lpSlideId], cfg.lpSubtitleId, lp.subtitle);
  doorsFinalizerSetShapeTextV16_(byId[cfg.lpSlideId], cfg.lpInsightId, lp.insight);
  doorsFinalizerWriteLpTableV16_(byId[cfg.lpSlideId], lp);

  var parts = doorsFinalizerMonthPartsV16_(reportMonth);
  doorsFinalizerSetShapeTextV16_(byId[cfg.keywordSlideId], cfg.keywordTitleId,'新規提案キーワード｜' + parts.year + '年' + parts.month + '月');
  doorsFinalizerSetShapeTextV16_(byId[cfg.keywordSlideId], cfg.keywordSubtitleId,reportMonth + ' レポート｜人が採用し、シートへ転記した候補のみ掲載');
  var transformRepair = doorsFinalizerRepairVerticalFlipV16_(byId[cfg.keywordSlideId], cfg.keywordNoteId);

  var chartSync = doorsFinalizerSyncChartsV162_(byId);

  var afterSlides = pres.getSlides();
  if (afterSlides.length !== cfg.expectedSlideCount) throw new Error('Finalizer regression: slide count changed.');
  var afterIds = afterSlides.map(function(s){ return s.getObjectId(); });
  var beforeIds = slides.map(function(s){ return s.getObjectId(); });
  if (JSON.stringify(afterIds) !== JSON.stringify(beforeIds)) throw new Error('Finalizer regression: slide order/objectIds changed.');

  return {status:'SUCCESS',version:cfg.version,reportMonth:reportMonth,lp:lp,keywordNoteTransformRepair:transformRepair,chartSync:chartSync,slideCount:afterSlides.length};
}

function diagnoseDoorsMonthlyProductionFinalizerV16(reportMonth) {
  doorsFinalizerAssertMonthV16_(reportMonth);
  var cfg = DOORS_MONTHLY_FINALIZER_V16;
  var ss = SpreadsheetApp.openById(cfg.spreadsheetId);
  var panel = ss.getSheetByName(cfg.panelSheet);
  var pres = SlidesApp.openById(cfg.presentationId);
  var slides = pres.getSlides();
  var ids = slides.map(function(s){ return s.getObjectId(); });
  var lpData; var lpError='';
  try { lpData = doorsFinalizerBuildLpInsightV16_(ss, reportMonth); }
  catch (e) { lpError = String(e && e.message ? e.message : e); }
  var checks = {
    panelPresent: !!panel,
    panelMonth: !!panel && String(panel.getRange('B3').getDisplayValue() || '').trim() === reportMonth,
    slideCount36: slides.length === cfg.expectedSlideCount,
    lpSlidePresent: ids.indexOf(cfg.lpSlideId) >= 0,
    keywordSlidePresent: ids.indexOf(cfg.keywordSlideId) >= 0,
    sessionsSlidePresent: ids.indexOf(cfg.sessionsSlideId) >= 0,
    aiTrendSlidePresent: ids.indexOf(cfg.aiTrendSlideId) >= 0,
    readingSlidePresent: ids.indexOf(cfg.readingSlideId) >= 0,
    lpDataReady: !lpError
  };
  var failed = Object.keys(checks).filter(function(k){ return checks[k] !== true; });
  return {status: failed.length ? 'BLOCK' : 'PASS',productionChanged:false,version:cfg.version,reportMonth:reportMonth,checks:checks,failed:failed,lpError:lpError,lp:lpData || null};
}

function doorsFinalizerSyncChartsV162_(byId) {
  return {
    sessions: doorsFinalizerSyncSessionsBarsV162_(byId.p7),
    aiTrend: doorsFinalizerSyncAiTrendV162_(byId.p14),
    readingRate: doorsFinalizerSyncReadingBarsV162_(byId.p13)
  };
}

function doorsFinalizerSyncSessionsBarsV162_(slide) {
  var barIds = ['p7_i5','p7_i8','p7_i11','p7_i14','p7_i17','p7_i20','p7_i23'];
  var labelIds = ['p7_i6','p7_i9','p7_i12','p7_i15','p7_i18','p7_i21','p7_i24'];
  var values = labelIds.map(function(id){ return doorsFinalizerParseNumberV162_(doorsFinalizerShapeTextV162_(slide,id)); });
  if (values.some(function(v){ return !isFinite(v) || v < 0; })) throw new Error('P.6 chart sync BLOCK: invalid session labels ' + JSON.stringify(values));
  var max = Math.max.apply(null, values);
  if (!(max > 0)) throw new Error('P.6 chart sync BLOCK: max sessions <= 0');
  var baseline=4754880, maxHeight=2300000, intrinsicHeight=3000000;
  barIds.forEach(function(id,i){
    var h=maxHeight*values[i]/max, top=baseline-h;
    doorsFinalizerSetElementTransformV162_(doorsFinalizerGetElementV16_(slide,id),{scaleY:h/intrinsicHeight,translateY:top});
    doorsFinalizerSetElementTransformV162_(doorsFinalizerGetElementV16_(slide,labelIds[i]),{translateY:Math.max(1450000,top-290000)});
  });
  return {status:'PASS',values:values,max:max};
}

function doorsFinalizerSyncAiTrendV162_(slide) {
  var valueIds = ['p14_v1','p14_v2','p14_v3','p14_v4','p14_v5','p14_v6','p14_v7'];
  var lineIds = ['p14_u_line1','p14_u_line2','p14_u_line3','p14_u_line4','p14_u_line5','p14_u_line6'];
  var tickIds = ['p14_y400','p14_y350','p14_y300','p14_y250'];
  var values = valueIds.map(function(id){ return doorsFinalizerParseNumberV162_(doorsFinalizerShapeTextV162_(slide,id)); });
  if (values.some(function(v){ return !isFinite(v) || v < 0; })) throw new Error('P.12 chart sync BLOCK: invalid AI/LLM labels ' + JSON.stringify(values));

  var rawMin=Math.min.apply(null,values), rawMax=Math.max.apply(null,values), tickStep=50;
  var yMin=Math.floor(rawMin/tickStep)*tickStep;
  var yMax=yMin+200;
  if (rawMax>yMax) { yMax=Math.ceil(rawMax/tickStep)*tickStep+tickStep; yMin=yMax-200; }
  if (rawMin<yMin) { yMin=Math.floor(rawMin/tickStep)*tickStep; yMax=yMin+200; }

  var plotTop=2390000, plotBottom=4490000, x0=3300000, dx=700000, intrinsic=3000000, labelOffset=250000;
  function yFor(v){ return plotTop + (yMax-v)/(yMax-yMin)*(plotBottom-plotTop); }
  var ys=values.map(yFor);

  lineIds.forEach(function(id,i){
    var y1=ys[i], y2=ys[i+1];
    doorsFinalizerSetElementTransformV162_(doorsFinalizerGetElementV16_(slide,id),{
      scaleX:dx/intrinsic,scaleY:(y2-y1)/intrinsic,shearX:0,shearY:0,translateX:x0+dx*i,translateY:y1
    });
  });

  valueIds.forEach(function(id,i){
    doorsFinalizerSetElementTransformV162_(doorsFinalizerGetElementV16_(slide,id),{translateY:ys[i]-labelOffset});
  });

  var tickValues=[yMax,yMax-50,yMax-100,yMax-150];
  tickIds.forEach(function(id,i){
    var v=tickValues[i];
    doorsFinalizerSetShapeTextV16_(slide,id,String(Math.round(v)));
    doorsFinalizerSetElementTransformV162_(doorsFinalizerGetElementV16_(slide,id),{translateY:yFor(v)});
  });
  return {status:'PASS',values:values,yMin:yMin,yMax:yMax,tickStep:tickStep};
}

function doorsFinalizerSyncReadingBarsV162_(slide) {
  var valueIds=['p13_i13','p13_i17','p13_i21'];
  var barIds=['p13_i12','p13_i16','p13_i20'];
  var values=valueIds.map(function(id){ return doorsFinalizerParsePercentV162_(doorsFinalizerShapeTextV162_(slide,id)); });
  if (values.some(function(v){ return !isFinite(v) || v < 0; })) throw new Error('P.14 chart sync BLOCK: invalid reading-rate labels ' + JSON.stringify(values));
  var max=Math.max.apply(null,values);
  if (!(max>0)) throw new Error('P.14 chart sync BLOCK: max reading rate <= 0');
  var fullScaleX=1.2192;
  barIds.forEach(function(id,i){ doorsFinalizerSetElementTransformV162_(doorsFinalizerGetElementV16_(slide,id),{scaleX:fullScaleX*values[i]/max}); });
  return {status:'PASS',values:values,max:max};
}

function doorsFinalizerSetElementTransformV162_(el,patch) {
  var t=el.getTransform();
  if (!t || typeof SlidesApp.newAffineTransformBuilder !== 'function') throw new Error('AffineTransformBuilder unavailable for ' + el.getObjectId());
  function choose(name,current){ return Object.prototype.hasOwnProperty.call(patch,name) ? patch[name] : current; }
  var b=SlidesApp.newAffineTransformBuilder()
    .setScaleX(choose('scaleX',t.getScaleX())).setScaleY(choose('scaleY',t.getScaleY()))
    .setShearX(choose('shearX',t.getShearX())).setShearY(choose('shearY',t.getShearY()))
    .setTranslateX(choose('translateX',t.getTranslateX())).setTranslateY(choose('translateY',t.getTranslateY()));
  el.setTransform(b.build());
}

function doorsFinalizerShapeTextV162_(slide,id){ var el=doorsFinalizerGetElementV16_(slide,id); return String(el.asShape().getText().asString() || '').trim(); }
function doorsFinalizerParseNumberV162_(text){ var s=String(text||'').replace(/,/g,''); var m=s.match(/-?\d+(?:\.\d+)?/); return m ? Number(m[0]) : NaN; }
function doorsFinalizerParsePercentV162_(text){ return doorsFinalizerParseNumberV162_(text); }

function doorsFinalizerBuildLpInsightV16_(ss, reportMonth) {
  var cfg=DOORS_MONTHLY_FINALIZER_V16;
  var prevMonth=doorsFinalizerPrevMonthV16_(reportMonth);
  var sh=ss.getSheetByName(cfg.gscPageSheet);
  if (!sh) throw new Error('Missing sheet: ' + cfg.gscPageSheet);
  var values=sh.getDataRange().getDisplayValues();
  if (values.length<2) throw new Error('GSC page SSOT is empty.');
  var h=values[0];
  function ix(name){ var i=h.indexOf(name); if(i<0) throw new Error('Missing GSC header: '+name); return i; }
  var cMonth=ix('month'),cRank=ix('rank'),cPage=ix('page'),cTitle=ix('title'),cClicks=ix('clicks'),cImp=ix('impressions'),cCtr=ix('ctr'),cPos=ix('position');
  var current=[],prevByUrl={};
  values.slice(1).forEach(function(r){
    var m=String(r[cMonth]||'').trim(),url=String(r[cPage]||'').trim(); if(!url)return;
    var obj={month:m,rank:Number(r[cRank]),url:url,title:String(r[cTitle]||'').trim(),clicks:Number(r[cClicks]),impressions:Number(r[cImp]),ctr:Number(r[cCtr]),position:Number(r[cPos])};
    if(m===prevMonth)prevByUrl[url]=obj;
    if(m===reportMonth && url.indexOf('/doors/contents/')>=0 && isFinite(obj.clicks))current.push(obj);
  });
  current.sort(function(a,b){return a.rank-b.rank;}); current=current.slice(0,10);
  if(current.length!==10) throw new Error('Expected 10 current /contents/ GSC pages; got '+current.length);
  current.forEach(function(x){ var p=prevByUrl[x.url]; x.prevClicks=p&&isFinite(p.clicks)?p.clicks:0; x.prevPosition=p&&isFinite(p.position)?p.position:NaN; x.delta=x.clicks-x.prevClicks; });
  var totalDelta=current.reduce(function(s,x){return s+x.delta;},0);
  var declines=current.filter(function(x){return x.delta<0;}),gains=current.filter(function(x){return x.delta>0;});
  var pool=(totalDelta<0?declines:gains).slice().sort(function(a,b){return Math.abs(b.delta)-Math.abs(a.delta);}).slice(0,3);
  var denom=(totalDelta<0?declines:gains).reduce(function(s,x){return s+Math.abs(x.delta);},0);
  var topAbs=pool.reduce(function(s,x){return s+Math.abs(x.delta);},0),share=denom?topAbs/denom:0;
  var approxTenth=Math.max(0,Math.min(10,Math.round(share*10))),jpMonth=Number(reportMonth.slice(5,7)),jpPrev=Number(prevMonth.slice(5,7));
  var title,insight;
  if(totalDelta<0&&pool.length){ title='流入LP｜クリック減の約'+approxTenth+'割が上位3ページに集中しました'; insight='主要10ページ中'+declines.length+'ページでクリックが減少。特に'+pool.map(function(x){return '「'+x.title+'」'+doorsFinalizerSignedIntV16_(x.delta);}).join('、')+'の3ページで減少幅の約'+Math.round(share*100)+'%を占めます。まず3ページの表示回数・CTR・順位をクエリ単位で確認します。'; }
  else if(totalDelta>0&&pool.length){ title='流入LP｜クリック増の約'+approxTenth+'割が上位3ページに集中しました'; insight='主要10ページ中'+gains.length+'ページでクリックが増加。特に'+pool.map(function(x){return '「'+x.title+'」'+doorsFinalizerSignedIntV16_(x.delta);}).join('、')+'の3ページで増加幅の約'+Math.round(share*100)+'%を占めます。伸長要因をクエリ・CTR・順位の順で確認します。'; }
  else { title='流入LP｜主要10ページの前月差は概ね横ばいです'; insight='主要10ページ全体のクリック差は小幅です。個別ページの表示回数・CTR・順位を確認し、変化が大きいページだけを優先して確認します。'; }
  return {title:title,subtitle:'Search Consoleのページ別実績を、'+jpMonth+'月と'+jpPrev+'月の同一条件で比較します。',insight:insight,totalDelta:totalDelta,declines:declines.length,gains:gains.length,top3:pool.map(function(x){return {title:x.title,delta:x.delta};}),share:share,currentMonth:reportMonth,prevMonth:prevMonth,currentMonthLabel:jpMonth+'月',prevMonthLabel:jpPrev+'月',rows:current};
}

function doorsFinalizerWriteLpTableV16_(slide,lp){
  var cfg=DOORS_MONTHLY_FINALIZER_V16,el=doorsFinalizerGetElementV16_(slide,cfg.lpTableId);
  if(el.getPageElementType()!==SlidesApp.PageElementType.TABLE) throw new Error('LP table target is not TABLE: '+cfg.lpTableId);
  var table=el.asTable(); if(table.getNumRows()!==11||table.getNumColumns()!==7) throw new Error('LP table contract changed: '+table.getNumRows()+'x'+table.getNumColumns());
  var headers=['ページ',lp.currentMonthLabel+'Click',lp.prevMonthLabel+'Click','差',lp.currentMonthLabel+'Imp','CTR','順位'];
  for(var c=0;c<7;c++) table.getCell(0,c).getText().setText(headers[c]);
  lp.rows.forEach(function(x,i){ var r=i+1,prevPos=x.prevPosition,posDelta=(isFinite(prevPos)&&isFinite(x.position))?prevPos-x.position:null; var posText=doorsFinalizerFixedV16_(x.position,2)+'位'; if(posDelta!=null&&Math.abs(posDelta)>=0.005)posText+='\n'+(posDelta>0?'↑':'↓')+Math.abs(posDelta).toFixed(2); var vals=[x.title,doorsFinalizerCommaV16_(x.clicks),doorsFinalizerCommaV16_(x.prevClicks),doorsFinalizerSignedIntV16_(x.delta),doorsFinalizerCommaV16_(x.impressions),doorsFinalizerPercentV16_(x.ctr),posText]; for(var cc=0;cc<7;cc++)table.getCell(r,cc).getText().setText(vals[cc]); doorsFinalizerSetDeltaCellFillV16_(table.getCell(r,3),x.delta); doorsFinalizerSetDeltaCellFillV16_(table.getCell(r,6),posDelta==null?0:posDelta); });
}
function doorsFinalizerSetDeltaCellFillV16_(cell,delta){ var fill=cell.getFill(); if(delta<0)fill.setSolidFill('#FFEBEB'); else if(delta>0)fill.setSolidFill('#E8F2FF'); else fill.setSolidFill('#FFFFFF'); }
function doorsFinalizerCommaV16_(n){return Math.round(Number(n)||0).toLocaleString('en-US');}
function doorsFinalizerPercentV16_(v){var n=Number(v);if(!isFinite(n))return '—';if(n>1)n=n/100;return(n*100).toFixed(2)+'%';}
function doorsFinalizerFixedV16_(n,d){var x=Number(n);return isFinite(x)?x.toFixed(d):'—';}
function doorsFinalizerRepairVerticalFlipV16_(slide,objectId){var el=doorsFinalizerGetElementV16_(slide,objectId),t=el.getTransform();if(!t||Number(t.getScaleY())>=0)return false;if(typeof SlidesApp.newAffineTransformBuilder!=='function')throw new Error('Cannot repair negative scaleY: AffineTransformBuilder unavailable.');var b=SlidesApp.newAffineTransformBuilder().setScaleX(t.getScaleX()).setScaleY(Math.abs(t.getScaleY())).setShearX(t.getShearX()).setShearY(t.getShearY()).setTranslateX(t.getTranslateX()).setTranslateY(t.getTranslateY());el.setTransform(b.build());return true;}
function doorsFinalizerSetShapeTextV16_(slide,id,text){var el=doorsFinalizerGetElementV16_(slide,id);el.asShape().getText().setText(String(text));}
function doorsFinalizerGetElementV16_(slide,id){var found=null;slide.getPageElements().some(function(el){if(el.getObjectId()===id){found=el;return true;}return false;});if(!found)throw new Error('Missing page element '+id+' on '+slide.getObjectId());return found;}
function doorsFinalizerPrevMonthV16_(month){doorsFinalizerAssertMonthV16_(month);var y=Number(month.slice(0,4)),m=Number(month.slice(5,7))-1;if(m===0){y--;m=12;}return String(y)+'-'+('0'+m).slice(-2);}
function doorsFinalizerMonthPartsV16_(month){doorsFinalizerAssertMonthV16_(month);return{year:Number(month.slice(0,4)),month:Number(month.slice(5,7))};}
function doorsFinalizerSignedIntV16_(n){return(n>0?'+':'')+String(Math.round(n));}
function doorsFinalizerAssertMonthV16_(month){if(!/^20\d{2}-(0[1-9]|1[0-2])$/.test(String(month||'')))throw new Error('Invalid reportMonth: '+month);}

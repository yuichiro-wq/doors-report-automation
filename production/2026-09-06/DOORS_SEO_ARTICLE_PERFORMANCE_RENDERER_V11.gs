/**
 * DOORS SEO article-performance renderer v1.0 - RELEASE CANDIDATE / NOT DEPLOYED
 *
 * Purpose:
 * - Render P17-P22 (slide objectIds p20..p21_detail_3) from Production SSOT.
 * - Stable population: Article Master order, 新規18 + リライト18 with nonblank publish/rewrite month.
 * - Metrics: current/previous PV + page-level GSC clicks/impressions/CTR/average position.
 * - Does NOT claim page-level GSC average position is an exact target-query rank.
 *
 * Safety:
 * - Default target is the verified canary deck only.
 * - Production deck is blocked unless allowProduction === true.
 * - Preflight completes before first Slides write.
 * - Only six known slide objectIds and renderer-owned/reference-image objects are modified.
 * - Page order and all slide objectIds are preserved.
 */

var DOORS_SEO_PERF_RENDER_V1 = Object.freeze({
  spreadsheetId: '1q1B0JGdLaBOCkDWa4yLNlSFozzf8eaky2sTvTL9BAgo',
  productionPresentationId: '17qy7JYXMWmjhUbPrmvpURKOkT1TqsF8-jkM74jxw5-4',
  canaryPresentationId: '1GSPeWBo04LK_sJgsnORQqPy9J41Va8X-571mj9wpWkg',
  articleMasterSheet: '記事マスター',
  performanceSheet: '記事別パフォーマンス',
  queryKpiSheet: '対策クエリKPI',
  expectedNew: 18,
  expectedRewrite: 18,
  slides: Object.freeze([
    {slideId:'p20',           type:'新規', start:0,  titleId:'p20_detail_t1', subtitleId:'p20_detail_s1', bodyId:'p20_detail_b1', refImageId:'p17_refdesign_img'},
    {slideId:'p20_detail_2',  type:'新規', start:6,  titleId:'p20_detail_t2', subtitleId:'p20_detail_s2', bodyId:'p20_detail_b2', refImageId:'p18_refdesign_img'},
    {slideId:'p20_detail_3',  type:'新規', start:12, titleId:'p20_detail_t3', subtitleId:'p20_detail_s3', bodyId:'p20_detail_b3', refImageId:'p19_refdesign_img'},
    {slideId:'p21',           type:'リライト', start:0,  titleId:'p21_detail_t1', subtitleId:'p21_detail_s1', bodyId:'p21_detail_b1', refImageId:'p20_refdesign_img'},
    {slideId:'p21_detail_2',  type:'リライト', start:6,  titleId:'p21_detail_t2', subtitleId:'p21_detail_s2', bodyId:'p21_detail_b2', refImageId:'p21_refdesign_img'},
    {slideId:'p21_detail_3',  type:'リライト', start:12, titleId:'p21_detail_t3', subtitleId:'p21_detail_s3', bodyId:'p21_detail_b3', refImageId:'p22_refdesign_img'}
  ]),
  legacyStyleMarkerPrefix: 'DOORS_SEO_NEGATIVE_STYLE_PROD_V1|',
  rowsPerSlide: 6
});

function renderDoorsSeoArticlePerformanceCanaryV1(reportMonth) {
  return renderDoorsSeoArticlePerformanceV1(
    reportMonth,
    DOORS_SEO_PERF_RENDER_V1.canaryPresentationId,
    false
  );
}

function renderDoorsSeoArticlePerformanceProductionV1(reportMonth) {
  return renderDoorsSeoArticlePerformanceV1(
    reportMonth,
    DOORS_SEO_PERF_RENDER_V1.productionPresentationId,
    true
  );
}

function renderDoorsSeoArticlePerformanceV1(reportMonth, presentationId, allowProduction) {
  doorsSeoPerfAssertMonthV1_(reportMonth);
  if (!presentationId) throw new Error('presentationId required');
  if (presentationId === DOORS_SEO_PERF_RENDER_V1.productionPresentationId && allowProduction !== true) {
    throw new Error('Production render blocked: explicit allowProduction=true required.');
  }

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(3000)) throw new Error('Another SEO performance render is running.');
  try {
    var state = doorsSeoPerfPreflightV1_(reportMonth, presentationId);
    var pres = SlidesApp.openById(presentationId);

    DOORS_SEO_PERF_RENDER_V1.slides.forEach(function(def) {
      var rows = (def.type === '新規' ? state.newRows : state.rewriteRows)
        .slice(def.start, def.start + DOORS_SEO_PERF_RENDER_V1.rowsPerSlide);
      doorsSeoPerfRenderOneSlideV1_(pres, def, rows, reportMonth);
    });

    pres.saveAndClose();
    return doorsSeoPerfVerifyV1_(reportMonth, presentationId, state.beforeSlideIds);
  } finally {
    lock.releaseLock();
  }
}

function doorsSeoPerfPreflightV1_(reportMonth, presentationId) {
  var ss = SpreadsheetApp.openById(DOORS_SEO_PERF_RENDER_V1.spreadsheetId);
  var master = doorsSeoPerfObjectsV1_(ss.getSheetByName(DOORS_SEO_PERF_RENDER_V1.articleMasterSheet));
  var perf = doorsSeoPerfObjectsV1_(ss.getSheetByName(DOORS_SEO_PERF_RENDER_V1.performanceSheet))
    .filter(function(r){ return doorsSeoPerfMonthV1_(r.month) === reportMonth; });

  var perfById = {};
  perf.forEach(function(r){ perfById[String(r.article_id || '')] = r; });

  var queryKpi = doorsSeoPerfObjectsV1_(ss.getSheetByName(DOORS_SEO_PERF_RENDER_V1.queryKpiSheet))
    .filter(function(r){ return doorsSeoPerfMonthV1_(r.month) === reportMonth; });
  var queryByUrl = {};
  queryKpi.forEach(function(r){
    queryByUrl[String(r.article_url || '').trim()] = r;
  });

  var newMaster = master.filter(function(r){ return String(r.article_type || '') === '新規'; });
  var rewriteMaster = master.filter(function(r){
    return String(r.article_type || '') === 'リライト' &&
      String(r.publish_or_rewrite_month || '').trim() !== '';
  });

  if (newMaster.length !== DOORS_SEO_PERF_RENDER_V1.expectedNew) {
    throw new Error('New population mismatch: ' + newMaster.length);
  }
  if (rewriteMaster.length !== DOORS_SEO_PERF_RENDER_V1.expectedRewrite) {
    throw new Error('Rewrite population mismatch: ' + rewriteMaster.length);
  }

  function join(row, idx) {
    var id = String(row.article_id || '');
    var p = perfById[id];
    if (!p) throw new Error('Performance row missing: ' + reportMonth + ' / ' + id);
    return {
      no: idx + 1,
      articleId: id,
      label: doorsSeoPerfShortLabelV1_(row.title, id),
      currentPv: Number(p.page_views || 0),
      previousPv: p.previous_month_page_views === '' ? null : Number(p.previous_month_page_views),
      changeRate: p.change_rate === '' ? null : doorsSeoPerfPercentV1_(p.change_rate),
      gscClicks: Number(p.gsc_clicks || 0),
      gscImpressions: Number(p.gsc_impressions || 0),
      gscCtr: doorsSeoPerfPercentV1_(p.gsc_ctr),
      gscPosition: p.gsc_position === '' ? null : Number(p.gsc_position),
      targetQuery: queryByUrl[String(row.article_url || '').trim()] || null
    };
  }

  var newRows = newMaster.map(join);
  var rewriteRows = rewriteMaster.map(join);

  var pres = SlidesApp.openById(presentationId);
  var beforeSlideIds = pres.getSlides().map(function(s){ return s.getObjectId(); });
  var byId = {};
  pres.getSlides().forEach(function(s){ byId[s.getObjectId()] = s; });

  DOORS_SEO_PERF_RENDER_V1.slides.forEach(function(def) {
    var slide = byId[def.slideId];
    if (!slide) throw new Error('Slide missing: ' + def.slideId);
    [def.titleId, def.subtitleId, def.bodyId].forEach(function(id) {
      if (!doorsSeoPerfElementByIdV1_(slide, id)) throw new Error('Element missing: ' + id);
    });
  });

  return {newRows:newRows, rewriteRows:rewriteRows, beforeSlideIds:beforeSlideIds};
}

function doorsSeoPerfRenderOneSlideV1_(pres, def, rows, reportMonth) {
  var slide = pres.getSlides().filter(function(s){ return s.getObjectId() === def.slideId; })[0];
  if (!slide) throw new Error('Slide missing: ' + def.slideId);

  // Remove the old full-slide reference image which otherwise hides editable content.
  var ref = doorsSeoPerfElementByIdV1_(slide, def.refImageId);
  if (ref) ref.remove();

  // Remove only old renderer-owned change overlays; never touch other shapes.
  slide.getPageElements().slice().forEach(function(el) {
    var title = '';
    try { title = String(el.getTitle() || ''); } catch (e) {}
    if (title.indexOf(DOORS_SEO_PERF_RENDER_V1.legacyStyleMarkerPrefix) === 0) el.remove();
  });

  var page = Math.floor(def.start / 6) + 1;
  var title = def.type === '新規'
    ? '新規記事｜' + doorsSeoPerfMonthLabelV1_(reportMonth) + 'PV・前月差・対策KW順位（' + page + '/3）'
    : 'リライト記事｜' + doorsSeoPerfMonthLabelV1_(reportMonth) + 'PV・前月差・GSC平均順位（' + page + '/3）';
  var subtitle = def.type === '新規'
    ? '記事｜' + doorsSeoPerfMonthLabelV1_(reportMonth) + 'PV｜前月PV｜増減｜対策KW順位'
    : '記事｜' + doorsSeoPerfMonthLabelV1_(reportMonth) + 'PV｜前月PV｜増減｜GSC平均順位';

  doorsSeoPerfSetTextV1_(slide, def.titleId, title);
  doorsSeoPerfSetTextV1_(slide, def.subtitleId, subtitle);
  doorsSeoPerfSetTextV1_(slide, def.bodyId, doorsSeoPerfBodyV1_(rows, def.type));

  // Color the percentage token itself. Cell-fill styling can be layered separately
  // without changing the data contract.
  doorsSeoPerfColorBodyChangesV1_(slide, def.bodyId, rows);
}

function doorsSeoPerfBodyV1_(rows, articleType) {
  return rows.map(function(r) {
    var prev = r.previousPv == null ? '—' : doorsSeoPerfIntV1_(r.previousPv);
    var change = r.changeRate == null ? '—' : doorsSeoPerfSignedPctV1_(r.changeRate);
    var rank;
    if (articleType === '新規') {
      rank = r.targetQuery && String(r.targetQuery.data_status || '') === 'OK' &&
        String(r.targetQuery.target_position || '').trim() !== ''
        ? Number(r.targetQuery.target_position).toFixed(2)
        : '—';
    } else {
      rank = r.gscPosition == null ? '—' : r.gscPosition.toFixed(1);
    }
    return ('0' + r.no).slice(-2) + ' ' + r.label + '｜' +
      doorsSeoPerfIntV1_(r.currentPv) + '｜' + prev + '｜' + change + '｜' + rank;
  }).join('\n');
}

function doorsSeoPerfColorBodyChangesV1_(slide, bodyId, rows) {
  var el = doorsSeoPerfElementByIdV1_(slide, bodyId);
  var text = el.asShape().getText();
  var full = text.asString();
  rows.forEach(function(r) {
    if (r.changeRate == null || r.previousPv == null) return;
    var token = doorsSeoPerfSignedPctV1_(r.changeRate);
    var idx = full.indexOf(token);
    if (idx < 0) return;
    var color = r.changeRate > 0 ? '#1E60C2' : (r.changeRate < 0 ? '#CC0000' : '#1F2A37');
    text.getRange(idx, idx + token.length).getTextStyle().setForegroundColor(color).setBold(r.changeRate !== 0);
  });
}

function doorsSeoPerfVerifyV1_(reportMonth, presentationId, beforeSlideIds) {
  var pres = SlidesApp.openById(presentationId);
  var after = pres.getSlides().map(function(s){ return s.getObjectId(); });
  if (JSON.stringify(after) !== JSON.stringify(beforeSlideIds)) {
    throw new Error('Slide objectId/order regression detected.');
  }
  DOORS_SEO_PERF_RENDER_V1.slides.forEach(function(def) {
    var slide = pres.getSlides().filter(function(s){ return s.getObjectId() === def.slideId; })[0];
    var title = doorsSeoPerfElementByIdV1_(slide, def.titleId).asShape().getText().asString();
    var body = doorsSeoPerfElementByIdV1_(slide, def.bodyId).asShape().getText().asString();
    if (title.indexOf(doorsSeoPerfMonthLabelV1_(reportMonth)) < 0) throw new Error('Title month mismatch: ' + def.slideId);
    if (body.split('\n').filter(Boolean).length !== 6) throw new Error('Expected 6 body lines: ' + def.slideId);
    if (doorsSeoPerfElementByIdV1_(slide, def.refImageId)) throw new Error('Reference image still exists: ' + def.refImageId);
  });
  return {
    status:'PASS',
    reportMonth:reportMonth,
    presentationId:presentationId,
    slideCount:after.length,
    objectIdSequenceUnchanged:true,
    updatedSlides:DOORS_SEO_PERF_RENDER_V1.slides.map(function(x){return x.slideId;})
  };
}

function doorsSeoPerfObjectsV1_(sheet) {
  if (!sheet) throw new Error('Sheet missing');
  var values = sheet.getDataRange().getDisplayValues();
  var h = values.shift();
  return values.filter(function(r){return r.some(function(v){return v !== '';});}).map(function(r){
    var o={}; h.forEach(function(k,i){o[String(k)]=r[i];}); return o;
  });
}
function doorsSeoPerfElementByIdV1_(slide,id) {
  var a=slide.getPageElements(); for(var i=0;i<a.length;i++) if(a[i].getObjectId()===id) return a[i]; return null;
}
function doorsSeoPerfSetTextV1_(slide,id,value) {
  var el=doorsSeoPerfElementByIdV1_(slide,id); if(!el) throw new Error('Element missing: '+id);
  el.asShape().getText().setText(String(value));
}
function doorsSeoPerfPercentV1_(v) {
  var s=String(v==null?'':v).replace(/,/g,'').trim(); if(!s)return null;
  if(s.indexOf('%')>=0)return Number(s.replace('%',''))/100;
  var n=Number(s); return isFinite(n)?n:null;
}
function doorsSeoPerfSignedPctV1_(ratio) {
  if(ratio==null)return '—'; var p=ratio*100; return (p>0?'+':'')+p.toFixed(1)+'%';
}
function doorsSeoPerfIntV1_(n) { return Math.round(Number(n||0)).toLocaleString('en-US'); }
function doorsSeoPerfMonthV1_(v) {
  var s=String(v||''); var m=s.match(/^(\d{4})[-\/]?(\d{1,2})/);
  return m?m[1]+'-'+('0'+m[2]).slice(-2):s;
}
function doorsSeoPerfMonthLabelV1_(m) { return Number(String(m).split('-')[1])+'月'; }
function doorsSeoPerfAssertMonthV1_(m) {
  if(!/^\d{4}-(0[1-9]|1[0-2])$/.test(String(m))) throw new Error('reportMonth must be YYYY-MM: '+m);
}
function doorsSeoPerfShortLabelV1_(title,id) {
  var labels = {
    'seminar-break_away_from_0-hit':'EC 0件',
    'seminar_retail_data_ai':'小売データ×AI',
    'dx_consulting':'DXコンサル',
    'multimodal-ai-examples':'マルチモーダルAI事例',
    'manufacturing-ai-agents-use-cases':'製造業AIエージェント',
    'understanding-rag':'RAG',
    'what-is-vertical-ai-agent':'Vertical AI',
    'best-ai-agent-frameworks-comparison':'AIエージェントFW比較',
    'multi-agent-system':'マルチエージェント',
    'multi_modal_ai':'マルチモーダルLLM',
    'physical_ai':'フィジカルAI',
    'about_asi':'ASI',
    'about_soverignai':'ソブリンAI',
    'about_digital_immune_system':'デジタル免疫',
    'about_multimodal_ai':'マルチモーダルAI',
    'ai_agent_use_cases':'AIエージェント活用事例',
    'difference_ai_agents_and_gen-ai':'AIエージェントと生成AIの違い',
    'ai_agent_security':'AIエージェントセキュリティ',
    'about_generative_ai':'生成AIとは','dx_it':'DXとは','dx_white_paper_2023':'DX戦略',
    'about_manufacturing_dx':'製造業DX','about_agi':'AGI','about_scm':'SCM','about_gx':'GX',
    'about_esg':'ESG','about_dynamic_pricing':'ダイナミックプライシング',
    'data_driven_management':'データドリブン経営','dx_cdo':'CDO',
    '01_about_data_scientist':'データサイエンティスト','01_about_llm':'LLM',
    'dx_ai_2021_1':'DX×AI','machine_learning':'機械学習','dx_theory_logistics':'物流DX',
    'dx_action_plan':'DX推進指標','dx_humanresources_2021_1':'DX人材'
  };
  return labels[id] || String(title || id);
}

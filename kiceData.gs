/*******************************************************
 * kiceData.gs — 엔진·코퍼스 빌더가 시트를 읽고 쓰는 명령 (kice-2.0.0 · 재구축 계획서 v5 §7-1, §17 S3)
 *
 *   corpus      GET   Corpus·ExprDict 를 쪽 단위로 읽는다 (엔진이 검색용으로 통째 받아 둔다)
 *   prompts     GET   prompt 시트의 활성 프롬프트 + Rule 시트의 활성 규칙
 *   image       GET   Drive 그림 파일을 base64 로
 *   dsread      GET   사본 「KICE_latex」 Data_DS / 구 「기출과 윤문」 Data_Latex 읽기
 *   corpus_put  POST  Corpus·AnswerKey·ExprDict·EvalSet 에 키 기준 upsert
 *
 * 시트를 쓰는 주체는 GAS 하나다(v1 A1). 엔진은 Sheets API 를 직접 쓰지 않는다.
 * 이 명령들은 배치 상태(KICE_AUTO_STATE_V2)와 무관하다.
 *******************************************************/

var KDATA_PAGE_DEFAULT = 200;
var KDATA_PAGE_MAX     = 500;
var KDATA_ETAG_PROP    = 'KDATA_ETAG';                 // corpus_put 이 쓸 때마다 갱신
var KDATA_IMAGE_MAX    = 8 * 1024 * 1024;              // image 한 장 상한 (base64 전 바이트)
var KDATA_READ_SHEETS  = ['Corpus', 'ExprDict', 'AnswerKey', 'EvalSet', 'Feedback', 'Opinions', 'QueCtx'];
var KDATA_PUT_SHEETS   = { Corpus: '문항ID', AnswerKey: '문항ID', ExprDict: '표현', EvalSet: '문항ID' };   // 시트 → 키 열
var KDATA_LOCK_HEADER  = '잠금';
var KDATA_LOCKED_COLS  = ['LaTeX', '정답', '대단원'];  // 잠금 행에서 force 없이는 못 바꾸는 열 (v5 §18 N13)
var KDATA_RULE_SHEET   = 'Rule';
var KDATA_PROMPT_SHEET = 'prompt';


/* =========================================================
 * 공용
 * ========================================================= */

function kdata_int_(v, def, min, max) {
  var n = parseInt(v, 10);
  if (!isFinite(n)) n = def;
  if (min !== undefined && n < min) n = min;
  if (max !== undefined && n > max) n = max;
  return n;
}

function kdata_isTrue_(v) {
  return v === true || v === 1 || String(v).toUpperCase() === 'TRUE' || String(v) === '1';
}

/** 셀 값을 JSON 으로 내보낼 값으로 — 날짜는 ISO 문자열 */
function kdata_out_(v) {
  if (v instanceof Date) return v.toISOString();
  return v;
}

function kdata_etag_(sh) {
  var t = PropertiesService.getScriptProperties().getProperty(KDATA_ETAG_PROP) || '0';
  return t + ':' + (sh ? sh.getLastRow() : 0);
}

/**
 * 시트 한 장을 쪽 단위로 읽는다.
 *   opts: {offset, limit, cols(쉼표 헤더 이름), filterCol(1-base), keywords[]}
 * 반환 {headers, rows, total, offset, next}  — offset 은 데이터 행 기준 0-base(2행이 0)
 */
function kdata_readPage_(sh, opts) {
  var lastRow = sh.getLastRow(), lastCol = Math.max(1, sh.getLastColumn());
  var head = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h == null ? '' : h).trim(); });
  var pick = [];
  if (opts.cols) {
    String(opts.cols).split(',').map(function (s) { return s.trim(); }).filter(Boolean).forEach(function (name) {
      var i = head.indexOf(name);
      if (i < 0) throw new Error('없는 열 이름: ' + name);
      pick.push(i);
    });
  } else {
    for (var c = 0; c < head.length; c++) pick.push(c);
  }
  var headers = pick.map(function (i) { return head[i]; });
  var total = Math.max(0, lastRow - 1);
  var offset = kdata_int_(opts.offset, 0, 0);
  var limit = kdata_int_(opts.limit, KDATA_PAGE_DEFAULT, 1, KDATA_PAGE_MAX);

  // 키워드 필터가 있으면 그 열만 먼저 읽어 걸린 행 번호를 고른다
  if (opts.keywords && opts.keywords.length && total > 0) {
    var fcol = kdata_int_(opts.filterCol, 1, 1, lastCol);
    var keys = sh.getRange(2, fcol, total, 1).getValues();
    var match = kice_keyMatcher_(opts.keywords);
    var hit = [];
    for (var i = 0; i < keys.length; i++) { var k = String(keys[i][0] == null ? '' : keys[i][0]).trim(); if (k && match(k)) hit.push(i + 2); }
    var slice = hit.slice(offset, offset + limit);
    var rowsOut = [];
    if (slice.length) {
      var minR = slice[0], maxR = slice[slice.length - 1];
      var block = sh.getRange(minR, 1, maxR - minR + 1, lastCol).getValues();
      slice.forEach(function (r) {
        var v = block[r - minR];
        rowsOut.push({ row: r, values: pick.map(function (ci) { return kdata_out_(v[ci]); }) });
      });
    }
    return { headers: headers, rows: rowsOut, total: hit.length, offset: offset,
             next: (offset + slice.length < hit.length) ? offset + slice.length : null };
  }

  var n = Math.max(0, Math.min(limit, total - offset));
  var out = [];
  if (n > 0) {
    var vals = sh.getRange(2 + offset, 1, n, lastCol).getValues();
    for (var j = 0; j < vals.length; j++) {
      out.push({ row: 2 + offset + j, values: pick.map(function (ci) { return kdata_out_(vals[j][ci]); }) });
    }
  }
  return { headers: headers, rows: out, total: total, offset: offset, next: (offset + n < total) ? offset + n : null };
}


/* =========================================================
 * corpus (GET)
 *   sheet(기본 Corpus) · offset · limit(기본 200, 최대 500) · cols(쉼표 헤더 이름, 생략하면 전부)
 * 응답 {ok, sheet, headers, rows:[{row, values[]}], total, offset, next, etag}
 * ========================================================= */
function kdata_apiCorpus_(p) {
  p = p || {};
  var name = String(p.sheet || 'Corpus');
  if (KDATA_READ_SHEETS.indexOf(name) < 0) return { ok: false, reason: '읽을 수 없는 시트: ' + name };
  var sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) return { ok: false, reason: '시트 없음: ' + name };
  var page = kdata_readPage_(sh, { offset: p.offset, limit: p.limit, cols: p.cols });
  return kice_merge_({ ok: true, sheet: name, etag: kdata_etag_(sh) }, page);
}


/* =========================================================
 * prompts (GET)
 * 응답 {ok, prompts:{key: value}, rules:[{from, to, note}], at}
 *   prompt 시트: A key · B value · C enabled      (enabled 가 TRUE 인 것만)
 *   Rule 시트  : A from · B to · C enabled · D note
 * ========================================================= */
function kdata_loadPrompts_() {
  var sh = SpreadsheetApp.getActive().getSheetByName(KDATA_PROMPT_SHEET);
  var out = {};
  if (!sh || sh.getLastRow() < 2) return out;
  var values = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();
  values.forEach(function (r) {
    var k = String(r[0] || '').trim();
    if (!k || !kdata_isTrue_(r[2])) return;
    out[k] = String(r[1] == null ? '' : r[1]);
  });
  return out;
}

/** Rule 시트의 활성 규칙 — 1.0.0 의 _loadRules_ 와 같은 선별·정렬 */
function kdata_loadRules_() {
  var sh = SpreadsheetApp.getActive().getSheetByName(KDATA_RULE_SHEET);
  if (!sh || sh.getLastRow() < 2) return [];
  var values = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues();
  var rules = [];
  values.forEach(function (r) {
    var from = String(r[0] || '').trim(), to = String(r[1] || '').trim(), note = String(r[3] || '').trim();
    if (!kdata_isTrue_(r[2]) || !from || from === to) return;
    if (to && to.indexOf(from) !== -1) return;
    rules.push({ from: from, to: to, note: note });
  });
  rules.sort(function (a, b) { return b.from.length - a.from.length; });
  return rules;
}

function kdata_apiPrompts_() {
  return { ok: true, prompts: kdata_loadPrompts_(), rules: kdata_loadRules_(), at: new Date().toISOString() };
}


/* =========================================================
 * image (GET)
 *   id(Drive 파일 ID) 또는 url(Drive 링크)
 * 응답 {ok, id, name, mime, bytes, base64}
 * ========================================================= */
function kdata_fileId_(s) {
  s = String(s || '');
  var m = s.match(/\/d\/([-\w]{25,})/) || s.match(/[?&]id=([-\w]{25,})/) || s.match(/^([-\w]{25,})$/);
  return m ? m[1] : '';
}

function kdata_apiImage_(p) {
  p = p || {};
  var id = kdata_fileId_(p.id || p.url);
  if (!id) return { ok: false, reason: 'id 또는 url 에서 Drive 파일 ID 를 찾지 못했습니다.' };
  var file = DriveApp.getFileById(id);
  var mime = String(file.getMimeType() || '');
  if (mime.indexOf('image/') !== 0 && mime !== 'application/pdf') return { ok: false, reason: '그림·PDF 가 아닌 파일: ' + mime };
  var size = file.getSize();
  if (size > KDATA_IMAGE_MAX) return { ok: false, reason: '파일이 너무 큽니다: ' + size + '바이트', bytes: size };
  var blob = file.getBlob();
  return { ok: true, id: id, name: file.getName(), mime: mime, bytes: size, base64: Utilities.base64Encode(blob.getBytes()) };
}


/* =========================================================
 * dsread (GET) — 읽기 전용
 *   src=latex  : 속성 LATEX_KICE_SS_ID 의 Data_DS      (사본 변환 결과)
 *   src=legacy : 속성 LEGACY_SS_ID 의 Data_Latex       (구 「기출과 윤문」 — 구 ID·대단원·#태그 승계용)
 *   keywords(쉼표) · col(키워드를 맞출 열 번호, 기본 1) · offset · limit · cols
 * 응답 {ok, src, sheet, headers, rows:[{row, values[]}], total, offset, next}
 * ========================================================= */
var KDATA_DS_SRC = {
  latex:  { prop: 'LATEX_KICE_SS_ID', sheet: 'Data_DS' },
  legacy: { prop: 'LEGACY_SS_ID',     sheet: 'Data_Latex' }
};

function kdata_apiDsread_(p) {
  p = p || {};
  var def = KDATA_DS_SRC[String(p.src || '')];
  if (!def) return { ok: false, reason: 'src 는 latex 또는 legacy' };
  var ssId = (PropertiesService.getScriptProperties().getProperty(def.prop) || '').trim();
  if (!ssId) return { ok: false, reason: '스크립트 속성 ' + def.prop + ' 미설정' };
  var sh = SpreadsheetApp.openById(ssId).getSheetByName(def.sheet);
  if (!sh) return { ok: false, reason: def.sheet + ' 시트 없음 (' + def.prop + ')' };
  var page = kdata_readPage_(sh, { offset: p.offset, limit: p.limit, cols: p.cols,
                                   keywords: kice_parseKeywords_(p.keywords), filterCol: p.col });
  return kice_merge_({ ok: true, src: String(p.src), sheet: def.sheet }, page);
}


/* =========================================================
 * corpus_put (POST) — 키 기준 upsert (v5 §17 S3)
 *   sheet  : Corpus | AnswerKey | ExprDict | EvalSet
 *   rows   : [{헤더 이름: 값, …}]  — 키 열(Corpus·AnswerKey·EvalSet = 문항ID, ExprDict = 표현)은 필수. 200행 이내
 *   force  : true 면 잠금 행의 보호 열도 바꾼다
 *   dryRun : true 면 쓰지 않고 바뀔 행 수만 돌려준다
 *
 * 규칙
 *  - 열은 위치가 아니라 **헤더 이름**으로 찾는다. 요청에 없는 열은 건드리지 않는다. 시트에 없는 열 이름은 거부.
 *  - Corpus 의 `잠금` 이 표시된 행은 LaTeX·정답·대단원 을 force 없이는 바꾸지 않는다(나머지 열은 바꾼다).
 *  - 셀 5만 자 한도를 넘는 값은 잘라 적고 truncated 에 알린다. 전문은 부르는 쪽이 보관한다.
 * 응답 {ok, sheet, updated, inserted, unchanged, lockedSkipped:[키…], truncated:[키:열…], dryRun}
 * ========================================================= */
function kdata_apiCorpusPut_(p) {
  p = p || {};
  var name = String(p.sheet || '');
  var keyHeader = KDATA_PUT_SHEETS[name];
  if (!keyHeader) return { ok: false, reason: 'sheet 는 ' + Object.keys(KDATA_PUT_SHEETS).join(' | ') };
  var rowsIn = p.rows;
  if (!Array.isArray(rowsIn) || !rowsIn.length) return { ok: false, reason: 'rows 가 비어 있습니다.' };
  if (rowsIn.length > 200) return { ok: false, reason: '한 번에 200행까지 (' + rowsIn.length + '행)' };
  var force = kdata_isTrue_(p.force), dryRun = kdata_isTrue_(p.dryRun);

  return kice_withLock_(function () {
    var sh = SpreadsheetApp.getActive().getSheetByName(name);
    if (!sh) return { ok: false, reason: '시트 없음: ' + name };
    var map = kice_headerMap_(sh);
    var keyCol = map[keyHeader];
    if (!keyCol) return { ok: false, reason: '키 열이 없습니다: ' + keyHeader };

    // 열 이름 검사 — 하나라도 모르는 이름이면 아무것도 쓰지 않는다
    var unknown = {};
    rowsIn.forEach(function (r) { for (var h in (r || {})) if (!map[h]) unknown[h] = true; });
    var unk = Object.keys(unknown);
    if (unk.length) return { ok: false, reason: '시트에 없는 열 이름: ' + unk.join(', ') };

    var width = sh.getLastColumn();
    var last = kice_lastDataRow_(sh, width);
    var data = (last >= 2) ? sh.getRange(2, 1, last - 1, width).getValues() : [];
    var index = {};
    for (var i = 0; i < data.length; i++) { var k = String(data[i][keyCol - 1] == null ? '' : data[i][keyCol - 1]).trim(); if (k && index[k] === undefined) index[k] = i; }
    var lockCol = map[KDATA_LOCK_HEADER] || 0;

    var touched = {}, appended = [], appendIndex = {};
    var updated = 0, inserted = 0, unchanged = 0, lockedSkipped = [], truncated = [], bad = 0;

    rowsIn.forEach(function (r) {
      r = r || {};
      var key = String(r[keyHeader] == null ? '' : r[keyHeader]).trim();
      if (!key) { bad++; return; }
      var isNew = index[key] === undefined && appendIndex[key] === undefined;
      var target;
      if (index[key] !== undefined) target = data[index[key]];
      else if (appendIndex[key] !== undefined) target = appended[appendIndex[key]];
      else { target = []; for (var c = 0; c < width; c++) target.push(''); }
      var locked = !isNew && lockCol && kice_hasValue_(target[lockCol - 1]) && String(target[lockCol - 1]).toUpperCase() !== 'FALSE';
      var changed = false, skippedLock = false;
      for (var h in r) {
        var col = map[h];
        var val = r[h];
        if (val === undefined) continue;
        if (val === null) val = '';
        if (typeof val === 'object') val = JSON.stringify(val);
        if (typeof val === 'string' && val.length > KICE_CELL_MAX) { val = kice_clip_(val); truncated.push(key + ':' + h); }
        if (locked && !force && KDATA_LOCKED_COLS.indexOf(h) >= 0) {
          if (String(target[col - 1]) !== String(val)) skippedLock = true;
          continue;
        }
        if (target[col - 1] !== val && String(target[col - 1]) !== String(val)) { target[col - 1] = val; changed = true; }
      }
      if (skippedLock) lockedSkipped.push(key);
      if (isNew) { appendIndex[key] = appended.length; appended.push(target); inserted++; }
      else if (changed) { if (index[key] !== undefined) touched[index[key]] = true; updated++; }
      else unchanged++;
    });

    if (!dryRun) {
      var tr = Object.keys(touched).map(Number).sort(function (a, b) { return a - b; });
      if (tr.length) {                                   // 바뀐 행들을 이웃끼리 묶어 한 번씩 쓴다
        var s = tr[0], e = tr[0];
        var flush = function (a, b) { sh.getRange(a + 2, 1, b - a + 1, width).setValues(data.slice(a, b + 1)); };
        for (var t = 1; t < tr.length; t++) {
          if (tr[t] === e + 1) { e = tr[t]; continue; }
          flush(s, e); s = e = tr[t];
        }
        flush(s, e);
      }
      if (appended.length) {
        var start = Math.max(2, last + 1);
        if (sh.getMaxRows() < start + appended.length - 1) sh.insertRowsAfter(sh.getMaxRows(), start + appended.length - 1 - sh.getMaxRows());
        sh.getRange(start, 1, appended.length, width).setValues(appended);
      }
      if (tr.length || appended.length) {
        PropertiesService.getScriptProperties().setProperty(KDATA_ETAG_PROP, String(Date.now()));
        kice_log_('corpus_put', name, 'upsert ' + updated + ' / insert ' + inserted + ' / 그대로 ' + unchanged +
                  (lockedSkipped.length ? ' / 잠금 보호 ' + lockedSkipped.length : '') + (truncated.length ? ' / 잘림 ' + truncated.length : ''));
      }
    }
    var out = { ok: true, sheet: name, updated: updated, inserted: inserted, unchanged: unchanged,
                lockedSkipped: lockedSkipped, truncated: truncated, dryRun: dryRun };
    if (bad) out.badRows = bad;
    return out;
  });
}

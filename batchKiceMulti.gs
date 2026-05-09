/*******************************************************
 * <배치 컨트롤: 3사 병렬 호출>
 *
 * - batch_startQueAuto(): 행 범위 입력 후 Claude/GPT/Gemini 동시 처리
 * - batch_continueQueAuto(): 이어달리기
 * - batch_stopQueAuto(): 수동 중지
 *
 * Que 시트 열 배치:
 *   A = ID, B = latex, C = chapter
 *   D = (비움)
 *   E = Claude edits (HTML)
 *   F = GPT edits (HTML)
 *   G = Gemini edits (HTML)
 *
 * 의존:
 *  - findSimilarFromB2()
 *  - review_rewriteAllProviders()  ← rewriteKiceMulti.gs
 *******************************************************/

var BATCH_SHEET_QUE    = 'Que';
var BATCH_SHEET_REVIEW = '문항검토';

var QUE_BATCH_KEY   = 'QUE_BATCH_STATE_V3';
var SAFE_RUN_MS     = 5 * 60 * 1000 + 20 * 1000;
var RESUME_AFTER_MS = 60 * 1000;
var PER_ROW_SLEEP_MS = 2000;

// Que 시트 열 번호
var COL_QUE_ID      = 1;  // A
var COL_QUE_LATEX   = 2;  // B
var COL_QUE_CHAPTER = 3;  // C
// D = 4 (비움)
var COL_QUE_CLAUDE  = 5;  // E
var COL_QUE_GPT     = 6;  // F
var COL_QUE_GEMINI  = 7;  // G


/** =========================
 * 1) 시작 — 행 범위 입력 (provider 선택 없음)
 * ========================= */
function batch_startQueAuto() {
  var ss = SpreadsheetApp.getActive();
  var ui = SpreadsheetApp.getUi();

  var shQue = ss.getSheetByName(BATCH_SHEET_QUE);
  var shRev = ss.getSheetByName(BATCH_SHEET_REVIEW);
  if (!shQue || !shRev) {
    ss.toast('Que 또는 문항검토 시트를 찾지 못했습니다.', '오류', 5);
    return;
  }

  var props = PropertiesService.getScriptProperties();
  if (props.getProperty(QUE_BATCH_KEY)) {
    ss.toast('이미 배치가 실행 중입니다. 중지 후 다시 시작하세요.', '배치', 5);
    return;
  }

  // 행 범위 입력
  var rowResp = ui.prompt(
    'Que 자동 배치 (Claude + GPT + Gemini 병렬)',
    '처리할 행을 입력해줘 (예: 2,5,7-10)',
    ui.ButtonSet.OK_CANCEL
  );
  if (rowResp.getSelectedButton() !== ui.Button.OK) return;

  var rows = _parseRowSpec_(String(rowResp.getResponseText() || '').trim());
  if (!rows.length) {
    ss.toast('유효한 행이 없습니다.', '중단', 5);
    return;
  }

  var state = {
    rows: rows,
    idx: 0,
    startedAt: new Date().toISOString()
  };
  props.setProperty(QUE_BATCH_KEY, JSON.stringify(state));

  _deleteTriggersByHandler_('batch_continueQueAuto');
  ss.toast('3사 병렬 | 총 ' + rows.length + '행 자동 처리 시작', '배치', 5);

  batch_continueQueAuto();
}


/** =========================
 * 2) 이어달리기
 * ========================= */
function batch_continueQueAuto() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;

  var startMs = Date.now();
  var ss = SpreadsheetApp.getActive();
  var props = PropertiesService.getScriptProperties();

  try {
    var raw = props.getProperty(QUE_BATCH_KEY);
    if (!raw) return;

    var state;
    try { state = JSON.parse(raw); } catch (e) {
      ss.toast('배치 상태 JSON 파싱 실패 → 중지합니다.', '오류', 6);
      _clearQueBatchState_(); _deleteTriggersByHandler_('batch_continueQueAuto');
      return;
    }

    var rows = Array.isArray(state.rows) ? state.rows : [];
    var idx  = Number(state.idx || 0);

    var shQue = ss.getSheetByName(BATCH_SHEET_QUE);
    var shRev = ss.getSheetByName(BATCH_SHEET_REVIEW);
    if (!shQue || !shRev) {
      ss.toast('Que 또는 문항검토 시트를 찾지 못했습니다. 배치 중지.', '오류', 6);
      _clearQueBatchState_(); _deleteTriggersByHandler_('batch_continueQueAuto');
      return;
    }

    // 안전장치: 강제 종료에 대비하여 이어달리기 트리거를 미리 예약
    _scheduleResumeTrigger_();

    while (idx < rows.length) {
      // 안전 시간 체크
      var elapsed = Date.now() - startMs;
      if (elapsed > SAFE_RUN_MS) {
        state.idx = idx;
        props.setProperty(QUE_BATCH_KEY, JSON.stringify(state));
        ss.toast(idx + '/' + rows.length + '까지 처리. 곧 이어서 실행됩니다.', '배치', 6);
        _scheduleResumeTrigger_();
        return;
      }

      var r = rows[idx];

      // 현재 진행 상태 저장 (강제 종료 방어)
      state.idx = idx;
      props.setProperty(QUE_BATCH_KEY, JSON.stringify(state));

      try {
        ss.toast('(' + (idx + 1) + '/' + rows.length + ') row ' + r + ' [3사 병렬]', '배치', 3);

        var id      = String(shQue.getRange(r, COL_QUE_ID).getDisplayValue() || '').trim();
        var latex   = String(shQue.getRange(r, COL_QUE_LATEX).getDisplayValue() || '').trim();
        var chapter = String(shQue.getRange(r, COL_QUE_CHAPTER).getDisplayValue() || '').trim();

        if (!latex) {
          shQue.getRange(r, COL_QUE_CLAUDE).setValue('');
          shQue.getRange(r, COL_QUE_GPT).setValue('');
          shQue.getRange(r, COL_QUE_GEMINI).setValue('');
          ss.toast('row ' + r + ': latex 비어있음 → 스킵', '배치', 3);
          idx++;
          continue;
        }

        // 문항검토 시트에 입력 세팅
        ss.setActiveSheet(shRev);
        shRev.getRange('B2').setValue(latex);
        shRev.getRange('C2').setValue(chapter);

        // 코드1: 유사문항 검색 (1회)
        findSimilarFromB2({ openViewer: false });

        // 코드2: 3사 LLM 병렬 호출
        var allResults = review_rewriteAllProviders();

        if (allResults.empty) {
          shQue.getRange(r, COL_QUE_CLAUDE).setValue('수정 구절 없음');
          shQue.getRange(r, COL_QUE_GPT).setValue('수정 구절 없음');
          shQue.getRange(r, COL_QUE_GEMINI).setValue('수정 구절 없음');
        } else {
          var refs = allResults.refs || [];

          // E열: Claude
          _writeProviderResult_(shQue, r, COL_QUE_CLAUDE, allResults.claude, refs);
          // F열: GPT
          _writeProviderResult_(shQue, r, COL_QUE_GPT, allResults.gpt, refs);
          // G열: Gemini
          _writeProviderResult_(shQue, r, COL_QUE_GEMINI, allResults.gemini, refs);
        }

        if (id) ss.toast('완료: ' + id + ' (row ' + r + ')', '배치', 2);
        if (PER_ROW_SLEEP_MS > 0) Utilities.sleep(PER_ROW_SLEEP_MS);

      } catch (errRow) {
        var msg = (errRow && errRow.message) ? errRow.message : String(errRow);
        // GAS 실행 시간 초과 감지
        if (msg.indexOf('제한') !== -1 || msg.indexOf('time') !== -1 || msg.indexOf('limit') !== -1) {
          state.idx = idx;
          props.setProperty(QUE_BATCH_KEY, JSON.stringify(state));
          _scheduleResumeTrigger_();
          ss.toast('시간 초과 감지 → row ' + r + '부터 이어서 실행 예정', '배치', 6);
          return;
        }
        ss.toast('row ' + r + ' 실패: ' + msg, '오류', 6);
        shQue.getRange(r, COL_QUE_CLAUDE).setValue('(실패) ' + msg);
        shQue.getRange(r, COL_QUE_GPT).setValue('(실패) ' + msg);
        shQue.getRange(r, COL_QUE_GEMINI).setValue('(실패) ' + msg);
      }

      idx++;
    }

    _clearQueBatchState_();
    _deleteTriggersByHandler_('batch_continueQueAuto');
    ss.toast('자동 배치 완료 ✅ (3사 병렬)', '배치', 6);

  } finally {
    lock.releaseLock();
  }
}


/** =========================
 * 3) 수동 중지
 * ========================= */
function batch_stopQueAuto() {
  var ss = SpreadsheetApp.getActive();
  _clearQueBatchState_();
  _deleteTriggersByHandler_('batch_continueQueAuto');
  ss.toast('자동 배치 중지됨', '배치', 5);
}


/* ===========================
 * provider별 결과를 Que 셀에 쓰기
 * =========================== */

function _writeProviderResult_(shQue, row, col, providerResult, refs) {
  if (!providerResult) {
    shQue.getRange(row, col).setValue('수정 구절 없음');
    return;
  }

  if (providerResult.error) {
    shQue.getRange(row, col).setValue('(실패) ' + providerResult.error);
    return;
  }

  var edits = providerResult.edits;
  if (!edits || edits.length === 0) {
    shQue.getRange(row, col).setValue('수정 구절 없음');
    return;
  }

  var html = _buildEditsHtml_(edits, refs);
  shQue.getRange(row, col).setValue(html);
}


/* ===========================
 * edits 배열 → HTML 변환
 * =========================== */

function _buildEditsHtml_(edits, refs) {
  var blocks = [];

  for (var i = 0; i < edits.length; i++) {
    var e = edits[i];

    var refIdx = (e.source_index !== null && e.source_index >= 1 && e.source_index <= refs.length)
      ? (e.source_index - 1) : -1;

    var source = refIdx >= 0 ? (refs[refIdx].source || '') : '';
    var link   = refIdx >= 0 ? (refs[refIdx].imageLink || '') : '';

    var text = '[[원본]] ' + e.original
      + '\n[[수정]] ' + e.revised
      + '\n[[근거]] ' + (e.evidence_quote || '(없음)')
      + '\n[[이유]] ' + e.reason;

    var sourceHtml = _escapeHtml_(source);
    var detailHtml = _escapeHtml_(text).replace(/\n/g, '<br>');
    var linkHtml   = link
      ? '<a href="' + _escapeHtml_(link) + '" target="_blank" rel="noopener noreferrer">원본link</a>'
      : '';

    var headParts = [];
    if (sourceHtml && sourceHtml.trim()) headParts.push(sourceHtml);
    if (linkHtml && linkHtml.trim()) headParts.push(linkHtml);

    blocks.push(headParts.join(' | ') + '<br>' + detailHtml);
  }

  return blocks.join('<br><br>');
}


/* ===========================
 * 상태/트리거 유틸
 * =========================== */

function _clearQueBatchState_() {
  PropertiesService.getScriptProperties().deleteProperty(QUE_BATCH_KEY);
}

function _scheduleResumeTrigger_() {
  _deleteTriggersByHandler_('batch_continueQueAuto');
  ScriptApp.newTrigger('batch_continueQueAuto').timeBased().after(RESUME_AFTER_MS).create();
}

function _deleteTriggersByHandler_(handlerName) {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction && triggers[i].getHandlerFunction() === handlerName) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}


/* ===========================
 * 행 스펙 파서
 * =========================== */

function _parseRowSpec_(spec) {
  var s = String(spec || '').replace(/\s+/g, '');
  if (!s) return [];
  var out = {};
  var parts = s.split(',');
  for (var p = 0; p < parts.length; p++) {
    var part = parts[p];
    if (!part) continue;
    if (/^\d+$/.test(part)) {
      out[Number(part)] = true;
    } else if (/^\d+\-\d+$/.test(part)) {
      var ab = part.split('-');
      var a = Number(ab[0]), b = Number(ab[1]);
      var start = Math.min(a, b), end = Math.max(a, b);
      for (var r = start; r <= end; r++) out[r] = true;
    }
  }
  var result = [];
  for (var key in out) { var n = Number(key); if (isFinite(n) && n >= 2) result.push(n); }
  result.sort(function(a, b) { return a - b; });
  return result;
}


/* ===========================
 * HTML 유틸
 * =========================== */

function _escapeHtml_(s) {
  return String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}
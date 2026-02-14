/*******************************************************
 * <배치 컨트롤: 멀티 LLM> (Claude / GPT / Gemini 선택)
 *
 * - batch_startQueAuto(): provider 선택 + 행 범위 입력 후 자동 처리
 * - batch_continueQueAuto(): 이어달리기
 * - batch_stopQueAuto(): 수동 중지
 *
 * 의존:
 *  - findSimilarFromB2({openViewer:false})
 *  - review_rewriteToKiceStyle({provider:'...', openViewer:false})
 *  - _buildQueE_HtmlFromReview_()
 *******************************************************/

var BATCH_SHEET_QUE    = 'Que';
var BATCH_SHEET_REVIEW = '문항검토';

var QUE_BATCH_KEY   = 'QUE_BATCH_STATE_V2';
var SAFE_RUN_MS     = 5 * 60 * 1000 + 20 * 1000;
var RESUME_AFTER_MS = 60 * 1000;
var PER_ROW_SLEEP_MS = 2000;


/** =========================
 * 1) 시작 — provider 선택 + 행 범위 입력
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

  // provider 선택
  var providerResp = ui.prompt(
    'LLM 선택',
    'claude / gpt / gemini 중 입력:',
    ui.ButtonSet.OK_CANCEL
  );
  if (providerResp.getSelectedButton() !== ui.Button.OK) return;

  var provider = String(providerResp.getResponseText() || '').trim().toLowerCase();
  if (provider !== 'claude' && provider !== 'gpt' && provider !== 'gemini') {
    ss.toast('유효하지 않은 provider: ' + provider, '오류', 5);
    return;
  }

  // 행 범위 입력
  var rowResp = ui.prompt(
    'Que 자동 배치 (' + provider.toUpperCase() + ')',
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
    provider: provider,
    startedAt: new Date().toISOString()
  };
  props.setProperty(QUE_BATCH_KEY, JSON.stringify(state));

  _deleteTriggersByHandler_('batch_continueQueAuto');
  ss.toast(provider.toUpperCase() + ' | 총 ' + rows.length + '행 자동 처리 시작', '배치', 5);

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
    var idx = Number(state.idx || 0);
    var provider = String(state.provider || 'claude').toLowerCase();

    var shQue = ss.getSheetByName(BATCH_SHEET_QUE);
    var shRev = ss.getSheetByName(BATCH_SHEET_REVIEW);
    if (!shQue || !shRev) {
      ss.toast('Que 또는 문항검토 시트를 찾지 못했습니다. 배치 중지.', '오류', 6);
      _clearQueBatchState_(); _deleteTriggersByHandler_('batch_continueQueAuto');
      return;
    }

    // ★ 안전장치: 강제 종료에 대비하여 이어달리기 트리거를 미리 예약
    // 정상 완료 시 삭제됨. 강제 종료되면 이 트리거가 살아서 자동 재개.
    _scheduleResumeTrigger_();

    while (idx < rows.length) {
      // ★ 안전 시간 체크 — API 호출 전에 미리 확인
      var elapsed = Date.now() - startMs;
      if (elapsed > SAFE_RUN_MS) {
        state.idx = idx;
        props.setProperty(QUE_BATCH_KEY, JSON.stringify(state));
        ss.toast(idx + '/' + rows.length + '까지 처리 (' + provider.toUpperCase() + '). 곧 이어서 실행됩니다.', '배치', 6);
        _scheduleResumeTrigger_();
        return;
      }

      var r = rows[idx];

      // ★ 행 처리 시작 전에 현재 진행 상태를 미리 저장 (강제 종료 방어)
      state.idx = idx;
      props.setProperty(QUE_BATCH_KEY, JSON.stringify(state));

      try {
        ss.toast('(' + (idx + 1) + '/' + rows.length + ') row ' + r + ' [' + provider.toUpperCase() + ']', '배치', 3);

        var id      = String(shQue.getRange(r, 1).getDisplayValue() || '').trim();
        var latex   = String(shQue.getRange(r, 2).getDisplayValue() || '').trim();
        var chapter = String(shQue.getRange(r, 3).getDisplayValue() || '').trim();

        if (!latex) {
          shQue.getRange(r, 4).setValue('');
          shQue.getRange(r, 5).setValue('');
          ss.toast('row ' + r + ': latex 비어있음 → 스킵', '배치', 3);
          idx++;
          continue;
        }

        ss.setActiveSheet(shRev);
        shRev.getRange('B2').setValue(latex);
        shRev.getRange('C2').setValue(chapter);

        // 코드1: 유사문항 검색
        findSimilarFromB2({ openViewer: false });

        // 코드2: LLM 윤문 (provider 전달)
        review_rewriteToKiceStyle({ provider: provider, openViewer: false });

        // Que!D = 윤문 결과
        var finalFull = String(shRev.getRange('B3').getDisplayValue() || '').trim();
        shQue.getRange(r, 4).setValue(finalFull);

        // Que!E = HTML
        var html = _buildQueE_HtmlFromReview_(shRev);
        shQue.getRange(r, 5).setValue(html || '수정 구절 없음');

        if (id) ss.toast('완료: ' + id + ' (row ' + r + ') [' + provider.toUpperCase() + ']', '배치', 2);
        if (PER_ROW_SLEEP_MS > 0) Utilities.sleep(PER_ROW_SLEEP_MS);

      } catch (errRow) {
        var msg = (errRow && errRow.message) ? errRow.message : String(errRow);
        // GAS 실행 시간 초과도 여기서 잡힘
        if (msg.indexOf('제한') !== -1 || msg.indexOf('time') !== -1 || msg.indexOf('limit') !== -1) {
          // 시간 초과로 추정 → 이어달리기 예약 후 종료
          state.idx = idx; // 현재 행 재시도
          props.setProperty(QUE_BATCH_KEY, JSON.stringify(state));
          _scheduleResumeTrigger_();
          ss.toast('시간 초과 감지 → row ' + r + '부터 이어서 실행 예정', '배치', 6);
          return;
        }
        ss.toast('row ' + r + ' 실패: ' + msg, '오류', 6);
        shQue.getRange(r, 4).setValue('(실패) ' + msg);
        shQue.getRange(r, 5).setValue('(실패) ' + msg);
      }

      idx++;
    }

    _clearQueBatchState_();
    _deleteTriggersByHandler_('batch_continueQueAuto');
    ss.toast('자동 배치 완료 ✅ (' + provider.toUpperCase() + ')', '배치', 6);

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
 * HTML 생성
 * =========================== */

function _escapeHtml_(s) {
  return String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

function _buildQueE_HtmlFromReview_(shRev) {
  var startRow = 21;
  var lastRow = shRev.getLastRow();
  if (lastRow < startRow) return '';
  var numRows = lastRow - startRow + 1;
  var vals = shRev.getRange(startRow, 1, numRows, 4).getValues();
  var rowBlocks = [];
  for (var i = 0; i < vals.length; i++) {
    var A = String(vals[i][0] || '').trim();
    var C = String(vals[i][2] || '').trim();
    var D = String(vals[i][3] || '').trim();
    if (!D || D === '수정 구절 없음') continue;
    var sourceHtml = _escapeHtml_(A);
    var detailHtml = _escapeHtml_(D).replace(/\n/g, '<br>');
    var linkHtml = C ? '<a href="' + _escapeHtml_(C) + '" target="_blank" rel="noopener noreferrer">원본link</a>' : '';
    var headParts = [];
    if (sourceHtml && sourceHtml.trim()) headParts.push(sourceHtml);
    if (linkHtml && linkHtml.trim()) headParts.push(linkHtml);
    rowBlocks.push(headParts.join(' | ') + '<br>' + detailHtml);
  }
  return rowBlocks.length === 0 ? '' : rowBlocks.join('<br><br>');
}
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
 *  - kiceAuto.gs (2026-10-02, 계획서 v4 §4.6): 소유 판별·러너 배치 훅
 *
 * [2026-10-02 · KICE 자동화 v4 §4.6] 변경 요지
 *  ① batch_continueQueAuto() 가 결과를 돌려준다: 'locked'|'idle'|'yield'|'stopped'|'done'|'fatal'
 *     (메뉴·트리거 호출자는 무시 → 영향 없음)
 *  ② 회 예산: 웹앱 tick(__KICE_TICK) 이면 KICE_TICK_BUDGET_MS, 메뉴·트리거면 SAFE_RUN_MS
 *  ③ 러너 배치(state.kice)에는 이어하기 트리거를 걸지 않는다 (D14)
 *  ④ 러너 배치: 루프 진입 시 남아 있는 inflight = 직전 실행이 그 행에서 강제 종료 → (실패) 기록·건너뜀 (D15)
 *  ⑤ 상태를 쓰는 모든 지점은 kice_saveQue_ (내 배치일 때만 저장) — 아니면 'stopped' 로 즉시 끝 (D17·A1)
 *     → ⏹️ / 원격 stop 뒤에 돌던 실행이 상태를 되살리지 못하고, 남의 배치 상태를 지우지도 않는다
 *  ⑥ 러너 배치: 행 처리 직전 inflight 기록 + heartbeat, 행이 끝나는 모든 경로에서 inflight 삭제
 *  ⑦ 완료: 소유 확인 → (러너 배치면) 완료 훅 kice_onBatchDone_ (아카이브 이관) → 소유일 때만 상태·트리거 삭제
 *  ⑧ 치명 분기(JSON 파싱 실패·시트 없음): 러너 배치면 AUTO error 로 알린 뒤 삭제
 *  ⑨ 러너 배치: 행 catch 의 '제한/time/limit' 분기를 타지 않고 (실패) 로 기록 → 다음 행.
 *     연속 3행 같은 예외 또는 3사 모두 (실패) 면 전역 장애로 보고 error (A3)
 *  ⑩ findSimilarFromB2 호출 직전 문항검토!A5:E20 초기화 + 시트를 인자로 넘김.
 *     러너 배치에서 유사문항 0개면 (실패) 유사문항 0개 기록 후 다음 행 (D22)
 *  - toast 는 웹앱·트리거 실행에서도 예외가 되지 않도록 kice_toast_ 로 감쌌다
 *  - [1-C P1] 'stopped' 로 빠질 때 남아 있는 상태가 메뉴 배치면 그 배치용 이어하기 트리거를 건다
 *    (⏹️ 직후 ▶️: 새 배치의 첫 실행이 옛 실행의 락에 막혀 'locked' 로 끝나므로, 옛 실행이 넘겨줘야 한다)
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
 * 1) 시작 — 행 범위 입력 (provider 선택 없음)   ※ 변경 없음
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
 *  반환: 'locked' | 'idle' | 'yield' | 'stopped' | 'done' | 'fatal'   (메뉴·트리거는 무시)
 * ========================= */
function batch_continueQueAuto() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return 'locked';                                   // ①

  var startMs = Date.now();
  var budget  = __KICE_TICK ? KICE_TICK_BUDGET_MS : SAFE_RUN_MS;             // ②
  var ss = SpreadsheetApp.getActive();
  var props = PropertiesService.getScriptProperties();

  try {
    var raw = props.getProperty(QUE_BATCH_KEY);
    if (!raw) return 'idle';

    var state;
    try { state = JSON.parse(raw); } catch (e) {
      kice_toast_(ss, '배치 상태 JSON 파싱 실패 → 중지합니다.', '오류', 6);
      var autoP = kice_loadAuto_();                                           // ⑧ state 가 없으므로 AUTO 로 판단
      if (autoP && autoP.stage === 'rewrite') kice_onBatchFail_('배치 상태 JSON 파싱 실패');
      _clearQueBatchState_(); _deleteTriggersByHandler_('batch_continueQueAuto');
      return 'fatal';
    }

    var rows = Array.isArray(state.rows) ? state.rows : [];
    var idx  = Number(state.idx || 0);
    var isKice = !!state.kice;                                                // 러너 배치인가

    var shQue = ss.getSheetByName(BATCH_SHEET_QUE);
    var shRev = ss.getSheetByName(BATCH_SHEET_REVIEW);
    if (!shQue || !shRev) {
      kice_toast_(ss, 'Que 또는 문항검토 시트를 찾지 못했습니다. 배치 중지.', '오류', 6);
      if (isKice) kice_onBatchFail_('Que 또는 문항검토 시트 없음');             // ⑧
      _clearQueBatchState_(); _deleteTriggersByHandler_('batch_continueQueAuto');
      return 'fatal';
    }

    // 안전장치: 강제 종료에 대비하여 이어달리기 트리거를 미리 예약 (러너 배치는 트리거 없음)
    _scheduleResumeTrigger_(state);                                           // ③

    // ④ 러너 배치: 직전 실행이 행 처리 중 죽었으면 그 행을 독 행으로 처리
    if (isKice && state.inflight) {
      var dead = kice_takeDeadInflight_(state, shQue, idx);
      idx = dead.idx;
      if (!kice_saveQue_(state)) return kice_stoppedExit_();
      kice_toast_(ss, 'row ' + dead.row + ' 강제 종료 흔적 → 건너뜀', '배치', 4);
    }

    while (idx < rows.length) {
      // 안전 시간 체크
      var elapsed = Date.now() - startMs;
      if (elapsed > budget) {
        state.idx = idx;
        if (!kice_saveQue_(state)) return kice_stoppedExit_();                          // ②·⑤ 소유일 때만 저장
        kice_toast_(ss, idx + '/' + rows.length + '까지 처리. 곧 이어서 실행됩니다.', '배치', 6);
        _scheduleResumeTrigger_(state);
        return 'yield';
      }

      var r = rows[idx];

      // 현재 진행 상태 저장 (강제 종료 방어) — 내 배치가 아니면 여기서 끝 (⏹️ / stop / 남의 배치)
      state.idx = idx;
      if (!kice_saveQue_(state)) return kice_stoppedExit_();                            // ⑤ (D17)

      var rowErr = null;          // 이 행의 예외 메시지 (A3 집계용)
      var rowVals = null;         // 이 행에 기록한 E/F/G (A3 집계용)

      try {
        kice_toast_(ss, '(' + (idx + 1) + '/' + rows.length + ') row ' + r + ' [3사 병렬]', '배치', 3);

        var id      = String(shQue.getRange(r, COL_QUE_ID).getDisplayValue() || '').trim();
        var latex   = String(shQue.getRange(r, COL_QUE_LATEX).getDisplayValue() || '').trim();
        var chapter = String(shQue.getRange(r, COL_QUE_CHAPTER).getDisplayValue() || '').trim();

        if (!latex) {
          shQue.getRange(r, COL_QUE_CLAUDE).setValue('');
          shQue.getRange(r, COL_QUE_GPT).setValue('');
          shQue.getRange(r, COL_QUE_GEMINI).setValue('');
          kice_toast_(ss, 'row ' + r + ': latex 비어있음 → 스킵', '배치', 3);
          idx++;
          if (isKice) { state.idx = idx; state.failStreak = { msg: '', n: 0, allFail: false }; if (!kice_saveQue_(state)) return kice_stoppedExit_(); }
          continue;
        }

        // ⑥ 러너 배치: 행 처리 직전 inflight 기록 (실행이 죽어도 남는다) + heartbeat
        if (isKice) {
          state.inflight = { row: r, at: Date.now() };
          if (!kice_saveQue_(state)) return kice_stoppedExit_();
          kice_heartbeat_();
        }

        // 문항검토 시트에 입력 세팅
        ss.setActiveSheet(shRev);
        shRev.getRange('B2').setValue(latex);
        shRev.getRange('C2').setValue(chapter);

        // ⑩ 코드1: 유사문항 검색 (1회) — 직전 행의 결과가 남지 않도록 출력 영역을 먼저 비운다
        shRev.getRange('A5:E20').clearContent();
        findSimilarFromB2({ sheet: shRev });

        var noRefs = false;
        if (isKice) {                                                           // D22
          var refVals = shRev.getRange('D6:D15').getValues();
          noRefs = !refVals.some(function (x) { return String(x[0] || '').trim() !== ''; });
        }

        if (noRefs) {
          var nr = '(실패) 유사문항 0개';
          shQue.getRange(r, COL_QUE_CLAUDE, 1, 3).setValues([[nr, nr, nr]]);
          rowVals = [nr, nr, nr];
          kice_toast_(ss, 'row ' + r + ': 유사문항 0개 → (실패) 기록', '배치', 3);
        } else {
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
          if (isKice) rowVals = shQue.getRange(r, COL_QUE_CLAUDE, 1, 3).getValues()[0];
        }

        if (id) kice_toast_(ss, '완료: ' + id + ' (row ' + r + ')', '배치', 2);
        if (PER_ROW_SLEEP_MS > 0) Utilities.sleep(PER_ROW_SLEEP_MS);

      } catch (errRow) {
        var msg = (errRow && errRow.message) ? errRow.message : String(errRow);

        // GAS 실행 시간 초과 감지 — 메뉴·트리거 배치만 (러너 배치는 ⑨: 보통 행 실패로 처리)
        if (!isKice && (msg.indexOf('제한') !== -1 || msg.indexOf('time') !== -1 || msg.indexOf('limit') !== -1)) {
          state.idx = idx;
          if (!kice_saveQue_(state)) return kice_stoppedExit_();                        // ⑤
          _scheduleResumeTrigger_(state);
          kice_toast_(ss, '시간 초과 감지 → row ' + r + '부터 이어서 실행 예정', '배치', 6);
          return 'yield';
        }
        kice_toast_(ss, 'row ' + r + ' 실패: ' + msg, '오류', 6);
        shQue.getRange(r, COL_QUE_CLAUDE).setValue('(실패) ' + msg);
        shQue.getRange(r, COL_QUE_GPT).setValue('(실패) ' + msg);
        shQue.getRange(r, COL_QUE_GEMINI).setValue('(실패) ' + msg);
        rowErr = msg;
      }

      idx++;

      // ⑥⑨ 러너 배치: 행 종료 처리 — inflight 삭제, 전역 장애 집계(A3), 소유 저장
      if (isKice) {
        delete state.inflight;
        state.idx = idx;

        var fs = state.failStreak || { msg: '', n: 0, allFail: false };
        if (rowErr) {
          fs = { msg: rowErr, n: (fs.n > 0 && !fs.allFail && fs.msg === rowErr) ? fs.n + 1 : 1, allFail: false };
        } else {
          var allFail = !!rowVals && rowVals.every(function (v) { return String(v || '').indexOf('(실패)') === 0; });
          if (allFail) {
            fs = { msg: fs.allFail && fs.n > 0 ? fs.msg : String(rowVals[0] || '').slice(0, 200), n: (fs.n > 0 && fs.allFail) ? fs.n + 1 : 1, allFail: true };
          } else {
            fs = { msg: '', n: 0, allFail: false };
          }
        }
        state.failStreak = fs;

        if (!kice_saveQue_(state)) return kice_stoppedExit_();

        if (fs.n >= KICE_FAIL_STREAK_N) {
          kice_onBatchFail_('연속 실패 ' + fs.n + '행: ' + fs.msg);
          kice_clearQueIfOwned_(state);
          kice_toast_(ss, '연속 실패 ' + fs.n + '행 → 배치 중단', '오류', 6);
          return 'fatal';
        }
      }
    }

    // ⑦ 완료 — 아직 내 배치인지 확인한 뒤에만 훅·삭제
    if (!kice_queOwned_(state)) return kice_stoppedExit_();
    if (isKice) kice_onBatchDone_(state);
    kice_clearQueIfOwned_(state);
    kice_toast_(ss, '자동 배치 완료 ✅ (3사 병렬)', '배치', 6);
    return 'done';

  } finally {
    lock.releaseLock();
  }
}


/** =========================
 * 3) 수동 중지   ※ 변경 없음 — 상태를 지우면 돌던 실행은 다음 저장 지점에서 'stopped' 로 끝난다
 * ========================= */
function batch_stopQueAuto() {
  var ss = SpreadsheetApp.getActive();
  _clearQueBatchState_();
  _deleteTriggersByHandler_('batch_continueQueAuto');
  ss.toast('자동 배치 중지됨', '배치', 5);
}


/* ===========================
 * provider별 결과를 Que 셀에 쓰기   ※ 변경 없음
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
 * edits 배열 → HTML 변환   ※ 변경 없음
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

/** ③ 이어하기 트리거 — 러너 배치(state.kice)에는 걸지 않는다(D14). 기존 트리거는 항상 정리. */
function _scheduleResumeTrigger_(state) {
  _deleteTriggersByHandler_('batch_continueQueAuto');
  if (state && state.kice) return;
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
 * 행 스펙 파서   ※ 변경 없음
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
 * HTML 유틸   ※ 변경 없음
 * =========================== */

function _escapeHtml_(s) {
  return String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

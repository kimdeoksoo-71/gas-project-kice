/*******************************************************
 * <코드3 : 배치컨트롤(완전 자동 트리거 이어달리기)>
 *
 * - batch_startQueAuto(): 사용자 입력(행 범위) 받고 상태 저장 후 즉시 1회 처리 시작
 * - batch_continueQueAuto(): 제한시간 전에 끊고 트리거로 이어달리기
 * - batch_stopQueAuto(): 수동 중지(상태/트리거 제거)
 *
 * 의존:
 *  - <코드1> findSimilarFromB2({openViewer:false})
 *  - <코드2> review_rewriteToKiceStyle_gpt52({openViewer:false})  // 네가 옵션화 수정한 상태
 *  - Que!E HTML 생성: _buildQueE_HtmlFromReview_()
 *******************************************************/

const BATCH_SHEET_QUE    = 'Que';
const BATCH_SHEET_REVIEW = '문항검토';

// 상태 저장 키
const QUE_BATCH_KEY = 'QUE_BATCH_STATE_V1';

// 안전 실행 시간(밀리초) : 5분 20초 정도에서 끊기
const SAFE_RUN_MS = 5 * 60 * 1000 + 20 * 1000;

// 다음 이어달리기까지 대기(밀리초)
const RESUME_AFTER_MS = 60 * 1000;

// 과도한 연속 호출 완화(필요 없으면 0으로)
const PER_ROW_SLEEP_MS = 150;


/** =========================
 * 1) 시작(사용자 입력 받고 자동 이어달리기 시작)
 * ========================= */
function batch_startQueAuto() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();

  const shQue = ss.getSheetByName(BATCH_SHEET_QUE);
  const shRev = ss.getSheetByName(BATCH_SHEET_REVIEW);
  if (!shQue || !shRev) {
    ss.toast('Que 또는 문항검토 시트를 찾지 못했습니다.', '오류', 5);
    return;
  }

  // 이미 실행 중이면 중복 방지
  const props = PropertiesService.getScriptProperties();
  const existing = props.getProperty(QUE_BATCH_KEY);
  if (existing) {
    ss.toast('이미 배치가 실행 중입니다. 중지 후 다시 시작하세요.', '배치', 5);
    return;
  }

  const resp = ui.prompt(
    'Que 자동 배치 시작',
    '처리할 행을 입력해줘 (예: 2,5,7-10)',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  const rows = _parseRowSpec_(String(resp.getResponseText() || '').trim());
  if (!rows.length) {
    ss.toast('유효한 행이 없습니다.', '중단', 5);
    return;
  }

  // 상태 저장
  const state = {
    rows,
    idx: 0,
    startedAt: new Date().toISOString(),
  };
  props.setProperty(QUE_BATCH_KEY, JSON.stringify(state));

  // 혹시 남아있는 트리거가 있으면 정리
  _deleteTriggersByHandler_('batch_continueQueAuto');

  ss.toast(`총 ${rows.length}행 자동 처리 시작`, '배치', 5);

  // 즉시 1회 처리(트리거 기다리지 말고)
  batch_continueQueAuto();
}


/** =========================
 * 2) 이어달리기(트리거가 호출하거나 start가 직접 호출)
 * ========================= */
function batch_continueQueAuto() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return; // 동시 실행 방지

  const startMs = Date.now();
  const ss = SpreadsheetApp.getActive();
  const props = PropertiesService.getScriptProperties();

  try {
    const raw = props.getProperty(QUE_BATCH_KEY);
    if (!raw) {
      // 상태가 없으면 할 일 없음
      return;
    }

    let state;
    try {
      state = JSON.parse(raw);
    } catch (e) {
      ss.toast('배치 상태 JSON 파싱 실패 → 중지합니다.', '오류', 6);
      _clearQueBatchState_();
      _deleteTriggersByHandler_('batch_continueQueAuto');
      return;
    }

    const rows = Array.isArray(state.rows) ? state.rows : [];
    let idx = Number(state.idx || 0);

    const shQue = ss.getSheetByName(BATCH_SHEET_QUE);
    const shRev = ss.getSheetByName(BATCH_SHEET_REVIEW);
    if (!shQue || !shRev) {
      ss.toast('Que 또는 문항검토 시트를 찾지 못했습니다. 배치 중지.', '오류', 6);
      _clearQueBatchState_();
      _deleteTriggersByHandler_('batch_continueQueAuto');
      return;
    }

    // 처리 루프
    while (idx < rows.length) {
      // 제한시간 전에 안전 종료
      if (Date.now() - startMs > SAFE_RUN_MS) {
        state.idx = idx;
        props.setProperty(QUE_BATCH_KEY, JSON.stringify(state));

        ss.toast(`시간 보호 종료. ${idx}/${rows.length}까지 처리. 곧 이어서 실행됩니다.`, '배치', 6);
        _scheduleResumeTrigger_();
        return;
      }

      const r = rows[idx];

      try {
        ss.toast(`(${idx + 1}/${rows.length}) row ${r} 처리중...`, '배치', 3);

        const id      = String(shQue.getRange(r, 1).getDisplayValue() || '').trim(); // A
        const latex   = String(shQue.getRange(r, 2).getDisplayValue() || '').trim(); // B
        const chapter = String(shQue.getRange(r, 3).getDisplayValue() || '').trim(); // C

        if (!latex) {
          shQue.getRange(r, 4).setValue(''); // D
          shQue.getRange(r, 5).setValue(''); // E
          ss.toast(`row ${r}: latex 비어있음 → 스킵`, '배치', 3);
          idx++;
          continue;
        }

        // 디버깅 편의(선택): 문항검토로
        ss.setActiveSheet(shRev);

        // 문항검토 입력
        shRev.getRange('B2').setValue(latex);
        shRev.getRange('C2').setValue(chapter);

        // <코드1> 검색(뷰어 OFF)
        findSimilarFromB2({ openViewer: false });

        // <코드2> 윤문(뷰어 OFF)
        review_rewriteToKiceStyle_gpt52({ openViewer: false });

        // Que!D = 문항검토!B3
        const finalFull = String(shRev.getRange('B3').getDisplayValue() || '').trim();
        shQue.getRange(r, 4).setValue(finalFull);

        // Que!E = HTML
        const html = _buildQueE_HtmlFromReview_(shRev);
        shQue.getRange(r, 5).setValue(html || '수정 구절 없음');

        if (id) ss.toast(`완료: ${id} (row ${r})`, '배치', 2);

        if (PER_ROW_SLEEP_MS > 0) Utilities.sleep(PER_ROW_SLEEP_MS);

      } catch (errRow) {
        const msg = (errRow && errRow.message) ? errRow.message : String(errRow);
        ss.toast(`row ${r} 실패: ${msg}`, '오류', 6);
        shQue.getRange(r, 4).setValue(`(실패) ${msg}`);
        shQue.getRange(r, 5).setValue(`(실패) ${msg}`);
        // 실패해도 다음 행
      }

      idx++;
    }

    // 다 끝남
    _clearQueBatchState_();
    _deleteTriggersByHandler_('batch_continueQueAuto');
    ss.toast('자동 배치 완료 ✅', '배치', 6);

  } finally {
    lock.releaseLock();
  }
}


/** =========================
 * 3) 수동 중지
 * ========================= */
function batch_stopQueAuto() {
  const ss = SpreadsheetApp.getActive();
  _clearQueBatchState_();
  _deleteTriggersByHandler_('batch_continueQueAuto');
  ss.toast('자동 배치 중지됨', '배치', 5);
}


/** =========================
 * 상태/트리거 유틸
 * ========================= */
function _clearQueBatchState_() {
  PropertiesService.getScriptProperties().deleteProperty(QUE_BATCH_KEY);
}

function _scheduleResumeTrigger_() {
  // 중복 트리거 방지
  _deleteTriggersByHandler_('batch_continueQueAuto');

  ScriptApp.newTrigger('batch_continueQueAuto')
    .timeBased()
    .after(RESUME_AFTER_MS)
    .create();
}

function _deleteTriggersByHandler_(handlerName) {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction && t.getHandlerFunction() === handlerName) {
      ScriptApp.deleteTrigger(t);
    }
  });
}


/** =========================
 * 행 스펙 파서: "2,5,7-10" → [2,5,7,8,9,10]
 * ========================= */
function _parseRowSpec_(spec) {
  const out = new Set();
  const s = String(spec || '').replace(/\s+/g, '');
  if (!s) return [];

  const parts = s.split(',').filter(Boolean);
  for (const p of parts) {
    if (/^\d+$/.test(p)) {
      out.add(Number(p));
    } else if (/^\d+\-\d+$/.test(p)) {
      const [a, b] = p.split('-').map(Number);
      const start = Math.min(a, b);
      const end   = Math.max(a, b);
      for (let r = start; r <= end; r++) out.add(r);
    }
  }

  return Array.from(out)
    .filter(n => Number.isFinite(n) && n >= 2)
    .sort((a, b) => a - b);
}


/** =========================
 * HTML 생성(네가 수정한 "출처 | 원본link" 1줄 헤더 포함)
 * ========================= */
function _escapeHtml_(s) {
  return String(s ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function _buildQueE_HtmlFromReview_(shRev) {
  const startRow = 21;
  const lastRow = shRev.getLastRow();
  if (lastRow < startRow) return '';

  const numRows = lastRow - startRow + 1;

  // ✅ 표시값 말고 실제 값으로 읽기
  const vals = shRev.getRange(startRow, 1, numRows, 4).getValues(); // A~D

  const rowBlocks = [];

  for (let i = 0; i < vals.length; i++) {
    const A = String(vals[i][0] || '').trim(); // source
    const C = String(vals[i][2] || '').trim(); // url
    const D = String(vals[i][3] || '').trim(); // detail (수정 내용)

    // ✅ "수정 구절 없음" 또는 빈 행은 완전히 무시
    if (!D || D === '수정 구절 없음') continue;

    const sourceHtml = _escapeHtml_(A);
    const detailHtml = _escapeHtml_(D).replace(/\n/g, '<br>');

    // ✅ 링크가 있을 때만 원본link 생성
    const linkHtml = C
      ? `<a href="${_escapeHtml_(C)}" target="_blank" rel="noopener noreferrer">원본link</a>`
      : '';

    // ✅ 출처 | 원본link 를 한 줄로 (없는 항목은 자동 제거)
    const head = [sourceHtml, linkHtml]
      .filter(s => s && s.trim())
      .join(' | ');

    rowBlocks.push(`${head}<br>${detailHtml}`);
  }

  // ✅ 실제 수정 내용이 하나도 없으면 HTML 자체를 만들지 않음
  if (rowBlocks.length === 0) {
    return '';
  }

  // 행 간 구분은 빈 줄
  return rowBlocks.join('<br><br>');
}



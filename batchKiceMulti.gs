/*******************************************************
 * batchKiceMulti.gs — 메뉴에서 시작하는 Que 윤문 (kice-2.0.0 · 재구축 계획서 v5 §17 S6, §18 N3·N14)
 *
 * 2.0.0 에서 바뀐 것
 *  - 이 프로젝트는 LLM 을 부르지 않는다. 메뉴는 "이 행들을 윤문해 달라"는 요청만 상태에 적고,
 *    맥미니의 kice_engine 이 claim → submit 으로 처리한다. 트리거·이어달리기 루프가 없다.
 *  - 단건 윤문도 이 길이다: Que 에 행을 붙여 넣고(A id · B Latex) ▶ Que 윤문 요청.
 *  - 메뉴 배치의 결과는 아카이브로 옮기지 않고 Que 에 남는다. 다음 러너 start 의 잔재 이관이 옮긴다.
 *
 * 상태는 kiceAuto.gs 의 KICE_AUTO_STATE_V2 / QUE_BATCH_STATE_V4 를 같이 쓴다.
 * 메뉴 배치의 runId 는 `menu~<시각>` 이고 q.menu = true 다.
 *******************************************************/

var KICE_MENU_AFTER_RUNNER_MS = 2 * 60 * 1000;      // 러너 배치가 끝난 직후에는 러너가 결과를 읽어 갈 틈을 준다

/** ▶ Que 윤문 요청 — 행 범위 입력 */
function batch_startQueAuto() {
  var ss = SpreadsheetApp.getActive();
  var ui = SpreadsheetApp.getUi();
  var shQue = ss.getSheetByName(KICE_QUE_SHEET);
  if (!shQue) { kice_toast_(ss, 'Que 시트를 찾지 못했습니다.', '오류', 5); return; }

  var rowResp = ui.prompt('Que 윤문 요청 (엔진이 처리: Claude + Gemini)',
    '처리할 행을 입력해 주세요 (예: 2,5,7-10)\nA열 id · B열 Latex 가 있는 행만 대상입니다.',
    ui.ButtonSet.OK_CANCEL);
  if (rowResp.getSelectedButton() !== ui.Button.OK) return;
  var rows = _parseRowSpec_(String(rowResp.getResponseText() || '').trim());
  if (!rows.length) { kice_toast_(ss, '유효한 행이 없습니다.', '중단', 5); return; }
  if (rows.length > KICE_ROWS_MAX) { ui.alert('한 번에 ' + KICE_ROWS_MAX + '행까지입니다.'); return; }

  // 대상 행 점검 — id·Latex 가 없는 행은 빼고, 이미 결과가 있는 행은 지울지 묻는다
  var read = kice_readQueRows_(shQue, rows);
  var valid = [], hasResult = 0;
  read.forEach(function (x) {
    if (!kice_hasValue_(x.v[COL_QUE_ID - 1]) || !kice_hasValue_(x.v[COL_QUE_LATEX - 1])) return;
    valid.push(x.row);
    if (kice_hasValue_(x.v[4]) || kice_hasValue_(x.v[5]) || kice_hasValue_(x.v[6])) hasResult++;
  });
  if (!valid.length) { ui.alert('id(A)와 Latex(B)가 모두 있는 행이 없습니다.'); return; }
  if (hasResult) {
    var ok = ui.alert('이미 결과가 있는 행 ' + hasResult + '개',
      '그 행들의 E·F·G 를 지우고 다시 윤문합니다. 지난 결과는 아카이브에 없으면 사라집니다.\n계속할까요?', ui.ButtonSet.YES_NO);
    if (ok !== ui.Button.YES) return;
  }

  var res = kice_withLock_(function () {
    var auto = kice_loadAuto_(), q = kice_readQueRaw_();
    if (auto && !kice_isEnded_(auto) && (kice_now_() - Number(auto.hb || 0)) <= KICE_TAKEOVER_MS) {
      return { ok: false, reason: '이미 실행 중입니다 — ' + auto.runId + ' (단계 ' + auto.stage + '). 끝난 뒤 다시 요청하거나 ⏹ 로 중지해 주세요.' };
    }
    if (auto && !auto.menu && kice_isEnded_(auto) && (kice_now_() - Number(auto.hb || 0)) < KICE_MENU_AFTER_RUNNER_MS) {
      return { ok: false, reason: '러너 배치가 방금 끝났습니다. 2분 뒤에 다시 요청해 주세요.' };
    }
    if (q) kice_clearQue_();
    if (auto && !kice_isEnded_(auto)) kice_markStopped_(auto, 'takeover by menu');

    valid.forEach(function (r) { shQue.getRange(r, COL_QUE_CLAUDE, 1, 3).clearContent(); });
    SpreadsheetApp.flush();

    var startedAt = kice_iso_();
    var runId = 'menu~' + startedAt.replace(/[-:.TZ]/g, '').slice(0, 14);
    kice_saveQue_({ startedAt: startedAt, kice: runId, menu: true, rows: valid, inflight: {}, poison: [], poisonCount: 0 });
    kice_saveAuto_({
      runId: runId, stage: 'rewrite', keywords: [], expect: null, startedAt: startedAt, hb: kice_now_(),
      total: valid.length, idx: 0, menu: true,
      src: null, residue: null, summary: null, archive: null, error: '', stoppedFrom: '', by: ''
    });
    kice_log_(runId, 'rewrite', '메뉴 요청 ' + valid.length + '행 (' + valid.slice(0, 12).join(',') + (valid.length > 12 ? '…' : '') + ') — 엔진 대기');
    return { ok: true, runId: runId, rows: valid.length };
  });

  if (!res.ok) { ui.alert('요청하지 못했습니다', res.reason === 'locked' ? '다른 작업이 진행 중입니다. 잠시 뒤 다시 눌러 주세요.' : res.reason, ui.ButtonSet.OK); return; }
  kice_toast_(ss, res.rows + '행을 엔진에 요청했습니다. 진행은 📋 상태에서 볼 수 있고, 결과는 Que 에 채워집니다.', 'KICE 윤문', 8);
}

/** ⏹ 중지 — 진행 중인 배치(메뉴·러너)를 멈춘다. 이미 채워진 행은 그대로 둔다. */
function batch_stopQueAuto() {
  var ss = SpreadsheetApp.getActive();
  var r = kice_apiStop_('menu');
  if (!r.ok) { kice_toast_(ss, '다른 작업이 진행 중입니다. 잠시 뒤 다시 눌러 주세요.', '중지', 5); return; }
  kice_toast_(ss, r.stoppedFrom ? '윤문을 중지했습니다' + (r.done !== null ? ' (끝난 행 ' + r.done + '개는 Que 에 있습니다)' : '') : '진행 중인 윤문이 없습니다.', 'KICE 윤문', 6);
}


/* ===========================
 * 행 스펙 파서 (1.0.0 그대로)
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
      if (end - start > 5000) continue;
      for (var r = start; r <= end; r++) out[r] = true;
    }
  }
  var result = [];
  for (var key in out) { var n = Number(key); if (isFinite(n) && n >= 2) result.push(n); }
  result.sort(function (a, b) { return a - b; });
  return result;
}

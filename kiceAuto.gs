/*******************************************************
 * kiceAuto.gs — KICE 윤문 자동화 (러너 연동) · 계획서 v4 §4.2~§4.5·§4.7
 *
 * 역할
 *  - 러너(맥미니 audit_runner)가 웹앱(RemoteApi.gs)으로 부르는 start / tick / status / stop 의 본체
 *  - batchKiceMulti.gs 가 쓰는 공용 함수: 소유 판별(kice_queOwned_), 소유 저장(kice_saveQue_),
 *    heartbeat, 독 행(inflight) 처리, 완료 훅(kice_onBatchDone_ → 아카이브 이관), 실패 훅
 *
 * 상태 (ScriptProperties)
 *  - QUE_BATCH_STATE_V3 : 기존 {rows, idx, startedAt} + 러너 배치일 때만 kice/inflight/poison/poisonCount/failStreak
 *  - KICE_AUTO_STATE    : {runId, startedAt, stage, keywords, expect, total, src, residue, summary, archive,
 *                          error, stoppedFrom, by, hb(ms), updatedAt(ISO)}
 *                         stage ∈ loading | rewrite | archive | done | error | stopped
 *  - LATEX_SS_ID        : Latex 변환 스프레드시트 ID (1회 설정)
 *
 * 원칙
 *  - 상태를 쓰거나 지우는 모든 지점은 "아직 내 배치인가"(startedAt·kice 동일)를 확인한 뒤에만 (D17·A1)
 *  - status 는 읽기만 한다 (D21). 저장은 stop 과 배치 엔진만.
 *  - 러너 배치에는 트리거를 걸지 않는다 (D14)
 *  - 아카이브에는 끝에만 붙인다 (뷰어 경계 864행, D23)
 *******************************************************/

var KICE_AUTO_KEY       = 'KICE_AUTO_STATE';
var KICE_LATEX_SS_PROP  = 'LATEX_SS_ID';
var KICE_LATEX_SRC_SHEET = 'Data_DS';
var KICE_LOG_SHEET      = 'KICE_Log';
var KICE_ARCHIVE_SHEET  = '아카이브';
var KICE_LIST_CAP       = 30;                       // 상태 속성 9KB 보호 — 목록은 30개까지
var KICE_TAKEOVER_MS    = 15 * 60 * 1000;           // heartbeat 가 이만큼 멈춘 러너 배치는 버려진 것으로 보고 인수
var KICE_END_STAGES     = ['done', 'error', 'stopped'];
var KICE_FAIL_STREAK_N  = 3;                        // A3: 연속 3행 실패 → 전역 장애

var __KICE_TICK = false;                            // 이 실행이 웹앱 tick 인가 (실행 단위 전역)
var KICE_TICK_BUDGET_MS = 3.5 * 60 * 1000;          // D20 — R6 실측 뒤 조정


/* =========================================================
 * 0. 작은 유틸
 * ========================================================= */

function kice_now_() { return Date.now(); }
function kice_iso_() { return new Date().toISOString(); }

function kice_merge_(a, b) {
  var o = {};
  for (var k in a) o[k] = a[k];
  for (var k2 in b) o[k2] = b[k2];
  return o;
}

function kice_nfc_(s) {
  s = String(s == null ? '' : s);
  return s.normalize ? s.normalize('NFC') : s;
}

/** 셀 비교용 문자열 (trim 없음, 개행만 통일) */
function kice_cell_(v) {
  return String(v == null ? '' : v).replace(/\r\n/g, '\n');
}
function kice_sameCell_(a, b) { return kice_cell_(a) === kice_cell_(b); }
function kice_hasValue_(v) { return String(v == null ? '' : v).trim() !== ''; }

/** 1~n열 중 아무 값이 있는 마지막 행 (없으면 1) — latex-convert 의 dlds_lastDataRow_ 복제 */
function kice_lastDataRow_(sh, n) {
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return 1;
  var v = sh.getRange(2, 1, lastRow - 1, n).getValues();
  for (var i = v.length - 1; i >= 0; i--) {
    for (var c = 0; c < v[i].length; c++) {
      if (String(v[i][c]).trim() !== '') return i + 2;
    }
  }
  return 1;
}

/** audition pv_keyMatcher_ 복제: NFC·소문자·부분 문자열 (D18) */
function kice_keyMatcher_(keywords) {
  var kws = keywords.map(function (k) { return kice_nfc_(k).toLowerCase(); });
  return function (key) {
    var kl = kice_nfc_(key).toLowerCase();
    for (var i = 0; i < kws.length; i++) if (kl.indexOf(kws[i]) !== -1) return true;
    return false;
  };
}

function kice_parseKeywords_(raw) {
  var seen = {}, out = [];
  String(raw || '').split(/[,\n;]+/).forEach(function (s) {
    s = s.trim();
    if (s && !seen[s]) { seen[s] = true; out.push(s); }
  });
  return out;
}

/** 목록을 30개까지만 (9KB 보호) */
function kice_capList_(arr) {
  arr = Array.isArray(arr) ? arr : [];
  return arr.length > KICE_LIST_CAP ? arr.slice(0, KICE_LIST_CAP) : arr;
}

/** KICE_Log 시트에 1행 — 실패해도 본 흐름을 막지 않는다 */
function kice_log_(runId, stage, msg) {
  try {
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName(KICE_LOG_SHEET);
    if (!sh) { sh = ss.insertSheet(KICE_LOG_SHEET); sh.appendRow(['time', 'run', 'stage', 'message']); }
    sh.appendRow([new Date(), String(runId || ''), String(stage || ''), String(msg || '')]);
  } catch (_) {}
}

/** 웹앱·트리거 실행에서도 안전한 toast */
function kice_toast_(ss, msg, title, sec) {
  try { ss.toast(msg, title || '배치', sec || 5); } catch (_) {}
}


/* =========================================================
 * 1. AUTO 상태
 * ========================================================= */

function kice_loadAuto_() {
  var s = PropertiesService.getScriptProperties().getProperty(KICE_AUTO_KEY);
  if (!s) return null;
  try { return JSON.parse(s); } catch (_) { return null; }
}

function kice_saveAuto_(auto) {
  auto.updatedAt = kice_iso_();
  PropertiesService.getScriptProperties().setProperty(KICE_AUTO_KEY, JSON.stringify(auto));
  return auto;
}

/** AUTO.hb(ms)·updatedAt 갱신 — 행 시작마다 */
function kice_heartbeat_() {
  var auto = kice_loadAuto_();
  if (!auto || auto.stage !== 'rewrite') return;     // 1-C P3: stop 과 겹쳐 'stopped' 를 되돌리지 않게
  auto.hb = kice_now_();
  kice_saveAuto_(auto);
}

function kice_isEnded_(auto) {
  return !!auto && KICE_END_STAGES.indexOf(auto.stage) >= 0;
}

/** AUTO 를 'stopped' 로 (진행 상태일 때만). 반환: stoppedFrom 또는 '' */
function kice_markStopped_(auto, by) {
  if (!auto || kice_isEnded_(auto)) return '';
  var from = auto.stage;
  auto.stoppedFrom = from;
  auto.stage = 'stopped';
  auto.by = by || '';
  auto.hb = kice_now_();
  kice_saveAuto_(auto);
  kice_log_(auto.runId, 'stopped', (by || '중지') + ' (from ' + from + ')');
  return from;
}

/** 실패 훅 — AUTO stage:'error' */
function kice_onBatchFail_(reason) {
  var auto = kice_loadAuto_();
  if (!auto || kice_isEnded_(auto)) return;
  auto.stage = 'error';
  auto.error = String(reason || '');
  auto.hb = kice_now_();
  kice_saveAuto_(auto);
  kice_log_(auto.runId, 'error', auto.error);
}


/* =========================================================
 * 2. QUE 상태 — 소유 판별 (D17·A1)
 * ========================================================= */

/** QUE_BATCH_STATE_V3 를 매번 새로 읽어 파싱 (없거나 깨지면 null) */
function kice_readQueRaw_() {
  var raw = PropertiesService.getScriptProperties().getProperty(QUE_BATCH_KEY);
  if (!raw) return null;
  try { return JSON.parse(raw); } catch (_) { return null; }
}

/** 지금 저장된 상태가 아직 "내 배치"인가 — startedAt 과 kice 가 모두 같아야 한다 */
function kice_queOwned_(state) {
  if (!state) return false;
  var live = kice_readQueRaw_();
  return !!live
    && String(live.startedAt || '') === String(state.startedAt || '')
    && String(live.kice || '') === String(state.kice || '');
}

/** 소유일 때만 저장. 반환 true=저장함 / false=쓰지 않음(⏹️·stop·남의 배치) */
function kice_saveQue_(state) {
  if (!kice_queOwned_(state)) return false;
  PropertiesService.getScriptProperties().setProperty(QUE_BATCH_KEY, JSON.stringify(state));
  return true;
}

/** 소유일 때만 상태·트리거 삭제 */
function kice_clearQueIfOwned_(state) {
  if (!kice_queOwned_(state)) return false;
  _clearQueBatchState_();
  _deleteTriggersByHandler_('batch_continueQueAuto');
  return true;
}

/**
 * 'stopped' 로 빠질 때 부른다 (1-C P1). 지금 저장된 상태가 남의 **메뉴 배치**면 그 배치용 이어하기 트리거를 건다.
 * ⏹️ 직후 ▶️ 로 새 배치를 시작하면 새 배치의 첫 실행은 옛 실행이 쥔 락에 막혀 'locked' 로 끝나고 트리거도 없다.
 * 락을 쥔 옛 실행만이 "곧 락이 풀린다"는 것을 알므로 여기서 넘겨준다. 러너 배치는 tick 이 돌리므로 걸지 않는다.
 */
function kice_stoppedExit_() {
  var live = kice_readQueRaw_();
  if (live && !live.kice) _scheduleResumeTrigger_(live);
  return 'stopped';
}

/**
 * 독 행 처리 (D15·§4.6 ④) — 루프 진입 직후, 첫 행 전에 호출. 러너 배치(state.kice)만.
 * 직전 실행이 행 처리 도중 강제 종료됐으면 state.inflight 가 남아 있다 → 그 행을 (실패) 로 기록하고 건너뛴다.
 * 반환: { row, idx }  (row=0 이면 inflight 없음). 저장은 호출자가 kice_saveQue_ 로.
 */
function kice_takeDeadInflight_(state, shQue, idx) {
  var f = state.inflight;
  if (!f || !f.row) return { row: 0, idx: idx };
  var dead = Number(f.row);
  var note = '(실패) GAS 6분 제한으로 강제 종료 — 건너뜀';
  try { shQue.getRange(dead, COL_QUE_CLAUDE, 1, 3).setValues([[note, note, note]]); } catch (_) {}
  state.poison = kice_capList_(state.poison || []);
  if (state.poison.length < KICE_LIST_CAP && state.poison.indexOf(dead) < 0) state.poison.push(dead);
  state.poisonCount = Number(state.poisonCount || 0) + 1;
  var rows = Array.isArray(state.rows) ? state.rows : [];
  var di = rows.indexOf(dead);
  if (di >= 0 && di >= idx) idx = di + 1;          // 죽은 행 다음부터
  delete state.inflight;
  state.idx = idx;
  kice_log_(state.kice, 'poison', 'row ' + dead + ' — 직전 실행이 이 행에서 강제 종료됨(6분 제한 추정), 건너뜀');
  return { row: dead, idx: idx };
}


/* =========================================================
 * 3. 아카이브 — 색인·중복 제외·이어 붙이기·요약
 * ========================================================= */

function kice_archiveSheet_() {
  var sh = SpreadsheetApp.getActive().getSheetByName(KICE_ARCHIVE_SHEET);
  if (!sh) throw new Error(KICE_ARCHIVE_SHEET + ' 시트를 찾을 수 없습니다.');
  return sh;
}

/** 아카이브 A열만 전체 읽어 {id: [row…]} (ids 에 있는 것만) */
function kice_archiveIndex_(sh, ids) {
  var want = {};
  (ids || []).forEach(function (id) { id = kice_cell_(id); if (id) want[id] = true; });
  var out = {};
  var last = sh.getLastRow();
  if (last < 2) return out;
  var vals = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < vals.length; i++) {
    var id = kice_cell_(vals[i][0]);
    if (!id || !want[id]) continue;
    (out[id] = out[id] || []).push(i + 2);
  }
  return out;
}

/**
 * Que 행(A:G 7열 배열)들 중 아카이브에 "같은 id + 같은 E·F·G" 가 이미 있는 행을 제외 (D8′).
 * 후보 id 의 행만 E:G 를 읽는다 (K1-08).
 * 반환 { keep: rows7[], dup: n }
 */
function kice_dedupeAgainstArchive_(sh, rows7) {
  if (!rows7.length) return { keep: [], dup: 0 };
  var index = kice_archiveIndex_(sh, rows7.map(function (r) { return r[0]; }));
  var cache = {};                                   // row -> [E,F,G]
  var keep = [], dup = 0;
  rows7.forEach(function (r) {
    var id = kice_cell_(r[0]);
    var hits = index[id] || [];
    var isDup = false;
    for (var i = 0; i < hits.length && !isDup; i++) {
      var row = hits[i];
      if (!cache[row]) cache[row] = sh.getRange(row, COL_QUE_CLAUDE, 1, 3).getValues()[0];
      var a = cache[row];
      isDup = kice_sameCell_(a[0], r[4]) && kice_sameCell_(a[1], r[5]) && kice_sameCell_(a[2], r[6]);
    }
    if (isDup) dup++; else keep.push(r);
  });
  return { keep: keep, dup: dup };
}

/** 아카이브 끝에 7열 행들을 이어 붙인다 → {start, end} (없으면 {start:0, end:0}) */
function kice_appendArchive_(sh, rows7) {
  if (!rows7.length) return { start: 0, end: 0 };
  var last = kice_lastDataRow_(sh, 7);
  var start = last >= 2 ? last + 1 : 2;
  if (sh.getMaxRows() < start + rows7.length - 1) sh.insertRowsAfter(sh.getMaxRows(), start + rows7.length - 1 - sh.getMaxRows());
  if (sh.getMaxColumns() < 7) sh.insertColumnsAfter(sh.getMaxColumns(), 7 - sh.getMaxColumns());
  sh.getRange(start, 1, rows7.length, 7).setValues(rows7);
  return { start: start, end: start + rows7.length - 1 };
}

/** 결과 칸 분류: edit(HTML 등) / none(수정 구절 없음) / fail((실패)로 시작) / empty */
function kice_classifyCell_(v) {
  var s = kice_cell_(v).trim();
  if (!s) return 'empty';
  if (s === '수정 구절 없음' || s === '수정구절 없음') return 'none';
  if (s.indexOf('(실패)') === 0) return 'fail';
  return 'edit';
}

/** provider 별 집계 */
function kice_summarize_(rows7) {
  var mk = function () { return { edit: 0, none: 0, fail: 0, empty: 0 }; };
  var sum = { rows: rows7.length, claude: mk(), gpt: mk(), gemini: mk() };
  rows7.forEach(function (r) {
    sum.claude[kice_classifyCell_(r[4])]++;
    sum.gpt[kice_classifyCell_(r[5])]++;
    sum.gemini[kice_classifyCell_(r[6])]++;
  });
  return sum;
}


/* =========================================================
 * 4. 완료 훅 (§4.5) — batch_continueQueAuto 가 소유를 확인한 뒤에만 부른다
 * ========================================================= */

function kice_onBatchDone_(state) {
  var auto = kice_loadAuto_();
  if (!auto || !state || String(state.kice || '') !== String(auto.runId || '') || auto.stage !== 'rewrite') return;

  try {
    auto.stage = 'archive';
    auto.hb = kice_now_();
    kice_saveAuto_(auto);

    var ss = SpreadsheetApp.getActive();
    var shQue = ss.getSheetByName(BATCH_SHEET_QUE);
    var shArc = kice_archiveSheet_();

    var rows = Array.isArray(state.rows) ? state.rows : [];
    var rows7 = [];
    if (rows.length) {
      var minR = Math.min.apply(null, rows), maxR = Math.max.apply(null, rows);
      var vals = shQue.getRange(minR, 1, maxR - minR + 1, 7).getValues();
      rows7 = vals.filter(function (r) { return kice_hasValue_(r[0]) || kice_hasValue_(r[1]); });
    }

    var dd = kice_dedupeAgainstArchive_(shArc, rows7);
    var range = kice_appendArchive_(shArc, dd.keep);

    var summary = kice_summarize_(rows7);
    summary.archived = dd.keep.length;
    summary.dup = dd.dup;
    summary.poisonCount = Number(state.poisonCount || 0);
    summary.countMismatch = !!(auto.src && auto.src.countMismatch);
    summary.alreadyDone = Number(auto.src && auto.src.alreadyDone || 0);

    auto.stage = 'done';
    auto.summary = summary;
    auto.archive = range;
    auto.hb = kice_now_();
    kice_saveAuto_(auto);
    kice_log_(auto.runId, 'done',
      '윤문 ' + summary.rows + '행 → 아카이브 ' + (range.start ? range.start + '~' + range.end : '추가 없음') +
      (dd.dup ? ' (중복 제외 ' + dd.dup + ')' : '') +
      ' · Claude edit ' + summary.claude.edit + ' / GPT ' + summary.gpt.edit + ' / Gemini ' + summary.gemini.edit +
      ' · 실패 C' + summary.claude.fail + ' G' + summary.gpt.fail + ' M' + summary.gemini.fail +
      (summary.poisonCount ? ' · poison ' + summary.poisonCount : ''));
  } catch (e) {
    kice_onBatchFail_('archive: ' + ((e && e.message) || e));
  }
}


/* =========================================================
 * 5. start (§4.3)
 * ========================================================= */

function kice_latexSsId_() {
  return (PropertiesService.getScriptProperties().getProperty(KICE_LATEX_SS_PROP) || '').trim();
}

function kice_apiStart_(p) {
  p = p || {};
  var keywords = kice_parseKeywords_(p.keywords);
  var runId = String(p.runId || '').trim();
  if (!keywords.length) return { ok: false, reason: 'keywords 파라미터가 비어 있습니다.' };
  if (!runId) return { ok: false, reason: 'runId 파라미터가 비어 있습니다.' };
  var expect = (p.expect === undefined || p.expect === null || String(p.expect) === '') ? null : Number(p.expect);
  if (expect !== null && !isFinite(expect)) expect = null;
  var limit = parseInt(p.limit, 10);
  if (!isFinite(limit) || limit < 1) limit = 0;

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { ok: false, reason: '잠금 대기 초과 — 이미 실행 중' };

  try {
    var ss = SpreadsheetApp.getActive();
    var auto = kice_loadAuto_();

    // ① 멱등 — 같은 runId 가 이미 시작됐고 loading 이 아니면 그대로 인정
    if (auto && auto.runId === runId && auto.stage !== 'loading') {
      return { ok: true, adopted: true, runId: runId, startedAt: auto.startedAt, stage: auto.stage, rows: auto.total || 0 };
    }

    // ② 가드 / 인수
    var q = kice_readQueRaw_();
    if (q) {
      var qRows = Array.isArray(q.rows) ? q.rows.length : 0;
      if (!q.kice) {
        // 1-C P4: 락을 잡았으므로 지금 도는 실행은 없다. 트리거까지 없으면 멈춘 배치 → 덕수님이 정리해야 한다
        var hasTrig = false;
        try { hasTrig = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'batch_continueQueAuto'; }); } catch (_) {}
        return { ok: false, stalled: !hasTrig,
                 reason: '이미 실행 중 — 메뉴 Que 자동윤문(행 ' + (Number(q.idx) || 0) + '/' + qRows + ')' +
                         (hasTrig ? '' : ' · 실행·트리거가 없어 멈춘 배치로 보임 → 메뉴 ⏹️ Que 자동윤문 중지로 정리해 주세요') };
      }
      var stale = !auto || auto.runId !== q.kice || q.kice === runId ||
                  (kice_now_() - Number(auto.hb || 0)) > KICE_TAKEOVER_MS;
      if (!stale) {
        return { ok: false, reason: '이미 실행 중 — 러너 배치 ' + q.kice + ' (행 ' + (Number(q.idx) || 0) + '/' + qRows + ')' };
      }
      _clearQueBatchState_();
      _deleteTriggersByHandler_('batch_continueQueAuto');
      if (auto && auto.runId !== runId) kice_markStopped_(auto, 'takeover by ' + runId);
      kice_log_(runId, 'takeover', '버려진 러너 배치 인수: ' + q.kice);
    }

    // ③ AUTO 선기록 (loading)
    auto = {
      runId: runId, stage: 'loading', keywords: keywords, expect: expect,
      startedAt: kice_iso_(), hb: kice_now_(), total: 0,
      src: null, residue: null, summary: null, archive: null, error: '', stoppedFrom: '', by: ''
    };
    kice_saveAuto_(auto);
    kice_log_(runId, 'start', '키워드: ' + keywords.join(' | ') + (expect !== null ? ' · expect ' + expect : '') + (limit ? ' · limit ' + limit : ''));

    var shQue = ss.getSheetByName(BATCH_SHEET_QUE);
    var shRev = ss.getSheetByName(BATCH_SHEET_REVIEW);
    if (!shQue || !shRev) {
      kice_onBatchFail_('Que 또는 문항검토 시트를 찾지 못했습니다.');
      return { ok: false, reason: 'Que 또는 문항검토 시트를 찾지 못했습니다.' };
    }
    var shArc = kice_archiveSheet_();

    // ④ 대상 선택 (D18·D19)
    var ssId = kice_latexSsId_();
    if (!ssId) { kice_onBatchFail_('스크립트 속성 LATEX_SS_ID 가 비어 있습니다.'); return { ok: false, reason: 'LATEX_SS_ID 미설정' }; }
    var src = SpreadsheetApp.openById(ssId).getSheetByName(KICE_LATEX_SRC_SHEET);
    if (!src) { kice_onBatchFail_('Latex 변환 파일에 ' + KICE_LATEX_SRC_SHEET + ' 시트가 없습니다.'); return { ok: false, reason: 'Latex 변환 파일에 Data_DS 시트 없음' }; }

    var matches = kice_keyMatcher_(keywords);
    var cand = [];                                   // [id, latex]
    var matched = 0, emptyB = 0;
    var srcLast = src.getLastRow();
    if (srcLast >= 2) {
      var sv = src.getRange(2, 1, srcLast - 1, 2).getValues();
      for (var i = 0; i < sv.length; i++) {
        var key = String(sv[i][0] || '').trim();
        if (!key || !matches(key)) continue;
        var b = sv[i][1];
        if (!kice_hasValue_(b)) { emptyB++; continue; }
        matched++;
        cand.push([key, b]);
      }
    }

    // alreadyDone: 아카이브에 같은 id + 같은 B 가 있으면 제외
    var alreadyDone = 0;
    if (cand.length) {
      var index = kice_archiveIndex_(shArc, cand.map(function (c) { return c[0]; }));
      var bCache = {};
      cand = cand.filter(function (c) {
        var hits = index[kice_cell_(c[0])] || [];
        for (var h = 0; h < hits.length; h++) {
          var row = hits[h];
          if (bCache[row] === undefined) bCache[row] = shArc.getRange(row, 2).getValue();
          if (kice_sameCell_(bCache[row], c[1])) { alreadyDone++; return false; }
        }
        return true;
      });
    }
    var countMismatch = (!limit && expect !== null) ? ((matched + emptyB) !== expect) : false;
    if (limit && cand.length > limit) cand = cand.slice(0, limit);

    auto.src = { matched: matched, emptyB: emptyB, alreadyDone: alreadyDone, countMismatch: countMismatch, expect: expect, limit: limit || 0 };
    if (countMismatch) kice_log_(runId, 'warn', '개수 불일치: Data_DS 일치 ' + (matched + emptyB) + '행 ≠ Latex ' + expect + '행 (경고만, 진행)');

    // ⑤ 잔재 이관 (D8′)
    var residue = { archived: 0, dup: 0, dropped: 0 };
    try {
      var qLast = shQue.getLastRow();
      if (qLast >= 2) {
        var qv = shQue.getRange(2, 1, qLast - 1, 7).getValues();
        var withResult = [];
        qv.forEach(function (r) {
          var any = kice_hasValue_(r[0]) || kice_hasValue_(r[1]);
          var hasRes = kice_hasValue_(r[4]) || kice_hasValue_(r[5]) || kice_hasValue_(r[6]);
          if (hasRes) withResult.push(r);
          else if (any) residue.dropped++;
        });
        if (withResult.length) {
          var dd = kice_dedupeAgainstArchive_(shArc, withResult);
          var rg = kice_appendArchive_(shArc, dd.keep);
          residue.archived = dd.keep.length;
          residue.dup = dd.dup;
          if (rg.start) residue.range = rg;
        }
      }
    } catch (eRes) {
      var rmsg = '잔재 아카이브 실패: ' + ((eRes && eRes.message) || eRes);
      kice_onBatchFail_(rmsg);
      return { ok: false, reason: rmsg };
    }
    kice_log_(runId, 'residue', 'Que 잔재 이관 ' + residue.archived + '행' + (residue.dup ? ' / 중복 제외 ' + residue.dup : '') + (residue.dropped ? ' / 결과 없음 ' + residue.dropped : ''));

    // ⑥ Que 초기화
    var qLast2 = shQue.getLastRow();
    if (qLast2 >= 2) shQue.getRange(2, 1, qLast2 - 1, 7).clearContent();

    // 적재 대상 0행 → 즉시 완료
    if (!cand.length) {
      auto.stage = 'done';
      auto.total = 0;
      auto.residue = residue;
      auto.summary = kice_merge_(kice_summarize_([]), { archived: 0, dup: 0, poisonCount: 0, countMismatch: countMismatch, alreadyDone: alreadyDone });
      auto.archive = { start: 0, end: 0 };
      auto.hb = kice_now_();
      kice_saveAuto_(auto);
      kice_log_(runId, 'done', '적재 대상 0행 (일치 ' + matched + ' / 빈 B ' + emptyB + ' / 기윤문 ' + alreadyDone + ')');
      SpreadsheetApp.flush();
      return { ok: true, runId: runId, startedAt: auto.startedAt, stage: 'done', rows: 0, src: auto.src, residue: residue };
    }

    // ⑦ 적재 (C 는 비움)
    shQue.getRange(2, 1, cand.length, 2).setValues(cand);
    SpreadsheetApp.flush();

    // ⑧ 상태 — 순서 고정: QUE 먼저, 그 다음 AUTO rewrite
    var rows = [];
    for (var r = 2; r <= cand.length + 1; r++) rows.push(r);
    var qState = { rows: rows, idx: 0, startedAt: auto.startedAt, kice: runId };
    PropertiesService.getScriptProperties().setProperty(QUE_BATCH_KEY, JSON.stringify(qState));

    auto.stage = 'rewrite';
    auto.total = cand.length;
    auto.residue = residue;
    auto.hb = kice_now_();
    kice_saveAuto_(auto);
    kice_log_(runId, 'rewrite', '적재 ' + cand.length + '행 (일치 ' + matched + ' / 빈 B ' + emptyB + ' / 기윤문 제외 ' + alreadyDone + ')');

    // ⑨ 행 처리는 시작하지 않는다 — 러너의 tick 이 돌린다
    return { ok: true, runId: runId, startedAt: auto.startedAt, stage: 'rewrite', rows: cand.length, src: auto.src, residue: residue };

  } finally {
    lock.releaseLock();
  }
}


/* =========================================================
 * 6. tick (§4.4)
 * ========================================================= */

function kice_apiTick_() {
  var auto = kice_loadAuto_(), q = kice_readQueRaw_();
  if (!auto) return { ok: false, reason: 'KICE 상태 없음' };
  var base = { ok: true, runId: auto.runId, total: auto.total || 0 };

  if (kice_isEnded_(auto))                                        // A2: 끝난 뒤의 tick 은 정상 응답
    return kice_merge_(base, { result: 'idle', stage: auto.stage, idx: auto.total || 0 });
  if (auto.stage === 'loading')
    return kice_merge_(base, { result: 'loading', stage: 'loading', idx: 0 });
  if (!q || String(q.kice || '') !== String(auto.runId) || String(q.startedAt || '') !== String(auto.startedAt || ''))
    return kice_merge_(base, { result: 'stopped', stage: 'stopped', derived: true, idx: q ? (Number(q.idx) || 0) : 0 });

  __KICE_TICK = true;
  var t0 = kice_now_();
  var res = batch_continueQueAuto();                              // 'locked'|'idle'|'yield'|'stopped'|'done'|'fatal'
  auto = kice_loadAuto_(); q = kice_readQueRaw_();
  return kice_merge_(base, {
    result: res,
    stage: auto ? auto.stage : null,
    idx: q ? (Number(q.idx) || 0) : (auto ? (auto.total || 0) : 0),
    total: auto ? (auto.total || 0) : base.total,
    elapsedMs: kice_now_() - t0
  });
}


/* =========================================================
 * 7. status (읽기만, D21) · stop (§4.7)
 * ========================================================= */

function kice_apiStatus_() {
  var q = kice_readQueRaw_();                                     // 순서: QUE → AUTO
  var auto = kice_loadAuto_();
  var derived = null;
  if (auto && auto.stage === 'rewrite' &&
      (!q || String(q.kice || '') !== String(auto.runId) || String(q.startedAt || '') !== String(auto.startedAt || ''))) {
    derived = { stage: 'stopped', stoppedFrom: 'rewrite', by: 'menu' };
  }
  var progress = {
    idx: q ? (Number(q.idx) || 0) : (auto && kice_isEnded_(auto) ? (auto.total || 0) : 0),
    total: auto ? (auto.total || 0) : 0,
    currentRow: (q && q.inflight && q.inflight.row) ? q.inflight.row : null,
    lastHeartbeat: auto ? (Number(auto.hb) || null) : null,       // ms epoch
    poisonCount: q ? Number(q.poisonCount || 0) : (auto && auto.summary ? Number(auto.summary.poisonCount || 0) : 0)
  };
  var hasTick = false;
  try {
    hasTick = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'batch_continueQueAuto'; });
  } catch (_) {}
  return {
    ok: true,
    version: (typeof RAPI !== 'undefined' && RAPI.VERSION) || '',
    state: auto || null,
    derived: derived,
    progress: progress,
    logTail: (typeof rapi_logTail_ === 'function') ? rapi_logTail_((typeof RAPI !== 'undefined' && RAPI.LOG_TAIL) || 20) : [],
    hasTick: hasTick,
    at: kice_iso_()
  };
}

function kice_apiStop_(by) {
  var q = kice_readQueRaw_();
  var auto = kice_loadAuto_();
  // 1-C P2: QUE 는 **이 런의 러너 배치**일 때만 지운다 — 덕수님 메뉴 배치(kice 없음)·다른 런의 배치는 건드리지 않는다
  var mine = !!(q && q.kice && auto && q.kice === auto.runId);
  if (mine) {
    _clearQueBatchState_();
    _deleteTriggersByHandler_('batch_continueQueAuto');
  }
  var from = kice_markStopped_(auto, by || '원격 중지');
  return { ok: true, stoppedFrom: from, cleared: mine };
}


/* =========================================================
 * 8. 메뉴 '📋 KICE 자동 상태' (웹앱 경로에서는 부르지 않는다)
 * ========================================================= */

function kice_showStatus() {
  var ui = SpreadsheetApp.getUi();
  var st = kice_apiStatus_();
  var a = st.state;
  if (!a) { ui.alert('KICE 자동 상태', '러너 실행 이력이 없습니다.', ui.ButtonSet.OK); return; }
  var lines = [
    'runId: ' + a.runId,
    'stage: ' + a.stage + (st.derived ? '  (파생: ' + st.derived.stage + ' by ' + st.derived.by + ')' : ''),
    '진행: ' + st.progress.idx + ' / ' + st.progress.total + (st.progress.currentRow ? '  (현재 행 ' + st.progress.currentRow + ')' : ''),
    '마지막 heartbeat: ' + (st.progress.lastHeartbeat ? new Date(st.progress.lastHeartbeat).toLocaleString() : '-'),
    '키워드: ' + ((a.keywords || []).join(' | ') || '-'),
    '시작: ' + (a.startedAt || '-')
  ];
  if (a.src) lines.push('대상: 일치 ' + a.src.matched + ' / 빈 B ' + a.src.emptyB + ' / 기윤문 제외 ' + a.src.alreadyDone + (a.src.countMismatch ? '  ⚠️ 개수 불일치(expect ' + a.src.expect + ')' : ''));
  if (a.residue) lines.push('잔재 이관: ' + a.residue.archived + '행' + (a.residue.dup ? ' / 중복 ' + a.residue.dup : '') + (a.residue.dropped ? ' / 결과 없음 ' + a.residue.dropped : ''));
  if (a.summary) {
    var s = a.summary;
    lines.push('요약: ' + s.rows + '행 — Claude edit ' + s.claude.edit + ' · GPT ' + s.gpt.edit + ' · Gemini ' + s.gemini.edit +
               ' / 실패 ' + (s.claude.fail + s.gpt.fail + s.gemini.fail) + ' / poison ' + (s.poisonCount || 0));
  }
  if (a.archive && a.archive.start) lines.push('아카이브: ' + a.archive.start + '~' + a.archive.end + '행');
  if (a.error) lines.push('오류: ' + a.error);
  if (a.stoppedFrom) lines.push('중지: from ' + a.stoppedFrom + (a.by ? ' by ' + a.by : ''));
  ui.alert('KICE 자동 상태', lines.join('\n'), ui.ButtonSet.OK);
}

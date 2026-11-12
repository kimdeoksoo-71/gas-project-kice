/*******************************************************
 * kiceAuto.gs — KICE 윤문 배치의 상태와 명령 본체 (kice-2.0.0 · 재구축 계획서 v5 §7)
 *
 * 2.0.0 에서 바뀐 것
 *  - 이 프로젝트는 LLM 을 부르지 않는다. 일은 맥미니의 kice_engine 이 하고, 여기는 시트 창구다.
 *    start(러너·메뉴) → claim / heartbeat / submit(엔진) → finish(엔진) → 아카이브 이관.
 *  - `문항검토` 시트 의존 없음 (새 파일 「KICE 윤문」에는 그 시트가 없다 — v5 §17 E6).
 *  - start 순서: 잔재 이관 → 기윤문(alreadyDone) 판정 (BACKLOG 10/3 ④⑤).
 *  - status 의 progress.idx 는 "끝난 행 수"다. 끝난 뒤에도 total 로 바뀌지 않는다 (BACKLOG 10/3).
 *
 * 상태 (ScriptProperties)
 *  - KICE_AUTO_STATE_V2 : {runId, startedAt, stage, keywords, expect, total, src, residue, summary, archive,
 *                          error, stoppedFrom, by, hb(ms), idx, menu, updatedAt(ISO)}
 *                         stage ∈ loading | rewrite | archive | done | error | stopped
 *  - QUE_BATCH_STATE_V4 : {startedAt, kice(runId), menu, rows[], inflight{row:{lease,engineId,at,n}}, poison[], poisonCount}
 *                         rewrite·archive 동안 반드시 살아 있다 — status 의 derived 판정이 이것을 본다 (v5 §20 F4).
 *
 * 러너와의 불변 계약 (v5 §7-3) — 러너 코드는 kice-1.0.0 때와 같다
 *  - start 응답 {ok, runId, startedAt, stage, rows, src, residue} · adopted
 *  - status 응답 state / derived / progress{idx,total,currentRow,lastHeartbeat,poisonCount} / summary / archive / src / residue
 *  - tick 은 항상 ok:true, 잠금 없음, locked·loading 을 오래 돌려주지 않는다 (v5 §20 F3)
 *  - "이미 실행 중" 이 들어간 reason 은 러너가 기다린다(busy)
 *
 * 원칙
 *  - 상태를 읽고 쓰는 명령(start·claim·submit·heartbeat·finish·stop·메뉴)은 모두 스크립트 잠금 안에서 (v5 §7-4 ①)
 *  - claim·submit·heartbeat 는 stage 가 rewrite 일 때만 (②)
 *  - 행의 완료 = Que E열과 G열이 모두 비어 있지 않음. 완료 목록을 속성에 쌓지 않는다 (⑤)
 *  - 아카이브에는 끝에만 붙인다
 *******************************************************/

var KICE_AUTO_KEY        = 'KICE_AUTO_STATE_V2';
var QUE_BATCH_KEY        = 'QUE_BATCH_STATE_V4';
var KICE_LATEX_SS_PROP   = 'LATEX_SS_ID';
var KICE_LATEX_SRC_SHEET = 'Data_DS';
var KICE_QUE_SHEET       = 'Que';
var KICE_CTX_SHEET       = 'QueCtx';
var KICE_OPN_SHEET       = 'Opinions';
var KICE_LOG_SHEET       = 'KICE_Log';
var KICE_ARCHIVE_SHEET   = '아카이브';
var KICE_LIST_CAP        = 30;                      // 상태 속성 9KB 보호 — 목록은 30개까지
var KICE_ROWS_MAX        = 1200;                    // 한 배치의 행 수 상한 (rows 배열이 속성 9KB 를 넘지 않게)
var KICE_TAKEOVER_MS     = 15 * 60 * 1000;          // heartbeat 가 이만큼 멈춘 배치·행은 버려진 것으로 본다
var KICE_MAX_CLAIMS      = 3;                       // 같은 행이 이만큼 claim 될 차례면 (실패)로 확정 (v5 §20 F7)
var KICE_END_STAGES      = ['done', 'error', 'stopped'];
var KICE_CELL_MAX        = 49000;                   // Sheets 셀 한도 5만 자 보호
var KICE_LOCK_WAIT_MS    = 10000;

// Que 시트 열 번호
var COL_QUE_ID     = 1;  // A
var COL_QUE_LATEX  = 2;  // B
var COL_QUE_CLAUDE = 5;  // E  Claude
var COL_QUE_MERGED = 6;  // F  합의 (1.0.0 에서는 GPT)
var COL_QUE_GEMINI = 7;  // G  Gemini

// QueCtx 헤더 (없으면 만든다). 엔진이 submit 의 ctx 로 채우는 열은 뒤쪽 다섯.
var KICE_CTX_HEADERS = ['사설 문항 id', '해설', '그림 링크', 'given_answer', '그림 요약', '조건→역할', '조건 유형 태그', '2028범위 판정'];
var KICE_CTX_FROM_ENGINE = { figSummary: '그림 요약', roles: '조건→역할', tags: '조건 유형 태그', scope: '2028범위 판정' };

var KICE_OPN_HEADERS = ['의견ID', '문항 id', '대상 구절', '신뢰도', 'claude_제안', 'gemini_제안', '카테고리', '이유', '근거',
                        '충돌', '힌트 누출', '답 일치 검증', '모델ID', '사고 수준', '프롬프트 버전'];
var KICE_OPN_FIELDS = {                             // submit 의 opinions[] 필드 이름 → 시트 헤더
  id: '의견ID', itemId: '문항 id', target: '대상 구절', confidence: '신뢰도', claude: 'claude_제안', gemini: 'gemini_제안',
  category: '카테고리', reason: '이유', evidence: '근거', conflict: '충돌', hint: '힌트 누출', answerCheck: '답 일치 검증',
  model: '모델ID', effort: '사고 수준', promptVersion: '프롬프트 버전'
};

var KICE_NONE_LINES = ['수정 구절 없음', '수정구절 없음', '합의 없음', '범위 밖'];   // 첫 줄이 이것이면 none (v5 §8-4)


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

/** 셀에 넣을 문자열 — 5만 자 한도를 넘으면 잘라 표시한다 */
function kice_clip_(v) {
  var s = String(v == null ? '' : v);
  return s.length > KICE_CELL_MAX ? s.slice(0, KICE_CELL_MAX) + '\n…[잘림 ' + s.length + '자]' : s;
}

/** 1~n열 중 아무 값이 있는 마지막 행 (없으면 1) */
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

/** NFC·소문자·부분 문자열 일치 (audition pv_keyMatcher_ 와 같은 규칙) */
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

function kice_toast_(ss, msg, title, sec) {
  try { ss.toast(msg, title || 'KICE 윤문', sec || 5); } catch (_) {}
}

/** 스크립트 잠금 안에서 fn 실행. 잠금을 못 잡으면 {ok:false, reason:'locked'} (v5 §7-4 ①) */
function kice_withLock_(fn) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(KICE_LOCK_WAIT_MS)) return { ok: false, reason: 'locked' };
  try { return fn(); } finally { lock.releaseLock(); }
}

/** 시트의 1행 헤더 → {이름: 열 번호(1-base)}. names 가 주어지면 없는 헤더를 오른쪽에 만든다. */
function kice_headerMap_(sh, names) {
  var lastCol = Math.max(1, sh.getLastColumn());
  var head = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var map = {};
  for (var c = 0; c < head.length; c++) {
    var h = String(head[c] == null ? '' : head[c]).trim();
    if (h && !map[h]) map[h] = c + 1;
  }
  if (names) {
    var next = lastCol + ((head.length === 1 && !String(head[0] || '').trim()) ? 0 : 1);
    if (next < 1) next = 1;
    names.forEach(function (n) {
      if (map[n]) return;
      if (sh.getMaxColumns() < next) sh.insertColumnsAfter(sh.getMaxColumns(), next - sh.getMaxColumns());
      sh.getRange(1, next).setValue(n);
      map[n] = next;
      next++;
    });
  }
  return map;
}

/** 셀의 첫 줄 (HTML 의 <br> 도 줄바꿈으로 본다) */
function kice_firstLine_(v) {
  var s = kice_cell_(v).trim();
  if (!s) return '';
  return s.split(/\n|<br\s*\/?>/i)[0].trim();
}

function kice_randomLease_() {
  return Utilities.getUuid().replace(/-/g, '').slice(0, 12);
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

function kice_isEnded_(auto) {
  return !!auto && KICE_END_STAGES.indexOf(auto.stage) >= 0;
}

/** AUTO 를 'stopped' 로 (진행 상태일 때만). 반환: stoppedFrom 또는 '' */
function kice_markStopped_(auto, by, doneCount) {
  if (!auto || kice_isEnded_(auto)) return '';
  var from = auto.stage;
  auto.stoppedFrom = from;
  auto.stage = 'stopped';
  auto.by = by || '';
  auto.hb = kice_now_();
  if (doneCount !== undefined && doneCount !== null) auto.idx = doneCount;
  kice_saveAuto_(auto);
  kice_log_(auto.runId, 'stopped', (by || '중지') + ' (from ' + from + ')');
  return from;
}

/** 실패 훅 — AUTO stage:'error' */
function kice_fail_(reason) {
  var auto = kice_loadAuto_();
  if (!auto || kice_isEnded_(auto)) return;
  auto.stage = 'error';
  auto.error = String(reason || '');
  auto.hb = kice_now_();
  kice_saveAuto_(auto);
  kice_log_(auto.runId, 'error', auto.error);
}


/* =========================================================
 * 2. QUE 상태 — 소유 판별
 * ========================================================= */

function kice_readQueRaw_() {
  var raw = PropertiesService.getScriptProperties().getProperty(QUE_BATCH_KEY);
  if (!raw) return null;
  try { return JSON.parse(raw); } catch (_) { return null; }
}

function kice_saveQue_(q) {
  PropertiesService.getScriptProperties().setProperty(QUE_BATCH_KEY, JSON.stringify(q));
}

function kice_clearQue_() {
  PropertiesService.getScriptProperties().deleteProperty(QUE_BATCH_KEY);
}

/** AUTO 와 QUE 가 같은 배치를 가리키는가 */
function kice_paired_(auto, q) {
  return !!auto && !!q &&
    String(q.kice || '') === String(auto.runId || '') &&
    String(q.startedAt || '') === String(auto.startedAt || '');
}

/**
 * 엔진 요청(claim 뒤의 submit·heartbeat·finish)의 공통 검사.
 * 통과하면 {auto, q}, 아니면 {err:{ok:false, reason, …}}.
 *   stages: 허용하는 stage 목록
 */
function kice_guard_(p, stages) {
  var auto = kice_loadAuto_(), q = kice_readQueRaw_();
  if (!auto) return { err: { ok: false, reason: 'not_running', stage: null } };
  if (stages.indexOf(auto.stage) < 0) return { err: { ok: false, reason: 'not_running', stage: auto.stage } };
  if (!kice_paired_(auto, q)) return { err: { ok: false, reason: 'not_running', stage: 'stopped' } };
  if (String(p.runId || '') !== String(auto.runId) || String(p.startedAt || '') !== String(auto.startedAt)) {
    return { err: { ok: false, reason: 'not_owner', runId: auto.runId } };
  }
  return { auto: auto, q: q };
}

/** Que 의 rows 행들을 읽어 [{row, v:[A..G]}] — rows 는 오름차순이 아니어도 된다 */
function kice_readQueRows_(shQue, rows) {
  if (!rows || !rows.length) return [];
  var minR = Math.min.apply(null, rows), maxR = Math.max.apply(null, rows);
  var vals = shQue.getRange(minR, 1, maxR - minR + 1, 7).getValues();
  return rows.map(function (r) { return { row: r, v: vals[r - minR] }; });
}

/** 한 행이 끝났는가 — E 와 G 가 모두 차 있다 (v5 §7-4 ⑤) */
function kice_rowDone_(v) {
  return kice_hasValue_(v[COL_QUE_CLAUDE - 1]) && kice_hasValue_(v[COL_QUE_GEMINI - 1]);
}

function kice_countDone_(shQue, rows) {
  var n = 0;
  kice_readQueRows_(shQue, rows).forEach(function (x) { if (kice_rowDone_(x.v)) n++; });
  return n;
}

function kice_inflightRows_(q) {
  var out = [];
  var f = (q && q.inflight) || {};
  for (var k in f) out.push(Number(k));
  out.sort(function (a, b) { return a - b; });
  return out;
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
 * Que 행(A:G 7열 배열)들 중 아카이브에 "같은 id + 같은 E·F·G" 가 이미 있는 행을 제외.
 * 셀 전문 비교 — 그래서 셀에는 시각·runId 같은 가변 값을 넣지 않고, 제출 뒤에는 고치지 않는다 (v5 §8-4).
 */
function kice_dedupeAgainstArchive_(sh, rows7) {
  if (!rows7.length) return { keep: [], dup: 0 };
  var index = kice_archiveIndex_(sh, rows7.map(function (r) { return r[0]; }));
  var cache = {};
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

/** 아카이브 끝에 7열 행들을 이어 붙인다 → {start, end} */
function kice_appendArchive_(sh, rows7) {
  if (!rows7.length) return { start: 0, end: 0 };
  var last = kice_lastDataRow_(sh, 7);
  var start = last >= 2 ? last + 1 : 2;
  if (sh.getMaxRows() < start + rows7.length - 1) sh.insertRowsAfter(sh.getMaxRows(), start + rows7.length - 1 - sh.getMaxRows());
  if (sh.getMaxColumns() < 7) sh.insertColumnsAfter(sh.getMaxColumns(), 7 - sh.getMaxColumns());
  sh.getRange(start, 1, rows7.length, 7).setValues(rows7);
  return { start: start, end: start + rows7.length - 1 };
}

/**
 * 결과 칸 분류 (v5 §8-4): empty / none(첫 줄이 '수정 구절 없음'·'합의 없음'·'범위 밖') / fail('(실패)'로 시작) / edit
 */
function kice_classifyCell_(v) {
  var first = kice_firstLine_(v);
  if (!first) return kice_cell_(v).trim() ? 'edit' : 'empty';
  if (KICE_NONE_LINES.indexOf(first) >= 0) return 'none';
  if (first.indexOf('(실패)') === 0) return 'fail';
  return 'edit';
}

/**
 * 열별 집계. 키 이름 claude / gpt / gemini 는 러너 장부(K_edit_*)가 읽으므로 그대로 둔다.
 * 2.0.0 에서 gpt = F열(합의) 집계다 (v5 §7-3).
 */
function kice_summarize_(rows7) {
  var mk = function () { return { edit: 0, none: 0, fail: 0, empty: 0 }; };
  var sum = { rows: rows7.length, claude: mk(), gpt: mk(), gemini: mk(), outOfScope: 0 };
  rows7.forEach(function (r) {
    sum.claude[kice_classifyCell_(r[4])]++;
    sum.gpt[kice_classifyCell_(r[5])]++;
    sum.gemini[kice_classifyCell_(r[6])]++;
    if (kice_firstLine_(r[4]) === '범위 밖') sum.outOfScope++;
  });
  return sum;
}

/**
 * 완료 처리 — finish 가 부른다. 잠금 안, auto.stage 는 rewrite 또는 (재진입) archive.
 *  - 러너 배치: 중복 제외 → 아카이브 끝에 이관 → summary → done
 *  - 메뉴 배치(q.menu): 아카이브로 옮기지 않고 Que 에 남긴다. 다음 러너 start 의 잔재 이관이 옮긴다 (v5 §18 N14)
 * 도중에 실행이 죽으면 stage 가 archive 로 남는다 → finish 재호출이 다시 실행하고 dedupe 가 중복 행을 막는다 (v5 §7-4 ⑥).
 */
function kice_completeBatch_(auto, q) {
  try {
    auto.stage = 'archive';
    auto.hb = kice_now_();
    kice_saveAuto_(auto);

    var ss = SpreadsheetApp.getActive();
    var shQue = ss.getSheetByName(KICE_QUE_SHEET);
    var rows = Array.isArray(q.rows) ? q.rows : [];
    var rows7 = kice_readQueRows_(shQue, rows).map(function (x) { return x.v; })
      .filter(function (r) { return kice_hasValue_(r[0]) || kice_hasValue_(r[1]); });

    var range = { start: 0, end: 0 }, dd = { keep: [], dup: 0 };
    if (!q.menu) {
      var shArc = kice_archiveSheet_();
      dd = kice_dedupeAgainstArchive_(shArc, rows7);
      range = kice_appendArchive_(shArc, dd.keep);
    }

    var summary = kice_summarize_(rows7);
    summary.archived = dd.keep.length;
    summary.dup = dd.dup;
    summary.poisonCount = Number(q.poisonCount || 0);
    summary.countMismatch = !!(auto.src && auto.src.countMismatch);
    summary.alreadyDone = Number(auto.src && auto.src.alreadyDone || 0);

    auto.stage = 'done';
    auto.summary = summary;
    auto.archive = range;
    auto.idx = rows7.length;
    auto.hb = kice_now_();
    kice_saveAuto_(auto);
    kice_clearQue_();
    kice_log_(auto.runId, 'done',
      '윤문 ' + summary.rows + '행 → ' + (q.menu ? 'Que 에 보존(메뉴 배치)' : '아카이브 ' + (range.start ? range.start + '~' + range.end : '추가 없음')) +
      (dd.dup ? ' (중복 제외 ' + dd.dup + ')' : '') +
      ' · Claude edit ' + summary.claude.edit + ' / 합의 ' + summary.gpt.edit + ' / Gemini ' + summary.gemini.edit +
      ' · 실패 C' + summary.claude.fail + ' G' + summary.gemini.fail +
      (summary.outOfScope ? ' · 범위 밖 ' + summary.outOfScope : '') +
      (summary.poisonCount ? ' · 포기 ' + summary.poisonCount : ''));
    return { ok: true, stage: 'done', summary: summary, archive: range };
  } catch (e) {
    var msg = 'archive: ' + ((e && e.message) || e);
    kice_fail_(msg);
    return { ok: false, reason: msg };
  }
}


/* =========================================================
 * 4. start (러너)
 * ========================================================= */

function kice_latexSsId_() {
  return (PropertiesService.getScriptProperties().getProperty(KICE_LATEX_SS_PROP) || '').trim();
}

/** 본문에서 Drive 그림 링크만 뽑는다 (`![name](url)` 마크다운 — Data_DS 의 그림 표기) */
function kice_figLinks_(text) {
  var out = [], seen = {};
  var re = /!\[[^\]]*\]\((https?:\/\/[^)\s]+)\)/g, m;
  var s = String(text || '');
  while ((m = re.exec(s)) !== null) { if (!seen[m[1]]) { seen[m[1]] = true; out.push(m[1]); } }
  return out;
}

/**
 * 진행 중인 배치가 있는지 본다. 잠금 안에서 호출.
 * 반환: null(시작해도 됨) 또는 {ok:false, reason} (거절). 버려진 배치는 여기서 인수 처리한다.
 */
function kice_takeoverOrBusy_(auto, q, newRunId) {
  if (!q && !(auto && !kice_isEnded_(auto))) return null;
  var alive = auto && !kice_isEnded_(auto) && (kice_now_() - Number(auto.hb || 0)) <= KICE_TAKEOVER_MS;
  var same = auto && auto.runId === newRunId;
  if (alive && !same && kice_paired_(auto, q)) {
    return { ok: false, reason: '이미 실행 중 — ' + (q.menu ? '메뉴 배치 ' : '러너 배치 ') + auto.runId +
                                ' (' + (Array.isArray(q.rows) ? q.rows.length : 0) + '행, 단계 ' + auto.stage + ')' };
  }
  // 버려진 배치(heartbeat 15분 정지) · 짝이 안 맞는 잔재 · 같은 runId 의 loading 잔재 → 인수
  kice_clearQue_();
  if (auto && !kice_isEnded_(auto) && !same) kice_markStopped_(auto, 'takeover by ' + newRunId);
  if (q) kice_log_(newRunId, 'takeover', '버려진 배치 인수: ' + (q.kice || '(runId 없음)'));
  return null;
}

/**
 * Que 에 남은 결과 행을 아카이브로 옮긴다 (잔재 이관). 잠금 안에서 호출.
 * 반환 {archived, dup, dropped, range?}
 */
function kice_moveResidue_(shQue, shArc) {
  var residue = { archived: 0, dup: 0, dropped: 0 };
  var qLast = shQue.getLastRow();
  if (qLast < 2) return residue;
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
  return residue;
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
  if (!lock.tryLock(KICE_LOCK_WAIT_MS)) return { ok: false, reason: '잠금 대기 초과 — 이미 실행 중' };

  try {
    var ss = SpreadsheetApp.getActive();
    var auto = kice_loadAuto_();

    // ① 멱등 — 같은 runId 가 이미 시작됐고 loading 이 아니면 그대로 인정
    if (auto && auto.runId === runId && auto.stage !== 'loading') {
      return { ok: true, adopted: true, runId: runId, startedAt: auto.startedAt, stage: auto.stage, rows: auto.total || 0 };
    }

    // ② 가드 / 인수
    var busy = kice_takeoverOrBusy_(auto, kice_readQueRaw_(), runId);
    if (busy) return busy;

    // ③ AUTO 선기록 (loading)
    auto = {
      runId: runId, stage: 'loading', keywords: keywords, expect: expect,
      startedAt: kice_iso_(), hb: kice_now_(), total: 0, idx: 0, menu: false,
      src: null, residue: null, summary: null, archive: null, error: '', stoppedFrom: '', by: ''
    };
    kice_saveAuto_(auto);
    kice_log_(runId, 'start', '키워드: ' + keywords.join(' | ') + (expect !== null ? ' · expect ' + expect : '') + (limit ? ' · limit ' + limit : ''));

    var shQue = ss.getSheetByName(KICE_QUE_SHEET);
    if (!shQue) { kice_fail_('Que 시트를 찾지 못했습니다.'); return { ok: false, reason: 'Que 시트를 찾지 못했습니다.' }; }
    var shArc = kice_archiveSheet_();

    // ④ 잔재 이관 — 기윤문 판정보다 먼저 (지난 배치가 Que 에 남긴 결과가 아카이브에 들어간 뒤에 판정해야 재런 때 중복 윤문이 없다)
    var residue;
    try {
      residue = kice_moveResidue_(shQue, shArc);
    } catch (eRes) {
      var rmsg = '잔재 아카이브 실패: ' + ((eRes && eRes.message) || eRes);
      kice_fail_(rmsg);
      return { ok: false, reason: rmsg };
    }
    kice_log_(runId, 'residue', 'Que 잔재 이관 ' + residue.archived + '행' + (residue.dup ? ' / 중복 제외 ' + residue.dup : '') + (residue.dropped ? ' / 결과 없음 ' + residue.dropped : ''));

    // ⑤ 대상 선택 — Data_DS A열로 먼저 고르고, 걸린 구간만 A:D 를 읽는다
    var ssId = kice_latexSsId_();
    if (!ssId) { kice_fail_('스크립트 속성 LATEX_SS_ID 가 비어 있습니다.'); return { ok: false, reason: 'LATEX_SS_ID 미설정' }; }
    var src = SpreadsheetApp.openById(ssId).getSheetByName(KICE_LATEX_SRC_SHEET);
    if (!src) { kice_fail_('Latex 변환 파일에 ' + KICE_LATEX_SRC_SHEET + ' 시트가 없습니다.'); return { ok: false, reason: 'Latex 변환 파일에 Data_DS 시트 없음' }; }

    var matches = kice_keyMatcher_(keywords);
    var cand = [];                                   // [id, problem, solution, givenAnswer]
    var matched = 0, emptyB = 0;
    var srcLast = src.getLastRow();
    if (srcLast >= 2) {
      var keys = src.getRange(2, 1, srcLast - 1, 1).getValues();
      var hitRows = [];
      for (var i = 0; i < keys.length; i++) {
        var key = String(keys[i][0] || '').trim();
        if (key && matches(key)) hitRows.push(i + 2);
      }
      if (hitRows.length) {
        var minR = hitRows[0], maxR = hitRows[hitRows.length - 1];
        var block = src.getRange(minR, 1, maxR - minR + 1, 4).getValues();
        hitRows.forEach(function (r) {
          var v = block[r - minR];
          if (!kice_hasValue_(v[1])) { emptyB++; return; }
          matched++;
          cand.push([String(v[0] || '').trim(), v[1], v[2], v[3]]);
        });
      }
    }

    // ⑥ 기윤문 제외: 아카이브에 같은 id + 같은 B 가 있으면 뺀다
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
    if (cand.length > KICE_ROWS_MAX) {
      kice_fail_('적재 대상 ' + cand.length + '행이 상한 ' + KICE_ROWS_MAX + '행을 넘습니다. 키워드를 나눠 주세요.');
      return { ok: false, reason: '적재 대상이 상한(' + KICE_ROWS_MAX + '행)을 넘음: ' + cand.length + '행' };
    }

    auto.src = { matched: matched, emptyB: emptyB, alreadyDone: alreadyDone, countMismatch: countMismatch, expect: expect, limit: limit || 0 };
    if (countMismatch) kice_log_(runId, 'warn', '개수 불일치: Data_DS 일치 ' + (matched + emptyB) + '행 ≠ Latex ' + expect + '행 (경고만, 진행)');

    // ⑦ Que·QueCtx 초기화
    var qLast2 = shQue.getLastRow();
    if (qLast2 >= 2) shQue.getRange(2, 1, qLast2 - 1, 7).clearContent();
    var shCtx = ss.getSheetByName(KICE_CTX_SHEET) || ss.insertSheet(KICE_CTX_SHEET);
    var ctxMap = kice_headerMap_(shCtx, KICE_CTX_HEADERS);
    var cLast = shCtx.getLastRow();
    if (cLast >= 2) shCtx.getRange(2, 1, cLast - 1, shCtx.getLastColumn()).clearContent();

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

    // ⑧ 적재 — Que A:B, QueCtx(같은 행 번호에 id·해설·그림 링크·given_answer)
    shQue.getRange(2, 1, cand.length, 2).setValues(cand.map(function (c) { return [c[0], c[1]]; }));
    var width = shCtx.getLastColumn();
    var ctxRows = cand.map(function (c) {
      var row = [];
      for (var k = 0; k < width; k++) row.push('');
      row[ctxMap['사설 문항 id'] - 1] = c[0];
      row[ctxMap['해설'] - 1] = kice_clip_(c[2]);
      row[ctxMap['그림 링크'] - 1] = kice_figLinks_(c[1]).concat(kice_figLinks_(c[2])).join('\n');
      row[ctxMap['given_answer'] - 1] = String(c[3] == null ? '' : c[3]);
      return row;
    });
    if (shCtx.getMaxRows() < cand.length + 1) shCtx.insertRowsAfter(shCtx.getMaxRows(), cand.length + 1 - shCtx.getMaxRows());
    shCtx.getRange(2, 1, ctxRows.length, width).setValues(ctxRows);
    SpreadsheetApp.flush();

    // ⑨ 상태 — 순서 고정: QUE 먼저, 그 다음 AUTO rewrite
    var rows = [];
    for (var r = 2; r <= cand.length + 1; r++) rows.push(r);
    kice_saveQue_({ startedAt: auto.startedAt, kice: runId, menu: false, rows: rows, inflight: {}, poison: [], poisonCount: 0 });

    auto.stage = 'rewrite';
    auto.total = cand.length;
    auto.residue = residue;
    auto.hb = kice_now_();
    kice_saveAuto_(auto);
    kice_log_(runId, 'rewrite', '적재 ' + cand.length + '행 (일치 ' + matched + ' / 빈 B ' + emptyB + ' / 기윤문 제외 ' + alreadyDone + ') — 엔진 대기');

    return { ok: true, runId: runId, startedAt: auto.startedAt, stage: 'rewrite', rows: cand.length, src: auto.src, residue: residue };

  } finally {
    lock.releaseLock();
  }
}


/* =========================================================
 * 5. 엔진 명령 — claim / heartbeat / submit / finish
 * ========================================================= */

/** QueCtx 에서 row 행을 읽되, 그 행의 id 가 Que 의 id 와 같을 때만 쓴다 (메뉴 배치는 QueCtx 가 없을 수 있다) */
function kice_ctxOf_(shCtx, ctxMap, ctxVals, row, id) {
  var empty = { solution: '', figLinks: [], givenAnswer: '' };
  if (!shCtx || !ctxVals) return empty;
  var v = ctxVals[row - 2];
  if (!v) return empty;
  if (String(v[ctxMap['사설 문항 id'] - 1] || '').trim() !== String(id || '').trim()) return empty;
  return {
    solution: String(v[ctxMap['해설'] - 1] || ''),
    figLinks: String(v[ctxMap['그림 링크'] - 1] || '').split('\n').map(function (s) { return s.trim(); }).filter(Boolean),
    givenAnswer: String(v[ctxMap['given_answer'] - 1] || '')
  };
}

/**
 * claim — 엔진이 처리할 행을 받아 간다.
 *   n(기본 3, 최대 10) · engineId
 * 응답 {ok, runId, startedAt, total, done, remaining, inflight, items:[{row, id, problem, solution, figLinks[], givenAnswer, lease, attempt}]}
 *  - 배치가 없으면 {ok:false, reason:'not_running', stage}
 *  - 남은 행도 inflight 도 없으면 items:[] · remaining:0 → 엔진이 finish 를 부른다
 */
function kice_apiClaim_(p) {
  p = p || {};
  var n = parseInt(p.n, 10);
  if (!isFinite(n) || n < 1) n = 3;
  if (n > 10) n = 10;
  var engineId = String(p.engineId || '').slice(0, 80);

  return kice_withLock_(function () {
    var auto = kice_loadAuto_(), q = kice_readQueRaw_();
    if (!auto || auto.stage !== 'rewrite') return { ok: false, reason: 'not_running', stage: auto ? auto.stage : null };
    if (!kice_paired_(auto, q)) return { ok: false, reason: 'not_running', stage: 'stopped' };

    var ss = SpreadsheetApp.getActive();
    var shQue = ss.getSheetByName(KICE_QUE_SHEET);
    var rows = Array.isArray(q.rows) ? q.rows : [];
    var read = kice_readQueRows_(shQue, rows);
    q.inflight = q.inflight || {};
    var now = kice_now_();

    var shCtx = ss.getSheetByName(KICE_CTX_SHEET), ctxMap = null, ctxVals = null;
    if (shCtx && shCtx.getLastRow() >= 2) {
      ctxMap = kice_headerMap_(shCtx, KICE_CTX_HEADERS);
      ctxVals = shCtx.getRange(2, 1, shCtx.getLastRow() - 1, shCtx.getLastColumn()).getValues();
    }

    var items = [], done = 0, changed = false, poisoned = [];
    read.forEach(function (x) {
      if (kice_rowDone_(x.v)) {
        done++;
        if (q.inflight[x.row]) { delete q.inflight[x.row]; changed = true; }
        return;
      }
      var f = q.inflight[x.row];
      if (f && (now - Number(f.at || 0)) <= KICE_TAKEOVER_MS) return;          // 다른 claim 이 살아 있다
      if (items.length >= n) return;
      var attempt = f ? Number(f.n || 1) + 1 : 1;
      if (attempt > KICE_MAX_CLAIMS) {                                          // 엔진이 이 행을 끝내지 못한다 → 포기 (v5 §20 F7)
        var note = '(실패) 엔진이 이 행을 ' + (attempt - 1) + '회 끝내지 못함 — 건너뜀';
        shQue.getRange(x.row, COL_QUE_CLAUDE, 1, 3).setValues([[note, note, note]]);
        delete q.inflight[x.row];
        q.poison = kice_capList_(q.poison || []);
        if (q.poison.length < KICE_LIST_CAP && q.poison.indexOf(x.row) < 0) q.poison.push(x.row);
        q.poisonCount = Number(q.poisonCount || 0) + 1;
        poisoned.push(x.row);
        done++; changed = true;
        return;
      }
      var lease = kice_randomLease_();
      q.inflight[x.row] = { lease: lease, engineId: engineId, at: now, n: attempt };
      changed = true;
      var id = String(x.v[COL_QUE_ID - 1] || '').trim();
      var ctx = kice_ctxOf_(shCtx, ctxMap, ctxVals, x.row, id);
      items.push({ row: x.row, id: id, problem: String(x.v[COL_QUE_LATEX - 1] || ''), solution: ctx.solution,
                   figLinks: ctx.figLinks, givenAnswer: ctx.givenAnswer, lease: lease, attempt: attempt });
    });

    if (changed) kice_saveQue_(q);
    if (items.length || poisoned.length) { auto.hb = now; kice_saveAuto_(auto); }
    poisoned.forEach(function (r) { kice_log_(auto.runId, 'poison', 'row ' + r + ' — ' + (KICE_MAX_CLAIMS) + '회 claim 뒤에도 미완료, (실패)로 확정'); });
    if (items.length) kice_log_(auto.runId, 'claim', 'rows ' + items.map(function (it) { return it.row + (it.attempt > 1 ? '(재' + it.attempt + ')' : ''); }).join(',') + ' · ' + engineId);

    var inflightN = kice_inflightRows_(q).length;
    return { ok: true, runId: auto.runId, startedAt: auto.startedAt, menu: !!q.menu, total: rows.length, done: done,
             remaining: rows.length - done, inflight: inflightN, items: items };
  });
}

/**
 * heartbeat — 엔진이 LLM 호출 중에 3분마다 부른다.
 *   runId · startedAt · leases = "row:lease,row:lease"
 * 응답 {ok, alive:[row…], lost:[row…]} — lost 는 엔진이 버려야 하는 작업
 */
function kice_apiHeartbeat_(p) {
  p = p || {};
  return kice_withLock_(function () {
    var g = kice_guard_(p, ['rewrite']);
    if (g.err) return g.err;
    var auto = g.auto, q = g.q, now = kice_now_();
    var alive = [], lost = [];
    var pairs = Array.isArray(p.leases) ? p.leases : String(p.leases || '').split(',');
    pairs.forEach(function (s) {
      var row, lease;
      if (s && typeof s === 'object') { row = Number(s.row); lease = String(s.lease || ''); }
      else { var ab = String(s || '').split(':'); row = Number(ab[0]); lease = String(ab[1] || ''); }
      if (!row) return;
      var f = (q.inflight || {})[row];
      if (f && f.lease === lease) { f.at = now; alive.push(row); } else lost.push(row);
    });
    if (alive.length) kice_saveQue_(q);
    auto.hb = now;
    kice_saveAuto_(auto);
    return { ok: true, alive: alive, lost: lost };
  });
}

/**
 * submit (POST) — 한 행의 결과.
 *   runId · startedAt · row · lease · id · cells{E,F,G} · opinions[] · ctx{figSummary, roles, tags, scope}
 * 검사: 단계 rewrite → 소유(runId+startedAt) → 행의 id 일치 → lease 일치.
 * 응답이 유실돼 같은 결과를 다시 보낸 경우(lease 는 이미 풀림)는 셀이 같으면 ok 로 답한다.
 */
function kice_apiSubmit_(p) {
  p = p || {};
  var row = Number(p.row);
  var cells = p.cells || {};
  var E = kice_clip_(cells.E), F = kice_clip_(cells.F), G = kice_clip_(cells.G);
  if (!row || row < 2) return { ok: false, reason: 'row 가 없습니다.' };
  if (!kice_hasValue_(E) || !kice_hasValue_(G)) return { ok: false, reason: 'cells.E 와 cells.G 는 비울 수 없습니다 (의견이 없으면 "수정 구절 없음").' };
  if (!kice_hasValue_(F)) F = '합의 없음';

  return kice_withLock_(function () {
    var g = kice_guard_(p, ['rewrite']);
    if (g.err) return g.err;
    var auto = g.auto, q = g.q;
    if ((q.rows || []).indexOf(row) < 0) return { ok: false, reason: 'row_not_in_batch', row: row };

    var ss = SpreadsheetApp.getActive();
    var shQue = ss.getSheetByName(KICE_QUE_SHEET);
    var cur = shQue.getRange(row, 1, 1, 7).getValues()[0];
    var id = String(cur[COL_QUE_ID - 1] || '').trim();
    if (String(p.id || '').trim() !== id) return { ok: false, reason: 'row_mismatch', row: row, id: id };

    var f = (q.inflight || {})[row];
    if (!f || f.lease !== String(p.lease || '')) {
      var same = kice_sameCell_(cur[COL_QUE_CLAUDE - 1], E) && kice_sameCell_(cur[COL_QUE_MERGED - 1], F) && kice_sameCell_(cur[COL_QUE_GEMINI - 1], G);
      if (same) return { ok: true, row: row, duplicate: true };
      return { ok: false, reason: 'stale_lease', row: row };
    }

    shQue.getRange(row, COL_QUE_CLAUDE, 1, 3).setValues([[E, F, G]]);
    var nOpn = 0, warn = '';
    try { nOpn = kice_upsertOpinions_(ss, id, p.opinions); } catch (eO) { warn = 'opinions: ' + ((eO && eO.message) || eO); }
    try { kice_writeCtx_(ss, row, id, p.ctx); } catch (eC) { warn += (warn ? ' / ' : '') + 'ctx: ' + ((eC && eC.message) || eC); }

    delete q.inflight[row];
    kice_saveQue_(q);
    auto.hb = kice_now_();
    kice_saveAuto_(auto);
    kice_log_(auto.runId, 'submit', 'row ' + row + ' ' + id + ' · C ' + kice_classifyCell_(E) + ' / 합의 ' + kice_classifyCell_(F) + ' / G ' + kice_classifyCell_(G) +
              ' · 의견 ' + nOpn + (warn ? ' ⚠ ' + warn : ''));
    var out = { ok: true, row: row, opinions: nOpn };
    if (warn) out.warn = warn;
    return out;
  });
}

/** Opinions 시트에 의견ID 로 upsert. 반환: 쓴 의견 수 */
function kice_upsertOpinions_(ss, itemId, opinions) {
  if (!Array.isArray(opinions) || !opinions.length) return 0;
  var sh = ss.getSheetByName(KICE_OPN_SHEET) || ss.insertSheet(KICE_OPN_SHEET);
  var map = kice_headerMap_(sh, KICE_OPN_HEADERS);
  var width = sh.getLastColumn();
  var idCol = map['의견ID'];
  var last = sh.getLastRow();
  var index = {};
  if (last >= 2) {
    var ids = sh.getRange(2, idCol, last - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) { var k = String(ids[i][0] || '').trim(); if (k) index[k] = i + 2; }
  }
  var appendRows = [], n = 0;
  opinions.forEach(function (o) {
    if (!o) return;
    var oid = String(o.id || '').trim();
    if (!oid) return;
    var rowVals = [];
    for (var c = 0; c < width; c++) rowVals.push('');
    for (var field in KICE_OPN_FIELDS) {
      var col = map[KICE_OPN_FIELDS[field]];
      var val = (field === 'itemId' && (o[field] === undefined || o[field] === null)) ? itemId : o[field];
      if (val !== undefined && val !== null) rowVals[col - 1] = kice_clip_(typeof val === 'object' ? JSON.stringify(val) : val);
    }
    if (index[oid]) sh.getRange(index[oid], 1, 1, width).setValues([rowVals]);
    else { index[oid] = -1; appendRows.push(rowVals); }
    n++;
  });
  if (appendRows.length) {
    var start = Math.max(2, kice_lastDataRow_(sh, width) + 1);
    if (sh.getMaxRows() < start + appendRows.length - 1) sh.insertRowsAfter(sh.getMaxRows(), start + appendRows.length - 1 - sh.getMaxRows());
    sh.getRange(start, 1, appendRows.length, width).setValues(appendRows);
  }
  return n;
}

/** QueCtx 의 row 행에 엔진이 만든 분석 결과를 적는다 (id 가 다르면 id 부터 새로 적는다 — 메뉴 배치) */
function kice_writeCtx_(ss, row, id, ctx) {
  if (!ctx || typeof ctx !== 'object') return;
  var sh = ss.getSheetByName(KICE_CTX_SHEET) || ss.insertSheet(KICE_CTX_SHEET);
  var map = kice_headerMap_(sh, KICE_CTX_HEADERS);
  if (sh.getMaxRows() < row) sh.insertRowsAfter(sh.getMaxRows(), row - sh.getMaxRows());
  var idCell = sh.getRange(row, map['사설 문항 id']);
  if (String(idCell.getValue() || '').trim() !== String(id || '').trim()) {
    sh.getRange(row, 1, 1, sh.getLastColumn()).clearContent();
    idCell.setValue(id);
  }
  for (var field in KICE_CTX_FROM_ENGINE) {
    var val = ctx[field];
    if (val === undefined || val === null) continue;
    sh.getRange(row, map[KICE_CTX_FROM_ENGINE[field]]).setValue(kice_clip_(typeof val === 'object' ? JSON.stringify(val) : val));
  }
}

/**
 * finish — 엔진이 마지막 submit 뒤에 부른다.
 *   runId · startedAt
 *  - rewrite: 미완료 행이나 inflight 가 남아 있으면 {ok:false, reason:'remaining', remaining, inflight}
 *  - archive 에 15분 넘게 머문 배치: 아카이브 단계를 다시 실행한다 (재진입)
 *  - 이미 done 이면 {ok:true, stage:'done', already:true}
 */
function kice_apiFinish_(p) {
  p = p || {};
  return kice_withLock_(function () {
    var auto = kice_loadAuto_(), q = kice_readQueRaw_();
    if (!auto) return { ok: false, reason: 'not_running', stage: null };
    if (String(p.runId || '') !== String(auto.runId) || String(p.startedAt || '') !== String(auto.startedAt)) {
      return { ok: false, reason: 'not_owner', runId: auto.runId };
    }
    if (auto.stage === 'done') return { ok: true, stage: 'done', already: true, summary: auto.summary, archive: auto.archive };
    if (auto.stage !== 'rewrite' && auto.stage !== 'archive') return { ok: false, reason: 'not_running', stage: auto.stage };
    if (!kice_paired_(auto, q)) return { ok: false, reason: 'not_running', stage: 'stopped' };

    if (auto.stage === 'archive') {
      if ((kice_now_() - Number(auto.hb || 0)) <= KICE_TAKEOVER_MS) return { ok: false, reason: 'archiving', stage: 'archive' };
      kice_log_(auto.runId, 'archive', '아카이브 단계 재진입 (15분 넘게 멈춤)');
      return kice_completeBatch_(auto, q);
    }

    var shQue = SpreadsheetApp.getActive().getSheetByName(KICE_QUE_SHEET);
    var rows = Array.isArray(q.rows) ? q.rows : [];
    var done = kice_countDone_(shQue, rows);
    var inflightN = kice_inflightRows_(q).length;
    if (done < rows.length || inflightN) return { ok: false, reason: 'remaining', remaining: rows.length - done, inflight: inflightN };
    return kice_completeBatch_(auto, q);
  });
}


/* =========================================================
 * 6. tick (러너 펌프) — 아무 일도 하지 않는다
 * ========================================================= */

/**
 * 러너(kice-1.0.0 시절 코드)의 펌프가 30초마다 부른다. 2.0.0 에서는 엔진이 일하므로 상태만 알려 준다.
 * 지킬 것 (v5 §7-1, §20 F3): 항상 ok:true · 잠금을 잡지 않는다 · stage 는 실제 값.
 * 펌프는 stage 가 rewrite 가 아니거나 result 가 stopped 면 끝난다.
 */
function kice_apiTick_() {
  var auto = kice_loadAuto_(), q = kice_readQueRaw_();
  if (!auto) return { ok: false, reason: 'KICE 상태 없음' };
  var base = { ok: true, runId: auto.runId, total: auto.total || 0, engine: true };
  if (auto.stage === 'loading') return kice_merge_(base, { result: 'loading', stage: 'loading', idx: 0 });
  if (auto.stage === 'rewrite' && !kice_paired_(auto, q)) {
    return kice_merge_(base, { result: 'stopped', stage: 'stopped', derived: true, idx: Number(auto.idx || 0) });
  }
  return kice_merge_(base, { result: 'idle', stage: auto.stage, idx: Number(auto.idx || 0) });
}


/* =========================================================
 * 7. status (읽기만) · stop
 * ========================================================= */

function kice_apiStatus_() {
  var q = kice_readQueRaw_();                                     // 순서: QUE → AUTO
  var auto = kice_loadAuto_();
  var derived = null;
  if (auto && auto.stage === 'rewrite' && !kice_paired_(auto, q)) {
    derived = { stage: 'stopped', stoppedFrom: 'rewrite', by: 'menu' };
  }
  var idx = auto ? Number(auto.idx || 0) : 0;
  var inflightRows = [];
  if (auto && (auto.stage === 'rewrite' || auto.stage === 'archive') && kice_paired_(auto, q)) {
    try {
      var shQue = SpreadsheetApp.getActive().getSheetByName(KICE_QUE_SHEET);
      idx = kice_countDone_(shQue, Array.isArray(q.rows) ? q.rows : []);
    } catch (_) {}
    inflightRows = kice_inflightRows_(q);
  }
  var progress = {
    idx: idx,                                                     // 끝난 행 수 (끝난 뒤에도 그대로)
    total: auto ? (auto.total || 0) : 0,
    currentRow: inflightRows.length ? inflightRows[0] : null,
    lastHeartbeat: auto ? (Number(auto.hb) || null) : null,       // ms epoch — 마지막 claim·submit·heartbeat
    poisonCount: q ? Number(q.poisonCount || 0) : (auto && auto.summary ? Number(auto.summary.poisonCount || 0) : 0),
    inflight: inflightRows.length
  };
  return {
    ok: true,
    version: (typeof RAPI !== 'undefined' && RAPI.VERSION) || '',
    state: auto || null,
    derived: derived,
    progress: progress,
    logTail: (typeof rapi_logTail_ === 'function') ? rapi_logTail_((typeof RAPI !== 'undefined' && RAPI.LOG_TAIL) || 20) : [],
    hasTick: false,
    at: kice_iso_()
  };
}

/**
 * stop — 러너(건너뜀·/stop)와 메뉴 ⏹ 가 부른다. QUE 는 지금의 AUTO 와 짝이 맞을 때만 지운다.
 * inflight 를 모두 지우고 stage 를 stopped 로 둔다. 그 뒤의 submit 은 not_running 으로 거부된다 (v5 §7-4 ⑧).
 */
function kice_apiStop_(by) {
  return kice_withLock_(function () {
    var q = kice_readQueRaw_();
    var auto = kice_loadAuto_();
    var mine = kice_paired_(auto, q);
    var done = null;
    if (mine) {
      try { done = kice_countDone_(SpreadsheetApp.getActive().getSheetByName(KICE_QUE_SHEET), Array.isArray(q.rows) ? q.rows : []); } catch (_) {}
      kice_clearQue_();
    }
    var from = kice_markStopped_(auto, by || '원격 중지', done);
    return { ok: true, stoppedFrom: from, cleared: mine, done: done };
  });
}


/* =========================================================
 * 8. 메뉴 '📋 상태'
 * ========================================================= */

function kice_showStatus() {
  var ui = SpreadsheetApp.getUi();
  var st = kice_apiStatus_();
  var a = st.state;
  if (!a) { ui.alert('KICE 윤문 상태', '실행 이력이 없습니다.', ui.ButtonSet.OK); return; }
  var lines = [
    'runId: ' + a.runId + (a.menu ? '  (메뉴 배치)' : ''),
    'stage: ' + a.stage + (st.derived ? '  (파생: ' + st.derived.stage + ' by ' + st.derived.by + ')' : ''),
    '진행: ' + st.progress.idx + ' / ' + st.progress.total +
      (st.progress.inflight ? '  (엔진 처리 중 ' + st.progress.inflight + '행, 첫 행 ' + st.progress.currentRow + ')' : ''),
    '마지막 heartbeat: ' + (st.progress.lastHeartbeat ? new Date(st.progress.lastHeartbeat).toLocaleString() : '-'),
    '키워드: ' + ((a.keywords || []).join(' | ') || '-'),
    '시작: ' + (a.startedAt || '-')
  ];
  if (a.src) lines.push('대상: 일치 ' + a.src.matched + ' / 빈 B ' + a.src.emptyB + ' / 기윤문 제외 ' + a.src.alreadyDone + (a.src.countMismatch ? '  ⚠️ 개수 불일치(expect ' + a.src.expect + ')' : ''));
  if (a.residue) lines.push('잔재 이관: ' + a.residue.archived + '행' + (a.residue.dup ? ' / 중복 ' + a.residue.dup : '') + (a.residue.dropped ? ' / 결과 없음 ' + a.residue.dropped : ''));
  if (a.summary) {
    var s = a.summary;
    lines.push('요약: ' + s.rows + '행 — Claude edit ' + s.claude.edit + ' · 합의 ' + s.gpt.edit + ' · Gemini ' + s.gemini.edit +
               ' / 실패 ' + (s.claude.fail + s.gemini.fail) + ' / 범위 밖 ' + (s.outOfScope || 0) + ' / 포기 ' + (s.poisonCount || 0));
  }
  if (a.archive && a.archive.start) lines.push('아카이브: ' + a.archive.start + '~' + a.archive.end + '행');
  if (a.error) lines.push('오류: ' + a.error);
  if (a.stoppedFrom) lines.push('중지: from ' + a.stoppedFrom + (a.by ? ' by ' + a.by : ''));
  if (a.stage === 'rewrite' && st.progress.lastHeartbeat && (Date.now() - st.progress.lastHeartbeat) > 5 * 60 * 1000) {
    lines.push('', '⚠️ 엔진 신호가 5분 넘게 없습니다. 맥미니의 kice_engine 이 떠 있는지 확인해 주세요.');
  }
  ui.alert('KICE 윤문 상태', lines.join('\n'), ui.ButtonSet.OK);
}

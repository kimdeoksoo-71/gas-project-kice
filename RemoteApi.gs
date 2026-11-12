/**
 * RemoteApi.gs — 헤드리스 원격 제어 API (kice-2.0.0 · 재구축 계획서 v5 §7)
 * 프로젝트: gas-project-kice — 구글 시트 「KICE 윤문」(Drive KICE윤문/)에 바인딩
 *
 * 부르는 쪽
 *   러너(audit_runner)      : ping · start · status · stop · tick           (GET, 1.0.0 때와 같은 계약)
 *   엔진(kice_engine)       : claim · heartbeat · submit(POST) · finish · corpus · prompts · image
 *   코퍼스 빌더(kice_engine): dsread · corpus_put(POST)
 *
 * ── 배포 ─────────────────────────────────────────────────────────────
 *   웹앱 / 실행: 나 / 액세스: 모든 사용자
 *   ⚠️ 코드를 고친 뒤에는 **배포 관리 → 연필(수정) → 버전: 새 버전 → 배포**. "새 배포"는 URL 이 바뀐다.
 *
 * ── 인증 ─────────────────────────────────────────────────────────────
 *   스크립트 속성 REMOTE_TOKEN 과 `token` 비교(GET 은 쿼리, POST 는 JSON 본문). 불일치면 빈 200.
 *
 * ── POST ─────────────────────────────────────────────────────────────
 *   본문은 JSON 하나. {token, cmd, …}. 쿼리 파라미터와 합쳐 GET 과 같은 디스패처로 간다(본문이 우선).
 *   웹앱 POST 는 302 를 돌려주므로 호출하는 쪽이 Location 을 GET 으로 따라가야 응답을 받는다.
 *
 * ── 스크립트 속성 ────────────────────────────────────────────────────
 *   REMOTE_TOKEN      필수
 *   LATEX_SS_ID       사설 런의 원본 「Latex 변환」 — start 가 Data_DS 를 읽는다
 *   LATEX_KICE_SS_ID  사본 「KICE_latex」 — dsread src=latex
 *   LEGACY_SS_ID      「기출과 윤문(2011~)」 — dsread src=legacy
 *   QV_LAYOUT_SWITCH_ROW_2  아카이브 뷰어의 두 번째 경계 행(전환 때 기록)
 *   GITHUB_TOKEN · GITHUB_REPO  프롬프트 push
 */

const RAPI = {
  VERSION:    'kice-2.0.0',
  PROJECT:    'kice',
  TOKEN_PROP: 'REMOTE_TOKEN',
  LOG_TAIL:   20,
  LOG_SHEET:  'KICE_Log',
};

function doGet(e) {
  const p = (e && e.parameter) || {};
  return rapi_handle_(p);
}

function doPost(e) {
  let p = {};
  const q = (e && e.parameter) || {};
  Object.keys(q).forEach(k => { p[k] = q[k]; });
  try {
    const body = e && e.postData && e.postData.contents ? JSON.parse(e.postData.contents) : {};
    if (body && typeof body === 'object') Object.keys(body).forEach(k => { p[k] = body[k]; });
  } catch (err) {
    if (!rapi_auth_(p.token)) return ContentService.createTextOutput('');
    return rapi_json_({ ok: false, reason: 'bad_json: ' + String((err && err.message) || err) });
  }
  return rapi_handle_(p);
}

function rapi_handle_(p) {
  if (!rapi_auth_(String(p.token || ''))) return ContentService.createTextOutput('');   // 빈 200
  try {
    return rapi_json_(rapi_dispatch_(p));
  } catch (err) {
    return rapi_json_({ ok: false, reason: String((err && err.message) || err) });
  }
}

function rapi_dispatch_(p) {
  switch (p.cmd) {
    case 'ping':       return { ok: true, project: RAPI.PROJECT, version: RAPI.VERSION, at: new Date().toISOString() };
    // 러너
    case 'start':      return kice_apiStart_(p);
    case 'tick':       return kice_apiTick_();
    case 'status':     return kice_apiStatus_();
    case 'stop':       return kice_apiStop_('원격 중지 (RemoteApi)');
    // 엔진
    case 'claim':      return kice_apiClaim_(p);
    case 'heartbeat':  return kice_apiHeartbeat_(p);
    case 'submit':     return kice_apiSubmit_(p);
    case 'finish':     return kice_apiFinish_(p);
    case 'corpus':     return kdata_apiCorpus_(p);
    case 'prompts':    return kdata_apiPrompts_();
    case 'image':      return kdata_apiImage_(p);
    // 코퍼스 빌더
    case 'dsread':     return kdata_apiDsread_(p);
    case 'corpus_put': return kdata_apiCorpusPut_(p);
    default:           return { ok: false, reason: 'unknown cmd: ' + p.cmd };
  }
}

/* =================================================
 * 내부 유틸
 * ================================================= */

/** 토큰 비교. 상수시간. */
function rapi_auth_(given) {
  const want = PropertiesService.getScriptProperties().getProperty(RAPI.TOKEN_PROP);
  if (!want || !given) return false;
  if (want.length !== given.length) return false;
  let diff = 0;
  for (let i = 0; i < want.length; i++) diff |= want.charCodeAt(i) ^ given.charCodeAt(i);
  return diff === 0;
}

function rapi_json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** KICE_Log 마지막 n행 → [{time, run, stage, message}] */
function rapi_logTail_(n) {
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(RAPI.LOG_SHEET);
    if (!sh) return [];
    const last = sh.getLastRow();
    if (last < 2) return [];
    const from = Math.max(2, last - n + 1);
    return sh.getRange(from, 1, last - from + 1, 4).getValues().map(r => ({
      time: r[0] instanceof Date ? r[0].toISOString() : String(r[0]),
      run: String(r[1] || ''), stage: String(r[2] || ''), message: String(r[3] || ''),
    }));
  } catch (_) { return []; }
}

/**
 * 설정 확인 (편집기에서 1회 실행). **토큰을 만들지도, 로그에 찍지도 않는다.**
 * 이 함수를 실행하면 이 프로젝트가 쓰는 권한(스프레드시트·Drive·외부 요청) 승인 창이 뜬다(처음 1회).
 * 속성 값은 프로젝트 설정 → 스크립트 속성에 직접 넣는다.
 */
function rapi_setupCheck() {
  const props = PropertiesService.getScriptProperties();
  const tok = props.getProperty(RAPI.TOKEN_PROP);
  if (!tok) console.log('REMOTE_TOKEN 미설정 — 스크립트 속성에 추가하세요.');
  else console.log('REMOTE_TOKEN 설정됨: 길이 %s, 형식 %s', tok.length, (tok.length >= 32 && /^[A-Za-z0-9]+$/.test(tok)) ? '정상' : '⚠️ 비정상(공백·줄바꿈 섞임?)');

  [['LATEX_SS_ID', 'Data_DS'], ['LATEX_KICE_SS_ID', 'Data_DS'], ['LEGACY_SS_ID', 'Data_Latex']].forEach(pair => {
    const id = (props.getProperty(pair[0]) || '').trim();
    if (!id) { console.log('%s 미설정', pair[0]); return; }
    try {
      const ss = SpreadsheetApp.openById(id);
      const sh = ss.getSheetByName(pair[1]);
      console.log('%s 확인: 「%s」 %s %s (마지막 행 %s)', pair[0], ss.getName(), pair[1], sh ? '있음' : '⚠️ 없음', sh ? sh.getLastRow() : '-');
    } catch (e) {
      console.log('%s 열기 실패: %s', pair[0], (e && e.message) || e);
    }
  });

  const ss = SpreadsheetApp.getActive();
  ['Que', '아카이브', 'prompt', 'Rule', 'KICE_Log', 'Corpus', 'AnswerKey', 'ExprDict', 'QueCtx', 'Opinions', 'Feedback', 'EvalSet'].forEach(n => {
    if (!ss.getSheetByName(n)) console.log('⚠️ 시트 없음: %s', n);
  });
  try { DriveApp.getRootFolder().getName(); console.log('Drive 권한 확인'); } catch (e) { console.log('Drive 권한 없음: %s', (e && e.message) || e); }
  console.log('버전 %s · 아카이브 마지막 행 %s', RAPI.VERSION, ss.getSheetByName('아카이브') ? ss.getSheetByName('아카이브').getLastRow() : '-');
}

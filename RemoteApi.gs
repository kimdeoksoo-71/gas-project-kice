/**
 * RemoteApi.gs — 헤드리스 원격 제어 API
 * 프로젝트: gas-project-kice (기출과 윤문 2011~)
 *
 * latex-convert / audition 의 RemoteApi.gs 와 같은 꼴이다. 맥미니 `audit_runner` 가 이 웹앱을
 * 불러 KICE 윤문 배치를 시작(start)·구동(tick)·조회(status)·중지(stop)한다. 본체는 kiceAuto.gs.
 *
 * ── 배포 ─────────────────────────────────────────────────────────────
 *   웹앱 / 실행: 나(kimdeoksoo@gmail.com) / 액세스: 모든 사용자
 *
 * ⚠️⚠️ 코드를 고친 뒤에는 반드시 **배포 관리 → 연필(수정) → 버전: 새 버전 → 배포**.
 *      "새 배포"는 URL이 바뀌고 옛 배포가 살아남아, 러너가 에러 없이 옛 코드를
 *      계속 부르게 된다. `VERSION`을 러너가 대조해 막는다.
 *
 * ── 인증 ─────────────────────────────────────────────────────────────
 *   스크립트 속성 `REMOTE_TOKEN`(32자 이상 난수)과 쿼리 파라미터 `token` 비교.
 *   불일치 시 **빈 200 응답** — 엔드포인트의 존재 자체를 드러내지 않는다 (C7).
 *   토큰은 맥미니에서 만들어 `~/audit_runner/secrets/remote_tokens.json` 의 "kice" 에 둔다 (프로젝트별 별도).
 *
 * ── 명령 ─────────────────────────────────────────────────────────────
 *   ping   → { ok, project, version, at }
 *   start  → kice_apiStart_(p)   : keywords(필수) · runId(필수) · expect(경고용) · limit(시험용)
 *   tick   → kice_apiTick_()     : batch_continueQueAuto() 1회 (예산 KICE_TICK_BUDGET_MS)
 *   status → kice_apiStatus_()   : 읽기만 { ok, version, state, derived, progress, logTail, hasTick, at }
 *   stop   → kice_apiStop_()     : QUE 삭제·트리거 삭제·AUTO stopped
 */

const RAPI = {
  VERSION:    'kice-1.0.0',
  PROJECT:    'kice',
  TOKEN_PROP: 'REMOTE_TOKEN',
  LOG_TAIL:   20,
  LOG_SHEET:  'KICE_Log',
};

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (!rapi_auth_(p.token)) return ContentService.createTextOutput('');   // C7: 빈 200

  try {
    switch (p.cmd) {
      case 'ping':   return rapi_json_({ ok: true, project: RAPI.PROJECT, version: RAPI.VERSION, at: new Date().toISOString() });
      case 'start':  return rapi_json_(kice_apiStart_(p));
      case 'tick':   return rapi_json_(kice_apiTick_());
      case 'status': return rapi_json_(kice_apiStatus_());
      case 'stop':   return rapi_json_(kice_apiStop_('원격 중지 (RemoteApi)'));
      default:       return rapi_json_({ ok: false, reason: 'unknown cmd: ' + p.cmd });
    }
  } catch (err) {
    return rapi_json_({ ok: false, reason: String((err && err.message) || err) });
  }
}

/* =================================================
 * 내부 유틸 (latex-convert 의 RemoteApi.gs 와 동일)
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
 * `REMOTE_TOKEN` 설정 확인 (편집기에서 1회 실행). **토큰을 만들지도, 로그에 찍지도 않는다.**
 * 토큰은 맥미니에서 암호용 난수로 생성한다(`~/audit_runner/secrets/remote_tokens.json`의 "kice").
 *   → 프로젝트 설정 → 스크립트 속성 → `REMOTE_TOKEN`에 그 값을 붙여 넣은 뒤 이 함수로 확인.
 * 함께 `LATEX_SS_ID`(Latex 변환 스프레드시트 ID)도 확인한다. 이 함수를 실행하면 SpreadsheetApp·
 * ScriptApp 등 이 프로젝트가 쓰는 권한 승인 창이 뜬다(처음 1회).
 */
function rapi_setupToken() {
  const props = PropertiesService.getScriptProperties();
  const tok = props.getProperty(RAPI.TOKEN_PROP);
  if (!tok) { console.log('REMOTE_TOKEN 미설정 — 스크립트 속성에 추가하세요.'); }
  else {
    const ok = tok.length >= 32 && /^[A-Za-z0-9]+$/.test(tok);
    console.log('REMOTE_TOKEN 설정됨: 길이 %s, 형식 %s', tok.length, ok ? '정상' : '⚠️ 비정상(공백·줄바꿈 섞임?)');
  }
  const ssId = (props.getProperty('LATEX_SS_ID') || '').trim();
  if (!ssId) { console.log('LATEX_SS_ID 미설정 — Latex 변환 스프레드시트 ID를 스크립트 속성에 추가하세요.'); return; }
  try {
    const sh = SpreadsheetApp.openById(ssId).getSheetByName('Data_DS');
    console.log('LATEX_SS_ID 확인: Data_DS %s (마지막 행 %s)', sh ? '있음' : '⚠️ 없음', sh ? sh.getLastRow() : '-');
  } catch (e) {
    console.log('LATEX_SS_ID 열기 실패: %s', (e && e.message) || e);
  }
}

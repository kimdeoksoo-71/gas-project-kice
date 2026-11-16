/**
 * RefViewer.gs — 기출 사례 보기 (kice-2.0.3 · v5 §32-7)
 *
 * 의견 블록의 `[[사례]]` 줄(엔진이 적는다 — `[[빈도]]` 가 센 바로 그 기출 문항ID 들)을 사이드바·아카이브 뷰어가 「사례 보기」 링크로 바꾼다.
 * 링크는 웹앱의 새 탭 화면(?view=refs)을 연다. 문항 원본 이미지(Corpus 「통이미지 링크」)를 최근 것부터 한 장씩, 화살표 키로 넘긴다.
 *
 * 접근: 웹앱은 익명 접근이라 표(ticket)가 있어야 화면과 그림을 준다. 표는 시트에서 사이드바·뷰어를 연 사람에게만 주고 12시간마다 바뀐다
 *       (직전 표도 받는다 — 열어 둔 탭이 끊기지 않게). REMOTE_TOKEN 과는 별개다. 그림은 Corpus 에 적힌 파일만 준다.
 */
var KV = { PROP: 'KV_TICKET', URL_PROP: 'KV_WEBAPP_URL', TTL_MS: 12 * 3600 * 1000, MAX_IDS: 400 };

function kv_state_() {
  try { return JSON.parse(PropertiesService.getScriptProperties().getProperty(KV.PROP) || '{}') || {}; } catch (e) { return {}; }
}

function kv_ticket_() {
  var st = kv_state_(), now = Date.now();
  if (!st.cur || now - Number(st.at || 0) > KV.TTL_MS) {
    st = { cur: Utilities.getUuid().replace(/-/g, ''), prev: (st.cur && now - Number(st.at || 0) < 2 * KV.TTL_MS) ? st.cur : '', at: now };
    PropertiesService.getScriptProperties().setProperty(KV.PROP, JSON.stringify(st));
  }
  return st.cur;
}

function kv_ticketOk_(t) {
  t = String(t || '');
  if (t.length < 16) return false;
  var st = kv_state_();
  if (Date.now() - Number(st.at || 0) > 2 * KV.TTL_MS) return false;
  return t === st.cur || t === st.prev;
}

/** 사이드바·아카이브 뷰어에 주는 링크의 바탕 {url, t}. 웹앱 주소를 알 수 없으면 url 이 빈 글자다(뷰어는 링크 없이 문항 수만 보인다). */
function kv_base_() {
  var url = '';
  try { url = String(PropertiesService.getScriptProperties().getProperty(KV.URL_PROP) || ''); } catch (e) {}
  if (!url) { try { url = String(ScriptApp.getService().getUrl() || ''); } catch (e) { url = ''; } }
  if (!url) return { url: '', t: '' };
  return { url: url, t: kv_ticket_() };
}

/** `2025-09-C.11.15~2025-06-P.27` → ['2025-09-C-11', '2025-09-C-15', '2025-06-P-27'] (뷰어가 주소를 줄이려고 이렇게 묶는다) */
function kv_expand_(s) {
  var out = [];
  String(s || '').split('~').forEach(function (g) {
    var a = g.split('.');
    if (a.length < 2 || !/^[\w-]+$/.test(a[0])) return;
    for (var i = 1; i < a.length; i++) if (/^\w+$/.test(a[i])) out.push(a[0] + '-' + a[i]);
  });
  return out.slice(0, KV.MAX_IDS);
}

function kv_page_(p) {
  if (!kv_ticketOk_(p.t)) {
    return HtmlService.createHtmlOutput('<p style="font:15px sans-serif;margin:40px">링크가 만료되었습니다. 시트에서 사이드바(또는 아카이브 뷰어)를 다시 연 뒤 「사례 보기」를 눌러 주세요.</p>').setTitle('기출 사례');
  }
  var tpl = HtmlService.createTemplateFromFile('RefViewerPage');
  tpl.boot = JSON.stringify({ t: String(p.t), ids: kv_expand_(p.ids), q: String(p.q || '').slice(0, 200) }).replace(/</g, '\\u003c');
  return tpl.evaluate().setTitle('기출 사례');
}

function kv_corpus_() {
  var sh = SpreadsheetApp.getActive().getSheetByName('Corpus');
  if (!sh || sh.getLastRow() < 2) return { H: {}, rows: [] };
  var v = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  var H = {};
  v[0].forEach(function (h, i) { H[String(h).trim()] = i; });
  return { H: H, rows: v.slice(1) };
}

/** 화면이 부른다: 문항ID 목록 → [{id, label, has}] (요청한 순서 그대로). */
function kv_refMeta(t, ids) {
  if (!kv_ticketOk_(t)) return { ok: false, reason: 'expired' };
  var c = kv_corpus_(), H = c.H, by = {};
  if (H['문항ID'] == null) return { ok: false, reason: 'Corpus 시트에 문항ID 열이 없습니다.' };
  c.rows.forEach(function (r) { by[String(r[H['문항ID']]).trim()] = r; });
  var col = function (r, k) { return H[k] == null ? '' : String(r[H[k]] == null ? '' : r[H[k]]).trim(); };
  var items = (ids || []).slice(0, KV.MAX_IDS).map(function (id) {
    var r = by[String(id).trim()];
    if (!r) return { id: String(id), label: '(Corpus 에 없음)', has: false };
    var year = col(r, '출제연도').replace(/\.0$/, ''), no = col(r, '번호').replace(/\.0$/, '');
    return { id: String(id), label: [year ? year + '년 출제' : '', col(r, '시험'), col(r, '과목'), no ? no + '번' : ''].filter(String).join(' · '),
             has: !!kdata_fileId_(col(r, '통이미지 링크')) };
  });
  return { ok: true, items: items };
}

/** 화면이 부른다: 문항ID → 원본 이미지 {ok, mime, base64}. Corpus 에 적힌 파일만 준다. */
function kv_refImage(t, id) {
  if (!kv_ticketOk_(t)) return { ok: false, reason: 'expired' };
  var c = kv_corpus_(), H = c.H;
  if (H['문항ID'] == null || H['통이미지 링크'] == null) return { ok: false, reason: 'Corpus 시트에 문항ID·통이미지 링크 열이 없습니다.' };
  id = String(id || '').trim();
  for (var i = 0; i < c.rows.length; i++) {
    if (String(c.rows[i][H['문항ID']]).trim() !== id) continue;
    var out = kdata_apiImage_({ url: String(c.rows[i][H['통이미지 링크']] || '') });
    return out.ok ? { ok: true, id: id, mime: out.mime, base64: out.base64 } : { ok: false, id: id, reason: out.reason };
  }
  return { ok: false, id: id, reason: 'Corpus 에 없는 문항입니다.' };
}

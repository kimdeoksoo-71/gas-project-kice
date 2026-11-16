/**
 * CorpusSearch.gs — 기출 표현 검색 (v5 §32-9)
 *
 * 사이드바 뷰어 맨 위의 검색 칸에 표현을 적으면 Corpus 의 LaTeX 에서 찾아 문항 수·연도별 문항 수와 「기출 N문항 보기」 링크를 보여 준다.
 * 링크는 사례 보기 화면(RefViewer)이다 — 원본 이미지를 최근 것부터 화살표 키로.
 * 일치: 글자 그대로, 다만 공백(띄어쓰기·줄바꿈)은 무시한다 — 변환된 LaTeX 는 `$x$ 축` 처럼 띄어쓰기가 원문과 다르다.
 */
var KS = { MAX_ITEMS: 400, CTX: 45, LINK_MAX: 300 };

/** 사례 보기 화면(RefViewer) 주소 — 문항ID 를 시험별로 묶어 줄인다. 주소를 알 수 없으면 빈 글자. */
function ks_refsUrl_(ids, q) {
  var base = { url: '', t: '' };
  try { base = kv_base_(); } catch (e) {}
  if (!base.url || !ids.length) return '';
  var groups = [], at = {};
  ids.forEach(function (id) { var k = id.lastIndexOf('-'), p = id.slice(0, k); if (!(p in at)) { at[p] = groups.length; groups.push([p]); } groups[at[p]].push(id.slice(k + 1)); });
  return base.url + (base.url.indexOf('?') < 0 ? '?' : '&') + 'view=refs&t=' + encodeURIComponent(base.t)
    + '&ids=' + groups.map(function (g) { return g.join('.'); }).join('~') + '&q=' + encodeURIComponent(String(q).slice(0, 80));
}

/**
 * 사이드바의 검색 칸이 부른다. → {ok, total, hits, years, url, shown} — url 은 사례 보기 화면(최근 LINK_MAX 문항까지).
 */
function ks_searchLink(q, inScope) {
  var r = ks_search(q, { inScope: !!inScope });
  if (!r.ok) return r;
  var ids = r.items.slice(0, KS.LINK_MAX).map(function (x) { return x.id; });
  return { ok: true, q: r.q, total: r.total, hits: r.hits, years: r.years, shown: ids.length, url: ks_refsUrl_(ids, q) };
}

/** 공백을 뺀 글과, 그 글의 각 글자가 원문의 몇 번째 글자인지. */
function ks_squeeze_(text) {
  var out = [], at = [];
  for (var i = 0; i < text.length; i++) {
    var ch = text.charAt(i);
    if (/\s/.test(ch)) continue;
    out.push(ch); at.push(i);
  }
  return { s: out.join(''), at: at };
}

/**
 * q: 찾을 표현 · opts.inScope: 2028 범위 안 문항만.
 * 응답 {ok, q, total, hits, years:[[연도, 문항 수]], items:[{id, label, scope, n, pre, hit, post, link}], truncated}
 */
function ks_search(q, opts) {
  opts = opts || {};
  var nq = String(q || '').replace(/\s+/g, '');
  if (nq.length < 2) return { ok: false, reason: '두 글자 이상 입력해 주세요.' };
  var c = kv_corpus_(), H = c.H;
  if (H['문항ID'] == null || H['LaTeX'] == null) return { ok: false, reason: 'Corpus 시트에 문항ID·LaTeX 열이 없습니다.' };
  var col = function (r, k) { return H[k] == null ? '' : String(r[H[k]] == null ? '' : r[H[k]]).trim(); };
  var items = [], years = {}, hits = 0;
  c.rows.forEach(function (r) {
    var id = col(r, '문항ID'), text = String(r[H['LaTeX']] == null ? '' : r[H['LaTeX']]);
    if (!id || !text) return;
    var scope = col(r, '2028범위');
    if (opts.inScope && scope !== '범위 안') return;
    var z = ks_squeeze_(text), first = z.s.indexOf(nq);
    if (first < 0) return;
    var n = 0, k = first;
    while (k >= 0) { n++; k = z.s.indexOf(nq, k + nq.length); }
    var a = z.at[first], b = z.at[first + nq.length - 1] + 1;
    var flat = function (s) { return s.replace(/\s+/g, ' '); };
    var year = col(r, '출제연도').replace(/\.0$/, ''), no = col(r, '번호').replace(/\.0$/, '');
    years[year] = (years[year] || 0) + 1; hits += n;
    items.push({ id: id, scope: scope, n: n, link: col(r, '통이미지 링크'),
                 label: [year ? year + '년 출제' : '', col(r, '시험'), col(r, '과목'), no ? no + '번' : ''].filter(String).join(' · '),
                 pre: (a > KS.CTX ? '…' : '') + flat(text.slice(Math.max(0, a - KS.CTX), a)), hit: flat(text.slice(a, b)),
                 post: flat(text.slice(b, b + KS.CTX)) + (b + KS.CTX < text.length ? '…' : '') });
  });
  items.sort(function (x, y) { var p = x.id.slice(0, 7), q2 = y.id.slice(0, 7); return p < q2 ? 1 : p > q2 ? -1 : (x.id < y.id ? -1 : x.id > y.id ? 1 : 0); });   // 최근 시험부터
  return { ok: true, q: String(q), total: items.length, hits: hits, truncated: items.length > KS.MAX_ITEMS,
           years: Object.keys(years).sort().map(function (y) { return [y, years[y]]; }), items: items.slice(0, KS.MAX_ITEMS) };
}

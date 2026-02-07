/*************************************************
 * Token_Stat 생성 v4 (TEXT / MATH + PATTERN / FORM)
 * + PATTERN 자동 n-gram(PATG_) 추가
 * + FORM 확장(형태/조합/구조)
 * + 숫자 토큰 제거(특히 MATH)
 * + df 필터 + 타입별 cap(총 5000 이하)
 * + ✅ (요구) MATH df=1 토큰은 생성/출력하지 않음
 *
 * source: Data_Latex!C (latex)
 *
 * Token_Stat 구조:
 * A token
 * B token_type: TEXT | MATH | PATTERN | FORM
 * C df
 * D tf
 * E weight = idf * typeBoost
 * F example
 *************************************************/

const TS_BOOST = {
  TEXT: 1.0,
  MATH: 1.0,
  PATTERN: 2.5,
  FORM: 1.7
};

const TS_CAP = {
  TEXT: 2500,
  MATH: 1500,
  PATTERN: 800,
  FORM: 200
};

// PATG(df) 필터
const TS_PATG_MIN_DF = 3;
const TS_PATG_MAX_DF_RATIO = 0.15;

// ✅ 추가: MATH df=1 제거
const TS_MATH_MIN_DF = 2;

// (옵션) FORM도 df=1 제거하고 싶으면 2로 바꿔
const TS_FORM_MIN_DF = 1;

function buildTokenStatFromDataLatex() {
  const ss = SpreadsheetApp.getActive();
  const dataSheet = ss.getSheetByName('Data_Latex');
  const statSheet = ss.getSheetByName('Token_Stat');

  if (!dataSheet || !statSheet) {
    ss.toast('❌ Data_Latex 또는 Token_Stat 시트가 없습니다.', '오류', 5);
    return;
  }

  const lastRow = dataSheet.getLastRow();
  if (lastRow < 2) {
    ss.toast('Data_Latex에 데이터가 없습니다.', '오류', 5);
    return;
  }

  const data = dataSheet.getRange(2, 3, lastRow - 1, 1).getValues(); // C열
  const N = data.length;

  /** token -> { type, df, tf, example } */
  const map = new Map();

  data.forEach(row => {
    const latex = String(row[0] || '');
    if (!latex) return;

    const seenInThisProblem = new Set();

    // 1) 수식 추출 ($...$)
    const mathRawList = [];
    const mathMatches = [...latex.matchAll(/\$(.+?)\$/g)];
    mathMatches.forEach(m => {
      const raw = String(m[1] ?? '');
      mathRawList.push(raw);

      tokenizeMath_(raw).forEach(tok => {
        registerToken_(map, tok, 'MATH', latex, seenInThisProblem);
      });
    });

    // 2) 텍스트 부분
    const textOnly = latex.replace(/\$(.+?)\$/g, ' ');

    tokenizeText_(textOnly).forEach(tok => {
      registerToken_(map, tok, 'TEXT', latex, seenInThisProblem);
    });

    // 3) PATTERN(핵심 regex)
    extractPatternTokens_(textOnly).forEach(tok => {
      registerToken_(map, tok, 'PATTERN', latex, seenInThisProblem);
    });

    // 4) PATTERN n-gram(PATG_)
    extractPatternNgrams_(textOnly).forEach(tok => {
      registerToken_(map, tok, 'PATTERN', latex, seenInThisProblem);
    });

    // 5) FORM(수식 구조)
    extractFormTokens_(mathRawList).forEach(tok => {
      registerToken_(map, tok, 'FORM', latex, seenInThisProblem);
    });
  });

  // 6) rows 생성(+ weight)
  let rows = [];
  map.forEach((v, token) => {
    const df = Math.max(1, v.df);
    const idf = Math.log(N / df);
    const boost = TS_BOOST[v.type] || 1.0;
    const weight = idf * boost;
    rows.push([token, v.type, v.df, v.tf, weight, v.example]);
  });

  // 7) df 필터 + 기타 정리
  rows = rows.filter(r => {
    const token = String(r[0]);
    const type = String(r[1]);
    const df = Number(r[2]);

    if (!isFinite(df)) return false;

    // TEXT 1글자 제거
    if (type === 'TEXT' && token.length <= 1) return false;

    // MATH 순수 숫자 제거(방어)
    if (type === 'MATH' && /^\d+$/.test(token)) return false;

    // ✅ 핵심: MATH df=1 제거
    if (type === 'MATH' && df < TS_MATH_MIN_DF) return false;

    // (옵션) FORM df=1도 제거하고 싶으면 TS_FORM_MIN_DF=2로
    if (type === 'FORM' && df < TS_FORM_MIN_DF) return false;

    // PATG_*에만 df 필터 적용
    if (type === 'PATTERN' && token.startsWith('PATG_')) {
      if (df < TS_PATG_MIN_DF) return false;
      if (df > N * TS_PATG_MAX_DF_RATIO) return false;
    }

    return true;
  });

  // 8) 타입별 cap
  rows = capRowsByType_(rows);

  // 9) 출력
  statSheet.clearContents();
  statSheet.getRange(1, 1, 1, 6).setValues([
    ['token', 'token_type', 'df', 'tf', 'weight', 'example']
  ]);

  rows.sort((a, b) => (b[4] - a[4]) || String(a[0]).localeCompare(String(b[0])));

  if (rows.length > 0) {
    statSheet.getRange(2, 1, rows.length, 6).setValues(rows);
  }

  ss.toast(`✅ Token_Stat v4 완료 (총 ${rows.length}개, N=${N})`, '완료', 4);
}

/*************************************************
 * TEXT 토큰화 (숫자 제거)
 *************************************************/
function tokenizeText_(text) {
  return String(text || '')
    .toLowerCase()
    .replace(/[0-9]/g, ' ')
    .replace(/[^\p{L}]/gu, ' ')
    .split(/\s+/)
    .filter(t => t.length >= 1);
}

/*************************************************
 * MATH 토큰화
 * - 순수 숫자는 제거
 *************************************************/
function tokenizeMath_(math) {
  return String(math || '')
    .replace(/\\[a-zA-Z]+/g, m => ` ${m} `)
    .replace(/[^a-zA-Z0-9\\]/g, ' $& ')
    .split(/\s+/)
    .filter(t => t.length >= 1)
    .filter(t => !/^\d+$/.test(t));
}

/*************************************************
 * PATTERN 토큰(6개 유지)
 *************************************************/
function extractPatternTokens_(textOnly) {
  const s = String(textOnly || '').toLowerCase();
  const out = new Set();

  if (/성립/.test(s)) out.add('PAT_성립');
  if (/다음(은|을)/.test(s)) out.add('PAT_다음');
  if (/모든\s*자연수/.test(s) || /자연수\s*n/.test(s)) out.add('PAT_모든자연수');
  if (/실수\s*전체/.test(s) || /실수\s*전체의\s*집합/.test(s)) out.add('PAT_실수전체');
  if (/를\s*만족/.test(s)) out.add('PAT_를만족');
  if (/라\s*하자/.test(s)) out.add('PAT_라하자');

  return Array.from(out);
}

/*************************************************
 * PATTERN n-gram(PATG_) 생성
 *************************************************/
function extractPatternNgrams_(textOnly) {
  const toks = tokenizeText_(textOnly).filter(t => t.length >= 2);

  const STOP = new Set([
    '다음','각','모든','서로','대하여','하여','한다','이다','인','에서','의','을','를','은','는','이','가',
    '그리고','또는','혹은','또','또한','때','경우','조건','만족','성립','하여라','구하여라','구하라',
    '값','최댓값','최소값','최솟값','정답','보기','중','중에서',
    '실수','정수','자연수','양수','음수','유리수','무리수',
    '함수','수열','항','점','선분','직선','원','삼각형','사각형','다각형','좌표','평면',
    '구간','범위','일때','이라','이라면','이고','이며','라고'
  ]);

  const clean = toks.filter(t => !STOP.has(t));
  if (clean.length < 2) return [];

  const out = new Set();
  const addN = (n) => {
    for (let i = 0; i + n <= clean.length; i++) {
      out.add('PATG_' + clean.slice(i, i + n).join('_'));
    }
  };

  addN(2); addN(3); addN(4);
  return Array.from(out);
}

/*************************************************
 * FORM 토큰(확장)
 *************************************************/
function extractFormTokens_(mathRawList) {
  const out = new Set();

  (mathRawList || []).forEach(raw => {
    const s = String(raw || '');

    // 기존 7개
    if (s.includes('\\int')) out.add('FORM_int');
    if (s.includes('\\sum')) out.add('FORM_sum');
    if (s.includes('\\sin') || s.includes('\\cos') || s.includes('\\tan')) out.add('FORM_trig');
    if (s.includes('\\log') || s.includes('\\ln')) out.add('FORM_log');
    if (s.includes('\\lim')) out.add('FORM_lim');
    if (s.includes('\\sqrt')) out.add('FORM_sqrt');
    if (s.includes('\\frac')) out.add('FORM_frac');

    // 형태/구조
    if (/_/.test(s)) out.add('FORM_subscript');
    if (/\^/.test(s)) out.add('FORM_superscript');

    if (s.includes('\\left|') || s.includes('\\right|') || s.includes('|')) out.add('FORM_abs');
    if (s.includes('\\begin{cases}') || s.includes('\\cases')) out.add('FORM_piecewise');
    if (/[()[\]]/.test(s)) out.add('FORM_interval');

    const eqCount = (s.match(/=/g) || []).length;
    if (eqCount >= 2) out.add('FORM_eq_chain');

    const hasFrac = s.includes('\\frac');
    const hasSqrt = s.includes('\\sqrt');
    const hasLog  = s.includes('\\log') || s.includes('\\ln');
    const hasLim  = s.includes('\\lim');
    const hasInt  = s.includes('\\int');
    const hasSum  = s.includes('\\sum');

    if (hasFrac && hasSqrt) out.add('FORM_frac+sqrt');
    if (hasLog && hasLim) out.add('FORM_log+lim');
    if (hasInt && hasFrac) out.add('FORM_int+frac');
    if (hasSum && hasFrac) out.add('FORM_sum+frac');
  });

  return Array.from(out);
}

/*************************************************
 * 타입별 cap 적용
 *************************************************/
function capRowsByType_(rows) {
  const by = { TEXT: [], MATH: [], PATTERN: [], FORM: [] };

  rows.forEach(r => {
    const type = String(r[1]);
    if (!by[type]) return;
    by[type].push(r);
  });

  Object.keys(by).forEach(k => by[k].sort((a, b) => (b[4] - a[4])));

  return [
    ...by.TEXT.slice(0, TS_CAP.TEXT),
    ...by.MATH.slice(0, TS_CAP.MATH),
    ...by.PATTERN.slice(0, TS_CAP.PATTERN),
    ...by.FORM.slice(0, TS_CAP.FORM),
  ];
}

/*************************************************
 * 토큰 등록
 *************************************************/
function registerToken_(map, token, type, example, seenSet) {
  const tok = String(token || '').trim();
  if (!tok) return;

  if (!map.has(tok)) {
    map.set(tok, {
      type,
      df: 0,
      tf: 0,
      example: String(example || '').slice(0, 80)
    });
  }

  const obj = map.get(tok);
  obj.tf++;

  if (!seenSet.has(tok)) {
    obj.df++;
    seenSet.add(tok);
  }
}

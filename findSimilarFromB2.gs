/*****************************************************
 * Find Similar (TF-IDF + Cosine, TEXT/MATH 분리)
 * + Token_Stat의 PATTERN/FORM 적극 반영
 * + (복구) C2 대단원 필터: C2에 값이 있으면 같은 chapter만 검색
 *
 * Input (현재 시트):
 *  - B2: query latex(본문+수식)
 *  - C2: query chapter (대단원)  ✅ 있으면 필터
 *  - E2: topN (1~50, 비어있으면 DEFAULT_TOP_N)
 *  - F2: text weight
 *  - G2: math weight
 *
 * Output 보호:
 *  - A5:E20까지만 사용(21행 이하 보호)
 *****************************************************/

const DEFAULT_TOP_N   = 10;
const DEFAULT_W_TEXT  = 0.7;
const DEFAULT_W_MATH  = 0.3;

const OUTPUT_HEADER_ROW = 5;
const OUTPUT_START_ROW  = 6;
const OUTPUT_LAST_ROW   = 20;   // 21행 이하 보호
const OUTPUT_LATEX_COL  = 4;

function findSimilarFromB2(options) {
  options = options || {};
  const openViewer = (options.openViewer !== false); // 기본 true

  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getActiveSheet();

  const dataSheet = ss.getSheetByName('Data_Latex');
  const statSheet = ss.getSheetByName('Token_Stat');

  if (!dataSheet || !statSheet) {
    ss.toast('Data_Latex 또는 Token_Stat 시트를 찾을 수 없습니다.', '오류', 5);
    return;
  }

  // 1) 입력
  const query = String(sheet.getRange('B2').getDisplayValue() || '').trim();
  if (!query) {
    ss.toast('B2 셀에 비교할 문항을 입력하세요.', '안내', 5);
    return;
  }

  // ✅ (복구) C2 대단원 필터
  const queryChapter = String(sheet.getRange('C2').getDisplayValue() || '').trim();
  const useChapterFilter = !!queryChapter;

  // 출력 가능 최대 개수(6~20행 => 15개)
  const maxOutput = OUTPUT_LAST_ROW - OUTPUT_START_ROW + 1;

  const wantN = FS_clampInt_(sheet.getRange('E2').getValue(), 1, 50, DEFAULT_TOP_N);
  const topNCount = Math.min(wantN, maxOutput);

  const { wText, wMath } = FS_readWeightsFG_(sheet); // F2/G2

  // 2) Token_Stat 로드 (PATTERN->TEXT, FORM->MATH)
  const { idfText, idfMath } = FS_loadIdfMaps_(statSheet);

  // 3) Query 벡터
  const qTokens = FS_extractTokensFromLatex_(query);

  const qTextTokens = FS_filterStopText_(
    (qTokens.textTokens || []).concat(qTokens.patternTokens || [])
  );
  const qMathTokens = FS_filterNoiseMath_(
    (qTokens.mathTokens || []).concat(qTokens.formTokens || [])
  );

  const qTextVec = FS_buildTfidfVector_(qTextTokens, idfText);
  const qMathVec = FS_buildTfidfVectorWithMathExtras_(qTokens.mathRawList, qMathTokens, idfMath);

  const qTextNorm = FS_vecNorm_(qTextVec);
  const qMathNorm = FS_vecNorm_(qMathVec);

  // 4) Data scan
  const data = dataSheet.getDataRange().getValues();
  const scored = [];

  for (let i = 1; i < data.length; i++) { // 2행부터
    const row = data[i];
    const source    = row[8]; // I
    const chapter   = String(row[9] ?? '').trim(); // J
    const driveLink = row[1]; // B
    const latex     = row[2]; // C
    if (!latex) continue;

    // ✅ (복구) 대단원 필터: C2가 비어있지 않으면 같은 chapter만
    if (useChapterFilter && chapter !== queryChapter) continue;

    const dTokens = FS_extractTokensFromLatex_(latex);

    const dTextTokens = FS_filterStopText_(
      (dTokens.textTokens || []).concat(dTokens.patternTokens || [])
    );
    const dMathTokens = FS_filterNoiseMath_(
      (dTokens.mathTokens || []).concat(dTokens.formTokens || [])
    );

    const dTextVec = FS_buildTfidfVector_(dTextTokens, idfText);
    const dMathVec = FS_buildTfidfVectorWithMathExtras_(dTokens.mathRawList, dMathTokens, idfMath);

    const textScore = FS_cosine_(qTextVec, qTextNorm, dTextVec);
    const mathScore = FS_cosine_(qMathVec, qMathNorm, dMathVec);

    const score = wText * textScore + wMath * mathScore;
    if (score > 0) scored.push({ score, source, chapter, driveLink, latex, rowIndex: i + 1 });
  }

  // 5) 정렬 + 상위 N
  scored.sort((a, b) => b.score - a.score || a.rowIndex - b.rowIndex);
  const topN = scored.slice(0, topNCount);

  // source 기준 내림차순(유지)
  topN.sort((a, b) => String(b.source).localeCompare(String(a.source)));

  // 6) 출력(보호 범위만)
  sheet.getRange(OUTPUT_HEADER_ROW, 1, OUTPUT_LAST_ROW - OUTPUT_HEADER_ROW + 1, 5).clearContent();

  sheet.getRange(OUTPUT_HEADER_ROW, 1, 1, 5).setValues([
    ['source', 'chapter', 'drive_link', 'Latex', 'score']
  ]);

  topN.forEach((item, idx) => {
    const r = OUTPUT_START_ROW + idx;
    sheet.getRange(r, 1).setValue(item.source);
    sheet.getRange(r, 2).setValue(item.chapter);
    sheet.getRange(r, 3).setValue(item.driveLink);
    sheet.getRange(r, 4).setValue(item.latex);
    sheet.getRange(r, 5).setValue(item.score);
  });

  ss.toast(
    useChapterFilter
      ? `유사문항 ${topN.length}개 (대단원 필터 ON: "${queryChapter}", TEXT=${wText.toFixed(2)}, MATH=${wMath.toFixed(2)})`
      : `유사문항 ${topN.length}개 (대단원 필터 OFF, TEXT=${wText.toFixed(2)}, MATH=${wMath.toFixed(2)})`,
    '검색',
    3
  );

  if (openViewer) FS_openViewerOnResults_(ss, sheet, topN.length);
}

/* ---------------------------
 * Viewer open
 * --------------------------- */
function FS_openViewerOnResults_(ss, sheet, count) {
  if (!count || count <= 0) {
    ss.toast('표시할 검색 결과가 없습니다.', '안내', 3);
    return;
  }
  const safeCount = Math.min(count, OUTPUT_LAST_ROW - OUTPUT_START_ROW + 1);
  const range = sheet.getRange(OUTPUT_START_ROW, OUTPUT_LATEX_COL, safeCount, 1);
  sheet.setActiveRange(range);

  try {
    if (typeof LV !== 'undefined' && LV && typeof LV.openDialog === 'function') {
      LV.openDialog();
    } else if (typeof lv_openDialog === 'function') {
      lv_openDialog();
    } else {
      ss.toast('LatexViewer(LV) 함수를 찾지 못했습니다. LatexViewer.gs가 있는지 확인해줘.', '오류', 5);
    }
  } catch (err) {
    ss.toast(`Viewer 오픈 실패: ${err && err.message ? err.message : err}`, '오류', 5);
  }
}

/* ---------------------------
 * Weights: F2/G2 (정규화)
 * --------------------------- */
function FS_readWeightsFG_(sheet) {
  let t = Number(sheet.getRange('F2').getValue());
  let m = Number(sheet.getRange('G2').getValue());

  if (!isFinite(t)) t = DEFAULT_W_TEXT;
  if (!isFinite(m)) m = DEFAULT_W_MATH;

  t = Math.max(0, t);
  m = Math.max(0, m);

  const sum = t + m;
  if (sum <= 0) return { wText: DEFAULT_W_TEXT, wMath: DEFAULT_W_MATH };

  return { wText: t / sum, wMath: m / sum };
}

/* ---------------------------
 * Load maps: PATTERN->idfText, FORM->idfMath
 * --------------------------- */
function FS_loadIdfMaps_(statSheet) {
  const last = statSheet.getLastRow();
  const n = Math.max(last - 1, 0);
  const vals = n ? statSheet.getRange(2, 1, n, 6).getValues() : [];

  const idfText = new Map();
  const idfMath = new Map();

  vals.forEach(r => {
    const token = String(r[0] || '').trim();
    const type  = String(r[1] || '').trim();
    const w     = Number(r[4]); // weight (=idf*boost)
    if (!token || !isFinite(w)) return;

    if (type === 'TEXT' || type === 'PATTERN') idfText.set(token, w);
    else if (type === 'MATH' || type === 'FORM') idfMath.set(token, w);
  });

  return { idfText, idfMath };
}

/* ---------------------------
 * Extract tokens (+PATTERN/+FORM)
 * --------------------------- */
function FS_extractTokensFromLatex_(latex) {
  const s = String(latex || '');

  const mathRawList = [];
  const mathTokens = [];
  const mathMatches = [...s.matchAll(/\$(.+?)\$/g)];
  mathMatches.forEach(m => {
    const raw = String(m[1] ?? '');
    mathRawList.push(raw);
    mathTokens.push(...FS_tokenizeMathLikeStat_(raw));
  });

  const textOnly = s.replace(/\$(.+?)\$/g, ' ');
  const textTokens = FS_tokenizeTextLikeStat_(textOnly);

  const patternTokens = FS_extractPatternTokensForSearch_(textOnly);
  const formTokens = FS_extractFormTokensForSearch_(mathRawList);

  return { textTokens, mathTokens, mathRawList, patternTokens, formTokens };
}

/* ---------------------------
 * Tokenize (Token_Stat 규칙 호환)
 * --------------------------- */
function FS_tokenizeTextLikeStat_(text) {
  return String(text || '')
    .toLowerCase()
    .replace(/[0-9]/g, ' ')
    .replace(/[^\p{L}]/gu, ' ')
    .split(/\s+/)
    .filter(t => t.length >= 1);
}

function FS_tokenizeMathLikeStat_(math) {
  return String(math || '')
    .replace(/\\[a-zA-Z]+/g, m => ` ${m} `)
    .replace(/[^a-zA-Z0-9\\]/g, ' $& ')
    .split(/\s+/)
    .filter(t => t.length >= 1)
    .filter(t => !/^\d+$/.test(t)); // 순수 숫자 제거
}

/* ---------------------------
 * PATTERN (PAT_ + PATG_)
 * --------------------------- */
function FS_extractPatternTokensForSearch_(textOnly) {
  const out = [];

  const s = String(textOnly || '').toLowerCase();
  if (/성립/.test(s)) out.push('PAT_성립');
  if (/다음(은|을)/.test(s)) out.push('PAT_다음');
  if (/모든\s*자연수/.test(s) || /자연수\s*n/.test(s)) out.push('PAT_모든자연수');
  if (/실수\s*전체/.test(s) || /실수\s*전체의\s*집합/.test(s)) out.push('PAT_실수전체');
  if (/를\s*만족/.test(s)) out.push('PAT_를만족');
  if (/라\s*하자/.test(s)) out.push('PAT_라하자');

  out.push(...FS_extractPatternNgramsForSearch_(textOnly));
  return out;
}

function FS_extractPatternNgramsForSearch_(textOnly) {
  const toks = FS_tokenizeTextLikeStat_(textOnly).filter(t => t.length >= 2);

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

/* ---------------------------
 * FORM tokens
 * --------------------------- */
function FS_extractFormTokensForSearch_(mathRawList) {
  const out = new Set();

  (mathRawList || []).forEach(raw => {
    const s = String(raw || '');

    if (s.includes('\\int')) out.add('FORM_int');
    if (s.includes('\\sum')) out.add('FORM_sum');
    if (s.includes('\\sin') || s.includes('\\cos') || s.includes('\\tan')) out.add('FORM_trig');
    if (s.includes('\\log') || s.includes('\\ln')) out.add('FORM_log');
    if (s.includes('\\lim')) out.add('FORM_lim');
    if (s.includes('\\sqrt')) out.add('FORM_sqrt');
    if (s.includes('\\frac')) out.add('FORM_frac');

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

/* ---------------------------
 * Stopwords / Noise
 * --------------------------- */
function FS_filterStopText_(tokens) {
  const STOP = new Set([
    '다음','각','모든','서로','대하여','하여','한다','이다','인','에서','의','을','를','은','는','이','가',
    '그리고','또는','혹은','또','또한','때','경우','조건','만족','성립','하여라','구하여라','구하라',
    '값','최댓값','최소값','최솟값','정답','보기','다음은','중','중에서',
    '실수','정수','자연수','양수','음수','유리수','무리수',
    '함수','수열','항','점','선분','직선','원','삼각형','사각형','다각형','좌표','평면',
    '구간','범위','일때','일때의','라할','라하면','라고','이라','이라면','이고','이며'
  ]);

  return (tokens || []).filter(t => {
    if (!t) return false;
    if (t.length <= 1) return false;
    if (STOP.has(t)) return false;
    return true;
  });
}

function FS_filterNoiseMath_(tokens) {
  const NOISE = new Set(['\\|', '*', '/', '.', ',', ':', ';']);
  return (tokens || []).filter(t => {
    if (!t) return false;
    if (NOISE.has(t)) return false;
    if (/^\s+$/.test(t)) return false;
    return true;
  });
}

/* ---------------------------
 * TF-IDF vector
 * --------------------------- */
function FS_buildTfidfVector_(tokens, idfMap) {
  const tf = new Map();
  (tokens || []).forEach(tok => tf.set(tok, (tf.get(tok) || 0) + 1));

  const vec = new Map();
  tf.forEach((cnt, tok) => {
    const w = idfMap.get(tok);
    if (!w) return;
    const val = (1 + Math.log(cnt)) * w;
    if (val > 0) vec.set(tok, val);
  });

  return vec;
}

/* ---------------------------
 * MATH vector + composite tokens
 * --------------------------- */
function FS_buildTfidfVectorWithMathExtras_(mathRawList, baseMathTokens, idfMap) {
  const tf = new Map();
  (baseMathTokens || []).forEach(tok => tf.set(tok, (tf.get(tok) || 0) + 1));

  const extras = FS_makeMathCompositeTokens_(mathRawList);
  extras.forEach(tok => tf.set(tok, (tf.get(tok) || 0) + 1));

  const vec = new Map();
  tf.forEach((cnt, tok) => {
    let w = idfMap.get(tok);
    if (!w) {
      w = FS_virtualIdfForComposite_(tok, idfMap);
      if (!w) return;
    }
    const val = (1 + Math.log(cnt)) * w;
    if (val > 0) vec.set(tok, val);
  });

  return vec;
}

function FS_makeMathCompositeTokens_(mathRawList) {
  const out = [];
  (mathRawList || []).forEach(raw => {
    const s = String(raw || '');

    for (const m of s.matchAll(/([a-zA-Z])\s*_\s*(\{[^}]+\}|[a-zA-Z0-9]+)/g)) {
      const v = m[1];
      const sub = m[2].replace(/[{}]/g, '');
      if (sub) out.push(`${v}_${sub}`);
    }
    for (const m of s.matchAll(/([a-zA-Z0-9])\s*\^\s*(\{[^}]+\}|[a-zA-Z0-9]+)/g)) {
      const v = m[1];
      const sup = m[2].replace(/[{}]/g, '');
      if (sup) out.push(`${v}^${sup}`);
    }

    if (s.includes('\\frac')) out.push('\\frac');
    if (s.includes('\\sqrt')) out.push('\\sqrt');
    if (s.includes('\\log')) out.push('\\log');
    if (s.includes('\\ln')) out.push('\\ln');
    if (s.includes('\\sin')) out.push('\\sin');
    if (s.includes('\\cos')) out.push('\\cos');
    if (s.includes('\\tan')) out.push('\\tan');
    if (s.includes('\\sum')) out.push('\\sum');
    if (s.includes('\\lim')) out.push('\\lim');
    if (s.includes('\\int')) out.push('\\int');
  });

  return out;
}

function FS_virtualIdfForComposite_(tok, idfMap) {
  if (tok.startsWith('\\')) {
    const w = idfMap.get(tok);
    return w || 0;
  }

  const parts = tok.match(/[a-zA-Z0-9]+/g) || [];
  let sum = 0, cnt = 0;

  for (const p of parts) {
    const w = idfMap.get(p);
    if (w) { sum += w; cnt++; }
  }
  if (cnt === 0) return 0;

  return (sum / cnt) * 1.1;
}

/* ---------------------------
 * Cosine utilities
 * --------------------------- */
function FS_vecNorm_(vec) {
  let s = 0;
  vec.forEach(v => { s += v * v; });
  return Math.sqrt(s);
}

function FS_dot_(a, b) {
  const [small, large] = (a.size <= b.size) ? [a, b] : [b, a];
  let s = 0;
  small.forEach((v, k) => {
    const u = large.get(k);
    if (u) s += v * u;
  });
  return s;
}

function FS_cosine_(qVec, qNorm, dVec) {
  const dNorm = FS_vecNorm_(dVec);
  if (qNorm === 0 || dNorm === 0) return 0;
  return FS_dot_(qVec, dVec) / (qNorm * dNorm);
}

/* ---------------------------
 * Misc
 * --------------------------- */
function FS_clampInt_(value, min, max, fallback) {
  const n = Number(value);
  if (!isFinite(n)) return fallback;
  const k = Math.floor(n);
  if (k < min) return min;
  if (k > max) return max;
  return k;
}

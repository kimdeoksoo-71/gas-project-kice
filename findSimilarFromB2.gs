/*****************************************************
 * Find Similar (TF-IDF + Cosine)
 *
 * ▸ Token_Stat의 TEXT/MATH/PATTERN/FORM 가중치를 활용
 * ▸ C2 대단원 필터: 값이 있으면 같은 chapter만 검색
 *
 * Input (문항검토 시트):
 *  B2: query latex (본문+수식)
 *  C2: chapter (대단원, 비우면 전체 검색)
 *  E2: topN (기본 10)
 *  F2: text weight (기본 0.7)
 *  G2: math weight (기본 0.3)
 *
 * Output:
 *  A5:E20 (21행 이하 보호)
 *****************************************************/


/* =========================================================
 * 상수
 * ========================================================= */

const FS_DEFAULT_TOP_N  = 10;
const FS_DEFAULT_W_TEXT = 0.7;
const FS_DEFAULT_W_MATH = 0.3;

const FS_HEADER_ROW = 5;
const FS_START_ROW  = 6;
const FS_LAST_ROW   = 20;   // 21행 이하 보호
const FS_LATEX_COL  = 4;    // D열

// composite 토큰의 가상 idf 배율
const FS_VIRTUAL_IDF_BOOST = 1.1;

// ── 공용 Stopword (빌드와 검색에서 동일하게 사용) ──
const FS_STOP_TEXT = new Set([
  '다음','다음은','각','모든','서로','대하여','하여','한다','이다','인',
  '에서','의','을','를','은','는','이','가',
  '그리고','또는','혹은','또','또한','때','경우','조건','만족','성립',
  '하여라','구하여라','구하라',
  '값','최댓값','최소값','최솟값','정답','보기','중','중에서',
  '실수','정수','자연수','양수','음수','유리수','무리수',
  '함수','수열','항','점','선분','직선','원','삼각형','사각형','다각형',
  '좌표','평면','구간','범위',
  '일때','일때의','이라','이라면','이고','이며','라고','라할','라하면'
]);

const FS_STOP_MATH = new Set(['\\|', '*', '/', '.', ',', ':', ';']);


/* =========================================================
 * 메인 함수
 * ========================================================= */

function findSimilarFromB2(options) {
  options = options || {};

  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getActiveSheet();
  var dataSheet = ss.getSheetByName('Data_Latex');
  var statSheet = ss.getSheetByName('Token_Stat');

  if (!dataSheet || !statSheet) {
    ss.toast('Data_Latex 또는 Token_Stat 시트를 찾을 수 없습니다.', '오류', 5);
    return;
  }

  // ── 입력 ──
  var query = String(sheet.getRange('B2').getDisplayValue() || '').trim();
  if (!query) {
    ss.toast('B2 셀에 비교할 문항을 입력하세요.', '안내', 5);
    return;
  }

  var queryChapter = String(sheet.getRange('C2').getDisplayValue() || '').trim();
  var useChapterFilter = !!queryChapter;

  var maxOutput = FS_LAST_ROW - FS_START_ROW + 1;  // 15
  var topNCount = Math.min(
    _clampInt_(sheet.getRange('E2').getValue(), 1, maxOutput, FS_DEFAULT_TOP_N),
    maxOutput
  );

  var weights = _readWeights_(sheet);

  // ── Token_Stat 로드 ──
  var idfMaps = _loadIdfMaps_(statSheet);

  // ── Query 벡터 ──
  var qTokens    = _extractTokens_(query);
  var qTextToks  = _filterStopText_(qTokens.textTokens.concat(qTokens.patternTokens));
  var qMathToks  = _filterStopMath_(qTokens.mathTokens.concat(qTokens.formTokens));
  var qTextVec   = _buildTfidfVec_(qTextToks, idfMaps.text);
  var qMathVec   = _buildMathVec_(qTokens.mathRawList, qMathToks, idfMaps.math);
  var qTextNorm  = _vecNorm_(qTextVec);
  var qMathNorm  = _vecNorm_(qMathVec);

  // ── Data 스캔 ──
  var data = dataSheet.getDataRange().getValues();
  var scored = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var latex   = row[2];  // C
    if (!latex) continue;

    var chapter = String(row[9] || '').trim();  // J
    if (useChapterFilter && chapter !== queryChapter) continue;

    var dTokens   = _extractTokens_(latex);
    var dTextToks = _filterStopText_(dTokens.textTokens.concat(dTokens.patternTokens));
    var dMathToks = _filterStopMath_(dTokens.mathTokens.concat(dTokens.formTokens));
    var dTextVec  = _buildTfidfVec_(dTextToks, idfMaps.text);
    var dMathVec  = _buildMathVec_(dTokens.mathRawList, dMathToks, idfMaps.math);

    var textScore = _cosine_(qTextVec, qTextNorm, dTextVec);
    var mathScore = _cosine_(qMathVec, qMathNorm, dMathVec);
    var score = weights.wText * textScore + weights.wMath * mathScore;

    if (score > 0) {
      scored.push({
        score: score,
        source: row[8],     // I
        chapter: chapter,
        driveLink: row[1],  // B
        latex: latex,
        rowIndex: i + 1
      });
    }
  }

  // ── 정렬 + 출력 ──
  scored.sort(function(a, b) { return b.score - a.score || a.rowIndex - b.rowIndex; });
  var topN = scored.slice(0, topNCount);
  topN.sort(function(a, b) { return String(b.source).localeCompare(String(a.source)); });

  _writeResults_(sheet, topN);

  ss.toast(
    (useChapterFilter
      ? '유사문항 ' + topN.length + '개 (대단원: "' + queryChapter + '"'
      : '유사문항 ' + topN.length + '개 (전체 검색')
    + ', TEXT=' + weights.wText.toFixed(2) + ', MATH=' + weights.wMath.toFixed(2) + ')',
    '검색', 3
  );
}


/* =========================================================
 * 입력 읽기
 * ========================================================= */

function _readWeights_(sheet) {
  var t = Number(sheet.getRange('F2').getValue());
  var m = Number(sheet.getRange('G2').getValue());
  if (!isFinite(t)) t = FS_DEFAULT_W_TEXT;
  if (!isFinite(m)) m = FS_DEFAULT_W_MATH;
  t = Math.max(0, t);
  m = Math.max(0, m);
  var sum = t + m;
  if (sum <= 0) return { wText: FS_DEFAULT_W_TEXT, wMath: FS_DEFAULT_W_MATH };
  return { wText: t / sum, wMath: m / sum };
}


/* =========================================================
 * Token_Stat 로드
 * ========================================================= */

function _loadIdfMaps_(statSheet) {
  var last = statSheet.getLastRow();
  var n = Math.max(last - 1, 0);
  var vals = n ? statSheet.getRange(2, 1, n, 6).getValues() : [];

  var textMap = new Map();
  var mathMap = new Map();

  for (var i = 0; i < vals.length; i++) {
    var token = String(vals[i][0] || '').trim();
    var type  = String(vals[i][1] || '').trim();
    var w     = Number(vals[i][4]);
    if (!token || !isFinite(w)) continue;

    if (type === 'TEXT' || type === 'PATTERN') textMap.set(token, w);
    else if (type === 'MATH' || type === 'FORM') mathMap.set(token, w);
  }

  return { text: textMap, math: mathMap };
}


/* =========================================================
 * 토큰 추출
 * ========================================================= */

function _extractTokens_(latex) {
  var s = String(latex || '');

  // 수식 추출
  var mathRawList  = [];
  var mathTokens   = [];
  var mathMatches  = s.match(/\$(.+?)\$/g) || [];

  for (var i = 0; i < mathMatches.length; i++) {
    var raw = mathMatches[i].slice(1, -1);  // $ 제거
    mathRawList.push(raw);
    mathTokens = mathTokens.concat(_tokenizeMath_(raw));
  }

  // 텍스트 추출
  var textOnly    = s.replace(/\$(.+?)\$/g, ' ');
  var textTokens  = _tokenizeText_(textOnly);

  // PATTERN + FORM
  var patternTokens = _extractPatternTokens_(textOnly);
  var formTokens    = _extractFormTokens_(mathRawList);

  return {
    textTokens: textTokens,
    mathTokens: mathTokens,
    mathRawList: mathRawList,
    patternTokens: patternTokens,
    formTokens: formTokens
  };
}

function _tokenizeText_(text) {
  return String(text || '')
    .toLowerCase()
    .replace(/[0-9]/g, ' ')
    .replace(/[^\p{L}]/gu, ' ')
    .split(/\s+/)
    .filter(function(t) { return t.length >= 1; });
}

function _tokenizeMath_(math) {
  return String(math || '')
    .replace(/\\[a-zA-Z]+/g, function(m) { return ' ' + m + ' '; })
    .replace(/[^a-zA-Z0-9\\]/g, function(m) { return ' ' + m + ' '; })
    .split(/\s+/)
    .filter(function(t) { return t.length >= 1; })
    .filter(function(t) { return !/^\d+$/.test(t); });
}


/* =========================================================
 * PATTERN 토큰 (PAT_ + PATG_)
 * ========================================================= */

function _extractPatternTokens_(textOnly) {
  var out = [];
  var s = String(textOnly || '').toLowerCase();

  // 고정 패턴 (PAT_)
  if (/성립/.test(s))                                     out.push('PAT_성립');
  if (/다음(은|을)/.test(s))                               out.push('PAT_다음');
  if (/모든\s*자연수/.test(s) || /자연수\s*n/.test(s))      out.push('PAT_모든자연수');
  if (/실수\s*전체/.test(s) || /실수\s*전체의\s*집합/.test(s)) out.push('PAT_실수전체');
  if (/를\s*만족/.test(s))                                 out.push('PAT_를만족');
  if (/라\s*하자/.test(s))                                 out.push('PAT_라하자');

  // n-gram (PATG_) — 공용 stopword 사용
  var toks = _tokenizeText_(textOnly).filter(function(t) { return t.length >= 2; });
  var clean = toks.filter(function(t) { return !FS_STOP_TEXT.has(t); });

  if (clean.length >= 2) {
    var seen = {};
    var addN = function(n) {
      for (var i = 0; i + n <= clean.length; i++) {
        var key = 'PATG_' + clean.slice(i, i + n).join('_');
        if (!seen[key]) { seen[key] = true; out.push(key); }
      }
    };
    addN(2); addN(3); addN(4);
  }

  return out;
}


/* =========================================================
 * FORM 토큰
 * ========================================================= */

function _extractFormTokens_(mathRawList) {
  var out = {};  // Set 대용

  for (var i = 0; i < (mathRawList || []).length; i++) {
    var s = String(mathRawList[i] || '');

    // 기본
    if (s.indexOf('\\int') !== -1)  out['FORM_int'] = 1;
    if (s.indexOf('\\sum') !== -1)  out['FORM_sum'] = 1;
    if (s.indexOf('\\sin') !== -1 || s.indexOf('\\cos') !== -1 || s.indexOf('\\tan') !== -1) out['FORM_trig'] = 1;
    if (s.indexOf('\\log') !== -1 || s.indexOf('\\ln') !== -1)  out['FORM_log'] = 1;
    if (s.indexOf('\\lim') !== -1)  out['FORM_lim'] = 1;
    if (s.indexOf('\\sqrt') !== -1) out['FORM_sqrt'] = 1;
    if (s.indexOf('\\frac') !== -1) out['FORM_frac'] = 1;

    // 구조
    if (/_/.test(s))  out['FORM_subscript'] = 1;
    if (/\^/.test(s)) out['FORM_superscript'] = 1;
    if (s.indexOf('\\left|') !== -1 || s.indexOf('\\right|') !== -1 || s.indexOf('|') !== -1) out['FORM_abs'] = 1;
    if (s.indexOf('\\begin{cases}') !== -1 || s.indexOf('\\cases') !== -1) out['FORM_piecewise'] = 1;
    if (/[()[\]]/.test(s)) out['FORM_interval'] = 1;

    var eqCount = (s.match(/=/g) || []).length;
    if (eqCount >= 2) out['FORM_eq_chain'] = 1;

    // 조합
    var hasFrac = s.indexOf('\\frac') !== -1;
    var hasSqrt = s.indexOf('\\sqrt') !== -1;
    var hasLog  = s.indexOf('\\log') !== -1 || s.indexOf('\\ln') !== -1;
    var hasLim  = s.indexOf('\\lim') !== -1;
    var hasInt  = s.indexOf('\\int') !== -1;
    var hasSum  = s.indexOf('\\sum') !== -1;

    if (hasFrac && hasSqrt) out['FORM_frac+sqrt'] = 1;
    if (hasLog && hasLim)   out['FORM_log+lim'] = 1;
    if (hasInt && hasFrac)  out['FORM_int+frac'] = 1;
    if (hasSum && hasFrac)  out['FORM_sum+frac'] = 1;
  }

  return Object.keys(out);
}


/* =========================================================
 * Stopword 필터 (공용 상수 사용)
 * ========================================================= */

function _filterStopText_(tokens) {
  return (tokens || []).filter(function(t) {
    if (!t || t.length <= 1) return false;
    if (FS_STOP_TEXT.has(t)) return false;
    return true;
  });
}

function _filterStopMath_(tokens) {
  return (tokens || []).filter(function(t) {
    if (!t) return false;
    if (FS_STOP_MATH.has(t)) return false;
    if (/^\s+$/.test(t)) return false;
    return true;
  });
}


/* =========================================================
 * TF-IDF 벡터
 * ========================================================= */

function _buildTfidfVec_(tokens, idfMap) {
  var tf = new Map();
  for (var i = 0; i < (tokens || []).length; i++) {
    var tok = tokens[i];
    tf.set(tok, (tf.get(tok) || 0) + 1);
  }

  var vec = new Map();
  tf.forEach(function(cnt, tok) {
    var w = idfMap.get(tok);
    if (!w) return;
    var val = (1 + Math.log(cnt)) * w;
    if (val > 0) vec.set(tok, val);
  });

  return vec;
}

/** MATH 벡터: 기본 토큰 + composite (subscript/superscript 패턴) */
function _buildMathVec_(mathRawList, baseMathTokens, idfMap) {
  var tf = new Map();
  for (var i = 0; i < (baseMathTokens || []).length; i++) {
    var tok = baseMathTokens[i];
    tf.set(tok, (tf.get(tok) || 0) + 1);
  }

  // composite 토큰 (a_n, x^2 등 구조적 패턴)
  var composites = _makeCompositeTokens_(mathRawList);
  for (var j = 0; j < composites.length; j++) {
    var ct = composites[j];
    tf.set(ct, (tf.get(ct) || 0) + 1);
  }

  var vec = new Map();
  tf.forEach(function(cnt, tok) {
    var w = idfMap.get(tok);
    if (!w) {
      w = _virtualIdf_(tok, idfMap);
      if (!w) return;
    }
    var val = (1 + Math.log(cnt)) * w;
    if (val > 0) vec.set(tok, val);
  });

  return vec;
}

/**
 * composite 토큰: subscript/superscript 패턴만 추출
 * (기존 코드에서 \frac, \sqrt 등을 중복 push하던 문제 제거)
 */
function _makeCompositeTokens_(mathRawList) {
  var out = [];
  for (var i = 0; i < (mathRawList || []).length; i++) {
    var s = String(mathRawList[i] || '');

    // a_n, f_{n+1} 등
    var subMatches = s.match(/([a-zA-Z])\s*_\s*(\{[^}]+\}|[a-zA-Z0-9]+)/g) || [];
    for (var j = 0; j < subMatches.length; j++) {
      var m = subMatches[j].match(/([a-zA-Z])\s*_\s*(\{[^}]+\}|[a-zA-Z0-9]+)/);
      if (m) {
        var sub = m[2].replace(/[{}]/g, '');
        if (sub) out.push(m[1] + '_' + sub);
      }
    }

    // x^2, n^k 등
    var supMatches = s.match(/([a-zA-Z0-9])\s*\^\s*(\{[^}]+\}|[a-zA-Z0-9]+)/g) || [];
    for (var k = 0; k < supMatches.length; k++) {
      var m2 = supMatches[k].match(/([a-zA-Z0-9])\s*\^\s*(\{[^}]+\}|[a-zA-Z0-9]+)/);
      if (m2) {
        var sup = m2[2].replace(/[{}]/g, '');
        if (sup) out.push(m2[1] + '^' + sup);
      }
    }
  }
  return out;
}

function _virtualIdf_(tok, idfMap) {
  // LaTeX 명령어는 직접 조회
  if (tok.charAt(0) === '\\') return idfMap.get(tok) || 0;

  // composite는 구성 요소의 idf 평균
  var parts = tok.match(/[a-zA-Z0-9]+/g) || [];
  var sum = 0, cnt = 0;
  for (var i = 0; i < parts.length; i++) {
    var w = idfMap.get(parts[i]);
    if (w) { sum += w; cnt++; }
  }
  if (cnt === 0) return 0;
  return (sum / cnt) * FS_VIRTUAL_IDF_BOOST;
}


/* =========================================================
 * Cosine 유사도
 * ========================================================= */

function _vecNorm_(vec) {
  var s = 0;
  vec.forEach(function(v) { s += v * v; });
  return Math.sqrt(s);
}

function _cosine_(qVec, qNorm, dVec) {
  var dNorm = _vecNorm_(dVec);
  if (qNorm === 0 || dNorm === 0) return 0;

  // dot product (작은 쪽 기준 순회)
  var dot = 0;
  var small = qVec, large = dVec;
  if (qVec.size > dVec.size) { small = dVec; large = qVec; }
  small.forEach(function(v, k) {
    var u = large.get(k);
    if (u) dot += v * u;
  });

  return dot / (qNorm * dNorm);
}


/* =========================================================
 * 출력
 * ========================================================= */

function _writeResults_(sheet, topN) {
  // 출력 영역 초기화
  sheet.getRange(FS_HEADER_ROW, 1, FS_LAST_ROW - FS_HEADER_ROW + 1, 5).clearContent();

  // 헤더
  sheet.getRange(FS_HEADER_ROW, 1, 1, 5).setValues([
    ['source', 'chapter', 'drive_link', 'Latex', 'score']
  ]);

  // 데이터
  if (topN.length === 0) return;

  var rows = [];
  for (var i = 0; i < topN.length; i++) {
    rows.push([
      topN[i].source,
      topN[i].chapter,
      topN[i].driveLink,
      topN[i].latex,
      topN[i].score
    ]);
  }
  sheet.getRange(FS_START_ROW, 1, rows.length, 5).setValues(rows);
}


/* =========================================================
 * 유틸
 * ========================================================= */

function _clampInt_(value, min, max, fallback) {
  var n = Number(value);
  if (!isFinite(n)) return fallback;
  var k = Math.floor(n);
  return Math.max(min, Math.min(max, k));
}
/*******************************************************
 * 문항검토: 기출 스타일 문장 다듬기 (멀티 LLM)
 *
 * ▸ provider별 API 호출을 통합하여 동일한 후처리 파이프라인 공유.
 * ▸ 배치 코드에서 options.provider = 'claude'|'gpt'|'gemini' 로 선택.
 *
 * [prompt 시트 키]
 *   SYSTEM_KICE_CLAUDE / SYSTEM_KICE_GPT / SYSTEM_KICE_GEMINI
 *   OUTPUT_FORMAT_KICE  (공통, 필수)
 *
 * [스크립트 속성]
 *   ANTHROPIC_API_KEY / OPENAI_API_KEY / GEMINI_API_KEY
 *******************************************************/


/* =========================================================
 * 상수 정의 — 모델·API·파라미터는 여기서만 관리
 * ========================================================= */

// ── 모델명 (업그레이드 시 여기만 수정) ──
var MODEL_CLAUDE = 'claude-sonnet-4-6';
var MODEL_GPT    = 'gpt-5.4';
var MODEL_GEMINI = 'gemini-2.5-flash';

// ── API 엔드포인트 ──
var API_URL_CLAUDE = 'https://api.anthropic.com/v1/messages';
var API_URL_GPT    = 'https://api.openai.com/v1/chat/completions';
var API_URL_GEMINI_BASE = 'https://generativelanguage.googleapis.com/v1beta/models/';

// ── 생성 파라미터 ──
var TEMPERATURE_CLAUDE = 0.15;
var TEMPERATURE_GPT    = 0.15;
var TEMPERATURE_GEMINI = 0.10;
var MAX_TOKENS_CLAUDE  = 4096;
var MAX_TOKENS_GPT     = 4096;
var MAX_TOKENS_GEMINI  = 65536;

// ── 시트명 ──
var SHEET_NAME        = '문항검토';
var RULE_SHEET_NAME   = 'Rule';
var PROMPT_SHEET_NAME = 'prompt';

// ── provider별 prompt 시트 키 ──
var PROVIDER_PROMPT_KEYS = {
  claude: 'SYSTEM_KICE_CLAUDE',
  gpt:    'SYSTEM_KICE_GPT',
  gemini: 'SYSTEM_KICE_GEMINI'
};

// ── provider별 스크립트 속성 키 ──
var PROVIDER_API_KEY_NAMES = {
  claude: 'ANTHROPIC_API_KEY',
  gpt:    'OPENAI_API_KEY',
  gemini: 'GEMINI_API_KEY'
};

// ── 공통 출력 형식 키 ──
var PROMPT_KEY_OUTPUT_FORMAT = 'OUTPUT_FORMAT_KICE';

// ── 재시도 설정 ──
var MAX_RETRIES     = 4;
var INITIAL_WAIT_MS = 5000;


/* =========================================================
 * 메인 함수
 * ========================================================= */

/**
 * @param {Object} options
 *   - provider: 'claude' | 'gpt' | 'gemini' (기본: 'claude')
 */
function review_rewriteToKiceStyle(options) {
  options = options || {};
  var provider = String(options.provider || 'claude').toLowerCase();

  if (!PROVIDER_PROMPT_KEYS[provider]) {
    throw new Error('지원하지 않는 provider: ' + provider);
  }

  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) throw new Error('시트를 찾을 수 없습니다: ' + SHEET_NAME);

  var apiKey = PropertiesService.getScriptProperties()
    .getProperty(PROVIDER_API_KEY_NAMES[provider]);
  if (!apiKey) throw new Error('스크립트 속성이 비어있습니다: ' + PROVIDER_API_KEY_NAMES[provider]);

  // 0) 기존 출력 초기화
  _clearOutput_(sh);

  // 1) 입력 읽기
  var targetRaw = String(sh.getRange('B2').getValue() || '').trim();
  var chapter   = String(sh.getRange('C2').getValue() || '').trim();

  if (!targetRaw) {
    sh.getRange('B3').setValue('수정구절 없음');
    sh.getRange('D21').setValue('수정 구절 없음');
    return;
  }

  // 2) Rule 로드 + 선치환
  var rules = _loadRules_();
  var preRule = _applyRulesWithLog_(targetRaw, rules);
  var target = preRule.text;
  var ruleEditsAll = preRule.edits.slice();

  // 3) 관련 기출문항 로드
  var sources    = _flatCol_(sh, 'A6:A15');
  var imageLinks = _flatCol_(sh, 'C6:C15');
  var refsText   = _flatCol_(sh, 'D6:D15');

  var refs = [];
  for (var i = 0; i < refsText.length; i++) {
    if (refsText[i]) {
      refs.push({ source: sources[i] || '', imageLink: imageLinks[i] || '', text: refsText[i] });
    }
  }

  // 4) 프롬프트 구성 + LLM 호출
  var systemPrompt = _buildSystemPrompt_(provider, rules);
  var userMessage  = _buildUserMessage_(target, chapter, refs);
  var parsed       = _callLLM_(provider, apiKey, systemPrompt, userMessage);

  // 5) 수정 구절 필터링 + 검증
  var editsRaw = Array.isArray(parsed.edits) ? parsed.edits : [];

  var llmEdits = editsRaw
    // 공백만 다른 수정 제거
    .filter(function(e) { return !_isWhitespaceOnlyEdit_(e.original, e.revised); })
    // evidence_quote 검증
    .filter(function(e) { return _validateEvidence_(e, refs); })
    // 기출에 이미 있는 original은 수정 금지
    .filter(function(e) {
      var orig = String(e.original || '').trim();
      if (orig && _existsInAnyRef_(orig, refs)) {
        var rev = String(e.revised || '').trim();
        if (!rev || !_existsInAnyRef_(rev, refs)) return false;
      }
      return true;
    })
    .map(function(e) {
      var rawIdx = _sanitizeSourceIndex_(e.source_index, refs.length);
      var verified = _verifyOrFindSourceIndex_(rawIdx, String(e.revised || '').trim(), refs);
      return {
        source_index: verified,
        original: String(e.original || '').trim(),
        revised: String(e.revised || '').trim(),
        evidence_quote: String(e.evidence_quote || '').trim(),
        reason: String(e.reason || '').trim(),
        _kind: provider.toUpperCase()
      };
    });

  // 6) 역방향 수정 되돌리기 목록
  var rejectedEdits = editsRaw
    .filter(function(e) { return !_isWhitespaceOnlyEdit_(e.original, e.revised); })
    .filter(function(e) {
      var orig = String(e.original || '').trim();
      var rev  = String(e.revised || '').trim();
      return orig && _existsInAnyRef_(orig, refs) && (!rev || !_existsInAnyRef_(rev, refs));
    });

  // 7) rewritten_full 처리
  var rewrittenFromModel = String(parsed.rewritten_full || '').trim();

  if (rewrittenFromModel && rewrittenFromModel.replace(/\s+/g, '') === target.replace(/\s+/g, '')) {
    rewrittenFromModel = target;
  }

  // 잘못된 수정 되돌리기
  if (rewrittenFromModel && rejectedEdits.length > 0) {
    for (var ri = 0; ri < rejectedEdits.length; ri++) {
      var rej = rejectedEdits[ri];
      var rejRevised  = String(rej.revised || '').trim();
      var rejOriginal = String(rej.original || '').trim();
      if (rejRevised && rejOriginal && rewrittenFromModel.indexOf(rejRevised) !== -1) {
        rewrittenFromModel = rewrittenFromModel.split(rejRevised).join(rejOriginal);
      }
    }
  }

  var baseFull = rewrittenFromModel || target;
  var postRule = _applyRulesWithLog_(baseFull, rules);
  var finalFull = postRule.text;

  if (postRule.edits.length) {
    postRule.edits.forEach(function(e) {
      ruleEditsAll.push({
        from: e.from, to: e.to, count: e.count,
        note: e.note, sample: e.sample, phase: 'POST'
      });
    });
  }

  sh.getRange('B3').setValue(finalFull);

  // 8) 출력 (D21 이하)
  _writeOutput_(sh, ruleEditsAll, llmEdits, refs, parsed.has_edits);
}


/* =========================================================
 * 프롬프트 구성
 * ========================================================= */

function _buildSystemPrompt_(provider, rules) {
  var promptKey = PROVIDER_PROMPT_KEYS[provider];
  var basePrompt = _getPromptValue_(promptKey);
  var outputFormat = _getPromptValue_(PROMPT_KEY_OUTPUT_FORMAT);

  var lines = [basePrompt, ''];

  // 강제 치환 규칙 (Rule) — 동적 삽입 (이건 프롬프트가 아니라 런타임 데이터)
  if (Array.isArray(rules) && rules.length > 0) {
    lines.push('<forced_rules>');
    lines.push('아래는 강제 수정 규칙이다. 이미 원문에 선적용되어 있다.');
    lines.push('- "to" 표현이 최종 결과(rewritten_full)에 그대로 유지되어야 한다.');
    lines.push('- "from" 표현이 최종 결과에 남아 있으면 안 된다.');
    rules.forEach(function(r, i) {
      var note = r.note ? ' (note: ' + r.note + ')' : '';
      lines.push('(' + (i + 1) + ') "' + r.from + '" → "' + r.to + '"' + note);
    });
    lines.push('</forced_rules>');
    lines.push('');
  }

  lines.push(outputFormat);
  return lines.join('\n');
}

function _buildUserMessage_(target, chapter, refs) {
  var lines = [];

  if (chapter) lines.push('<chapter>' + chapter + '</chapter>');
  lines.push('');

  var segments = _splitToSegments_(target);
  lines.push('<target>');
  lines.push('[원문 전체]');
  lines.push(target);
  lines.push('');
  lines.push('[구절별 분해]');
  for (var s = 0; s < segments.length; s++) {
    lines.push('S' + (s + 1) + ': ' + segments[s]);
  }
  lines.push('</target>');
  lines.push('');

  lines.push('<references>');
  if (refs.length === 0) {
    lines.push('(관련 기출문항 없음)');
  } else {
    refs.forEach(function(r, i) {
      lines.push('<ref index="' + (i + 1) + '" source="' + (r.source || '') + '">');
      lines.push(r.text);
      lines.push('</ref>');
    });
  }
  lines.push('</references>');

  return lines.join('\n');
}


/* =========================================================
 * LLM 호출 (provider별 분기 → JSON 파싱)
 * ========================================================= */

function _callLLM_(provider, apiKey, systemPrompt, userMessage) {
  var jsonText;

  if (provider === 'claude')     jsonText = _callClaude_(apiKey, systemPrompt, userMessage);
  else if (provider === 'gpt')   jsonText = _callGPT_(apiKey, systemPrompt, userMessage);
  else if (provider === 'gemini') jsonText = _callGemini_(apiKey, systemPrompt, userMessage);
  else throw new Error('알 수 없는 provider: ' + provider);

  jsonText = _fixLatexEscapesInJson_(jsonText);

  try {
    return JSON.parse(jsonText);
  } catch (e) {
    var preview = jsonText.length > 500
      ? jsonText.substring(0, 300) + '\n...\n' + jsonText.substring(jsonText.length - 200)
      : jsonText;
    throw new Error(provider.toUpperCase() + ' JSON 파싱 실패:\n' + preview);
  }
}

/* ── Claude ── */
function _callClaude_(apiKey, systemPrompt, userMessage) {
  var body = {
    model: MODEL_CLAUDE,
    max_tokens: MAX_TOKENS_CLAUDE,
    temperature: TEMPERATURE_CLAUDE,
    system: systemPrompt,
    messages: [
      { role: 'user', content: userMessage },
      { role: 'assistant', content: '{' }
    ],
    stop_sequences: ['\n}']
  };

  var resp = _httpPost_(API_URL_CLAUDE, body, {
    'x-api-key': apiKey,
    'anthropic-version': '2023-06-01'
  });

  var rawText = '';
  if (resp && Array.isArray(resp.content)) {
    for (var i = 0; i < resp.content.length; i++) {
      if (resp.content[i].type === 'text') { rawText = resp.content[i].text || ''; break; }
    }
  }
  return '{' + rawText.trim() + '}';
}

/* ── GPT ── */
function _callGPT_(apiKey, systemPrompt, userMessage) {
  var body = {
    model: MODEL_GPT,
    max_completion_tokens: MAX_TOKENS_GPT,
    temperature: TEMPERATURE_GPT,
    response_format: { type: 'json_object' },
    messages: [
      { role: 'system', content: systemPrompt },
      { role: 'user',   content: userMessage }
    ]
  };

  var resp = _httpPost_(API_URL_GPT, body, {
    'Authorization': 'Bearer ' + apiKey
  });

  if (resp && resp.choices && resp.choices.length > 0 && resp.choices[0].message) {
    return resp.choices[0].message.content.trim();
  }
  throw new Error('GPT 응답에서 content를 추출하지 못했습니다.');
}

/* ── Gemini ── */
function _callGemini_(apiKey, systemPrompt, userMessage) {
  var body = {
    contents: [
      { role: 'user', parts: [{ text: userMessage }] }
    ],
    systemInstruction: {
      parts: [{ text: systemPrompt }]
    },
    generationConfig: {
      temperature: TEMPERATURE_GEMINI,
      maxOutputTokens: MAX_TOKENS_GEMINI,
      responseMimeType: 'application/json'
    }
  };

  var url = API_URL_GEMINI_BASE + MODEL_GEMINI + ':generateContent?key=' + apiKey;
  var resp = _httpPost_(url, body, {});

  if (resp && resp.candidates && resp.candidates.length > 0) {
    var candidate = resp.candidates[0];

    if (candidate.finishReason === 'SAFETY' || candidate.finishReason === 'BLOCKED') {
      throw new Error('Gemini: 안전 필터에 의해 응답이 차단되었습니다.');
    }
    if (candidate.finishReason === 'MAX_TOKENS') {
      throw new Error('Gemini: 출력 토큰 제한을 초과했습니다.');
    }

    var parts = candidate.content && candidate.content.parts;
    if (parts) {
      for (var i = 0; i < parts.length; i++) {
        if (parts[i].text) {
          var text = parts[i].text.trim()
            .replace(/^```json\s*/i, '').replace(/\s*```$/i, '');
          var first = text.indexOf('{');
          var last  = text.lastIndexOf('}');
          if (first !== -1 && last > first) return text.substring(first, last + 1);
          return text;
        }
      }
    }
  }

  var detail = '';
  if (resp && resp.promptFeedback && resp.promptFeedback.blockReason) {
    detail = ' (사유: ' + resp.promptFeedback.blockReason + ')';
  }
  throw new Error('Gemini 응답에서 유효한 텍스트를 추출하지 못했습니다.' + detail);
}


/* =========================================================
 * HTTP (지수 백오프 재시도)
 * ========================================================= */

function _httpPost_(url, body, extraHeaders) {
  var headers = {};
  for (var key in extraHeaders) headers[key] = extraHeaders[key];

  var options = {
    method: 'post',
    contentType: 'application/json',
    headers: headers,
    muteHttpExceptions: true,
    payload: JSON.stringify(body)
  };

  var lastCode = 0, lastText = '';

  for (var attempt = 0; attempt <= MAX_RETRIES; attempt++) {
    var res = UrlFetchApp.fetch(url, options);
    lastCode = res.getResponseCode();
    lastText = res.getContentText();

    if (lastCode >= 200 && lastCode < 300) return JSON.parse(lastText);

    var retryable = (lastCode === 429 || lastCode === 529 || lastCode === 500 || lastCode === 503);
    if (!retryable || attempt >= MAX_RETRIES) break;

    Utilities.sleep(INITIAL_WAIT_MS * Math.pow(2, attempt));
  }

  throw new Error('API 오류 (' + lastCode + '):\n' + lastText.substring(0, 500));
}


/* =========================================================
 * evidence_quote 검증
 * ========================================================= */

function _validateEvidence_(edit, refs) {
  var revised = String(edit.revised || '').trim();
  if (!revised) return false;

  // 핵심 검증: revised가 어떤 ref에든 글자 그대로 포함되는지
  if (!_existsInAnyRef_(revised, refs)) return false;

  // 보너스 검증: evidence_quote가 있으면 실제 ref와 대조
  var evidence = String(edit.evidence_quote || '').trim();
  var srcIdx = _sanitizeSourceIndex_(edit.source_index, refs.length);

  if (evidence && srcIdx !== null) {
    var refText      = (refs[srcIdx - 1].text || '').replace(/\s+/g, '');
    var evidenceNorm = evidence.replace(/\s+/g, '');
    // evidence가 ref에 없으면 source_index가 틀린 것이지만, 수정 자체는 유지
    // (source_index는 _verifyOrFindSourceIndex_에서 후보정됨)
  }

  return true;
}


/* =========================================================
 * 출력
 * ========================================================= */

function _writeOutput_(sh, ruleEditsAll, llmEdits, refs, hasEditsFromModel) {
  var startRow = 21;
  var maxRows  = 200;

  var ruleOut = ruleEditsAll.map(function(e) {
    return { _kind: 'RULE', source: 'RULE', imageLink: '', text: _formatRuleEdit_(e) };
  });

  var llmOut = llmEdits.map(function(e) {
    var refIdx = (e.source_index !== null && e.source_index >= 1 && e.source_index <= refs.length)
      ? (e.source_index - 1) : -1;
    return {
      _kind: e._kind,
      source: refIdx >= 0 ? (refs[refIdx].source || '') : '',
      imageLink: refIdx >= 0 ? (refs[refIdx].imageLink || '') : '',
      text: '[[원본]] ' + e.original
        + '\n[[수정]] ' + e.revised
        + '\n[[근거]] ' + (e.evidence_quote || '(없음)')
        + '\n[[이유]] ' + e.reason
    };
  });

  var outItems = ruleOut.concat(llmOut);

  if (outItems.length === 0 || (!hasEditsFromModel && ruleOut.length === 0)) {
    sh.getRange('D21').setValue('수정 구절 없음');
    return;
  }

  var rowCount = Math.min(outItems.length, maxRows);
  var outA = [], outC = [], outD = [];
  for (var j = 0; j < rowCount; j++) {
    outA.push([outItems[j].source || '']);
    outC.push([outItems[j].imageLink || '']);
    outD.push([outItems[j].text || '']);
  }

  sh.getRange(startRow, 1, rowCount, 1).setValues(outA);
  sh.getRange(startRow, 3, rowCount, 1).setValues(outC);
  sh.getRange(startRow, 4, rowCount, 1).setValues(outD);
}


/* =========================================================
 * Rule: 로드 + 적용
 * ========================================================= */

function _loadRules_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(RULE_SHEET_NAME);
  if (!sh) return [];
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  var values = sh.getRange(2, 1, lastRow - 1, 4).getValues();
  var rules = [];

  for (var i = 0; i < values.length; i++) {
    var from    = String(values[i][0] || '').trim();
    var to      = String(values[i][1] || '').trim();
    var enabled = values[i][2];
    var note    = String(values[i][3] || '').trim();
    var isOn = (enabled === true) || (String(enabled).toUpperCase() === 'TRUE');
    if (!isOn || !from || from === to) continue;
    if (to && to.indexOf(from) !== -1) continue;
    rules.push({ from: from, to: to, note: note });
  }

  rules.sort(function(a, b) { return b.from.length - a.from.length; });
  return rules;
}

function _applyRulesWithLog_(text, rules) {
  var cur = String(text || '');
  if (!cur || !Array.isArray(rules) || rules.length === 0) return { text: cur, edits: [] };

  var edits = [];
  for (var i = 0; i < rules.length; i++) {
    var r = rules[i];
    var count = _countOccurrences_(cur, r.from);
    if (count <= 0) continue;
    var idx = cur.indexOf(r.from);
    var sample = cur.slice(Math.max(0, idx - 20), Math.min(cur.length, idx + r.from.length + 20));
    cur = cur.split(r.from).join(r.to);
    edits.push({ from: r.from, to: r.to, count: count, note: r.note, sample: sample });
  }
  return { text: cur, edits: edits };
}

function _formatRuleEdit_(e) {
  var phase = e.phase || 'PRE';
  var reason = 'Rule(' + phase + '): "' + e.from + '" → "' + e.to + '" | ' + e.count + '회';
  if (e.note) reason += ' | note: ' + e.note;
  return '[[원본]] ' + e.from + '\n[[수정]] ' + e.to + '\n[[이유]] ' + reason;
}


/* =========================================================
 * 구절 분해
 * ========================================================= */

function _splitToSegments_(text) {
  if (!text) return [];
  var rawLines = text.split(/\n/);
  var segments = [];
  for (var i = 0; i < rawLines.length; i++) {
    var line = rawLines[i].trim();
    if (!line) continue;
    var subs = _splitBySentenceEnd_(line);
    for (var j = 0; j < subs.length; j++) {
      var seg = subs[j].trim();
      if (seg) segments.push(seg);
    }
  }
  return segments;
}

function _splitBySentenceEnd_(line) {
  var results = [], current = '', inLatex = false;
  for (var i = 0; i < line.length; i++) {
    var ch = line[i];
    if (ch === '$') { inLatex = !inLatex; current += ch; continue; }
    current += ch;
    if (!inLatex && (ch === '.' || ch === '?' || ch === '!')) {
      var next = (i + 1 < line.length) ? line[i + 1] : '';
      if (next === ' ' || next === '\t' || i + 1 >= line.length) {
        results.push(current.trim());
        current = '';
      }
    }
  }
  if (current.trim()) results.push(current.trim());
  return results;
}


/* =========================================================
 * Gemini LaTeX 이스케이프 수정
 * ========================================================= */

function _fixLatexEscapesInJson_(text) {
  if (!text) return text;
  try { JSON.parse(text); return text; } catch (e) { /* 수정 필요 */ }

  var s = text;
  for (var iter = 0; iter < 100; iter++) {
    try { JSON.parse(s); return s; } catch (e) {
      var msg = String(e.message || '');
      var posMatch = msg.match(/position\s+(\d+)/i) || msg.match(/column\s+(\d+)/i);
      if (!posMatch) return _fixAllLatexBackslashes_(s);

      var pos = Number(posMatch[1]);
      if (pos < 1 || pos >= s.length) return _fixAllLatexBackslashes_(s);

      var bsPos = -1;
      for (var k = pos; k >= Math.max(0, pos - 5); k--) {
        if (s[k] === '\\') { bsPos = k; break; }
      }
      if (bsPos === -1 && pos < s.length && s[pos] === '\\') bsPos = pos;
      if (bsPos === -1) return _fixAllLatexBackslashes_(s);

      s = s.substring(0, bsPos) + '\\' + s.substring(bsPos);
    }
  }
  return s;
}

function _fixAllLatexBackslashes_(text) {
  var PLACEHOLDER = '\x00DBL\x00';
  var s = text.split('\\\\').join(PLACEHOLDER);
  s = s.replace(/\\([a-zA-Z])/g, function(match, ch) { return '\\\\' + ch; });
  s = s.replace(/\\\\\\\\([bnrft])(?![a-zA-Z])/g, '\\\\$1');
  s = s.split(PLACEHOLDER).join('\\\\');
  return s;
}


/* =========================================================
 * 유틸리티
 * ========================================================= */

function _clearOutput_(sh) {
  sh.getRange('B3').clearContent();
  var startRow = 21, lastRow = sh.getMaxRows(), lastCol = sh.getMaxColumns();
  var numRows = lastRow - startRow + 1;
  if (numRows > 0) sh.getRange(startRow, 1, numRows, lastCol).clearContent();
}

function _flatCol_(sh, range) {
  return sh.getRange(range).getValues().map(function(r) { return String(r[0] || '').trim(); });
}

function _countOccurrences_(text, sub) {
  if (!sub) return 0;
  var count = 0, pos = 0;
  while (true) {
    var idx = text.indexOf(sub, pos);
    if (idx === -1) break;
    count++; pos = idx + sub.length;
  }
  return count;
}

function _isWhitespaceOnlyEdit_(original, revised) {
  var o = String(original || ''), r = String(revised || '');
  return o.replace(/\s+/g, '') === r.replace(/\s+/g, '');
}

function _existsInAnyRef_(phrase, refs) {
  if (!phrase || !refs || refs.length === 0) return false;
  var norm = phrase.replace(/\s+/g, '');
  for (var i = 0; i < refs.length; i++) {
    if ((refs[i].text || '').replace(/\s+/g, '').indexOf(norm) !== -1) return true;
  }
  return false;
}

function _sanitizeSourceIndex_(x, refCount) {
  var n = Number(x);
  if (!isFinite(n) || n % 1 !== 0) return null;
  if (refCount <= 0 || n < 1 || n > refCount) return null;
  return n;
}

function _verifyOrFindSourceIndex_(claimedIdx, revised, refs) {
  if (!revised || !refs || refs.length === 0) return claimedIdx;
  var norm = revised.replace(/\s+/g, '');
  if (claimedIdx !== null && claimedIdx >= 1 && claimedIdx <= refs.length) {
    if ((refs[claimedIdx - 1].text || '').replace(/\s+/g, '').indexOf(norm) !== -1) return claimedIdx;
  }
  for (var i = 0; i < refs.length; i++) {
    if ((refs[i].text || '').replace(/\s+/g, '').indexOf(norm) !== -1) return i + 1;
  }
  return claimedIdx;
}


/* =========================================================
 * prompt 시트 읽기
 * ========================================================= */

function _getPromptValue_(key) {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(PROMPT_SHEET_NAME);
  if (!sh) throw new Error('prompt 시트를 찾을 수 없습니다.');

  var lastRow = sh.getLastRow();
  if (lastRow < 2) throw new Error('prompt 시트에 데이터가 없습니다.');

  var values = sh.getRange(2, 1, lastRow - 1, 3).getValues();
  for (var i = 0; i < values.length; i++) {
    var k = String(values[i][0] || '').trim();
    var v = String(values[i][1] || '').trim();
    var enabled = values[i][2];
    if (k === key && (enabled === true || String(enabled).toUpperCase() === 'TRUE')) {
      if (!v) throw new Error('prompt 시트의 ' + key + ' 값이 비어있습니다.');
      return v;
    }
  }
  throw new Error('prompt 시트에서 활성화된 key를 찾지 못했습니다: ' + key);
}
/*******************************************************
 * 문항검토: 기출 스타일 문장 다듬기 (멀티 LLM: Claude / GPT / Gemini)
 *
 * provider별 API 호출을 통합하여 동일한 후처리 파이프라인 공유.
 * 배치 코드에서 options.provider = 'claude'|'gpt'|'gemini' 로 선택.
 *
 * [prompt 시트 키]
 *   SYSTEM_KICE_CLAUDE   — Claude용 system prompt
 *   SYSTEM_KICE_GPT      — GPT용 system prompt
 *   SYSTEM_KICE_GEMINI   — Gemini용 system prompt
 *   OUTPUT_FORMAT_KICE   — 출력 형식 (공통, 선택)
 *
 * [스크립트 속성]
 *   ANTHROPIC_API_KEY    — Claude
 *   OPENAI_API_KEY       — GPT
 *   GEMINI_API_KEY       — Gemini
 *******************************************************/

var SHEET_NAME = '문항검토';
var RULE_SHEET_NAME = 'Rule';
var PROMPT_SHEET_NAME = 'prompt';

// provider별 prompt 시트 키 매핑
var PROVIDER_PROMPT_KEYS = {
  claude: 'SYSTEM_KICE_CLAUDE',
  gpt:    'SYSTEM_KICE_GPT',
  gemini: 'SYSTEM_KICE_GEMINI'
};

// provider별 API 키 스크립트 속성명
var PROVIDER_API_KEY_NAMES = {
  claude: 'ANTHROPIC_API_KEY',
  gpt:    'OPENAI_API_KEY',
  gemini: 'GEMINI_API_KEY'
};

// 공통 출력 형식 키
var PROMPT_KEY_OUTPUT_FORMAT = 'OUTPUT_FORMAT_KICE';

// 재시도 설정
var MAX_RETRIES = 4;
var INITIAL_WAIT_MS = 5000;


/* ===========================
 * 메인 함수
 * =========================== */

/**
 * @param {Object} options
 *   - provider: 'claude' | 'gpt' | 'gemini' (기본: 'claude')
 *   - openViewer: boolean (기본: true)
 */
function review_rewriteToKiceStyle(options) {
  options = options || {};
  var provider = String(options.provider || 'claude').toLowerCase();
  var openViewer = (options.openViewer !== false);

  // provider 유효성 검증
  if (!PROVIDER_PROMPT_KEYS[provider]) {
    throw new Error('지원하지 않는 provider: ' + provider + ' (claude/gpt/gemini 중 선택)');
  }

  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) throw new Error('시트를 찾을 수 없습니다: ' + SHEET_NAME);

  var apiKeyName = PROVIDER_API_KEY_NAMES[provider];
  var apiKey = PropertiesService.getScriptProperties().getProperty(apiKeyName);
  if (!apiKey) throw new Error('스크립트 속성 ' + apiKeyName + '가 비어있습니다.');

  // 0) 기존 출력 초기화
  _clearOutput_(sh);

  // 1) 입력 읽기
  var targetRaw = String(sh.getRange('B2').getValue() || '').trim();
  var chapter = String(sh.getRange('C2').getValue() || '').trim();

  if (!targetRaw) {
    sh.getRange('B3').setValue('수정구절 없음');
    sh.getRange('D21').setValue('수정 구절 없음');
    if (openViewer) _selectResultRangeAndOpenViewer_(ss, sh, 21, 21);
    return;
  }

  // 2) Rule 로드 + 선치환
  var rules = _loadRules_();
  var preRule = _applyRulesWithLog_(targetRaw, rules);
  var target = preRule.text;
  var ruleEditsAll = preRule.edits.slice();

  // 3) 관련 기출문항 로드
  var sources = _flatten_(sh.getRange('A6:A15').getValues());
  var imageLinks = _flatten_(sh.getRange('C6:C15').getValues());
  var refsText = _flatten_(sh.getRange('D6:D15').getValues());

  var refs = [];
  for (var i = 0; i < refsText.length; i++) {
    if (refsText[i]) {
      refs.push({ source: sources[i] || '', imageLink: imageLinks[i] || '', text: refsText[i] });
    }
  }

  // 4) provider별 프롬프트 구성 + LLM 호출
  var systemPrompt = _buildSystemPrompt_(provider, rules);
  var userMessage = _buildUserMessage_(target, chapter, refs);
  var parsed = _callAndParseLLM_(provider, apiKey, systemPrompt, userMessage);

  // 5) 수정 구절 필터링 + source_index 검증 (공통)
  var editsRaw = Array.isArray(parsed.edits) ? parsed.edits : [];
  var llmEdits = editsRaw
    .filter(function(e) { return !_isWhitespaceOnlyEdit_(e.original, e.revised); })
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
        reason: String(e.reason || '').trim(),
        _kind: provider.toUpperCase()
      };
    });

  // 6) 역방향 수정 되돌리기 목록
  var rejectedEdits = editsRaw
    .filter(function(e) { return !_isWhitespaceOnlyEdit_(e.original, e.revised); })
    .filter(function(e) {
      var orig = String(e.original || '').trim();
      var rev = String(e.revised || '').trim();
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
      var rejRevised = String(rej.revised || '').trim();
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
      ruleEditsAll.push({ from: e.from, to: e.to, count: e.count, note: e.note, sample: e.sample, phase: 'POST' });
    });
  }

  sh.getRange('B3').setValue(finalFull);

  // 8) 출력 (D21 이하)
  var startRow = 21;
  var maxRows = 200;

  var ruleOut = ruleEditsAll.map(function(e) {
    return { _kind: 'RULE', source: 'RULE', imageLink: '', text: _formatRuleEdit_(e) };
  });

  var llmOut = llmEdits.map(function(e) {
    var refIdx = (e.source_index !== null && e.source_index >= 1 && e.source_index <= refs.length) ? (e.source_index - 1) : -1;
    return {
      _kind: e._kind,
      source: refIdx >= 0 ? (refs[refIdx].source || '') : '',
      imageLink: refIdx >= 0 ? (refs[refIdx].imageLink || '') : '',
      text: '[[원본]] ' + e.original + '\n[[수정]] ' + e.revised + '\n[[이유]] ' + e.reason
    };
  });

  var outItems = ruleOut.concat(llmOut);

  if (outItems.length === 0 || (!parsed.has_edits && ruleOut.length === 0)) {
    sh.getRange('D21').setValue('수정 구절 없음');
    if (openViewer) _selectResultRangeAndOpenViewer_(ss, sh, 21, 21);
    return;
  }

  var rowCount = Math.min(outItems.length, maxRows);
  var outA = [], outC = [], outD = [];
  for (var j = 0; j < rowCount; j++) {
    var it = outItems[j];
    outA.push([it.source || '']);
    outC.push([it.imageLink || '']);
    outD.push([it.text || '']);
  }

  sh.getRange(startRow, 1, rowCount, 1).setValues(outA);
  sh.getRange(startRow, 3, rowCount, 1).setValues(outC);
  sh.getRange(startRow, 4, rowCount, 1).setValues(outD);

  var lastRow = startRow + rowCount - 1;
  if (openViewer) _selectResultRangeAndOpenViewer_(ss, sh, startRow, lastRow);
}


/* ===========================
 * 프롬프트 구성
 * =========================== */

/**
 * provider별 system prompt 구성
 * prompt 시트에서 provider 전용 키를 읽고, Rule + 출력 형식을 합친다.
 */
function _buildSystemPrompt_(provider, rules) {
  var promptKey = PROVIDER_PROMPT_KEYS[provider];
  var basePrompt = _getPromptValue_(promptKey);

  var lines = [];
  lines.push(basePrompt);
  lines.push('');

  // 강제 치환 규칙 (Rule) — 동적 삽입
  if (Array.isArray(rules) && rules.length > 0) {
    lines.push('<forced_rules>');
    lines.push('아래는 강제 수정 규칙이다. 이미 원문에 선적용되어 있다.');
    lines.push('- "to" 표현이 최종 결과(rewritten_full)에 그대로 유지되어야 한다. 절대 변경하지 마라.');
    lines.push('- "from" 표현이 최종 결과에 남아 있으면 안 된다.');
    lines.push('- 문법이 어색해지면 Rule의 to는 건드리지 말고 주변 문장만 다듬어라.');
    rules.forEach(function(r, i) {
      var note = r.note ? ' (note: ' + r.note + ')' : '';
      lines.push('(' + (i + 1) + ') "' + r.from + '" → "' + r.to + '"' + note);
    });
    lines.push('</forced_rules>');
    lines.push('');
  }

  // 출력 형식
  var outputFormat = _getPromptValueOrDefault_(PROMPT_KEY_OUTPUT_FORMAT, null);
  if (outputFormat) {
    lines.push(outputFormat);
  } else {
    lines.push(_defaultOutputFormat_());
  }

  return lines.join('\n');
}

/** 기본 출력 형식 */
function _defaultOutputFormat_() {
  var lines = [];
  lines.push('<work_process>');
  lines.push('아래 절차를 반드시 따라 검토하라:');
  lines.push('');
  lines.push('1단계: 구절별 대조');
  lines.push('- <target>에 S1, S2, S3, ... 로 번호가 매겨진 구절이 있다.');
  lines.push('- 각 구절(S1~Sn)을 하나씩 꺼내어, <references>의 모든 기출문항(ref 1~N)과 단어·표현 단위로 대조하라.');
  lines.push('- 모든 구절을 빠짐없이 점검하라. 어떤 구절도 건너뛰지 마라.');
  lines.push('');
  lines.push('2단계: 수정 판단');
  lines.push('- 대조 결과, 기출문항의 표현과 다른 부분만 수정 대상이다.');
  lines.push('- 수정 방향: "기출에 없는 표현 → 기출에 있는 표현" (역방향 금지)');
  lines.push('- 기출문항에 이미 등장하는 표현은 절대 수정하지 않는다.');
  lines.push('- 수정할 것이 없으면 억지로 수정하지 않는다.');
  lines.push('');
  lines.push('3단계: 결과 출력');
  lines.push('</work_process>');
  lines.push('');
  lines.push('<output_format>');
  lines.push('반드시 아래 JSON 구조만 출력하라. JSON 외의 텍스트는 일절 포함하지 마라.');
  lines.push('{');
  lines.push('  "has_edits": boolean,');
  lines.push('  "rewritten_full": "수정된 문항 전문 (수정 없으면 원문 그대로)",');
  lines.push('  "edits": [');
  lines.push('    {');
  lines.push('      "source_index": revised 구절이 글자 그대로 포함된 기출문항 번호(정수, 1~N),');
  lines.push('      "original": "원문에서 수정 대상 구절",');
  lines.push('      "revised": "수정된 구절",');
  lines.push('      "reason": "수정 이유 (간결하게)"');
  lines.push('    }');
  lines.push('  ]');
  lines.push('}');
  lines.push('');
  lines.push('- 원문과 rewritten_full을 공백 제거 후 비교하여 동일하면 has_edits = false, edits = []로 반환한다.');
  lines.push('- 띄어쓰기만 변경된 구절은 edits에 포함하지 않는다.');
  lines.push('');
  lines.push('[자기 점검] edits 배열의 각 항목에 대해 출력 전에 반드시 확인하라:');
  lines.push('1. original이 기출문항에 이미 등장하는 표현이면 → 해당 수정을 취소하라.');
  lines.push('2. revised가 기출문항 어디에도 글자 그대로 등장하지 않으면 → 해당 수정을 취소하라.');
  lines.push('3. 수정 방향은 반드시 "기출에 없는 표현 → 기출에 있는 표현"이어야 한다.');
  lines.push('</output_format>');
  return lines.join('\n');
}

/** User Message — 구절별 분해 포함 */
function _buildUserMessage_(target, chapter, refs) {
  var lines = [];

  if (chapter) lines.push('<chapter>' + chapter + '</chapter>');
  lines.push('');

  var segments = _splitToSegments_(target);
  lines.push('<target>');
  lines.push('[원문 전체]');
  lines.push(target);
  lines.push('');
  lines.push('[구절별 분해] (각 구절을 기출문항과 빠짐없이 대조하라)');
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


/* ===========================
 * LLM 호출 + JSON 파싱 (통합)
 * =========================== */

/**
 * provider별 API 호출 → JSON 파싱까지 통합
 * 반환: { has_edits, rewritten_full, edits[] }
 */
function _callAndParseLLM_(provider, apiKey, systemPrompt, userMessage) {
  var jsonText;

  if (provider === 'claude') {
    jsonText = _callClaude_(apiKey, systemPrompt, userMessage);
  } else if (provider === 'gpt') {
    jsonText = _callGPT_(apiKey, systemPrompt, userMessage);
  } else if (provider === 'gemini') {
    jsonText = _callGemini_(apiKey, systemPrompt, userMessage);
  } else {
    throw new Error('알 수 없는 provider: ' + provider);
  }

  // LaTeX 백슬래시 이스케이프 수정 (Gemini가 JSON 안에서 \frac 등을 이중 이스케이프 안 하는 문제)
  jsonText = _fixLatexEscapesInJson_(jsonText);

  var parsed;
  try {
    parsed = JSON.parse(jsonText);
  } catch (e) {
    var preview = jsonText.length > 500
      ? jsonText.substring(0, 300) + '\n...(중략)...\n' + jsonText.substring(jsonText.length - 200)
      : jsonText;
    var codes = '';
    for (var ci = 0; ci < Math.min(5, jsonText.length); ci++) {
      codes += jsonText.charCodeAt(ci) + ' ';
    }
    throw new Error(provider.toUpperCase() + ' JSON 파싱 실패 (len=' + jsonText.length + ', first5codes=' + codes.trim() + '):\n' + preview);
  }

  return parsed;
}


/* ===========================
 * Claude API
 * =========================== */

function _callClaude_(apiKey, systemPrompt, userMessage) {
  var body = {
    model: 'claude-sonnet-4-20250514',
    max_tokens: 4096,
    temperature: 0.15,
    system: systemPrompt,
    messages: [
      { role: 'user', content: userMessage },
      { role: 'assistant', content: '{' }
    ],
    stop_sequences: ['\n}']
  };

  var resp = _httpPostWithRetry_(
    'https://api.anthropic.com/v1/messages',
    body,
    { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' }
  );

  // prefill '{' + 응답 + stop '}' 복원
  var rawText = '';
  if (resp && Array.isArray(resp.content)) {
    for (var i = 0; i < resp.content.length; i++) {
      if (resp.content[i].type === 'text') {
        rawText = resp.content[i].text || '';
        break;
      }
    }
  }
  return '{' + rawText.trim() + '}';
}


/* ===========================
 * OpenAI GPT API (Chat Completions)
 * =========================== */

function _callGPT_(apiKey, systemPrompt, userMessage) {
  var body = {
    model: 'gpt-5.2',
    max_completion_tokens: 4096,
    temperature: 0.15,
    response_format: { type: 'json_object' },
    messages: [
      { role: 'system', content: systemPrompt },
      { role: 'user', content: userMessage }
    ]
  };

  var resp = _httpPostWithRetry_(
    'https://api.openai.com/v1/chat/completions',
    body,
    { 'Authorization': 'Bearer ' + apiKey }
  );

  // Chat Completions 응답 파싱
  if (resp && resp.choices && resp.choices.length > 0) {
    var msg = resp.choices[0].message;
    if (msg && msg.content) return msg.content.trim();
  }
  throw new Error('GPT 응답에서 content를 추출하지 못했습니다.');
}


/* ===========================
 * Google Gemini API
 * =========================== */

function _callGemini_(apiKey, systemPrompt, userMessage) {
  // 1. 요청 바디 구성
  // 400 에러를 방지하기 위해 표준 규격인 스네이크 케이스(snake_case)를 우선 사용합니다.
  var body = {
    contents: [
      { 
        role: 'user', 
        parts: [{ text: userMessage }] 
      }
    ],
    systemInstruction: {
      parts: [{ text: systemPrompt }]
    },
    generationConfig: {
      // 수학 문제의 엄밀한 교정을 위해 온도를 낮추어 일관성을 확보합니다.
      temperature: 0.1, 
      maxOutputTokens: 65536,
      responseMimeType: 'application/json'
      // 400 에러의 원인이 된 thinkingConfig는 제거하여 안정성을 최우선으로 합니다.
    }
  };

  // 2. 모델 URL 설정 (추천하신 gemini-2.5-flash 사용)
  var url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' + apiKey;

  // 3. API 호출
  var resp = _httpPostWithRetry_(url, body, {});

  // 4. Gemini 응답 파싱
  if (resp && resp.candidates && resp.candidates.length > 0) {
    var candidate = resp.candidates[0];

    // 안전 필터 및 종료 사유 확인
    if (candidate.finishReason === 'SAFETY' || candidate.finishReason === 'BLOCKED') {
      throw new Error('Gemini: 안전 필터에 의해 응답이 차단되었습니다.');
    }
    
    if (candidate.finishReason === 'MAX_TOKENS') {
      throw new Error('Gemini: 출력 토큰 제한을 초과했습니다.');
    }

    var parts = candidate.content && candidate.content.parts;
    if (parts && parts.length > 0) {
      // 여러 파트 중 실제 텍스트(JSON)가 포함된 파트 검색
      for (var i = 0; i < parts.length; i++) {
        if (parts[i].text) {
          var text = parts[i].text.trim();
          
          // Markdown 코드 블록 제거 및 JSON 객체 순수 추출
          text = text.replace(/^```json\s*/i, '').replace(/\s*```$/i, '');
          
          var firstBrace = text.indexOf('{');
          var lastBrace = text.lastIndexOf('}');
          if (firstBrace !== -1 && lastBrace > firstBrace) {
            return text.substring(firstBrace, lastBrace + 1);
          }
          return text; // 중괄호가 없는 경우 전체 텍스트 반환
        }
      }
    }
  }

  // 응답 실패 시 상세 에러 메시지 구성
  var errorDetail = '';
  if (resp && resp.promptFeedback && resp.promptFeedback.blockReason) {
    errorDetail = ' (사유: ' + resp.promptFeedback.blockReason + ')';
  }
  throw new Error('Gemini 응답에서 유효한 텍스트를 추출하지 못했습니다.' + errorDetail);
}


/* ===========================
 * HTTP 공통 (지수 백오프 재시도)
 * =========================== */

function _httpPostWithRetry_(url, body, extraHeaders) {
  var headers = {};
  for (var key in extraHeaders) {
    headers[key] = extraHeaders[key];
  }

  var options = {
    method: 'post',
    contentType: 'application/json',
    headers: headers,
    muteHttpExceptions: true,
    payload: JSON.stringify(body)
  };

  var lastCode = 0;
  var lastText = '';

  for (var attempt = 0; attempt <= MAX_RETRIES; attempt++) {
    var res = UrlFetchApp.fetch(url, options);
    lastCode = res.getResponseCode();
    lastText = res.getContentText();

    if (lastCode >= 200 && lastCode < 300) {
      return JSON.parse(lastText);
    }

    var retryable = (lastCode === 429 || lastCode === 529 || lastCode === 500 || lastCode === 503);
    if (!retryable || attempt >= MAX_RETRIES) break;

    var waitMs = INITIAL_WAIT_MS * Math.pow(2, attempt);
    Utilities.sleep(waitMs);
  }

  throw new Error('API 오류 (' + lastCode + '):\n' + lastText.substring(0, 500));
}


/* ===========================
 * 구절 분해
 * =========================== */

function _splitToSegments_(text) {
  if (!text) return [];
  var rawLines = text.split(/\n/);
  var segments = [];
  for (var i = 0; i < rawLines.length; i++) {
    var line = rawLines[i].trim();
    if (!line) continue;
    var subSegs = _splitBySentenceEnd_(line);
    for (var j = 0; j < subSegs.length; j++) {
      var seg = subSegs[j].trim();
      if (seg) segments.push(seg);
    }
  }
  return segments;
}

function _splitBySentenceEnd_(line) {
  var results = [];
  var current = '';
  var inLatex = false;
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


/* ===========================
 * Rule: 로드 + 적용
 * =========================== */

function _loadRules_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(RULE_SHEET_NAME);
  if (!sh) return [];
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  var values = sh.getRange(2, 1, lastRow - 1, 4).getValues();
  var rules = [];
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var from = String(row[0] || '').trim();
    var to = String(row[1] || '').trim();
    var enabled = row[2];
    var note = String(row[3] || '').trim();
    var isEnabled = (enabled === true) || (String(enabled).toUpperCase() === 'TRUE');
    if (!isEnabled || !from || from === to) continue;
    if (to && to.indexOf(from) !== -1) continue;
    rules.push({ from: from, to: to, note: note });
  }
  rules.sort(function(a, b) { return b.from.length - a.from.length; });
  return rules;
}

function _applyRulesWithLog_(text, rules) {
  var src = String(text || '');
  if (!src || !Array.isArray(rules) || rules.length === 0) return { text: src, edits: [] };
  var cur = src;
  var edits = [];
  for (var i = 0; i < rules.length; i++) {
    var r = rules[i];
    var from = String(r.from || '');
    var to = String(r.to || '');
    if (!from) continue;
    var count = _countOccurrences_(cur, from);
    if (count <= 0) continue;
    var idx = cur.indexOf(from);
    var sample = _contextSample_(cur, idx, from.length, 20);
    cur = cur.split(from).join(to);
    edits.push({ from: from, to: to, count: count, note: String(r.note || ''), sample: sample });
  }
  return { text: cur, edits: edits };
}

function _countOccurrences_(text, sub) {
  if (!sub) return 0;
  var count = 0, pos = 0;
  while (true) { var idx = text.indexOf(sub, pos); if (idx === -1) break; count++; pos = idx + sub.length; }
  return count;
}

function _contextSample_(text, idx, len, pad) {
  if (idx < 0) return '';
  return text.slice(Math.max(0, idx - pad), Math.min(text.length, idx + len + pad));
}

function _formatRuleEdit_(e) {
  var from = String(e.from || ''), to = String(e.to || '');
  var count = Number(e.count || 0), note = String(e.note || '').trim();
  var phase = e.phase ? String(e.phase) : 'PRE';
  var reason = 'Rule(' + phase + '): "' + from + '" → "' + to + '" | ' + count + '회';
  if (note) reason += ' | note: ' + note;
  return '[[원본]] ' + from + '\n[[수정]] ' + to + '\n[[이유]] ' + reason;
}

/**
 * Gemini LaTeX 이스케이프 수정
 *
 * Gemini가 JSON 문자열 안에서 LaTeX 백슬래시(\frac, \mathrm 등)를
 * 이중 이스케이프하지 않아 JSON 파싱이 실패하는 문제를 수정.
 *
 * 전략: JSON.parse를 시도하고, 실패하면 에러 위치 근처의 잘못된
 * 이스케이프를 찾아 수정한 뒤 재시도. 최대 100회 반복.
 */
function _fixLatexEscapesInJson_(text) {
  if (!text) return text;

  // 먼저 그대로 파싱 시도
  try { JSON.parse(text); return text; } catch (e) { /* 수정 필요 */ }

  var s = text;
  var MAX_ITER = 100;

  for (var iter = 0; iter < MAX_ITER; iter++) {
    try {
      JSON.parse(s);
      return s; // 성공
    } catch (e) {
      // 에러 메시지에서 위치 추출: "position N" 또는 "column N"
      var msg = String(e.message || '');
      var posMatch = msg.match(/position\s+(\d+)/i) || msg.match(/column\s+(\d+)/i);
      if (!posMatch) {
        // 위치를 찾지 못하면 전체 치환 시도 후 반환
        return _fixAllLatexBackslashes_(s);
      }

      var pos = Number(posMatch[1]);
      if (pos < 1 || pos >= s.length) return _fixAllLatexBackslashes_(s);

      // 에러 위치 앞뒤에서 잘못된 백슬래시 이스케이프 찾기
      // 에러 위치에서 거꾸로 가장 가까운 백슬래시 찾기
      var bsPos = -1;
      for (var k = pos; k >= Math.max(0, pos - 5); k--) {
        if (s[k] === '\\') { bsPos = k; break; }
      }
      // 앞에서 못 찾으면 에러 위치 자체가 백슬래시인 경우
      if (bsPos === -1 && pos < s.length && s[pos] === '\\') bsPos = pos;
      if (bsPos === -1) return _fixAllLatexBackslashes_(s);

      // 해당 백슬래시를 이중으로 만들기
      s = s.substring(0, bsPos) + '\\' + s.substring(bsPos);
    }
  }

  return s;
}

/**
 * 최후의 수단: 알려진 LaTeX 명령 패턴을 전부 이중 이스케이프
 */
function _fixAllLatexBackslashes_(text) {
  // JSON 유효 이스케이프: \", \\, \/, \b, \f, \n, \r, \t, \uXXXX
  // 그 외 \+알파벳은 모두 LaTeX으로 간주
  // 단, 이미 \\인 경우 보호

  var PLACEHOLDER = '\x00DBL\x00';
  var s = text.split('\\\\').join(PLACEHOLDER);

  // \+알파벳(소문자만, 1글자 이상)이면서 JSON 유효 이스케이프가 아닌 것
  // JSON 유효 1글자 이스케이프: b, f, n, r, t (u는 \uXXXX로 별도)
  // 하지만 \frac에서 \f도 LaTeX이므로, 뒤에 알파벳이 더 이어지면 LaTeX
  s = s.replace(/\\([a-zA-Z])/g, function(match, ch) {
    var idx = match.length; // 매치 후 위치
    // 실제 원본에서 이 매치 다음 문자를 볼 수 없으므로
    // 1글자 JSON 이스케이프(\n, \t, \r, \b, \f)와 LaTeX 명령을 구분하기 어려움
    // → 일단 전부 이중 이스케이프하고, JSON 유효 이스케이프는 복원
    return '\\\\' + ch;
  });

  // JSON 유효 이스케이프 복원: \\\\n → \\n, \\\\t → \\t 등
  // 단, 뒤에 알파벳이 바로 이어지면 LaTeX이므로 복원하지 않음
  s = s.replace(/\\\\\\\\([bnrft])(?![a-zA-Z])/g, '\\\\$1');

  s = s.split(PLACEHOLDER).join('\\\\');
  return s;
}


/* ===========================
 * 유틸리티
 * =========================== */

function _clearOutput_(sh) {
  sh.getRange('B3').clearContent();
  var startRow = 21, lastRow = sh.getMaxRows(), lastCol = sh.getMaxColumns();
  var numRows = lastRow - startRow + 1;
  if (numRows > 0) sh.getRange(startRow, 1, numRows, lastCol).clearContent();
}

function _flatten_(arr2d) {
  var result = [];
  for (var i = 0; i < arr2d.length; i++) result.push(String(arr2d[i][0] || '').trim());
  return result;
}

function _isWhitespaceOnlyEdit_(original, revised) {
  var o = String(original || ''), r = String(revised || '');
  if (o === r) return true;
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
  var revisedNorm = revised.replace(/\s+/g, '');
  if (claimedIdx !== null && claimedIdx >= 1 && claimedIdx <= refs.length) {
    if ((refs[claimedIdx - 1].text || '').replace(/\s+/g, '').indexOf(revisedNorm) !== -1) return claimedIdx;
  }
  for (var i = 0; i < refs.length; i++) {
    if ((refs[i].text || '').replace(/\s+/g, '').indexOf(revisedNorm) !== -1) return i + 1;
  }
  return claimedIdx;
}

function _selectResultRangeAndOpenViewer_(ss, sh, startRow, lastRow) {
  ss.setActiveSheet(sh);
  var rg = sh.getRange(startRow, 4, Math.max(1, lastRow - startRow + 1), 1);
  rg.activate(); ss.setActiveRange(rg);
  try {
    if (typeof LV !== 'undefined' && LV && typeof LV.openDialog === 'function') LV.openDialog();
    else if (typeof lv_openDialog === 'function') lv_openDialog();
    else ss.toast('LatexViewer(LV) 함수를 찾지 못했습니다.', '안내', 3);
  } catch (e) { ss.toast('Viewer 오픈 실패: ' + (e.message || e), '오류', 5); }
}


/* ===========================
 * prompt 시트 읽기
 * =========================== */

function _getPromptValue_(key) {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(PROMPT_SHEET_NAME);
  if (!sh) throw new Error('prompt 시트를 찾을 수 없습니다: ' + PROMPT_SHEET_NAME);
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

function _getPromptValueOrDefault_(key, defaultValue) {
  try { return _getPromptValue_(key); } catch (e) { return defaultValue; }
}
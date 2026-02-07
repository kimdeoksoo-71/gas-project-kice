/*******************************************************
 * 문항검토: 기출 스타일 문장 다듬기 (GPT-5.2 / Responses API)
 * + 출력 후 Viewer(Dialog) 자동 호출
 *
 * ✅ 추가 요구(반영됨):
 * - Rule 시트(A:from,B:to,C:enabled,D:note)에 따른 강제 치환(선치환 + 후치환)
 * - Rule로 강제 수정된 구절도 D21 이하에 함께 출력(무엇을 수정했는지 확인)
 * - Rule 치환이 문법을 깨면 GPT가 주변 문장으로 자연스럽게 다듬도록 프롬프트에 명시
 * - "수정 구절" 출력 시, source(A6:A15)에 대응하는 이미지 링크(C6:C15)를 C열(21행~)에 함께 기록
 *******************************************************/

const SHEET_NAME = '문항검토';
const PROMPT_SHEET_NAME = 'prompt';
const PROMPT_KEY_INSTRUCTIONS = 'INSTRUCTIONS_KICE_REWRITE';
const RULE_SHEET_NAME = 'Rule';

function review_rewriteToKiceStyle_gpt52(options) {
  options = options || {};
  const openViewer = (options.openViewer !== false); // 기본 true
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) throw new Error(`시트를 찾을 수 없습니다: ${SHEET_NAME}`);

  const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!apiKey) throw new Error('스크립트 속성 OPENAI_API_KEY가 비어있습니다.');

  // 0) 실행 시작 시: B3 + 21행 이하 전체(모든 열) 기존 내용 지우기
  _clearOutput_(sh);

  // 1) 입력 읽기
  const targetRaw = String(sh.getRange('B2').getValue() ?? '').trim();
  const chapter = String(sh.getRange('C2').getValue() ?? '').trim();

  if (!targetRaw) {
    sh.getRange('B3').setValue('수정구절 없음');
    sh.getRange('D21').setValue('수정 구절 없음');
    if (openViewer) _selectResultRangeAndOpenViewer_(ss, sh, 21, 21);
    return;
  }

  // 1-1) Rule 로드 + 선치환 (원문)
  const rules = _loadRules_(); // enabled만, from 길이 내림차순 정렬
  const preRule = _applyRulesWithLog_(targetRaw, rules);
  const target = preRule.text;           // GPT에게 줄 문항(선치환 반영)
  const ruleEditsAll = [...preRule.edits]; // Rule 로그 누적(선/후치환)

  // ✅ 관련 기출문항: source(A) + imageLink(C) + text(D)
  const sources = sh.getRange('A6:A15').getValues().flat().map(v => String(v ?? '').trim());
  const imageLinks = sh.getRange('C6:C15').getValues().flat().map(v => String(v ?? '').trim());
  const refsText = sh.getRange('D6:D15').getValues().flat().map(v => String(v ?? '').trim());

  // 유효한 관련문항만 추림
  const refs = [];
  for (let i = 0; i < refsText.length; i++) {
    if (refsText[i]) {
      refs.push({
        // refs에서의 1-based index가 source_index로 사용됨
        source: sources[i] || '',
        imageLink: imageLinks[i] || '',
        text: refsText[i],
      });
    }
  }

  // 2) 프롬프트 구성 (Rule 포함)
  const prompt = _buildPrompt_(target, chapter, refs, rules);

  // 3) GPT 호출 (Structured Outputs)
  const schema = _buildJsonSchema_();
  const responseJson = _callOpenAIResponses_(apiKey, {
    model: 'gpt-5.2',
    instructions: _buildInstructions_(),
    input: [{ role: 'user', content: prompt }],
    text: {
      format: { type: 'json_schema', name: 'rewrite_result', strict: true, schema }
    },
    temperature: 0.2,
    max_output_tokens: 2500
  });

  // 4) 모델 출력(JSON) 파싱
  const outText = _extractOutputText_(responseJson);
  if (!outText) throw new Error('모델 출력 텍스트를 추출하지 못했습니다.');

  let data;
  try {
    data = JSON.parse(outText);
  } catch (e) {
    throw new Error('모델 출력이 JSON 파싱에 실패했습니다.\n' + outText);
  }

  // 5) GPT 수정 구절 필터링 (띄어쓰기/줄바꿈-only 수정 구절 제거)
  const editsRaw = Array.isArray(data.edits) ? data.edits : [];
  const gptEdits = editsRaw
    .filter(e => !_isWhitespaceOnlyEdit_(e.original, e.revised))
    .map(e => ({
      source_index: _sanitizeSourceIndex_(e.source_index, refs.length),
      original: String(e.original ?? '').trim(),
      revised: String(e.revised ?? '').trim(),
      reason: String(e.reason ?? '').trim(),
      _kind: 'GPT'
    }));

  // 6) rewritten_full 결정 + 후치환(Rule) (최종 결과에도 Rule 남아있으면 무조건 제거)
  let rewrittenFromModel = String(data.rewritten_full ?? '').trim();

  // ✅ [추가] rewritten_full이 띄어쓰기만 바뀐 경우 → 원문(target)으로 되돌리기
  if (
    rewrittenFromModel &&
    rewrittenFromModel.replace(/\s+/g, '') === target.replace(/\s+/g, '')
  ) {
    rewrittenFromModel = target;
  }

  const baseFull = rewrittenFromModel || target || targetRaw;


  const postRule = _applyRulesWithLog_(baseFull, rules);
  const finalFull = postRule.text;
  if (postRule.edits.length) ruleEditsAll.push(...postRule.edits.map(e => ({ ...e, phase: 'POST' })));

  // 최종 결과 저장(B3): Rule 후치환까지 끝난 값
  sh.getRange('B3').setValue(finalFull);

  // 7) 출력(D21 이하): Rule edits + GPT edits 함께 출력
  const startRow = 21;
  const maxRows = 200;

  // Rule 로그를 출력용 엔트리로 변환 (A="RULE", C="", D=상세)
  const ruleOut = (ruleEditsAll || []).map(e => ({
    _kind: 'RULE',
    source: 'RULE',
    imageLink: '',
    text: _formatRuleEdit_(e)
  }));

  // GPT 로그를 출력용 엔트리로 변환 (기존 매핑 유지)
  const gptOut = (gptEdits || []).map(e => {
    const refIdx = (e.source_index >= 1 && e.source_index <= refs.length) ? (e.source_index - 1) : 0;
    return {
      _kind: 'GPT',
      source: refs[refIdx]?.source || '',
      imageLink: refs[refIdx]?.imageLink || '',
      text: `[[원본]] ${e.original}\n[[수정]] ${e.revised}\n[[이유]] ${e.reason}`
    };
  });

  const outItems = [...ruleOut, ...gptOut];

  // 8) "수정구절 없음" 처리: (Rule/GPT 모두 없으면)
  if (outItems.length === 0 || (!data.has_edits && ruleOut.length === 0)) {
    sh.getRange('D21').setValue('수정 구절 없음');
    if (openViewer) _selectResultRangeAndOpenViewer_(ss, sh, 21, 21);
    return;
  }

  const rowCount = Math.min(outItems.length, maxRows);
  const outA = new Array(rowCount);
  const outC = new Array(rowCount);
  const outD = new Array(rowCount);

  for (let i = 0; i < rowCount; i++) {
    const it = outItems[i];
    outA[i] = [it.source || ''];
    outC[i] = [it.imageLink || ''];
    outD[i] = [it.text || ''];
  }

  sh.getRange(startRow, 1, rowCount, 1).setValues(outA); // A열
  sh.getRange(startRow, 3, rowCount, 1).setValues(outC); // C열(이미지 링크; Rule은 빈칸)
  sh.getRange(startRow, 4, rowCount, 1).setValues(outD); // D열

  // 9) 출력 완료 후 viewer 자동 호출 (D열 결과 구간 선택)
  const lastRow = startRow + rowCount - 1;
  if (openViewer) _selectResultRangeAndOpenViewer_(ss, sh, startRow, lastRow);
}

/** =========================
 * Rule: load + apply
 * ========================= */

/**
 * Rule 시트(A:from,B:to,C:enabled,D:note), 헤더 1행
 * enabled만 로드, from 길이 내림차순 정렬(겹침 완화)
 */
function _loadRules_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(RULE_SHEET_NAME);
  if (!sh) return []; // Rule 없으면 그냥 패스

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  const values = sh.getRange(2, 1, lastRow - 1, 4).getValues(); // A~D
  const rules = [];

  for (const row of values) {
    const from = String(row[0] ?? '').trim();
    const to = String(row[1] ?? '').trim();
    const enabled = row[2];
    const note = String(row[3] ?? '').trim();

    const isEnabled = (enabled === true) || (String(enabled).toUpperCase() === 'TRUE');
    if (!isEnabled) continue;
    if (!from) continue;
    if (from === to) continue;

    // 무한치환 방지: to 안에 from이 들어가면 위험(단순 버전에서는 skip)
    if (to && to.includes(from)) continue;

    rules.push({ from, to, note });
  }

  // from 길이 내림차순(긴 구절 먼저 치환)
  rules.sort((a, b) => (b.from.length - a.from.length));
  return rules;
}

/**
 * 단순 포함 매칭 치환 + 로그
 * - split/join으로 "모든 등장" 치환
 * - edits: from/to/count/note/sample(문맥)
 */
function _applyRulesWithLog_(text, rules) {
  const src = String(text ?? '');
  if (!src || !Array.isArray(rules) || rules.length === 0) {
    return { text: src, edits: [] };
  }

  let cur = src;
  const edits = [];

  for (const r of rules) {
    const from = String(r.from ?? '');
    const to = String(r.to ?? '');
    if (!from) continue;

    const count = _countOccurrences_(cur, from);
    if (count <= 0) continue;

    // 문맥 샘플(첫 번째 발생 위치 주변 20자)
    const idx = cur.indexOf(from);
    const sample = _contextSample_(cur, idx, from.length, 20);

    cur = cur.split(from).join(to);

    edits.push({
      from,
      to,
      count,
      note: String(r.note ?? ''),
      sample
    });
  }

  return { text: cur, edits };
}

function _countOccurrences_(text, sub) {
  if (!sub) return 0;
  let count = 0;
  let pos = 0;
  while (true) {
    const idx = text.indexOf(sub, pos);
    if (idx === -1) break;
    count++;
    pos = idx + sub.length;
  }
  return count;
}

function _contextSample_(text, idx, len, pad) {
  if (idx < 0) return '';
  const start = Math.max(0, idx - pad);
  const end = Math.min(text.length, idx + len + pad);
  return text.slice(start, end);
}

// ✅ Rule 로그를 기존 UI 형식으로 맞춤
function _formatRuleEdit_(e) {
  const from = String(e.from ?? '');
  const to = String(e.to ?? '');
  const count = Number(e.count ?? 0);
  const note = String(e.note ?? '').trim();
  const phase = e.phase ? String(e.phase) : '';

  // 이유 문구: Rule 규칙 + (POST/PRE) + 적용횟수 + note
  const reasonParts = [];
  reasonParts.push(`Rule${phase ? `(${phase})` : ''}: "${from}" → "${to}"`);
  reasonParts.push(`${count}회`);
  if (note) reasonParts.push(`note: ${note}`);

  return `[[원본]] ${from}\n[[수정]] ${to}\n[[이유]] ${reasonParts.join(' | ')}`;
}

/** =========================
 * 기존 유틸
 * ========================= */

/** 출력 영역 정리: B3 + 21행 이하 전체(모든 열) */
function _clearOutput_(sh) {
  sh.getRange('B3').clearContent();

  const startRow = 21;
  const lastRow = sh.getMaxRows();
  const lastCol = sh.getMaxColumns();

  const numRows = lastRow - startRow + 1;
  if (numRows > 0) sh.getRange(startRow, 1, numRows, lastCol).clearContent();
}

/** 띄어쓰기(공백류, 줄바꿈 포함)만 바뀐 수정인지 판정 */
function _isWhitespaceOnlyEdit_(original, revised) {
  const o = String(original ?? '');
  const r = String(revised ?? '');
  if (o === r) return true;
  return o.replace(/\s+/g, '') === r.replace(/\s+/g, '');
}

/** source_index 보정: 1~refCount, 아니면 1로 */
function _sanitizeSourceIndex_(x, refCount) {
  const n = Number(x);
  if (!Number.isFinite(n) || n % 1 !== 0) return 1;
  if (refCount <= 0) return 1;
  if (n < 1) return 1;
  if (n > refCount) return 1;
  return n;
}

/** 결과 범위(D열)를 선택 상태로 만든 뒤 Viewer(Dialog) 오픈 */
function _selectResultRangeAndOpenViewer_(ss, sh, startRow, lastRow) {
  ss.setActiveSheet(sh);

  const rg = sh.getRange(startRow, 4, Math.max(1, lastRow - startRow + 1), 1); // D열
  rg.activate();
  ss.setActiveRange(rg);

  // LV가 없거나 openDialog가 없으면 그냥 토스트만
  try {
    if (typeof LV !== 'undefined' && LV && typeof LV.openDialog === 'function') {
      LV.openDialog();
    } else if (typeof lv_openDialog === 'function') {
      lv_openDialog();
    } else {
      ss.toast('LatexViewer(LV) 함수를 찾지 못했습니다.', '안내', 3);
    }
  } catch (e) {
    ss.toast(`Viewer 오픈 실패: ${e && e.message ? e.message : e}`, '오류', 5);
  }
}


/** 사용자 프롬프트 본문 (Rule 포함) */
function _buildPrompt_(target, chapter, refs, rules) {
  const lines = [];
  lines.push('다음은 "검토할 문항(원문)"과 "관련 기출문항"이다.');
  if (chapter) lines.push(`대단원(참고): ${chapter}`);
  lines.push('');

  lines.push('[검토할 문항(원문)]');
  lines.push(target);
  lines.push('');

  // ✅ Rule(강제 치환) 지시
  if (Array.isArray(rules) && rules.length > 0) {
    lines.push('[강제 수정 규칙(Rule)]');
    lines.push('아래 규칙의 "to" 표현은 최종 결과에서 절대 변경하지 마라(문법이 어색하면 주변 문장만 수정).');
    lines.push('또한 아래 규칙의 "from" 표현은 최종 결과에 남아있으면 안 된다.');
    rules.forEach((r, i) => {
      const note = r.note ? ` | note: ${r.note}` : '';
      lines.push(`- (${i + 1}) from: ${r.from}  ->  to: ${r.to}${note}`);
    });
    lines.push('');
  }

  lines.push('[관련 기출문항]');
  if (refs.length === 0) {
    lines.push('(없음)');
  } else {
    refs.forEach((r, i) => {
      lines.push(`(${i + 1}) source: ${r.source || '(빈값)'}`);
      lines.push(r.text);
      lines.push('');
    });
  }

  lines.push('');
  lines.push('요구사항에 맞춰 "검토할 문항"의 문장 스타일만 기출 스타일(수능형)로 다듬어라.');
  lines.push('문법/호응이 어색하면 Rule의 to를 건드리지 말고 주변 문장만 다듬어 자연스럽게 만들어라.');
  lines.push('수정 구절마다 근거가 된 관련 기출문항 번호(1~N)를 source_index에 적어라.');
  lines.push('반드시 JSON만 출력하라.');
  return lines.join('\n');
}

/** 모델 지시문 */
function _buildInstructions_() {
  return _getPromptValue_(PROMPT_KEY_INSTRUCTIONS);
}

/** Structured Outputs용 JSON Schema */
function _buildJsonSchema_() {
  return {
    type: 'object',
    additionalProperties: false,
    properties: {
      has_edits: { type: 'boolean' },
      rewritten_full: { type: 'string' },
      edits: {
        type: 'array',
        items: {
          type: 'object',
          additionalProperties: false,
          properties: {
            source_index: { type: 'integer', description: '근거가 된 관련 기출문항 번호(1~N)' },
            original: { type: 'string' },
            revised: { type: 'string' },
            reason: { type: 'string' }
          },
          required: ['source_index', 'original', 'revised', 'reason']
        }
      }
    },
    required: ['has_edits', 'rewritten_full', 'edits']
  };
}

/** OpenAI Responses API 호출 */
function _callOpenAIResponses_(apiKey, payload) {
  const url = 'https://api.openai.com/v1/responses';
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + apiKey },
    muteHttpExceptions: true,
    payload: JSON.stringify(payload)
  };

  const res = UrlFetchApp.fetch(url, options);
  const code = res.getResponseCode();
  const text = res.getContentText();

  if (code < 200 || code >= 300) throw new Error(`OpenAI API 오류 (${code}):\n${text}`);
  return JSON.parse(text);
}

/** Responses API 응답에서 assistant output_text 추출 */
function _extractOutputText_(resp) {
  const output = resp && resp.output;
  if (!Array.isArray(output)) return '';

  for (const item of output) {
    if (item && item.type === 'message' && item.role === 'assistant' && Array.isArray(item.content)) {
      for (const part of item.content) {
        if (part && (part.type === 'output_text' || part.type === 'text') && typeof part.text === 'string') {
          return part.text.trim();
        }
      }
    }
  }
  return '';
}

/** prompt 시트에서 값 읽는 함수 */
function _getPromptValue_(key) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(PROMPT_SHEET_NAME);
  if (!sh) throw new Error(`prompt 시트를 찾을 수 없습니다: ${PROMPT_SHEET_NAME}`);

  // A:key, B:value, C:enabled
  const lastRow = sh.getLastRow();
  if (lastRow < 2) throw new Error('prompt 시트에 데이터가 없습니다.');

  const values = sh.getRange(2, 1, lastRow - 1, 3).getValues(); // A~C
  for (const [k, v, enabled] of values) {
    if (String(k).trim() === key && (enabled === true || String(enabled).toUpperCase() === 'TRUE')) {
      const text = String(v ?? '').trim();
      if (!text) throw new Error(`prompt 시트의 ${key} 값이 비어있습니다.`);
      return text;
    }
  }
  throw new Error(`prompt 시트에서 활성화된 key를 찾지 못했습니다: ${key}`);
}

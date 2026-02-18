/**
 * OpenAI API로 "현재 키로 접근 가능한 모델" 전부 나열
 * - Logger.log로 출력
 * - {models: string[]} 형태로 반환
 *
 * 사전 준비:
 * Script Properties에 OPENAI_API_KEY 저장
 */
function listOpenAIModels() {
  const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!apiKey) throw new Error('Script Properties에 OPENAI_API_KEY가 없어.');

  const url = 'https://api.openai.com/v1/models';
  const res = UrlFetchApp.fetch(url, {
    method: 'get',
    muteHttpExceptions: true,
    headers: {
      Authorization: 'Bearer ' + apiKey
    }
  });

  const code = res.getResponseCode();
  const text = res.getContentText();

  if (code < 200 || code >= 300) {
    throw new Error(`OpenAI API Error (${code}): ${text}`);
  }

  const json = JSON.parse(text);
  const models = (json.data || [])
    .map(m => m && m.id)
    .filter(Boolean)
    .sort((a, b) => a.localeCompare(b));

  Logger.log('=== 사용 가능한 모델 목록 (%s개) ===', models.length);
  models.forEach((id, i) => Logger.log('%s) %s', String(i + 1).padStart(3, '0'), id));

  return { models };
}

/**
 * (옵션) 모델 목록을 현재 시트에 뿌리기
 * - A열: model_id
 * - B열: fetched_at
 */
function dumpOpenAIModelsToSheet() {
  const { models } = listOpenAIModels();
  const sheet = SpreadsheetApp.getActiveSheet();

  const now = new Date();
  const rows = models.map(id => [id, now]);

  sheet.getRange(1, 1, 1, 2).setValues([['model_id', 'fetched_at']]);
  if (rows.length) sheet.getRange(2, 1, rows.length, 2).setValues(rows);
  sheet.autoResizeColumns(1, 2);
}

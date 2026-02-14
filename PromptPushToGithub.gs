/**
 * 'prompt' 시트의 각 키를 prompt/ 폴더 내 개별 .txt 파일로 GitHub에 Push합니다.
 */
function pushIndividualPromptsToGithub() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('prompt');
  if (!sh) throw new Error('prompt 시트를 찾을 수 없습니다.');

  var props = PropertiesService.getScriptProperties();
  var GITHUB_TOKEN = props.getProperty('GITHUB_TOKEN');
  var REPO = props.getProperty('GITHUB_REPO');
  
  if (!GITHUB_TOKEN || !REPO) {
    throw new Error('스크립트 속성에 GITHUB_TOKEN 또는 GITHUB_REPO가 설정되지 않았습니다.');
  }

  // 1. 시트 데이터 읽기 (Key: A열, Value: B열, Enabled: C열)
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return;
  var data = sh.getRange(2, 1, lastRow - 1, 3).getValues();

  var successCount = 0;
  var errorCount = 0;

  // 2. 각 행을 순회하며 개별 파일로 Push
  data.forEach(function(row) {
    var key = String(row[0]).trim();
    var content = String(row[1]).trim();
    var enabled = (row[2] === true || String(row[2]).toUpperCase() === 'TRUE');

    // 키가 비어있거나 비활성화된 경우 건너뜀
    if (!key || !enabled) return;

    var filePath = 'prompt/' + key + '.txt';
    var commitMessage = 'Update prompt: ' + key + ' (' + new Date().toLocaleString() + ')';

    try {
      // 기존 파일 SHA 확인
      var sha = _getGithubFileSha(REPO, filePath, GITHUB_TOKEN);

      // GitHub API 호출
      var url = 'https://api.github.com/repos/' + REPO + '/contents/' + filePath;
      var payload = {
        message: commitMessage,
        content: Utilities.base64Encode(content, Utilities.Charset.UTF_8),
        sha: sha
      };

      var options = {
        method: 'put',
        contentType: 'application/json',
        headers: {
          'Authorization': 'Bearer ' + GITHUB_TOKEN,
          'Accept': 'application/vnd.github+json'
        },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      };

      var response = UrlFetchApp.fetch(url, options);
      if (response.getResponseCode() === 200 || response.getResponseCode() === 201) {
        successCount++;
      } else {
        console.error('실패: ' + key + ' - ' + response.getContentText());
        errorCount++;
      }
    } catch (e) {
      console.error('에러 발생: ' + key + ' - ' + e.message);
      errorCount++;
    }
  });

  SpreadsheetApp.getUi().alert('작업 완료\n성공: ' + successCount + '건\n실패: ' + errorCount + '건');
}

/**
 * 기존 파일의 SHA를 가져오는 유틸리티 함수
 */
function _getGithubFileSha(repo, path, token) {
  var url = 'https://api.github.com/repos/' + repo + '/contents/' + encodeURIComponent(path);
  var options = {
    method: 'get',
    headers: { 'Authorization': 'Bearer ' + token },
    muteHttpExceptions: true
  };
  var res = UrlFetchApp.fetch(url, options);
  if (res.getResponseCode() === 200) {
    return JSON.parse(res.getContentText()).sha;
  }
  return null;
}
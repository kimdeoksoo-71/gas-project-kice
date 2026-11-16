/**
 * 사이드바 열기
 */
function showCellPreviewSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('선택 영역 미리보기')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * 선택된 모든 셀의 내용을 배열로 가져오는 함수
 */
function getActiveRangeContents() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getActiveRange();
  const values = range.getDisplayValues(); // 화면에 보이는 값 그대로 가져옴
  const startRow = range.getRow();
  const startCol = range.getColumn();
  
  let contents = [];
  
  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      contents.push({
        value: values[r][c] || '',
        address: sheet.getRange(startRow + r, startCol + c).getA1Notation()
      });
    }
  }
  
  // 의견 셀(E·F·G, 새 형식)의 [[ID]] 줄 → Feedback 시트의 판정. 사이드바가 의견마다 채택/기각 버튼을 붙인다 (kice-2.0.2)
  const want = {};
  contents.forEach(c => {
    const re = /\[\[ID\]\][ \t]*([^\n<]+)/g; let m;
    while ((m = re.exec(c.value))) { const id = m[1].trim(); if (id.split('#').length === 3) want[id] = true; }
  });
  let feedback = {};
  try { feedback = QV.feedbackByIds(want); } catch (e) { feedback = {}; }

  return {
    feedback: feedback,
    contents: contents,
    sheetName: sheet.getName(),
    rangeAddress: range.getA1Notation()
  };
}
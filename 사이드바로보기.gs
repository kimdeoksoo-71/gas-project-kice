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
  
  return {
    contents: contents,
    sheetName: sheet.getName(),
    rangeAddress: range.getA1Notation()
  };
}
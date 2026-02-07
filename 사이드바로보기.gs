/**
 * 사이드바 열기
 */
function showCellPreviewSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('셀 미리보기');
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * 현재 활성 시트의 활성 셀 내용을 가져오는 함수
 * 사이드바에서 주기적으로 이걸 호출해서 새 내용을 받아감
 */
function getActiveCellContent() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const cell = sheet.getActiveCell();
  // 표시되는 값으로 가져오는 편이 자연스러움
  const value = cell.getDisplayValue();
  return {
    value: value || '',
    sheetName: sheet.getName(),
    address: cell.getA1Notation()
  };
}

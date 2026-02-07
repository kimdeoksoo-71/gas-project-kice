function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('🧩 사용자도구');

  // ---- 기타 도구 ----
  menu.addItem('뷰어 : 원본과 비교', 'lv_openDialog');
  menu.addItem('뷰어 : 키워드검색 결과', 'sv_openSearchViewer');
  menu.addItem('🪟 Que 결과 뷰어', 'qv_openDialog')


  menu.addItem('사이드바','showCellPreviewSidebar');

  // ---- 유사문항 검색과 문항별 윤문 ----
  menu.addSeparator();
  menu.addItem('기출 유사문항 검색 : 문항검토B2','findSimilarFromB2');
  menu.addItem('키워드 검색 : 검색', 'searchDataLatexByKeywords');
  
  // ---- 문항 윤문 ----
  menu.addSeparator();
  menu.addItem('🧵 Que 자동윤문 시작', 'batch_startQueAuto')
  menu.addItem('⛔ Que 자동윤문 중지', 'batch_stopQueAuto')
  menu.addItem('한 문항 윤문 제안', 'review_rewriteToKiceStyle_gpt52');
  
  // ---- 검색어로 문항 나열하여 Latex 변환 ----
  menu.addSeparator();
  menu.addItem('검색어로 문항나열', 'runSearchAndAppend');
  menu.addItem('행 범위 입력 → Latex 변환', 'mpb_runRange');
  menu.addItem('BatchStart', 'mpb_batchStart');
  
  // ---- Token 테이블 ----
  menu.addSeparator();
  menu.addItem('Token_Stat 만들기', 'buildTokenStatFromDataLatex');
  
  menu.addToUi();
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('🍋 앱 메뉴');

  // ---- 기타 도구 ----
  menu.addItem('뷰어 : 원본과 비교', 'lv_openDialog');
  menu.addItem('뷰어 : 키워드검색 결과', 'sv_openSearchViewer');
  menu.addItem('사이드바','showCellPreviewSidebar');
  menu.addItem('📂 Que 결과 뷰어', 'qv_openDialog');

  // ---- 유사문항 검색과 문항별 윤문 ----
  menu.addSeparator();
  menu.addItem('🔍 기출 유사문항 검색 : 문항검토B2','findSimilarFromB2');
  menu.addItem('🔑 키워드 검색 : 검색', 'searchDataLatexByKeywords');
  
  // ---- 문항 윤문 ----
  menu.addSeparator();
  menu.addItem('▶️ Que 자동윤문 시작', 'batch_startQueAuto');
  menu.addItem('⏹️ Que 자동윤문 중지', 'batch_stopQueAuto');
  menu.addSeparator
  menu.addItem('현재 프롬프트를 github에 push', 'pushIndividualPromptsToGithub')
  
  // ---- Token 테이블 ----
  menu.addSeparator();
  menu.addItem('Token_Stat 만들기', 'buildTokenStatFromDataLatex');
  
  menu.addToUi();
}

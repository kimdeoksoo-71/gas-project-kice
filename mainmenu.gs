/** 메뉴 (kice-2.0.0) — 검색·Token_Stat 메뉴는 기존 파일 「기출과 윤문(2011~)」에 남아 있다 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('🍋 KICE 윤문');

  // ---- 윤문 ----
  menu.addItem('▶ Que 윤문 요청', 'batch_startQueAuto');
  menu.addItem('⏹ 중지', 'batch_stopQueAuto');
  menu.addItem('📋 상태', 'kice_showStatus');

  // ---- 보기 ----
  menu.addSeparator();
  menu.addItem('🅰️ 아카이브 뷰어', 'qv_openDialog');
  menu.addItem('➡️ 사이드바 뷰어', 'showCellPreviewSidebar');

  // ---- 프롬프트 ----
  menu.addSeparator();
  menu.addItem('✅ 현재 프롬프트를 github에 push', 'pushIndividualPromptsToGithub');

  menu.addToUi();
}

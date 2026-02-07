/*************************************************
 * Que 결과 뷰어 (Dialog)
 * - 사용자 입력: 시작행/마지막행
 * - Que(A:id,B:latex,C:chapter,D:finalfull,E:rewrittens)
 * - E가 "수정 구절 없음"인 행은 제외
 * - 좌: (상단) id+latex  / (하단) finalfull
 * - 우: rewrittens(HTML 그대로 innerHTML 렌더)
 *************************************************/

const QV = (function () {
  const TITLE = 'Que Viewer';
  const SHEET = 'Que';

  // 사용자별 상태 저장 키
  const KEY = 'QV_RANGE_V1';

  function openDialog() {
    const ss = SpreadsheetApp.getActive();
    const ui = SpreadsheetApp.getUi();

    // 1) Alert(=prompt)로 시작행/마지막행 받기
    const resp = ui.prompt(
      'Que 결과 보기',
      '시작행,마지막행을 입력해줘 (예: 2,50)',
      ui.ButtonSet.OK_CANCEL
    );
    if (resp.getSelectedButton() !== ui.Button.OK) return;

    const txt = String(resp.getResponseText() || '').trim();
    const m = txt.match(/^(\d+)\s*,\s*(\d+)$/);
    if (!m) {
      ss.toast('입력 형식 오류: "시작행,마지막행" 예) 2,50', '안내', 5);
      return;
    }

    let startRow = Number(m[1]);
    let endRow = Number(m[2]);
    if (!Number.isFinite(startRow) || !Number.isFinite(endRow)) return;

    if (startRow > endRow) [startRow, endRow] = [endRow, startRow];
    startRow = Math.max(2, Math.floor(startRow));
    endRow = Math.max(startRow, Math.floor(endRow));

    // 상태 저장(사용자별)
    PropertiesService.getUserProperties().setProperty(KEY, JSON.stringify({
      sheet: SHEET,
      startRow,
      endRow,
      t: Date.now()
    }));

    // 2) Dialog 오픈
    const html = HtmlService.createHtmlOutputFromFile('QueViewerDialog')
      .setTitle(TITLE)
      .setWidth(1200)
      .setHeight(1400);

    SpreadsheetApp.getUi().showModalDialog(html, TITLE);
  }

  function getPayload() {
    const ss = SpreadsheetApp.getActive();
    const stRaw = PropertiesService.getUserProperties().getProperty(KEY);
    if (!stRaw) {
      return { signature: 'no_state', sheetName: SHEET, items: [] };
    }

    let st;
    try {
      st = JSON.parse(stRaw);
    } catch (e) {
      return { signature: 'bad_state', sheetName: SHEET, items: [] };
    }

    const sh = ss.getSheetByName(st.sheet || SHEET);
    if (!sh) {
      return { signature: 'no_sheet', sheetName: st.sheet || SHEET, items: [] };
    }

    const startRow = Number(st.startRow);
    const endRow = Number(st.endRow);
    const h = Math.max(0, endRow - startRow + 1);
    if (h <= 0) return { signature: 'empty_range', sheetName: sh.getName(), items: [] };

    // A~E 읽기
    const vals = sh.getRange(startRow, 1, h, 5).getValues();

    const items = [];
    for (let i = 0; i < vals.length; i++) {
      const row = startRow + i;
      const id = String(vals[i][0] ?? '').trim();
      const latex = String(vals[i][1] ?? '').trim();
      const finalfull = String(vals[i][3] ?? '').trim();
      const rew = String(vals[i][4] ?? '').trim();

      // 2) E가 "수정 구절 없음"이면 제외
      if (!rew || rew === '수정 구절 없음') continue;

      items.push({
        row,
        id,
        latex,
        finalfull,
        rewrittens_html: rew
      });
    }

    const signature =
      `${sh.getName()}|${startRow}-${endRow}|cnt:${items.length}|t:${st.t}|` +
      items.map(it => `${it.row}:${it.rewrittens_html.length}`).join(',');

    return {
      signature,
      sheetName: sh.getName(),
      range: { startRow, endRow },
      count: items.length,
      items
    };
  }

  return { openDialog, getPayload };
})();

function qv_openDialog() {
  QV.openDialog();
}

function qv_getPayload() {
  return QV.getPayload();
}

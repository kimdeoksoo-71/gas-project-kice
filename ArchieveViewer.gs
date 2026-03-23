/*************************************************
 * 아카이브 결과 뷰어 (Dialog)
 * - 사용자 입력: 시작행-마지막행
 * - 아카이브(A:id, B:Latex, C:chapter, D:claude, E:gpt, F:gemini)
 * - 좌: B(Latex)
 * - 우: D(Claude), E(GPT), F(Gemini) 순차 표시
 *************************************************/

const QV = (function () {
  const TITLE = '아카이브 Viewer';
  const SHEET = '아카이브';
  const KEY = 'QV_RANGE_V1';

  function openDialog() {
    const ss = SpreadsheetApp.getActive();
    const ui = SpreadsheetApp.getUi();

    const resp = ui.prompt(
      '아카이브 결과 보기',
      '시작행-마지막행을 입력해줘 (예: 2-50)',
      ui.ButtonSet.OK_CANCEL
    );
    if (resp.getSelectedButton() !== ui.Button.OK) return;

    const txt = String(resp.getResponseText() || '').trim();
    const m = txt.match(/^(\d+)\s*-\s*(\d+)$/);
    if (!m) {
      ss.toast('입력 형식 오류: "시작행-마지막행" 예) 2-50', '안내', 5);
      return;
    }

    let startRow = Number(m[1]);
    let endRow = Number(m[2]);
    if (!Number.isFinite(startRow) || !Number.isFinite(endRow)) return;

    if (startRow > endRow) [startRow, endRow] = [endRow, startRow];
    startRow = Math.max(2, Math.floor(startRow));
    endRow = Math.max(startRow, Math.floor(endRow));

    PropertiesService.getUserProperties().setProperty(KEY, JSON.stringify({
      sheet: SHEET,
      startRow,
      endRow,
      t: Date.now()
    }));

    const html = HtmlService.createHtmlOutputFromFile('ArchiveViewer')
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

    // A~F 읽기 (6열)
    const vals = sh.getRange(startRow, 1, h, 6).getValues();

    const items = [];
    for (let i = 0; i < vals.length; i++) {
      const row = startRow + i;
      const id = String(vals[i][0] ?? '').trim();
      const latex = String(vals[i][1] ?? '').trim();
      const claude = String(vals[i][3] ?? '').trim();
      const gpt = String(vals[i][4] ?? '').trim();
      const gemini = String(vals[i][5] ?? '').trim();

      // D,E,F 모두 비어있으면 제외
      if (!claude && !gpt && !gemini) continue;

      items.push({
        row,
        id,
        latex,
        claude,
        gpt,
        gemini
      });
    }

    const signature =
      `${sh.getName()}|${startRow}-${endRow}|cnt:${items.length}|t:${st.t}|` +
      items.map(it => `${it.row}:${it.claude.length}:${it.gpt.length}:${it.gemini.length}`).join(',');

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
/*************************************************
 * 아카이브 결과 뷰어 (Dialog)
 * - 사용자 입력: 시작행-마지막행
 * - 아카이브 열 배치는 Que 와 같다: A id · B Latex · C chapter · D (비움) · E Claude · F GPT · G Gemini
 *   (2026-10-02 실측: 864행부터 이 배치. 863행 이전은 형식이 혼재 — D HTML·E 윤문 전문·F HTML / D·E·F 결과 셋 / E·F 두 칸 …)
 * - 좌: B(Latex)
 * - 우 [v4 D23]:
 *     row ≥ QV_LAYOUT_SWITCH_ROW(864): E/F/G 를 Claude / GPT / Gemini 로 표시
 *     row <  864: D~G 중 값이 있는 열을 "D열(옛)" 처럼 열 이름 그대로 표시 (회사 이름을 추측해 붙이지 않음)
 * - 호환: 예전 필드 claude/gpt/gemini 도 새 배치(E/F/G) 기준으로 계속 채운다
 *************************************************/

const QV = (function () {
  const TITLE = '아카이브 Viewer';
  const SHEET = '아카이브';
  const KEY = 'QV_RANGE_V1';
  const QV_LAYOUT_SWITCH_ROW = 864;          // 이 행부터 E/F/G = Claude/GPT/Gemini (아카이브는 끝에만 붙이므로 고정)
  const OLD_COLS = [                          // 863행 이전: 열 이름 그대로
    { idx: 3, label: 'D열(옛)' },
    { idx: 4, label: 'E열(옛)' },
    { idx: 5, label: 'F열(옛)' },
    { idx: 6, label: 'G열(옛)' }
  ];

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

    // A~G 읽기 (7열) — 시트 열이 7보다 적으면 있는 만큼만
    const nCols = Math.min(7, sh.getMaxColumns());
    const vals = sh.getRange(startRow, 1, h, nCols).getValues();
    const cell = (row, i) => String((row[i] === undefined || row[i] === null) ? '' : row[i]).trim();

    const items = [];
    for (let i = 0; i < vals.length; i++) {
      const row = startRow + i;
      const v = vals[i];
      const id = cell(v, 0);
      const latex = cell(v, 1);

      let blocks;
      let layout;
      if (row >= QV_LAYOUT_SWITCH_ROW) {
        layout = 'new';
        blocks = [
          { label: 'Claude', cls: 'claude', html: cell(v, 4) },
          { label: 'GPT',    cls: 'gpt',    html: cell(v, 5) },
          { label: 'Gemini', cls: 'gemini', html: cell(v, 6) }
        ].filter(b => b.html);
      } else {
        layout = 'old';
        blocks = OLD_COLS
          .map(c => ({ label: c.label, cls: 'old', html: cell(v, c.idx) }))
          .filter(b => b.html);
      }

      // 결과가 하나도 없으면 제외
      if (!blocks.length) continue;

      items.push({
        row,
        id,
        latex,
        layout,
        blocks,
        // 호환 필드 (새 배치 기준)
        claude: cell(v, 4),
        gpt: cell(v, 5),
        gemini: cell(v, 6)
      });
    }

    const signature =
      `${sh.getName()}|${startRow}-${endRow}|cnt:${items.length}|t:${st.t}|` +
      items.map(it => `${it.row}:${it.layout}:${it.blocks.map(b => b.html.length).join('/')}`).join(',');

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

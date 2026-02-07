/*************************************************
 * LaTeX 비교 뷰어 (Dialog)
 * - 왼쪽: 현재 시트의 B2(원본) 표시 (독립 스크롤)
 * - 오른쪽: 같은 열에서 선택된 여러 셀(유사문항) 목록 표시 (독립 스크롤)
 *   + 각 항목 헤더에 A열(source) 함께 표시
 *   + (요청) '문항검토' 시트에서 C열(6행 이하)의 드라이브 링크를 "원본link"로 노출
 *************************************************/

const LV = (function () {
  const TITLE = 'Viewer';

  function openDialog() {
    const html = HtmlService.createHtmlOutputFromFile('ViewerDialog')
      .setTitle(TITLE)
      .setWidth(1800)
      .setHeight(1400);
    SpreadsheetApp.getUi().showModalDialog(html, TITLE);
  }

  /** ✅ 왼쪽 패널: 항상 현재 시트 B2 */
  function getB2Content_() {
    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getActiveSheet();
    const cell = sheet.getRange('B2');
    return {
      sheetName: sheet.getName(),
      a1: 'B2',
      row: 2,
      col: 2,
      text: String(cell.getDisplayValue() ?? '')
    };
  }

  /**
   * ✅ 오른쪽 패널: 선택된 여러 셀(ActiveRangeList)에서
   * "활성 셀의 열"만 추출 → 행 오름차순 → items 반환
   * + 각 row의 A열(source)
   * + (요청) 문항검토 시트라면 C열 원본 링크도 포함
   */
  function getSelectionContents_() {
    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getActiveSheet();
    const activeCell = sheet.getActiveCell();
    const targetCol = activeCell.getColumn();

    const rangeList = sheet.getActiveRangeList();
    const ranges = rangeList ? rangeList.getRanges() : [sheet.getActiveRange()];

    const rowsSet = new Set();

    ranges.forEach(rg => {
      const r1 = rg.getRow();
      const c1 = rg.getColumn();
      const nr = rg.getNumRows();
      const nc = rg.getNumColumns();

      // targetCol을 포함하는 범위만: 해당 열의 행들을 수집
      if (targetCol >= c1 && targetCol <= c1 + nc - 1) {
        for (let i = 0; i < nr; i++) rowsSet.add(r1 + i);
      }
    });

    // 혹시라도 비면 activeCell만
    if (rowsSet.size === 0) rowsSet.add(activeCell.getRow());

    const rows = Array.from(rowsSet).sort((a, b) => a - b);

    // ✅ 배치 읽기: A열(source) + C열(link) + targetCol(유사문항)
    const minRow = rows[0];
    const maxRow = rows[rows.length - 1];
    const h = maxRow - minRow + 1;

    const aDisp   = sheet.getRange(minRow, 1, h, 1).getDisplayValues();            // A열
    const simDisp = sheet.getRange(minRow, targetCol, h, 1).getDisplayValues();    // 선택 열

    // (요청) 문항검토 시트일 때만 C열도 읽음
    const isReviewSheet = (sheet.getName() === '문항검토');
    const cVals = isReviewSheet
      ? sheet.getRange(minRow, 3, h, 1).getValues()  // C열(링크는 getValues로)
      : null;

    const colA1 = columnToLetter_(targetCol);

    const items = rows.map(r => {
      const i = r - minRow;

      // C열 링크: 문항검토 & row>=6일 때만 의미있게 사용
      let origLink = '';
      if (isReviewSheet && r >= 6) {
        const v = cVals?.[i]?.[0];
        origLink = (v && typeof v === 'string') ? v.trim() : String(v ?? '').trim();
      }

      return {
        row: r,
        col: targetCol,
        a1: `${colA1}${r}`,
        text: String(simDisp[i][0] ?? ''),
        source: String(aDisp[i][0] ?? ''),
        origLink
      };
    });

    const signature =
      `${sheet.getName()}|C${targetCol}|` +
      items.map(it => `${it.a1}:${(it.origLink || '').length}`).join(',');

    return {
      sheetName: sheet.getName(),
      col: targetCol,
      count: items.length,
      signature,
      items
    };
  }

  /**
   * ✅ 좌(B2) + 우(선택셀목록) 한 번에 반환
   */
  function getComparePayload() {
    const left = getB2Content_();
    const right = getSelectionContents_();
    const signature = `${left.sheetName}|B2|len:${left.text.length}|${right.signature}`;
    return { signature, left, right };
  }

  /** col number → A1 letter (1->A, 27->AA) */
  function columnToLetter_(column) {
    let temp = '';
    let letter = '';
    let col = Number(column);
    while (col > 0) {
      temp = (col - 1) % 26;
      letter = String.fromCharCode(temp + 65) + letter;
      col = (col - temp - 1) / 26;
    }
    return letter;
  }

  return { openDialog, getComparePayload };
})();

function lv_openDialog() {
  LV.openDialog();
}

function lv_getComparePayload() {
  return LV.getComparePayload();
}

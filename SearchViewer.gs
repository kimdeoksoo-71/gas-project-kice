/*************************************************
 * 검색 시트 전용 뷰어 (Dialog)
 * - 오른쪽: '검색' 시트에서 선택된 셀들을 모아서 표시
 *   + 각 항목 헤더에 A열(source), C열(구글드라이브 이미지 링크) 표시
 * - 왼쪽: 검색어/검색논리 표시
 *   (맨 위) G2 대단원
 *   H2&H3&H4 / I2&I3&I4 / J2&J3&J4 (줄바꿈)
 *
 * ✅ 기존 LV / lv_* / ViewerDialog 와 충돌 없음
 *************************************************/

const SV = (function () {
  const TITLE = '키워드 검색 Viewer';

  function openDialog() {
    const html = HtmlService.createHtmlOutputFromFile('SearchViewerDialog')
      .setTitle(TITLE)
      .setWidth(1800)
      .setHeight(1400);
    SpreadsheetApp.getUi().showModalDialog(html, TITLE);
  }

  /** 왼쪽: G2(대단원) + 검색어 블록(H/I/J 묶음) */
  function getLeftQueryBlock_() {
    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName('검색');
    if (!sh) throw new Error("시트 '검색'을 찾을 수 없습니다.");

    const major = String(sh.getRange('G2').getDisplayValue() ?? '').trim();

    const H = sh.getRange('H2:H4').getDisplayValues().flat();
    const I = sh.getRange('I2:I4').getDisplayValues().flat();
    const J = sh.getRange('J2:J4').getDisplayValues().flat();

    const line = (arr) => arr.map(v => String(v ?? '').trim()).filter(Boolean).join(' & ');

    const lines = [];
    if (major) lines.push(`대단원: ${major}`);
    lines.push(line(H));
    lines.push(line(I));
    lines.push(line(J));

    const text = lines.filter(Boolean).join('\n');

    return {
      sheetName: sh.getName(),
      a1: 'G2 / H2:H4 / I2:I4 / J2:J4',
      row: 2,
      col: 7,
      text: text || '(검색어 없음)'
    };
  }

  /** 오른쪽: 선택된 모든 셀 수집 + 해당 행 A/C 같이 */
  function getSelectionContents_() {
    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getActiveSheet();
    if (sheet.getName() !== '검색') {
      throw new Error("이 뷰어는 '검색' 시트에서만 사용합니다.");
    }

    const rangeList = sheet.getActiveRangeList();
    const ranges = rangeList ? rangeList.getRanges() : [sheet.getActiveRange()];

    const cells = [];
    ranges.forEach(rg => {
      const r1 = rg.getRow();
      const c1 = rg.getColumn();
      const nr = rg.getNumRows();
      const nc = rg.getNumColumns();
      for (let i = 0; i < nr; i++) {
        for (let j = 0; j < nc; j++) {
          cells.push({ row: r1 + i, col: c1 + j });
        }
      }
    });

    // 중복 제거 + 정렬(행 우선, 열 차선)
    const uniq = Array.from(new Map(cells.map(x => [`${x.row},${x.col}`, x])).values())
      .sort((a, b) => (a.row - b.row) || (a.col - b.col));

    if (uniq.length === 0) {
      const ac = sheet.getActiveCell();
      uniq.push({ row: ac.getRow(), col: ac.getColumn() });
    }

    // ✅ A(source), C(link) 배치 읽기
    const rows = uniq.map(x => x.row);
    const minRow = Math.min(...rows);
    const maxRow = Math.max(...rows);
    const h = maxRow - minRow + 1;

    const aDisp = sheet.getRange(minRow, 1, h, 1).getDisplayValues(); // A: source
    const cVals = sheet.getRange(minRow, 3, h, 1).getValues();        // C: drive link

    const items = uniq.map(pos => {
      const i = pos.row - minRow;
      const cell = sheet.getRange(pos.row, pos.col);

      const source = String(aDisp[i][0] ?? '');

      const rawLink = cVals?.[i]?.[0];
      const imgLink =
        (rawLink && typeof rawLink === 'string')
          ? rawLink.trim()
          : String(rawLink ?? '').trim();

      return {
        row: pos.row,
        col: pos.col,
        a1: `${columnToLetter_(pos.col)}${pos.row}`,
        text: String(cell.getDisplayValue() ?? ''),
        source,
        imgLink
      };
    });

    const signature =
      `${sheet.getName()}|SV|` +
      items.map(it => `${it.a1}:${it.text.length}:${(it.imgLink || '').length}`).join(',');

    return { sheetName: sheet.getName(), count: items.length, signature, items };
  }

  function getComparePayload() {
    const left = getLeftQueryBlock_();
    const right = getSelectionContents_();
    const signature = `${left.sheetName}|SVLEFT|len:${left.text.length}|${right.signature}`;
    return { signature, left, right };
  }

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

/** ✅ 검색 시트 전용 실행 함수(새 이름) */
function sv_openSearchViewer() {
  SV.openDialog();
}

/** ✅ HTML에서 호출하는 서버 함수(새 이름) */
function sv_getComparePayload() {
  return SV.getComparePayload();
}

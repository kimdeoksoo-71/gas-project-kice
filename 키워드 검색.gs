function searchDataLatexByKeywords() {
  const ss = SpreadsheetApp.getActive();
  const searchSheet = ss.getSheetByName('검색');
  const dataSheet   = ss.getSheetByName('Data_Latex');

  if (!searchSheet || !dataSheet) {
    ss.toast('"검색" 시트 또는 "Data_Latex" 시트를 찾을 수 없습니다.', '검색', 6);
    return;
  }

  // --- 1. 검색 조건 읽기 ---
  const chapter = String(searchSheet.getRange('G2').getDisplayValue() ?? '').trim();

  const getKeywords = (colLetter) => {
    return searchSheet
      .getRange(colLetter + '2:' + colLetter + '4')
      .getValues()
      .flat()
      .map(v => String(v ?? '').trim())
      .filter(v => v);
  };

  const kw1 = getKeywords('H');
  const kw2 = getKeywords('I');
  const kw3 = getKeywords('J');

  if (kw1.length === 0 && kw2.length === 0 && kw3.length === 0) {
    ss.toast('검색어가 없습니다. H/I/J 열에 검색어를 입력해주세요.', '검색', 6);
    return;
  }

  const data = dataSheet.getDataRange().getValues();
  const result = [];

  const matchGroup = (text, keywords) => {
    if (!keywords.length) return false;
    return keywords.every(k => text.indexOf(k) !== -1);
  };

  // --- 2. 조건에 맞는 행 찾기 ---
  for (let r = 1; r < data.length; r++) {
    const row = data[r];

    const colI = row[8]; // source
    const colJ = row[9]; // chapter
    const colB = row[1]; // drive_link
    const colC = row[2]; // latex

    if (!colC) continue;
    if (chapter && String(colJ) !== chapter) continue;

    const text = String(colC);

    const ok =
      matchGroup(text, kw1) ||
      matchGroup(text, kw2) ||
      matchGroup(text, kw3);

    if (!ok) continue;

    // ✅ 검색 시트 출력 형식: A=source, B=chapter, C=drive_link, D=latex
    result.push([colI, colJ, colB, colC]);
  }

  // --- 3. source(A열) 기준 내림차순 정렬 ---
  result.sort((a, b) => String(b[0]).localeCompare(String(a[0])));

  // --- 4. '검색' 시트에 출력 ---
  const lastRow = searchSheet.getLastRow();
  if (lastRow > 1) {
    searchSheet.getRange(2, 1, lastRow - 1, 4).clearContent();
  }

  if (result.length === 0) {
    ss.toast('조건에 맞는 문항을 찾지 못했습니다.', '검색', 6);
    return;
  }

  searchSheet.getRange(2, 1, result.length, 4).setValues(result);

  // --- ✅ 5. 결과 범위 선택 + 뷰어 자동 호출 ---
  // 뷰어는 "선택된 셀"을 보여주니까, latex가 들어있는 D열을 자동 선택해줌
  searchSheet.activate();
  const selectRange = searchSheet.getRange(2, 4, result.length, 1); // D2:D
  selectRange.activate();

  ss.toast(`✅ ${result.length}개 결과 출력 + 뷰어 실행`, '검색', 5);

  // 뷰어 호출(네가 만든 함수)
  sv_openSearchViewer();
}

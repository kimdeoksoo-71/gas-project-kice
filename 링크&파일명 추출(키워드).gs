/***********************
 * PBM: IMAGE 폴더 PNG 검색 → Data1!A:B 기록
 * - 메뉴: PBM 도구 > 이미지 검색·기록 실행
 ***********************/

function runSearchAndAppend() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('PNG 파일명 검색', '파일명에 포함될 단어를 입력하세요 (예: 2021, 미적분 등)', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const keyword = (res.getResponseText() || '').trim();
  if (!keyword) { ui.alert('검색어가 비어 있습니다.'); return; }

  const folder = getFolderByPath('PBMAI/IMAGE'); // \, / 모두 허용
  const pairs = collectPngNameUrlPairs(folder, keyword);
  if (pairs.length === 0) { ui.alert('조건에 맞는 PNG 파일이 없습니다.'); return; }

  // 한글 가나다 + 숫자 자연 정렬
  pairs.sort((a, b) => a[0].localeCompare(b[0], 'ko', { sensitivity: 'base', numeric: true }));

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('Data1'); if (!sh) sh = ss.insertSheet('Data1');
  const startRow = Math.max(2, sh.getLastRow() + 1);
  sh.getRange(startRow, 1, pairs.length, 2).setValues(pairs);

  ui.alert(`총 ${pairs.length}건을 기록했습니다. (시작행: ${startRow})`);
}

/** 경로 기반 폴더 찾기: 'PBMAI/IMAGE' 또는 'PBMAI\\IMAGE' 모두 허용 */
function getFolderByPath(pathLike) {
  const path = pathLike.replace(/\\/g, '/').replace(/^\/+|\/+$/g, '');
  const parts = path.split('/');
  let cur = DriveApp.getRootFolder(); // 내 드라이브 루트
  for (const name of parts) {
    const it = cur.getFoldersByName(name);
    if (!it.hasNext()) throw new Error(`폴더를 찾을 수 없습니다: ${name} (경로: ${pathLike})`);
    cur = it.next(); // 동일 이름 여러 개면 첫 번째 사용
  }
  return cur;
}

/** 폴더(1레벨) 내 PNG만, 파일명에 keyword 포함된 항목 → [filename, url] */
function collectPngNameUrlPairs(folder, keyword) {
  const kw = keyword.toLowerCase();
  const out = [];
  const files = folder.getFiles();
  while (files.hasNext()) {
    const f = files.next();
    const name = f.getName(); if (!name) continue;
    const lower = name.toLowerCase();
    if (lower.endsWith('.png') && lower.indexOf(kw) !== -1) out.push([name, f.getUrl()]);
  }
  return out;
}

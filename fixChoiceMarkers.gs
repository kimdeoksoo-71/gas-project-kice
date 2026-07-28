/*******************************************************
 * (패치 E) Data_Latex 선지 번호 원상복구 — 1회 실행용
 *
 * DB 구축 당시 원문자 선지 번호(①~⑤)가 괄호 숫자
 * ((1)~(5))로 일괄 변환된 것을 되돌린다.
 *
 * ▸ C열(latex)에서 "줄 첫머리"의 (1)~(5)만 ①~⑤로 변환
 *   → 본문·수식 속 괄호 숫자(f(1) 등)는 건드리지 않음
 * ▸ 실행 전 스프레드시트 사본 백업 권장
 * ▸ 실행 후 "Token_Stat 만들기"를 한 번 다시 실행 권장
 *
 * 효과: refs에 올바른 표기가 실려 LLM 혼란의 근원이
 * 사라지고, "(1)로 바꾸라"는 edit은 revised가 기출에
 * 없게 되어 기존 검증(_existsInAnyRef_)만으로도 자동
 * 폐기된다. (패치 A·B는 이중 방어로 그대로 유지)
 *******************************************************/

function fixChoiceMarkersInDataLatex() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('Data_Latex');
  if (!sh) {
    ss.toast('Data_Latex 시트를 찾을 수 없습니다.', '오류', 5);
    return;
  }

  var last = sh.getLastRow();
  if (last < 2) {
    ss.toast('Data_Latex에 데이터가 없습니다.', '오류', 5);
    return;
  }

  var ui = SpreadsheetApp.getUi();
  var resp = ui.alert(
    '선지 번호 원상복구',
    'Data_Latex C열의 줄 첫머리 (1)~(5)를 ①~⑤로 변환합니다.\n' +
    '실행 전 스프레드시트 사본을 백업했습니까?',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp !== ui.Button.OK) return;

  var rng = sh.getRange(2, 3, last - 1, 1);   // C열
  var vals = rng.getValues();
  var MAP = { '1': '①', '2': '②', '3': '③', '4': '④', '5': '⑤' };
  var changedDocs = 0, changedMarks = 0;

  for (var i = 0; i < vals.length; i++) {
    var s = String(vals[i][0] || '');
    if (!s) continue;

    // 줄 시작의 "(n) " 또는 "(n)\t" 만 변환
    var t = s.replace(/(^|\n)\(([1-5])\)[ \t]?/g, function(m, pre, d) {
      changedMarks++;
      return pre + MAP[d] + ' ';
    });

    if (t !== s) { vals[i][0] = t; changedDocs++; }
  }

  if (changedDocs === 0) {
    ss.toast('변환 대상이 없습니다.', '완료', 5);
    return;
  }

  rng.setValues(vals);
  ss.toast(
    '선지 번호 변환 완료: ' + changedDocs + '개 문항, ' + changedMarks + '개 번호',
    '완료', 6
  );
  Logger.log('fixChoiceMarkersInDataLatex: %s docs, %s markers', changedDocs, changedMarks);
}
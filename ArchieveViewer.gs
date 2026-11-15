/*************************************************
 * 아카이브 결과 뷰어 (Dialog)
 * - 사용자 입력: 시작행-마지막행
 * - 아카이브 열 배치는 Que 와 같다: A id · B Latex · C chapter · D (비움) · E Claude · F GPT · G Gemini
 *   (2026-10-02 실측: 864행부터 이 배치. 863행 이전은 형식이 혼재 — D HTML·E 윤문 전문·F HTML / D·E·F 결과 셋 / E·F 두 칸 …)
 * - 좌: B(Latex)
 * - 우 [v4 D23]:
 *     row ≥ QV_LAYOUT_SWITCH_ROW(864): E/F/G 를 Claude / GPT / Gemini 로 표시
 *     row <  864: D~G 중 값이 있는 열을 "D열(옛)" 처럼 열 이름 그대로 표시 (회사 이름을 추측해 붙이지 않음)
 *     row ≥ 속성 QV_LAYOUT_SWITCH_ROW_2: E/F/G 를 Claude / 합의 / Gemini 로 표시 (kice-2.0.0 — F열의 뜻이 GPT 에서 합의로 바뀐 경계.
 *         전환 때 아카이브 증분 복사를 마친 뒤 마지막 행 + 1 을 스크립트 속성에 적는다. 속성이 없으면 이 구분은 쓰지 않는다)
 * - v2 행은 의견마다 채택/기각 버튼을 붙인다. 판정은 Feedback 시트에만 적는다 (셀은 고치지 않는다 — v5 §8-4)
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

  /** 두 번째 경계 행 — 스크립트 속성 QV_LAYOUT_SWITCH_ROW_2 (없거나 숫자가 아니면 0 = 구분 없음) */
  function switchRow2_() {
    const n = Number(PropertiesService.getScriptProperties().getProperty('QV_LAYOUT_SWITCH_ROW_2'));
    return (isFinite(n) && n >= QV_LAYOUT_SWITCH_ROW) ? Math.floor(n) : 0;
  }

  function getPayload() {
    const ss = SpreadsheetApp.getActive();
    const row2 = switchRow2_();
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
      if (row2 && row >= row2) {
        layout = 'v2';
        blocks = [
          { label: 'Claude', cls: 'claude', html: cell(v, 4) },
          { label: '합의',   cls: 'merged', html: cell(v, 5) },
          { label: 'Gemini', cls: 'gemini', html: cell(v, 6) }
        ].filter(b => b.html);
      } else if (row >= QV_LAYOUT_SWITCH_ROW) {
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
      switchRow2: row2,
      feedback: feedbackFor_(ss, items),
      range: { startRow, endRow },
      count: items.length,
      items
    };
  }

  /* ---------- Feedback (v5 §8-4, 부록 A-8) ----------
   * 셀은 제출 뒤 고치지 않는다. 채택/기각은 Feedback 시트(의견ID · 판정 · 메모 · 반영상태)에만 적는다. */
  const FB_SHEET = 'Feedback';
  const FB_HEADERS = ['의견ID', '판정', '메모', '반영상태'];
  const FB_VERDICTS = ['채택', '기각', ''];
  const ID_RE = /\[\[ID\]\]\s*([^<\n]+)/g;

  function fbSheet_(ss) {
    const sh = ss.getSheetByName(FB_SHEET) || ss.insertSheet(FB_SHEET);
    if (sh.getLastRow() < 1) sh.getRange(1, 1, 1, FB_HEADERS.length).setValues([FB_HEADERS]);
    return sh;
  }

  /** 보이는 행들의 의견ID 에 대한 판정 {의견ID: {verdict, memo, applied}} */
  function feedbackFor_(ss, items) {
    const want = {};
    items.forEach(it => {
      if (it.layout !== 'v2') return;
      it.blocks.forEach(b => { let m; ID_RE.lastIndex = 0; while ((m = ID_RE.exec(b.html))) want[m[1].trim()] = true; });
    });
    return feedbackByIds_(ss, want);
  }

  /** want = {의견ID: true} → {의견ID: {verdict, memo, applied}}. 사이드바(사이드바로보기.gs)도 쓴다. */
  function feedbackByIds_(ss, want) {
    const out = {};
    if (!Object.keys(want).length) return out;
    const sh = ss.getSheetByName(FB_SHEET);
    if (!sh || sh.getLastRow() < 2) return out;
    sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues().forEach(r => {
      const id = String(r[0] || '').trim();
      if (id && want[id]) out[id] = { verdict: String(r[1] || ''), memo: String(r[2] || ''), applied: String(r[3] || '') };
    });
    return out;
  }

  /** 의견 하나의 판정을 적는다. verdict = '채택' | '기각' | ''(판정 취소). 이미 반영된 판정은 바꾸지 않는다. */
  function setFeedback(id, verdict, memo) {
    id = String(id || '').trim();
    verdict = String(verdict || '');
    if (!id || id.split('#').length !== 3) return { ok: false, reason: '의견ID 형식이 아닙니다' };
    if (FB_VERDICTS.indexOf(verdict) < 0) return { ok: false, reason: '판정은 채택·기각만' };
    const lock = LockService.getDocumentLock();
    if (!lock.tryLock(10000)) return { ok: false, reason: '다른 기록이 진행 중입니다. 다시 눌러 주세요' };
    try {
      const sh = fbSheet_(SpreadsheetApp.getActive());
      const last = sh.getLastRow();
      let row = 0, cur = null;
      if (last >= 2) {
        const vals = sh.getRange(2, 1, last - 1, 4).getValues();
        for (let i = 0; i < vals.length; i++) if (String(vals[i][0] || '').trim() === id) { row = i + 2; cur = vals[i]; break; }
      }
      if (cur && String(cur[3] || '').trim()) return { ok: false, reason: '이미 반영된 판정입니다(' + cur[3] + ')' };
      const m = (memo === undefined || memo === null) ? (cur ? String(cur[2] || '') : '') : String(memo).slice(0, 2000);
      if (row) sh.getRange(row, 2, 1, 2).setValues([[verdict, m]]);
      else if (verdict || m) sh.getRange(Math.max(2, last + 1), 1, 1, 4).setValues([[id, verdict, m, '']]);
      return { ok: true, id: id, verdict: verdict, memo: m };
    } finally {
      lock.releaseLock();
    }
  }

  return { openDialog, getPayload, setFeedback, feedbackByIds: (want) => feedbackByIds_(SpreadsheetApp.getActive(), want) };
})();

function qv_setFeedback(id, verdict, memo) {
  return QV.setFeedback(id, verdict, memo);
}

function qv_openDialog() {
  QV.openDialog();
}

function qv_getPayload() {
  return QV.getPayload();
}

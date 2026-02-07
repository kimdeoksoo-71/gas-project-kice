/*************************************************
 * Mathpix 대량 변환 — 커서 없는 “다음 N개” 방식 (충돌-안전)
 * 전역 노출: mpb_onOpen, mpb_runRange, mpb_start, mpb_stop, mpb_reset
 *************************************************/
const _MPB = (function () {
  // ===== 내부 상수/설정 (전역 충돌 방지) =====
  const CFG = {
    SHEET_NAME: 'Data_Latex',
    EXPECTED_HEADER: ['filename','drive_link','latex','text','status','attempts','last_error','processed_at'],
    COLS: { filename:1, drive_link:2, latex:3, text:4, status:5, attempts:6, last_error:7, processed_at:8 },
    SPREADSHEET_ID: SpreadsheetApp.getActive().getId(),
    STOP_FLAG_PROP: 'MPB_STOP_FLAG',
    HANDLER_NAME: 'mpb__processNextBatchStatus', // 트리거 핸들러(글로벌)
    // 배치 파라미터
    TIME_BUDGET_MS: 5 * 60 * 1000,
    SAFETY_GAP_MS: 20 * 1000,
    MAX_ROWS_PER_RUN: 120,
    MAX_ATTEMPTS: 5,
    INPROGRESS_RETRY_MS: 30 * 60 * 1000
  };

  // ===== 내부 유틸 =====
  function getSheet_() {
    const ss = SpreadsheetApp.openById(CFG.SPREADSHEET_ID);
    return ss.getSheetByName(CFG.SHEET_NAME);
  }
  function ensureHeader_(sh) {
    const hdr = sh.getRange(1,1,1,CFG.EXPECTED_HEADER.length).getValues()[0];
    if (hdr.join('|') !== CFG.EXPECTED_HEADER.join('|')) {
      sh.getRange(1,1,1,CFG.EXPECTED_HEADER.length).setValues([CFG.EXPECTED_HEADER]);
    }
  }
  function writeRowResult_(sheet,row,{latex,text,status,attempts,last_error,processed_at}) {
    sheet.getRange(row, CFG.COLS.latex, 1, 6).setValues([[
      latex||'', text||'', status||'', attempts||0, last_error||'', processed_at||''
    ]]);
  }
  function getMathpixCreds_() {
    const sp = PropertiesService.getScriptProperties();
    const app_id  = (sp.getProperty('MATHPIX_APP_ID')  || '').trim();
    const app_key = (sp.getProperty('MATHPIX_APP_KEY') || '').trim();
    if (!app_id || !app_key) throw new Error('Mathpix 키가 비어 있습니다.');
    return { app_id, app_key };
  }
  function extractFileId_(s){ const m=(s||'').match(/[-\w]{25,}/); if(!m) throw new Error('fileId parse failed: '+s); return m[0]; }
  function callMathpixWithRetry_({url, creds, payload, tries=6, base=700, jitter=400, maxWait=15000}) {
    const opt = {
      method:'post', contentType:'application/json',
      headers:{ 'app_id':creds.app_id, 'app_key':creds.app_key },
      payload: JSON.stringify(payload), muteHttpExceptions:true
    };
    for (let i=0;i<tries;i++){
      const res = UrlFetchApp.fetch(url, opt);
      const code = res.getResponseCode();
      if (code>=200 && code<300) return JSON.parse(res.getContentText());
      if (code===401) throw new Error('401 Unauthorized: OCR API 키/조직/헤더 확인. 응답: '+res.getContentText());
      if ([429,500,502,503,504].includes(code) && i<tries-1) {
        const delay = Math.min(base*Math.pow(2,i)+Math.random()*jitter, maxWait);
        Utilities.sleep(delay); continue;
      }
      throw new Error('Mathpix error '+code+': '+res.getContentText());
    }
    throw new Error('Mathpix retry exhausted');
  }


  // ===== 수동 단발 처리 =====
  function runRange_() {
    const ui = SpreadsheetApp.getUi();
    const sh = getSheet_();
    if (!sh) { ui.alert(`시트 "${CFG.SHEET_NAME}" 없음`); return; }
    ensureHeader_(sh);

    const resp = ui.prompt('변환 범위', '예: 2-100 (단일 행이면 5)', ui.ButtonSet.OK_CANCEL);
    if (resp.getSelectedButton() !== ui.Button.OK) return;

    const input = (resp.getResponseText()||'').trim();
    let s,e, m = input.match(/^\s*(\d+)\s*-\s*(\d+)\s*$/);
    if (m) { s=+m[1]; e=+m[2]; } else if (/^\d+$/.test(input)) { s=+input; e=s; }
    else { ui.alert('형식 오류. 예: 2-100 또는 5'); return; }
    if (e < s) [s,e] = [e,s];

    const lock = LockService.getScriptLock(); lock.waitLock(30000);
    let success = 0, total = 0;
    try {
      const creds = getMathpixCreds_();
      s = Math.max(2, s); e = Math.min(sh.getLastRow(), e);
      if (e < s) { ui.alert('해당 구간에 데이터가 없습니다.'); return; }

      for (let row=s; row<=e; row++) {
        total++;
        try {
          const vals = sh.getRange(row,1,1,8).getValues()[0];
          const linkOrId = (vals[CFG.COLS.drive_link-1]||'').toString().trim();

          const attempts = Number(vals[CFG.COLS.attempts-1]||0) + 1;
          sh.getRange(row, CFG.COLS.attempts).setValue(attempts);
          sh.getRange(row, CFG.COLS.status).setValue('in_progress');
          sh.getRange(row, CFG.COLS.processed_at).setValue(new Date());

          if (!linkOrId) {
            writeRowResult_(sh,row,{latex:'',text:'',status:'error',attempts,last_error:'missing_drive_link',processed_at:new Date()});
            continue;
          }

          const fileId = extractFileId_(linkOrId);
          const blob = DriveApp.getFileById(fileId).getBlob();
          const dataUrl = 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());

          const result = callMathpixWithRetry_({
            url: 'https://api.mathpix.com/v3/text',
            creds,
            payload: {
              src: dataUrl,
              formats: ['text'],
              rm_spaces: true,
              math_inline_delimiters: ['$', '$'],
              math_block_delimiters: ['$$','$$'],
              enable_tables: true,
              confidence_threshold: 0.0
            }
          });

          const merged = result.text || '';
          sh.getRange(row, CFG.COLS.latex).setValue(merged);
          sh.getRange(row, CFG.COLS.text).setValue('');
          sh.getRange(row, CFG.COLS.status).setValue('done');
          sh.getRange(row, CFG.COLS.last_error).setValue('');
          sh.getRange(row, CFG.COLS.processed_at).setValue(new Date());
          success++;

        } catch (e) {
          writeRowResult_(sh,row,{latex:'',text:'',status:'error',
            attempts:Number(sh.getRange(row,CFG.COLS.attempts).getValue()||0)||1,
            last_error:String(e).slice(0,500), processed_at:new Date()});
        }
        Utilities.sleep(200);
      }
      ui.alert(`완료: ${success}/${total} 행 성공`);
    } finally { lock.releaseLock(); }
  }

  // ===== 배치: 커서 없이 “다음 N개” 선별 =====
  function start_() {
    const sp = PropertiesService.getScriptProperties();
    sp.deleteProperty(CFG.STOP_FLAG_PROP); // 킬스위치 해제

    // 중복 트리거 제거 후 생성
    const namesToKill = [CFG.HANDLER_NAME, 'mpb__processNextBatchStatus'];
    ScriptApp.getProjectTriggers().forEach(t => {
      if (namesToKill.includes(t.getHandlerFunction())) ScriptApp.deleteTrigger(t);
    });
    ScriptApp.newTrigger(CFG.HANDLER_NAME).timeBased().everyMinutes(1).create();

    try { SpreadsheetApp.getUi().alert('배치 시작: 매 회차 “다음 N개” 자동 처리'); } catch(_) {}
  }

  function stop_() {
    const sp = PropertiesService.getScriptProperties();
    sp.setProperty(CFG.STOP_FLAG_PROP, '1'); // 킬스위치 ON
    const namesToKill = [CFG.HANDLER_NAME, 'mpb__processNextBatchStatus'];
    let count = 0;
    ScriptApp.getProjectTriggers().forEach(t => {
      if (namesToKill.includes(t.getHandlerFunction())) { ScriptApp.deleteTrigger(t); count++; }
    });
    try { SpreadsheetApp.getUi().alert(`배치 중지 요청 완료. 삭제된 트리거: ${count}개`); } catch(_) {}
  }

  function reset_() {
    const sp = PropertiesService.getScriptProperties();
    sp.deleteProperty(CFG.STOP_FLAG_PROP);
    try { SpreadsheetApp.getUi().alert('배치 상태 초기화 완료'); } catch(_) {}
  }

  // ===== 트리거 핸들러 (글로벌에서 래핑 호출) =====
  function processNextBatchStatus_() {
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(30000)) return; // 동시 실행 방지

    try {
      const sp = PropertiesService.getScriptProperties();
      if (sp.getProperty(CFG.STOP_FLAG_PROP) === '1') return; // 즉시 종료

      const t0 = Date.now();
      const deadline = t0 + CFG.TIME_BUDGET_MS - CFG.SAFETY_GAP_MS;

      const sh = getSheet_();
      if (!sh) { stop_(); return; }
      ensureHeader_(sh);

      const lastRow = sh.getLastRow();
      if (lastRow < 2) { stop_(); return; }

      // 전체 스캔 → 처리대상 N개 선별
      const rng = sh.getRange(2, 1, lastRow - 1, 8).getValues();
      const now = Date.now();
      const targets = [];
      for (let i = 0; i < rng.length; i++) {
        const rowIdx = i + 2;
        const vals = rng[i];
        const status = (vals[CFG.COLS.status - 1] || '').toString().trim();      // E
        const attempts = Number(vals[CFG.COLS.attempts - 1] || 0);               // F
        const processedAt = new Date(vals[CFG.COLS.processed_at - 1] || 0).getTime(); // H
        const linkOrId = (vals[CFG.COLS.drive_link - 1] || '').toString().trim();// B

        if (status === 'done') continue;

        if (status === 'in_progress') {
          const age = now - (isNaN(processedAt) ? 0 : processedAt);
          if (age >= 0 && age < CFG.INPROGRESS_RETRY_MS) continue;
        }

        if (attempts >= CFG.MAX_ATTEMPTS) continue;

        if (!linkOrId) { targets.push({ row: rowIdx, mode: 'error_missing' }); }
        else { targets.push({ row: rowIdx, mode: 'normal', linkOrId }); }

        if (targets.length >= CFG.MAX_ROWS_PER_RUN) break;
      }

      if (targets.length === 0) { stop_(); return; }

      const creds = getMathpixCreds_();
      let processed = 0;

      for (const t of targets) {
        if (sp.getProperty(CFG.STOP_FLAG_PROP) === '1') break;
        if (Date.now() > deadline) break;

        const row = t.row;

        try {
          const attempts = Number(sh.getRange(row, CFG.COLS.attempts).getValue() || 0) + 1;
          sh.getRange(row, CFG.COLS.attempts).setValue(attempts);
          sh.getRange(row, CFG.COLS.status).setValue('in_progress');
          sh.getRange(row, CFG.COLS.processed_at).setValue(new Date());

          if (t.mode === 'error_missing') {
            writeRowResult_(sh, row, {
              latex:'', text:'', status:'error', attempts,
              last_error:'missing_drive_link', processed_at:new Date()
            });
            processed++;
            continue;
          }

          const fileId = extractFileId_(t.linkOrId);
          const blob = DriveApp.getFileById(fileId).getBlob();
          const dataUrl = 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());

          const result = callMathpixWithRetry_({
            url: 'https://api.mathpix.com/v3/text',
            creds,
            payload: {
              src: dataUrl,
              formats: ['text'],
              rm_spaces: true,
              math_inline_delimiters: ['$', '$'],
              math_block_delimiters: ['$$', '$$'],
              enable_tables: true,
              confidence_threshold: 0.0
            }
          });

          const merged = result.text || '';
          sh.getRange(row, CFG.COLS.latex).setValue(merged);
          sh.getRange(row, CFG.COLS.text).setValue('');
          sh.getRange(row, CFG.COLS.status).setValue('done');
          sh.getRange(row, CFG.COLS.last_error).setValue('');
          sh.getRange(row, CFG.COLS.processed_at).setValue(new Date());

        } catch (e) {
          writeRowResult_(sh, row, {
            latex:'', text:'', status:'error',
            attempts:Number(sh.getRange(row, CFG.COLS.attempts).getValue() || 0) || 1,
            last_error:String(e).slice(0,500), processed_at:new Date()
          });
        }

        Utilities.sleep(300);
        processed++;
      }

      if (processed === 0) stop_();

    } finally {
      try { lock.releaseLock(); } catch (_) {}
    }
  }

  // 외부에 노출할 API
  return {
    runRange_: runRange_,
    start_: start_,
    stop_: stop_,
    reset_: reset_,
    processNextBatchStatus_: processNextBatchStatus_,
  };
})();

function mpb_runRange() {
  _MPB.runRange_();
}

function mpb_start() {
  _MPB.start_();
}

function mpb_stop() {
  _MPB.stop_();
}

function mpb_reset() {
  _MPB.reset_();
}

function mpb__processNextBatchStatus() {
  _MPB.processNextBatchStatus_();
}

// 가짜 GAS 런타임 — 시트는 2차원 배열, 속성은 객체
const fs = require('fs'), vm = require('vm'), path = require('path');
function makeEnv(repo) {
  const props = {};
  const books = {};                       // ssId -> {name, sheets:{name: rows[][]}}
  let clock = 1800000000000;
  function Sheet(book, name) {
    const data = () => book.sheets[name];
    const pad = (r, c) => { const d = data(); while (d.length < r) d.push([]); for (let i = 0; i < r; i++) while (d[i].length < c) d[i].push(''); };
    const api = {
      getName: () => name,
      getLastRow: () => { const d = data(); for (let i = d.length - 1; i >= 0; i--) if (d[i].some(v => v !== '' && v != null)) return i + 1; return 0; },
      getLastColumn: () => { let m = 0; data().forEach(r => { for (let i = r.length - 1; i >= 0; i--) if (r[i] !== '' && r[i] != null) { m = Math.max(m, i + 1); break; } }); return m; },
      getMaxRows: () => Math.max(1000, data().length), getMaxColumns: () => 40,
      insertRowsAfter: () => {}, insertColumnsAfter: () => {},
      appendRow: (r) => { const n = api.getLastRow(); pad(n + 1, r.length); r.forEach((v, i) => data()[n][i] = v); },
      getRange: (row, col, nr, nc) => {
        if (typeof row === 'string') throw new Error('A1 미지원');
        nr = nr || 1; nc = nc || 1;
        return {
          getValues: () => { pad(row + nr - 1, col + nc - 1); const o = []; for (let i = 0; i < nr; i++) { const rr = []; for (let j = 0; j < nc; j++) rr.push(data()[row - 1 + i][col - 1 + j]); o.push(rr); } return o; },
          getValue: () => { pad(row, col); return data()[row - 1][col - 1]; },
          setValues: (v) => { if (v.length !== nr || v.some(r => r.length !== nc)) throw new Error(`setValues 크기 불일치 ${v.length}x${v[0] && v[0].length} vs ${nr}x${nc}`); pad(row + nr - 1, col + nc - 1); for (let i = 0; i < nr; i++) for (let j = 0; j < nc; j++) data()[row - 1 + i][col - 1 + j] = v[i][j]; },
          setValue: (v) => { pad(row, col); data()[row - 1][col - 1] = v; },
          clearContent: () => { pad(row + nr - 1, col + nc - 1); for (let i = 0; i < nr; i++) for (let j = 0; j < nc; j++) data()[row - 1 + i][col - 1 + j] = ''; },
        };
      },
    };
    return api;
  }
  function Book(id) {
    const b = books[id];
    return { getName: () => b.name, getSheetByName: (n) => b.sheets[n] ? Sheet(b, n) : null,
             insertSheet: (n) => { b.sheets[n] = []; return Sheet(b, n); }, toast: () => {} };
  }
  let lockHeld = false, lockFail = false;
  const sandbox = {
    console: { log: () => {} }, JSON, Math, Date: class extends Date { constructor(...a) { if (a.length) super(...a); else super(clock); } static now() { return clock; } },
    PropertiesService: { getScriptProperties: () => ({ getProperty: k => (k in props ? props[k] : null), setProperty: (k, v) => { props[k] = String(v); }, deleteProperty: k => { delete props[k]; } }),
                         getUserProperties: () => ({ getProperty: () => null, setProperty: () => {} }) },
    LockService: { getScriptLock: () => ({ tryLock: () => { if (lockFail || lockHeld) return false; lockHeld = true; return true; }, releaseLock: () => { lockHeld = false; } }) },
    SpreadsheetApp: { getActive: () => Book('main'), openById: (id) => { if (!books[id]) throw new Error('no ss ' + id); return Book(id); }, flush: () => {}, getUi: () => { throw new Error('no ui'); } },
    Utilities: { getUuid: () => 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, () => Math.floor(Math.random() * 16).toString(16)), base64Encode: () => 'BASE64' },
    ContentService: { createTextOutput: (s) => ({ _t: s, setMimeType() { return this; } }), MimeType: { JSON: 'json' } },
    ScriptApp: { getProjectTriggers: () => [] }, DriveApp: {}, HtmlService: {},
  };
  vm.createContext(sandbox);
  ['kiceAuto.gs', 'kiceData.gs', 'batchKiceMulti.gs', 'RemoteApi.gs', 'ArchieveViewer.gs', 'mainmenu.gs'].forEach(f => {
    let src = fs.readFileSync(path.join(repo, f), 'utf8').replace(/^const (RAPI|QV) =/m, 'var $1 =');
    vm.runInContext(src, sandbox, { filename: f });
  });
  return { sandbox, props, books, tick: (ms) => { clock += ms; }, setLockFail: (v) => { lockFail = v; },
           get: (p) => JSON.parse(sandbox.doGet({ parameter: p })._t || 'null'),
           post: (body, q) => JSON.parse(sandbox.doPost({ parameter: q || {}, postData: { contents: JSON.stringify(body) } })._t || 'null') };
}
module.exports = { makeEnv };

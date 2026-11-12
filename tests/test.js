const { makeEnv } = require('./mock.js');
const REPO = process.argv[2];
let pass = 0, fail = 0;
function ok(name, cond, extra) { if (cond) pass++; else { fail++; console.log('  ✗', name, extra !== undefined ? JSON.stringify(extra).slice(0, 300) : ''); } }
const TOK = 'a'.repeat(40);
function fresh(nRows) {
  const env = makeEnv(REPO);
  env.props.REMOTE_TOKEN = TOK; env.props.LATEX_SS_ID = 'latex'; env.props.LATEX_KICE_SS_ID = 'klatex'; env.props.LEGACY_SS_ID = 'legacy';
  env.books.main = { name: 'KICE 윤문', sheets: {
    Que: [['id', 'Latex', '', '', 'claude', 'gpt', 'gemini']],
    '아카이브': [['id', 'Latex', '', '', 'c', 'g', 'm'], ['OLD_1공통01', 'old B', '', '', 'e1', 'f1', 'g1']],
    KICE_Log: [['time', 'run', 'stage', 'message']],
    QueCtx: [['사설 문항 id', '해설', '그림 링크', '그림 요약', '조건→역할', '조건 유형 태그', '2028범위 판정']],
    Opinions: [['의견ID', '문항 id', '대상 구절', '신뢰도', 'claude_제안', 'gemini_제안', '카테고리', '이유', '근거', '충돌', '힌트 누출', '답 일치 검증', '모델ID', '사고 수준', '프롬프트 버전']],
    Corpus: [['문항ID', '파이프라인 키', 'LaTeX', '정답', '대단원', '잠금']],
    AnswerKey: [['문항ID', '정답', '배점', '출처']], ExprDict: [['표현', '비권장 변형']], EvalSet: [['문항ID', '구분']],
    prompt: [['key', 'value', 'enabled'], ['SYSTEM_KICE_CLAUDE', 'P1', true], ['SYSTEM_KICE_GPT', 'P2', false]],
    Rule: [['from', 'to', 'enabled', 'note'], ['미분불가능한', '미분가능하지 않은', true, 'n'], ['x', 'x', true, '']],
  } };
  const ds = [['id', 'problem', 'given_solution', 'given_answer']];
  ds.push(['다른세트_1공통01', 'other', 'sol', '1']);
  for (let i = 1; i <= nRows; i++) ds.push([`26K99_1공통${String(i).padStart(2, '0')}`, `문제 ${i} ![f](https://drive.google.com/file/d/AAAAAAAAAAAAAAAAAAAAAAAAAAAA${i}/view)`, `해설 ${i}`, String(i)]);
  ds.push(['26K99_1공통99', '', 'sol', '']);   // 빈 B
  env.books.latex = { name: 'Latex 변환', sheets: { Data_DS: ds } };
  env.books.klatex = { name: 'KICE_latex', sheets: { Data_DS: [['id', 'problem'], ['2609평가원_1공통01', 'p1'], ['2609평가원_1공통02', 'p2'], ['2511수능_1공통01', 'p3']] } };
  env.books.legacy = { name: '기출과 윤문', sheets: { Data_Latex: [['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'ID'], ['x', 'y', '', '', '', '', '', '', '26091공통01']] } };
  const g = (p) => env.get(Object.assign({ token: TOK }, p));
  const po = (b) => env.post(Object.assign({ token: TOK }, b));
  return { env, g, po, que: () => env.books.main.sheets.Que, arc: () => env.books.main.sheets['아카이브'] };
}
function cellsOf(i) { return { E: `c-edit ${i}<br>[[원본]] a`, F: i % 2 ? '합의 없음' : `[합의] x<br>y ${i}`, G: i % 3 ? '수정 구절 없음' : `g-edit ${i}` }; }

console.log('S1 인증·ping·tick(상태 없음)');
{ const { env, g } = fresh(3);
  ok('토큰 없으면 빈 응답', env.get({ cmd: 'ping' }) === null);
  ok('틀린 토큰 POST 빈 응답', env.post({ token: 'x', cmd: 'ping' }) === null);
  ok('ping 버전', g({ cmd: 'ping' }).version === 'kice-2.0.0');
  ok('tick 상태 없음 ok:false', g({ cmd: 'tick' }).ok === false);
  ok('claim 상태 없음 not_running', g({ cmd: 'claim' }).reason === 'not_running');
  ok('unknown cmd', g({ cmd: 'zz' }).ok === false);
}

console.log('S2 정상 흐름: start → claim → heartbeat → submit → finish → 아카이브');
{ const { env, g, po, que, arc } = fresh(5);
  const s = g({ cmd: 'start', keywords: '26K99', runId: 'run-1', expect: '6' });
  ok('start ok rewrite 5행', s.ok && s.stage === 'rewrite' && s.rows === 5, s);
  ok('src 집계', s.src.matched === 5 && s.src.emptyB === 1 && s.src.countMismatch === false, s.src);
  ok('Que 적재', que()[1][0] === '26K99_1공통01' && que()[5][0] === '26K99_1공통05');
  ok('QueCtx 헤더에 given_answer 추가', env.books.main.sheets.QueCtx[0].includes('given_answer'));
  let st = g({ cmd: 'status' });
  ok('status 계약 필드', st.state.startedAt === s.startedAt && st.state.stage === 'rewrite' && st.derived === null && st.progress.idx === 0 && st.progress.total === 5 && st.progress.currentRow === null && typeof st.progress.lastHeartbeat === 'number' && st.progress.poisonCount === 0, st.progress);
  let t = g({ cmd: 'tick' });
  ok('tick ok:true idle rewrite', t.ok === true && t.result === 'idle' && t.stage === 'rewrite', t);
  ok('start 멱등 adopted', g({ cmd: 'start', keywords: '26K99', runId: 'run-1' }).adopted === true);
  const busy = g({ cmd: 'start', keywords: '26K99', runId: 'run-2' });
  ok('다른 런 start 는 "이미 실행 중"', busy.ok === false && busy.reason.includes('이미 실행 중'), busy);
  const c = g({ cmd: 'claim', n: '3', engineId: 'e1' });
  ok('claim 3행', c.ok && c.items.length === 3 && c.items[0].row === 2 && c.runId === 'run-1' && c.remaining === 5, c);
  ok('claim 항목: 해설·그림·정답', c.items[0].solution === '해설 1' && c.items[0].figLinks.length === 1 && c.items[0].givenAnswer === '1' && c.items[0].lease.length === 12, c.items[0]);
  const c2 = g({ cmd: 'claim', n: '3', engineId: 'e2' });
  ok('두 번째 claim 은 다른 행', c2.items.length === 2 && c2.items[0].row === 5 && c2.inflight === 5, c2);
  st = g({ cmd: 'status' });
  ok('status currentRow=2, inflight 5', st.progress.currentRow === 2 && st.progress.inflight === 5, st.progress);
  env.tick(3 * 60 * 1000);
  const hb = g({ cmd: 'heartbeat', runId: 'run-1', startedAt: s.startedAt, leases: c.items.map(i => i.row + ':' + i.lease).join(',') + ',6:WRONG' });
  ok('heartbeat alive 3 lost 1', hb.ok && hb.alive.length === 3 && hb.lost.length === 1 && hb.lost[0] === 6, hb);
  ok('finish 는 남은 행 때문에 거절', g({ cmd: 'finish', runId: 'run-1', startedAt: s.startedAt }).reason === 'remaining');
  const all = c.items.concat(c2.items);
  all.forEach((it, i) => {
    const r = po({ cmd: 'submit', runId: 'run-1', startedAt: s.startedAt, row: it.row, lease: it.lease, id: it.id, cells: cellsOf(i),
                   opinions: [{ id: it.id + '#aaaaaa#' + i, target: 't', confidence: '합의', claude: 'c', gemini: 'g', category: '표기·문체' }], ctx: { roles: { a: 1 }, tags: 'T1', scope: '범위 안' } });
    ok('submit row ' + it.row, r.ok && r.opinions === 1 && !r.warn, r);
  });
  ok('Que E·F·G 기록', que()[1][4].startsWith('c-edit 0') && que()[1][5].startsWith('[합의]') && que()[2][5] === '합의 없음');
  const dupRe = po({ cmd: 'submit', runId: 'run-1', startedAt: s.startedAt, row: all[0].row, lease: all[0].lease, id: all[0].id, cells: cellsOf(0) });
  ok('같은 내용 재제출 = duplicate ok', dupRe.ok && dupRe.duplicate === true, dupRe);
  const stale = po({ cmd: 'submit', runId: 'run-1', startedAt: s.startedAt, row: all[0].row, lease: all[0].lease, id: all[0].id, cells: cellsOf(9) });
  ok('다른 내용 재제출 = stale_lease', stale.reason === 'stale_lease', stale);
  st = g({ cmd: 'status' });
  ok('status idx=5', st.progress.idx === 5 && st.progress.currentRow === null, st.progress);
  ok('Opinions 5행', env.books.main.sheets.Opinions.length === 6 && env.books.main.sheets.Opinions[1][1] === '26K99_1공통01');
  const ctxH = env.books.main.sheets.QueCtx[0];
  ok('QueCtx 분석 기록', env.books.main.sheets.QueCtx[1][ctxH.indexOf('조건 유형 태그')] === 'T1' && env.books.main.sheets.QueCtx[1][ctxH.indexOf('해설')] === '해설 1');
  const c3 = g({ cmd: 'claim' });
  ok('남은 것 없는 claim', c3.ok && c3.items.length === 0 && c3.remaining === 0 && c3.inflight === 0, c3);
  const f = g({ cmd: 'finish', runId: 'run-1', startedAt: s.startedAt });
  ok('finish done + 아카이브 3~7', f.ok && f.stage === 'done' && f.archive.start === 3 && f.archive.end === 7, f);
  ok('summary 키 유지', f.summary.rows === 5 && f.summary.claude.edit === 5 && f.summary.gpt.edit === 3 && f.summary.gpt.none === 2 && f.summary.gemini.none === 3 && 'poisonCount' in f.summary && 'outOfScope' in f.summary, f.summary);
  ok('아카이브 내용', arc()[2][0] === '26K99_1공통01' && arc()[6][4].startsWith('c-edit 4'));
  st = g({ cmd: 'status' });
  ok('끝난 뒤 status: done, idx=5(=처리 수), archive·src·residue', st.state.stage === 'done' && st.progress.idx === 5 && st.state.archive.start === 3 && st.state.src.matched === 5 && st.state.residue.archived === 0 && st.state.summary.rows === 5, st.progress);
  t = g({ cmd: 'tick' });
  ok('끝난 뒤 tick: ok idle done', t.ok && t.result === 'idle' && t.stage === 'done', t);
  ok('finish 재호출 already', g({ cmd: 'finish', runId: 'run-1', startedAt: s.startedAt }).already === true);
  ok('끝난 뒤 submit not_running', po({ cmd: 'submit', runId: 'run-1', startedAt: s.startedAt, row: 2, lease: 'x', id: 'x', cells: cellsOf(1) }).reason === 'not_running');
  // 재런: 잔재 이관 뒤 기윤문 판정 → 전부 alreadyDone, 중복 이관 없음
  const s2 = g({ cmd: 'start', keywords: '26K99', runId: 'run-1~r1', expect: '6' });
  ok('재런: 0행 done, alreadyDone 5, 잔재는 중복 제외', s2.ok && s2.stage === 'done' && s2.rows === 0 && s2.src.alreadyDone === 5 && s2.residue.archived === 0 && s2.residue.dup === 5, s2);
  ok('아카이브 행 수 그대로(7)', arc().filter(r => r[0]).length === 7);
}

console.log('S3 stop 뒤 늦은 submit · 잔재 이관 순서(④⑤)');
{ const { env, g, po, que, arc } = fresh(4);
  const s = g({ cmd: 'start', keywords: '26K99', runId: 'r3' });
  const c = g({ cmd: 'claim', n: '2' });
  po({ cmd: 'submit', runId: 'r3', startedAt: s.startedAt, row: 2, lease: c.items[0].lease, id: c.items[0].id, cells: cellsOf(0) });
  const sp = g({ cmd: 'stop' });
  ok('stop: cleared, done=1', sp.ok && sp.cleared && sp.done === 1 && sp.stoppedFrom === 'rewrite', sp);
  let st = g({ cmd: 'status' });
  ok('stop 뒤 status stopped, idx=1', st.state.stage === 'stopped' && st.progress.idx === 1 && st.derived === null, st.progress);
  ok('stop 뒤 tick: stage stopped', g({ cmd: 'tick' }).stage === 'stopped');
  ok('늦은 submit 거부 not_running', po({ cmd: 'submit', runId: 'r3', startedAt: s.startedAt, row: 3, lease: c.items[1].lease, id: c.items[1].id, cells: cellsOf(1) }).reason === 'not_running');
  ok('늦은 heartbeat 거부', g({ cmd: 'heartbeat', runId: 'r3', startedAt: s.startedAt, leases: '3:x' }).reason === 'not_running');
  ok('늦은 claim 거부', g({ cmd: 'claim' }).reason === 'not_running');
  const s2 = g({ cmd: 'start', keywords: '26K99', runId: 'r3b' });
  ok('다음 런: 잔재 1행 이관 후 그 문항은 기윤문 → 3행만 적재', s2.ok && s2.residue.archived === 1 && s2.residue.dropped === 3 && s2.src.alreadyDone === 1 && s2.rows === 3, s2);
  ok('옛 세대 submit 은 not_owner', po({ cmd: 'submit', runId: 'r3', startedAt: s.startedAt, row: 2, lease: 'x', id: que()[1][0], cells: cellsOf(1) }).reason === 'not_owner');
}

console.log('S4 lease 만료 → 재claim → 옛 submit stale · 3회째 포기');
{ const { env, g, po, que } = fresh(2);
  const s = g({ cmd: 'start', keywords: '26K99', runId: 'r4' });
  const c1 = g({ cmd: 'claim', n: '1' });
  env.tick(10 * 60 * 1000);
  ok('10분: 아직 재claim 안 됨(다른 행을 줌)', g({ cmd: 'claim', n: '1' }).items[0].row === 3);
  env.tick(16 * 60 * 1000);
  const c2 = g({ cmd: 'claim', n: '1' });
  ok('16분 뒤 재claim, attempt 2, 새 lease', c2.items[0].row === 2 && c2.items[0].attempt === 2 && c2.items[0].lease !== c1.items[0].lease, c2.items);
  ok('옛 lease submit = stale_lease', po({ cmd: 'submit', runId: 'r4', startedAt: s.startedAt, row: 2, lease: c1.items[0].lease, id: c1.items[0].id, cells: cellsOf(0) }).reason === 'stale_lease');
  env.tick(16 * 60 * 1000);
  const c3 = g({ cmd: 'claim', n: '1' });
  ok('attempt 3 허용', c3.items.length === 1 && c3.items[0].attempt === 3, c3.items);
  env.tick(16 * 60 * 1000);
  const c4 = g({ cmd: 'claim', n: '5' });
  ok('4번째 차례: (실패) 확정, 그 행은 주지 않음', que()[1][4].startsWith('(실패)') && !c4.items.some(i => i.row === 2), c4);
  const st = g({ cmd: 'status' });
  ok('poisonCount 1', st.progress.poisonCount === 1, st.progress);
}

console.log('S5 인수(15분 정지) · archive 재진입 · 잠금');
{ const { env, g, po } = fresh(2);
  const s = g({ cmd: 'start', keywords: '26K99', runId: 'r5' });
  env.tick(16 * 60 * 1000);
  const s2 = g({ cmd: 'start', keywords: '26K99', runId: 'r5-next' });
  ok('heartbeat 16분 정지 → 다음 런이 인수', s2.ok && s2.stage === 'rewrite' && s2.rows === 2, s2);
  const c = g({ cmd: 'claim', n: '5' });
  c.items.forEach((it, i) => po({ cmd: 'submit', runId: 'r5-next', startedAt: s2.startedAt, row: it.row, lease: it.lease, id: it.id, cells: cellsOf(i) }));
  // 아카이브 도중 죽은 상황을 만든다: stage 를 archive 로 바꿔 둠
  const a = JSON.parse(env.props.KICE_AUTO_STATE_V2); a.stage = 'archive'; env.props.KICE_AUTO_STATE_V2 = JSON.stringify(a);
  ok('archive 15분 안: archiving 으로 거절', g({ cmd: 'finish', runId: 'r5-next', startedAt: s2.startedAt }).reason === 'archiving');
  env.tick(16 * 60 * 1000);
  const f = g({ cmd: 'finish', runId: 'r5-next', startedAt: s2.startedAt });
  ok('archive 재진입 완료', f.ok && f.stage === 'done' && f.archive.start === 3, f);
  env.setLockFail(true);
  ok('잠금 실패: claim locked', g({ cmd: 'claim' }).reason === 'locked' || g({ cmd: 'claim' }).reason === 'not_running');
  ok('잠금 실패: start 는 "이미 실행 중"', g({ cmd: 'start', keywords: 'x', runId: 'z' }).reason.includes('이미 실행 중'));
  ok('잠금 실패여도 tick·status 는 응답', g({ cmd: 'tick' }).ok === true && g({ cmd: 'status' }).ok === true);
  env.setLockFail(false);
}

console.log('S6 submit 검사 · 분류');
{ const { env, g, po } = fresh(2);
  const s = g({ cmd: 'start', keywords: '26K99', runId: 'r6' });
  const c = g({ cmd: 'claim', n: '2' });
  const it = c.items[0];
  ok('E 빈 값 거절', po({ cmd: 'submit', runId: 'r6', startedAt: s.startedAt, row: it.row, lease: it.lease, id: it.id, cells: { E: '', F: '', G: 'x' } }).ok === false);
  ok('id 불일치 row_mismatch', po({ cmd: 'submit', runId: 'r6', startedAt: s.startedAt, row: it.row, lease: it.lease, id: 'WRONG', cells: cellsOf(0) }).reason === 'row_mismatch');
  ok('배치 밖 행', po({ cmd: 'submit', runId: 'r6', startedAt: s.startedAt, row: 99, lease: 'x', id: 'x', cells: cellsOf(0) }).reason === 'row_not_in_batch');
  ok('startedAt 불일치 not_owner', po({ cmd: 'submit', runId: 'r6', startedAt: 'zzz', row: it.row, lease: it.lease, id: it.id, cells: cellsOf(0) }).reason === 'not_owner');
  po({ cmd: 'submit', runId: 'r6', startedAt: s.startedAt, row: it.row, lease: it.lease, id: it.id, cells: { E: '범위 밖', F: '범위 밖', G: '범위 밖' } });
  const it2 = c.items[1];
  po({ cmd: 'submit', runId: 'r6', startedAt: s.startedAt, row: it2.row, lease: it2.lease, id: it2.id, cells: { E: '(실패) 3회 오류', F: '', G: '[분석 없음]<br>[[원본]] a' } });
  const f = g({ cmd: 'finish', runId: 'r6', startedAt: s.startedAt });
  ok('분류: 범위 밖=none·outOfScope 1, (실패)=fail, F 빈칸→합의 없음', f.summary.outOfScope === 1 && f.summary.claude.none === 1 && f.summary.claude.fail === 1 && f.summary.gpt.none === 2 && f.summary.gemini.edit === 1, f.summary);
  const C = env.sandbox.kice_classifyCell_;
  ok('classify 기본', C('') === 'empty' && C('수정 구절 없음') === 'none' && C('수정구절 없음\n') === 'none' && C('합의 없음<br>') === 'none' && C('(실패) x') === 'fail' && C('[합의] a') === 'edit');
}

console.log('S7 데이터 명령: prompts · corpus · dsread · corpus_put');
{ const { env, g, po } = fresh(1);
  const pr = g({ cmd: 'prompts' });
  ok('prompts: 활성만, 규칙 1건', pr.ok && pr.prompts.SYSTEM_KICE_CLAUDE === 'P1' && !('SYSTEM_KICE_GPT' in pr.prompts) && pr.rules.length === 1, pr);
  const d = g({ cmd: 'dsread', src: 'latex', keywords: '2609평가원' });
  ok('dsread latex 키워드 2행', d.ok && d.total === 2 && d.rows[0].values[0] === '2609평가원_1공통01' && d.next === null, d);
  ok('dsread legacy col=9', g({ cmd: 'dsread', src: 'legacy', keywords: '2609', col: '9' }).total === 1);
  ok('dsread 잘못된 src', g({ cmd: 'dsread', src: 'x' }).ok === false);
  const dry = po({ cmd: 'corpus_put', sheet: 'Corpus', dryRun: true, rows: [{ '문항ID': '2026-09-C-01', 'LaTeX': 'L1', '정답': 3 }] });
  ok('dryRun 은 쓰지 않음', dry.ok && dry.inserted === 1 && env.books.main.sheets.Corpus.length === 1, dry);
  const p1 = po({ cmd: 'corpus_put', sheet: 'Corpus', rows: [{ '문항ID': '2026-09-C-01', 'LaTeX': 'L1', '정답': 3 }, { '문항ID': '2026-09-C-02', 'LaTeX': 'L2' }] });
  ok('insert 2', p1.ok && p1.inserted === 2 && env.books.main.sheets.Corpus[1][2] === 'L1', p1);
  const p2 = po({ cmd: 'corpus_put', sheet: 'Corpus', rows: [{ '문항ID': '2026-09-C-01', '대단원': '11' }, { '문항ID': '2026-09-C-02', 'LaTeX': 'L2' }] });
  ok('지정 열만 갱신, 나머지 유지, 같은 값은 unchanged', p2.updated === 1 && p2.unchanged === 1 && env.books.main.sheets.Corpus[1][2] === 'L1' && env.books.main.sheets.Corpus[1][4] === '11', p2);
  env.books.main.sheets.Corpus[1][5] = 'Y';   // 사람이 잠금
  const p3 = po({ cmd: 'corpus_put', sheet: 'Corpus', rows: [{ '문항ID': '2026-09-C-01', 'LaTeX': 'CHANGED', '파이프라인 키': 'k1' }] });
  ok('잠금 행: LaTeX 보호, 다른 열은 갱신', p3.lockedSkipped.length === 1 && env.books.main.sheets.Corpus[1][2] === 'L1' && env.books.main.sheets.Corpus[1][1] === 'k1', p3);
  const p4 = po({ cmd: 'corpus_put', sheet: 'Corpus', force: true, rows: [{ '문항ID': '2026-09-C-01', 'LaTeX': 'FORCED' }] });
  ok('force 면 갱신', p4.updated === 1 && env.books.main.sheets.Corpus[1][2] === 'FORCED', p4);
  ok('모르는 열 이름 거부', po({ cmd: 'corpus_put', sheet: 'Corpus', rows: [{ '문항ID': 'x', '없는열': 1 }] }).ok === false && env.books.main.sheets.Corpus.length === 3);
  ok('허용 안 된 시트 거부', po({ cmd: 'corpus_put', sheet: 'Que', rows: [{}] }).ok === false);
  const big = po({ cmd: 'corpus_put', sheet: 'Corpus', rows: [{ '문항ID': '2026-09-C-03', 'LaTeX': 'x'.repeat(60000) }] });
  ok('5만 자 초과 잘림 표시', big.truncated.length === 1 && env.books.main.sheets.Corpus[3][2].length < 50000, big.truncated);
  const co = g({ cmd: 'corpus', sheet: 'Corpus', cols: '문항ID,정답', limit: '2' });
  ok('corpus 쪽 읽기·열 선택·etag', co.ok && co.headers.length === 2 && co.rows.length === 2 && co.total === 3 && co.next === 2 && /:\d+$/.test(co.etag), co);
  ok('corpus 없는 열 이름', g({ cmd: 'corpus', cols: '없음' }).ok === false);
  ok('ExprDict 키=표현', po({ cmd: 'corpus_put', sheet: 'ExprDict', rows: [{ '표현': '부분의 넓이', '비권장 변형': '둘러싸인 도형의 넓이' }] }).inserted === 1);
  ok('bad JSON', JSON.parse(env.sandbox.doPost({ parameter: { token: TOK }, postData: { contents: '{bad' } })._t).reason.startsWith('bad_json'));
}

console.log('S8 메뉴 배치(상태만) — finish 는 아카이브로 옮기지 않는다');
{ const { env, g, po, que, arc } = fresh(1);
  que().push(['단건_1공통01', 'B1', '', '', '', '', '']); que().push(['단건_1공통02', 'B2', '', '', '', '', '']);
  const st0 = '2026-11-13T00:00:00.000Z';
  env.props.QUE_BATCH_STATE_V4 = JSON.stringify({ startedAt: st0, kice: 'menu~1', menu: true, rows: [2, 3], inflight: {}, poison: [], poisonCount: 0 });
  env.props.KICE_AUTO_STATE_V2 = JSON.stringify({ runId: 'menu~1', stage: 'rewrite', startedAt: st0, hb: Date.now() + 1e15, total: 2, idx: 0, menu: true });
  const a = JSON.parse(env.props.KICE_AUTO_STATE_V2); a.hb = env.sandbox.kice_now_(); env.props.KICE_AUTO_STATE_V2 = JSON.stringify(a);
  const busy = g({ cmd: 'start', keywords: '26K99', runId: 'rr' });
  ok('메뉴 배치 중 러너 start = 이미 실행 중', busy.ok === false && busy.reason.includes('이미 실행 중') && busy.reason.includes('메뉴'), busy);
  const c = g({ cmd: 'claim', n: '5' });
  ok('메뉴 배치 claim: 해설 없음(QueCtx 미일치)', c.menu === true && c.items.length === 2 && c.items[0].solution === '', c);
  c.items.forEach((it, i) => po({ cmd: 'submit', runId: 'menu~1', startedAt: st0, row: it.row, lease: it.lease, id: it.id, cells: cellsOf(i), ctx: { tags: 'm' } }));
  const n0 = arc().filter(r => r[0]).length;
  const f = g({ cmd: 'finish', runId: 'menu~1', startedAt: st0 });
  ok('메뉴 finish: done, 아카이브 그대로, Que 에 결과', f.ok && f.archive.start === 0 && arc().filter(r => r[0]).length === n0 && que()[1][4].startsWith('c-edit'), f);
  const s = g({ cmd: 'start', keywords: '26K99', runId: 'after-menu' });
  ok('다음 러너 start 가 메뉴 결과 2행을 이관', s.ok && s.residue.archived === 2, s.residue);
  const sp = env.sandbox.kice_apiStop_('menu');
  const st = g({ cmd: 'status' });
  ok('메뉴 ⏹: 러너가 읽는 by=menu', sp.ok && st.state.stage === 'stopped' && st.state.by === 'menu', st.state);
}

console.log(`\n결과: 통과 ${pass} / 실패 ${fail}`);
process.exit(fail ? 1 : 0);

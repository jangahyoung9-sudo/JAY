/* =====================================================================
 * 병원 상담 통계 주간 대시보드 - 앱 코드
 * 구성: [1] 날짜 도구  [2] 엑셀 읽기/헤더 감지  [3] 행 정제  [4] 주차 계산 엔진
 *       [5] 내장 자가진단  [6] 화면(UI)  [7] 검증용 공개 함수
 * 원칙: 날짜는 연/월/일 '숫자'로만 다룹니다. (new Date('2026-10-02')는 UTC로 읽혀 하루 밀림)
 * ===================================================================== */
(function () {
'use strict';

/* ---------- [0] 작은 도구 ---------- */
const $ = (s) => document.querySelector(s);
const pad2 = (n) => String(n).padStart(2, '0');
const esc = (s) => String(s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const WD = ['일', '월', '화', '수', '목', '금', '토'];

/* =====================================================================
 * [1] 날짜 도구 — 모든 날짜를 "1970-01-01 = 0"인 정수(일 수)로 바꿔 계산합니다.
 *     타임존의 영향을 전혀 받지 않습니다.
 * ===================================================================== */
// 연/월/일 → 일 수
function daysFromCivil(y, m, d) {
  y -= m <= 2 ? 1 : 0;
  const era = Math.floor(y / 400);
  const yoe = y - era * 400;
  const doy = Math.floor((153 * (m + (m > 2 ? -3 : 9)) + 2) / 5) + d - 1;
  const doe = yoe * 365 + Math.floor(yoe / 4) - Math.floor(yoe / 100) + doy;
  return era * 146097 + doe - 719468;
}
// 일 수 → {y,m,d}
function civilFromDays(z) {
  z += 719468;
  const era = Math.floor(z / 146097);
  const doe = z - era * 146097;
  const yoe = Math.floor((doe - Math.floor(doe / 1460) + Math.floor(doe / 36524) - Math.floor(doe / 146096)) / 365);
  const doy = doe - (365 * yoe + Math.floor(yoe / 4) - Math.floor(yoe / 100));
  const mp = Math.floor((5 * doy + 2) / 153);
  const d = doy - Math.floor((153 * mp + 2) / 5) + 1;
  const m = mp + (mp < 10 ? 3 : -9);
  return { y: yoe + era * 400 + (m <= 2 ? 1 : 0), m, d };
}
const dowOf = (n) => (((n + 4) % 7) + 7) % 7;               // 0=일 … 3=수 … 6=토
const isoOf = (n) => { const c = civilFromDays(n); return c.y + '-' + pad2(c.m) + '-' + pad2(c.d); };
const mdOf = (n) => { const c = civilFromDays(n); return pad2(c.m) + '-' + pad2(c.d); };
const mdwOf = (n) => mdOf(n) + '(' + WD[dowOf(n)] + ')';
// 실제 달력에 있는 날짜인지 확인 후 일 수 반환 (없으면 null)
function ymdToDays(y, m, d) {
  if (!(y >= 1900 && y <= 2200 && m >= 1 && m <= 12 && d >= 1 && d <= 31)) return null;
  const n = daysFromCivil(y, m, d);
  const c = civilFromDays(n);
  return c.y === y && c.m === m && c.d === d ? n : null;
}
function parseISO(s) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s || '');
  return m ? ymdToDays(+m[1], +m[2], +m[3]) : null;
}
// 가장 가까운 수요일로 맞춤 (기본값/회의일 보정용)
function nearestWednesday(n) {
  const diff = 3 - dowOf(n);
  return n + (((diff + 10) % 7) - 3);
}

/* 날짜 정규화기: 엑셀 시리얼 / 20261002 / 2026.10.02 / 2026-10-02 14:30 / 2026년 10월 2일 …
 * 반환: {state:'ok', n} | {state:'empty'} | {state:'error'} */
function parseDateCell(v) {
  if (v === null || v === undefined) return { state: 'empty' };
  if (v instanceof Date) {                       // 방어용(보통 쓰이지 않음)
    const n = ymdToDays(v.getFullYear(), v.getMonth() + 1, v.getDate());
    return n === null ? { state: 'error' } : { state: 'ok', n };
  }
  if (typeof v === 'number') return fromNumber(v);
  const s = String(v).trim();
  if (s === '') return { state: 'empty' };
  if (/^\d+(\.\d+)?$/.test(s)) return fromNumber(parseFloat(s));   // "46000" 같은 숫자 텍스트
  const m = /(\d{4})\s*[년.\-\/]\s*(\d{1,2})\s*[월.\-\/]\s*(\d{1,2})/.exec(s);
  if (m) {
    const n = ymdToDays(+m[1], +m[2], +m[3]);
    if (n !== null) return { state: 'ok', n };
  }
  return { state: 'error' };
}
function fromNumber(x) {
  if (Number.isInteger(x) && x >= 19000101 && x <= 22001231) {   // 20261002 형태
    const n = ymdToDays(Math.floor(x / 10000), Math.floor(x / 100) % 100, x % 100);
    return n === null ? { state: 'error' } : { state: 'ok', n };
  }
  if (x >= 20000 && x < 80000) return { state: 'ok', n: Math.floor(x) - 25569 };  // 엑셀 시리얼(25569 = 1970-01-01)
  return { state: 'error' };
}

/* =====================================================================
 * [2] 엑셀 읽기 · 헤더 자동 감지
 * ===================================================================== */
const normText = (v) => (v === null || v === undefined ? '' : String(v)).replace(/\s+/g, '').toLowerCase();
const colLetter = (i) => { let s = ''; i++; while (i > 0) { const r = (i - 1) % 26; s = String.fromCharCode(65 + r) + s; i = Math.floor((i - 1) / 26); } return s; };

// 제목 키워드 → 어떤 종류의 열인지 (검사 순서가 중요: 환자번호 → 날짜 → 환자명 → 담당자)
const HEADER_RULES = [
  ['no',    /(환자|차트|등록|병록|진료)(번호|no\.?|id|코드)/],
  ['date',  /상담(일자|일시|일|날짜)|일자|날짜|내원일|방문일|^date$/],
  ['name',  /환자명|환자이름|환자성명|^환자$|성명|^이름$/],
  ['staff', /담당|상담자|상담사|상담원|직원|작성자/],
];
function classifyHeader(cell) {
  const t = normText(cell);
  if (!t || t.length > 20) return null;
  for (const [kind, re] of HEADER_RULES) if (re.test(t)) return kind;
  return null;
}
// 한 행을 헤더로 가정했을 때의 열 매핑 (없으면 -1)
function mappingForRow(row) {
  const map = { date: -1, no: -1, name: -1, staff: -1 };
  (row || []).forEach((cell, i) => {
    const k = classifyHeader(cell);
    if (k && map[k] < 0) map[k] = i;
  });
  return map;
}
const mappingOk = (m) => m.date >= 0 && (m.no >= 0 || m.name >= 0);
// 상위 20행 중 키워드가 가장 많이 걸린 행을 헤더로 판정
function detectHeader(grid) {
  let best = { row: -1, score: 0, map: null };
  for (let r = 0; r < Math.min(20, grid.length); r++) {
    const map = mappingForRow(grid[r]);
    const score = Object.values(map).filter((x) => x >= 0).length;
    if (score > best.score) best = { row: r, score, map };
  }
  if (best.row < 0) return { ok: false, row: 0, map: { date: -1, no: -1, name: -1, staff: -1 }, score: 0 };
  return { ok: best.score >= 2 && mappingOk(best.map), row: best.row, map: best.map, score: best.score };
}
/* ---- 컬럼 설정 기억: 열 '제목 글자'만 이 PC 브라우저(localStorage)에 저장 (환자 데이터는 저장 안 함) ---- */
const MEM_KEY = 'consultDash.columnMaps.v1';
function memLoad() { try { return JSON.parse(localStorage.getItem(MEM_KEY) || '[]'); } catch (e) { return []; } }
function memSave(list) { try { localStorage.setItem(MEM_KEY, JSON.stringify(list)); return true; } catch (e) { return false; } }
const sigOf = (row) => (row || []).map(normText).filter(Boolean).join('|');     // 제목줄 지문
function memFind(grid) {                                  // 저장된 양식과 같은 제목줄이 있으면 그 설정을 돌려줌
  const list = memLoad();
  if (!list.length) return null;
  for (let r = 0; r < Math.min(20, grid.length); r++) {
    const sig = sigOf(grid[r]);
    const e = sig.split('|').length >= 2 ? list.find((x) => x.sig === sig) : null;
    if (!e) continue;
    const cells = (grid[r] || []).map(normText), map = {};
    for (const k of ['date', 'no', 'name', 'staff']) map[k] = e.fields[k] ? cells.indexOf(e.fields[k]) : -1;
    if (mappingOk(map)) return { ok: true, row: r, map, score: 9, remembered: true };
  }
  return null;
}
function memRemember(grid, headerRow, map) {              // 수동 지정을 저장 (같은 제목줄이면 덮어씀)
  const header = (grid[headerRow] || []).map(normText), sig = sigOf(grid[headerRow]);
  if (sig.split('|').length < 2 || !header[map.date]) return false;
  const fields = {};
  for (const k of ['date', 'no', 'name', 'staff']) fields[k] = map[k] >= 0 ? header[map[k]] || '' : '';
  const list = memLoad().filter((x) => x.sig !== sig);
  list.push({ sig, fields });
  return memSave(list.slice(-30));
}

// 워크북 → 시트별 2차원 배열. raw:true 이므로 날짜 셀은 시리얼 숫자로 옵니다.
function readWorkbook(arrayBuffer) {
  const wb = XLSX.read(arrayBuffer, { type: 'array' });
  const sheets = wb.SheetNames.map((name) => ({
    name,
    grid: XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: null, blankrows: true }),
  }));
  // 헤더 감지에 성공하는 첫 시트를 선택 (없으면 가장 점수 높은 시트)
  let pick = 0, bestScore = -1, found = false;
  sheets.forEach((s, i) => {
    s.det = memFind(s.grid) || detectHeader(s.grid);   // 기억된 설정이 자동 감지보다 우선
    if (!found && s.det.ok) { pick = i; found = true; }
    else if (!found && s.det.score > bestScore) { pick = i; bestScore = s.det.score; }
  });
  return { sheets, pick };
}

/* =====================================================================
 * [3] 행 정제 — 한 행씩 읽어 날짜 변환, 합계행/날짜없음 표시
 * ===================================================================== */
const SUM_RE = /^(합계|소계|총계|누계|계|total|subtotal|grandtotal)$|합계|소계|총계/i;
function isSumRow(cells) {
  return cells.some((c) => {
    if (typeof c !== 'string') return false;
    const t = c.replace(/\s+/g, '');
    return t.length > 0 && t.length <= 10 && SUM_RE.test(t);
  });
}
const idText = (v) => {                       // 환자번호/이름 정리 (숫자 1001.0 → "1001")
  if (v === null || v === undefined) return '';
  if (typeof v === 'number') return String(v);
  return String(v).replace(/\s+/g, ' ').trim();
};
const dateText = (v) => (v === null || v === undefined ? '' : String(v));

// grid의 headerRow 다음 줄부터 정제. 반환: 정제된 행 배열
function prepareRows(grid, headerRow, map) {
  const rows = [];
  for (let r = headerRow + 1; r < grid.length; r++) {
    const cells = grid[r] || [];
    if (cells.every((c) => c === null || c === undefined || String(c).trim() === '')) continue;   // 완전 빈 줄 무시
    const pno = map.no >= 0 ? idText(cells[map.no]) : '';
    const pname = map.name >= 0 ? idText(cells[map.name]) : '';
    const dv = map.date >= 0 ? cells[map.date] : null;
    const row = {
      excelRow: r + 1, dateRaw: dateText(dv), pno, pname,
      staff: map.staff >= 0 ? idText(cells[map.staff]) : '',
      n: null, status: 'ok',
      // 환자 식별자: 환자번호 우선 → 없으면 환자명. 접두사로 서로 섞이지 않게 함
      key: pno ? 'N:' + pno : (pname ? 'M:' + pname : null),
    };
    if (isSumRow(cells)) row.status = 'sum';
    else {
      const p = parseDateCell(dv);
      if (p.state === 'ok') row.n = p.n;
      else row.status = p.state === 'empty' ? 'noDate' : 'dateError';
    }
    rows.push(row);
  }
  return rows;
}

/* =====================================================================
 * [4] 주차 계산 엔진
 *  - 주 = 수요일~화요일. 회의일 W(수): 이번주 = W-7~W-1, 전주 = W-14~W-8
 *  - 주차 번호 = floor((날짜 − 그 해 첫 수요일)/7) + 1, 첫 수요일 이전은 Week 0 (YTD에는 포함)
 *    (첫 수요일 주를 'Week 1'로 표시합니다. 기획의 floor 식은 0부터 시작하지만,
 *     Week 0을 '첫 수요일 이전'으로 따로 쓰므로 겹치지 않게 +1 했습니다.)
 * ===================================================================== */
function analyze(rows, meetingN) {
  const thisEnd = meetingN - 1, thisStart = meetingN - 7;
  const Y = civilFromDays(thisEnd).y;                        // 기준 연도 = 이번주 마지막 날의 연도
  const jan1 = daysFromCivil(Y, 1, 1);
  const firstWed = jan1 + ((3 - dowOf(jan1) + 7) % 7);        // 그 해 첫 수요일
  const weekIdx = (n) => (n < firstWed ? 0 : Math.floor((n - firstWed) / 7) + 1);
  const kThis = weekIdx(thisEnd);
  const kMin = firstWed > jan1 ? 0 : 1;                       // 1/1이 수요일이면 Week 0 없음

  const weeks = [];
  for (let k = kMin; k <= kThis; k++) {
    const start = k === 0 ? jan1 : firstWed + 7 * (k - 1);
    const end = k === 0 ? firstWed - 1 : start + 6;
    weeks.push({ k, start, end, ids: new Set(), consults: 0 });
  }
  const c = { loaded: rows.length, sum: 0, noDate: 0, dateError: 0, prevYear: 0, after: 0, counted: 0, noPatientNo: 0, noPatientId: 0 };
  const excluded = [];
  const monthMap = new Map();      // 월(1~12) → {ids, c}  (월간 요약용: 달력 월 기준)
  const staffMap = new Map();      // 담당자 → {주차 → {ids, c}}  (직원별 통계용)
  const why = { sum: '합계/소계 행', noDate: '날짜 없음', dateError: '날짜 해석 실패', prevYear: '전년도', after: '회의일 이후' };
  const skip = (r, key) => { c[key]++; excluded.push({ excelRow: r.excelRow, reason: why[key], dateRaw: r.dateRaw, who: r.pno || r.pname }); };

  for (const r of rows) {
    if (r.status !== 'ok') { skip(r, r.status); continue; }
    if (r.n < jan1) { skip(r, 'prevYear'); continue; }
    if (r.n > thisEnd) { skip(r, 'after'); continue; }
    c.counted++;
    const w = weeks[weekIdx(r.n) - kMin];
    w.consults++;                                              // 상담건수 = 행 수
    if (r.key) w.ids.add(r.key); else c.noPatientId++;         // 주 안에서 중복 제거
    if (!r.pno) c.noPatientNo++;
    const mo = civilFromDays(r.n).m;
    if (!monthMap.has(mo)) monthMap.set(mo, { ids: new Set(), c: 0 });
    const mv = monthMap.get(mo); mv.c++; if (r.key) mv.ids.add(r.key);
    // 직원별: 같은 규칙을 담당자 단위로 적용 (주 안 중복제거, 상담건수=행 수)
    const sn = r.staff || '(미지정)';
    if (!staffMap.has(sn)) staffMap.set(sn, new Map());
    const sw = staffMap.get(sn), wk = weekIdx(r.n);
    if (!sw.has(wk)) sw.set(wk, { ids: new Set(), c: 0 });
    const sc = sw.get(wk); sc.c++; if (r.key) sc.ids.add(r.key);
  }
  // 1월 ~ 기준주 끝이 속한 달까지 (마지막 달은 기준주 끝까지만 집계 = 진행 중)
  const lastM = civilFromDays(thisEnd).m, months = [];
  let mcum = 0;
  for (let m = 1; m <= lastM; m++) {
    const v = monthMap.get(m) || { ids: new Set(), c: 0 };
    const first = daysFromCivil(Y, m, 1), nextFirst = m === 12 ? daysFromCivil(Y + 1, 1, 1) : daysFromCivil(Y, m + 1, 1);
    mcum += v.c;
    months.push({ m, patients: v.ids.size, consults: v.c, cumC: mcum, start: first, end: Math.min(nextFirst - 1, thisEnd), partial: thisEnd < nextFirst - 1 });
  }
  const staff = [...staffMap.entries()].map(([name, sw]) => {
    const g = (k) => (sw.has(k) ? sw.get(k) : { ids: new Set(), c: 0 });
    const o = { name, thisP: g(kThis).ids.size, thisC: g(kThis).c, prevP: 0, prevC: 0, ytdP: 0, ytdC: 0 };
    if (kThis - 1 >= kMin) { o.prevP = g(kThis - 1).ids.size; o.prevC = g(kThis - 1).c; }
    for (const [k, v] of sw) if (k <= kThis) { o.ytdP += v.ids.size; o.ytdC += v.c; }   // YTD 환자 = 주별 환자수의 합
    return o;
  }).sort((a, b) => b.thisC - a.thisC || b.ytdC - a.ytdC || a.name.localeCompare(b.name, 'ko'));
  let cp = 0, cc = 0;
  for (const w of weeks) { w.patients = w.ids.size; cp += w.patients; cc += w.consults; w.cumP = cp; w.cumC = cc; delete w.ids; }
  const zero = { patients: 0, consults: 0, cumP: 0, cumC: 0 };
  const cur = weeks[kThis - kMin];
  const prev = kThis - 1 >= kMin ? weeks[kThis - 1 - kMin] : zero;
  return {
    meetingN, thisStart, thisEnd, prevStart: thisStart - 7, prevEnd: thisEnd - 7, year: Y, kThis, weeks,
    this: { patients: cur.patients, consults: cur.consults },
    prev: { patients: prev.patients, consults: prev.consults },
    ytd: { patients: cur.cumP, consults: cur.cumC },
    prevYtd: { patients: prev.cumP, consults: prev.cumC },
    counts: c, excluded, staff, months,
    excludedTotal: c.sum + c.noDate + c.dateError + c.prevYear + c.after,
  };
}

/* =====================================================================
 * [5] 내장 자가진단 — 화면을 열 때마다 콘솔에서 돌아가는 간단 테스트(T1~T4, T6, T7, T9)
 * ===================================================================== */
function selfTest() {
  const out = [];
  const t = (name, ok) => out.push({ name, ok: !!ok });
  const mk = (iso, id) => ({ status: 'ok', n: parseISO(iso), key: id ? 'N:' + id : null, pno: id || '', pname: '', excelRow: 0, dateRaw: iso });
  const M = (s) => parseISO(s);
  // T1 기획 예시: 회의일 10-07 → 환자3/상담5, YTD 5/8
  const base = [mk('2026-01-05', 'A'), mk('2026-09-24', 'B'), mk('2026-09-24', 'B'),
    mk('2026-09-30', 'C'), mk('2026-10-02', 'C'), mk('2026-10-02', 'D'), mk('2026-10-05', 'E'), mk('2026-10-06', 'D')];
  const r1 = analyze(base, M('2026-10-07'));
  t('T1 이번주 환자3/상담5', r1.this.patients === 3 && r1.this.consults === 5);
  t('T1 YTD 환자5/상담8', r1.ytd.patients === 5 && r1.ytd.consults === 8);
  // T2 같은 날 같은 환자 2회 → 환자1·상담2
  const r2 = analyze([mk('2026-10-01', 'X'), mk('2026-10-01', 'X')], M('2026-10-07'));
  t('T2 환자1/상담2', r2.this.patients === 1 && r2.this.consults === 2);
  // T3 화 10-06 / 수 10-07 → 서로 다른 주
  const r3 = analyze([mk('2026-10-06', 'X'), mk('2026-10-07', 'X')], M('2026-10-14'));
  t('T3 두 주 각각 1명', r3.this.patients === 1 && r3.prev.patients === 1 && r3.ytd.patients === 2);
  // T4 01-01~01-06 → Week 0, YTD 포함
  const r4 = analyze([mk('2026-01-01', 'X'), mk('2026-01-06', 'Y')], M('2026-01-14'));
  t('T4 Week 0 포함', r4.weeks[0].k === 0 && r4.prev.patients === 2 && r4.ytd.patients === 2);
  // T6 날짜 변환 (46297 = 2026-10-02)
  t('T6 시리얼 46297', parseDateCell(46297).n === M('2026-10-02'));
  t('T6 2026.10.02', parseDateCell('2026.10.02').n === M('2026-10-02'));
  t('T6 시각 포함 텍스트', parseDateCell('2026-10-02 14:30').n === M('2026-10-02'));
  t('T6 20261002', parseDateCell(20261002).n === M('2026-10-02'));
  t('T6 오류/빈칸 구분', parseDateCell('날짜미정').state === 'error' && parseDateCell('  ').state === 'empty');
  t('날짜 하루 밀림 없음(요일)', WD[dowOf(M('2026-10-07'))] === '수');
  // T7 합계행
  t('T7 합계행 감지', isSumRow(['합계', null, 12]) && isSumRow([null, '소 계']) && !isSumRow(['2026-10-02', '김합']));
  // T9 전년도 제외
  const r9 = analyze([mk('2025-12-20', 'Z'), mk('2026-10-01', 'X')], M('2026-10-07'));
  t('T9 전년도 제외', r9.counts.prevYear === 1 && r9.this.consults === 1);
  return out;
}

/* =====================================================================
 * [6] 화면(UI)
 * ===================================================================== */
const state = { fileName: '', sheets: [], sheetIdx: 0, headerRow: 0, map: null, rows: [], meetingN: null, result: null, chart: null };

function showMsg(text) { const el = $('#msg'); el.hidden = !text; el.textContent = text || ''; }

/* --- 파일 읽기 --- */
function handleFile(file) {
  if (!file) return;
  if (!/\.(xlsx|xls)$/i.test(file.name)) { showMsg('엑셀 파일(.xlsx, .xls)만 열 수 있습니다.'); return; }
  const fr = new FileReader();
  fr.onload = () => {
    try { loadBuffer(fr.result, file.name); }
    catch (e) { console.error(e); showMsg('파일을 읽지 못했습니다. 엑셀에서 열리는 파일인지 확인해 주세요. (' + e.message + ')'); }
  };
  fr.onerror = () => showMsg('파일을 읽는 중 오류가 났습니다.');
  fr.readAsArrayBuffer(file);
}
function loadBuffer(buf, name) {
  const wb = readWorkbook(buf);
  state.fileName = name; state.sheets = wb.sheets;
  selectSheet(wb.pick);
  $('#dropzone').hidden = true; $('#fileBar').hidden = false;
  $('#fileName').textContent = name;
}
function selectSheet(i) {
  state.sheetIdx = i;
  const s = state.sheets[i];
  // 시트 선택 드롭다운 (시트가 여러 개일 때만 보임)
  $('#sheetWrap').hidden = state.sheets.length < 2;
  $('#sheetSelect').innerHTML = state.sheets.map((x, j) => '<option value="' + j + '"' + (j === i ? ' selected' : '') + '>' + esc(x.name) + '</option>').join('');
  state.headerRow = s.det.row; state.map = Object.assign({}, s.det.map);
  renderMapping();
  if (s.det.ok) { $('#mapPanel').open = false; showMsg(s.det.remembered ? '✓ 저장해 둔 컬럼 설정을 적용했습니다. (제목줄 ' + (s.det.row + 1) + '행)' : ''); recompute(); }
  else {
    $('#result').hidden = true; $('#mapPanel').open = true;
    showMsg('헤더(상담일자/환자번호/환자명 등)를 자동으로 찾지 못했습니다. 아래 "컬럼 지정"에서 직접 골라 주세요.');
  }
  $('#mapPanel').hidden = false;
}

/* --- 컬럼 수동 지정 UI --- */
function renderMapping() {
  const grid = state.sheets[state.sheetIdx].grid;
  const header = grid[state.headerRow] || [];
  let width = header.length;
  for (let r = 0; r < Math.min(grid.length, 50); r++) width = Math.max(width, (grid[r] || []).length);
  const opts = ['<option value="-1">(사용 안 함)</option>'];
  for (let i = 0; i < width; i++) opts.push('<option value="' + i + '">' + colLetter(i) + '열 · ' + esc(header[i] === null || header[i] === undefined ? '(빈 제목)' : header[i]) + '</option>');
  [['#mapDate', 'date'], ['#mapNo', 'no'], ['#mapName', 'name'], ['#mapStaff', 'staff']].forEach(([sel, k]) => {
    $(sel).innerHTML = opts.join(''); $(sel).value = String(state.map[k]);
  });
  $('#headerRowInput').value = state.headerRow + 1;
}
function applyMapping() {
  const map = { date: +$('#mapDate').value, no: +$('#mapNo').value, name: +$('#mapName').value, staff: +$('#mapStaff').value };
  if (!mappingOk(map)) { showMsg('상담일자 열과, 환자번호 또는 환자명 열 중 하나는 꼭 지정해야 합니다.'); return; }
  state.map = map; state.headerRow = Math.max(0, (+$('#headerRowInput').value || 1) - 1);
  recompute();
  showMsg(memRemember(state.sheets[state.sheetIdx].grid, state.headerRow, state.map) ? '✓ 이 엑셀 양식의 컬럼 설정을 이 PC 브라우저에 기억했습니다. 다음에 같은 양식은 자동으로 적용됩니다.' : '컬럼 설정을 적용했습니다. (이 브라우저에서는 설정을 저장할 수 없습니다)');
}

/* --- 계산 → 화면 --- */
function recompute() {
  const s = state.sheets[state.sheetIdx];
  state.rows = prepareRows(s.grid, state.headerRow, state.map);
  // [2단계 확인용] 정제된 행 목록을 콘솔(F12)에 출력
  console.log('[정제된 행] 시트=' + s.name + ', 헤더=' + (state.headerRow + 1) + '행, 총 ' + state.rows.length + '행');
  console.table(state.rows.slice(0, 300).map((r) => ({ 엑셀행: r.excelRow, 상태: r.status, 날짜: r.n === null ? '' : isoOf(r.n), 원본날짜: r.dateRaw, 환자번호: r.pno, 환자명: r.pname, 담당자: r.staff })));
  if (state.meetingN === null) setMeeting(nearestWednesday(todayN()));
  else render();
}
function todayN() { const d = new Date(); return daysFromCivil(d.getFullYear(), d.getMonth() + 1, d.getDate()); }
function setMeeting(n) {
  const snapped = nearestWednesday(n);
  $('#meetingNote').textContent = snapped !== n ? '선택한 날짜가 수요일이 아니라서 가까운 수요일(' + mdwOf(snapped) + ')로 맞췄습니다.' : '';
  state.meetingN = snapped;
  $('#meetingDate').value = isoOf(snapped);
  if (state.rows.length || state.sheets.length) render();
}

function fmtDelta(cur, prev, label) {
  const diff = cur - prev;
  if (diff === 0) return '<span class="delta same">― 변동 없음 <small>' + label + '</small></span>';
  const pct = prev === 0 ? '(전주 0)' : '(' + (diff > 0 ? '+' : '−') + Math.abs(diff / prev * 100).toFixed(1) + '%)';
  return '<span class="delta ' + (diff > 0 ? 'up' : 'down') + '">' + (diff > 0 ? '▲' : '▼') + ' ' + Math.abs(diff) + ' ' + pct + ' <small>' + label + '</small></span>';
}

function render() {
  if (state.meetingN === null || !state.sheets.length) return;
  const R = analyze(state.rows, state.meetingN);
  state.result = R;
  $('#result').hidden = false;
  // 인쇄용 머리글(화면에서는 숨김, 인쇄/PDF에서만 보임)
  const nowD = new Date();
  $('#printHead').textContent = '병원 상담 통계 주간 보고 · 회의일 ' + isoOf(R.meetingN) + '(' + WD[dowOf(R.meetingN)] + ') · 이번주 ' + mdwOf(R.thisStart) + ' ~ ' + mdwOf(R.thisEnd) +
    ' · 파일 ' + state.fileName + ' · 출력 ' + nowD.getFullYear() + '-' + pad2(nowD.getMonth() + 1) + '-' + pad2(nowD.getDate());
  const c = R.counts;

  // 경고/품질 배너
  const noMapNo = state.map.no < 0;
  const warn = $('#warnBanner');
  if (c.noPatientNo > 0 || c.noPatientId > 0) {
    let t = '⚠ ' + (noMapNo ? '환자번호 열이 없어 환자명으로 환자를 구분했습니다.' : '환자번호가 비어 있는 ' + c.noPatientNo + '행은 환자명으로 구분했습니다.') +
      ' 동명이인은 한 명으로 계산될 수 있으니 환자수는 참고용으로 봐 주세요.';
    if (c.noPatientId > 0) t += ' (번호·이름이 모두 없는 ' + c.noPatientId + '행은 상담건수에만 포함, 환자수에서는 제외)';
    warn.textContent = t; warn.hidden = false;
  } else warn.hidden = true;
  $('#qLoaded').textContent = c.loaded; $('#qExcluded').textContent = R.excludedTotal;
  $('#qNoPno').textContent = c.noPatientNo; $('#qDateErr').textContent = c.dateError;
  $('#qDetail').textContent = '집계에 반영된 행 ' + c.counted + '건 · 제외 내역: 합계/소계 ' + c.sum + ' · 날짜 없음 ' + c.noDate + ' · 날짜 오류 ' + c.dateError +
    ' · 전년도 ' + c.prevYear + ' · 회의일 이후 ' + c.after + '  (기준 연도 ' + R.year + '년)';
  $('#excludedTable').innerHTML = '<tr><th>엑셀 행</th><th>사유</th><th>원본 날짜</th><th>환자</th></tr>' +
    R.excluded.slice(0, 300).map((e) => '<tr><td>' + e.excelRow + '</td><td>' + e.reason + '</td><td>' + esc(e.dateRaw) + '</td><td>' + esc(e.who) + '</td></tr>').join('') +
    (R.excluded.length > 300 ? '<tr><td colspan="4">… 외 ' + (R.excluded.length - 300) + '행</td></tr>' : '');

  // KPI 카드 4개
  const tp = mdwOf(R.thisStart) + ' ~ ' + mdwOf(R.thisEnd);
  const card = (cls, lab, num, unit, per, delta) => '<div class="kpi ' + cls + '"><div class="lab">' + lab + '</div><div class="num">' + num + '<small>' + unit + '</small></div><div class="per">' + per + '</div>' + delta + '</div>';
  $('#kpis').innerHTML =
    card('', '주간 환자수', R.this.patients, '명', tp, fmtDelta(R.this.patients, R.prev.patients, '전주 대비')) +
    card('', '주간 상담건수', R.this.consults, '건', tp, fmtDelta(R.this.consults, R.prev.consults, '전주 대비')) +
    card('ytd', 'YTD 환자수', R.ytd.patients, '명', '01-01 ~ ' + mdOf(R.thisEnd) + ' (주별 합)', fmtDelta(R.ytd.patients, R.prevYtd.patients, '전주 누계 대비')) +
    card('ytd', 'YTD 상담건수', R.ytd.consults, '건', '01-01 ~ ' + mdOf(R.thisEnd), fmtDelta(R.ytd.consults, R.prevYtd.consults, '전주 누계 대비'));

  renderTable(R); renderChart(R); renderStaff(R); renderMonths(R);
}

function shownWeeks(R) {
  const n = +$('#rangeSelect').value;
  return n > 0 ? R.weeks.slice(-n) : R.weeks;
}
const weekLabel = (w) => 'W' + w.k;
const weekPeriod = (w) => mdOf(w.start) + ' ~ ' + mdOf(w.end);

function renderTable(R) {
  const ws = shownWeeks(R);
  $('#weekTable').innerHTML = '<thead><tr><th>주차</th><th>기간</th><th>환자수</th><th>상담건수</th><th>누적환자</th><th>누적상담</th></tr></thead><tbody>' +
    ws.map((w) => '<tr' + (w.k === R.kThis ? ' class="cur"' : '') + '><td>' + weekLabel(w) + '</td><td>' + weekPeriod(w) + '</td><td>' + w.patients + '</td><td>' + w.consults + '</td><td>' + w.cumP + '</td><td>' + w.cumC + '</td></tr>').join('') + '</tbody>';
}

/* --- 월간 요약표 (달력 월 기준) --- */
function monthTable(R) {
  const head = ['월', '집계 기간', '환자수', '상담건수', '누적상담건수', '전월 대비(상담)'];
  const body = R.months.map((x, i) => {
    const d = i === 0 ? null : x.consults - R.months[i - 1].consults;
    return [x.m + '월' + (x.partial ? ' (진행 중)' : ''), mdOf(x.start) + ' ~ ' + mdOf(x.end), x.patients, x.consults, x.cumC,
      d === null ? '' : d === 0 ? '―' : (d > 0 ? '▲ ' : '▼ ') + Math.abs(d)];
  });
  return { head, body };
}
function renderMonths(R) {
  const t = monthTable(R);
  $('#monthTable').innerHTML = '<thead><tr>' + t.head.map((h) => '<th>' + h + '</th>').join('') + '</tr></thead><tbody>' +
    t.body.map((r, i) => '<tr' + (i === t.body.length - 1 ? ' class="cur"' : '') + '>' + r.map((v) => '<td>' + esc(v) + '</td>').join('') + '</tr>').join('') + '</tbody>';
}

/* --- 직원별 통계표 --- */
function staffTable(R) {
  const head = ['담당자', '주간 환자수', '주간 상담건수', '전주 상담건수', '전주 대비', 'YTD 환자수', 'YTD 상담건수'];
  const chg = (cur, prev) => { const d = cur - prev; return d === 0 ? '―' : (d > 0 ? '▲ ' : '▼ ') + Math.abs(d); };
  const body = R.staff.map((x) => [x.name, x.thisP, x.thisC, x.prevC, chg(x.thisC, x.prevC), x.ytdP, x.ytdC]);
  const sum = (f) => R.staff.reduce((a, x) => a + x[f], 0);
  body.push(['합계(직원별 합)', sum('thisP'), sum('thisC'), sum('prevC'), chg(sum('thisC'), sum('prevC')), sum('ytdP'), sum('ytdC')]);
  return { head, body };
}
function renderStaff(R) {
  const show = state.map.staff >= 0 && R.staff.length > 0;
  state.hasStaff = show;
  $('#staffCard').hidden = !show;
  if (!show) return;
  const t = staffTable(R);
  $('#staffTable').innerHTML = '<thead><tr>' + t.head.map((h) => '<th>' + h + '</th>').join('') + '</tr></thead><tbody>' +
    t.body.map((r, i) => '<tr' + (i === t.body.length - 1 ? ' class="cur"' : '') + '>' + r.map((v) => '<td>' + esc(v) + '</td>').join('') + '</tr>').join('') + '</tbody>';
}

function renderChart(R) {
  const ws = shownWeeks(R);
  if (state.chart) state.chart.destroy();
  if (typeof Chart === 'undefined') return;
  Chart.defaults.font.family = "'Malgun Gothic','Apple SD Gothic Neo','Noto Sans KR',sans-serif";
  state.chart = new Chart($('#chart'), {
    data: {
      labels: ws.map(weekLabel),
      datasets: [
        { type: 'line', label: '환자수', data: ws.map((w) => w.patients), borderColor: '#17375e', backgroundColor: '#17375e', borderWidth: 2.5, pointRadius: 3.5, tension: 0.15, order: 1 },
        { type: 'bar', label: '상담건수', data: ws.map((w) => w.consults), backgroundColor: ws.map((w) => (w.k === R.kThis ? '#1f6fb5' : '#a9c6e3')), order: 2 },
      ],
    },
    options: {
      responsive: true, maintainAspectRatio: false, animation: false,
      scales: { y: { beginAtZero: true, ticks: { precision: 0 }, grid: { color: '#e6ebf1' } }, x: { grid: { display: false } } },
      plugins: { tooltip: { callbacks: { title: (it) => { const w = ws[it[0].dataIndex]; return weekLabel(w) + ' (' + weekPeriod(w) + ')'; } } } },
    },
  });
  $('#chartBox').setAttribute('aria-label', '주간 추이: 이번주 환자 ' + R.this.patients + '명, 상담 ' + R.this.consults + '건');
}

/* --- 표 복사: HTML 표 + 텍스트(탭 구분)를 함께 복사 → 한글/PPT/엑셀에 '선 있는 표'로 붙음 --- */
function buildCopy(kind) {
  const R = state.result, ws = shownWeeks(R);
  let head, body, hl = (i) => ws[i].k === R.kThis;       // hl: 강조할 줄
  if (kind === 'staff') { const t = staffTable(R); head = t.head; body = t.body; hl = (i) => i === body.length - 1; }
  else if (kind === 'month') { const t = monthTable(R); head = t.head; body = t.body; hl = (i) => i === body.length - 1; }
  else { head = ['주차', '기간', '환자수', '상담건수', '누적환자', '누적상담']; body = ws.map((w) => [weekLabel(w), weekPeriod(w), w.patients, w.consults, w.cumP, w.cumC]); }
  const cell = 'border:1px solid #000000;padding:4px 8px;text-align:center;';
  let html = '<table border="1" cellspacing="0" cellpadding="4" style="border-collapse:collapse;border:1px solid #000000;font-family:\'맑은 고딕\',sans-serif;font-size:11pt;">';
  html += '<tr>' + head.map((h) => '<th style="' + cell + 'background-color:#dfe6ee;font-weight:bold;">' + h + '</th>').join('') + '</tr>';
  body.forEach((row, i) => { html += '<tr>' + row.map((v) => '<td style="' + cell + (hl(i) ? 'background-color:#e4effa;font-weight:bold;' : '') + '">' + esc(v) + '</td>').join('') + '</tr>'; });
  html += '</table>';
  const text = [head].concat(body).map((r) => r.join('\t')).join('\n');
  return { html, text };
}
async function copyTable(kind) {
  if (!state.result) return;
  const { html, text } = buildCopy(kind);
  const done = (msg) => { const el = $({ staff: '#copyStatusStaff', month: '#copyStatusMonth' }[kind] || '#copyStatus'); el.textContent = msg; setTimeout(() => { el.textContent = ''; }, 3000); };
  try {   // 1순위: 최신 클립보드 API
    if (!(navigator.clipboard && window.ClipboardItem)) throw new Error('no clipboard api');
    await navigator.clipboard.write([new ClipboardItem({ 'text/html': new Blob([html], { type: 'text/html' }), 'text/plain': new Blob([text], { type: 'text/plain' }) })]);
    done('✓ 복사됨'); return;
  } catch (e) { /* file:// 에서 막히면 아래 폴백 */ }
  try {   // 2순위: execCommand('copy') — copy 이벤트에서 HTML/텍스트를 직접 지정
    const holder = document.createElement('div');
    holder.style.cssText = 'position:fixed;left:-9999px;top:0;'; holder.innerHTML = html; document.body.appendChild(holder);
    const range = document.createRange(); range.selectNodeContents(holder);
    const sel = window.getSelection(); sel.removeAllRanges(); sel.addRange(range);
    const onCopy = (ev) => { ev.clipboardData.setData('text/html', html); ev.clipboardData.setData('text/plain', text); ev.preventDefault(); };
    document.addEventListener('copy', onCopy, { once: true });
    const ok = document.execCommand('copy');
    document.removeEventListener('copy', onCopy); sel.removeAllRanges(); holder.remove();
    done(ok ? '✓ 복사됨' : '복사에 실패했습니다. 표를 마우스로 선택해 Ctrl+C 해 주세요.');
  } catch (e) { done('복사에 실패했습니다. 표를 마우스로 선택해 Ctrl+C 해 주세요.'); }
}

/* --- 결과 엑셀 다운로드: 요약 / 주간표 / 직원별 / 월간 / 제외된 행 시트 (이 PC 안에서 파일 생성) --- */
function buildResultSheets(R) {
  const c = R.counts, rng = (a, b) => isoOf(a) + ' ~ ' + isoOf(b);
  const sheets = [];
  sheets.push(['요약', [
    ['병원 상담 통계 주간 요약'], ['회의일', isoOf(R.meetingN)], ['이번주', rng(R.thisStart, R.thisEnd)], ['전주', rng(R.prevStart, R.prevEnd)],
    ['파일', state.fileName], [],
    ['지표', '이번주', '전주', '전주 대비 증감', 'YTD (1/1~이번주 끝)', '전주 말 YTD'],
    ['환자수', R.this.patients, R.prev.patients, R.this.patients - R.prev.patients, R.ytd.patients, R.prevYtd.patients],
    ['상담건수', R.this.consults, R.prev.consults, R.this.consults - R.prev.consults, R.ytd.consults, R.prevYtd.consults], [],
    ['데이터 품질'], ['로드 행수', c.loaded], ['집계 반영 행수', c.counted], ['제외 행수 합계', R.excludedTotal],
    ['  합계/소계 행', c.sum], ['  날짜 없음', c.noDate], ['  날짜 오류', c.dateError], ['  전년도', c.prevYear], ['  회의일 이후', c.after],
    ['환자번호 없는 행수', c.noPatientNo], [],
    ['YTD 환자수는 주별 환자수의 합입니다(주 간 중복제거 없음).'],
  ]]);
  sheets.push(['주간표', [['주차', '시작일', '종료일', '환자수', '상담건수', '누적환자', '누적상담']].concat(
    R.weeks.map((w) => ['W' + w.k, isoOf(w.start), isoOf(w.end), w.patients, w.consults, w.cumP, w.cumC]))]);
  if (R.staff.length && state.hasStaff) { const t = staffTable(R); sheets.push(['직원별', [t.head].concat(t.body)]); }
  const mt = monthTable(R); sheets.push(['월간', [mt.head].concat(mt.body)]);
  sheets.push(['제외된 행', [['엑셀 행', '사유', '원본 날짜', '환자']].concat(R.excluded.map((e) => [e.excelRow, e.reason, e.dateRaw, e.who]))]);
  return sheets;
}
function downloadResult() {
  if (!state.result) return;
  const wb = XLSX.utils.book_new();
  buildResultSheets(state.result).forEach(([name, aoa]) => XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), name));
  const data = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
  a.download = '상담통계_회의일' + isoOf(state.result.meetingN) + '.xlsx';
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 2000);
}

/* --- 이벤트 연결 --- */
function bind() {
  const dz = $('#dropzone'), fi = $('#fileInput');
  dz.addEventListener('click', () => fi.click());
  dz.addEventListener('keydown', (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); fi.click(); } });
  fi.addEventListener('change', () => { handleFile(fi.files[0]); fi.value = ''; });
  $('#otherFileBtn').addEventListener('click', () => fi.click());
  ['dragenter', 'dragover'].forEach((t) => window.addEventListener(t, (e) => { e.preventDefault(); dz.classList.add('over'); }));
  ['dragleave', 'drop'].forEach((t) => window.addEventListener(t, (e) => { e.preventDefault(); dz.classList.remove('over'); }));
  window.addEventListener('drop', (e) => { if (e.dataTransfer && e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]); });
  $('#sheetSelect').addEventListener('change', (e) => selectSheet(+e.target.value));
  $('#applyMapBtn').addEventListener('click', applyMapping);
  $('#clearMemBtn').addEventListener('click', () => { memSave([]); showMsg('저장해 둔 컬럼 설정을 모두 지웠습니다.'); });
  $('#headerRowInput').addEventListener('change', () => {          // 헤더 행 바꾸면 그 행으로 열 이름 다시 인식
    state.headerRow = Math.max(0, (+$('#headerRowInput').value || 1) - 1);
    state.map = mappingForRow(state.sheets[state.sheetIdx].grid[state.headerRow]); renderMapping();
  });
  $('#meetingDate').addEventListener('change', (e) => { const n = parseISO(e.target.value); if (n !== null) setMeeting(n); });
  $('#prevWeekBtn').addEventListener('click', () => setMeeting(state.meetingN === null ? nearestWednesday(todayN()) - 7 : state.meetingN - 7));
  $('#nextWeekBtn').addEventListener('click', () => setMeeting(state.meetingN === null ? nearestWednesday(todayN()) + 7 : state.meetingN + 7));
  $('#rangeSelect').addEventListener('change', () => { if (state.result) { renderTable(state.result); renderChart(state.result); } });
  $('#downloadBtn').addEventListener('click', downloadResult);
  $('#printBtn').addEventListener('click', () => window.print());
  $('#copyBtn').addEventListener('click', () => copyTable('week'));
  $('#copyMonthBtn').addEventListener('click', () => copyTable('month'));
  $('#copyStaffBtn').addEventListener('click', () => copyTable('staff'));
}

/* =====================================================================
 * [7] 검증용 공개 함수 — verify.py가 헤드리스 브라우저에서 호출합니다(화면과 같은 계산 코드를 사용).
 * ===================================================================== */
function analyzeWorkbook(arrayBuffer, meetingISO) {
  const wb = readWorkbook(arrayBuffer);
  const s = wb.sheets[wb.pick];
  const rows = prepareRows(s.grid, s.det.row, s.det.map);
  const R = analyze(rows, parseISO(meetingISO));
  const weeks = {};
  R.weeks.forEach((w) => { if (w.patients || w.consults) weeks[w.k] = [w.patients, w.consults]; });
  const c = R.counts;
  return {
    detected: s.det.ok, sheet: s.name, headerRow: s.det.row + 1, map: s.det.map,
    this: R.this, prev: R.prev, ytd: R.ytd, weeks_nonzero: weeks,
    counts: { loaded: c.loaded, sum: c.sum, no_date: c.noDate, date_error: c.dateError, prev_year: c.prevYear, after: c.after, counted: c.counted, no_patient_no: c.noPatientNo },
    staff: Object.fromEntries(R.staff.map((x) => [x.name, { this: [x.thisP, x.thisC], prev: [x.prevP, x.prevC], ytd: [x.ytdP, x.ytdC] }])),
    months: Object.fromEntries(R.months.filter((x) => x.patients || x.consults).map((x) => [x.m, [x.patients, x.consults]])),
    thisRange: [isoOf(R.thisStart), isoOf(R.thisEnd)],
  };
}
window.DashboardAPI = { analyzeWorkbook, parseDateCell, analyze, selfTest };

/* --- 시작 --- */
bind();
const results = selfTest();
const passed = results.filter((r) => r.ok).length;
results.forEach((r) => { if (!r.ok) console.error('[자가진단 실패] ' + r.name); });
console.log('[자가진단] ' + passed + '/' + results.length + ' 통과');
$('#foot').textContent = '내장 자가진단 ' + passed + '/' + results.length + ' 통과' + (passed === results.length ? ' ✓' : ' — 일부 실패(콘솔 F12 확인)');
setMeeting(nearestWednesday(todayN()));     // 회의일 기본값 = 오늘과 가장 가까운 수요일
})();

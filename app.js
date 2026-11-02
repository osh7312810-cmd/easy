/* 법인카드 집근처 사용 분석 - 브라우저 단독 동작 (데이터는 서버로 전송되지 않음) */
'use strict';

// ---------- 상수 ----------
const FIELDS = {
  seq: '순번', card: '카드번호', user: '이용자명', date: '이용일', merchant: '가맹점명',
  amount: '이용금액', addr: '가맹점주소', bizNo: '사업자등록번호', approvalNo: '승인번호',
  taxType: '과세유형', slipType: '전표구분', usd: '미화금액', fx: '미화환율', time: '승인시간',
  mccCode: '가맹점업종코드', mccName: '가맹점업종명', fuel: '유종', unitPrice: '리터당단가',
};
const REQUIRED = ['user', 'date', 'merchant', 'amount', 'addr'];
const DEFAULT_HOLIDAYS = [
  // 2025
  '2025-01-01','2025-01-27','2025-01-28','2025-01-29','2025-01-30','2025-03-01','2025-03-03','2025-05-05','2025-05-06',
  '2025-06-03','2025-06-06','2025-08-15','2025-10-03','2025-10-05','2025-10-06','2025-10-07','2025-10-08','2025-10-09','2025-12-25',
  // 2026
  '2026-01-01','2026-02-16','2026-02-17','2026-02-18','2026-03-01','2026-03-02','2026-05-05','2026-05-24','2026-05-25',
  '2026-06-03','2026-06-06','2026-08-15','2026-08-17','2026-09-24','2026-09-25','2026-09-26','2026-10-03','2026-10-05',
  '2026-10-09','2026-12-25',
].join('\n');
// 주소가 PG사 본사로 찍히는 온라인 결제 판별용
const PG_PATTERNS = [/ - /, /이니시스/, /KCP/i, /나이스페이/, /토스페이먼츠/, /갤럭시아/, /다날/, /KG모빌리언스/, /NHN/i, /네이버페이/, /카카오페이/, /페이코/i];
const PG_MCC = ['PG쇼핑몰', '전자상거래', '통신판매'];
const DOW = ['일', '월', '화', '수', '목', '금', '토'];
const JUDGE_CLASS = { 고위험: 'b-high', 주의: 'b-warn', 정상: 'b-ok', 제외: 'b-mute', 확인불가: 'b-mute' };
const JUDGE_COLOR = { 고위험: '#d92d20', 주의: '#e8890c', 정상: '#12a150' };

// ---------- 저장소 ----------
const store = {
  get(k, d) { try { const v = localStorage.getItem(k); return v == null ? d : JSON.parse(v); } catch { return d; } },
  set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) { console.warn('저장 실패', e); } },
};
let settings = Object.assign({
  kakaoKey: '', radius: 1000, nightStart: 22, nightEnd: 6, exOnline: true, exForeign: true, holidays: DEFAULT_HOLIDAYS,
  geminiKey: '', geminiModel: 'gemini-3.5-flash', aiMask: true,
}, store.get('settings', {}));
// gemini-2.5 계열은 신규 사용자에게 종료됨
if (/^gemini-2\.5-(flash|pro)$/.test(settings.geminiModel)) settings.geminiModel = 'gemini-3.5-flash';
let employees = store.get('employees', {});   // name -> {address, radius, lat, lng, matched, status}
let geoCache = store.get('geoCache', {});     // key -> {lat, lng, method, matched} | {fail:true}
let txns = [];       // 정규화된 거래
let results = [];    // 분석 결과 (txns + 판정)

// ---------- 유틸 ----------
const $ = (s) => document.querySelector(s);
const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const won = (n) => (Number(n) || 0).toLocaleString('ko-KR');
const normHeader = (h) => String(h ?? '').replace(/\s+/g, '');
const cleanAddr = (a) => String(a ?? '').replace(/\s+/g, ' ').trim();
const pad = (n) => String(n).padStart(2, '0');
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const fmtDist = (m) => (m == null ? '-' : m < 1000 ? `${Math.round(m)}m` : `${(m / 1000).toFixed(1)}km`);

function toNumber(v) {
  if (typeof v === 'number') return v;
  const n = parseFloat(String(v ?? '').replace(/[,\s원]/g, ''));
  return isNaN(n) ? 0 : n;
}
function parseDate(v) {
  if (v instanceof Date) return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  if (typeof v === 'number' && v > 20000) {
    const d = XLSX.SSF.parse_date_code(v);
    return new Date(d.y, d.m - 1, d.d);
  }
  const m = String(v ?? '').match(/(\d{4})\D?(\d{1,2})\D?(\d{1,2})/);
  return m ? new Date(+m[1], +m[2] - 1, +m[3]) : null;
}
function parseTime(v) {
  if (typeof v === 'number' && v < 1) {
    const s = Math.round(v * 86400);
    return `${pad(Math.floor(s / 3600))}:${pad(Math.floor((s % 3600) / 60))}:${pad(s % 60)}`;
  }
  const str = String(v ?? '').trim();
  const m = str.match(/^(\d{1,2}):?(\d{2}):?(\d{2})?$/);
  return m ? `${pad(m[1])}:${m[2]}:${m[3] || '00'}` : str;
}
const ymd = (d) => (d ? `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}` : '');
function haversine(lat1, lng1, lat2, lng2) {
  const R = 6371000, rad = Math.PI / 180;
  const dLat = (lat2 - lat1) * rad, dLng = (lng2 - lng1) * rad;
  const a = Math.sin(dLat / 2) ** 2 + Math.cos(lat1 * rad) * Math.cos(lat2 * rad) * Math.sin(dLng / 2) ** 2;
  return 2 * R * Math.asin(Math.sqrt(a));
}

// ---------- 탭 ----------
document.querySelectorAll('nav button').forEach((b) => b.addEventListener('click', () => {
  document.querySelectorAll('nav button').forEach((x) => x.classList.toggle('active', x === b));
  document.querySelectorAll('.tab').forEach((t) => t.classList.toggle('active', t.id === 'tab-' + b.dataset.tab));
  if (b.dataset.tab === 'dash' && map) setTimeout(() => { map.relayout(); drawMap(); }, 50);
}));

// ---------- 카카오 SDK ----------
let kakaoReady = null, geocoder = null, places = null, map = null, mapObjs = [];
function loadKakao() {
  if (!settings.kakaoKey) return Promise.reject(new Error('카카오 JavaScript 키가 없습니다. 설정 탭에서 입력하세요.'));
  if (kakaoReady) return kakaoReady;
  if (location.protocol === 'file:') {
    $('#sdkState').textContent = '카카오 지도: file:// 에서는 동작하지 않음 (실행.bat 사용)';
    return Promise.reject(new Error('file:// 로 열려 있습니다. 실행.bat 으로 실행하세요.'));
  }
  $('#sdkState').textContent = '카카오 지도: 연결 중…';
  kakaoReady = new Promise((resolve, reject) => {
    const s = document.createElement('script');
    s.src = `https://dapi.kakao.com/v2/maps/sdk.js?appkey=${encodeURIComponent(settings.kakaoKey)}&libraries=services&autoload=false`;
    s.onload = () => kakao.maps.load(() => {
      geocoder = new kakao.maps.services.Geocoder();
      places = new kakao.maps.services.Places();
      $('#sdkState').textContent = '카카오 지도: 연결됨';
      initMap();
      resolve();
    });
    s.onerror = () => {
      kakaoReady = null;
      $('#sdkState').textContent = '카카오 지도: 연결 실패';
      reject(new Error('카카오 SDK 로드 실패 — JavaScript 키와 Web 도메인 등록(' + location.origin + '), 카카오맵 활성화를 확인하세요.'));
    };
    document.head.appendChild(s);
  });
  return kakaoReady;
}
function initMap() {
  const el = $('#map');
  el.textContent = '';
  map = new kakao.maps.Map(el, { center: new kakao.maps.LatLng(37.5665, 126.978), level: 7 });
  map.addControl(new kakao.maps.ZoomControl(), kakao.maps.ControlPosition.RIGHT);
}

const addressSearch = (q) => new Promise((res) => geocoder.addressSearch(q, (r, st) => res(st === kakao.maps.services.Status.OK ? r : [])));
const keywordSearch = (q) => new Promise((res) => places.keywordSearch(q, (r, st) => res(st === kakao.maps.services.Status.OK ? r : []), { size: 15 }));

/**
 * 주소 → 좌표. 카드사 주소는 띄어쓰기/상세주소가 지저분해서 단계적으로 시도한다.
 *  1) 정리한 전체 주소  2) 가맹점명 키워드 검색(같은 시군구 내 결과만)  3) 뒤에서부터 토큰을 줄여가며 재시도(근사)
 */
async function geocode(addr, merchant) {
  const a = cleanAddr(addr);
  const key = a + '||' + (merchant || '');
  if (geoCache[key]) return geoCache[key];
  let out = { fail: true };
  const tokens = a.split(' ');

  let r = a ? await addressSearch(a) : [];
  if (r.length) out = { lat: +r[0].y, lng: +r[0].x, method: '주소', matched: r[0].address_name };

  if (out.fail && merchant) {
    const name = merchant.replace(/\(주\)|주식회사|㈜/g, '').trim();
    const area = tokens.slice(0, 2).join(' ');
    const list = await keywordSearch(name);
    const hit = list.find((p) => tokens.length < 2 || (p.road_address_name || p.address_name || '').includes(tokens[1]));
    if (hit) out = { lat: +hit.y, lng: +hit.x, method: '가맹점명', matched: `${hit.place_name} (${hit.road_address_name || hit.address_name})` };
    else if (area) {
      const list2 = await keywordSearch(area + ' ' + name);
      if (list2.length) out = { lat: +list2[0].y, lng: +list2[0].x, method: '가맹점명', matched: `${list2[0].place_name} (${list2[0].road_address_name || list2[0].address_name})` };
    }
  }

  for (let n = tokens.length - 1; out.fail && n >= 3; n--) {
    r = await addressSearch(tokens.slice(0, n).join(' '));
    if (r.length) out = { lat: +r[0].y, lng: +r[0].x, method: '근사', matched: r[0].address_name };
  }

  geoCache[key] = out;
  store.set('geoCache', geoCache);
  return out;
}

// ---------- 파일 업로드 ----------
const drop = $('#drop'), fileInput = $('#file');
drop.addEventListener('click', () => fileInput.click());
drop.addEventListener('dragover', (e) => { e.preventDefault(); drop.classList.add('over'); });
drop.addEventListener('dragleave', () => drop.classList.remove('over'));
drop.addEventListener('drop', (e) => { e.preventDefault(); drop.classList.remove('over'); loadFiles(e.dataTransfer.files); });
fileInput.addEventListener('change', () => { loadFiles(fileInput.files); fileInput.value = ''; });

async function readSheetRows(file) {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: 'array', cellDates: true, codepage: 949 });
  const rows = [];
  for (const name of wb.SheetNames) rows.push(...XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: '' }).map((r) => r));
  return rows;
}

async function loadFiles(files) {
  if (!files || !files.length) return;
  const all = [], errors = [];
  for (const f of files) {
    try {
      const parsed = parseRows(await readSheetRows(f));
      all.push(...parsed);
    } catch (e) { errors.push(`${f.name}: ${e.message}`); }
  }
  if (errors.length) alert(errors.join('\n'));
  if (!all.length) return;
  setTransactions(all, [...files].map((f) => f.name).join(', '));
}

function parseRows(rows) {
  const hIdx = rows.findIndex((r) => r.some((c) => normHeader(c) === '이용자명') && r.some((c) => normHeader(c) === '가맹점명'));
  if (hIdx < 0) throw new Error('헤더 행(이용자명, 가맹점명 등)을 찾지 못했습니다.');
  const header = rows[hIdx].map(normHeader);
  const col = {};
  for (const [k, label] of Object.entries(FIELDS)) col[k] = header.indexOf(label);
  const missing = REQUIRED.filter((k) => col[k] < 0).map((k) => FIELDS[k]);
  if (missing.length) throw new Error('필수 컬럼 없음: ' + missing.join(', '));

  const out = [];
  for (const r of rows.slice(hIdx + 1)) {
    const get = (k) => (col[k] >= 0 ? r[col[k]] : '');
    if (!String(get('user')).trim() && !String(get('merchant')).trim()) continue;
    const date = parseDate(get('date'));
    out.push({
      seq: get('seq'), card: String(get('card')), user: String(get('user')).trim(), date, dateStr: ymd(date),
      merchant: String(get('merchant')).trim(), amount: toNumber(get('amount')), addr: cleanAddr(get('addr')),
      bizNo: String(get('bizNo')), approvalNo: String(get('approvalNo')), taxType: String(get('taxType')),
      slipType: String(get('slipType')), usd: toNumber(get('usd')), fx: toNumber(get('fx')), time: parseTime(get('time')),
      mccCode: String(get('mccCode')), mccName: String(get('mccName')).trim(), fuel: String(get('fuel')), unitPrice: toNumber(get('unitPrice')),
    });
  }
  return out;
}

function setTransactions(list, label) {
  txns = list;
  results = [];
  const users = [...new Set(txns.map((t) => t.user))].filter(Boolean);
  let added = 0;
  for (const u of users) if (!employees[u]) { employees[u] = { address: '', radius: '', status: '미입력' }; added++; }
  store.set('employees', employees);
  renderEmployees();
  const noHome = users.filter((u) => !employees[u].lat);
  $('#fileInfo').innerHTML = `<b>${esc(label)}</b> · ${txns.length.toLocaleString()}건 · 이용자 ${users.length}명`
    + (noHome.length ? ` · <span style="color:var(--warn)">집주소 미등록 ${noHome.length}명 (직원 집주소 탭)</span>` : '')
    + (added ? ` · 신규 이용자 ${added}명 추가됨` : '');
  $('#btnAnalyze').disabled = false;
  fillUserSelects(users);
  renderAll();
}

function fillUserSelects(users) {
  for (const id of ['#fUser', '#mapUser', '#aiUser']) {
    const sel = $(id), cur = sel.value;
    const first = id === '#mapUser' ? '<option value="">이용자 선택</option>' : '<option value="">전체 이용자</option>';
    sel.innerHTML = first + users.map((u) => `<option>${esc(u)}</option>`).join('');
    if (users.includes(cur)) sel.value = cur;
    else if (id === '#mapUser' && users.length) sel.value = users[0];
  }
}

// ---------- 분석 ----------
function holidaySet() { return new Set(String(settings.holidays).split(/[\s,]+/).filter(Boolean)); }
function isOnline(t) {
  return PG_MCC.some((m) => t.mccName.includes(m)) || PG_PATTERNS.some((p) => p.test(t.merchant));
}
function isForeign(t) { return t.usd > 0 || (t.addr && !/[가-힣]/.test(t.addr)); }
function inNight(time) {
  const h = parseInt(String(time).slice(0, 2), 10);
  if (isNaN(h)) return false;
  const { nightStart: s, nightEnd: e } = settings;
  return s > e ? h >= s || h < e : h >= s && h < e;
}

$('#btnAnalyze').addEventListener('click', runAnalysis);
async function runAnalysis() {
  try { await loadKakao(); } catch (e) { alert(e.message); return; }
  const btn = $('#btnAnalyze');
  btn.disabled = true;

  // 1) 집주소 좌표 (아직 없는 직원)
  const users = [...new Set(txns.map((t) => t.user))];
  for (const u of users) if (employees[u]?.address && employees[u].lat == null) await geocodeEmployee(u, false);
  store.set('employees', employees);
  renderEmployees();

  // 2) 가맹점 좌표
  const targets = txns.filter((t) => !(settings.exOnline && isOnline(t)) && !(settings.exForeign && isForeign(t)) && t.addr);
  const uniq = [...new Map(targets.map((t) => [cleanAddr(t.addr) + '||' + t.merchant, t])).values()];
  let done = 0;
  for (const t of uniq) {
    const cached = geoCache[cleanAddr(t.addr) + '||' + t.merchant];
    if (!cached) { await geocode(t.addr, t.merchant); await sleep(60); }
    done++;
    $('#prog').style.width = (done / uniq.length * 100) + '%';
    $('#progText').textContent = `가맹점 좌표 변환 ${done}/${uniq.length}`;
  }
  $('#progText').textContent = `완료 · 가맹점 ${uniq.length}곳`;
  btn.disabled = false;
  computeResults();
}

function computeResults() {
  const hol = holidaySet();
  results = txns.map((t) => {
    const r = { ...t, flags: [], dist: null, geo: null, judge: '정상', reason: '' };
    const dow = t.date ? t.date.getDay() : null;
    r.dow = dow == null ? '' : DOW[dow];
    const offDay = dow === 0 || dow === 6 || hol.has(t.dateStr);
    const night = inNight(t.time);
    if (offDay) r.flags.push(hol.has(t.dateStr) ? '공휴일' : '주말');
    if (night) r.flags.push('심야');
    if (/취소/.test(t.slipType)) r.flags.push('취소');

    if (settings.exOnline && isOnline(t)) { r.judge = '제외'; r.reason = '온라인/PG'; return r; }
    if (settings.exForeign && isForeign(t)) { r.judge = '제외'; r.reason = '해외'; return r; }
    const emp = employees[t.user];
    if (!emp || emp.lat == null) { r.judge = '확인불가'; r.reason = '집주소 미등록'; return r; }
    const g = geoCache[cleanAddr(t.addr) + '||' + t.merchant];
    if (!g || g.fail) { r.judge = '확인불가'; r.reason = '가맹점 위치 확인 실패'; return r; }

    r.geo = g;
    r.dist = haversine(emp.lat, emp.lng, g.lat, g.lng);
    const R = Number(emp.radius) || Number(settings.radius) || 1000;
    const near = r.dist <= R, adj = r.dist <= R * 2;
    const risky = offDay || night;
    if (near && risky) { r.judge = '고위험'; r.reason = `집 ${fmtDist(r.dist)} + ${r.flags.filter((f) => f !== '취소').join('·')}`; }
    else if (near) { r.judge = '주의'; r.reason = `집 ${fmtDist(r.dist)} 이내`; }
    else if (adj && risky) { r.judge = '주의'; r.reason = `집 인접 ${fmtDist(r.dist)} + ${r.flags.filter((f) => f !== '취소').join('·')}`; }
    if (g.method === '근사' && r.judge !== '정상') r.reason += ' (근사좌표)';
    return r;
  });
  $('#btnExport').disabled = !results.length;
  renderAll();
}

// ---------- 렌더링 ----------
function renderAll() { renderKpis(); renderUsers(); renderTx(); drawMap(); }

function renderKpis() {
  const src = results.length ? results : [];
  if (!src.length) {
    $('#kpis').innerHTML = txns.length ? `<div class="kpi"><div class="l">업로드 건수</div><div class="v">${txns.length.toLocaleString()}</div><div class="s">분석 실행 전</div></div>` : '';
    return;
  }
  const sum = (a) => a.reduce((s, x) => s + x.amount, 0);
  const valid = src.filter((r) => ['고위험', '주의', '정상'].includes(r.judge));
  const near = valid.filter((r) => r.judge !== '정상');
  const high = valid.filter((r) => r.judge === '고위험');
  const ex = src.filter((r) => r.judge === '제외');
  const unk = src.filter((r) => r.judge === '확인불가');
  const pct = valid.length ? Math.round(near.length / valid.length * 1000) / 10 : 0;
  $('#kpis').innerHTML = `
    <div class="kpi"><div class="l">분석 대상</div><div class="v">${valid.length.toLocaleString()}건</div><div class="s">${won(sum(valid))}원</div></div>
    <div class="kpi high"><div class="l">고위험 (집근처+휴일/심야)</div><div class="v">${high.length.toLocaleString()}건</div><div class="s">${won(sum(high))}원</div></div>
    <div class="kpi warn"><div class="l">집근처·인접 의심 (고위험+주의)</div><div class="v">${near.length.toLocaleString()}건</div><div class="s">${won(sum(near))}원 · 건수 비율 ${pct}%</div></div>
    <div class="kpi"><div class="l">제외 (온라인/해외)</div><div class="v">${ex.length.toLocaleString()}건</div><div class="s">${won(sum(ex))}원</div></div>
    <div class="kpi"><div class="l">확인불가</div><div class="v">${unk.length.toLocaleString()}건</div><div class="s">집주소 미등록 / 위치 실패</div></div>`;
}

function renderUsers() {
  const tb = $('#userTbl tbody');
  if (!results.length) { tb.innerHTML = '<tr><td colspan="7" class="muted">분석 결과가 없습니다</td></tr>'; return; }
  const by = {};
  for (const r of results) {
    const u = (by[r.user] ||= { n: 0, amt: 0, nn: 0, namt: 0, hi: 0, home: employees[r.user]?.lat != null });
    if (!['고위험', '주의', '정상'].includes(r.judge)) continue;
    u.n++; u.amt += r.amount;
    if (r.judge !== '정상') { u.nn++; u.namt += r.amount; }
    if (r.judge === '고위험') u.hi++;
  }
  const rows = Object.entries(by).sort((a, b) => b[1].namt - a[1].namt);
  tb.innerHTML = rows.map(([name, u]) => `<tr data-user="${esc(name)}">
    <td>${esc(name)}${u.home ? '' : ' <span class="badge b-mute">집주소 없음</span>'}</td>
    <td class="num">${u.n}</td><td class="num">${won(u.amt)}</td>
    <td class="num">${u.nn}</td><td class="num">${won(u.namt)}</td>
    <td class="num">${u.n ? (u.nn / u.n * 100).toFixed(1) + '%' : '-'}</td>
    <td class="num">${u.hi ? `<span class="badge b-high">${u.hi}</span>` : 0}</td></tr>`).join('');
  tb.querySelectorAll('tr').forEach((tr) => tr.addEventListener('click', () => {
    $('#fUser').value = tr.dataset.user; $('#mapUser').value = tr.dataset.user; renderTx(); drawMap();
  }));
}

function filteredResults() {
  const u = $('#fUser').value, j = $('#fJudge').value, q = $('#fText').value.trim().toLowerCase();
  return results.filter((r) => (!u || r.user === u) && (!j || r.judge === j)
    && (!q || `${r.merchant} ${r.addr} ${r.mccName}`.toLowerCase().includes(q)));
}
const JUDGE_ORDER = { 고위험: 0, 주의: 1, 확인불가: 2, 정상: 3, 제외: 4 };
function renderTx() {
  const tb = $('#txTbl tbody');
  if (!results.length) {
    tb.innerHTML = `<tr><td colspan="11" class="muted">${txns.length ? '“집근처 분석 실행”을 눌러주세요' : '분석 결과가 없습니다'}</td></tr>`;
    $('#txCount').textContent = '';
    return;
  }
  const list = filteredResults().sort((a, b) => JUDGE_ORDER[a.judge] - JUDGE_ORDER[b.judge] || (a.dist ?? 1e12) - (b.dist ?? 1e12));
  tb.innerHTML = list.map((r) => `<tr data-i="${results.indexOf(r)}">
    <td><span class="badge ${JUDGE_CLASS[r.judge]}">${r.judge}</span></td>
    <td>${esc(r.dateStr)} (${r.dow})</td><td>${esc(r.time)}</td><td>${esc(r.user)}</td>
    <td>${esc(r.merchant)}</td><td>${esc(r.mccName)}</td><td class="num">${won(r.amount)}</td>
    <td class="num">${fmtDist(r.dist)}</td>
    <td>${r.flags.map((f) => `<span class="flag">${f}</span>`).join('')}${esc(r.reason)}</td>
    <td title="${esc(r.geo?.matched || '')}">${esc(r.addr)}</td><td>${esc(r.geo?.method || '-')}</td></tr>`).join('');
  $('#txCount').textContent = `${list.length.toLocaleString()}건 · ${won(list.reduce((s, r) => s + r.amount, 0))}원`;
  tb.querySelectorAll('tr').forEach((tr) => tr.addEventListener('click', () => {
    tb.querySelectorAll('tr.sel').forEach((x) => x.classList.remove('sel'));
    tr.classList.add('sel');
    focusTx(results[+tr.dataset.i]);
  }));
}
['#fUser', '#fJudge'].forEach((s) => $(s).addEventListener('change', () => {
  if (s === '#fUser' && $('#fUser').value) { $('#mapUser').value = $('#fUser').value; drawMap(); }
  renderTx();
}));
$('#fText').addEventListener('input', renderTx);
$('#mapUser').addEventListener('change', drawMap);

// ---------- 지도 ----------
let openIw = null;
function clearMap() { mapObjs.forEach((o) => o.setMap(null)); mapObjs = []; openIw = null; }
function drawMap() {
  if (!map) return;
  clearMap();
  const user = $('#mapUser').value;
  const emp = employees[user];
  const bounds = new kakao.maps.LatLngBounds();
  let any = false;
  if (emp && emp.lat != null) {
    const home = new kakao.maps.LatLng(emp.lat, emp.lng);
    const R = Number(emp.radius) || Number(settings.radius) || 1000;
    mapObjs.push(new kakao.maps.CustomOverlay({ map, position: home, yAnchor: 1.3, content: `<div class="home-pin">🏠 ${esc(user)} 집</div>`, zIndex: 5 }));
    mapObjs.push(new kakao.maps.Circle({ map, center: home, radius: R, strokeWeight: 2, strokeColor: '#d92d20', strokeOpacity: .8, fillColor: '#d92d20', fillOpacity: .08 }));
    mapObjs.push(new kakao.maps.Circle({ map, center: home, radius: R * 2, strokeWeight: 1.5, strokeColor: '#e8890c', strokeOpacity: .8, strokeStyle: 'dash', fillOpacity: 0 }));
    bounds.extend(home); any = true;
  }
  // 같은 위치 가맹점은 묶어서 표시
  const groups = new Map();
  for (const r of results) {
    if (r.user !== user || !r.geo) continue;
    const k = `${r.geo.lat.toFixed(5)},${r.geo.lng.toFixed(5)}`;
    const g = groups.get(k) || { lat: r.geo.lat, lng: r.geo.lng, items: [], worst: '정상' };
    g.items.push(r);
    if (JUDGE_ORDER[r.judge] < JUDGE_ORDER[g.worst]) g.worst = r.judge;
    groups.set(k, g);
  }
  for (const g of groups.values()) {
    const pos = new kakao.maps.LatLng(g.lat, g.lng);
    const dot = document.createElement('div');
    dot.className = 'dot';
    dot.style.background = JUDGE_COLOR[g.worst] || '#98a2b3';
    if (g.items.length > 1) { dot.style.width = dot.style.height = '18px'; }
    const ov = new kakao.maps.CustomOverlay({ map, position: pos, content: dot, zIndex: g.worst === '고위험' ? 4 : g.worst === '주의' ? 3 : 2 });
    dot.addEventListener('click', () => showGroup(g, pos));
    mapObjs.push(ov);
    if (g.worst !== '정상' || !any) { bounds.extend(pos); any = true; }
  }
  if (any) map.setBounds(bounds, 40, 40, 40, 40);
}
function showGroup(g, pos) {
  if (openIw) openIw.setMap(null);
  const it = g.items;
  const total = it.reduce((s, r) => s + r.amount, 0);
  const html = document.createElement('div');
  html.className = 'iw';
  html.innerHTML = `<b>${esc(it[0].merchant)}</b>${it.length > 1 ? ` 외` : ''}<br>
    <span class="muted">${esc(it[0].geo.matched || it[0].addr)}</span><br>
    집과 거리 <b>${fmtDist(it[0].dist)}</b> · ${it.length}건 · ${won(total)}원
    <div style="margin-top:6px;max-height:140px;overflow:auto">${it.map((r) => `<div><span class="badge ${JUDGE_CLASS[r.judge]}">${r.judge}</span> ${esc(r.dateStr)}(${r.dow}) ${esc(r.time.slice(0, 5))} ${won(r.amount)}원</div>`).join('')}</div>`;
  html.addEventListener('click', () => { openIw.setMap(null); openIw = null; });
  openIw = new kakao.maps.CustomOverlay({ map, position: pos, content: html, yAnchor: 1, zIndex: 10 });
  mapObjs.push(openIw);
}
function focusTx(r) {
  if (!map || !r) return;
  if ($('#mapUser').value !== r.user) { $('#mapUser').value = r.user; drawMap(); }
  if (!r.geo) return;
  const pos = new kakao.maps.LatLng(r.geo.lat, r.geo.lng);
  map.setLevel(4); map.panTo(pos);
  const k = `${r.geo.lat.toFixed(5)},${r.geo.lng.toFixed(5)}`;
  const items = results.filter((x) => x.user === r.user && x.geo && `${x.geo.lat.toFixed(5)},${x.geo.lng.toFixed(5)}` === k);
  showGroup({ items }, pos);
  $('#map').scrollIntoView({ behavior: 'smooth', block: 'center' });
}

// ---------- 결과 엑셀 ----------
$('#btnExport').addEventListener('click', () => {
  const list = filteredResults();
  const rows = list.map((r) => ({
    판정: r.judge, 사유: [r.flags.join('·'), r.reason].filter(Boolean).join(' / '),
    '집과거리(m)': r.dist == null ? '' : Math.round(r.dist), 좌표기준: r.geo?.method || '', 좌표매칭주소: r.geo?.matched || '',
    순번: r.seq, 카드번호: r.card, 이용자명: r.user, 이용일: r.dateStr, 요일: r.dow, 가맹점명: r.merchant, 이용금액: r.amount,
    '가맹점 주소': r.addr, '사업자 등록번호': r.bizNo, 승인번호: r.approvalNo, 과세유형: r.taxType, 전표구분: r.slipType,
    미화금액: r.usd, 미화환율: r.fx, 승인시간: r.time, 가맹점업종코드: r.mccCode, 가맹점업종명: r.mccName, 유종: r.fuel, 리터당단가: r.unitPrice,
  }));
  const ws = XLSX.utils.json_to_sheet(rows);
  ws['!cols'] = Object.keys(rows[0] || {}).map((k) => ({ wch: ['가맹점 주소', '좌표매칭주소', '사유'].includes(k) ? 40 : k === '가맹점명' ? 28 : 12 }));
  const sum = {};
  for (const r of results) {
    const s = (sum[r.user] ||= { 이용자명: r.user, 분석건수: 0, 총금액: 0, 고위험건수: 0, 고위험금액: 0, 주의건수: 0, 주의금액: 0, 제외건수: 0, 확인불가건수: 0 });
    if (r.judge === '제외') { s.제외건수++; continue; }
    if (r.judge === '확인불가') { s.확인불가건수++; continue; }
    s.분석건수++; s.총금액 += r.amount;
    if (r.judge === '고위험') { s.고위험건수++; s.고위험금액 += r.amount; }
    if (r.judge === '주의') { s.주의건수++; s.주의금액 += r.amount; }
  }
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(Object.values(sum)), '이용자별요약');
  XLSX.utils.book_append_sheet(wb, ws, '거래상세');
  const d = new Date();
  XLSX.writeFile(wb, `법인카드_집근처분석_${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}.xlsx`);
});

// ---------- 직원 집주소 ----------
async function geocodeEmployee(name, render = true) {
  const e = employees[name];
  if (!e || !e.address) return;
  const r = await addressSearch(cleanAddr(e.address));
  if (r.length) Object.assign(e, { lat: +r[0].y, lng: +r[0].x, matched: r[0].road_address?.address_name || r[0].address_name, status: 'ok' });
  else {
    const p = await keywordSearch(cleanAddr(e.address));
    if (p.length) Object.assign(e, { lat: +p[0].y, lng: +p[0].x, matched: `${p[0].place_name} (${p[0].road_address_name || p[0].address_name})`, status: 'ok' });
    else Object.assign(e, { lat: null, lng: null, matched: '', status: 'fail' });
  }
  store.set('employees', employees);
  if (render) renderEmployees();
}

function renderEmployees() {
  const tb = $('#empTbl tbody');
  const names = Object.keys(employees).sort((a, b) => a.localeCompare(b, 'ko'));
  if (!names.length) { tb.innerHTML = '<tr><td colspan="5" class="muted">등록된 직원이 없습니다. 사용내역을 업로드하거나 직접 추가하세요.</td></tr>'; return; }
  tb.innerHTML = names.map((n) => {
    const e = employees[n];
    const st = e.lat != null ? `<span class="badge b-ok">확인</span> <span class="muted">${esc(e.matched || '')}</span>`
      : e.status === 'fail' ? '<span class="badge b-high">주소 확인 실패</span>'
      : e.address ? '<span class="badge b-warn">좌표 미확인</span>' : '<span class="badge b-mute">미입력</span>';
    return `<tr data-name="${esc(n)}"><td>${esc(n)}</td>
      <td><input type="text" class="eAddr" value="${esc(e.address || '')}" style="width:100%" placeholder="집주소 입력"></td>
      <td class="num"><input type="number" class="eRad" value="${esc(e.radius || '')}" placeholder="${settings.radius}" style="width:90px"></td>
      <td>${st}</td>
      <td><button class="btn eGeo">좌표 확인</button> <button class="btn eDel">삭제</button></td></tr>`;
  }).join('');
  tb.querySelectorAll('tr').forEach((tr) => {
    const n = tr.dataset.name;
    tr.style.cursor = 'default';
    tr.querySelector('.eAddr').addEventListener('change', (ev) => {
      employees[n].address = ev.target.value.trim(); employees[n].lat = null; employees[n].status = '';
      store.set('employees', employees); renderEmployees();
    });
    tr.querySelector('.eRad').addEventListener('change', (ev) => { employees[n].radius = ev.target.value; store.set('employees', employees); });
    tr.querySelector('.eGeo').addEventListener('click', async () => {
      try { await loadKakao(); } catch (e) { alert(e.message); return; }
      await geocodeEmployee(n);
      if (results.length) computeResults();
    });
    tr.querySelector('.eDel').addEventListener('click', () => {
      if (!confirm(`${n} 직원 정보를 삭제할까요?`)) return;
      delete employees[n]; store.set('employees', employees); renderEmployees();
    });
  });
}
$('#btnAddEmp').addEventListener('click', () => {
  const n = $('#newEmpName').value.trim();
  if (!n) return;
  employees[n] = { ...(employees[n] || {}), address: $('#newEmpAddr').value.trim(), lat: null, status: '' };
  store.set('employees', employees);
  $('#newEmpName').value = $('#newEmpAddr').value = '';
  renderEmployees();
});
$('#btnGeoAllEmp').addEventListener('click', async () => {
  try { await loadKakao(); } catch (e) { alert(e.message); return; }
  for (const n of Object.keys(employees)) if (employees[n].address) { await geocodeEmployee(n, false); await sleep(60); }
  renderEmployees();
  if (results.length) computeResults();
});
$('#btnEmpTpl').addEventListener('click', () => {
  const rows = Object.entries(employees).map(([n, e]) => ({ 이용자명: n, 집주소: e.address || '', 반경: e.radius || '' }));
  const ws = XLSX.utils.json_to_sheet(rows.length ? rows : [{ 이용자명: '홍길동', 집주소: '서울 송파구 올림픽로 300', 반경: '' }]);
  ws['!cols'] = [{ wch: 12 }, { wch: 50 }, { wch: 8 }];
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, '직원집주소');
  XLSX.writeFile(wb, '직원_집주소.xlsx');
});
$('#btnEmpImport').addEventListener('click', () => $('#empFile').click());
$('#empFile').addEventListener('change', async (ev) => {
  const f = ev.target.files[0]; ev.target.value = '';
  if (!f) return;
  const rows = await readSheetRows(f);
  const h = rows.findIndex((r) => r.some((c) => normHeader(c) === '이용자명'));
  if (h < 0) { alert('이용자명 컬럼을 찾지 못했습니다.'); return; }
  const hd = rows[h].map(normHeader);
  const iN = hd.indexOf('이용자명'), iA = hd.findIndex((x) => x === '집주소' || x === '주소'), iR = hd.indexOf('반경');
  let cnt = 0;
  for (const r of rows.slice(h + 1)) {
    const n = String(r[iN] ?? '').trim();
    if (!n) continue;
    const addr = iA >= 0 ? String(r[iA] ?? '').trim() : '';
    const prev = employees[n] || {};
    employees[n] = { ...prev, address: addr || prev.address || '', radius: iR >= 0 ? String(r[iR] ?? '') : prev.radius || '' };
    if (addr && addr !== prev.address) { employees[n].lat = null; employees[n].status = ''; }
    cnt++;
  }
  store.set('employees', employees);
  renderEmployees();
  alert(`${cnt}명 등록/갱신했습니다. "전체 좌표 확인"을 눌러주세요.`);
});

// ---------- 설정 ----------
function renderSettings() {
  $('#kakaoKey').value = settings.kakaoKey;
  $('#radius').value = settings.radius;
  $('#nightStart').value = settings.nightStart;
  $('#nightEnd').value = settings.nightEnd;
  $('#exOnline').checked = settings.exOnline;
  $('#exForeign').checked = settings.exForeign;
  $('#holidays').value = settings.holidays;
  $('#geminiKey').value = settings.geminiKey;
  $('#geminiModel').value = settings.geminiModel;
  $('#aiMask').checked = settings.aiMask;
  $('#cacheInfo').textContent = `저장된 가맹점 좌표 ${Object.keys(geoCache).length.toLocaleString()}건`;
}
$('#btnSaveKey').addEventListener('click', () => {
  const k = $('#kakaoKey').value.trim();
  const changed = k !== settings.kakaoKey;
  settings.kakaoKey = k; store.set('settings', settings);
  if (changed && kakaoReady) { location.reload(); return; }
  loadKakao().then(() => alert('카카오 지도에 연결되었습니다.')).catch((e) => alert(e.message));
});
$('#btnSaveSet').addEventListener('click', () => {
  Object.assign(settings, {
    radius: Number($('#radius').value) || 1000,
    nightStart: Number($('#nightStart').value) || 0, nightEnd: Number($('#nightEnd').value) || 0,
    exOnline: $('#exOnline').checked, exForeign: $('#exForeign').checked, holidays: $('#holidays').value,
  });
  store.set('settings', settings);
  renderEmployees();
  if (results.length) computeResults();
  alert('저장했습니다.');
});
$('#btnClearCache').addEventListener('click', () => {
  if (!confirm('저장된 가맹점 좌표를 모두 지웁니다. 다음 분석 때 다시 조회합니다.')) return;
  geoCache = {}; store.set('geoCache', geoCache); renderSettings();
});

// ---------- Gemini AI (이상징후 요약 · 경영진 보고서) ----------
const GEMINI_BASE = 'https://generativelanguage.googleapis.com/v1beta';
let lastReport = null;   // {title, html}

function saveGeminiFields() {
  settings.geminiKey = $('#geminiKey').value.trim();
  settings.geminiModel = $('#geminiModel').value.trim().replace(/^models\//, '') || 'gemini-3.5-flash';
  settings.aiMask = $('#aiMask').checked;
  store.set('settings', settings);
}

// 선택한 모델이 혼잡(503)·한도초과(429)·종료(404)·응답지연이면 아래 순서로 자동 전환
const GEMINI_FALLBACKS = ['gemini-3.5-flash', 'gemini-flash-latest', 'gemini-flash-lite-latest', 'gemini-3.1-flash-lite'];

async function callGeminiOnce(model, prompt, timeoutMs) {
  const ctl = new AbortController();
  const timer = setTimeout(() => ctl.abort(), timeoutMs);
  try {
    const res = await fetch(`${GEMINI_BASE}/models/${encodeURIComponent(model)}:generateContent`, {
      method: 'POST', signal: ctl.signal,
      headers: { 'Content-Type': 'application/json', 'x-goog-api-key': settings.geminiKey },
      body: JSON.stringify({ contents: [{ role: 'user', parts: [{ text: prompt }] }], generationConfig: { temperature: 0.3 } }),
    });
    const json = await res.json().catch(() => ({}));
    if (!res.ok) {
      const msg = json.error?.message || res.statusText;
      const err = new Error(`${res.status}: ${msg}`);
      err.retry = [404, 429, 500, 503, 504].includes(res.status);
      if (res.status === 400 && /API key/i.test(msg)) err.message = 'Gemini API 키가 올바르지 않습니다. 설정 탭에서 다시 입력하세요.';
      if (res.status === 403) err.message = 'Gemini API 키 권한이 없습니다 (403). AI Studio에서 키 상태를 확인하세요.';
      throw err;
    }
    const text = (json.candidates?.[0]?.content?.parts || []).filter((p) => !p.thought).map((p) => p.text || '').join('');
    if (!text) {
      const err = new Error('응답이 비어 있음 (' + (json.candidates?.[0]?.finishReason || json.promptFeedback?.blockReason || '사유 불명') + ')');
      err.retry = true;
      throw err;
    }
    return { text, model: json.modelVersion || model };
  } catch (e) {
    if (e.name === 'AbortError') { const err = new Error(`응답 시간 초과(${timeoutMs / 1000}초)`); err.retry = true; throw err; }
    if (e instanceof TypeError) { const err = new Error('네트워크 오류 — 인터넷 연결을 확인하세요.'); throw err; }
    throw e;
  } finally { clearTimeout(timer); }
}

/** @returns {Promise<{text:string, model:string}>} */
async function callGemini(prompt, { timeoutMs = 90000, onStatus = () => {} } = {}) {
  if (!settings.geminiKey) throw new Error('Gemini API 키가 없습니다. 설정 탭에서 입력하세요.');
  const models = [...new Set([settings.geminiModel, ...GEMINI_FALLBACKS])];
  const errors = [];
  for (const m of models) {
    onStatus(`${m} 호출 중…`);
    try { return await callGeminiOnce(m, prompt, timeoutMs); } catch (e) {
      errors.push(`· ${m} → ${e.message}`);
      if (!e.retry) throw new Error(e.message);
      onStatus(`${m} 응답 실패, 다른 모델로 재시도…`);
    }
  }
  throw new Error('모든 Gemini 모델 호출에 실패했습니다. 잠시 후 다시 시도하세요.\n\n' + errors.join('\n'));
}

$('#btnSaveGemini').addEventListener('click', async () => {
  saveGeminiFields();
  const st = $('#geminiState'), btn = $('#btnSaveGemini');
  btn.disabled = true;
  st.style.color = '';
  try {
    const r = await callGemini("연결 테스트입니다. 'OK' 한 단어로만 답하세요.", { timeoutMs: 20000, onStatus: (s) => { st.textContent = s; } });
    st.style.color = 'var(--ok)';
    st.textContent = `✔ 연결 성공 (응답 모델: ${r.model})`;
  } catch (e) {
    st.style.color = 'var(--high)';
    st.textContent = '✖ ' + e.message.split('\n')[0];
    alert(e.message);
  } finally { btn.disabled = false; }
});
$('#btnListModels').addEventListener('click', async () => {
  saveGeminiFields();
  if (!settings.geminiKey) { alert('Gemini API 키를 먼저 입력하세요.'); return; }
  try {
    const res = await fetch(`${GEMINI_BASE}/models?pageSize=200`, { headers: { 'x-goog-api-key': settings.geminiKey } });
    const json = await res.json();
    if (!res.ok) throw new Error(json.error?.message || res.statusText);
    const names = (json.models || [])
      .filter((m) => (m.supportedGenerationMethods || []).includes('generateContent') && /gemini/.test(m.name))
      .map((m) => m.name.replace(/^models\//, ''));
    $('#geminiModels').innerHTML = names.map((n) => `<option value="${esc(n)}">`).join('');
    alert(`사용 가능한 모델 ${names.length}개를 불러왔습니다. 모델 칸을 클릭해 선택하세요.\n\n${names.slice(0, 15).join('\n')}${names.length > 15 ? '\n…' : ''}`);
  } catch (e) { alert('모델 목록 조회 실패: ' + e.message); }
});

/** Gemini에 보낼 데이터. 카드번호·사업자번호·집주소는 넣지 않고, 이용자명은 옵션에 따라 가명 처리 */
function buildAiData(user) {
  const src = results.filter((r) => !user || r.user === user);
  const users = [...new Set(src.map((r) => r.user))].sort((a, b) => a.localeCompare(b, 'ko'));
  const alias = {};
  users.forEach((u, i) => { alias[u] = settings.aiMask ? `직원${String(i + 1).padStart(2, '0')}` : u; });
  const sum = (a) => a.reduce((s, r) => s + r.amount, 0);
  const pick = (j) => src.filter((r) => r.judge === j);
  const valid = src.filter((r) => ['고위험', '주의', '정상'].includes(r.judge));
  const flagged = src.filter((r) => r.judge === '고위험' || r.judge === '주의');
  const dates = src.map((r) => r.dateStr).filter(Boolean).sort();

  const byUser = users.map((u) => {
    const v = valid.filter((r) => r.user === u), h = v.filter((r) => r.judge === '고위험'), w = v.filter((r) => r.judge === '주의');
    return {
      이용자: alias[u], 분석건수: v.length, 총금액: sum(v), 고위험건수: h.length, 고위험금액: sum(h), 주의건수: w.length, 주의금액: sum(w),
      '집근처비율(%)': v.length ? Math.round((h.length + w.length) / v.length * 1000) / 10 : 0,
      집주소등록: employees[u]?.lat != null,
    };
  });
  const group = (list, keyFn) => {
    const m = {};
    for (const r of list) { const k = keyFn(r); (m[k] ||= { 건수: 0, 금액: 0 }); m[k].건수++; m[k].금액 += r.amount; }
    return Object.entries(m).map(([k, v]) => ({ 항목: k, ...v })).sort((a, b) => b.금액 - a.금액);
  };
  const repeat = group(flagged, (r) => `${alias[r.user]} | ${r.merchant}`).filter((x) => x.건수 >= 2);
  const order = { 고위험: 0, 주의: 1 };
  const tx = [...flagged].sort((a, b) => order[a.judge] - order[b.judge] || b.amount - a.amount).slice(0, 150).map((r) => ({
    판정: r.judge, 이용자: alias[r.user], 일자: r.dateStr, 요일: r.dow, 시간: r.time.slice(0, 5), 가맹점: r.merchant,
    업종: r.mccName, 금액: r.amount, '집과거리(m)': Math.round(r.dist), 사유: r.flags.filter((f) => f !== '취소').join('·') || '평일주간',
  }));

  const data = {
    분석기간: dates.length ? `${dates[0]} ~ ${dates[dates.length - 1]}` : '',
    분석대상: user ? alias[user] : `전체 ${users.length}명`,
    판정기준: {
      집근처반경m: Number(settings.radius), 인접반경m: Number(settings.radius) * 2, 심야: `${settings.nightStart}시~${settings.nightEnd}시`,
      고위험: '집근처 + (주말·공휴일 또는 심야)', 주의: '집근처(평일 주간) 또는 인접 + (주말·공휴일 또는 심야)',
      제외: [settings.exOnline && '온라인/PG 결제', settings.exForeign && '해외 결제'].filter(Boolean).join(', ') || '없음',
    },
    전체현황: {
      업로드건수: src.length, 업로드금액: sum(src), 분석대상건수: valid.length, 분석대상금액: sum(valid),
      고위험건수: pick('고위험').length, 고위험금액: sum(pick('고위험')), 주의건수: pick('주의').length, 주의금액: sum(pick('주의')),
      제외건수: pick('제외').length, 제외금액: sum(pick('제외')), 확인불가건수: pick('확인불가').length,
    },
    이용자별: byUser,
    의심거래_업종별: group(flagged, (r) => r.mccName || '미분류'),
    의심거래_요일시간대: {
      주말공휴일: flagged.filter((r) => r.flags.includes('주말') || r.flags.includes('공휴일')).length,
      심야: flagged.filter((r) => r.flags.includes('심야')).length,
      평일주간: flagged.filter((r) => !r.flags.some((f) => ['주말', '공휴일', '심야'].includes(f))).length,
    },
    반복사용_가맹점: repeat,
    의심거래_목록: tx,
    의심거래_목록_생략건수: Math.max(0, flagged.length - tx.length),
  };
  return { data, alias, src, flagged, valid };
}

const AI_RULES = `규칙:
- 아래 JSON에 있는 사실과 숫자만 사용하고, 없는 내용·수치를 지어내지 말 것.
- 금액은 "1,234,000원"처럼 천 단위 콤마를 붙일 것.
- 집근처 사용 자체가 곧 부정사용은 아님. "소명 확인 필요", "가능성" 등으로 표현하고 단정하지 말 것.
- 이용자 이름은 JSON에 적힌 그대로 사용할 것.
- 한국어 마크다운으로, 개조식(~함, ~임)으로 간결하게 작성할 것. 문서 제목(# 수준)은 쓰지 말 것.`;

function summaryPrompt(data) {
  return `당신은 기업 내부감사팀의 법인카드 모니터링 분석가입니다.
아래 JSON은 직원 법인카드 사용내역을 직원 집 위치와 비교한 "집근처 사용" 분석 결과입니다.
이상징후를 요약해 주세요.

형식:
## 핵심 요약
(3~5개 bullet, 가장 중요한 수치 포함)
## 주요 이상징후 패턴
(시간대·요일·업종·반복 가맹점 등 패턴별 bullet, 근거 수치와 대표 거래 예시)
## 이용자별 확인 필요 사항
(표: 이용자 | 위험도(상/중/하) | 주요 내용 | 소명 요청 사항)
## 추가 점검 권고
(3~5개 bullet)

${AI_RULES}

데이터(JSON):
${JSON.stringify(data)}`;
}

function reportPrompt(data) {
  return `당신은 기업 내부감사팀장입니다. 아래 JSON(법인카드 "집근처 사용" 분석 결과)을 바탕으로
CEO·CFO 등 경영진에게 보고할 보고서 본문을 작성하세요. 바쁜 경영진이 1~2분 안에 읽을 수 있게 핵심부터 씁니다.

형식 (이 순서와 제목을 그대로 사용):
## 1. 보고 개요
(목적, 분석 기간, 분석 방법을 2~3줄)
## 2. 핵심 결과
(경영진이 알아야 할 핵심 메시지 3개, 각 수치 포함)
## 3. 주요 이상징후
(패턴별로 정리, 대표 사례 포함)
## 4. 이용자별 리스크 평가
(표: 이용자 | 위험도(상/중/하) | 의심 건수·금액 | 평가 의견)
## 5. 시사점 및 원인 추정
## 6. 개선 권고안
(### 단기 조치 / ### 제도 개선 으로 구분)
## 7. 향후 조치 계획
(표: 조치 사항 | 담당 | 일정 — 일정은 "보고 후 1주 이내" 같은 상대 표현)

${AI_RULES}

데이터(JSON):
${JSON.stringify(data)}`;
}

/** AI 결과의 가명(직원01…)을 실명으로 되돌린다 (브라우저 안에서만) */
function unmask(text, alias) {
  if (!settings.aiMask) return text;
  const back = Object.fromEntries(Object.entries(alias).map(([real, a]) => [a, real]));
  const keys = Object.keys(back).sort((a, b) => b.length - a.length);
  if (!keys.length) return text;
  return text.replace(new RegExp(keys.join('|'), 'g'), (m) => back[m]);
}
const md = (text) => DOMPurify.sanitize(marked.parse(text));

function kpiGridHtml(src, valid) {
  const sum = (a) => a.reduce((s, r) => s + r.amount, 0);
  const h = src.filter((r) => r.judge === '고위험'), w = src.filter((r) => r.judge === '주의');
  const pct = valid.length ? ((h.length + w.length) / valid.length * 100).toFixed(1) : '0.0';
  return `<div class="kgrid">
    <div>분석 대상<b>${valid.length.toLocaleString()}건</b>${won(sum(valid))}원</div>
    <div>고위험<b style="color:#d92d20">${h.length.toLocaleString()}건</b>${won(sum(h))}원</div>
    <div>주의<b style="color:#e8890c">${w.length.toLocaleString()}건</b>${won(sum(w))}원</div>
    <div>집근처 의심 비율<b>${pct}%</b>건수 기준</div></div>`;
}
function flaggedTableHtml(flagged, limit = 20) {
  const order = { 고위험: 0, 주의: 1 };
  const list = [...flagged].sort((a, b) => order[a.judge] - order[b.judge] || b.amount - a.amount).slice(0, limit);
  if (!list.length) return '<p>해당 거래 없음</p>';
  return `<table><thead><tr><th>판정</th><th>이용자</th><th>일시</th><th>가맹점</th><th>업종</th><th class="num">금액</th><th class="num">집과 거리</th><th>사유</th></tr></thead><tbody>
    ${list.map((r) => `<tr><td>${r.judge}</td><td>${esc(r.user)}</td><td>${esc(r.dateStr)}(${r.dow}) ${esc(r.time.slice(0, 5))}</td>
      <td>${esc(r.merchant)}</td><td>${esc(r.mccName)}</td><td class="num">${won(r.amount)}원</td><td class="num">${fmtDist(r.dist)}</td>
      <td>${esc(r.reason)}</td></tr>`).join('')}</tbody></table>
    ${flagged.length > limit ? `<p class="muted">외 ${flagged.length - limit}건 (전체 목록은 대시보드 엑셀 다운로드 참조)</p>` : ''}`;
}

async function runAi(kind) {
  if (!results.length) { alert('먼저 대시보드에서 집근처 분석을 실행하세요.'); return; }
  saveGeminiFields();
  const user = $('#aiUser').value;
  const { data, alias, src, flagged, valid } = buildAiData(user);
  const btns = ['#btnAiSummary', '#btnAiReport'].map((s) => $(s));
  btns.forEach((b) => { b.disabled = true; });
  $('#aiState').textContent = `Gemini가 작성 중입니다… (보통 10~40초)`;
  try {
    const ai = await callGemini(kind === 'report' ? reportPrompt(data) : summaryPrompt(data), {
      onStatus: (s) => { $('#aiState').textContent = `Gemini 작성 중 — ${s} (보통 10~40초)`; },
    });
    const text = unmask(ai.text, alias);
    const today = ymd(new Date());
    const target = user || '전체 이용자';
    const dept = $('#rptDept').value.trim(), author = $('#rptAuthor').value.trim();
    store.set('rptMeta', { dept, author });
    const meta = `보고일 ${today}${dept ? ' · ' + esc(dept) : ''}${author ? ' · 작성 ' + esc(author) : ''} · 대상 ${esc(target)} · 기간 ${esc(data.분석기간)}`;
    const note = `<div class="note">※ 본 문서의 건수·금액은 분석 시스템이 계산한 값이며, 해석 문장은 AI(${esc(ai.model)})가 작성하였습니다.
      집근처 사용은 이상징후일 뿐 부정사용 확정이 아니므로, 반드시 사용자 소명과 증빙 확인을 거쳐 판단하시기 바랍니다.</div>`;
    let title, html;
    if (kind === 'report') {
      title = '법인카드 집근처 사용 점검 결과 보고';
      html = `<h1>${title}</h1><div class="meta">${meta}</div>${kpiGridHtml(src, valid)}${md(text)}
        <h2>[별첨] 주요 의심 거래 (상위 20건)</h2>${flaggedTableHtml(flagged)}${note}`;
    } else {
      title = '법인카드 이상징후 요약';
      html = `<h1>${title}</h1><div class="meta">${meta}</div>${kpiGridHtml(src, valid)}${md(text)}${note}`;
    }
    lastReport = { title, html };
    $('#aiOut').innerHTML = html;
    $('#aiState').textContent = `작성 완료 (${ai.model})`;
    ['#btnRptCopy', '#btnRptWord', '#btnRptHtml', '#btnRptPrint'].forEach((s) => { $(s).disabled = false; });
  } catch (e) {
    $('#aiState').textContent = '';
    alert(e.message);
  } finally {
    btns.forEach((b) => { b.disabled = false; });
  }
}
$('#btnAiSummary').addEventListener('click', () => runAi('summary'));
$('#btnAiReport').addEventListener('click', () => runAi('report'));

const REPORT_CSS = `body{font-family:"Malgun Gothic",sans-serif;font-size:11pt;line-height:1.7;color:#1d2330;max-width:900px;margin:24px auto;padding:0 16px}
h1{font-size:18pt;border-bottom:2px solid #1d2330;padding-bottom:6px}h2{font-size:13pt;margin-top:20px;border-left:4px solid #2f6fed;padding-left:8px}
h3{font-size:11.5pt}table{border-collapse:collapse;width:100%;font-size:9.5pt;margin:6px 0 12px}th,td{border:1px solid #c9ced6;padding:5px 7px;text-align:left}
th{background:#f2f4f7}.num{text-align:right}.meta,.muted,.note{color:#667085;font-size:9.5pt}.note{border-top:1px solid #ddd;margin-top:20px;padding-top:6px}
.kgrid{display:flex;gap:8px;margin:10px 0}.kgrid div{flex:1;border:1px solid #ddd;border-radius:6px;padding:6px 8px;font-size:9.5pt}.kgrid b{display:block;font-size:14pt}`;
function reportDoc() {
  return `<!doctype html><html lang="ko"><head><meta charset="utf-8"><title>${esc(lastReport.title)}</title><style>${REPORT_CSS}</style></head><body>${lastReport.html}</body></html>`;
}
function download(name, content, type) {
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob(['﻿' + content], { type }));
  a.download = name;
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 1000);
}
const reportFileName = (ext) => `${lastReport.title.replace(/\s+/g, '_')}_${ymd(new Date())}.${ext}`;
$('#btnRptWord').addEventListener('click', () => download(reportFileName('doc'), reportDoc(), 'application/msword'));
$('#btnRptHtml').addEventListener('click', () => download(reportFileName('html'), reportDoc(), 'text/html'));
$('#btnRptPrint').addEventListener('click', () => window.print());
$('#btnRptCopy').addEventListener('click', async () => {
  const html = $('#aiOut').innerHTML, text = $('#aiOut').innerText;
  try {
    await navigator.clipboard.write([new ClipboardItem({ 'text/html': new Blob([html], { type: 'text/html' }), 'text/plain': new Blob([text], { type: 'text/plain' }) })]);
  } catch { await navigator.clipboard.writeText(text); }
  $('#aiState').textContent = '클립보드에 복사했습니다 (메일·한글·워드에 붙여넣기 가능)';
});

// ---------- 시작 ----------
const rptMeta = store.get('rptMeta', {});
$('#rptDept').value = rptMeta.dept || '';
$('#rptAuthor').value = rptMeta.author || '';
renderSettings();
renderEmployees();
renderAll();
if (settings.kakaoKey) loadKakao().catch((e) => console.warn(e.message));

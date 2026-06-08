/* ===========================================================
   판매 베스트 대시보드 — Apps Script API
   -----------------------------------------------------------
   시트 컬럼 구조 (0-based):
   A(0)=이미지  B(1)=대분류  C(2)=소분류  D(3)=년도
   E(4)=소분류&시즌  F(5)=시즌  G(6)=상품명  H(7)=원가
   I(8)=판매가  J(9)=대표바코드
   K(10)=26년누적판매수량  L(11)=26년누적판매금액   ← 각 매장 누적
   M(12)=1월실판매수량  N(13)=1월낮판매금액
   O(14)=W1수량  P(15)=W1금액
   Q(16)=2주수량  R(17)=2주매출  ...

   26'PP 시트:
   U(20)=전체누적판매수량  V(21)=전체누적판매금액

   상품리스트 시트:
   A(0)=상품명  H(7)=대표이미지URL
   =========================================================== */

const STORE_NAMES = [
  '남악','부천','의정부','인천','유성','여수','율하','충장',
  '동수원','송파','제주','괴정','강서','순천','평택','고척',
  '원주','마리오','창원'
];

const STORE_COLORS = {
  '남악':'#1565c0','부천':'#2e7d32','의정부':'#e65100','인천':'#6a1b9a',
  '유성':'#00695c','여수':'#ad1457','율하':'#558b2f','충장':'#283593',
  '동수원':'#f57f17','송파':'#00838f','제주':'#4e342e','괴정':'#37474f',
  '강서':'#9e9d24','순천':'#bf360c','평택':'#4527a0','고척':'#01579b',
  '원주':'#1b5e20','마리오':'#ff8f00','창원':'#0d47a1'
};

// 고정 컬럼 인덱스
const COL_CAT   = 1;   // B: 대분류
const COL_NAME  = 6;   // G: 상품명
const COL_PRICE = 8;   // I: 판매가
const COL_CUM_QTY = 10; // K: 매장별 26년 누적 판매수량
const COL_CUM_REV = 11; // L: 매장별 26년 누적 판매금액
const COL_OV_QTY  = 20; // U: 26'PP 전체 누적 판매수량
const COL_OV_REV  = 21; // V: 26'PP 전체 누적 판매금액

const OVERALL_SHEET    = "26'PP";      // 전체 누적 시트명
const IMAGE_SHEET_NAME = '상품리스트'; // 이미지 URL 참조 시트

// ── 이미지 URL 조회 테이블 로드 ──────────────────────────
// 상품리스트 시트의 A열(상품명) → H열(대표이미지URL) 매핑 반환
function loadImageMap(ss) {
  const map   = {};
  const sheet = ss.getSheetByName(IMAGE_SHEET_NAME);
  if (!sheet) return map;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return map;

  // 헤더(1행) 건너뜀, A~H열(8컬럼)만 읽기
  const data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
  data.forEach(row => {
    const name = String(row[0] || '').trim();          // A열: 상품명
    const url  = String(row[7] || '').trim();          // H열: 대표이미지
    if (name && url && url.startsWith('http')) {
      map[name] = url;
    }
  });
  return map;
}

// ── 세션 진단 (1회 실행 함수) ────────────────────────────────
function debugSession() {
  const cookie = buildEZCookie();
  Logger.log('PHPSESSID 앞 10자: ' + cookie.substring(10, 20) + '...');

  // rt_status 테스트
  const r1 = UrlFetchApp.fetch('https://ecn5.ezadmin.co.kr/function.php', {
    method: 'POST',
    headers: { 'Cookie': cookie, 'Content-Type': 'application/x-www-form-urlencoded' },
    payload: 'template=main&action=rt_status', muteHttpExceptions: true
  });
  Logger.log('[rt_status] 코드: ' + r1.getResponseCode() + ' / 응답: ' + r1.getContentText().substring(0, 150));

  // 실제 E700 데이터 요청 테스트 (남악, 어제 1건)
  const today = new Date();
  today.setDate(today.getDate() - 1);
  const ymd = Utilities.formatDate(today, 'Asia/Seoul', 'yyyy-MM-dd');
  const r2 = UrlFetchApp.fetch('https://ecn5.ezadmin.co.kr/function.php', {
    method: 'POST',
    headers: { 'Cookie': cookie, 'Content-Type': 'application/x-www-form-urlencoded',
               'Referer': 'https://ecn5.ezadmin.co.kr/template.html?template=E700' },
    payload: 'template=E700&action=grid&str_desc_select=89&start_date=' + ymd + '&end_date=' + ymd + '&sorting=qty&limit=1&category=0&time_check=false',
    muteHttpExceptions: true
  });
  const h = r2.getContentText();
  Logger.log('[E700] 세션만료여부: ' + h.includes('세션이 종료'));
  Logger.log('[E700] 응답 앞부분: ' + h.substring(0, 200));
}

// ── 헤더 구조 확인 (1회 실행 함수) ──────────────────────
function debugHeaders() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(STORE_NAMES[0]);
  if (!sheet) { Logger.log('시트 없음'); return; }

  const lastCol = sheet.getLastColumn();
  Logger.log('총 열: ' + lastCol);

  for (let hRow = 1; hRow <= 2; hRow++) {
    const headers = sheet.getRange(hRow, 1, 1, lastCol).getValues()[0];
    headers.forEach((h, i) => {
      if (String(h || '').trim()) {
        Logger.log('행' + hRow + ' [' + i + '] = "' + String(h).replace(/\n/g, '\\n') + '"');
      }
    });
  }
}


// ── EZAdmin 순차 요청 + 세션 로테이션 자동 갱신 ──────────────
// EZAdmin은 매 요청마다 PHPSESSID를 교체함 → 순차로 최신 쿠키를 사용해야 함
function fetchEZBatch(requests) {
  const htmls = new Array(requests.length).fill('');
  const props = PropertiesService.getScriptProperties();

  for (let i = 0; i < requests.length; i++) {
    // 항상 최신 PHPSESSID로 쿠키 재구성
    const req = Object.assign({}, requests[i], {
      headers: Object.assign({}, requests[i].headers, { 'Cookie': buildEZCookie() })
    });

    const resp = UrlFetchApp.fetch(req.url, req);

    // 응답에서 새 PHPSESSID 자동 저장 (세션 로테이션 대응)
    const sc = [].concat(resp.getAllHeaders()['Set-Cookie'] || []);
    const ns = sc.map(c => c.split(';')[0]).find(c => /^PHPSESSID=.+/.test(c));
    if (ns) props.setProperty('PHPSESSID', ns.split('=')[1]);

    const h = resp.getContentText();
    if (!h.includes('세션이 종료')) {
      htmls[i] = h;
    } else {
      // 세션 만료 시 1회 즉시 재시도 (다른 서버로 라우팅될 수 있음)
      Utilities.sleep(300);
      const req2 = Object.assign({}, req, {
        headers: Object.assign({}, req.headers, { 'Cookie': buildEZCookie() })
      });
      const resp2 = UrlFetchApp.fetch(req2.url, req2);
      const sc2 = [].concat(resp2.getAllHeaders()['Set-Cookie'] || []);
      const ns2 = sc2.map(c => c.split(';')[0]).find(c => /^PHPSESSID=.+/.test(c));
      if (ns2) props.setProperty('PHPSESSID', ns2.split('=')[1]);
      htmls[i] = resp2.getContentText().includes('세션이 종료') ? '' : resp2.getContentText();
    }
  }

  return htmls.map(h => ({ getContentText: () => h }));
}

// ── 진입점 ────────────────────────────────────────────────
function doGet(e) {
  try {
    const p         = (e && e.parameter) ? e.parameter : {};
    const noCache   = p.nc === '1';
    const today     = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');
    const startDate = p.startDate || today;
    const endDate   = p.endDate   || today;
    const view      = p.view || 'overall'; // 'overall' | 'stores'

    return view === 'stores'
      ? getStoresView(startDate, endDate, noCache)
      : getOverallView(startDate, endDate, noCache);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── 기간별 일괄 계산 & 캐시 ────────────────────────────────────
// • 전체 베스트: str_desc_select 없이 1번 요청 → 이지어드민 전체 데이터와 정확히 일치
// • 매장별 베스트: EZ_STORES 순차 조회
function computeAndCachePeriod(startDate, endDate) {
  const cache = CacheService.getScriptCache();
  const now   = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy.MM.dd(EEE) HH:mm');

  // ── 1. 전체 베스트 ─────────────────────────────────────────────
  const allResp = UrlFetchApp.fetch('https://ecn5.ezadmin.co.kr/function.php', {
    method : 'POST',
    headers: { 'Cookie': buildEZCookie(), 'Content-Type': 'application/x-www-form-urlencoded',
               'Referer': 'https://ecn5.ezadmin.co.kr/template.html?template=E700' },
    payload: 'template=E700&action=grid' +
             '&start_date=' + startDate + '&end_date=' + endDate +
             '&sorting=qty&limit=500&category=0&time_check=false',
    muteHttpExceptions: true
  });

  // 세션 로테이션 저장
  const sc = [].concat(allResp.getAllHeaders()['Set-Cookie'] || []);
  const ns = sc.map(c => c.split(';')[0]).find(c => /^PHPSESSID=.+/.test(c));
  if (ns) PropertiesService.getScriptProperties().setProperty('PHPSESSID', ns.split('=')[1]);

  const allHtml = allResp.getContentText();
  if (allHtml.includes('세션이 종료')) throw new Error('session_expired');

  const overallItems = parseEZResponse(allHtml) || [];
  const overallResult = {
    updatedAt: now, view: 'overall', startDate, endDate,
    overall: aggregateEZProducts(overallItems).slice(0, 20)
  };
  try { cache.put('ez_overall_' + startDate + '_' + endDate, JSON.stringify(overallResult), 21600); } catch(_) {}
  Logger.log('  전체: ' + overallResult.overall.length + '개 상품');

  // ── 2. 매장별 베스트: EZ_STORES 순차 조회 ────────────────────────
  const storeRequests = EZ_STORES.map(store => ({
    url    : 'https://ecn5.ezadmin.co.kr/function.php',
    method : 'POST',
    headers: { 'Cookie': buildEZCookie(), 'Content-Type': 'application/x-www-form-urlencoded',
               'Referer': 'https://ecn5.ezadmin.co.kr/template.html?template=E700' },
    payload: 'template=E700&action=grid&str_desc_select=' + store.code +
             '&start_date=' + startDate + '&end_date=' + endDate +
             '&sorting=qty&limit=500&category=0&time_check=false',
    followRedirects: true, muteHttpExceptions: true
  }));

  const storeResponses = fetchEZBatch(storeRequests);

  const storeResults = EZ_STORES.map((store, idx) => {
    const html  = storeResponses[idx].getContentText();
    const items = (html && !html.includes('세션이 종료')) ? (parseEZResponse(html) || []) : [];
    return { name: store.name, color: store.color, products: aggregateEZProducts(items).slice(0, 10) };
  });

  const storesResult = {
    updatedAt: now, view: 'stores', startDate, endDate, stores: storeResults
  };
  try { cache.put('ez_stores_' + startDate + '_' + endDate, JSON.stringify(storesResult), 21600); } catch(_) {}

  return overallResult.overall.length;
}

// ── 전체 조회 응답에 매장 정보가 포함되는지 확인 (1회 실행용) ────
// 실행 후 로그에서 '상품 필드 목록'을 확인하세요.
// str_desc 또는 store 관련 필드가 있으면 매장별도 1번 요청으로 처리 가능
function debugAllStoresResponse() {
  const ymd = Utilities.formatDate(new Date(new Date().getTime() - 86400000), 'Asia/Seoul', 'yyyy-MM-dd');
  const resp = UrlFetchApp.fetch('https://ecn5.ezadmin.co.kr/function.php', {
    method : 'POST',
    headers: { 'Cookie': buildEZCookie(), 'Content-Type': 'application/x-www-form-urlencoded',
               'Referer': 'https://ecn5.ezadmin.co.kr/template.html?template=E700' },
    payload: 'template=E700&action=grid' +
             '&start_date=' + ymd + '&end_date=' + ymd +
             '&sorting=qty&limit=10&category=0&time_check=false',
    muteHttpExceptions: true
  }).getContentText();

  if (resp.includes('세션이 종료')) { Logger.log('❌ 세션 만료'); return; }

  const items = parseEZResponse(resp);
  Logger.log('상품 수: ' + items.length);
  if (items.length > 0) {
    Logger.log('상품 필드 목록: ' + Object.keys(items[0]).join(', '));
    Logger.log('첫 번째 상품 전체: ' + JSON.stringify(items[0]));
    Logger.log('두 번째 상품 전체: ' + JSON.stringify(items[1] || {}));
  }
}

// ── 전체 매장 뷰: str_desc_select 없이 1번 요청 → 이지어드민 전체와 일치 ─
function getOverallView(startDate, endDate, noCache) {
  const cacheKey = 'ez_overall_' + startDate + '_' + endDate;
  const cache    = CacheService.getScriptCache();
  if (!noCache) {
    const cached = cache.get(cacheKey);
    if (cached) return ContentService.createTextOutput(cached).setMimeType(ContentService.MimeType.JSON);
  }

  // str_desc_select 생략 = 전 매장 조회
  const resp = UrlFetchApp.fetch('https://ecn5.ezadmin.co.kr/function.php', {
    method : 'POST',
    headers: { 'Cookie': buildEZCookie(), 'Content-Type': 'application/x-www-form-urlencoded',
               'Referer': 'https://ecn5.ezadmin.co.kr/template.html?template=E700' },
    payload: 'template=E700&action=grid' +
             '&start_date=' + startDate + '&end_date=' + endDate +
             '&sorting=qty&limit=500&category=0&time_check=false',
    muteHttpExceptions: true
  });

  // 세션 로테이션 저장
  const sc = [].concat(resp.getAllHeaders()['Set-Cookie'] || []);
  const ns = sc.map(c => c.split(';')[0]).find(c => /^PHPSESSID=.+/.test(c));
  if (ns) PropertiesService.getScriptProperties().setProperty('PHPSESSID', ns.split('=')[1]);

  const html = resp.getContentText();
  if (html.includes('세션이 종료')) {
    return ContentService.createTextOutput(JSON.stringify({ error: 'session_expired' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const items   = parseEZResponse(html) || [];
  const overall = aggregateEZProducts(items).slice(0, 20);
  const result  = {
    updatedAt: Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy.MM.dd(EEE) HH:mm'),
    view: 'overall', startDate, endDate, overall
  };

  const jsonStr = JSON.stringify(result);
  try { cache.put(cacheKey, jsonStr, 21600); } catch(_) {}
  return ContentService.createTextOutput(jsonStr).setMimeType(ContentService.MimeType.JSON);
}

// ── 매장별 뷰: 19개 병렬 조회 ─────────────────────────────
function getStoresView(startDate, endDate, noCache) {
  const cacheKey = 'ez_stores_' + startDate + '_' + endDate;
  const cache    = CacheService.getScriptCache();
  if (!noCache) {
    const cached = cache.get(cacheKey);
    if (cached) return ContentService.createTextOutput(cached).setMimeType(ContentService.MimeType.JSON);
  }

  // 세션 유효성: 첫 번째 병렬 요청 결과로 판단 (rt_status 불필요)

  const cookie   = buildEZCookie(); // autoLogin으로 갱신된 쿠키 사용
  const requests = EZ_STORES.map(store => ({
    url    : 'https://ecn5.ezadmin.co.kr/function.php',
    method : 'POST',
    headers: { 'Cookie': cookie, 'Content-Type': 'application/x-www-form-urlencoded',
               'Referer': 'https://ecn5.ezadmin.co.kr/template.html?template=E700' },
    payload: 'template=E700&action=grid&str_desc_select=' + store.code +
             '&start_date=' + startDate + '&end_date=' + endDate +
             '&sorting=qty&limit=500&category=0&time_check=false',
    followRedirects: true, muteHttpExceptions: true
  }));

  const responses = fetchEZBatch(requests);

  const stores = EZ_STORES.map((store, idx) => {
    const html  = responses[idx].getContentText();
    const items = html.includes('세션이 종료') ? [] : (parseEZResponse(html) || []);
    return { name: store.name, color: store.color, products: aggregateEZProducts(items).slice(0, 10) };
  });

  const result = {
    updatedAt: Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy.MM.dd(EEE) HH:mm'),
    view: 'stores', startDate, endDate, stores
  };

  const jsonStr = JSON.stringify(result);
  try { cache.put(cacheKey, jsonStr, 21600); } catch(_) {}
  return ContentService.createTextOutput(jsonStr).setMimeType(ContentService.MimeType.JSON);
}

// ── 매장별 연간 누적 읽기 (월별 실판매수량·금액 합산) ──────
function readStoreCumulative(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return [];

  const lastCol = sheet.getLastColumn();

  // 헤더에서 월별 앵커(N월실판매수량) 인덱스 수집
  let monthAnchors = [];
  for (let hRow = 1; hRow <= 2; hRow++) {
    if (lastRow < hRow) continue;
    const headers = sheet.getRange(hRow, 1, 1, lastCol).getValues()[0];
    for (let m = 1; m <= 12; m++) {
      const anchor = headers.findIndex(h =>
        String(h || '').includes(m + '월') && String(h || '').includes('실판매수량')
      );
      if (anchor >= 0) monthAnchors.push(anchor);
    }
    if (monthAnchors.length > 0) break;
  }
  if (monthAnchors.length === 0) return [];

  // 마지막 앵커+2 까지만 읽기 (앵커+1 = 월별 금액)
  const maxAnchor = Math.max(...monthAnchors);
  const readCols  = Math.min(maxAnchor + 2, lastCol);
  const rows = sheet.getRange(3, 1, lastRow - 2, readCols).getValues();

  const products = [];
  rows.forEach(row => {
    const name = String(row[COL_NAME] || '').trim();
    if (!name) return;

    let totalQty = 0, totalRev = 0;
    monthAnchors.forEach(anchor => {
      totalQty += Number(row[anchor])     || 0;
      totalRev += Number(row[anchor + 1]) || 0;
    });
    if (totalQty <= 0 && totalRev <= 0) return;

    products.push({
      i: '',
      c: String(row[COL_CAT]   || '').trim() || '기타',
      n: name,
      p: Number(row[COL_PRICE]) || 0,
      q: totalQty,
      s: totalRev
    });
  });

  return products.sort((a, b) => b.s - a.s);
}

// ── 누적 시트 읽기 (26'PP 전체용) ────────────────────────
function readCumulativeSheet(sheet, qtyCol, revCol, maxReadCols) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return [];

  const lastCol = Math.min(sheet.getLastColumn(), maxReadCols + 1);
  const rows = sheet.getRange(3, 1, lastRow - 2, lastCol).getValues();

  const products = [];
  rows.forEach(row => {
    const name = String(row[COL_NAME] || '').trim();
    if (!name) return;

    const qty = Number(row[qtyCol]) || 0;
    const rev = Number(row[revCol]) || 0;
    if (qty <= 0 && rev <= 0) return;

    products.push({
      i: '',
      c: String(row[COL_CAT] || '').trim() || '기타',
      n: name,
      p: Number(row[COL_PRICE]) || 0,
      q: qty,
      s: rev
    });
  });

  return products.sort((a, b) => b.s - a.s);
}

// ── 주차별 시트 읽기 ──────────────────────────────────────
function readStoreWeekly(sheet, period) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return [];

  const lastCol = sheet.getLastColumn();

  // 헤더 행 1, 2 모두 시도 (한국 시트는 2행 헤더인 경우 많음)
  let cols = null;
  for (let hRow = 1; hRow <= 2 && !cols; hRow++) {
    if (lastRow < hRow) break;
    const headers = sheet.getRange(hRow, 1, 1, lastCol).getValues()[0];
    cols = findPeriodColumns(headers, period);
  }
  if (!cols) return [];

  const readCols = Math.max(cols.qty, cols.rev, COL_NAME, COL_PRICE) + 1;
  const rows = sheet.getRange(3, 1, lastRow - 2, Math.min(readCols, lastCol)).getValues();

  const products = [];
  rows.forEach(row => {
    const name = String(row[COL_NAME] || '').trim();
    if (!name) return;

    const qty = Number(row[cols.qty]) || 0;
    const rev = Number(row[cols.rev]) || 0;
    if (qty <= 0 && rev <= 0) return;

    products.push({
      i: '',
      c: String(row[COL_CAT] || '').trim() || '기타',
      n: name,
      p: Number(row[COL_PRICE]) || 0,
      q: qty,
      s: rev
    });
  });

  return products.sort((a, b) => b.s - a.s);
}

// ── 컬럼 탐색 유틸 ────────────────────────────────────────
function findCol(headers, name) {
  return headers.findIndex(h => String(h || '').trim() === name);
}

// period = 'current' | '1월1주' | '2월3주' ...
function findPeriodColumns(headers, period) {
  if (period === 'current') {
    const qty = findCol(headers, '금주판매');
    const rev = findCol(headers, '금주매출');
    if (qty >= 0 && rev >= 0) return { qty, rev };
    return { qty: COL_CUM_QTY, rev: COL_CUM_REV };
  }

  const m = period.match(/^(\d+)월(\d+)주$/);
  if (!m) return null;

  const monthNum = parseInt(m[1]);
  const weekNum  = parseInt(m[2]);

  const anchor = headers.findIndex(h =>
    String(h || '').includes(monthNum + '월') &&
    String(h || '').includes('실판매수량')
  );
  if (anchor < 0) return null;

  const offset = weekNum * 2;
  return { qty: anchor + offset, rev: anchor + offset + 1 };
}

// ── 전체 집계 (주차별 모드용) ─────────────────────────────
function aggregate(allStore, topN) {
  const map = {};
  allStore.forEach(({ products }) => {
    products.forEach(p => {
      if (!map[p.n]) {
        map[p.n] = { i: '', c: p.c, n: p.n, p: p.p, q: 0, s: 0 };
      }
      if (!map[p.n].p && p.p) map[p.n].p = p.p;
      map[p.n].q += p.q;
      map[p.n].s += p.s;
    });
  });
  return Object.values(map).sort((a, b) => b.s - a.s).slice(0, topN);
}

// ════════════════════════════════════════════════════════
// ── 이지어드민 E700 연동 ──────────────────────────────────
// ════════════════════════════════════════════════════════

const EZ_STORES = [
  { name:'남악',  code:89,  color:'#1565c0' },
  { name:'부천',  code:88,  color:'#2e7d32' },
  { name:'의정부',code:120, color:'#e65100' },
  { name:'인천',  code:131, color:'#6a1b9a' },
  { name:'유성',  code:132, color:'#00695c' },
  { name:'여수',  code:133, color:'#ad1457' },
  { name:'율하',  code:70,  color:'#558b2f' },
  { name:'충장',  code:139, color:'#283593' },
  { name:'동수원',code:140, color:'#f57f17' },
  { name:'송파',  code:138, color:'#00838f' },
  { name:'제주',  code:141, color:'#4e342e' },
  { name:'괴정',  code:143, color:'#37474f' },
  { name:'강서',  code:86,  color:'#9e9d24' },
  { name:'순천',  code:149, color:'#bf360c' },
  { name:'평택',  code:150, color:'#4527a0' },
  { name:'고척',  code:152, color:'#01579b' },
  { name:'원주',  code:156, color:'#1b5e20' },
  { name:'마리오',code:159, color:'#ff8f00' },
  { name:'창원',  code:162, color:'#0d47a1' }
];

// ── 인증 쿠키 생성 ──────────────────────────────────────────
function buildEZCookie() {
  const p      = PropertiesService.getScriptProperties();
  const sid    = p.getProperty('PHPSESSID')  || '';
  const domain = p.getProperty('ECN_DOMAIN') || '';
  const id     = p.getProperty('ECN_ID')     || '';
  const pw     = p.getProperty('ECN_PW')     || '';
  return 'PHPSESSID=' + sid +
         '; ecn_domain=' + domain + '; ecn_id=' + id + '; ecn_pw=' + pw +
         '; ecn_saveid=1; ecn_savepw=1';
}


// ── EZAdmin 단일 요청 ────────────────────────────────────────
function ezRequest(storeCode, startDate, endDate, limit) {
  limit = limit || 3000;
  const html = UrlFetchApp.fetch('https://ecn5.ezadmin.co.kr/function.php', {
    method : 'POST',
    headers: { 'Cookie': buildEZCookie(), 'Content-Type': 'application/x-www-form-urlencoded',
               'Referer': 'https://ecn5.ezadmin.co.kr/template.html?template=E700' },
    payload: 'template=E700&action=grid&str_desc_select=' + storeCode +
             '&start_date=' + startDate + '&end_date=' + endDate +
             '&sorting=qty&limit=' + limit + '&category=0&time_check=false',
    followRedirects: true, muteHttpExceptions: true
  }).getContentText();
  return html.includes('세션이 종료') ? null : html;
}

// ── Apps Script용 화요일 계산 ─────────────────────────────
function getThisTuesdayGS(d) {
  const day  = d.getDay();
  const diff = day === 0 ? -5 : day === 1 ? -6 : 2 - day;
  const tue  = new Date(d.getTime());
  tue.setDate(d.getDate() + diff);
  return tue;
}

// ── 사전 계산 & 캐시 워밍 (트리거에서 실행) ──────────────────
// 어제까지의 데이터만 사용 → 하루 1회 새벽 실행으로 충분
function precomputeAndCache() {
  const KST  = 'Asia/Seoul';
  const fmt  = d => Utilities.formatDate(d, KST, 'yyyy-MM-dd');
  const today = new Date();

  // 어제 (기준일)
  const yest = new Date(today); yest.setDate(today.getDate() - 1);

  // 금주: 이번 화요일 ~ 어제
  const thisTue = getThisTuesdayGS(today);

  // 전주: 저번 화요일 ~ 저번 월요일
  const lastMon = new Date(thisTue); lastMon.setDate(thisTue.getDate() - 1);
  const lastTue = new Date(lastMon); lastTue.setDate(lastMon.getDate() - 6);

  // 이번달: 1일 ~ 어제
  const monthStart = new Date(today.getFullYear(), today.getMonth(), 1);

  // 전월: 전월 1일 ~ 전월 말일
  const prevEnd   = new Date(monthStart); prevEnd.setDate(prevEnd.getDate() - 1);
  const prevStart = new Date(prevEnd.getFullYear(), prevEnd.getMonth(), 1);

  const periods = [
    { s: fmt(yest),      e: fmt(yest)      },  // 어제
    { s: fmt(thisTue),   e: fmt(yest)      },  // 금주 (화요일~어제)
    { s: fmt(lastTue),   e: fmt(lastMon)   },  // 전주
    { s: fmt(monthStart),e: fmt(yest)      },  // 이번달 (1일~어제)
    { s: fmt(prevStart), e: fmt(prevEnd)   }   // 전월
  ];

  // 세션 확인 (로드밸런서 대응: 최대 3회 재시도)
  const checkDate = Utilities.formatDate(new Date(new Date().getTime() - 86400000), 'Asia/Seoul', 'yyyy-MM-dd');
  const sessionOk = (() => {
    for (let i = 0; i < 3; i++) {
      const r = UrlFetchApp.fetch('https://ecn5.ezadmin.co.kr/function.php', {
        method : 'POST',
        headers: { 'Cookie': buildEZCookie(), 'Content-Type': 'application/x-www-form-urlencoded',
                   'Referer': 'https://ecn5.ezadmin.co.kr/template.html?template=E700' },
        payload : 'template=E700&action=grid&str_desc_select=89&start_date=' + checkDate +
                  '&end_date=' + checkDate + '&sorting=qty&limit=1&category=0&time_check=false',
        muteHttpExceptions: true
      });
      // 새 PHPSESSID 자동 저장
      const sc = [].concat(r.getAllHeaders()['Set-Cookie'] || []);
      const ns = sc.map(c => c.split(';')[0]).find(c => /^PHPSESSID=.+/.test(c));
      if (ns) PropertiesService.getScriptProperties().setProperty('PHPSESSID', ns.split('=')[1]);
      if (!r.getContentText().includes('세션이 종료')) return true;
    }
    return false;
  })();

  if (!sessionOk) {
    Logger.log('⚠️ 세션 만료 — PHPSESSID를 수동으로 갱신해 주세요.');
    return;
  }
  Logger.log('✅ 세션 확인 완료');

  // 각 기간 사전 계산: 19개 요청 1회로 전체+매장별 동시 캐싱
  periods.forEach(({ s, e }) => {
    try {
      const cnt = computeAndCachePeriod(s, e);
      Logger.log('✅ 캐시 완료 (전체+매장별): ' + s + ' ~ ' + e + ' / 전체 상품 수: ' + cnt);
    } catch(err) {
      Logger.log('❌ 오류: ' + s + ' ~ ' + e + ' — ' + err.message);
    }
  });

  Logger.log('🎉 사전 계산 완료 — ' + Utilities.formatDate(today, KST, 'HH:mm'));
}

// ── 세션 유지 ping (별도 30분 트리거용) ─────────────────────
function keepSessionAlive() {
  const ymd  = Utilities.formatDate(new Date(new Date().getTime() - 86400000), 'Asia/Seoul', 'yyyy-MM-dd');
  const html = UrlFetchApp.fetch('https://ecn5.ezadmin.co.kr/function.php', {
    method : 'POST',
    headers: { 'Cookie': buildEZCookie(), 'Content-Type': 'application/x-www-form-urlencoded',
               'Referer': 'https://ecn5.ezadmin.co.kr/template.html?template=E700' },
    payload : 'template=E700&action=grid&str_desc_select=89&start_date=' + ymd + '&end_date=' + ymd + '&sorting=qty&limit=1&category=0&time_check=false',
    muteHttpExceptions: true
  }).getContentText();

  if (html.includes('세션이 종료')) {
    Logger.log('⚠️ 세션 만료 — PHPSESSID 수동 갱신 필요');
  } else {
    Logger.log('✅ 세션 정상: ' + Utilities.formatDate(new Date(), 'Asia/Seoul', 'HH:mm'));
  }
}

// ── 매장 1개 데이터 가져오기 (null = 세션 만료) ──
function fetchEZStore(storeCode, startDate, endDate) {
  try {
    const resp = UrlFetchApp.fetch('https://ecn5.ezadmin.co.kr/function.php', {
      method: 'POST',
      headers: {
        'Cookie'      : buildEZCookie(),
        'Content-Type': 'application/x-www-form-urlencoded',
        'Referer'     : 'https://ecn5.ezadmin.co.kr/template.html?template=E700'
      },
      payload: 'template=E700&action=grid&str_desc_select=' + storeCode +
               '&start_date=' + startDate + '&end_date=' + endDate +
               '&sorting=qty&limit=3000&category=0&time_check=false',
      followRedirects: true,
      muteHttpExceptions: true
    });
    const html = resp.getContentText();
    if (html.includes('세션이 종료')) return null;
    return parseEZResponse(html);
  } catch(e) {
    return [];
  }
}

// ── HTML 응답에서 제품 배열 파싱 ──
function parseEZResponse(html) {
  const m = html.match(/'addRowData',\s*"product_id",\s*(\[\{[\s\S]*?\}\])\s*\)/);
  if (!m) return [];
  try { return JSON.parse(m[1]); } catch(e) { return []; }
}

// ── product_image HTML에서 URL 추출 ──
function extractEZImageUrl(imgHtml) {
  if (!imgHtml) return '';
  const m = imgHtml.match(/\/uploads\/dammom\/(https?:\/\/[^'"\s]+)/);
  return m ? m[1] : '';
}

// ── SKU 배열 → 상품명 기준 집계 ──
function aggregateEZProducts(items) {
  const map = {};
  items.forEach(item => {
    const qty = parseInt(item.real_qty)    || 0;
    const rev = parseInt(item.real_amount) || 0;
    if (qty <= 0 && rev <= 0) return;
    const name = String(item.product_name || '').trim();
    if (!name) return;
    const cat = String(item.category || '').split('>')[0].trim() || '기타';
    if (!map[name]) {
      map[name] = {
        i: extractEZImageUrl(item.product_image || ''),
        c: cat,
        n: name,
        p: parseInt(item.price) || 0,
        q: 0, s: 0
      };
    }
    map[name].q += qty;
    map[name].s += rev;
  });
  return Object.values(map).sort((a, b) => b.s - a.s);
}

// ── 로그인 → PHPSESSID 획득 ──────────────────────────────
function ezLogin() {
  const p      = PropertiesService.getScriptProperties();
  const domain = p.getProperty('ECN_DOMAIN') || '';
  const id     = p.getProperty('ECN_ID')     || '';
  const pw     = p.getProperty('ECN_PW')     || '';

  // 1단계: 로그인 페이지 요청 (PHPSESSID 쿠키 초기값 획득)
  const init = UrlFetchApp.fetch('https://ecn5.ezadmin.co.kr/', {
    followRedirects: false,
    muteHttpExceptions: true
  });
  const initCookies = (init.getAllHeaders()['Set-Cookie'] || []);
  const initSession = [].concat(initCookies).map(c => c.split(';')[0]).join('; ');

  // 2단계: 로그인 POST
  const loginResp = UrlFetchApp.fetch('https://ecn5.ezadmin.co.kr/function.php', {
    method: 'POST',
    headers: {
      'Cookie'      : initSession,
      'Content-Type': 'application/x-www-form-urlencoded'
    },
    payload: 'template=main&action=login&domain=' + domain +
             '&id=' + id + '&pw=' + encodeURIComponent(pw),
    followRedirects: false,
    muteHttpExceptions: true
  });

  Logger.log('로그인 응답 코드: ' + loginResp.getResponseCode());
  Logger.log('로그인 응답 Body: ' + loginResp.getContentText().substring(0, 300));

  const setCookies = loginResp.getAllHeaders()['Set-Cookie'] || [];
  const session = [].concat(setCookies).map(c => c.split(';')[0]).join('; ');
  Logger.log('받은 쿠키: ' + session);
  return session || initSession;
}

// ── 인증 테스트 (에디터에서 직접 실행) ──────────────────────
function testEZAdminAuth() {
  const html = UrlFetchApp.fetch('https://ecn5.ezadmin.co.kr/function.php', {
    method: 'POST',
    headers: {
      'Cookie'      : buildEZCookie(),
      'Content-Type': 'application/x-www-form-urlencoded'
    },
    payload: 'template=E700&action=grid&str_desc_select=89' +
             '&start_date=2026-05-01&end_date=2026-05-31' +
             '&sorting=qty&limit=5&category=0&time_check=false',
    muteHttpExceptions: true
  }).getContentText();

  const items = parseEZResponse(html);
  Logger.log('파싱된 상품 수: ' + items.length);
  if (items.length > 0) {
    Logger.log('첫 번째 상품: ' + items[0].product_name + ' / 실판매수량: ' + items[0].real_qty);
    Logger.log('✅ 인증 성공!');
  } else {
    Logger.log('❌ 인증 실패 — PHPSESSID 확인 필요');
    Logger.log('HTML 앞부분: ' + html.substring(0, 200));
  }
}

// ── 주차 목록 생성 ────────────────────────────────────────
function buildPeriodList(ss) {
  for (const storeName of STORE_NAMES) {
    const sheet = ss.getSheetByName(storeName);
    if (!sheet) continue;

    const lastCol = sheet.getLastColumn();
    if (lastCol < 14) continue;

    const lastRow = sheet.getLastRow();

    // 헤더 행 1, 2 모두 시도
    for (let hRow = 1; hRow <= 2; hRow++) {
      if (lastRow < hRow) continue;
      const headers = sheet.getRange(hRow, 1, 1, lastCol).getValues()[0];
      const periods = [];

      for (let m = 1; m <= 12; m++) {
        const anchor = headers.findIndex(h =>
          String(h || '').includes(m + '월') && String(h || '').includes('실판매수량')
        );
        if (anchor < 0) continue;

        for (let w = 1; w <= 6; w++) {
          const qIdx = anchor + w * 2;
          if (qIdx >= lastCol) break;
          if (String(headers[qIdx] || '').includes('실판매수량')) break;
          periods.push({ label: m + '월 ' + w + '주차', value: m + '월' + w + '주' });
        }
      }

      if (periods.length > 0) return periods;
    }
  }
  return [];
}

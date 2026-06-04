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


// ── 진입점 (이지어드민 API) ───────────────────────────────
function doGet(e) {
  try {
    const p       = (e && e.parameter) ? e.parameter : {};
    const noCache = p.nc === '1';
    const today   = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');
    const startDate = p.startDate || today;
    const endDate   = p.endDate   || today;

    // 캐시 조회 (10분)
    const cacheKey = 'ez_v2_' + startDate + '_' + endDate;
    const cache    = CacheService.getScriptCache();
    if (!noCache) {
      const cached = cache.get(cacheKey);
      if (cached) {
        return ContentService.createTextOutput(cached).setMimeType(ContentService.MimeType.JSON);
      }
    }

    // 첫 매장으로 세션 유효성 확인
    const firstItems = fetchEZStore(EZ_STORES[0].code, startDate, endDate);
    if (firstItems === null) {
      return ContentService
        .createTextOutput(JSON.stringify({ error: 'session_expired', message: 'PHPSESSID가 만료됐습니다. Apps Script 스크립트 속성에서 PHPSESSID를 갱신해 주세요.' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 전 매장 데이터 수집
    const allStore = [];
    allStore.push({
      name: EZ_STORES[0].name, color: EZ_STORES[0].color,
      products: aggregateEZProducts(firstItems)
    });
    EZ_STORES.slice(1).forEach(store => {
      try {
        const items = fetchEZStore(store.code, startDate, endDate);
        allStore.push({ name: store.name, color: store.color, products: aggregateEZProducts(items || []) });
      } catch(err) {
        allStore.push({ name: store.name, color: store.color, products: [] });
      }
    });

    // 전체 TOP 20 집계
    const map = {};
    allStore.forEach(({ products }) => {
      products.forEach(pr => {
        if (!map[pr.n]) map[pr.n] = { i: pr.i, c: pr.c, n: pr.n, p: pr.p, q: 0, s: 0 };
        map[pr.n].q += pr.q;
        map[pr.n].s += pr.s;
        if (!map[pr.n].i && pr.i) map[pr.n].i = pr.i;
      });
    });
    const overall = Object.values(map).sort((a, b) => b.s - a.s).slice(0, 20);

    const result = {
      updatedAt : Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy.MM.dd(EEE) HH:mm'),
      startDate, endDate, overall,
      stores: allStore.map(s => ({ ...s, products: s.products.slice(0, 10) }))
    };

    const jsonStr = JSON.stringify(result);
    try { cache.put(cacheKey, jsonStr, 600); } catch(_) {}

    return ContentService.createTextOutput(jsonStr).setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
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

// ── 인증 쿠키 생성 (PHPSESSID + remember-me) ──
function buildEZCookie() {
  const p      = PropertiesService.getScriptProperties();
  const sid    = p.getProperty('PHPSESSID')  || '';
  const domain = p.getProperty('ECN_DOMAIN') || '';
  const id     = p.getProperty('ECN_ID')     || '';
  const pw     = p.getProperty('ECN_PW')     || '';
  return 'PHPSESSID=' + sid +
         '; ecn_domain=' + domain +
         '; ecn_id=' + id +
         '; ecn_pw=' + pw +
         '; ecn_saveid=1; ecn_savepw=1';
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

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
/**
 * 첫 번째 매장 시트의 1~2행 헤더 전체를 출력합니다.
 */
function debugHeaders() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(STORE_NAMES[0]);
  if (!sheet) { Logger.log('시트 없음'); return; }

  const lastCol = sheet.getLastColumn();
  Logger.log('총 열: ' + lastCol);

  for (let hRow = 1; hRow <= 2; hRow++) {
    const headers = sheet.getRange(hRow, 1, 1, lastCol).getValues()[0];
    // 비어있지 않은 헤더만 출력
    headers.forEach((h, i) => {
      if (String(h || '').trim()) {
        Logger.log('행' + hRow + ' [' + i + '] = "' + String(h).replace(/\n/g, '\\n') + '"');
      }
    });
  }
}

// ── 진입점 ────────────────────────────────────────────────
function doGet(e) {
  try {
    const p      = (e && e.parameter) ? e.parameter : {};
    const mode   = p.mode   || 'weekly';
    const period = p.period || 'current';

    const ss     = SpreadsheetApp.getActiveSpreadsheet();
    const imgMap = loadImageMap(ss); // 상품리스트에서 이미지 URL 로드

    let overall = [], stores = [];

    if (mode === 'cumulative') {
      // ── 26년 누적 모드 ──
      const ppSheet = ss.getSheetByName(OVERALL_SHEET);
      overall = ppSheet ? readCumulativeSheet(ppSheet, COL_OV_QTY, COL_OV_REV, 22) : [];

      STORE_NAMES.forEach(name => {
        try {
          const sheet = ss.getSheetByName(name);
          if (!sheet) return;
          const products = readCumulativeSheet(sheet, COL_CUM_QTY, COL_CUM_REV, 12);
          stores.push({ name, color: STORE_COLORS[name] || '#1565c0', products: products.slice(0, 10) });
        } catch (err) {
          stores.push({ name, color: STORE_COLORS[name] || '#1565c0', products: [] });
        }
      });

    } else {
      // ── 주차별 모드 ──
      const allStore = [];
      STORE_NAMES.forEach(name => {
        try {
          const sheet = ss.getSheetByName(name);
          if (!sheet) return;
          const products = readStoreWeekly(sheet, period);
          allStore.push({ name, color: STORE_COLORS[name] || '#1565c0', products });
        } catch (err) {
          allStore.push({ name, color: STORE_COLORS[name] || '#1565c0', products: [] });
        }
      });
      overall = aggregate(allStore, 20);
      stores  = allStore.map(s => ({ name: s.name, color: s.color, products: s.products.slice(0, 10) }));
    }

    // ── 이미지 URL 적용 (상품리스트 매핑) ──
    const applyImg = prod => { prod.i = imgMap[prod.n] || ''; return prod; };
    overall = overall.map(applyImg);
    stores  = stores.map(s => ({ ...s, products: s.products.map(applyImg) }));

    const periods = buildPeriodList(ss);

    const result = {
      updatedAt : Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy.MM.dd(EEE) HH:mm'),
      mode, period, periods, overall, stores
    };

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── 누적 시트 읽기 ────────────────────────────────────────
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

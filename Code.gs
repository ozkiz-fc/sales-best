/* ===========================================================
   판매 베스트 대시보드 — Apps Script API
   -----------------------------------------------------------
   시트 컬럼 구조 (0-based):
   A(0)=이미지  B(1)=대분류  C(2)=소분류  D(3)=년도
   E(4)=소분류&시즌  F(5)=시즌  G(6)=상품명  H(7)=원가
   I(8)=판매가  J(9)=대표바코드
   K(10)=26년누적판매수량  L(11)=26년누적판매금액   ← 각 매장 누적
   M(12)=1월낮판매수량  N(13)=1월낮판매금액
   O(14)=W1수량  P(15)=W1금액
   Q(16)=2주수량  R(17)=2주매출  ...

   26'PP 시트:
   U(20)=전체누적판매수량  V(21)=전체누적판매금액
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

const OVERALL_SHEET = "26'PP"; // 전체 누적 시트명

// ── 진입점 ────────────────────────────────────────────────
function doGet(e) {
  try {
    const p      = (e && e.parameter) ? e.parameter : {};
    const mode   = p.mode   || 'weekly';
    const period = p.period || 'current';

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let overall = [], stores = [];

    if (mode === 'cumulative') {
      // ── 26년 누적 모드 ──
      // 전체: 26'PP 시트 U열/V열
      const ppSheet = ss.getSheetByName(OVERALL_SHEET);
      overall = ppSheet ? readCumulativeSheet(ppSheet, COL_OV_QTY, COL_OV_REV, 20) : [];

      // 매장별: 각 매장 시트 K열/L열 (12컬럼만 읽어 속도 개선)
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
// qtyCol, revCol: 0-based 컬럼 인덱스
// maxReadCols: 읽을 최대 컬럼 수 (속도 최적화)
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
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  const cols = findPeriodColumns(headers, period);
  if (!cols) return [];

  // 필요한 컬럼까지만 읽어 속도 개선
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
    // 금주 컬럼 없으면 K/L(누적) 사용
    return { qty: COL_CUM_QTY, rev: COL_CUM_REV };
  }

  const m = period.match(/^(\d+)월(\d+)주$/);
  if (!m) return null;

  const monthNum = parseInt(m[1]);
  const weekNum  = parseInt(m[2]);

  const anchor = headers.findIndex(h =>
    String(h || '').includes(monthNum + '월') &&
    String(h || '').includes('낮판매수량')
  );
  if (anchor < 0) return null;

  // anchor+2 = W1수량/금액, anchor+4 = 2주수량/금액 ...
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
  // 여러 매장 시트에서 "N월 낮판매수량" 헤더를 찾아 주차 목록 생성
  for (const storeName of STORE_NAMES) {
    const sheet = ss.getSheetByName(storeName);
    if (!sheet) continue;

    const lastCol = sheet.getLastColumn();
    if (lastCol < 14) continue; // 주차 컬럼이 없으면 건너뜀

    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const periods = [];

    for (let m = 1; m <= 12; m++) {
      const anchor = headers.findIndex(h =>
        String(h || '').includes(m + '월') && String(h || '').includes('낮판매수량')
      );
      if (anchor < 0) continue;

      for (let w = 1; w <= 6; w++) {
        const qIdx = anchor + w * 2;
        if (qIdx >= lastCol) break;
        // 다음 달 앵커면 중단
        if (String(headers[qIdx] || '').includes('낮판매수량')) break;
        periods.push({ label: m + '월 ' + w + '주차', value: m + '월' + w + '주' });
      }
    }

    if (periods.length > 0) return periods;
  }
  return [];
}

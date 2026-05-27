/* ===========================================================
   판매 베스트 대시보드 — Apps Script API
   -----------------------------------------------------------
   배포: 웹 앱으로 배포 > 액세스 권한 "모든 사람"

   시트 컬럼 구조 (0-based index):
   A(0)=이미지  B(1)=대분류  C(2)=소분류  D(3)=년도
   E(4)=소분류&시즌  F(5)=시즌  G(6)=상품명  H(7)=원가
   I(8)=판매가  J(9)=대표바코드  K(10)=금주판매  L(11)=금주매출
   M(12)=1월낮판매수량  N(13)=1월낮판매금액
   O(14)=W1수량  P(15)=W1금액
   Q(16)=2주수량  R(17)=2주매출 ...
   Y(24)=2월낮판매수량  Z(25)=2월낮판매금액 ...
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

// 시트 고정 컬럼 인덱스
const COL_CAT   = 1;  // B: 대분류 (의류/레인/슈즈/잡화/시즌)
const COL_NAME  = 6;  // G: 상품명
const COL_PRICE = 8;  // I: 판매가

// ── 진입점 ────────────────────────────────────────────────
function doGet(e) {
  try {
    const p      = (e && e.parameter) ? e.parameter : {};
    const mode   = p.mode   || 'weekly';
    const period = p.period || 'current';

    const ss       = SpreadsheetApp.getActiveSpreadsheet();
    const allStore = [];

    STORE_NAMES.forEach(name => {
      try {
        const sheet = ss.getSheetByName(name);
        if (!sheet) return;
        const products = readStore(sheet, mode, period);
        allStore.push({ name, color: STORE_COLORS[name] || '#1565c0', products });
      } catch (err) {
        allStore.push({ name, color: STORE_COLORS[name] || '#1565c0', products: [] });
      }
    });

    const overall = aggregate(allStore, 20);
    const stores  = allStore.map(s => ({ name: s.name, color: s.color, products: s.products.slice(0, 10) }));
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

// ── 매장 시트 읽기 ─────────────────────────────────────────
function readStore(sheet, mode, period) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) return [];

  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const rows    = sheet.getRange(3, 1, lastRow - 2, lastCol).getValues();

  let qtyIdx = -1, revIdx = -1;
  let cumCols = [];

  if (mode === 'cumulative') {
    cumCols = findCumulativeColumns(headers);
    if (cumCols.length === 0) {
      qtyIdx = findCol(headers, '금주판매');
      revIdx = findCol(headers, '금주매출');
    }
  } else {
    const cols = findPeriodColumns(headers, period);
    if (!cols) return [];
    qtyIdx = cols.qty;
    revIdx = cols.rev;
  }

  const products = [];
  rows.forEach(row => {
    // G열(index 6) = 상품명
    const name = String(row[COL_NAME] || '').trim();
    if (!name) return;

    let qty = 0, rev = 0;

    if (mode === 'cumulative' && cumCols.length > 0) {
      cumCols.forEach(c => {
        qty += Number(row[c.qi]) || 0;
        rev += Number(row[c.ri]) || 0;
      });
    } else {
      if (qtyIdx < 0 || revIdx < 0) return;
      qty = Number(row[qtyIdx]) || 0;
      rev = Number(row[revIdx]) || 0;
    }

    if (qty <= 0 && rev <= 0) return;

    // B열(index 1) = 대분류 (의류/레인/슈즈/잡화/시즌)
    const cat = String(row[COL_CAT] || '').trim() || '기타';

    products.push({
      i: '',                         // CellImage는 URL로 읽을 수 없음
      c: cat,
      n: name,
      p: Number(row[COL_PRICE]) || 0, // I열(index 8) = 판매가
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
    return (qty >= 0 && rev >= 0) ? { qty, rev } : null;
  }

  const m = period.match(/^(\d+)월(\d+)주$/);
  if (!m) return null;

  const monthNum = parseInt(m[1]);
  const weekNum  = parseInt(m[2]);

  // "N월 낮판매수량" 컬럼 위치(anchor) 탐색
  const anchor = headers.findIndex(h =>
    String(h || '').includes(monthNum + '월') &&
    String(h || '').includes('낮판매수량')
  );
  if (anchor < 0) return null;

  // anchor+2 = W1수량, anchor+3 = W1금액
  // anchor+4 = 2주수량, anchor+5 = 2주매출 ...
  const offset = weekNum * 2;
  return { qty: anchor + offset, rev: anchor + offset + 1 };
}

// 누적: 모든 "N월 낮판매수량/금액" 컬럼 수집
function findCumulativeColumns(headers) {
  const cols = [];
  for (let m = 1; m <= 12; m++) {
    const qi = headers.findIndex(h =>
      String(h || '').includes(m + '월') && String(h || '').includes('낮판매수량')
    );
    const ri = headers.findIndex(h =>
      String(h || '').includes(m + '월') && String(h || '').includes('낮판매금액')
    );
    if (qi >= 0 && ri >= 0) cols.push({ qi, ri });
  }
  return cols;
}

// ── 전체 집계 ─────────────────────────────────────────────
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

// ── 사용 가능 주차 목록 생성 ──────────────────────────────
// 헤더에서 "N월 낮판매수량" 컬럼을 찾아 주차 목록 생성
// 여러 매장 시트를 순서대로 검색하여 첫 번째로 데이터가 있는 시트 사용
function buildPeriodList(ss) {
  for (const storeName of STORE_NAMES) {
    const sheet = ss.getSheetByName(storeName);
    if (!sheet) continue;

    const lastCol = sheet.getLastColumn();
    if (lastCol < 1) continue;

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
        // 다음 달 앵커 컬럼이면 중단
        const colHeader = String(headers[qIdx] || '');
        if (colHeader.includes('낮판매수량')) break;
        periods.push({ label: m + '월 ' + w + '주차', value: m + '월' + w + '주' });
      }
    }

    if (periods.length > 0) return periods;
  }
  return [];
}

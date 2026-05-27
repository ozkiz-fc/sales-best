/* ===========================================================
   판매 베스트 대시보드 — Apps Script API
   -----------------------------------------------------------
   배포: 웹 앱으로 배포 > 액세스 권한 "모든 사람"
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

// 상품명 접두어 → 분류
const CAT_MAP = {
  '장화':'레인','우비':'레인','레인':'레인',
  '샌들':'슈즈','아쿠아':'슈즈','젤리':'슈즈','슬립온':'슈즈','구두':'슈즈',
  '원피스':'의류','세트':'의류','하의':'의류','상의':'의류',
  '아우터':'의류','수영복':'의류','래쉬가드':'의류'
};

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
        // 매장 시트 오류 시 빈 배열로 처리
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
      // 누적 컬럼이 없으면 금주 사용
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
    const name = String(row[1] || '').trim();
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

    const prefix = name.split('-')[0];
    products.push({
      i: String(row[0] || ''),
      c: CAT_MAP[prefix] || '기타',
      n: name,
      p: Number(row[2]) || 0,
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

  // anchor+0 = 월수량, anchor+1 = 월금액
  // anchor+2 = W1수량, anchor+3 = W1금액
  // anchor+4 = W2수량, anchor+5 = W2금액 ...
  const offset = weekNum * 2;
  return { qty: anchor + offset, rev: anchor + offset + 1 };
}

// 26년 누적: 모든 "N월 낮판매수량/금액" 컬럼 수집
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
        map[p.n] = { i: p.i || '', c: p.c, n: p.n, p: p.p, q: 0, s: 0 };
      }
      // 이미지/가격은 첫 번째 non-empty 값 유지
      if (!map[p.n].i && p.i) map[p.n].i = p.i;
      map[p.n].q += p.q;
      map[p.n].s += p.s;
    });
  });
  return Object.values(map).sort((a, b) => b.s - a.s).slice(0, topN);
}

// ── 사용 가능 주차 목록 생성 ──────────────────────────────
function buildPeriodList(ss) {
  const sheet = ss.getSheetByName(STORE_NAMES[0]);
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  // 2행(합계행) 먼저 시도, 합계가 모두 0이면 3행~ 데이터를 열 단위로 합산
  let checkRow = sheet.getRange(2, 1, 1, lastCol).getValues()[0];
  const rowSum = checkRow.reduce((a, b) => a + (Number(b) || 0), 0);
  if (rowSum === 0 && lastRow >= 3) {
    const dataRows = sheet.getRange(3, 1, lastRow - 2, lastCol).getValues();
    checkRow = Array(lastCol).fill(0).map((_, ci) =>
      dataRows.reduce((s, r) => s + (Number(r[ci]) || 0), 0)
    );
  }

  const periods = [];
  for (let m = 1; m <= 12; m++) {
    const anchor = headers.findIndex(h =>
      String(h || '').includes(m + '월') && String(h || '').includes('낮판매수량')
    );
    if (anchor < 0) continue;

    // 최대 6주까지 체크
    for (let w = 1; w <= 6; w++) {
      const qIdx = anchor + w * 2;
      if (qIdx >= checkRow.length) break;
      const total = Number(checkRow[qIdx]) || 0;
      if (total > 0) {
        periods.push({ label: m + '월 ' + w + '주차', value: m + '월' + w + '주' });
      }
    }
  }
  return periods;
}

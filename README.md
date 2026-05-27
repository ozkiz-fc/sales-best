# 🏆 판매 베스트 대시보드

오즈키즈 전체 매장 · 18개 지점의 주간 판매 현황을 보여주는 단일 파일 대시보드입니다.

## 미리보기

- **전체 판매 베스트 TOP 20** — 전 매장 합산 주간 판매량/매출 순위
- **매장별 판매 베스트 TOP 10** — 18개 지점별 개별 순위 카드

## 시작하기

### 1. Apps Script 준비

Google Sheets에 연결된 Apps Script 웹 앱이 다음 JSON 형식을 반환해야 합니다:

```json
{
  "updated": "2026.05.26(Tue) 20:27",
  "overall": [
    {
      "name": "장화-팅클텝LED",
      "category": "레인",
      "img": "https://...",
      "price": 29900,
      "qty": 183,
      "revenue": 5471700
    }
  ],
  "stores": [
    {
      "name": "남악",
      "color": "#1565c0",
      "items": [
        {
          "name": "세트-슈가체크",
          "img": "https://...",
          "qty": 10,
          "revenue": 499000
        }
      ]
    }
  ]
}
```

### 2. API_URL 설정

`index.html` 상단의 `API_URL` 값을 배포된 Apps Script URL로 교체합니다:

```js
const API_URL = 'https://script.google.com/macros/s/YOUR_SCRIPT_ID/exec';
```

### 3. 배포

파일 하나(`index.html`)만 있으면 됩니다. GitHub Pages, Netlify, 또는 로컬에서 바로 열 수 있습니다.

**GitHub Pages 배포:**
1. 이 저장소를 GitHub에 push
2. Settings → Pages → Branch: `main`, 폴더: `/ (root)` 선택
3. 저장하면 `https://<username>.github.io/<repo>/` 에서 접근 가능

## 파일 구조

```
.
└── index.html   # 대시보드 (HTML + CSS + JS 단일 파일)
```

## 커스터마이징

| 항목 | 위치 |
|------|------|
| 매장 헤더 색상 | `STORE_COLORS` 배열 |
| 표시 건수 (TOP N) | `renderOverall` / `renderStores` 호출부 |
| 색상 테마 | `:root` CSS 변수 |

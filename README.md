# BLUE NINE — Prototype v1.0

사내 AE를 위한 범용 광고 견적서 효율화 웹 솔루션의 프로토타입입니다.
설계 근거는 `Rule Book_260518.docx` (Source 1 ~ Source 37)이며, 모든 코드 주석에 Source 번호로 트레이서빌리티를 남겼습니다.

> 🚀 **배포 안내**: 무료 배포(Vercel + Render) 절차는 **[DEPLOY.md](./DEPLOY.md)** 참고.

## 폴더 구조

```
AI TF - pioneer/
├── backend/                         # FastAPI 백엔드 (완전 휘발성)
│   ├── main.py                      # API 엔트리포인트 + 라우팅
│   ├── schemas.py                   # Pydantic 데이터 모델
│   ├── session_store.py             # In-memory 세션 (Source 9, 10)
│   ├── monitoring.py                # Latency/Traffic/Token-Cost 미들웨어
│   ├── master_loader.py             # 제작단가기준집.pdf 로더 (Source 16, 36)
│   ├── validation.py                # 3단계 신호등 + 더블체크 (Source 28~33)
│   ├── exporter.py                  # Excel(살아있는 수식) Export (Source 25b)
│   ├── parsers/                     # 카테고리별 파서 (2단계 분기)
│   │   ├── __init__.py              #   라우터: (l1, l2) → 파서 클래스
│   │   ├── base.py                  #   공통 휴리스틱
│   │   ├── production_video.py      #   영상제작비
│   │   ├── production_radio.py      #   라디오제작비
│   │   ├── production_print.py      #   인쇄제작비
│   │   ├── production_btl.py        #   BTL제작비
│   │   ├── media_tvc.py             #   TVC 매체비
│   │   ├── media_radio.py           #   라디오 매체비
│   │   ├── media_print.py           #   PRINT 매체비
│   │   ├── media_digital.py         #   디지털 매체비
│   │   └── billing_status.py        #   Billing Status (삼각 검증용)
│   └── run.py                       # 개발용 uvicorn 런처
│
├── frontend/                        # React (Vite) 프론트엔드
│   ├── package.json
│   ├── vite.config.js               # /api → 127.0.0.1:8088 프록시
│   ├── index.html
│   └── src/
│       ├── main.jsx
│       ├── App.jsx                  # 상단 토글/배너/3컬럼 워크스페이스
│       ├── api.js                   # fetch 클라이언트 (Mode 헤더 포함)
│       ├── styles.css               # 일반 UI
│       ├── print.css                # @page A4 + @media print (Source 25a)
│       └── components/
│           ├── ModeToggle.jsx       # Fast / Precise (Source 3, 4)
│           ├── CategoryStepper.jsx  # 2단계 카테고리 분기 (Source 11)
│           ├── UploadForm.jsx       # 협력사 파일 업로드
│           ├── EstimateSheet.jsx    # 견적서 + 신호등 + 인라인 수정
│           ├── ExportBar.jsx        # 인쇄 / Excel 다운로드
│           └── MonitorPanel.jsx     # 운영 모니터링 (요구사항 4)
│
├── reference/                       # 학습 자산 (Source 26) — 휘발성 입력과 별개
│   ├── 제작단가기준집.pdf            # 마스터 데이터 (Source 16)
│   ├── 영상제작비견적서1·2.xlsx
│   ├── 라디오제작비견적서1.xlsx
│   ├── 인쇄제작비견적서1·2.xlsx
│   ├── BTL제작비견적서1·2·3.xlsx
│   └── Billing Status.xls
│
├── Rule Book_260518.docx            # 단일 진실 공급원 (SSOT)
└── requirements.txt
```

## 실행 방법

### 1. 백엔드 (FastAPI)
```powershell
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python -m backend.run
# → http://127.0.0.1:8088
# Swagger: http://127.0.0.1:8088/docs
```

### 2. 프론트엔드 (React + Vite)
```powershell
cd frontend
npm install
npm run dev
# → http://localhost:5173 (자동 /api 프록시)
```

크롬에서 http://localhost:5173 으로 접속.

## 핵심 요구사항 매핑

| 요구사항 | 구현 위치 | Rule Book Source |
|---|---|---|
| 2단계 카테고리 선택 UI | `CategoryStepper.jsx`, `GET /api/categories` | Source 11, 12, 17 |
| 동적 분기 → 템플릿 파서 매칭 | `backend/parsers/__init__.py:get_parser` | Source 26 |
| 완전 휘발성 (DB 없음) | `session_store.py` (RAM only), 30분 idle GC | Source 9, 10 |
| 마스터 데이터 메모리 적재 | `master_loader.MASTER.load()` (startup) | Source 16, 36 |
| Fast / Precise 토글 | `ModeToggle.jsx` + `X-Blue-Nine-Mode` header | Source 3, 4 |
| 3단계 신호등 (행 + 합계) | `validation.evaluate_document` | Source 29~33 |
| 빨강/노랑 인라인 수정 | `EstimateSheet.jsx > EditableNum` | Source 32 |
| Chrome A4 인쇄 | `print.css @page / @media print` | Source 25a |
| Excel(살아있는 수식) | `exporter.build_xlsx` (=D*E, =SUM, =ROUND) | Source 25b |
| 더블체크 (입출력 100% 일치) | `validation._close` + `evaluate_document` warnings | Source 28, 29 |
| 매체비 삼각 검증 | `validation.evaluate_triangle` + `parsers/billing_status.py` | Source 19, 20, 22 |
| 모니터링(Latency/Traffic/Cost) | `monitoring.BillingMiddleware`, `/api/monitor/*` | (요구사항 4) |
| Red 상태 Export 차단 | `main.estimate_export` 409 | Source 32 |
| 'BLUE NINE' 명칭 일관 노출 | 탭 타이틀, 상단 로고, 파일명 prefix, Excel 헤더/푸터 | Source 37 |

## API 요약

| Method | Path | 설명 |
|---|---|---|
| POST | `/api/session/new` | 신규 휘발성 세션 발급 |
| POST | `/api/session/destroy` | 세션 즉시 파기 |
| GET  | `/api/categories` | 2단계 카테고리 메타 |
| GET  | `/api/master/info` | 제작단가기준집 적재 상태 |
| GET  | `/api/master/items` | 마스터 항목 (read-only) |
| POST | `/api/master/reload` | (admin) 마스터 재적재 — `admin_key` 헤더 필요 |
| POST | `/api/estimate/parse` | 협력사 견적서 파싱 + 신호등 평가 |
| POST | `/api/estimate/update_row` | 행 수동 수정 후 재평가 |
| POST | `/api/estimate/{sid}/{eid}/export.xlsx` | Excel(수식 살아있음) |
| GET  | `/api/monitor/summary` | 호출수/Latency/토큰/비용 집계 |
| GET  | `/api/monitor/logs` | 최근 호출 로그 |

## 데이터 흐름

```
[브라우저] ── multipart 파일 + 카테고리(l1, l2) ──▶ [FastAPI]
                                                    │
                                                    ▼
                            ┌────────────────────────────────┐
                            │ parsers.get_parser(l1, l2)     │  ← 동적 매칭
                            ├────────────────────────────────┤
                            │ rows = parser.parse(blob)      │
                            │ validation.evaluate_document() │  ← 신호등
                            │ (media: + billing 삼각 검증)   │
                            ├────────────────────────────────┤
                            │ STORE.put_estimate (RAM only)  │  ← DB 없음
                            └────────────────────────────────┘
                                          │
                                          ▼
[브라우저] ◀── EstimateDocument JSON ── [FastAPI]
   │
   ├── 인라인 수정 → POST /api/estimate/update_row → 재평가 → 갱신
   ├── 인쇄  → window.print()  (print.css A4)
   └── Excel → POST /export.xlsx → openpyxl 수식 워크북

[brwsr close] ─ beforeunload ─▶ /api/session/destroy ─▶ STORE.drop()
```

## 더블체크 로직 (요구사항 & Source 28, 29)

`validation.evaluate_document` 내부:
- 행 단위: `unit_price × quantity ≈ amount`  (허용오차: `max(0.5원, 1e-4 × |max|)`)
- 시트 합계: `Σ(row.amount) + (자동계산된 대행수수료) == sum_jeongga + sum_outsourcing + sum_agency_fee`
- 매체비: 삼각 (광고주 청구 — Billing — 매체사 지급 + 수수료) 의 max Δ 가 허용오차 이내
- 어느 한 곳이라도 어긋나면 `doc.warnings` 에 누적되고 overall 신호등 ▶ Red.

## 운영 모드별 비용 모델 (모니터링)

`monitoring.py` 상수:
- `PRICE_IN_PER_MTOK_USD = 15.0`  / `PRICE_OUT_PER_MTOK_USD = 75.0`  (가정치)
- `BYTES_PER_TOKEN = 4`           (한국어 안전 추정)
- `FAST_DISCOUNT = 0.5`           (Fast 모드 이미지 임베딩 효과)

실제 운영 단계에서 사용 모델/벤더 정책에 맞춰 교체하세요.

## 알려진 한계 (프로토타입)

- `master_loader._parse_pdf_text` 는 단순 정규식 기반 — 실제 PDF 표 구조는 OCR/표추출 모듈로 교체 권장.
- `parsers/base.py` 는 휴리스틱 컬럼 추론. 비정형 셀 병합/멀티헤더는 Precise 모드에서 별도 OCR 파이프라인을 붙이는 전제.
- 인증/권한 시스템 미구현 — Admin 권한은 더미 헤더 `admin_key: BLUE_NINE_ADMIN`. 배포 전 SSO 연동 필요.

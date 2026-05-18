# BLUE NINE — Deployment Guide

데모용 무료 배포: 백엔드는 **Render**, 프론트엔드는 **Vercel** 사용.

```
┌────────────────────┐     HTTPS      ┌──────────────────────┐
│  Vercel            │ ─────────────▶ │  Render              │
│  (Vite SPA)        │   /api/...     │  (FastAPI + uvicorn) │
│  blue-nine.vercel  │ ◀───────────── │  blue-nine-api...    │
└────────────────────┘                └──────────────────────┘
                                                │
                                                ▼ in-memory
                                       (DB 없음 — Source 9, 10)
```

---

## 0. GitHub 준비

이 레포는 이미 `https://github.com/finklkdi-art/AI-TF-pioneer.git` 에 연결됨.

```powershell
cd "C:\Users\CHEIL\desktop\AI TF - pioneer"
git add .gitignore Procfile render.yaml runtime.txt backend/ frontend/ reference/ "Rule Book_260518.docx" README.md DEPLOY.md requirements.txt start_blue_nine.ps1
git commit -m "Feat: BLUE NINE prototype — deploy-ready for Vercel + Render"
git push origin main
```

> ⚠ `.gitignore` 가 `.venv/`, `__pycache__/`, `__buf__/`, 개인 PPT/Excel/PNG 등을 자동 제외함. `git status` 로 확인.

---

## 1. 백엔드 → Render (무료 플랜)

### 방법 A: Blueprint 자동 (권장)

1. https://dashboard.render.com → **New +** → **Blueprint**
2. GitHub repo `finklkdi-art/AI-TF-pioneer` 선택
3. `render.yaml` 자동 감지 → **Apply**
4. 1~3분 후 `blue-nine-api.onrender.com` 발급
5. **Verify**: `curl https://blue-nine-api.onrender.com/healthz` → `{"ok":true}`

### 방법 B: 수동 설정

1. **New +** → **Web Service** → repo 선택
2. 다음 값 입력:

| Field | Value |
|---|---|
| Name | `blue-nine-api` |
| Region | Singapore |
| Branch | `main` |
| Root Directory | (비움) |
| Runtime | Python 3 |
| Build Command | `pip install --upgrade pip && pip install -r backend/requirements.txt` |
| Start Command | `uvicorn backend.main:app --host 0.0.0.0 --port $PORT` |
| Instance Type | Free |

3. **Environment** 탭에서 추가:

   | Key | Value | 필수 여부 |
   |---|---|---|
   | `PYTHON_VERSION` | `3.11.9` | ✅ |
   | `BLUE_NINE_ENV` | `production` | ✅ |
   | `BLUE_NINE_ALLOWED_ORIGINS` | `https://blue-nine.vercel.app` | ✅ Vercel 도메인 확정 후 |
   | `BLUE_NINE_ADMIN_KEY` | `BLUE_NINE_ADMIN` (또는 본인 키) | 권장 |
   | `LLAMAPARSE_API_KEY` | (비워둠 — 향후 통합 시) | ❌ 선택 |
   | `ANTHROPIC_API_KEY` | (비워둠 — 향후 통합 시) | ❌ 선택 |

   > `LLAMAPARSE_API_KEY` / `ANTHROPIC_API_KEY` 는 **현재 코드 경로에서 사용되지 않습니다** — `.env` 스캐폴딩만 되어 있고 휴리스틱 파서가 동작합니다. 키를 비워둬도 `/`/`healthz` 응답에 `"active_pipeline":"heuristic"` 으로 표시됩니다.

> 💤 **Free 플랜 주의사항**: 15분 idle 후 sleep. 다시 깨우는 데 ~30초. 데모 5분 전 한 번 `curl /healthz` 호출로 워밍업.

---

## 2. 프론트엔드 → Vercel (무료)

### 단계

1. https://vercel.com/new → GitHub repo `finklkdi-art/AI-TF-pioneer` 선택
2. **Configure Project** 화면에서:

| Field | Value |
|---|---|
| Framework Preset | Vite |
| Root Directory | `frontend` |
| Build Command | (자동: `npm run build`) |
| Output Directory | (자동: `dist`) |
| Install Command | (자동: `npm install`) |

3. **Environment Variables** 에 추가:
   - `VITE_API_BASE_URL` = `https://blue-nine-api.onrender.com`   ← Render에서 발급된 URL
4. **Deploy** 클릭 → 1분 내 `https://blue-nine.vercel.app` (또는 자동 생성된 도메인) 발급
5. **Verify**:
   - 브라우저로 접속 → 광고주/캠페인 입력, 카테고리 선택, 견적서 업로드
   - DevTools Network 탭에서 `https://blue-nine-api.onrender.com/api/...` 호출 확인

---

## 3. CORS 잠금 (배포 직후 권장)

Vercel 도메인이 확정되면 Render 측에서 origin 제한:

```
Render Dashboard → blue-nine-api → Environment → Add
  BLUE_NINE_ALLOWED_ORIGINS = https://blue-nine.vercel.app,https://blue-nine-git-main-finklkdi-art.vercel.app
```

저장하면 Render 가 자동 재배포. 5초 후 `*` 허용이 화이트리스트로 좁혀짐.

---

## 4. 동작 확인 체크리스트

| 확인 항목 | 명령 / 위치 |
|---|---|
| 백엔드 health | `curl https://blue-nine-api.onrender.com/healthz` |
| Swagger 문서 | https://blue-nine-api.onrender.com/docs |
| 마스터 데이터 적재 | `curl .../api/master/info` 의 `item_count > 0` |
| 카테고리 메타 | `curl .../api/categories` |
| 모니터링 누적 | `curl .../api/monitor/summary` |
| 프론트 SPA | https://blue-nine.vercel.app 접속 → 콘솔 에러 0 |
| 파일 업로드 | UI에서 `영상제작비견적서1.xlsx` 업로드 → 22개 행 + 신호등 표시 |
| 인쇄 (A4) | Chrome → Ctrl+P → A4 미리보기 |
| Excel 다운로드 | 행 수정해 Green 만든 후 ⬇ Excel 클릭 |

---

## 5. 트러블슈팅

| 증상 | 원인 | 해결 |
|---|---|---|
| Vercel 빌드에서 `vercel.json: unknown` | Root Directory 미설정 | Settings → General → Root Directory = `frontend` |
| 프론트에서 CORS 에러 | `BLUE_NINE_ALLOWED_ORIGINS` 누락/오타 | Render env 의 origin 정확히 매칭 (https://, 끝 슬래시 없음) |
| 파일 업로드 시 502 | Render free 플랜 cold start | 1~2회 재시도 또는 사전 워밍업 |
| `master_loaded.item_count: 0` | PDF 파싱 휴리스틱 한계 | 운영시 별도 OCR 모듈 교체 — 데모엔 영향 없음 |
| 한글 파일명 깨짐 | Linux에서 NFC 정규화 차이 | `BLUE_NINE_MASTER_PDF` env 로 명시 |
| `Red 신호등 → Export 불가` (HTTP 409) | 의도된 동작 (Source 32) | UI에서 단가/수량/금액 수정해 Green 만든 후 재시도 |

---

## 6. 데모 당일 운영 팁

1. **데모 5분 전**: `curl https://blue-nine-api.onrender.com/healthz` 로 콜드스타트 미리 깨우기.
2. **시연 순서**:
   1. 메인 화면 → 휘발성 배너 강조 ("브라우저 종료 시 자동 파기" — Source 9, 10)
   2. Step 1: 제작비 → Step 2: 영상 → 토글로 Precise 선택
   3. `영상제작비견적서1.xlsx` 업로드 → 22행 + 신호등 노출
   4. Red 행 클릭 → "AI 판정 근거" 패널 노출 → 단가 수정 → Green 전환
   5. 상단 모니터링 패널의 latency/cost 누적 보여주기 (요구사항 4)
   6. ⬇ Excel 다운로드 → 시트 열어서 SUM 수식이 살아있음을 시연
   7. Ctrl+P → A4 미리보기 (Source 25a)
3. **로그 모니터**: `https://blue-nine-api.onrender.com/api/monitor/summary` 를 별도 탭으로 띄워두면 발표 중 실시간 호출 추이 노출 가능.

---

## 7. 발급 후 업데이트할 URL 메모

| 자원 | URL |
|---|---|
| GitHub | https://github.com/finklkdi-art/AI-TF-pioneer |
| Render (backend) | https://blue-nine-api.onrender.com ← 배포 후 갱신 |
| Vercel (frontend) | https://blue-nine.vercel.app ← 배포 후 갱신 |
| Swagger | https://blue-nine-api.onrender.com/docs |

"""
BLUE NINE — Backend FastAPI entrypoint.

설계 원칙 (Rule Book v1.0)
- 완전 휘발성 세션 (Source 9, 10): DB / 디스크 영속화 없음. 메모리만.
- 2단계 카테고리 분기 (Source 11, 12, 17) — 라우터 키: (l1, l2).
- 3단계 신호등 검증 (Source 29~32) — Row + Sheet 합계 동시.
- 입력 숫자 보존 (Source 28) — 시스템은 input 금액을 임의로 수정하지 않음.
- 모니터링 미들웨어 (요구사항 4) — Latency/Traffic/Token-Cost.
"""
from __future__ import annotations
import io
import os
import uuid
from datetime import datetime
from pathlib import Path
from typing import List, Optional

from fastapi import FastAPI, UploadFile, File, Form, HTTPException, Header
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse, JSONResponse
from pydantic import BaseModel

from .config import settings           # noqa: F401  (side effect: .env loaded)
from .schemas import (
    EstimateDocument, EstimateRow, ParseRequest, UpdateRowRequest,
    MasterRefItem, SourceFileSummary,
)
from .session_store import STORE
from .monitoring import BillingMiddleware, BUS
from .master_loader import MASTER
from .parsers import get_parser
from .parsers.billing_status import parse_billing
from .parsers.post_process import post_process_rows
from .services.pdf_pipeline import pdf_to_estimate_rows
from .jeongga_whitelist import generate_jeongga_rows
from .validation import evaluate_document, evaluate_triangle
from .exporter import build_xlsx


app = FastAPI(
    title="BLUE NINE API",
    version="1.0.0-prototype",
    description="사내 AE를 위한 범용 견적서 효율화 솔루션 — 휘발성 세션 기반.",
)

# CORS — 배포 환경에 따라 화이트리스트 적용.
# 로컬 개발(=Vite proxy, same-origin)에서는 사용되지 않음.
# Vercel 배포 후엔 BLUE_NINE_ALLOWED_ORIGINS 환경변수로 도메인 좁히기를 권장.
_origins_env = settings.allowed_origins.strip()
if _origins_env == "*" or not _origins_env:
    _allow_origins = ["*"]
    _allow_origin_regex = None
else:
    _allow_origins = [o.strip() for o in _origins_env.split(",") if o.strip()]
    _allow_origin_regex = None

app.add_middleware(
    CORSMiddleware,
    allow_origins=_allow_origins,
    allow_origin_regex=_allow_origin_regex,
    allow_credentials=False,        # 쿠키/세션 미사용 (Source 9, 10)
    allow_methods=["*"],
    allow_headers=["*"],
    expose_headers=["Content-Disposition"],   # Excel 다운로드 파일명 노출
)
app.add_middleware(BillingMiddleware)


@app.on_event("startup")
def _startup() -> None:
    # Source 16, 36 — '제작단가기준집.pdf' 마스터 메모리 적재
    cnt = MASTER.load()
    print(f"[BLUE NINE] master items loaded: {cnt}")


# ─────────────────────────────────────────────────────────────
# 세션
# ─────────────────────────────────────────────────────────────
@app.post("/api/session/new")
def session_new():
    sid = STORE.create()
    return {"session_id": sid, "ttl_sec": 30 * 60}


@app.post("/api/session/destroy")
def session_destroy(session_id: str = Form(...)):
    STORE.drop(session_id)
    return {"ok": True}


# ─────────────────────────────────────────────────────────────
# 카테고리 메타 (프론트 2단계 분기 UI 데이터원)
# ─────────────────────────────────────────────────────────────
@app.get("/api/categories")
def categories():
    return {
        "step1": [
            {"key": "production", "label": "제작비"},
            {"key": "media", "label": "매체비"},
        ],
        "step2": {
            "production": [
                {"key": "video",  "label": "영상",  "template": "영상제작비견적서"},
                {"key": "radio",  "label": "라디오","template": "라디오제작비견적서"},
                {"key": "print",  "label": "인쇄",  "template": "인쇄제작비견적서"},
                {"key": "btl",    "label": "BTL",   "template": "BTL제작비견적서"},
                {"key": "other",  "label": "기타",  "template": "Generic Production"},
            ],
            "media": [
                {"key": "tvc",     "label": "TVC",     "template": "Billing Status"},
                {"key": "radio",   "label": "라디오",   "template": "Billing Status"},
                {"key": "print",   "label": "PRINT",   "template": "Billing Status"},
                {"key": "digital", "label": "디지털",   "template": "Billing Status"},
                {"key": "other",   "label": "기타",    "template": "Billing Status"},
            ],
        },
    }


# ─────────────────────────────────────────────────────────────
# 마스터 데이터 (제작단가기준집)
# ─────────────────────────────────────────────────────────────
@app.get("/api/master/info")
def master_info():
    return MASTER.info()


@app.get("/api/master/items", response_model=List[MasterRefItem])
def master_items():
    return MASTER.items()


@app.post("/api/master/reload")
def master_reload(admin_key: str = Header(default="")):
    # Source 36 — admin 전용 reload.
    if admin_key != settings.admin_key:
        raise HTTPException(status_code=403, detail="admin only")
    cnt = MASTER.load()
    return {"reloaded": True, "items": cnt, "info": MASTER.info()}


# ─────────────────────────────────────────────────────────────
# 견적서 — 업로드 / 파싱 / 평가 / 수정 / Export
# ─────────────────────────────────────────────────────────────
def _is_billing_filename(name: str) -> bool:
    """파일명만으로 Billing Status 여부 판별 (Source 20)."""
    if not name:
        return False
    n = name.lower()
    return any(k in n for k in ("billing", "빌링", "집행내역", "매체집행"))


@app.post("/api/estimate/parse")
async def estimate_parse(
    session_id: str = Form(...),
    category_l1: str = Form(...),          # production | media
    category_l2: str = Form(...),          # video | radio | print | btl | other | tvc | digital ...
    mode: str = Form("precise"),           # fast | precise (Source 3, 4)
    client: Optional[str] = Form(None),
    campaign: Optional[str] = Form(None),
    version_label: str = Form("초안"),
    applied_count: int = Form(1),          # ✨ 정가항목 적용 건수 — AE 가 Step 3 에서 수동 설정
    files: List[UploadFile] = File(...),   # ✨ 다중 파일 입력
):
    """
    여러 개의 협력사 견적서 파일을 한 번에 수신.
    - 파일별로 휴리스틱 파서 실행 → rows 통합 (source_file 태깅).
    - 파일명이 'billing*' / '빌링*' 패턴이면 Billing Status 로 인식해 삼각 검증.
    - 결과는 단일 EstimateDocument 로 머지 후 신호등 평가.
    """
    if category_l1 not in ("production", "media"):
        raise HTTPException(400, "category_l1 must be production|media")
    if not files:
        raise HTTPException(400, "최소 1개 이상의 파일이 필요합니다.")

    parser = get_parser(category_l1, category_l2)

    all_rows: List[EstimateRow] = []
    sources: List[SourceFileSummary] = []
    billing_rows_aggregated: List[dict] = []
    doc_notes_buf: List[str] = []        # PDF 파이프라인 등에서 채워지는 메타 note

    for uf in files:
        fname = uf.filename or "uploaded.bin"
        try:
            blob = await uf.read()
        except Exception as e:
            sources.append(SourceFileSummary(filename=fname, error=f"read_failed: {e}"))
            continue
        STORE.put_upload(session_id, fname, blob)        # 휘발성 보관 (Source 9)

        # Billing Status 자동 식별 — 매체비 카테고리에서만 의미 있음
        if category_l1 == "media" and _is_billing_filename(fname):
            try:
                br = parse_billing(blob)
                billing_rows_aggregated.extend(br)
                sources.append(SourceFileSummary(
                    filename=fname, rows=len(br), role="billing", size_bytes=len(blob),
                ))
            except Exception as e:
                sources.append(SourceFileSummary(
                    filename=fname, role="billing", size_bytes=len(blob),
                    error=f"billing_parse_failed: {e}",
                ))
            continue

        # ── 파일 형식별 분기 ────────────────────────────────
        ext = (fname.rsplit(".", 1)[-1] or "").lower()

        if ext == "pdf":
            # PDF → LlamaParse + Anthropic 파이프라인 (Source 4 Precise 모드)
            try:
                pdf_rows, meta = pdf_to_estimate_rows(
                    blob,
                    category_l1=category_l1,
                    category_l2=category_l2,
                    mode=mode,
                    filename=fname,
                )
                for r in pdf_rows:
                    r.source_file = fname
                all_rows.extend(pdf_rows)
                sources.append(SourceFileSummary(
                    filename=fname, rows=len(pdf_rows), role="estimate-pdf",
                    size_bytes=len(blob),
                ))
                doc_notes_buf.append(
                    f"📄 LLM(PDF) '{fname}' — markdown {meta['markdown_length']:,} chars · "
                    f"신뢰도 {meta['llm_confidence']*100:.0f}%"
                )
                doc_notes_buf.extend(meta.get("llm_notes", []))
            except Exception as e:
                sources.append(SourceFileSummary(
                    filename=fname, role="estimate-pdf", size_bytes=len(blob),
                    error=f"pdf_pipeline_failed: {e}",
                ))
            continue

        # 그 외 (xlsx/xls/csv) — 휴리스틱 파서 (프로파일 매칭 우선)
        try:
            rows = parser.parse(blob, filename=fname)
            for r in rows:
                r.source_file = fname
            all_rows.extend(rows)
            sources.append(SourceFileSummary(
                filename=fname, rows=len(rows), role="estimate", size_bytes=len(blob),
            ))
        except Exception as e:
            sources.append(SourceFileSummary(
                filename=fname, role="estimate", size_bytes=len(blob),
                error=f"parse_failed: {e}",
            ))

    # ── 후처리 (단위노이즈 / 동의어 dedup / 0행 제거) ──────────────
    cleaned_rows, pp_stats = post_process_rows(all_rows)
    doc_notes_buf.append(
        f"🧹 후처리: 입력 {pp_stats['input']} → 정리 {len(cleaned_rows)}행 "
        f"(합계행 {pp_stats['dropped_aggregate']}, 단위 {pp_stats['unit_stripped']}, "
        f"중복병합 {pp_stats['merged_duplicates']}, 0원 {pp_stats['dropped_zero']})"
    )

    # ── 정가항목 주입 (Source 16; 2026-05-19 비즈니스 룰) ─────────
    # 카테고리별 표준 단가 세트 매핑 (영상/인쇄만 보유).
    # media + 그 외 production(radio/btl/other) 은 빈 리스트 → A=0.
    jeongga_rows = (
        generate_jeongga_rows(category_l2, applied_count)
        if category_l1 == "production" else []
    )
    merged_rows = jeongga_rows + cleaned_rows

    doc = EstimateDocument(
        estimate_id=f"est-{uuid.uuid4().hex[:8]}",
        session_id=session_id,
        version_label=version_label,
        category_l1=category_l1,                                  # type: ignore[arg-type]
        category_l2=category_l2,
        mode=mode,                                                # type: ignore[arg-type]
        client=client,
        campaign=campaign,
        issue_date=datetime.now().strftime("%Y-%m-%d"),
        rows=merged_rows,
        sources=sources,
    )
    if jeongga_rows:
        doc.notes.append(
            f"📐 정가항목 자동 주입: 카테고리 '{category_l2}' 표준 단가 × 적용 건수 {applied_count} (AE 수동 입력)"
        )
    elif category_l1 == "production":
        doc.notes.append(
            f"📐 정가항목 미적용: 카테고리 '{category_l2}' 에는 표준 단가 세트가 정의되지 않음"
        )

    # 매체비 삼각 검증 (Source 19~22) — billing role 파일이 있을 때만
    if category_l1 == "media" and billing_rows_aggregated:
        billing_sum = sum(r["charged"] for r in billing_rows_aggregated)
        media_paid_sum = sum(r["paid"] for r in billing_rows_aggregated)
        charged_sum = sum(r.amount for r in doc.rows if r.section == "매체청구액")
        if charged_sum == 0:
            charged_sum = sum(r.amount for r in doc.rows)
        fee_sum = max(billing_sum - media_paid_sum, 0)
        doc.triangle = evaluate_triangle(charged_sum, billing_sum, media_paid_sum, fee_sum)
        if not doc.triangle.consistent:
            doc.warnings.append(
                f"⚠ 매체비 삼각 검증 실패 — Δ={doc.triangle.delta:,.0f}"
            )

    # 파일별 결과를 notes 에 요약
    ok_files = [s for s in sources if not s.error]
    err_files = [s for s in sources if s.error]
    doc.notes.append(
        f"📥 입력 {len(files)}개 파일 처리 — 성공 {len(ok_files)}건, 실패 {len(err_files)}건"
    )
    doc.notes.extend(doc_notes_buf)
    for s in err_files:
        doc.warnings.append(f"⚠ '{s.filename}' 파싱 실패: {s.error}")

    evaluate_document(doc)
    STORE.put_estimate(session_id, doc)
    return doc


@app.get("/api/estimate/{session_id}", response_model=List[EstimateDocument])
def estimate_list(session_id: str):
    return STORE.list_estimates(session_id)


@app.get("/api/estimate/{session_id}/{estimate_id}", response_model=EstimateDocument)
def estimate_get(session_id: str, estimate_id: str):
    doc = STORE.get_estimate(session_id, estimate_id)
    if not doc:
        raise HTTPException(404, "estimate not found (휘발성 세션 만료 가능)")
    return doc


@app.post("/api/estimate/update_row", response_model=EstimateDocument)
def update_row(req: UpdateRowRequest):
    doc = STORE.get_estimate(req.session_id, req.estimate_id)
    if not doc:
        raise HTTPException(404, "estimate not found")
    for r in doc.rows:
        if r.id == req.row_id:
            for k, v in req.patch.items():
                if k in ("unit_price", "quantity", "amount"):
                    setattr(r, k, float(v))
                elif k in ("item_name", "vendor", "note", "section", "category"):
                    setattr(r, k, v)
            break
    evaluate_document(doc)
    STORE.put_estimate(req.session_id, doc)
    return doc


@app.post("/api/estimate/{session_id}/{estimate_id}/export.xlsx")
def estimate_export(session_id: str, estimate_id: str):
    doc = STORE.get_estimate(session_id, estimate_id)
    if not doc:
        raise HTTPException(404, "estimate not found")
    # 정책 변경 (2026-05-19): Red 신호등 상태에서도 다운로드 허용.
    # AE 가 검증 경고를 보면서도 즉시 청구 진행이 필요한 케이스가 있음.
    # 경고는 응답 헤더 X-Blue-Nine-Validation 으로 노출.
    data = build_xlsx(doc)
    bio = io.BytesIO(data)
    fname = f"BLUE_NINE_{doc.category_l1}_{doc.category_l2}_{doc.estimate_id}.xlsx"
    return StreamingResponse(
        bio,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": f'attachment; filename="{fname}"',
            "X-Blue-Nine-Validation": doc.overall_light,
        },
    )


# ─────────────────────────────────────────────────────────────
# 모니터링 API (Latency / Traffic / Token cost) — 요구사항 4
# ─────────────────────────────────────────────────────────────
@app.get("/api/monitor/summary")
def monitor_summary():
    return BUS.aggregate()


@app.get("/api/monitor/logs")
def monitor_logs(limit: int = 50):
    return BUS.snapshot(last_n=limit)


@app.get("/")
def root():
    return {
        "service": "BLUE NINE",
        "version": app.version,
        "rule_book": "v1.0",
        "memory_only": True,
        "env": settings.env,
        "master_loaded": MASTER.info(),
        "llm": settings.llm_status(),       # 안전 노출 (값 자체는 미공개)
    }


@app.get("/healthz")
def healthz():
    """Render / Vercel health probe."""
    return {"ok": True}

"""
PDF → 구조화된 EstimateRow 추출 파이프라인 (Source 4 Precise 모드 핵심 경로).

흐름:
  1) LlamaParse  : PDF → markdown 텍스트 (표 구조 보존)
  2) Anthropic   : markdown → 라인아이템 분류 (정가/외주/대행수수료 등)
  3) Validator   : Pydantic 검증 후 EstimateRow 로 변환

설계 원칙:
  - 키가 없으면 RuntimeError → main.py 가 graceful sources[].error 로 surfacing
  - 입력 숫자는 절대 손대지 않음 (Source 28)
  - 시스템 프롬프트는 캐시 가능한 stable prefix 로 구성 → 다중 PDF 처리 시 토큰 비용 절감
"""
from __future__ import annotations
import os
import re
import tempfile
import uuid
from typing import List, Optional, Literal

from pydantic import BaseModel, Field

from ..config import settings
from ..schemas import EstimateRow


# ─────────────────────────────────────────────────────────────
# Phase 1 — LlamaParse (PDF → markdown)
# ─────────────────────────────────────────────────────────────
def llama_parse_pdf_to_markdown(blob: bytes, *, mode: str = "precise", filename: str = "input.pdf") -> str:
    """LlamaParse 호출. mode 에 따라 fast/precise 분기 (Source 4)."""
    if not settings.llamaparse_api_key:
        raise RuntimeError("LLAMAPARSE_API_KEY 미설정 — PDF 파싱 불가")

    # llama-parse 0.5+ : new unified SDK 권장이나 prototype 단계엔 기존 패키지 사용
    from llama_parse import LlamaParse  # type: ignore[import-untyped]

    # Source 4 매핑:
    #   fast   → Vision Embedding 신속 스캔  ≈ parse_page_without_llm
    #   precise→ 전체 OCR + 라인 정밀 교차검증  ≈ parse_page_with_llm (default)
    parse_mode = "parse_page_without_llm" if mode == "fast" else "parse_page_with_llm"

    parser = LlamaParse(
        api_key=settings.llamaparse_api_key,
        result_type="markdown",
        parse_mode=parse_mode,
        language="ko",
        verbose=False,
        # PDF 가 견적서/표 위주이므로 leaf node 단위로 안 자르고 통째로 받기
        split_by_page=False,
    )

    # LlamaParse 는 path 기반 — 임시 파일로 우회 (휘발성 보장: finally 에서 즉시 삭제)
    suffix = os.path.splitext(filename)[1] or ".pdf"
    tmp = tempfile.NamedTemporaryFile(suffix=suffix, delete=False)
    try:
        tmp.write(blob)
        tmp.flush()
        tmp.close()
        documents = parser.load_data(tmp.name)
    finally:
        try:
            os.unlink(tmp.name)
        except OSError:
            pass

    return "\n\n".join(getattr(d, "text", "") for d in documents).strip()


# ─────────────────────────────────────────────────────────────
# Phase 2 — Anthropic (markdown → 구조화)
# ─────────────────────────────────────────────────────────────
Section = Literal["정가항목", "외주비", "대행수수료", "매체청구액", "매체지급액", "매체수수료", "기타"]


class LLMExtractedRow(BaseModel):
    section: Section = Field(..., description="섹션 분류")
    item_name: str = Field(..., description="라인아이템 명")
    vendor: Optional[str] = Field(None, description="협력사/매체사 명")
    unit_price: float = Field(0.0, description="단가 (숫자만, 원 단위)")
    quantity: float = Field(1.0, description="수량")
    amount: float = Field(0.0, description="금액 (단가×수량 이거나 명시된 합계)")
    note: Optional[str] = Field(None, description="비고 / 부연")


class LLMExtractedRows(BaseModel):
    rows: List[LLMExtractedRow] = Field(default_factory=list)
    overall_confidence: float = Field(
        0.9, ge=0.0, le=1.0,
        description="추출 전체 확신도 — 표 구조가 명확할수록 높음",
    )
    notes: List[str] = Field(default_factory=list, description="AE 가 참고할 만한 메모")


_SYSTEM_PROMPT = """\
당신은 한국 광고대행사의 견적서/단가표 PDF 를 분석해 라인아이템을 구조화 추출하는 분석가입니다.

[BLUE NINE 섹션 분류 규칙]
■ 제작비 (production) 카테고리:
  · 정가항목   — 회사 정가 항목 (기본료, Copy료, Creative Work료, Direction료, 기획관리비 등)
  · 외주비     — 협력사/PD 비용 (PD료, 촬영연출, POST프로덕션, 편집, 녹음, BGM, 디자인, 인쇄, 후반작업 등)
  · 대행수수료 — 광고대행사 수수료 (= 외주비 × 17.65% 또는 10%)

■ 매체비 (media) 카테고리:
  · 매체청구액 — 광고주에게 청구하는 매체비
  · 매체지급액 — 매체사에 지급하는 금액
  · 매체수수료 — 광고대행사가 수취하는 수수료

■ 분류 불가능한 경우만 '기타' 사용. 가능한 한 위 7개 안으로 매핑.

[엄수 규칙 — Source 28 (입력 보존 원칙)]
1. PDF 에 명시된 숫자만 사용. 추정/계산/반올림으로 새로운 숫자를 만들지 말 것.
2. 금액(amount) 칸이 비어 있고 단가(unit_price)·수량(quantity)이 모두 있으면 amount = unit_price × quantity 만 허용.
3. 헤더 행, 합계/소계/총계 행은 제외. 데이터 행만 추출.
4. 카테고리가 모호하면 overall_confidence 를 낮추고 notes 에 사유 명시 (예: "기타 제작비로 보임").
5. 추출 가능한 라인이 없으면 rows=[] 빈 배열 반환 + notes 설명.

[항목명 분류 힌트]
· "기본료", "Copy료", "Creative Work료", "Direction료", "기획관리비" → 정가항목
· "PD료", "Producer", "촬영", "촬영연출", "감독" → 외주비
· "POST", "편집", "EDIT", "DI", "2D", "3D", "VFX", "CG" → 외주비
· "녹음", "성우", "음향", "BGM" → 외주비
· "디자인", "리터칭", "보정", "일러스트" → 외주비
· "대행수수료", "Agency Fee" → 대행수수료
· "매체비", "Gross", "광고주청구" → 매체청구액
· "지급액", "Net", "매체사지급" → 매체지급액
· "매체수수료", "Commission" → 매체수수료

JSON 으로만 응답하며, 한글 항목명은 그대로 보존합니다.
"""


def anthropic_markdown_to_rows(
    markdown: str,
    *,
    category_l1: str,
    category_l2: str,
    mode: str = "precise",
) -> LLMExtractedRows:
    """Claude 호출 → 구조화 추출."""
    if not settings.anthropic_api_key:
        raise RuntimeError("ANTHROPIC_API_KEY 미설정 — LLM 분류 불가")

    import anthropic  # type: ignore[import-untyped]

    client = anthropic.Anthropic(api_key=settings.anthropic_api_key)

    user_msg = (
        f"카테고리: {category_l1} / {category_l2}\n"
        f"처리 모드: {mode}\n\n"
        f"[PDF 추출 markdown — 시작]\n"
        f"{markdown}\n"
        f"[PDF 추출 markdown — 끝]\n\n"
        f"위 내용에서 라인아이템을 BLUE NINE 섹션 분류 규칙에 맞춰 JSON 으로 추출하세요."
    )

    # messages.parse + Pydantic 으로 스키마 강제 + 자동 검증
    # cache_control 로 system prompt 캐싱 (반복 PDF 처리 시 ~90% 비용 절감)
    response = client.messages.parse(
        model=settings.anthropic_model,            # .env 의 ANTHROPIC_MODEL (default: claude-sonnet-4-6)
        max_tokens=8192,
        system=[{
            "type": "text",
            "text": _SYSTEM_PROMPT,
            "cache_control": {"type": "ephemeral"},
        }],
        messages=[{"role": "user", "content": user_msg}],
        output_format=LLMExtractedRows,
    )
    return response.parsed_output


# ─────────────────────────────────────────────────────────────
# Phase 3 — 두 단계 통합 + EstimateRow 변환
# ─────────────────────────────────────────────────────────────
def pdf_to_estimate_rows(
    blob: bytes,
    *,
    category_l1: str,
    category_l2: str,
    mode: str = "precise",
    filename: str = "input.pdf",
) -> tuple[List[EstimateRow], dict]:
    """
    PDF blob → List[EstimateRow] + 메타 정보 (markdown 길이, LLM 신뢰도 등).
    main.py 에서 호출 — 실패 시 raise.
    """
    md = llama_parse_pdf_to_markdown(blob, mode=mode, filename=filename)
    if not md.strip():
        raise RuntimeError("LlamaParse 결과 비어 있음 — PDF 가 이미지/스캔본일 가능성")

    extracted = anthropic_markdown_to_rows(
        md, category_l1=category_l1, category_l2=category_l2, mode=mode,
    )

    rows: List[EstimateRow] = []
    for r in extracted.rows:
        rows.append(EstimateRow(
            id=f"r-{uuid.uuid4().hex[:8]}",
            section=r.section,
            item_name=r.item_name[:80],
            vendor=r.vendor,
            unit_price=r.unit_price,
            quantity=r.quantity,
            amount=r.amount,
            note=r.note,
        ))

    meta = {
        "markdown_length": len(md),
        "llm_confidence": extracted.overall_confidence,
        "llm_notes": extracted.notes,
    }
    return rows, meta

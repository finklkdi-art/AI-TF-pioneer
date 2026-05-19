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
당신은 한국 광고대행사가 받은 협력사 견적서 PDF 를 분석해 라인아이템을 구조화 추출하는 분석가입니다.

[★★ 핵심 비즈니스 룰 ★★]
  · 협력사가 보내온 input PDF 에는 '정가항목'(기획료/카피료/크리에이티브 워크료/디렉션료/
    자료조사비/제작진행비)이 절대 포함되지 않습니다.
  · 따라서 어떤 행도 '정가항목' 으로 분류하지 마십시오. 모든 제작 관련 라인은
    '외주비' 로 분류합니다.
  · 만약 PDF 안에 '기획료', '카피료' 같은 명칭이 등장하더라도 그것은 협력사의
    내부 비용 분류일 뿐이므로 '외주비' 로 흡수합니다.

[BLUE NINE 섹션 키]
  - production: '외주비' (모든 협력사 제작 라인) / '대행수수료' (협력사가 명시한 경우만)
  - media:      '매체청구액' / '매체지급액' / '매체수수료'

[엄수 규칙 — Source 28 (입력 보존 원칙)]
1. PDF 에 명시된 숫자만 사용. 추정·계산·반올림으로 새 숫자를 만들지 말 것.
2. amount 칸이 비어 있고 unit_price·quantity 가 모두 있으면 amount = unit_price × quantity 만 허용.
3. 헤더 행, 합계/소계/총계 행은 제외. 데이터 행만 추출.
4. 항목명에 협력사·세부내역이 함께 있으면 가능한 한 그대로 보존 (예: "감독료 / 편집·CG연출 포함").
5. 추출 가능한 라인이 없으면 rows=[] 빈 배열 반환 + notes 설명.

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
    demoted = 0
    for r in extracted.rows:
        section = r.section
        item_name = r.item_name[:80]
        # ── 정가항목 차단 (server-side guardrail) ───────────────────
        # input 파일에는 정가항목이 없다는 비즈니스 룰. LLM 이 잘못 분류해도 외주비로 강등.
        if category_l1 == "production" and section == "정가항목":
            section = "외주비"
            demoted += 1
        rows.append(EstimateRow(
            id=f"r-{uuid.uuid4().hex[:8]}",
            section=section,
            item_name=item_name,
            vendor=r.vendor,
            unit_price=r.unit_price,
            quantity=r.quantity,
            amount=r.amount,
            note=r.note,
        ))
    if demoted:
        extracted.notes.append(
            f"⛓ '정가항목' {demoted}건을 input 룰에 따라 '외주비' 로 강등"
        )

    meta = {
        "markdown_length": len(md),
        "llm_confidence": extracted.overall_confidence,
        "llm_notes": extracted.notes,
    }
    return rows, meta

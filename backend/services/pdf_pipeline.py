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
  · input PDF 에는 '정가항목' (기획료/카피료/크리에이티브 워크료/디렉션료/자료조사비/제작진행비) 이
    절대 포함되지 않습니다. 어떤 행도 '정가항목' 으로 분류하지 말 것 — 모든 제작 관련 라인은
    '외주비' 로 분류합니다.
  · 'Editing' / '편집', 'Producing' / 'PD료', 'Recording' / '녹음실' 처럼 표기만 다르고
    실제로 같은 업무를 가리키는 항목은 사용자 친화적인 한국어 표기로 통일하여 추출합니다.

[BLUE NINE 섹션 키]
  - production: 모든 추출 라인은 '외주비'
  - media:      '매체청구액' / '매체지급액' / '매체수수료'

[항목명 정제 규칙 — 매우 중요]
  ① 단위·수량 텍스트 ('명/건/식/편/회/일/팀/개/인/시간/벌') 가 항목명 칸에 섞여 있으면 제거.
     예) '성우 2명' → '성우' (수량 2 는 quantity 필드로 분리), '편집 1편' → '편집'.
  ② 합계/소계/총계/VAT/부가세/대행수수료 라인은 절대 추출하지 말 것 — 시스템이 별도 계산.
  ③ 협력사명이 보이면 vendor 필드에 분리해 넣고, item_name 에는 업무 카테고리만 남길 것.
  ④ 숫자가 비어 있거나 '-', '0' 인 행은 추출하지 말 것 (살아있는 데이터만 노출).
  ⑤ 동일 의미 항목이 PDF 안에 2회 이상 나오면 한 번만 추출하고, amount 는 합산.

[자유 양식·개인사업자·프리랜서 견적서 항목 인식 가이드]
  · 일부 견적서 (특히 개인사업자/프리랜서 양식) 는 상단 '업무내용' 칸에 캠페인명만 적고
    실제 직종/역할은 청구인 정보 근처 (하단부) 에 별도로 표기합니다.
    예) 상단 = '삼성 비스포크 어댑테이션' (캠페인), 하단 = '성우' (실제 외주비 항목)
  · 이 경우 item_name 으로 **직종/역할 키워드** 를 우선 사용하세요:
    성우, 내레이터, 녹음기사, 엔지니어, 믹싱, 편집, 에디터, 디자이너, 아트디렉터,
    촬영기사, 촬영감독, 조명감독, PD, 감독, 조감독, 모델, 스타일리스트,
    메이크업, 헤어, 푸드스타일리스트, VFX, 2D, 3D, CG, DI 등.
  · 캠페인명 (예: '삼성 비스포크 ...') 은 item_name 으로 쓰지 말고 note 또는 무시.

[숫자 보존 규칙 — Source 28]
  · PDF 에 명시된 숫자만 사용. 추정·계산·반올림으로 새 숫자를 만들지 말 것.
  · amount 가 비어있고 unit_price·quantity 가 모두 있을 때만 amount = unit_price × quantity 허용.

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

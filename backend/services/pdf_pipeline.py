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
import io
import os
import re
import tempfile
import uuid
from typing import List, Optional, Literal

from pydantic import BaseModel, Field

from ..config import settings
from ..schemas import EstimateRow
from .bible_cache import get_bible_text, get_output_layout
from .intent_classifier import classify_purpose, format_purpose_context_for_llm


# ─────────────────────────────────────────────────────────────
# Phase 1 — PDF → markdown (LlamaParse 우선 + pdfplumber/pypdf 폴백)
# ─────────────────────────────────────────────────────────────
def _local_pdf_to_text(blob: bytes) -> str:
    """LlamaParse 없이 표준 라이브러리로 PDF 텍스트 추출.
    pdfplumber → pypdf 순으로 시도. 둘 다 실패하면 빈 문자열.
    표 구조는 보존되지 않지만 라인 단위 텍스트는 확보 → LLM 컨텍스트로 활용 가능."""
    # 1) pdfplumber
    try:
        import pdfplumber
        bio = io.BytesIO(blob)
        out: list[str] = []
        with pdfplumber.open(bio) as pdf:
            for i, page in enumerate(pdf.pages):
                # 텍스트 우선, 그 다음 표를 markdown-스타일로 직렬화
                text = page.extract_text() or ""
                if text.strip():
                    out.append(f"## Page {i+1}\n{text}")
                try:
                    tables = page.extract_tables() or []
                    for ti, tb in enumerate(tables):
                        if not tb: continue
                        rows = ["| " + " | ".join("" if c is None else str(c).strip() for c in r) + " |" for r in tb]
                        if rows:
                            sep = "| " + " | ".join("---" for _ in tb[0]) + " |"
                            out.append(f"### Page {i+1} Table {ti+1}\n" + "\n".join([rows[0], sep, *rows[1:]]))
                except Exception:
                    pass
        text = "\n\n".join(out).strip()
        if text:
            return text
    except ImportError:
        pass
    except Exception:
        pass
    # 2) pypdf
    try:
        from pypdf import PdfReader
        bio = io.BytesIO(blob)
        r = PdfReader(bio)
        chunks = []
        for i, page in enumerate(r.pages):
            t = page.extract_text() or ""
            if t.strip():
                chunks.append(f"## Page {i+1}\n{t}")
        return "\n\n".join(chunks).strip()
    except ImportError:
        return ""
    except Exception:
        return ""


def llama_parse_pdf_to_markdown(blob: bytes, *, mode: str = "precise", filename: str = "input.pdf") -> str:
    """PDF → markdown 변환. 3단계 폴백:
      ① LlamaParse (키 있고 import 성공 시) — 표 구조까지 정밀 추출
      ② pdfplumber — 텍스트 + 표 추출 (표 markdown 직렬화)
      ③ pypdf      — 단순 텍스트 추출
    어떤 단계든 결과 텍스트가 생기면 그것을 반환. 모두 실패 시 RuntimeError.
    """
    # ── ① LlamaParse 시도 (API 키 있을 때만)
    if settings.llamaparse_api_key:
        try:
            from llama_parse import LlamaParse  # type: ignore[import-untyped]
            parse_mode = "parse_page_without_llm" if mode == "fast" else "parse_page_with_llm"
            parser = LlamaParse(
                api_key=settings.llamaparse_api_key,
                result_type="markdown",
                parse_mode=parse_mode,
                language="ko",
                verbose=False,
                split_by_page=False,
            )
            suffix = os.path.splitext(filename)[1] or ".pdf"
            tmp = tempfile.NamedTemporaryFile(suffix=suffix, delete=False)
            try:
                tmp.write(blob); tmp.flush(); tmp.close()
                documents = parser.load_data(tmp.name)
            finally:
                try: os.unlink(tmp.name)
                except OSError: pass
            md = "\n\n".join(getattr(d, "text", "") for d in documents).strip()
            if md:
                return md
            # LlamaParse 가 빈 결과 → 폴백
        except ImportError:
            # llama_parse 패키지 자체가 없음 — 운영 환경 의존성 누락 시 자동 폴백
            pass
        except Exception:
            # 네트워크/쿼터/타임아웃 등 — 폴백
            pass

    # ── ② pdfplumber / ③ pypdf 폴백
    md = _local_pdf_to_text(blob)
    if md:
        return md

    # 어느 추출기도 텍스트를 못 뽑은 경우
    raise RuntimeError(
        "PDF 텍스트 추출 실패 — LlamaParse 호출 실패 + pdfplumber/pypdf 도 빈 결과"
    )


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


_SYSTEM_PROMPT_CORE = """\
당신은 한국 광고대행사가 받은 협력사 견적서 (PDF/Excel 표 markdown 직렬화) 를 분석해
라인아이템을 구조화 추출하는 시맨틱 파서입니다. **3단계 의도 중심 파이프라인**:

  STEP 1 — 청구 목적 탑다운 판별 (user message 의 [STEP 1] 블록 참조)
  STEP 2 — 표준 엔티티 (Ground Truth) 와 의미 매칭 (user message 의 [STEP 2] 블록 참조)
  STEP 3 — Output 양식 동기화 (sanitize: 단가/수량/금액 자동 보정)

[STEP 3 출력 sanitization 규칙]
  · 소스에 "성우료 1,500,000원" 처럼 합계만 있고 단가/수량이 비어있으면:
      unit_price=1500000, quantity=1, amount=1500000 으로 자동 채울 것 (에러 X).
  · 합계 행 (소계/총계/VAT/대행수수료/Total/(A)(B)(C)) 은 절대 추출 금지.
  · 의미 매칭이 명확하지 않은 행은 추출 금지 (rows=[] 반환 + notes 설명).

[1단계 — 의도 분류 (Intent Taxonomy)]
입력 문서의 각 행/블록에 대해 먼저 아래 분류 중 하나를 머릿속에서 결정:
  · 인적 용역  : 성우/모델/PD/감독/조감독/편집·믹싱·CG 작업자 등 사람의 일
  · 장비 대여  : 카메라/조명/스튜디오/이동차/특수기자재
  · 후반 작업  : DI/Editing/2D/3D/VFX/녹음
  · 수수료    : 대행수수료/위탁수수료 → ❌ 절대 추출 금지 (시스템 자동 산입)
  · 합계·소계 : 합계/소계/총계/VAT/부가세 → ❌ 절대 추출 금지
  · 정가항목  : 기획료/카피료/크리에이티브 워크료/디렉션료/자료조사비 → ❌ 절대 추출 금지
              (이 6종은 input 파일에 존재하지 않음 — AE 가 별도 입력)
처음 4개 분류만 추출 대상 = 모두 '외주비' 섹션.

[2단계 — 엔티티 묶음 (Entity Grouping)]
표가 정형이 아닐 경우 (개인사업자/프리랜서 자유양식 등):
  · 같은 행 정렬, 협력사명 인접, 비고란 텍스트 등을 모두 활용해
    [item_name (직종/업무) + vendor + unit_price + quantity + amount] 을 한 묶음으로.
  · 캠페인명 (예: '삼성 비스포크 어댑테이션') 은 절대 item_name 으로 쓰지 말 것.
    실제 외주비 항목은 직종/역할 키워드 (성우/모델/녹음기사/편집/믹싱/PD 등).

[3단계 — 자가 보정 (Self-Correction)]
  · 만약 unit_price × quantity ≠ amount 면:
      → 비고/산출물/세부내역 컬럼에서 "녹음 3건", "1분 이상 3편" 같은 수량 단서 재탐색.
      → 텍스트의 정수를 quantity 로 역산. 그래도 안 맞으면 amount 를 정답으로 두고
        unit_price = amount / quantity 로 재계산.
  · quantity 가 0 또는 공란인데 amount 가 있으면 비고에서 같은 방식으로 역산.

[4단계 — 정제 룰]
  · 단위 텍스트 ('명/건/식/편/회/일/팀/개/인/시간/벌') 가 item_name 에 섞이면 제거.
  · vendor 가 라벨 셀의 값 (예: '업체명') 이면 다음 셀의 진짜 회사명 사용.
  · 금액 0/공란 행은 추출 금지.
  · 동일 의미 항목 (Editing↔편집, 프로듀싱료↔PD료, 녹음실비↔녹음) 은 한 번만, amount 합산.

[5단계 — 출력 구조 강제 (Reference Output 레이아웃)]
출력은 '외주비' 섹션 하나로만 채움. 정가항목과 대행수수료는 시스템이 따로 처리.
JSON 으로만 응답.
"""


def _build_system_prompt() -> str:
    bible = get_bible_text()
    layout = get_output_layout()
    parts = [_SYSTEM_PROMPT_CORE]
    if bible:
        parts.append(
            "[Ground Truth — 제작단가기준집 (참고용, 의도 분류에만 활용)]\n"
            "아래 정가 항목들은 input 파일에 등장하지 않음. 등장해도 무시.\n"
            f"```\n{bible[:6000]}\n```"
        )
    if layout:
        parts.append(
            "[Reference Output 레이아웃 — 외주비 섹션 출력 패턴]\n"
            f"```\n{layout[:2500]}\n```"
        )
    return "\n\n".join(parts)


# Backwards-compat for callers that read _SYSTEM_PROMPT directly
_SYSTEM_PROMPT = """\
(deprecated stub — actual prompt is built dynamically by _build_system_prompt())
당신은 한국 광고대행사가 받은 협력사 견적서를 분석해 라인아이템을 구조화 추출하는 분석가입니다.

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

    # STEP 1+2 — 청구 목적 분류 + 템플릿 엔티티 컨텍스트
    purpose_rep = classify_purpose(markdown)
    purpose_block = format_purpose_context_for_llm(purpose_rep)

    user_msg = (
        f"카테고리: {category_l1} / {category_l2}\n"
        f"처리 모드: {mode}\n\n"
        f"{purpose_block}\n\n"
        f"[소스 문서 markdown — 시작]\n"
        f"{markdown}\n"
        f"[소스 문서 markdown — 끝]\n\n"
        f"위 [STEP 1] 결과의 청구 목적과 [STEP 2] 표준 엔티티를 우선시하여, "
        f"의미가 매칭되는 라인아이템만 인용해 JSON 으로 추출하세요. "
        f"인접 셀의 텍스트를 임의로 끌어오지 말고, 의미 매칭이 명확한 것만 포함."
    )

    # messages.parse + Pydantic 으로 스키마 강제 + 자동 검증
    # cache_control 로 system prompt 캐싱 (반복 호출 시 ~90% 비용 절감)
    system_prompt = _build_system_prompt()
    response = client.messages.parse(
        model=settings.anthropic_model,            # .env 의 ANTHROPIC_MODEL (default: claude-sonnet-4-6)
        max_tokens=8192,
        system=[{
            "type": "text",
            "text": system_prompt,
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

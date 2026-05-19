"""
시맨틱 xlsx 파서 — 셀 좌표 의존을 폐기한 의도·문맥 중심 추출 엔진.

흐름:
  1) xlsx 전체 시트를 markdown 표로 1차 구조화 (pandas → to_markdown)
  2) bible_cache 의 제작단가기준집 + output.xlsx 레이아웃을 컨텍스트로 주입
  3) Claude 시맨틱 파서로 라인아이템 추출 (intent → entity grouping → self-correction)
  4) 결과를 EstimateRow 로 변환

호출 조건: ground_truth/column_profile 둘 다 매칭 실패하거나, 강제 사용 시.
"""
from __future__ import annotations
import io
import uuid
from typing import List, Optional

import pandas as pd

from ..config import settings
from ..schemas import EstimateRow


def xlsx_to_markdown(blob: bytes, max_rows_per_sheet: int = 80) -> str:
    """xlsx 전체를 markdown 표로 직렬화. tabulate 없으면 to_string fallback."""
    bio = io.BytesIO(blob)
    try:
        xl = pd.ExcelFile(bio)
    except Exception:
        bio.seek(0)
        xl = pd.ExcelFile(bio, engine="xlrd")
    parts: List[str] = []
    for sh in xl.sheet_names:
        df = pd.read_excel(xl, sheet_name=sh, header=None)
        # 완전 빈 행/열 제거 후 상위 N 행만
        df = df.dropna(how="all").dropna(how="all", axis=1).head(max_rows_per_sheet)
        if df.empty:
            continue
        df = df.fillna("")
        parts.append(f"## Sheet: {sh}")
        try:
            parts.append(df.to_markdown(index=False))
        except Exception:
            parts.append(df.to_string(index=False))
    return "\n\n".join(parts)


def semantic_parse_xlsx(
    blob: bytes,
    *,
    filename: str,
    category_l1: str,
    category_l2: str,
    mode: str = "precise",
) -> Optional[List[EstimateRow]]:
    """Claude 시맨틱 파서 호출. 키 없거나 호출 실패 시 None 반환 (호출자가 폴백)."""
    if not settings.anthropic_api_key:
        return None

    md = xlsx_to_markdown(blob)
    if not md.strip():
        return []

    # 기존 anthropic_markdown_to_rows 재사용 — bible 이 시스템 프롬프트에 자동 주입됨
    from .pdf_pipeline import anthropic_markdown_to_rows

    try:
        extracted = anthropic_markdown_to_rows(
            md,
            category_l1=category_l1,
            category_l2=category_l2,
            mode=mode,
        )
    except Exception:
        return None

    rows: List[EstimateRow] = []
    for r in extracted.rows:
        section = r.section
        item_name = r.item_name[:80]
        # 정가항목 ↔ 대행수수료 강등 (정책: input 파일에 없음)
        if category_l1 == "production" and section in ("정가항목", "대행수수료"):
            section = "외주비"
        rows.append(EstimateRow(
            id=f"r-sem-{uuid.uuid4().hex[:8]}",
            section=section,
            item_name=item_name,
            vendor=r.vendor,
            unit_price=r.unit_price,
            quantity=r.quantity,
            amount=r.amount,
            note=r.note,
        ))
    return rows

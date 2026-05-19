"""
BLUE NINE — 정가항목 단가표 + 행 생성기 (Source 16).

핵심 비즈니스 로직 (2026-05-19 재설계):
  · 정가항목은 협력사/외주사가 보내오는 input 견적서에는 존재하지 않는다.
  · AE 가 화면(Step 3)에서 '적용 건수' (applied_count) 를 수동으로 지정하면,
    백엔드가 이 단가표에 따라 6개 정가행을 즉시 생성해 외주비 옆에 합산한다.
  · 따라서 파서는 input 파일에서 정가항목을 검출하려고 하지 않는다 — 모든
    파싱 결과는 '외주비' 로 분류한다.
"""
from __future__ import annotations
import uuid
from typing import List, Tuple

from .schemas import EstimateRow

# 사내 제작단가기준집 (Source 16) — 항목명과 표준 단가 (단위: 원)
BASE_UNIT_PRICES: List[Tuple[str, float]] = [
    ("기획료",            3_000_000),
    ("카피료",            2_000_000),
    ("크리에이티브 워크료",  2_000_000),
    ("디렉션료",          2_000_000),
    ("자료조사비",         500_000),
    ("제작진행비",        1_200_000),
]

CANONICAL_JEONGGA_ITEMS: List[str] = [name for name, _ in BASE_UNIT_PRICES]


def generate_jeongga_rows(applied_count: int) -> List[EstimateRow]:
    """
    AE 가 지정한 적용 건수만큼 6개 정가항목을 EstimateRow 리스트로 생성.

    각 행: unit_price = 표준 단가, quantity = applied_count, amount = unit_price × quantity.
    applied_count == 0 이면 unit_price 만 표기되고 quantity / amount 는 0 으로 노출 (output.xlsx 동일).
    """
    n = max(0, int(applied_count))
    rows: List[EstimateRow] = []
    for name, unit in BASE_UNIT_PRICES:
        rows.append(EstimateRow(
            id=f"r-jeongga-{uuid.uuid4().hex[:6]}",
            section="정가항목",
            item_name=name,
            unit_price=unit,
            quantity=float(n),
            amount=unit * n,
            source_file="(정가항목 / AE 수동 입력)",
            note=f"표준 단가 × 적용 건수 {n}",
        ))
    return rows

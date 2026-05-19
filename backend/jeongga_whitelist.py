"""
BLUE NINE — 정가항목 단가표 + 행 생성기 (Source 16).

핵심 비즈니스 룰 (2026-05-19):
  · 정가항목은 협력사가 보내오는 input 견적서에 존재하지 않는다.
  · AE 가 Step 3 화면에서 applied_count 를 직접 설정한다.
  · 백엔드는 (category_l2, applied_count) 조합에 따라 표준 정가 세트를 자동 주입.

카테고리별 노출 정책:
  - video / print : 6개 정가행 노출 + qty=applied_count
  - radio / btl / other / media-* : 정가항목 미노출 (빈 리스트 반환)
"""
from __future__ import annotations
import uuid
from typing import Dict, List, Tuple

from .schemas import EstimateRow


# ─────────────────────────────────────────────────────────────
# 카테고리별 표준 단가 (Source 16)
# ─────────────────────────────────────────────────────────────
# 영상제작비 — 제작단가기준집 4p / 30p / 31p 의 핵심 5종만 정가항목으로 인정.
# (제작진행비는 정가항목이 아닌 예외 항목 → 외주비 또는 별도 항목으로 재분류, 2026-05-19)
_VIDEO_PRICES: List[Tuple[str, float]] = [
    ("기획료",            3_000_000),
    ("카피료",            2_000_000),
    ("크리에이티브 워크료",  2_000_000),
    ("디렉션료",          2_000_000),
    ("자료조사비",         500_000),
]

# 인쇄(KV) 제작비 — 영상보다 낮은 표준
_PRINT_PRICES: List[Tuple[str, float]] = [
    ("기획료",            2_000_000),
    ("카피료",            1_500_000),
    ("크리에이티브 워크료",  2_000_000),
    ("디렉션료",          1_500_000),
    ("자료조사비",         300_000),
]

# 카테고리(l2 key) → 정가 세트 매핑.
# 키가 없으면 정가 미적용 (radio / btl / other / 모든 media).
BASE_UNIT_PRICES_BY_CATEGORY: Dict[str, List[Tuple[str, float]]] = {
    "video": _VIDEO_PRICES,
    "print": _PRINT_PRICES,
}


def is_jeongga_applicable(category_l2: str) -> bool:
    """카테고리에 표준 정가 세트가 정의되어 있는지."""
    return category_l2 in BASE_UNIT_PRICES_BY_CATEGORY


def generate_jeongga_rows(category_l2: str, applied_count: int) -> List[EstimateRow]:
    """
    (category_l2, applied_count) → 정가항목 EstimateRow 리스트.

    - 카테고리에 정가 세트가 없으면 빈 리스트 반환 (UI 도 카운터 미노출).
    - 정가 세트가 있는 경우 항상 모든 행을 노출 (output.xlsx 패턴).
      applied_count=0 이면 qty=0 / amount=0 로 노출 (단가는 그대로 표기).
    """
    prices = BASE_UNIT_PRICES_BY_CATEGORY.get(category_l2)
    if not prices:
        return []
    n = max(0, int(applied_count))
    rows: List[EstimateRow] = []
    for name, unit in prices:
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


# Legacy export (다른 모듈에서 import 할 수 있음 — 이름만 유지)
CANONICAL_JEONGGA_ITEMS: List[str] = [name for name, _ in _VIDEO_PRICES]

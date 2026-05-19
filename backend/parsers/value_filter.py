"""
금액 기반 유효 데이터 필터 — 자율 추론 엔진 최상단 제약 (Skeleton Rule).

모든 파서 (휴리스틱 / 컬럼 프로파일 / Ground Truth / LLM 시맨틱) 가 동일하게 사용:
  · is_valid_amount(v)        : 0/공란/N/A/-/Null/NaN 등은 모두 무효
  · detect_value_column(df)   : 거래가격/견적금액/단가/합계금액 헤더 컬럼 자동 감지

원칙: 거래가격 칸이 비어있거나 0 인 행은 '빈 서식' 일 뿐 — 절대 추출 대상이 아님.
"""
from __future__ import annotations
import re
from typing import Optional

import pandas as pd


# 거래가격/금액성 열의 헤더 후보 (오른쪽 컬럼 우선 매칭)
VALUE_HEADER_KEYWORDS: tuple[str, ...] = (
    "거래가격", "거래 가격",
    "견적금액", "견적 금액", "견적가격",
    "합계금액", "합계 금액", "최종견적", "최종 견적",
    "청구금액", "청구 금액", "청구액",
    "단가",
    "금  액", "금액",                # '금 액' 같은 공백 변형 흡수
    "Amount", "Price", "Total",
)

# 텍스트로 입력된 무효 토큰 (lower-cased 비교)
INVALID_AMOUNT_TOKENS: tuple[str, ...] = (
    "", "-", "—", "–", "n/a", "na", "null", "none", "공란", ".",
)

# 표 헤더 행으로 인정하기 위한 동행 키워드 (단순 라벨 셀과 구분)
_TABLE_HEADER_HINTS: tuple[str, ...] = (
    "구분", "항목", "거래가격", "비고", "수량", "단가", "단위",
    "세부", "협력사", "협력회사", "산출물", "VAT", "견적", "금액",
    "Item", "Unit", "Qty", "Amount",
)

# 전화번호 / 사업자번호 / 우편번호 패턴 — 금액으로 오인하지 않음
_PHONE_LIKE_RE = re.compile(r"\d{2,4}[-\s\.]\d{2,4}[-\s\.]\d{3,4}")
# 단일 라인아이템의 합리적인 상한선 (1조 = 일조). 그 이상은 거의 항상 ID/오타.
_AMOUNT_UPPER_BOUND: float = 1_000_000_000_000.0


def is_valid_amount(v) -> bool:
    """금액이 유효한 청구 데이터인가?

    무효 케이스:
      · None / NaN / 빈 문자열 / '-' / '—' / 'N/A' / 'Null' / '0' / 음수
      · 전화번호·사업자번호 패턴 (예: 010-4998-2512, 820-86-01523)
      · 1조 초과 (현실적 라인아이템 금액 상한)
    유효 케이스: 정수/실수 > 0, 콤마/원이 섞인 깨끗한 양수 텍스트
    """
    if v is None:
        return False
    if isinstance(v, float) and v != v:           # NaN
        return False
    if isinstance(v, (int, float)):
        return 0 < v < _AMOUNT_UPPER_BOUND
    s = str(v).strip()
    if s.lower() in INVALID_AMOUNT_TOKENS:
        return False
    if _PHONE_LIKE_RE.search(s):                  # 전화번호/사업자번호 차단
        return False
    cleaned = re.sub(r"[^\d.\-]", "", s)
    if not cleaned or cleaned in ("-", ".", "-."):
        return False
    try:
        n = float(cleaned)
    except (ValueError, TypeError):
        return False
    return 0 < n < _AMOUNT_UPPER_BOUND


def coerce_positive_amount(v) -> Optional[float]:
    """is_valid_amount 통과한 값을 float 으로 반환. 무효면 None."""
    if not is_valid_amount(v):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    cleaned = re.sub(r"[^\d.\-]", "", str(v))
    try:
        return float(cleaned)
    except (ValueError, TypeError):
        return None


def detect_value_column(df: pd.DataFrame, max_scan_rows: int = 20) -> Optional[int]:
    """상위 N 행에서 '진짜 표 헤더' 행을 찾아 거래가격 컬럼 위치 반환.

    개선 (2026-05-19):
      · 단순 라벨 셀 (예: R5 의 '거래가격: 1500000') 에 속지 않도록,
        ≥2 개의 다른 헤더 힌트 ('구분/항목/비고/수량/단가' 등) 가 함께 있는
        행만 '표 헤더' 로 인정.
      · 후보 행 중 헤더 힌트 점수가 가장 높은 행의 value 컬럼을 채택.
    """
    rows = min(max_scan_rows, df.shape[0])
    best_col: Optional[int] = None
    best_score = 0
    for ridx in range(rows):
        row = df.iloc[ridx].tolist()
        # 행 안의 텍스트 셀 수집
        cells: list[tuple[int, str]] = []
        for cidx, v in enumerate(row):
            if v is None or (isinstance(v, float) and pd.isna(v)):
                continue
            s = str(v).strip()
            if s:
                cells.append((cidx, s))
        if len(cells) < 2:
            continue
        # 헤더 힌트 점수
        header_score = sum(
            1 for _, s in cells
            if any(h.replace(" ", "") in s.replace(" ", "") for h in _TABLE_HEADER_HINTS)
        )
        if header_score < 2:                       # 진짜 표 헤더가 아님
            continue
        # value column 후보 (오른쪽 우선)
        for cidx, s in sorted(cells, key=lambda x: -x[0]):
            if any(kw.replace(" ", "") in s.replace(" ", "") for kw in VALUE_HEADER_KEYWORDS):
                if header_score > best_score:
                    best_col = cidx
                    best_score = header_score
                break
    return best_col

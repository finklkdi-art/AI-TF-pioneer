"""
파일명 패턴별 컬럼 프로파일 — 영상/인쇄 표준 견적서 양식의 열 위치를 정확히 매핑.

각 프로파일은:
  · item_cols   : 항목명 텍스트가 들어가는 열 (조합해서 항목명 형성)
  · vendor_col  : 협력사명이 들어가는 열 (없으면 None)
  · num_cols    : 단가/수량/금액 후보가 들어가는 열들 (왼쪽 → 오른쪽 = 단가 → 수량 → 금액 순으로 해석)
  · qty_col     : 수량 전용 열 (있으면 num_cols 의 추론보다 우선)
  · amount_col  : 금액 전용 열 (있으면 num_cols 의 마지막값보다 우선)
  · skip_rows   : 헤더 위 스킵할 행 수 (대략적인 hint)

엑셀 열 표기 (A=0, B=1, C=2, D=3, E=4, F=5, G=6, H=7, I=8, J=9, K=10, ...).
"""
from __future__ import annotations
import re
from dataclasses import dataclass, field
from typing import List, Optional, Tuple


@dataclass
class ColumnProfile:
    key: str
    filename_patterns: Tuple[str, ...]
    item_cols: List[int]
    num_cols: List[int]               # 단가·수량·금액 후보 (왼쪽부터 단가, 마지막=금액)
    vendor_col: Optional[int] = None
    qty_col: Optional[int] = None
    amount_col: Optional[int] = None
    note_col: Optional[int] = None


# ─── 프로파일 정의 ──────────────────────────────────────────
# 사용자 가이드 (2026-05-19):
#   · 영상제작비견적서1 / input1   : 항목 A,B  / 단가·금액 D,E,F
#   · 영상제작비견적서2 / input2   : 항목 B,C,D / 단가·금액 H,I,J
#   · 인쇄제작비견적서1            : 항목 A,B  / 단가·금액 I,J,K
#   · 인쇄제작비견적서2 / input5   : 항목 C,D  / 단가·금액 E,F,G,H,I
PROFILES: List[ColumnProfile] = [
    ColumnProfile(
        key="video_template_1",
        filename_patterns=("영상제작비견적서1", "input1"),
        item_cols=[0, 1],            # A, B
        num_cols=[3, 4, 5],          # D, E, F (사전견적/협의견적 단계들)
        vendor_col=2,                # C (세부내역) — 협력사로 사용 가능한 경우
        amount_col=5,                # F = 최종견적
        note_col=7,
    ),
    ColumnProfile(
        key="video_template_2",
        filename_patterns=("영상제작비견적서2", "input2"),
        item_cols=[1, 2, 3],         # B, C, D
        num_cols=[7, 8, 9],          # H, I, J
        amount_col=9,
        note_col=10,
    ),
    ColumnProfile(
        key="print_template_1",
        filename_patterns=("인쇄제작비견적서1",),
        item_cols=[0, 1],            # A, B
        vendor_col=4,                # E (협력사 컬럼이 있는 경우)
        num_cols=[8, 9, 10],         # I, J, K
        amount_col=10,
        note_col=11,
    ),
    ColumnProfile(
        key="print_template_2",
        filename_patterns=("인쇄제작비견적서2", "input5"),
        item_cols=[2, 3],            # C, D
        num_cols=[4, 5, 6, 7, 8],    # E, F, G, H, I
        amount_col=8,
    ),
]


def match_profile(filename: str) -> Optional[ColumnProfile]:
    """파일명에 가장 잘 맞는 프로파일을 반환. 매칭 없으면 None."""
    if not filename:
        return None
    base = filename.rsplit("/", 1)[-1].rsplit("\\", 1)[-1]
    base_lower = base.lower()
    # 가장 긴 패턴 우선 매칭
    for prof in sorted(PROFILES, key=lambda p: -max(len(pat) for pat in p.filename_patterns)):
        for pat in prof.filename_patterns:
            if pat in base or pat.lower() in base_lower:
                return prof
    return None

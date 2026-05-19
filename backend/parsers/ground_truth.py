"""
Ground Truth 셀 좌표 추출기 — 부서 표준 견적서 양식의 정확한 시트/셀을 알고 있는
특정 파일에 대해 칼같은 데이터 추출을 수행. 매칭 안 되는 파일은 column_profiles.py
→ 휴리스틱 순으로 폴백.

표준 input 5종 (Ground Truth):
  1) input1.xlsx — 시트 '포스트프로덕션'
       H3=공급자/회사명, A14=구분, B14=항목, F14=거래가격, H14=수량
  2) input2.xlsx — 시트 '녹음 표준견적서'
       J3=공급자, A12=구분, B12=항목, G12=거래가격, I12=비고('녹음3' 등에서 정수 추출 → 수량)
  3) input3.xlsx — 시트 '견적서 (개인사업자)'
       B19=업체명, F12=단가, E12=수량, G12=금액
  4) input4.pdf — input3 의 PDF 변환본 (LLM 파이프라인 사용)
  5) input5.xlsx — 단일 시트
       A=구분, B=항목, C=협력회사, D=사전견적, E=협의견적, G/H=비고(수량)
"""
# -*- coding: utf-8 -*-
from __future__ import annotations
import io
import re
import uuid
from dataclasses import dataclass, field
from typing import Callable, Dict, List, Optional, Tuple
import pandas as pd

from ..schemas import EstimateRow


# ─── 셀 좌표 ───────────────────────────────────────────────
def cell_addr_to_indices(addr: str) -> Tuple[int, int]:
    """엑셀 셀 주소 (예: 'H3') → (col_idx=7, row_idx=2) 0-인덱스."""
    m = re.match(r"^([A-Z]+)(\d+)$", addr.strip().upper())
    if not m:
        raise ValueError(f"invalid cell address: {addr!r}")
    col_letters, row_num = m.group(1), int(m.group(2))
    col = 0
    for ch in col_letters:
        col = col * 26 + (ord(ch) - ord("A") + 1)
    return col - 1, row_num - 1


def _get_cell(df: pd.DataFrame, addr: str):
    c, r = cell_addr_to_indices(addr)
    if r >= df.shape[0] or c >= df.shape[1]:
        return None
    v = df.iat[r, c]
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    return v


def _to_num(v) -> Optional[float]:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = re.sub(r"[^\d.\-]", "", str(v))
    if s in ("", "-", ".", "-."):
        return None
    try:
        return float(s)
    except ValueError:
        return None


def _qty_from_text(v) -> Optional[float]:
    """텍스트에 섞인 수량('녹음3, 효과 등' → 3) 추출. 숫자형이면 그대로."""
    if v is None:
        return None
    if isinstance(v, (int, float)) and not (isinstance(v, float) and pd.isna(v)):
        return float(v)
    s = str(v)
    nums = re.findall(r"\d+", s)
    if not nums:
        return None
    return float(nums[0])


def _strip_unit_text(s: Optional[str]) -> str:
    if not s:
        return ""
    out = str(s).strip()
    # 항목명 칸에 단위 텍스트가 섞여 들어온 경우 제거
    out = re.sub(r"\b\d+\s*(?:명|건|식|편|회|일|팀|개|인|시간|벌)\b", "", out)
    out = re.sub(r"(?<![가-힣A-Za-z])(?:명|건|식|편|회|일|팀|개|인|시간|벌)(?![가-힣A-Za-z])", "", out)
    out = re.sub(r"\s+", " ", out).strip(" /·,;:-—")
    return out


# ─── 프로파일 정의 ──────────────────────────────────────────
@dataclass
class GroundTruthProfile:
    """단일 시트에서 정해진 셀을 단일 행으로 추출하는 프로파일."""
    key: str
    filename_patterns: Tuple[str, ...]
    sheet_filter: str                       # 부분일치
    vendor_cell: Optional[str] = None
    section_cell: Optional[str] = None
    item_cell: Optional[str] = None
    unit_price_cell: Optional[str] = None
    quantity_cell: Optional[str] = None
    quantity_from_text: bool = False        # 셀이 텍스트면 첫 정수를 수량으로
    amount_cell: Optional[str] = None
    fallback_qty: float = 1.0


GROUND_TRUTH_PROFILES: List[GroundTruthProfile] = [
    GroundTruthProfile(
        key="input1_post",
        filename_patterns=("input1",),
        sheet_filter="포스트프로덕션",
        vendor_cell="H3", section_cell="A14", item_cell="B14",
        amount_cell="F14", quantity_cell="H14", quantity_from_text=True,
    ),
    GroundTruthProfile(
        key="input2_recording",
        filename_patterns=("input2",),
        sheet_filter="녹음 표준견적서",
        vendor_cell="J3", section_cell="A12", item_cell="B12",
        amount_cell="G12", quantity_cell="I12", quantity_from_text=True,
    ),
    GroundTruthProfile(
        key="input3_individual",
        filename_patterns=("input3",),
        sheet_filter="견적서 (개인사업자)",
        # B19 는 '업체명' 라벨 셀이므로 실제 vendor 값은 C19 에서 추출.
        vendor_cell="C19", item_cell="C12",         # 업무내용을 항목으로 사용
        unit_price_cell="F12", quantity_cell="E12", amount_cell="G12",
    ),
]


def find_profile(filename: str) -> Optional[GroundTruthProfile]:
    if not filename:
        return None
    base = filename.rsplit("/", 1)[-1].rsplit("\\", 1)[-1].lower()
    for prof in GROUND_TRUTH_PROFILES:
        for pat in prof.filename_patterns:
            if pat.lower() in base:
                return prof
    return None


def _select_sheet(xl: pd.ExcelFile, name_hint: str) -> Optional[str]:
    for sh in xl.sheet_names:
        if name_hint in sh:
            return sh
    return None


def extract_with_ground_truth(blob: bytes, filename: str) -> Optional[List[EstimateRow]]:
    """매칭되면 EstimateRow 리스트 반환, 아니면 None."""
    prof = find_profile(filename)
    if prof is None:
        return None

    bio = io.BytesIO(blob)
    try:
        xl = pd.ExcelFile(bio)
    except Exception:
        bio.seek(0)
        xl = pd.ExcelFile(bio, engine="xlrd")

    sheet = _select_sheet(xl, prof.sheet_filter)
    if sheet is None:
        return []      # 매칭됐지만 시트 없음 — 빈 결과로 명시

    df = pd.read_excel(xl, sheet_name=sheet, header=None)

    vendor = str(_get_cell(df, prof.vendor_cell) or "").strip() if prof.vendor_cell else None
    section_label = str(_get_cell(df, prof.section_cell) or "").strip() if prof.section_cell else ""
    raw_item = _get_cell(df, prof.item_cell) if prof.item_cell else None
    item_name = _strip_unit_text(str(raw_item) if raw_item else "")

    if not item_name and section_label:
        item_name = section_label[:60]
    if not item_name:
        return []

    amount = _to_num(_get_cell(df, prof.amount_cell)) if prof.amount_cell else None
    unit_price = _to_num(_get_cell(df, prof.unit_price_cell)) if prof.unit_price_cell else None

    qty: Optional[float] = None
    if prof.quantity_cell:
        raw_q = _get_cell(df, prof.quantity_cell)
        qty = _qty_from_text(raw_q) if prof.quantity_from_text else _to_num(raw_q)
    if qty is None or qty == 0:
        qty = prof.fallback_qty

    if amount is None and unit_price is not None:
        amount = unit_price * qty
    if amount is None or amount == 0:
        return []    # 0/공란 → 노출 안 함

    row = EstimateRow(
        id=f"r-gt-{uuid.uuid4().hex[:6]}",
        section="외주비",
        item_name=(section_label + " · " + item_name) if section_label and section_label != item_name else item_name,
        vendor=vendor[:60] if vendor else None,
        unit_price=unit_price if unit_price is not None else (amount / qty if qty else amount),
        quantity=qty,
        amount=amount,
    )
    return [row]


# ─── input5 (단일 시트, 컬럼 기반) — 마스터 규칙 ─────────────
def extract_input5(blob: bytes, filename: str) -> Optional[List[EstimateRow]]:
    """input5.xlsx 처럼 단일 시트의 정해진 컬럼 매핑.
    A=구분, B=항목, C=협력회사, D=사전견적, E=협의견적, G/H=비고
    """
    if not filename or "input5" not in filename.lower():
        return None
    bio = io.BytesIO(blob)
    xl = pd.ExcelFile(bio)
    df = pd.read_excel(xl, sheet_name=xl.sheet_names[0], header=None)
    HEADER_NOISE = (
        "JOB NO", "JOB 명", "제작 CD", "프로덕션 :", "촬영일자", "온에어일",
        "구분", "ART:", "CW :", "AE:", "감독 :", "PD :", "PM :", "실행예산",
        "완료견적", "첨부",
    )
    def _str_or_blank(v):
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return ""
        s = str(v).strip()
        return "" if s.lower() == "nan" else s

    out: List[EstimateRow] = []
    for ridx in range(df.shape[0]):
        row = df.iloc[ridx].tolist()
        sec    = _str_or_blank(row[0]) if 0 < len(row) else ""
        item   = _str_or_blank(row[1]) if 1 < len(row) else ""
        vendor = _str_or_blank(row[2]) if 2 < len(row) else ""
        pre    = _to_num(row[3]) if 3 < len(row) else None
        cur    = _to_num(row[4]) if 4 < len(row) else None
        qty_text = ""
        for ci in (6, 7):
            if ci < len(row):
                q = _str_or_blank(row[ci])
                if q: qty_text += " " + q
        qty = _qty_from_text(qty_text) or 1.0
        joined = (sec + " " + item).strip()
        # 헤더/메타 행 + 합계 행 강제 제외
        if any(k in joined for k in HEADER_NOISE):
            continue
        if any(k in joined for k in ("외주항목계", "합계", "총계", "소계", "총합")):
            continue
        # 금액: 협의 우선, 없으면 사전
        amount = cur if (cur is not None and cur != 0) else pre
        if amount is None or amount == 0:
            continue
        name_parts = []
        if sec:  name_parts.append(_strip_unit_text(sec))
        if item: name_parts.append(_strip_unit_text(item))
        name = " · ".join([p for p in name_parts if p])
        if not name:
            continue
        out.append(EstimateRow(
            id=f"r-gt5-{uuid.uuid4().hex[:6]}",
            section="외주비",
            item_name=name[:80],
            vendor=vendor[:60] if vendor else None,
            unit_price=amount / qty if qty else amount,
            quantity=qty,
            amount=amount,
        ))
    return out


def extract_ground_truth(blob: bytes, filename: str) -> Optional[List[EstimateRow]]:
    """모든 ground truth 추출기 시도 — 첫 매칭이 결과 반환."""
    if not filename:
        return None
    out = extract_with_ground_truth(blob, filename)
    if out is not None:
        return out
    out = extract_input5(blob, filename)
    if out is not None:
        return out
    return None

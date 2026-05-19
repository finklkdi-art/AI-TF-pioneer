"""
파서 베이스.
- 입력: 협력사 견적서 (xlsx/xls) bytes
- 출력: List[EstimateRow]  (정가/외주비/대행수수료 섹션 라벨링 포함)
- 더블체크: row.unit_price * row.quantity ≈ row.amount  (Source 28, 29, 32)
"""
from __future__ import annotations
import io
import re
import uuid
from typing import List, Tuple, Dict, Any, Optional
import pandas as pd

from ..schemas import EstimateRow
from .column_profiles import ColumnProfile, match_profile
from .ground_truth import extract_ground_truth
from .value_filter import is_valid_amount, detect_value_column


# 섹션 키워드 사전 — Source 15, 17
SECTION_KEYWORDS: Dict[str, List[str]] = {
    "정가항목": ["정가합계", "기본정가", "정가", "AGENCY", "기획료"],
    "외주비": ["외주비", "외주비합계", "외주", "협력사", "PD", "촬영연출", "POST프로덕션",
              "음향", "녹음", "BGM", "디자인", "인쇄", "후반작업"],
    "대행수수료": ["대행수수료", "대행료", "수수료", "AGENCY FEE", "Agency Fee"],
}

# 매체비 섹션 (Source 17) — 매체별 청구액/지급액/수수료
MEDIA_SECTION_KEYWORDS: Dict[str, List[str]] = {
    "매체청구액": ["청구액", "광고주청구", "GROSS", "Gross"],
    "매체지급액": ["지급액", "매체사지급", "NET", "Net"],
    "매체수수료": ["매체수수료", "수수료", "AGCY", "Commission"],
}


def _normalize(s: Any) -> str:
    if s is None:
        return ""
    return re.sub(r"\s+", "", str(s))


_PHONE_RE = re.compile(r"\d{2,4}[-\s\.]\d{2,4}[-\s\.]\d{3,4}")


def _to_num(v: Any) -> Optional[float]:
    """텍스트에서 양수 추출. 전화번호/사업자번호 패턴은 거부."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    raw = str(v).strip()
    # 전화번호/사업자번호 — 금액 아님
    if _PHONE_RE.search(raw):
        return None
    s = re.sub(r"[^\d.\-]", "", raw)
    if s in ("", "-", ".", "-."):
        return None
    try:
        return float(s)
    except ValueError:
        return None


def _read_workbook(blob: bytes) -> Dict[str, pd.DataFrame]:
    bio = io.BytesIO(blob)
    # xls/xlsx 자동 분기
    try:
        xl = pd.ExcelFile(bio)
    except Exception:
        bio.seek(0)
        xl = pd.ExcelFile(bio, engine="xlrd")
    return {sh: pd.read_excel(xl, sheet_name=sh, header=None) for sh in xl.sheet_names}


def _classify_section(row_text: str, *, media: bool = False) -> Optional[str]:
    if not row_text:
        return None
    dictionary = MEDIA_SECTION_KEYWORDS if media else SECTION_KEYWORDS
    for sect, kws in dictionary.items():
        for kw in kws:
            if kw and _normalize(kw) in _normalize(row_text):
                return sect
    return None


def _scan_rows(df: pd.DataFrame, *, media: bool = False) -> List[EstimateRow]:
    """행 단위 휴리스틱 추출.

    [최상단 룰 — 2026-05-19 금액 유효성 필터]
      거래가격/견적금액/단가 헤더 컬럼을 자동 감지해, 해당 칸이 '유효 금액'
      (None / NaN / 0 / 공란 / N/A / '-' 등 모두 무효) 인 행만 추출.
      매칭 없으면 last-numeric fallback. 어느 쪽이든 amount > 0 인 행만 통과.

    [Section 분류 규칙 — 2026-05-19 화이트리스트 강제]
      · production: 모든 행 '외주비' (정가는 시스템 자동 주입)
      · media     : 헤더 키워드 기반 (매체청구액/지급액/수수료)
    """
    # 최상단 — value column 감지 (오른쪽 컬럼 우선)
    value_col = detect_value_column(df)

    out: List[EstimateRow] = []
    current_section = "매체청구액" if media else None    # production 은 행별로 결정

    for ridx in range(df.shape[0]):
        row = df.iloc[ridx]
        cells = [c for c in row.tolist() if c is not None and not (isinstance(c, float) and pd.isna(c))]
        if not cells:
            continue
        # media 카테고리는 헤더 기반 섹션 유지
        if media:
            joined = " ".join(str(c) for c in cells)
            sect = _classify_section(joined, media=True)
            if sect:
                current_section = sect

        # 최상단 게이트 — STRICT: value_col 이 감지됐다면 그 컬럼만 유일한 진실.
        # value_col 이 비어있거나 무효이면 그 행 자체를 즉시 건너뜀.
        row_list = row.tolist()
        primary_amount = None
        if value_col is not None:
            if value_col >= len(row_list):
                continue
            if not is_valid_amount(row_list[value_col]):
                continue                              # 빈 서식 행 — 추출 안 함
            primary_amount = _to_num(row_list[value_col])
        else:
            # value_col 미감지 — last-numeric fallback. 적어도 한 개 유효 금액 있어야.
            nums_pre = [_to_num(v) for v in row_list if _to_num(v) is not None]
            if not nums_pre or not any(is_valid_amount(n) for n in nums_pre):
                continue

        # 숫자 후보 (지금은 통과한 행에 대해서만 계산)
        nums = [(j, _to_num(v)) for j, v in enumerate(row_list) if _to_num(v) is not None]
        if len(nums) < 1:
            continue

        # name = 가장 앞쪽의 텍스트 셀들 합치기 (최대 3개)
        text_cells = []
        for j, v in enumerate(row.tolist()):
            if _to_num(v) is not None:
                continue
            if v is None or (isinstance(v, float) and pd.isna(v)):
                continue
            s = str(v).strip()
            if not s:
                continue
            text_cells.append(s)
        if not text_cells:
            continue
        name = " / ".join(text_cells[:3])

        # 합계/소계/총계 행은 건너뜀
        if any(k in name for k in ("합계", "총계", "소계", "총합", "Total", "TOTAL")):
            continue

        # amount 결정: value_col 우선, 없으면 last numeric
        if primary_amount is not None and primary_amount > 0:
            amount = primary_amount
            # value_col 좌측의 숫자들을 단가/수량 후보로
            left_nums = [n for j, n in nums if j < (value_col or 999)]
        else:
            amount = nums[-1][1] or 0.0
            left_nums = [n for _, n in nums[:-1]]
        unit_price = 0.0
        quantity = 1.0
        if len(left_nums) >= 2:
            unit_price = left_nums[-2] or 0.0
            quantity = left_nums[-1] or 1.0
        elif len(left_nums) == 1:
            unit_price = left_nums[0] or 0.0

        # ── 최종 게이트 — amount 가 0 이면 절대 노출 안 함
        if not amount or amount <= 0:
            continue

        # ── 섹션 결정 ─────────────────────────────────────────────
        # 2026-05-19 비즈니스 룰: 정가항목은 input 파일에 존재하지 않음.
        # AE 가 별도로 applied_count 를 입력하면 backend.main 이 정가행을 주입.
        # 따라서 파서는 production 카테고리에서 모든 추출 행을 '외주비' 로 분류.
        if media:
            section = current_section or "매체청구액"
        else:
            section = "외주비"
        item_name_out = name[:80]

        out.append(EstimateRow(
            id=f"r-{uuid.uuid4().hex[:8]}",
            section=section,
            item_name=item_name_out,
            unit_price=unit_price,
            quantity=quantity,
            amount=amount,
        ))
    return out


def _scan_rows_with_profile(df: pd.DataFrame, profile: ColumnProfile) -> List[EstimateRow]:
    """프로파일 기반 컬럼 매핑 추출.

    각 행에서:
      - item_name = profile.item_cols 의 텍스트 셀 join
      - vendor    = profile.vendor_col 값 (있으면)
      - amount    = profile.amount_col 값 (있으면) / 없으면 profile.num_cols 마지막 숫자
      - unit_price/quantity = num_cols 에서 amount 앞 두 후보
    """
    out: List[EstimateRow] = []
    for ridx in range(df.shape[0]):
        row = df.iloc[ridx].tolist()
        # item_name 합성 — text only
        name_parts: List[str] = []
        for ci in profile.item_cols:
            if ci < len(row):
                v = row[ci]
                if v is None or (isinstance(v, float) and pd.isna(v)):
                    continue
                s = str(v).strip()
                if not s:
                    continue
                # 항목명 칸에 숫자만 들어있으면 스킵
                if _to_num(s) is not None:
                    continue
                name_parts.append(s)
        if not name_parts:
            continue
        name = " / ".join(name_parts)[:80]

        # vendor (선택)
        vendor: Optional[str] = None
        if profile.vendor_col is not None and profile.vendor_col < len(row):
            v = row[profile.vendor_col]
            if v is not None and not (isinstance(v, float) and pd.isna(v)):
                s = str(v).strip()
                if s and _to_num(s) is None:
                    vendor = s[:40]

        # 금액 결정 — amount_col 우선, 아니면 num_cols 마지막
        amount: float = 0.0
        if profile.amount_col is not None and profile.amount_col < len(row):
            a = _to_num(row[profile.amount_col])
            if a is not None:
                amount = a
        if amount == 0:
            for ci in reversed(profile.num_cols):
                if ci < len(row):
                    a = _to_num(row[ci])
                    if a is not None and a != 0:
                        amount = a
                        break

        # 수량 / 단가 후보
        quantity: float = 1.0
        unit_price: float = 0.0
        if profile.qty_col is not None and profile.qty_col < len(row):
            q = _to_num(row[profile.qty_col])
            if q is not None:
                quantity = q
        # num_cols 에서 amount 가 아닌 두 후보 → 첫 = 단가, 둘째 = 수량
        other_nums = []
        for ci in profile.num_cols:
            if ci >= len(row):
                continue
            n = _to_num(row[ci])
            if n is None:
                continue
            if profile.amount_col is not None and ci == profile.amount_col:
                continue
            other_nums.append(n)
        if other_nums:
            unit_price = other_nums[0]
            if len(other_nums) > 1 and profile.qty_col is None:
                quantity = other_nums[1] if other_nums[1] else quantity

        # 최상단 게이트 — 유효 금액이 아니면 노출 안 함
        if not is_valid_amount(amount):
            continue
        out.append(EstimateRow(
            id=f"r-{uuid.uuid4().hex[:8]}",
            section="외주비",       # 프로파일은 production input 전제 — 모두 외주비
            item_name=name,
            vendor=vendor,
            unit_price=unit_price,
            quantity=quantity,
            amount=amount,
        ))
    return out


class BaseParser:
    label = "generic"
    is_media = False

    def parse(self, blob: bytes, filename: Optional[str] = None) -> List[EstimateRow]:
        # 0) Ground Truth 우선 — input1~5 처럼 정확한 좌표를 아는 파일은 칼같이 추출
        if filename and not self.is_media:
            gt = extract_ground_truth(blob, filename)
            if gt is not None:
                return gt
        # 1) 파일명 + 컬럼 프로파일 매칭 (영상/인쇄 표준 견적서)
        profile = match_profile(filename) if (filename and not self.is_media) else None
        sheets = _read_workbook(blob)
        rows: List[EstimateRow] = []
        if profile is not None:
            for _, df in sheets.items():
                rows.extend(_scan_rows_with_profile(df, profile))
        else:
            # 2) 휴리스틱 폴백 (다양한 양식 대응)
            for _, df in sheets.items():
                rows.extend(_scan_rows(df, media=self.is_media))
        return rows


class GenericProductionParser(BaseParser):
    label = "production-generic"
    is_media = False


class GenericMediaParser(BaseParser):
    label = "media-generic"
    is_media = True

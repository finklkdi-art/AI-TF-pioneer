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


def _to_num(v: Any) -> Optional[float]:
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

    [Section 분류 규칙 — 2026-05-19 화이트리스트 강제]
      · production 카테고리:
          item_name 이 backend/jeongga_whitelist 의 6개 항목과 매칭되면 '정가항목',
          그 외 모든 행은 '외주비' 로 격리.
          → 섹션 헤더('정가합계', '외주비합계' 등) 키워드는 더 이상 분류에 영향을 주지 않음.
      · media 카테고리:
          기존 헤더 휴리스틱 유지 — '매체청구액 / 매체지급액 / 매체수수료'.
    """
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

        # 숫자 후보
        nums = [(j, _to_num(v)) for j, v in enumerate(row.tolist()) if _to_num(v) is not None]
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

        # 마지막 숫자를 amount, 그 앞을 unit_price/quantity 후보로
        amount = nums[-1][1] or 0.0
        unit_price = 0.0
        quantity = 1.0
        if len(nums) >= 3:
            unit_price = nums[-3][1] or 0.0
            quantity = nums[-2][1] or 1.0
        elif len(nums) == 2:
            unit_price = nums[-2][1] or 0.0
            quantity = 1.0

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
        sheets = _read_workbook(blob)
        # 1) 파일명으로 프로파일 매칭 시도 (production 만)
        profile = match_profile(filename) if (filename and not self.is_media) else None
        rows: List[EstimateRow] = []
        if profile is not None:
            # 모든 시트에 동일 프로파일 적용 (대체로 1~2 시트가 데이터)
            for _, df in sheets.items():
                rows.extend(_scan_rows_with_profile(df, profile))
        else:
            for _, df in sheets.items():
                rows.extend(_scan_rows(df, media=self.is_media))
        return rows


class GenericProductionParser(BaseParser):
    label = "production-generic"
    is_media = False


class GenericMediaParser(BaseParser):
    label = "media-generic"
    is_media = True

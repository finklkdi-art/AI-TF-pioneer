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
    """단순 휴리스틱 추출:
    각 행에서 (이름성 텍스트, 수량, 단가, 금액)을 찾아 EstimateRow 생성.
    """
    out: List[EstimateRow] = []
    current_section = "외주비" if not media else "매체청구액"
    for ridx in range(df.shape[0]):
        row = df.iloc[ridx]
        cells = [c for c in row.tolist() if c is not None and not (isinstance(c, float) and pd.isna(c))]
        if not cells:
            continue
        # 섹션 헤더 감지
        joined = " ".join(str(c) for c in cells)
        sect = _classify_section(joined, media=media)
        if sect:
            current_section = sect
        # 숫자 후보
        nums = [(j, _to_num(v)) for j, v in enumerate(row.tolist()) if _to_num(v) is not None]
        if len(nums) < 1:
            continue
        # name = 가장 앞쪽의 텍스트 셀들 합치기
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
        # 합계로 보이는 행은 건너뜀 (별도 누적)
        if any(k in name for k in ("합계", "총계", "소계", "총합", "Total", "TOTAL")):
            continue
        # 마지막 숫자를 amount, 그 앞을 unit_price/quantity 후보로
        nums_sorted = nums
        amount = nums_sorted[-1][1] or 0.0
        unit_price = 0.0
        quantity = 1.0
        if len(nums_sorted) >= 3:
            unit_price = nums_sorted[-3][1] or 0.0
            quantity = nums_sorted[-2][1] or 1.0
        elif len(nums_sorted) == 2:
            unit_price = nums_sorted[-2][1] or 0.0
            quantity = 1.0
        out.append(EstimateRow(
            id=f"r-{uuid.uuid4().hex[:8]}",
            section=current_section,
            item_name=name[:80],
            unit_price=unit_price,
            quantity=quantity,
            amount=amount,
        ))
    return out


class BaseParser:
    label = "generic"
    is_media = False

    def parse(self, blob: bytes) -> List[EstimateRow]:
        sheets = _read_workbook(blob)
        rows: List[EstimateRow] = []
        for _, df in sheets.items():
            rows.extend(_scan_rows(df, media=self.is_media))
        return rows


class GenericProductionParser(BaseParser):
    label = "production-generic"
    is_media = False


class GenericMediaParser(BaseParser):
    label = "media-generic"
    is_media = True

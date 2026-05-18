"""
Billing Status 파일 파서 (Source 20).
매체비 견적서와 1:1 매핑을 위해 매체사명 + 캠페인명 + 집행월 키 추출.
"""
from __future__ import annotations
import io
from typing import List, Dict, Any
import pandas as pd

KEY_HEADER_HINTS = {
    "media":    ["매체사", "매체사명", "매체", "Vendor", "Media"],
    "campaign": ["캠페인", "캠페인명", "Campaign", "광고주", "Brand", "세부 내역", "세부내역", "Detail"],
    "month":    ["집행월", "Month", "기간", "Billing Date", "요청 일자", "요청일자"],
    "charged":  ["청구 금액", "청구금액", "청구 합계", "청구액", "Gross", "광고주청구"],
    "paid":     ["지급액", "Net", "매체사지급", "지급 금액", "지급금액"],
    "fee_rate": ["수수료율", "Fee", "Commission"],
    "vat":      ["VAT", "부가세"],
    "total":    ["청구 합계", "청구합계", "Total", "합계"],
}


def _norm(s: Any) -> str:
    return "" if s is None else str(s).strip()


def _detect_header_row(df: pd.DataFrame, max_scan: int = 10) -> int:
    for i in range(min(max_scan, df.shape[0])):
        row = [_norm(c) for c in df.iloc[i].tolist()]
        hits = 0
        for hints in KEY_HEADER_HINTS.values():
            if any(any(h in c for h in hints) for c in row if c):
                hits += 1
        if hits >= 3:
            return i
    return 0


def _to_float(v: Any) -> float:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return 0.0
    s = str(v).replace(",", "").strip()
    if not s:
        return 0.0
    try:
        return float(s)
    except (ValueError, TypeError):
        return 0.0


def parse_billing(blob: bytes) -> List[Dict[str, Any]]:
    bio = io.BytesIO(blob)
    try:
        xl = pd.ExcelFile(bio)
    except Exception:
        bio.seek(0)
        xl = pd.ExcelFile(bio, engine="xlrd")
    rows: List[Dict[str, Any]] = []
    for sh in xl.sheet_names:
        df = pd.read_excel(xl, sheet_name=sh, header=None)
        if df.empty:
            continue
        hr = _detect_header_row(df)
        headers = [_norm(c) for c in df.iloc[hr].tolist()]
        col_map: Dict[str, int] = {}
        for key, hints in KEY_HEADER_HINTS.items():
            for j, h in enumerate(headers):
                if any(_norm(hint) in h for hint in hints):
                    col_map[key] = j
                    break
        # charged 또는 paid 또는 total 중 어느 하나라도 잡혀야 데이터 행으로 인식
        if not any(k in col_map for k in ("charged", "paid", "total")):
            continue
        for i in range(hr + 1, df.shape[0]):
            r = df.iloc[i].tolist()
            def pick(k: str):
                j = col_map.get(k)
                return r[j] if j is not None and j < len(r) else None
            # 합계/Total 행은 건너뜀
            first_text = " ".join(str(c) for c in r[:3] if pd.notna(c))
            if any(k in first_text for k in ("Total", "TOTAL", "합계", "총계", "소계")):
                continue
            charged_f = _to_float(pick("charged"))
            paid_f = _to_float(pick("paid"))
            total_f = _to_float(pick("total"))
            # charged 가 비어 있으면 total - VAT 로 폴백
            if charged_f == 0 and total_f > 0:
                vat_f = _to_float(pick("vat"))
                charged_f = total_f - vat_f
            # paid 가 없으면 charged 와 동일하게 취급 (이 파일 유형은 수수료 분리 없음)
            if paid_f == 0 and charged_f > 0:
                paid_f = charged_f
            if charged_f == 0 and paid_f == 0:
                continue
            rows.append({
                "media": _norm(pick("media")),
                "campaign": _norm(pick("campaign")),
                "month": _norm(pick("month")),
                "charged": charged_f,
                "paid": paid_f,
                "fee_rate": pick("fee_rate"),
                "sheet": sh,
            })
    return rows

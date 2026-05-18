"""
Billing Status 파일 파서 (Source 20).
매체비 견적서와 1:1 매핑을 위해 매체사명 + 캠페인명 + 집행월 키 추출.
"""
from __future__ import annotations
import io
from typing import List, Dict, Any
import pandas as pd

KEY_HEADER_HINTS = {
    "media": ["매체사", "매체", "매체사명", "Vendor", "Media"],
    "campaign": ["캠페인", "캠페인명", "Campaign", "광고주", "Brand"],
    "month": ["집행월", "월", "Month", "기간"],
    "charged": ["청구액", "Gross", "광고주청구"],
    "paid": ["지급액", "Net", "매체사지급"],
    "fee_rate": ["수수료율", "Fee", "Commission"],
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
        hr = _detect_header_row(df)
        headers = [_norm(c) for c in df.iloc[hr].tolist()]
        col_map: Dict[str, int] = {}
        for key, hints in KEY_HEADER_HINTS.items():
            for j, h in enumerate(headers):
                if any(hint in h for hint in hints):
                    col_map[key] = j
                    break
        if "charged" not in col_map and "paid" not in col_map:
            continue
        for i in range(hr + 1, df.shape[0]):
            r = df.iloc[i].tolist()
            def pick(k: str):
                j = col_map.get(k)
                return r[j] if j is not None and j < len(r) else None
            charged = pick("charged")
            paid = pick("paid")
            try:
                charged_f = float(str(charged).replace(",", "")) if charged not in (None, "") else 0.0
            except (ValueError, TypeError):
                charged_f = 0.0
            try:
                paid_f = float(str(paid).replace(",", "")) if paid not in (None, "") else 0.0
            except (ValueError, TypeError):
                paid_f = 0.0
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

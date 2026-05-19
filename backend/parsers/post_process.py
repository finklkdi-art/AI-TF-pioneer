"""
파싱 결과 후처리 — 5가지 정리 작업:
  1) 단위 노이즈 제거       : '명/건/식/편/회/일/팀/원' 등이 항목명 끝에 붙은 경우 제거.
  2) 합계/소계 행 강제 배제 : 항목명에 '합계 / 소계 / 총계 / VAT / 부가세' 포함 시 drop.
  3) 동의어 정규화          : '프로듀싱료' ↔ 'PD료', 'Editing' ↔ '편집' 등을 canonical 로 통일.
  4) (name, vendor) 키 dedup: 같은 행이 여러 파일에서 추출됐다면 amount 합산.
  5) 금액 0 행 제거         : amount 가 0 또는 빈 값인 행은 노출 안 함.
"""
from __future__ import annotations
import re
from typing import List, Dict, Tuple

from ..schemas import EstimateRow


# ─── 단위 노이즈 패턴 ──────────────────────────────────────
# 항목명 끝/중간에 단위만 떠있거나 숫자 뒤에 단위가 붙은 경우 제거.
# 한국어 단위: 명, 건, 식, 편, 회, 일, 팀, 원, 개, 시간, 인 등.
_UNIT_TOKEN_RE = re.compile(
    r"(?<![가-힣A-Za-z])(?:명|건|식|편|회|일|팀|개|인|시간|원|벌)(?![가-힣A-Za-z])",
)
# 항목명 안에 들어간 수량 표기 (예: '5명', '3편') — 단어 단위로 제거
_QTY_INLINE_RE = re.compile(r"\b\d+\s*(?:명|건|식|편|회|일|팀|개|인|시간|벌)\b")


def strip_unit_noise(name: str) -> str:
    s = name
    s = _QTY_INLINE_RE.sub("", s)
    s = _UNIT_TOKEN_RE.sub("", s)
    # 중복 공백·앞뒤 공백·트레일링 separator 정리
    s = re.sub(r"\s+", " ", s).strip(" /·,;:-—")
    return s


# ─── 합계/노이즈 행 검출 ──────────────────────────────────
_AGGREGATE_KEYWORDS = (
    "합계", "총계", "소계", "총합", "총 합계", "거래가격",
    "VAT", "부가세", "Vat", "TOTAL", "Total",
    "(A)", "(B)", "(C)", "(D)", "(E)", "(F)", "(G)", "(H)", "(I)", "(J)",
    "( A )", "( B )", "( C )", "( D )",
    "대행수수료", "수수료 ", "Agency Fee",
)


def is_aggregate_row(name: str) -> bool:
    if not name:
        return True
    n = name.replace(" ", "")
    for kw in _AGGREGATE_KEYWORDS:
        if kw.replace(" ", "") in n:
            return True
    return False


# ─── 동의어 정규화 ────────────────────────────────────────
# 동의어 → canonical 표기. 매칭은 정규화 후 부분일치.
_SYNONYM_MAP: Dict[str, str] = {
    "프로듀싱료": "PD료",
    "producing": "PD료",
    "producer": "PD료",
    "프로듀서": "PD료",

    "editing": "편집",
    "edit": "편집",
    "editor": "편집",

    "recording": "녹음",
    "녹음실": "녹음",
    "녹음실비": "녹음",

    "성우비": "성우료",
    "성우 비": "성우료",

    "음향": "음향료",
    "audio": "음향료",

    "촬영연출": "촬영",
    "촬영기사": "촬영",

    "후반작업": "후반",
    "post": "후반",
    "post production": "후반",
    "후반 작업": "후반",

    "스튜디오대관": "스튜디오",
    "스튜디오대관료": "스튜디오",
    "studio": "스튜디오",

    "art work": "아트워크",
    "artwork": "아트워크",
    "아트 워크": "아트워크",
}


def _normalize_for_key(name: str) -> str:
    s = (name or "").lower()
    s = re.sub(r"\s+", "", s)
    s = re.sub(r"[/\\,\.\-_·•()\[\]:;]", "", s)
    return s


def canonicalize_item_name(name: str) -> str:
    """동의어 매칭 후 canonical 명을 반환. 매칭 없으면 원본 그대로."""
    norm = _normalize_for_key(name)
    for alias, canonical in _SYNONYM_MAP.items():
        if _normalize_for_key(alias) in norm:
            return canonical
    return name


# ─── 후처리 파이프라인 ────────────────────────────────────
def post_process_rows(rows: List[EstimateRow]) -> Tuple[List[EstimateRow], dict]:
    """5단계 정리 후 (정리된 행 리스트, 통계 dict) 반환."""
    stats = dict(input=len(rows), dropped_aggregate=0, dropped_zero=0,
                 dropped_empty_name=0, merged_duplicates=0, unit_stripped=0)

    cleaned: List[EstimateRow] = []
    for r in rows:
        # 정가/대행수수료 자동행은 후처리에서 손대지 않음
        if r.section in ("정가항목", "대행수수료"):
            cleaned.append(r)
            continue

        # 1) 합계/노이즈 행 제외
        if is_aggregate_row(r.item_name):
            stats["dropped_aggregate"] += 1
            continue

        # 2) 단위 노이즈 제거
        original = r.item_name
        new_name = strip_unit_noise(original)
        if new_name != original:
            stats["unit_stripped"] += 1
        if not new_name.strip():
            stats["dropped_empty_name"] += 1
            continue

        # 3) 동의어 정규화는 dedup 키 산출에만 사용 (표시명은 보존)
        r.item_name = new_name
        cleaned.append(r)

    # 4) (canonical_key, vendor) 로 dedup + 금액 합산
    bucket: Dict[Tuple[str, str], EstimateRow] = {}
    keep_others: List[EstimateRow] = []
    for r in cleaned:
        if r.section != "외주비":
            keep_others.append(r)
            continue
        key = (_normalize_for_key(canonicalize_item_name(r.item_name)), (r.vendor or "").strip())
        existing = bucket.get(key)
        if existing is None:
            bucket[key] = r
        else:
            # 같은 키 — amount/qty 합산, source_file 병합
            existing.amount = (existing.amount or 0) + (r.amount or 0)
            existing.quantity = (existing.quantity or 0) + (r.quantity or 0)
            if r.source_file and existing.source_file and r.source_file not in existing.source_file:
                existing.source_file = f"{existing.source_file}, {r.source_file}"
            stats["merged_duplicates"] += 1

    merged = list(bucket.values())

    # 5) Output 양식 동기화 sanitize + 0 행 제거
    # "성우료 1,500,000" 처럼 unit/qty 없이 amount 만 있는 경우 자동 보정.
    stats["sanitized"] = 0
    final: List[EstimateRow] = []
    for r in merged:
        if not r.amount or r.amount == 0:
            stats["dropped_zero"] += 1
            continue
        if not r.quantity or r.quantity == 0:
            r.quantity = 1.0
            stats["sanitized"] += 1
        if not r.unit_price or r.unit_price == 0:
            r.unit_price = r.amount / (r.quantity or 1.0)
            stats["sanitized"] += 1
        final.append(r)

    # 다른 섹션(정가/대행수수료) + 정리된 외주비
    return keep_others + final, stats

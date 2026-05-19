"""
3단계 신호등 검증 (Source 28~33).
- 행 단위:  unit_price * quantity == amount  ?
- 마스터 매칭:  제작단가기준집.pdf 의 정가 기준과 항목명 매칭 ?
- 시트 합계: 정가/외주/대행 합계 + 더블체크
- 매체비 삼각 검증 (Source 19, 22)
"""
from __future__ import annotations
import uuid
from typing import List, Tuple, Optional

from .schemas import EstimateRow, EstimateDocument, TriangleCheck, LightColor
from .master_loader import MASTER

# 부동소수 비교 허용 오차 — Source 28, 29 (input 숫자 보존)
EPS_ABS = 0.5            # 원 단위 (반올림 흡수)
EPS_REL = 1e-4

# ─────────────────────────────────────────────────────────────
# 대행수수료 요율 — 사내 표준 (Source 17, 18)
# ─────────────────────────────────────────────────────────────
# 표준 요율 (전 항목 일괄 적용 기준)
DEFAULT_AGENCY_FEE_RATE = 0.1765         # 17.65%

# 향후 확장: 모델료/진행비 등 일부 항목에 별도 요율 적용 시 사용
REDUCED_AGENCY_FEE_RATE = 0.10           # 10%
REDUCED_RATE_KEYWORDS: Tuple[str, ...] = (
    # 활성화 시 이 키워드를 포함한 외주 행만 REDUCED 요율 적용.
    # 현재는 비활성 — 전 항목 17.65% 일괄 처리.
    # "모델료", "Model", "진행비"
)


def agency_fee_rate_for(item_name: str) -> float:
    """
    개별 외주 행에 적용할 요율 결정.
    - REDUCED_RATE_KEYWORDS 비어있으면 항상 DEFAULT_AGENCY_FEE_RATE 반환.
    - 향후 사내 정책 변경 시 위 상수만 수정하면 즉시 자동 분기.
    """
    if not REDUCED_RATE_KEYWORDS:
        return DEFAULT_AGENCY_FEE_RATE
    name = (item_name or "").lower()
    for kw in REDUCED_RATE_KEYWORDS:
        if kw.lower() in name:
            return REDUCED_AGENCY_FEE_RATE
    return DEFAULT_AGENCY_FEE_RATE


def calculate_agency_fee(outsourcing_rows: List[EstimateRow]) -> Tuple[float, float]:
    """
    외주비 행들에 대해 대행수수료 총액 계산.
    Returns: (총수수료, 실효요율) — 실효요율은 표시용 (가중평균).
    """
    fee = 0.0
    base = 0.0
    for r in outsourcing_rows:
        rate = agency_fee_rate_for(r.item_name)
        fee += r.amount * rate
        base += r.amount
    fee = round(fee)
    effective_rate = (fee / base) if base else DEFAULT_AGENCY_FEE_RATE
    return fee, effective_rate


def _close(a: float, b: float) -> bool:
    return abs(a - b) <= max(EPS_ABS, EPS_REL * max(abs(a), abs(b), 1.0))


def evaluate_row(row: EstimateRow) -> EstimateRow:
    """
    한 행에 대한 신호등 판정 + confidence 산출.
    """
    light: LightColor = "green"
    conf = 0.99
    reasons: List[str] = []

    qty = row.quantity if row.quantity not in (None, 0) else 1.0
    expected = (row.unit_price or 0.0) * qty
    if not _close(expected, row.amount):
        # 단가*수량과 합계 불일치 — 빨강
        light = "red"
        conf = min(conf, 0.70)
        reasons.append(
            f"수식오류: 단가({row.unit_price:,.0f}) × 수량({qty:g}) = {expected:,.0f} ≠ 금액({row.amount:,.0f})"
        )

    # 정가항목 행은 마스터 매칭 확인
    master_hit = MASTER.find(row.item_name) if row.section == "정가항목" else None
    if row.section == "정가항목":
        if master_hit and master_hit.unit_price is not None:
            if not _close(master_hit.unit_price, row.unit_price):
                # 정가 기준과 단가 차이
                light = "yellow" if light != "red" else "red"
                conf = min(conf, 0.92)
                reasons.append(
                    f"정가기준 단가({master_hit.unit_price:,.0f}) 와 입력 단가({row.unit_price:,.0f}) 불일치"
                )
            else:
                reasons.append("제작단가기준집과 단가 일치")
        else:
            light = "yellow" if light != "red" else "red"
            conf = min(conf, 0.92)
            reasons.append("정가기준 항목 매칭 실패 — 카테고리 모호")

    # 항목명에 '편집' 같은 모호 키워드만 있을 경우 yellow
    if light == "green" and row.item_name:
        ambiguous = ["편집", "촬영", "디자인", "기타"]
        if any(row.item_name.strip() == k for k in ambiguous):
            light = "yellow"
            conf = min(conf, 0.93)
            reasons.append("항목명 모호 — AI 추론 매핑 필요")

    # 필수 누락 — 외주비에서 단가/수량이 모두 0이면 위험
    if row.section == "외주비" and row.unit_price == 0 and row.amount == 0:
        light = "yellow" if light == "green" else light
        conf = min(conf, 0.91)
        reasons.append("외주비 단가/금액 모두 0 — 입력 누락 가능")

    row.confidence = conf
    row.light = light
    row.reasoning = " · ".join(reasons) if reasons else "OK"
    return row


def aggregate_sections(rows: List[EstimateRow]) -> Tuple[float, float, float]:
    a = sum(r.amount for r in rows if r.section == "정가항목")
    b = sum(r.amount for r in rows if r.section == "외주비")
    c = sum(r.amount for r in rows if r.section == "대행수수료")
    return a, b, c


# Legacy alias — 외부에서 import 할 수 있어 유지 (단일 항목 빠른 계산)
def compute_agency_fee(sum_b: float, rate: float = DEFAULT_AGENCY_FEE_RATE) -> float:
    return round(sum_b * rate)


_FEE_ROW_ID_PREFIX = "r-fee-auto-"


def _strip_auto_fee_rows(doc: EstimateDocument) -> None:
    """기존 자동생성 대행수수료 행을 제거 (재계산 idempotent 보장)."""
    doc.rows = [r for r in doc.rows if not (
        r.section == "대행수수료" and r.id.startswith(_FEE_ROW_ID_PREFIX)
    )]


def evaluate_document(doc: EstimateDocument) -> EstimateDocument:
    # 0) 이전 자동 대행수수료 행 제거 (update_row 등 재호출 시 중복 방지)
    _strip_auto_fee_rows(doc)

    # 1) 행 단위 평가
    for r in doc.rows:
        evaluate_row(r)

    # 2) 합계 산출 — 외주비/정가/대행수수료(수동 입력분)
    a, b, c_manual = aggregate_sections(doc.rows)
    doc.sum_jeongga = a
    doc.sum_outsourcing = b

    # 3) 대행수수료 자동 계산 (production 한정; Source 17 — 매체비엔 대행수수료 없음)
    #    외주비 합계 × 17.65% 를 합성 행으로 doc.rows 끝에 추가.
    if doc.category_l1 == "production" and c_manual == 0 and b > 0:
        outsourcing_rows = [r for r in doc.rows if r.section == "외주비"]
        fee_amount, eff_rate = calculate_agency_fee(outsourcing_rows)
        fee_row = EstimateRow(
            id=f"{_FEE_ROW_ID_PREFIX}{uuid.uuid4().hex[:6]}",
            section="대행수수료",
            item_name=f"대행수수료 ({eff_rate*100:.2f}%)",
            unit_price=fee_amount,
            quantity=1.0,
            amount=fee_amount,
            source_file="(자동 계산)",
            note=f"외주비 합계 {b:,.0f} × {eff_rate*100:.2f}%",
            confidence=1.0,
            light="green",
            reasoning=f"자동 산출: 외주비 {b:,.0f} × {eff_rate*100:.2f}% = {fee_amount:,.0f}",
        )
        doc.rows.append(fee_row)
        c_total = fee_amount
        doc.notes.append(
            f"⚡ 대행수수료 자동 산입: 외주비 {b:,.0f} × {DEFAULT_AGENCY_FEE_RATE*100:.2f}% = {fee_amount:,.0f}"
        )
    else:
        c_total = c_manual

    doc.sum_agency_fee = c_total
    doc.sum_total = a + b + c_total      # 거래가격 (VAT 별도)
    doc.vat = round(doc.sum_total * 0.1)
    doc.sum_with_vat = doc.sum_total + doc.vat

    # 4) 더블체크 — 모든 행 합산 (자동 fee 포함) == A + B + C
    all_rows_sum = sum(r.amount for r in doc.rows)
    if not _close(all_rows_sum, doc.sum_total):
        doc.warnings.append(
            f"⚠ 더블체크 실패: 모든 행 합계 {all_rows_sum:,.0f} "
            f"!= A({a:,.0f}) + B({b:,.0f}) + C({c_total:,.0f}) = {doc.sum_total:,.0f}"
        )

    # 4) overall 신호등 — 최악 케이스가 전체를 결정
    lights = [r.light for r in doc.rows]
    if "red" in lights or doc.warnings:
        doc.overall_light = "red"
        doc.overall_confidence = min((r.confidence for r in doc.rows), default=0.7)
    elif "yellow" in lights:
        doc.overall_light = "yellow"
        doc.overall_confidence = min((r.confidence for r in doc.rows), default=0.92)
    else:
        doc.overall_light = "green"
        doc.overall_confidence = min((r.confidence for r in doc.rows), default=0.99)
    return doc


def evaluate_triangle(
    media_charged_sum: float,
    billing_status_sum: float,
    media_paid: float,
    agency_fee: float,
) -> TriangleCheck:
    """Source 19, 22 — 매체비 삼각 검증."""
    abc = media_paid + agency_fee
    deltas = [
        abs(media_charged_sum - billing_status_sum),
        abs(media_charged_sum - abc),
        abs(billing_status_sum - abc),
    ]
    delta = max(deltas)
    consistent = delta <= max(EPS_ABS, EPS_REL * max(media_charged_sum, billing_status_sum, abc, 1.0))
    return TriangleCheck(
        media_charged_sum=media_charged_sum,
        billing_status_sum=billing_status_sum,
        media_paid_plus_fee=abc,
        consistent=consistent,
        delta=delta,
    )

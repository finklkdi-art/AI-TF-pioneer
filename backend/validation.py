"""
3단계 신호등 검증 (Source 28~33).
- 행 단위:  unit_price * quantity == amount  ?
- 마스터 매칭:  제작단가기준집.pdf 의 정가 기준과 항목명 매칭 ?
- 시트 합계: 정가/외주/대행 합계 + 더블체크
- 매체비 삼각 검증 (Source 19, 22)
"""
from __future__ import annotations
from typing import List, Tuple, Optional
import math

from .schemas import EstimateRow, EstimateDocument, TriangleCheck, LightColor
from .master_loader import MASTER

# 부동소수 비교 허용 오차 — Source 28, 29 (input 숫자 보존)
EPS_ABS = 0.5            # 원 단위 (반올림 흡수)
EPS_REL = 1e-4


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


def compute_agency_fee(sum_b: float, rate: float = 0.1765) -> float:
    # Source: 영상제작비 패턴 (B) * 17.65% (회사 정책에 따라 10% 등 변주)
    return round(sum_b * rate)


def evaluate_document(doc: EstimateDocument) -> EstimateDocument:
    # 1) 행 단위 평가
    for r in doc.rows:
        evaluate_row(r)
    # 2) 합계 산출
    a, b, c = aggregate_sections(doc.rows)
    doc.sum_jeongga = a
    doc.sum_outsourcing = b
    # 정가 견적서에 대행수수료가 비어있고 production 이면 자동 계산 *보조* — 단, input 행은 절대 수정하지 않음
    if c == 0 and doc.category_l1 == "production":
        c = compute_agency_fee(b)
        doc.notes.append("대행수수료 자동 계산: (B) × 17.65%")
    doc.sum_agency_fee = c
    doc.sum_total = a + b + c
    doc.vat = round(doc.sum_total * 0.1)
    doc.sum_with_vat = doc.sum_total + doc.vat

    # 3) 더블체크 — input vs output 일치 (Source 28, 29)
    input_sum = sum(r.amount for r in doc.rows)
    output_sum = doc.sum_jeongga + doc.sum_outsourcing + doc.sum_agency_fee
    # 대행수수료가 자동 계산된 경우 input_sum 에서 빼고 비교
    rows_c = sum(r.amount for r in doc.rows if r.section == "대행수수료")
    expected_from_rows = input_sum + (doc.sum_agency_fee - rows_c)
    if not _close(expected_from_rows, output_sum):
        doc.warnings.append(
            f"⚠ 더블체크 실패: 행 합계({input_sum:,.0f}) + 대행수수료({doc.sum_agency_fee:,.0f}) "
            f"!= 출력 합계({output_sum:,.0f})"
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

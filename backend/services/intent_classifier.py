"""
3단계 의도 중심 파싱 파이프라인 — STEP 1 (목적 판별) + STEP 2 (템플릿 바인딩).

문서 전체 텍스트를 키워드 세트로 훑어 청구 목적을 1순위로 정의. 좌표 기반
인접 셀 추출 대신 '이 견적서가 무엇을 청구하는가?' 를 먼저 결정하고, 그 목적에
대응하는 표준 엔티티 (녹음실/성우/오디오PD 등) 와 의미 매칭한 데이터만 인용.
"""
from __future__ import annotations
import re
from dataclasses import dataclass
from typing import Dict, List, Tuple


# ─── STEP 1 — 청구 목적 분류 키워드 ──────────────────────
# 각 purpose 별로 (키워드, 가중치) 튜플. 가중치는 식별력에 비례.
PURPOSE_KEYWORDS: Dict[str, List[Tuple[str, int]]] = {
    "AUDIO": [
        ("녹음실", 3), ("성우", 3), ("내레이션", 3), ("내레이터", 3),
        ("믹싱", 3), ("오디오 PD", 3), ("Audio PD", 3),
        ("Song 제작", 3), ("라이선스", 2), ("라이브러리", 2),
        ("Recording", 2), ("BGM", 2), ("음향", 2), ("효과음", 2),
        ("녹음", 1),  # 약한 단독 단어
    ],
    "DI_NTC": [
        ("DI", 3), ("NTC", 3), ("Color Grading", 3), ("색보정", 3), ("색 보정", 3),
        ("Telecine", 3), ("Converting", 3), ("컨버팅", 3),
    ],
    "EDIT_2D_3D": [
        ("Editing", 3), ("종합편집", 3), ("가편집", 3), ("타이틀", 2),
        ("VFX", 3), ("CGI", 3),
        ("합성", 2), ("2D", 2), ("3D", 2), ("CG", 2), ("편집", 1),
    ],
    "CF_PRODUCTION": [
        ("프로덕션 프로듀싱료", 4), ("프로덕션비", 3),
        ("CF프로덕션", 4), ("CF 프로덕션", 4),
        ("감독료", 3), ("연출료", 2),
        ("촬영기사", 2), ("촬영감독", 2), ("촬영조수", 2),
        ("조명감독", 2), ("조명기사", 2),
        ("카메라 기자재", 3), ("조명기자재", 2), ("특수기자재", 2),
        ("스튜디오비", 3), ("스튜디오 대관", 3),
        ("미술비", 2), ("Art Work", 2), ("미 술 비", 2),
        ("Special Effect", 2), ("특수촬영", 2),
    ],
    "PD_FEE": [
        ("프리랜서 PD", 4), ("독립 PD", 4),
        ("PD료", 3), ("프로듀서", 2),
    ],
    "MODEL": [
        ("모델료", 4), ("출연료", 3), ("주연모델", 3), ("조연모델", 3),
        ("초상권", 2), ("부분모델", 2),
    ],
    "PRINT_KV": [
        ("KV", 3), ("키비주얼", 3), ("키 비주얼", 3),
        ("인쇄", 2), ("리터칭", 3), ("일러스트", 2),
    ],
    "BTL": [
        ("BTL", 3), ("이벤트", 2), ("프로모션", 2),
        ("부스", 2), ("자재비", 2), ("운영비", 2), ("보험료", 1),
    ],
}


# ─── STEP 2 — purpose 별 표준 엔티티 (Ground Truth) ──────
# Reference 폴더의 표준 양식 + 제작단가기준집 기준의 핵심 항목.
# LLM 에게 이 항목들과 의미 매칭하여 input 에서 인용하도록 지시.
TEMPLATE_ENTITIES: Dict[str, List[str]] = {
    "AUDIO": [
        "녹음실 대관료",
        "성우료",
        "오디오 PD 연출료",
        "Song 제작 / 저작권 / 라이선스",
        "믹싱료",
        "BGM / 음원",
        "효과음 / 음향",
    ],
    "DI_NTC": [
        "DI / 색보정",
        "NTC / Converting",
        "Telecine",
    ],
    "EDIT_2D_3D": [
        "편집 / Editing (종합편집·가편집)",
        "2D 합성",
        "3D / CGI",
        "VFX",
        "타이틀 / 자막",
    ],
    "CF_PRODUCTION": [
        "프로덕션 프로듀싱료 / 감독료",
        "연출료 (조감독·조연출 포함)",
        "촬영 인건비 (촬영기사·촬영조수)",
        "조명 인건비 (조명기사·조명조수)",
        "카메라 / 조명 / 특수 기자재료",
        "Special Effect",
        "스튜디오 / 로케이션 대관",
        "미술비 / Art Work / Stylist",
        "촬영 진행비 / 보험료 / 모델료",
    ],
    "PD_FEE": [
        "PD 프로듀싱 용역료",
    ],
    "MODEL": [
        "모델료 (주연·조연)",
        "대역 모델료 / 부분모델 / 엑스트라",
        "초상권 사용료",
    ],
    "PRINT_KV": [
        "KV / 키비주얼 디자인",
        "촬영 / 보정 / 리터칭",
        "일러스트 / 그래픽",
    ],
    "BTL": [
        "이벤트 진행 인건비",
        "자재비 / 부스 설치",
        "운영비 / 보험료",
    ],
    "GENERIC": [
        "외주 인건비",
        "외주 장비/대관",
        "후반 작업비",
        "기타 외주비",
    ],
}


@dataclass
class PurposeReport:
    primary: str                              # 1순위 purpose key
    scores: Dict[str, int]                    # 전체 점수
    entities: List[str]                       # primary 의 표준 엔티티
    confidence: float                         # 0~1 (top score 비중)


def classify_purpose(text: str) -> PurposeReport:
    """문서 전체 텍스트에서 청구 목적을 탑다운 판별.

    각 purpose 의 키워드 발견 횟수 × 가중치 합산 → 가장 높은 점수의 purpose 반환.
    모든 점수 0 이면 GENERIC.
    """
    if not text:
        return PurposeReport("GENERIC", {}, TEMPLATE_ENTITIES["GENERIC"], 0.0)
    scores: Dict[str, int] = {}
    for purpose, kws in PURPOSE_KEYWORDS.items():
        s = 0
        for kw, weight in kws:
            # 케이스 무관, 단어 경계 (한글은 단어 경계가 약해서 단순 포함 검사)
            hits = len(re.findall(re.escape(kw), text, flags=re.IGNORECASE))
            s += hits * weight
        scores[purpose] = s
    primary = max(scores.items(), key=lambda x: x[1])
    primary_key, top_score = primary
    total = sum(scores.values())
    if top_score == 0:
        return PurposeReport("GENERIC", scores, TEMPLATE_ENTITIES["GENERIC"], 0.0)
    confidence = top_score / total if total > 0 else 0.0
    return PurposeReport(
        primary=primary_key,
        scores=scores,
        entities=TEMPLATE_ENTITIES.get(primary_key, TEMPLATE_ENTITIES["GENERIC"]),
        confidence=round(confidence, 3),
    )


def format_purpose_context_for_llm(rep: PurposeReport) -> str:
    """LLM user message 에 주입할 컨텍스트 블록을 markdown 으로 직렬화."""
    entities_md = "\n".join(f"  · {e}" for e in rep.entities)
    score_top = sorted(rep.scores.items(), key=lambda x: -x[1])[:5]
    score_md = ", ".join(f"{k}={v}" for k, v in score_top if v > 0) or "(키워드 매칭 없음)"
    return (
        f"[탑다운 분석 결과 — STEP 1]\n"
        f"청구 목적 (Primary Purpose): **{rep.primary}**  (confidence {rep.confidence*100:.0f}%)\n"
        f"키워드 매칭 점수: {score_md}\n\n"
        f"[STEP 2 — 이 목적에 대응하는 표준 엔티티 (Ground Truth)]\n"
        f"아래 항목들과 의미적으로 매칭되는 데이터만 input 문서에서 인용하세요.\n"
        f"인접 셀이 아닌 의미 매칭. 매칭 없으면 추출 금지.\n"
        f"{entities_md}"
    )

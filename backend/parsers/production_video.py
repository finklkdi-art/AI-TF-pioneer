"""영상제작비 견적서 파서 (Source 13).
레퍼런스: 영상제작비견적서1.xlsx, 영상제작비견적서2.xlsx

구조 패턴:
  - 정가항목 (기본료 / Copy료 / Creative Work료 / Direction료 / 기획관리비)
  - 정가합계 (A) — 광고대행사 정가 적용
  - 외주비 (PD / 촬영연출 / POST프로덕션 / 음향 / 보조PD / 성우 ...)
  - 외주비합계 (B)
  - 대행수수료 (C) = (B) * 17.65% (또는 10%)
  - 총합계 = A + B + C  (VAT 별도)
"""
from .base import BaseParser


KNOWN_JEONGGA_ITEMS = {
    "기본료", "Copy료", "Creative Work료", "Direction료", "기획관리비",
}

KNOWN_OUTSOURCING_HEADS = {
    "PD료", "Producer", "촬영연출", "촬영", "POST프로덕션", "POST",
    "편집", "녹음", "음향", "성우", "보조PD", "BGM",
    "VFX", "CG", "DI", "EDIT", "2D", "3D", "Art", "후반작업",
}


class ProductionVideoParser(BaseParser):
    label = "production-video"
    is_media = False

    def __init__(self):
        super().__init__()

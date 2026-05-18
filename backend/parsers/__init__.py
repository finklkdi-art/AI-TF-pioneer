"""
파서 라우터 — 2단계 카테고리 선택값에 따라 적절한 템플릿 파서를 선택 (Source 11, 26).

mapping:
  production / video  -> ProductionVideoParser   (영상제작비견적서)
  production / radio  -> ProductionRadioParser   (라디오제작비견적서, RCM)
  production / print  -> ProductionPrintParser   (인쇄제작비견적서)
  production / btl    -> ProductionBTLParser     (BTL제작비견적서)
  production / other  -> GenericProductionParser
  media / tvc         -> MediaTVCParser
  media / radio       -> MediaRadioParser
  media / print       -> MediaPrintParser
  media / digital     -> MediaDigitalParser
  media / other       -> MediaGenericParser
  (모든 media 파서는 Billing Status 와 1:1 매핑 시도)
"""
from __future__ import annotations
from typing import Dict, Tuple, Type
from .base import BaseParser, GenericProductionParser, GenericMediaParser
from .production_video import ProductionVideoParser
from .production_radio import ProductionRadioParser
from .production_print import ProductionPrintParser
from .production_btl import ProductionBTLParser
from .media_tvc import MediaTVCParser
from .media_radio import MediaRadioParser
from .media_print import MediaPrintParser
from .media_digital import MediaDigitalParser

_REGISTRY: Dict[Tuple[str, str], Type[BaseParser]] = {
    ("production", "video"): ProductionVideoParser,
    ("production", "radio"): ProductionRadioParser,
    ("production", "print"): ProductionPrintParser,
    ("production", "btl"): ProductionBTLParser,
    ("production", "other"): GenericProductionParser,
    ("media", "tvc"): MediaTVCParser,
    ("media", "radio"): MediaRadioParser,
    ("media", "print"): MediaPrintParser,
    ("media", "digital"): MediaDigitalParser,
    ("media", "other"): GenericMediaParser,
}


def get_parser(category_l1: str, category_l2: str) -> BaseParser:
    cls = _REGISTRY.get((category_l1, category_l2))
    if cls is None:
        # Source 27 — fallback. 가장 가까운 일반 파서.
        cls = GenericProductionParser if category_l1 == "production" else GenericMediaParser
    return cls()

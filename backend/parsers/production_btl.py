"""BTL(이벤트/프로모션)제작비 견적서 파서 (Source 12, 14).
인건비/자재비/운영비/보험료/기타.
"""
from .base import BaseParser


class ProductionBTLParser(BaseParser):
    label = "production-btl"
    is_media = False

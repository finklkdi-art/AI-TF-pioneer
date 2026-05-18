"""인쇄(KV/이미지)제작비 견적서 파서 (Source 12)."""
from .base import BaseParser


class ProductionPrintParser(BaseParser):
    label = "production-print"
    is_media = False

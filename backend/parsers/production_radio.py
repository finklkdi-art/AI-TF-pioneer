"""라디오제작비 견적서 파서 (Source 12).
레퍼런스: 라디오제작비견적서1.xlsx (RCM 시트)

구조 패턴: 정가항목 / 녹음실 / 성우비 / BGM / 후반작업.
"""
from .base import BaseParser


class ProductionRadioParser(BaseParser):
    label = "production-radio"
    is_media = False

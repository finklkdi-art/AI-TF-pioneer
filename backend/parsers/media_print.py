"""PRINT 매체비 파서."""
from .base import BaseParser


class MediaPrintParser(BaseParser):
    label = "media-print"
    is_media = True

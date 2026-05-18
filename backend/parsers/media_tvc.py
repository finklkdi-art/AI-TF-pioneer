"""TVC 매체비 파서 (Source 17~22)."""
from .base import BaseParser


class MediaTVCParser(BaseParser):
    label = "media-tvc"
    is_media = True

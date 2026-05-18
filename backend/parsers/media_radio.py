"""라디오 매체비 파서."""
from .base import BaseParser


class MediaRadioParser(BaseParser):
    label = "media-radio"
    is_media = True

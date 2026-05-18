"""디지털(Display/Search/SNS) 매체비 파서."""
from .base import BaseParser


class MediaDigitalParser(BaseParser):
    label = "media-digital"
    is_media = True

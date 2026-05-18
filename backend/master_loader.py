"""
'제작단가기준집.pdf' 마스터 데이터 로더 (Source 16, 36).
- 앱 구동 시 1회 메모리 적재.
- 일반 AE는 read-only. Admin 만 reload 가능.
- 휘발성 보안 원칙에 따라 디스크엔 원본 PDF 외 별도 캐시 저장하지 않음.
"""
from __future__ import annotations
import os
import re
from pathlib import Path
from threading import RLock
from typing import List, Dict, Optional

from .schemas import MasterRefItem


# BLUE_NINE_REF_DIR 환경변수로 운영 환경에서 외부 경로 주입 가능 (Source 36 — Admin sync).
from .config import settings as _settings
_default_ref_dir = Path(__file__).resolve().parent.parent / "reference"
REF_DIR = Path(_settings.ref_dir) if _settings.ref_dir else _default_ref_dir
MASTER_PDF_NAME = _settings.master_pdf


def _split_currency(num_str: str) -> Optional[float]:
    s = num_str.replace(",", "").replace("원", "").strip()
    if not s:
        return None
    try:
        return float(s)
    except ValueError:
        return None


def _parse_pdf_text(text: str) -> List[MasterRefItem]:
    """간소화된 룰: '항목명 ... 단가(숫자, 콤마 가능) (원/단위)' 형태 라인을 추출.
    실제 운영 PDF의 정밀 파싱은 별도 OCR/표 추출 모듈로 확장.
    """
    out: List[MasterRefItem] = []
    section = "정가항목"
    line_re = re.compile(r"^(?P<name>[^\d\n]{2,40}?)\s+(?P<price>[\d,]{3,})(?:\s*원)?(?:\s*/\s*(?P<unit>\S+))?\s*$")
    for raw in text.splitlines():
        line = raw.strip()
        if not line:
            continue
        if any(k in line for k in ("[정가", "[기본", "[기준")):
            section = line.strip("[]")
            continue
        m = line_re.match(line)
        if not m:
            continue
        price = _split_currency(m.group("price"))
        if price is None:
            continue
        out.append(MasterRefItem(
            code=f"M-{len(out)+1:04d}",
            section=section,
            item_name=m.group("name").strip(),
            unit_price=price,
            unit=m.group("unit"),
        ))
    return out


def _extract_pdf_text(path: Path) -> str:
    try:
        import pdfplumber  # type: ignore
        text_parts: List[str] = []
        with pdfplumber.open(str(path)) as pdf:
            for pg in pdf.pages:
                t = pg.extract_text() or ""
                text_parts.append(t)
        return "\n".join(text_parts)
    except Exception:
        try:
            from pypdf import PdfReader  # type: ignore
            r = PdfReader(str(path))
            return "\n".join((p.extract_text() or "") for p in r.pages)
        except Exception:
            return ""


class MasterStore:
    def __init__(self) -> None:
        self._lock = RLock()
        self._items: List[MasterRefItem] = []
        self._by_name: Dict[str, MasterRefItem] = {}
        self._loaded_path: Optional[str] = None
        self._version: int = 0

    def load(self, path: Optional[Path] = None) -> int:
        p = path or (REF_DIR / MASTER_PDF_NAME)
        if not p.exists():
            return 0
        text = _extract_pdf_text(p)
        items = _parse_pdf_text(text)
        with self._lock:
            self._items = items
            self._by_name = {it.item_name: it for it in items}
            self._loaded_path = str(p)
            self._version += 1
        return len(items)

    def items(self) -> List[MasterRefItem]:
        with self._lock:
            return list(self._items)

    def find(self, item_name: str) -> Optional[MasterRefItem]:
        if not item_name:
            return None
        with self._lock:
            if item_name in self._by_name:
                return self._by_name[item_name]
            # 느슨한 매칭 (포함 관계)
            for k, v in self._by_name.items():
                if item_name in k or k in item_name:
                    return v
        return None

    def info(self) -> Dict:
        with self._lock:
            return {
                "loaded_path": self._loaded_path,
                "version": self._version,
                "item_count": len(self._items),
            }


MASTER = MasterStore()

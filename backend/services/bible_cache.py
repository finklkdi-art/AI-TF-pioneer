"""
Ground Truth Bible 메모리 캐시 — 앱 부팅 시 1회 적재.

  · 제작단가기준집.pdf — 사내 정가 기준 (4p/30p/31p 기획료/카피료/디렉션료 등)
  · reference/output.xlsx — 표준 견적서 출력 레이아웃 (A·B·C 섹션 구조)

이 두 가지는 LLM 시맨틱 파서의 시스템 프롬프트에 매번 주입되어 input 문서의
'비즈니스적 의도' 추론에 사용. 휘발성 (Source 9) 위배 아님 — 회사 공통 자산.
"""
from __future__ import annotations
import io
from pathlib import Path
from typing import Optional

from ..config import settings


_BIBLE_TEXT: str = ""
_OUTPUT_LAYOUT: str = ""
_LOADED: bool = False


def _read_ref_dir() -> Path:
    if settings.ref_dir:
        return Path(settings.ref_dir)
    return Path(__file__).resolve().parent.parent.parent / "reference"


def _load_pdf_text(path: Path, max_chars: int = 8000) -> str:
    """제작단가기준집.pdf 의 텍스트 추출. 4p/30p/31p 를 우선."""
    if not path.exists():
        return ""
    try:
        import pdfplumber
        with pdfplumber.open(str(path)) as pdf:
            n = len(pdf.pages)
            priority_idx = [3, 29, 30]      # 0-indexed 4p, 30p, 31p
            ordered: list[str] = []
            for i in priority_idx:
                if 0 <= i < n:
                    t = pdf.pages[i].extract_text() or ""
                    if t.strip():
                        ordered.append(f"[Page {i+1}]\n{t}")
            for i in range(n):
                if i in priority_idx:
                    continue
                t = pdf.pages[i].extract_text() or ""
                if t.strip():
                    ordered.append(t)
                if sum(len(x) for x in ordered) > max_chars:
                    break
        joined = "\n\n".join(ordered)
        return joined[:max_chars]
    except Exception:
        try:
            from pypdf import PdfReader
            r = PdfReader(str(path))
            return "\n\n".join((p.extract_text() or "") for p in r.pages)[:max_chars]
        except Exception:
            return ""


def _load_xlsx_layout(path: Path, max_rows: int = 30) -> str:
    """reference/output.xlsx 의 첫 시트 레이아웃을 markdown 으로 직렬화."""
    if not path.exists():
        return ""
    try:
        import pandas as pd
        xl = pd.ExcelFile(path)
        df = pd.read_excel(xl, sheet_name=xl.sheet_names[0], header=None, nrows=max_rows)
        # to_markdown 은 tabulate 필요 — 없으면 to_string fallback
        try:
            return df.fillna("").to_markdown(index=False)
        except Exception:
            return df.fillna("").to_string(index=False)
    except Exception:
        return ""


def init_bible_cache() -> dict:
    """앱 부팅 시 호출. 멱등 — 두 번 호출돼도 안전."""
    global _BIBLE_TEXT, _OUTPUT_LAYOUT, _LOADED
    if _LOADED:
        return info()
    ref = _read_ref_dir()
    _BIBLE_TEXT = _load_pdf_text(ref / settings.master_pdf)
    _OUTPUT_LAYOUT = _load_xlsx_layout(ref / "output.xlsx")
    _LOADED = True
    return info()


def info() -> dict:
    return {
        "loaded": _LOADED,
        "bible_chars": len(_BIBLE_TEXT),
        "output_layout_chars": len(_OUTPUT_LAYOUT),
    }


def get_bible_text() -> str:
    return _BIBLE_TEXT


def get_output_layout() -> str:
    return _OUTPUT_LAYOUT

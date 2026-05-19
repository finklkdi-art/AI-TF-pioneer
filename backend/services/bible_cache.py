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
_PURPOSE_TEMPLATES: dict[str, str] = {}      # purpose key → 표준 시트 markdown
_LOADED: bool = False


# Purpose → 시트명 hint 매핑 (intent_classifier.PURPOSE_KEYWORDS 와 정렬)
PURPOSE_TO_SHEET_HINTS: dict[str, tuple[str, ...]] = {
    "AUDIO":         ("녹음 표준견적서", "녹음", "Recording"),
    "DI_NTC":        ("포스트프로덕션", "DI", "NTC", "Telecine"),
    "EDIT_2D_3D":    ("포스트프로덕션", "편집", "Editing", "합성"),
    "CF_PRODUCTION": ("CF프로덕션 표준견적서", "CF프로덕션", "프로덕션"),
    "PD_FEE":        ("PD 표준견적서", "PD"),
}

# 표준 시트들을 가져올 후보 파일 (우선순위 순)
_TEMPLATE_SOURCE_FILES: tuple[str, ...] = (
    "영상제작비견적서2.xlsx",
    "input1.xlsx",
    "input2.xlsx",
    "영상제작비견적서1.xlsx",
)


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


def _load_purpose_templates(ref_dir: Path) -> dict[str, str]:
    """reference 파일들 중 표준 시트를 찾아 purpose 별 markdown 으로 캐싱."""
    import pandas as pd
    out: dict[str, str] = {}
    for src_name in _TEMPLATE_SOURCE_FILES:
        p = ref_dir / src_name
        if not p.exists():
            continue
        try:
            xl = pd.ExcelFile(p)
        except Exception:
            continue
        for sh in xl.sheet_names:
            for purpose, hints in PURPOSE_TO_SHEET_HINTS.items():
                if purpose in out:
                    continue                       # 첫 매칭 우선
                if not any(h in sh for h in hints):
                    continue
                try:
                    df = pd.read_excel(xl, sheet_name=sh, header=None, nrows=20).fillna("")
                    try:
                        md = df.to_markdown(index=False)
                    except Exception:
                        md = df.to_string(index=False)
                    out[purpose] = f"### 표준 견적서 시트: {sh}  (출처: {src_name})\n\n{md}"
                except Exception:
                    pass
        if len(out) >= len(PURPOSE_TO_SHEET_HINTS):
            break
    return out


def init_bible_cache() -> dict:
    """앱 부팅 시 호출. 멱등 — 두 번 호출돼도 안전."""
    global _BIBLE_TEXT, _OUTPUT_LAYOUT, _PURPOSE_TEMPLATES, _LOADED
    if _LOADED:
        return info()
    ref = _read_ref_dir()
    _BIBLE_TEXT = _load_pdf_text(ref / settings.master_pdf)
    _OUTPUT_LAYOUT = _load_xlsx_layout(ref / "output.xlsx")
    _PURPOSE_TEMPLATES = _load_purpose_templates(ref)
    _LOADED = True
    return info()


def info() -> dict:
    return {
        "loaded": _LOADED,
        "bible_chars": len(_BIBLE_TEXT),
        "output_layout_chars": len(_OUTPUT_LAYOUT),
        "purpose_templates": {k: len(v) for k, v in _PURPOSE_TEMPLATES.items()},
    }


def get_bible_text() -> str:
    return _BIBLE_TEXT


def get_output_layout() -> str:
    return _OUTPUT_LAYOUT


def get_template_for_purpose(purpose: str) -> str:
    """STEP 2 — 판별된 purpose 에 대응하는 표준 견적서 시트 markdown 반환."""
    return _PURPOSE_TEMPLATES.get(purpose, "")

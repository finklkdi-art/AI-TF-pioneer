"""
BLUE NINE — central config / secret accessor.

설계 원칙:
- 키가 없어도 앱은 정상 부팅한다 (Source 9 휘발성 원칙과 무관 — 단순히 LLM 미연동 시 휴리스틱으로 동작).
- 키가 채워지면 향후 Precise 모드의 LlamaParse / Anthropic 통합 진입점에서 활용.
- 로컬에서는 `.env` 파일이 있으면 자동 로드. 운영(Render)에서는 환경변수가 직접 주입됨.
"""
from __future__ import annotations
import os
from pathlib import Path
from typing import Optional


def _load_dotenv_if_present() -> None:
    """python-dotenv 가 설치되어 있으면 repo_root/.env 자동 로드.
    설치 안 되어 있어도 무시 (운영 환경엔 OS env vars 가 이미 주입됨)."""
    try:
        from dotenv import load_dotenv  # type: ignore
    except ImportError:
        return
    candidates = [
        Path(__file__).resolve().parent.parent / ".env",   # repo_root/.env
        Path.cwd() / ".env",                                # CWD 기준
    ]
    for p in candidates:
        if p.exists():
            load_dotenv(p, override=False)
            return


_load_dotenv_if_present()


def _get(key: str, default: Optional[str] = None) -> Optional[str]:
    v = os.environ.get(key)
    if v is None or v.strip() == "":
        return default
    return v.strip()


class Settings:
    """singleton-스타일 설정 접근자."""

    # ── 운영 환경
    env: str = _get("BLUE_NINE_ENV", "development") or "development"
    port: int = int(_get("PORT", "8088") or "8088")
    host: str = _get("HOST", "127.0.0.1") or "127.0.0.1"
    reload: bool = (_get("RELOAD", "true") or "true").lower() in ("1", "true", "yes", "on")
    allowed_origins: str = _get("BLUE_NINE_ALLOWED_ORIGINS", "*") or "*"
    admin_key: str = _get("BLUE_NINE_ADMIN_KEY", "BLUE_NINE_ADMIN") or "BLUE_NINE_ADMIN"

    # ── 마스터 데이터
    ref_dir: Optional[str] = _get("BLUE_NINE_REF_DIR")
    master_pdf: str = _get("BLUE_NINE_MASTER_PDF", "제작단가기준집.pdf") or "제작단가기준집.pdf"

    # ── LLM 키 (스캐폴딩 — 현재 미사용, 키가 없어도 None 으로 안전 noop)
    llamaparse_api_key: Optional[str] = _get("LLAMAPARSE_API_KEY")
    anthropic_api_key: Optional[str] = _get("ANTHROPIC_API_KEY")
    anthropic_model: str = _get("ANTHROPIC_MODEL", "claude-sonnet-4-6") or "claude-sonnet-4-6"

    @property
    def is_production(self) -> bool:
        return self.env == "production"

    @property
    def has_llm_keys(self) -> bool:
        """향후 통합 코드에서 분기용. 현재는 정보 표시 목적만."""
        return bool(self.llamaparse_api_key and self.anthropic_api_key)

    def llm_status(self) -> dict:
        """루트 / health 응답에 노출 가능한 안전한 상태 (값 자체는 노출 안 함)."""
        return {
            "llamaparse_configured": bool(self.llamaparse_api_key),
            "anthropic_configured": bool(self.anthropic_api_key),
            "anthropic_model": self.anthropic_model if self.anthropic_api_key else None,
            "active_pipeline": "llm" if self.has_llm_keys else "heuristic",
        }


settings = Settings()

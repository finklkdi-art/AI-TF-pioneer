"""
Dev / Prod launcher.

로컬: `python -m backend.run`            → 127.0.0.1:8088, reload on
Render: `uvicorn backend.main:app ...`    → render.yaml 의 startCommand 사용
        (이 파일은 직접 사용되지 않지만, 동일 동작을 보장하기 위해 env 를 따라감)
"""
import os
import uvicorn


def _bool(v: str | None, default: bool = False) -> bool:
    if v is None:
        return default
    return v.lower() in ("1", "true", "yes", "on")


if __name__ == "__main__":
    is_prod = os.environ.get("BLUE_NINE_ENV", "development") == "production"
    host = os.environ.get("HOST", "0.0.0.0" if is_prod else "127.0.0.1")
    port = int(os.environ.get("PORT", "8088"))
    reload = _bool(os.environ.get("RELOAD"), default=not is_prod)
    uvicorn.run("backend.main:app", host=host, port=port, reload=reload)

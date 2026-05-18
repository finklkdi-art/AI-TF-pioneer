"""
모니터링 미들웨어 (요구사항 4):
- Latency, 트래픽 호출수, 예상 토큰 사용량, API 과금 비용($) 산출
- /api/monitor/* 엔드포인트로 노출
- 휘발성 (프로세스 메모리)
"""
from __future__ import annotations
import time
from collections import deque
from threading import RLock
from typing import Deque, Dict, List
from datetime import datetime

from starlette.middleware.base import BaseHTTPMiddleware
from starlette.requests import Request
from starlette.responses import Response

from .schemas import ApiLogEntry

# 단가 가정 — Anthropic Claude Opus 4.x 정도의 모델 사용 가정
# 실측치와 무관하게 데모용. 운영 단계에서 실제 모델/벤더에 맞춰 교체.
PRICE_IN_PER_MTOK_USD = 15.0    # $/1M input token
PRICE_OUT_PER_MTOK_USD = 75.0   # $/1M output token
# 1 token ≒ 4 bytes 가정 (한국어는 더 길지만 안전 추정)
BYTES_PER_TOKEN = 4
# Fast 모드 보정 — 이미지 임베딩 처리는 평균적으로 30~50% 토큰 절감 (Source 4)
FAST_DISCOUNT = 0.5


def estimate_tokens_and_cost(bytes_in: int, bytes_out: int, mode_hint: str = "precise"):
    tok_in = max(1, bytes_in // BYTES_PER_TOKEN)
    tok_out = max(1, bytes_out // BYTES_PER_TOKEN)
    if mode_hint == "fast":
        tok_in = int(tok_in * FAST_DISCOUNT)
        tok_out = int(tok_out * FAST_DISCOUNT)
    cost = (tok_in / 1_000_000) * PRICE_IN_PER_MTOK_USD \
         + (tok_out / 1_000_000) * PRICE_OUT_PER_MTOK_USD
    return tok_in + tok_out, round(cost, 6)


class MonitorBus:
    """링버퍼 형태의 메모리 로그 — 휘발성."""

    def __init__(self, cap: int = 1000):
        self._logs: Deque[ApiLogEntry] = deque(maxlen=cap)
        self._lock = RLock()

    def push(self, entry: ApiLogEntry) -> None:
        with self._lock:
            self._logs.append(entry)

    def snapshot(self, last_n: int = 200) -> List[ApiLogEntry]:
        with self._lock:
            return list(self._logs)[-last_n:]

    def aggregate(self) -> Dict[str, float]:
        with self._lock:
            if not self._logs:
                return {
                    "total_calls": 0, "avg_latency_ms": 0.0,
                    "p95_latency_ms": 0.0, "total_tokens": 0,
                    "total_cost_usd": 0.0, "error_rate": 0.0,
                }
            lats = sorted(e.latency_ms for e in self._logs)
            total = len(lats)
            p95 = lats[int(total * 0.95) - 1] if total >= 20 else lats[-1]
            errs = sum(1 for e in self._logs if e.status >= 500)
            return {
                "total_calls": total,
                "avg_latency_ms": round(sum(lats) / total, 2),
                "p95_latency_ms": round(p95, 2),
                "total_tokens": sum(e.est_tokens for e in self._logs),
                "total_cost_usd": round(sum(e.est_cost_usd for e in self._logs), 6),
                "error_rate": round(errs / total, 4),
            }


BUS = MonitorBus()


class BillingMiddleware(BaseHTTPMiddleware):
    """모든 /api/* 요청에 대해 자동 로깅."""

    async def dispatch(self, request: Request, call_next):
        t0 = time.perf_counter()
        # body 길이 (Content-Length 우선)
        try:
            bytes_in = int(request.headers.get("content-length", "0") or 0)
        except ValueError:
            bytes_in = 0
        mode_hint = request.headers.get("x-blue-nine-mode", "precise")
        status = 500
        bytes_out = 0
        response: Response | None = None
        try:
            response = await call_next(request)
            status = response.status_code
            cl = response.headers.get("content-length")
            if cl:
                bytes_out = int(cl)
            else:
                # body length 추정 — streaming 응답엔 0이어도 OK
                bytes_out = 0
        finally:
            elapsed_ms = (time.perf_counter() - t0) * 1000.0
            tokens, cost = estimate_tokens_and_cost(bytes_in, bytes_out, mode_hint)
            BUS.push(ApiLogEntry(
                timestamp=datetime.utcnow().isoformat(),
                method=request.method,
                path=str(request.url.path),
                status=status,
                latency_ms=round(elapsed_ms, 2),
                bytes_in=bytes_in,
                bytes_out=bytes_out,
                est_tokens=tokens,
                est_cost_usd=cost,
            ))
        return response

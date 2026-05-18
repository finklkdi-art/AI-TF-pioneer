"""휘발성 세션 저장소 (Source 9, 10)
- DB/디스크 영속화 없음. 프로세스 메모리에만 보관.
- idle timeout 경과 시 자동 파기.
"""
from __future__ import annotations
import time
import uuid
from threading import RLock
from typing import Dict, Optional, List
from .schemas import EstimateDocument

IDLE_TIMEOUT_SEC = 30 * 60   # 30분 idle 자동 파기


class _Session:
    __slots__ = ("session_id", "last_seen", "estimates", "uploads", "consents")

    def __init__(self, sid: str):
        self.session_id = sid
        self.last_seen = time.time()
        self.estimates: Dict[str, EstimateDocument] = {}
        # 협력사 파일은 raw bytes 자체를 메모리 보관 (DB 미사용 — Source 9)
        self.uploads: Dict[str, bytes] = {}
        self.consents: Dict[str, bool] = {}


class SessionStore:
    def __init__(self) -> None:
        self._sessions: Dict[str, _Session] = {}
        self._lock = RLock()

    def _gc(self) -> None:
        now = time.time()
        dead = [sid for sid, s in self._sessions.items() if now - s.last_seen > IDLE_TIMEOUT_SEC]
        for sid in dead:
            self._sessions.pop(sid, None)

    def create(self) -> str:
        with self._lock:
            self._gc()
            sid = uuid.uuid4().hex
            self._sessions[sid] = _Session(sid)
            return sid

    def get(self, sid: str) -> _Session:
        with self._lock:
            self._gc()
            if sid not in self._sessions:
                # 세션이 만료/없을 경우 자동 생성하지 않고 신규 발급
                self._sessions[sid] = _Session(sid)
            s = self._sessions[sid]
            s.last_seen = time.time()
            return s

    def drop(self, sid: str) -> None:
        with self._lock:
            self._sessions.pop(sid, None)

    def list_estimates(self, sid: str) -> List[EstimateDocument]:
        return list(self.get(sid).estimates.values())

    def put_estimate(self, sid: str, doc: EstimateDocument) -> None:
        self.get(sid).estimates[doc.estimate_id] = doc

    def get_estimate(self, sid: str, eid: str) -> Optional[EstimateDocument]:
        return self.get(sid).estimates.get(eid)

    def put_upload(self, sid: str, name: str, data: bytes) -> None:
        self.get(sid).uploads[name] = data

    def get_upload(self, sid: str, name: str) -> Optional[bytes]:
        return self.get(sid).uploads.get(name)


STORE = SessionStore()

"""BLUE NINE — Pydantic models (Source 7, 8, 11, 12, 17 / Rule Book v1.0)."""
from __future__ import annotations
from typing import List, Optional, Literal, Dict, Any
from pydantic import BaseModel, Field
from datetime import datetime


CategoryLevel1 = Literal["production", "media"]
CategoryLevel2_Production = Literal["video", "radio", "print", "btl", "other"]
CategoryLevel2_Media = Literal["tvc", "radio", "print", "digital", "other"]
ProcessingMode = Literal["fast", "precise"]
LightColor = Literal["green", "yellow", "red"]


class EstimateRow(BaseModel):
    """A single line item in an estimate."""
    id: str
    section: str = Field(..., description="정가항목/외주비/대행수수료 등 섹션 키")
    category: Optional[str] = None
    item_name: str = ""
    detail: Optional[str] = None
    vendor: Optional[str] = Field(None, description="협력사 / 매체사")
    unit_price: float = 0.0
    quantity: float = 1.0
    amount: float = 0.0
    note: Optional[str] = None
    confidence: float = 1.0
    light: LightColor = "green"
    reasoning: Optional[str] = None
    editable_fields: List[str] = Field(default_factory=lambda: ["unit_price", "quantity", "amount"])


class TriangleCheck(BaseModel):
    """매체비 삼각 검증 (Source 19, 20, 22)."""
    media_charged_sum: float = 0.0          # 광고주 청구액 합계
    billing_status_sum: float = 0.0         # Billing Status 청구액 합계
    media_paid_plus_fee: float = 0.0        # 매체사 지급액 + 대행수수료
    consistent: bool = True
    delta: float = 0.0


class EstimateDocument(BaseModel):
    """완성된 견적서 한 건 — 휘발성 세션 메모리에만 존재 (Source 9, 10)."""
    estimate_id: str
    session_id: str
    version_label: str = "초안"            # 사전견적/1차/2차/최종 (Source 8)
    category_l1: CategoryLevel1
    category_l2: str
    mode: ProcessingMode
    client: Optional[str] = None
    campaign: Optional[str] = None
    job_no: Optional[str] = None
    issue_date: str = ""
    rows: List[EstimateRow] = []
    sum_jeongga: float = 0.0          # (A) 정가합계
    sum_outsourcing: float = 0.0      # (B) 외주비합계
    sum_agency_fee: float = 0.0       # (C) 대행수수료
    sum_total: float = 0.0
    vat: float = 0.0
    sum_with_vat: float = 0.0
    overall_light: LightColor = "green"
    overall_confidence: float = 1.0
    triangle: Optional[TriangleCheck] = None
    warnings: List[str] = []
    notes: List[str] = []
    created_at: str = Field(default_factory=lambda: datetime.utcnow().isoformat())


class ParseRequest(BaseModel):
    session_id: str
    category_l1: CategoryLevel1
    category_l2: str
    mode: ProcessingMode = "precise"
    client: Optional[str] = None
    campaign: Optional[str] = None
    version_label: str = "초안"


class UpdateRowRequest(BaseModel):
    session_id: str
    estimate_id: str
    row_id: str
    patch: Dict[str, Any]


class MasterRefItem(BaseModel):
    code: str
    section: str
    item_name: str
    unit_price: Optional[float] = None
    unit: Optional[str] = None
    note: Optional[str] = None


class ApiLogEntry(BaseModel):
    timestamp: str
    method: str
    path: str
    status: int
    latency_ms: float
    bytes_in: int
    bytes_out: int
    est_tokens: int
    est_cost_usd: float

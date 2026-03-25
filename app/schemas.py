from pydantic import BaseModel
from typing import Optional
from datetime import datetime


# --- Rate ---

class RateCreate(BaseModel):
    charge_type: str
    pickup_code: Optional[str] = None
    odcy_code: Optional[str] = None
    dest_code: Optional[str] = None
    container_type: Optional[str] = None
    unit_price: float
    memo: Optional[str] = None


class RateUpdate(BaseModel):
    charge_type: Optional[str] = None
    pickup_code: Optional[str] = None
    odcy_code: Optional[str] = None
    dest_code: Optional[str] = None
    container_type: Optional[str] = None
    unit_price: Optional[float] = None
    memo: Optional[str] = None


class RateResponse(BaseModel):
    id: int
    charge_type: str
    pickup_code: Optional[str]
    odcy_code: Optional[str]
    dest_code: Optional[str]
    container_type: Optional[str]
    unit_price: float
    memo: Optional[str]
    created_at: Optional[datetime]

    class Config:
        from_attributes = True


# --- Session ---

class SessionSummary(BaseModel):
    id: int
    filename: str
    uploaded_at: Optional[datetime]
    total_rows: int
    trkv_pass: int
    trkv_fail: int
    trkv_no_rate: int
    storage_pass: int
    storage_fail: int
    storage_no_rate: int
    handling_pass: int
    handling_fail: int
    handling_no_rate: int
    shuttle_pass: int
    shuttle_fail: int
    shuttle_no_rate: int
    total_diff: float

    class Config:
        from_attributes = True


# --- Result ---

class ResultRow(BaseModel):
    id: int
    session_id: int
    row_number: int
    container_no: Optional[str]
    transport_date: Optional[str]
    pickup_code: Optional[str]
    pickup_name: Optional[str]
    odcy_code: Optional[str]
    odcy_name: Optional[str]
    dest_code: Optional[str]
    dest_name: Optional[str]
    container_type: Optional[str]
    dg_flag: Optional[bool]
    trkv_actual: Optional[float]
    trkv_expected: Optional[float]
    trkv_diff: Optional[float]
    trkv_status: Optional[str]
    storage_actual: Optional[float]
    storage_expected: Optional[float]
    storage_diff: Optional[float]
    storage_status: Optional[str]
    handling_actual: Optional[float]
    handling_expected: Optional[float]
    handling_diff: Optional[float]
    handling_status: Optional[str]
    shuttle_actual: Optional[float]
    shuttle_expected: Optional[float]
    shuttle_diff: Optional[float]
    shuttle_status: Optional[str]
    overall_status: Optional[str]
    memo: Optional[str]

    class Config:
        from_attributes = True

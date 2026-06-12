"""
자동 F5 입력 API 라우터
대시보드 토글이 호출. 온이면 15분마다 OS 에 F5 키 입력(활성 창 새로고침).
"""
from typing import Optional

from fastapi import APIRouter
from pydantic import BaseModel

from app.services import autopress_service

router = APIRouter()


@router.get("/status")
def status():
    return autopress_service.get_status()


class ToggleRequest(BaseModel):
    enabled: bool
    interval_sec: Optional[int] = None


@router.post("/toggle")
def toggle(body: ToggleRequest):
    return autopress_service.set_enabled(body.enabled, body.interval_sec)


class IntervalRequest(BaseModel):
    interval_sec: int


@router.post("/interval")
def set_interval(body: IntervalRequest):
    return autopress_service.set_interval(body.interval_sec)


@router.post("/test")
def test_press():
    """F5 1회 + 마우스 1px 이동 즉시 실행(동작 확인용)."""
    f5 = autopress_service.send_f5()
    mouse = autopress_service.move_mouse_1px()
    return {"f5": f5, "mouse": mouse}

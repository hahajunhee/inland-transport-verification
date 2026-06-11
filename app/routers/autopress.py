"""
자동 F5 입력 API 라우터
대시보드 토글이 호출. 온이면 15분마다 OS 에 F5 키 입력(활성 창 새로고침).
"""
from fastapi import APIRouter
from pydantic import BaseModel

from app.services import autopress_service

router = APIRouter()


@router.get("/status")
def status():
    return autopress_service.get_status()


class ToggleRequest(BaseModel):
    enabled: bool


@router.post("/toggle")
def toggle(body: ToggleRequest):
    return autopress_service.set_enabled(body.enabled)


@router.post("/test")
def test_press():
    """F5 1회 즉시 전송(동작 확인용)."""
    ok = autopress_service.send_f5()
    return {"sent": ok}

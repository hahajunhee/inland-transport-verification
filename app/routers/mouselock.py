"""
마우스 고정 API 라우터
- 핫키 Alt+1 로 토글되는 기능의 상태 조회/수동 토글.
"""
from fastapi import APIRouter

from app.services import mouselock_service

router = APIRouter()


@router.get("/status")
def status():
    return mouselock_service.get_status()


@router.post("/toggle")
def toggle():
    """Alt+1 과 동일한 토글(수동/테스트용)."""
    mouselock_service.request_toggle()
    return {"requested": True}


@router.post("/off")
def off():
    """켜져 있으면 강제 해제."""
    mouselock_service.force_off()
    return {"requested": True}

from fastapi import APIRouter, HTTPException
from typing import List, Optional
from app.schemas import RateCreate, RateUpdate, RateResponse
from app.services import rate_service

router = APIRouter()


@router.get("", response_model=List[RateResponse])
def list_rates(
    charge_type: Optional[str] = None,
    pickup_code: Optional[str] = None,
    dest_code: Optional[str] = None,
):
    return rate_service.get_all_rates(charge_type, pickup_code, dest_code)


@router.post("", response_model=RateResponse)
def create_rate(body: RateCreate):
    return rate_service.create_rate(body.model_dump())


@router.put("/{rate_id}", response_model=RateResponse)
def update_rate(rate_id: int, body: RateUpdate):
    rate = rate_service.update_rate(rate_id, body.model_dump(exclude_unset=True))
    if not rate:
        raise HTTPException(status_code=404, detail="요율을 찾을 수 없습니다.")
    return rate


@router.delete("/{rate_id}")
def delete_rate(rate_id: int):
    ok = rate_service.delete_rate(rate_id)
    if not ok:
        raise HTTPException(status_code=404, detail="요율을 찾을 수 없습니다.")
    return {"ok": True}

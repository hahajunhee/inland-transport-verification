from sqlalchemy.orm import Session
from sqlalchemy import and_, or_
from typing import Optional
from app.models import TransportRate


def find_rate(
    db: Session,
    charge_type: str,
    pickup_code: Optional[str],
    odcy_code: Optional[str],
    dest_code: Optional[str],
    container_type: Optional[str],
) -> Optional[TransportRate]:
    """
    charge_type 필수, 나머지 필드는 exact match 또는 DB에 NULL(any) 등록된 것과 매칭.
    더 구체적인 규칙(NULL 필드 수 적은 것) 우선 반환.
    """
    candidates = (
        db.query(TransportRate)
        .filter(TransportRate.charge_type == charge_type)
        .filter(
            or_(TransportRate.pickup_code == pickup_code, TransportRate.pickup_code == None)
        )
        .filter(
            or_(TransportRate.odcy_code == odcy_code, TransportRate.odcy_code == None)
        )
        .filter(
            or_(TransportRate.dest_code == dest_code, TransportRate.dest_code == None)
        )
        .filter(
            or_(TransportRate.container_type == container_type, TransportRate.container_type == None)
        )
        .all()
    )

    if not candidates:
        return None

    # 가장 구체적인 매칭 우선 (NULL 필드 수가 적은 것)
    def specificity(rate: TransportRate) -> int:
        score = 0
        if rate.pickup_code is not None:
            score += 1
        if rate.odcy_code is not None:
            score += 1
        if rate.dest_code is not None:
            score += 1
        if rate.container_type is not None:
            score += 1
        return score

    candidates.sort(key=specificity, reverse=True)
    return candidates[0]


def get_all_rates(db: Session, charge_type: Optional[str] = None, pickup_code: Optional[str] = None, dest_code: Optional[str] = None):
    q = db.query(TransportRate)
    if charge_type:
        q = q.filter(TransportRate.charge_type == charge_type)
    if pickup_code:
        q = q.filter(TransportRate.pickup_code == pickup_code)
    if dest_code:
        q = q.filter(TransportRate.dest_code == dest_code)
    return q.order_by(TransportRate.charge_type, TransportRate.id).all()


def create_rate(db: Session, data: dict) -> TransportRate:
    rate = TransportRate(**data)
    db.add(rate)
    db.commit()
    db.refresh(rate)
    return rate


def update_rate(db: Session, rate_id: int, data: dict) -> Optional[TransportRate]:
    rate = db.query(TransportRate).filter(TransportRate.id == rate_id).first()
    if not rate:
        return None
    for k, v in data.items():
        if v is not None:
            setattr(rate, k, v)
    db.commit()
    db.refresh(rate)
    return rate


def delete_rate(db: Session, rate_id: int) -> bool:
    rate = db.query(TransportRate).filter(TransportRate.id == rate_id).first()
    if not rate:
        return False
    db.delete(rate)
    db.commit()
    return True

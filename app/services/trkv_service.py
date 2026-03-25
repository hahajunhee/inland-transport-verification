from typing import Optional
from sqlalchemy.orm import Session
from app.models import TRKVPortMapping, TRKVRoute, TRKVContainerTier


# ─── 포트명 매핑 ──────────────────────────────────────────────────────

def resolve_port(db: Session, name: Optional[str]) -> Optional[str]:
    """엑셀 포트명 → 부산신항/부산북항. 매핑 없으면 원본 그대로 반환."""
    if not name:
        return name
    mapping = db.query(TRKVPortMapping).filter(TRKVPortMapping.excel_name == name.strip()).first()
    return mapping.port_type if mapping else name.strip()


def get_all_port_mappings(db: Session):
    return db.query(TRKVPortMapping).order_by(TRKVPortMapping.id).all()


def create_port_mapping(db: Session, excel_name: str, port_type: str) -> TRKVPortMapping:
    obj = TRKVPortMapping(excel_name=excel_name.strip(), port_type=port_type)
    db.add(obj)
    db.commit()
    db.refresh(obj)
    return obj


def update_port_mapping(db: Session, mapping_id: int, excel_name: str, port_type: str) -> Optional[TRKVPortMapping]:
    obj = db.query(TRKVPortMapping).filter(TRKVPortMapping.id == mapping_id).first()
    if not obj:
        return None
    obj.excel_name = excel_name.strip()
    obj.port_type = port_type
    db.commit()
    db.refresh(obj)
    return obj


def delete_port_mapping(db: Session, mapping_id: int) -> bool:
    obj = db.query(TRKVPortMapping).filter(TRKVPortMapping.id == mapping_id).first()
    if not obj:
        return False
    db.delete(obj)
    db.commit()
    return True


# ─── 구간요율 ─────────────────────────────────────────────────────────

def get_all_routes(db: Session):
    return db.query(TRKVRoute).order_by(TRKVRoute.pickup_port, TRKVRoute.departure_name, TRKVRoute.dest_port).all()


def create_route(db: Session, data: dict) -> TRKVRoute:
    obj = TRKVRoute(**data)
    db.add(obj)
    db.commit()
    db.refresh(obj)
    return obj


def update_route(db: Session, route_id: int, data: dict) -> Optional[TRKVRoute]:
    obj = db.query(TRKVRoute).filter(TRKVRoute.id == route_id).first()
    if not obj:
        return None
    for k, v in data.items():
        setattr(obj, k, v)
    db.commit()
    db.refresh(obj)
    return obj


def delete_route(db: Session, route_id: int) -> bool:
    obj = db.query(TRKVRoute).filter(TRKVRoute.id == route_id).first()
    if not obj:
        return False
    db.delete(obj)
    db.commit()
    return True


# ─── 컨테이너 티어 ───────────────────────────────────────────────────

def get_all_container_tiers(db: Session):
    return db.query(TRKVContainerTier).order_by(TRKVContainerTier.cont_type, TRKVContainerTier.is_dg).all()


def bulk_save_container_tiers(db: Session, items: list[dict]) -> list[TRKVContainerTier]:
    """[{cont_type, is_dg, tier_number}, ...] 일괄 저장 (upsert)"""
    results = []
    for item in items:
        obj = (
            db.query(TRKVContainerTier)
            .filter(
                TRKVContainerTier.cont_type == item["cont_type"],
                TRKVContainerTier.is_dg == item["is_dg"],
            )
            .first()
        )
        if obj:
            obj.tier_number = item.get("tier_number")
        else:
            obj = TRKVContainerTier(**item)
            db.add(obj)
        results.append(obj)
    db.commit()
    for r in results:
        db.refresh(r)
    return results


def update_container_tier(db: Session, tier_id: int, tier_number: Optional[int]) -> Optional[TRKVContainerTier]:
    obj = db.query(TRKVContainerTier).filter(TRKVContainerTier.id == tier_id).first()
    if not obj:
        return None
    obj.tier_number = tier_number
    db.commit()
    db.refresh(obj)
    return obj


# ─── 핵심 요율 조회 ──────────────────────────────────────────────────

def get_trkv_expected(
    db: Session,
    pickup_name: Optional[str],
    departure_name: Optional[str],
    dest_name: Optional[str],
    cont_type: Optional[str],
    dg_raw: Optional[str],
) -> Optional[float]:
    """
    TRKV 예상 금액 반환. 설정 누락 시 None 반환 → NO_RATE 처리.

    1. pickup_name / dest_name → 포트 매핑으로 부산신항/부산북항 해석
    2. (cont_type, dg_raw) → TRKVContainerTier.tier_number 조회
    3. (pickup_port, departure_name, dest_port) → TRKVRoute 조회 후 tier{N} 반환
    """
    # 1. 포트 해석
    pickup_port = resolve_port(db, pickup_name)
    dest_port   = resolve_port(db, dest_name)

    # 2. D/G 판단 (엑셀 원본값 "X"이면 DG)
    is_dg = str(dg_raw or "").strip().upper() == "X"

    # 3. 컨테이너 티어 조회
    ct = str(cont_type or "").strip()
    tier_row = (
        db.query(TRKVContainerTier)
        .filter(TRKVContainerTier.cont_type == ct, TRKVContainerTier.is_dg == is_dg)
        .first()
    )
    if not tier_row or tier_row.tier_number is None:
        return None

    tier_num = tier_row.tier_number  # 1~6

    # 4. 구간요율 조회
    dep = str(departure_name or "").strip()
    route = (
        db.query(TRKVRoute)
        .filter(
            TRKVRoute.pickup_port == pickup_port,
            TRKVRoute.departure_name == dep,
            TRKVRoute.dest_port == dest_port,
        )
        .first()
    )
    if not route:
        return None

    # 5. 티어번호에 해당하는 단가 반환
    price = getattr(route, f"tier{tier_num}", None)
    return price  # None이면 NO_RATE

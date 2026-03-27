from typing import Optional
from app import data_store


# ─── 포트명 매핑 ──────────────────────────────────────────────────────

def resolve_port(name: Optional[str]) -> Optional[str]:
    """엑셀 포트명 → 부산신항/부산북항. 매핑 없으면 원본 그대로 반환."""
    if not name:
        return name
    name = name.strip()
    items = data_store.load("port_mappings.json")
    for m in items:
        if m["excel_name"] == name:
            return m["port_type"]
    return name


def get_all_port_mappings() -> list:
    return sorted(data_store.load("port_mappings.json"), key=lambda x: x["id"])


def create_port_mapping(excel_name: str, port_type: str) -> dict:
    items = data_store.load("port_mappings.json")
    # 중복 체크
    if any(m["excel_name"] == excel_name.strip() for m in items):
        raise ValueError("이미 등록된 포트명입니다.")
    obj = {
        "id": data_store.next_id(items),
        "excel_name": excel_name.strip(),
        "port_type": port_type,
    }
    items.append(obj)
    data_store.save("port_mappings.json", items)
    return obj


def update_port_mapping(mapping_id: int, excel_name: str, port_type: str) -> Optional[dict]:
    items = data_store.load("port_mappings.json")
    for i, m in enumerate(items):
        if m["id"] == mapping_id:
            items[i]["excel_name"] = excel_name.strip()
            items[i]["port_type"] = port_type
            data_store.save("port_mappings.json", items)
            return items[i]
    return None


def delete_port_mapping(mapping_id: int) -> bool:
    items = data_store.load("port_mappings.json")
    new_items = [m for m in items if m["id"] != mapping_id]
    if len(new_items) == len(items):
        return False
    data_store.save("port_mappings.json", new_items)
    return True


# ─── 구간요율 ─────────────────────────────────────────────────────────

def get_all_routes() -> list:
    items = data_store.load("trkv_routes.json")
    return sorted(items, key=lambda x: (x.get("pickup_port", ""), x.get("departure_name", ""), x.get("dest_port", "")))


def create_route(data: dict) -> dict:
    items = data_store.load("trkv_routes.json")
    obj = {"id": data_store.next_id(items), **data}
    items.append(obj)
    data_store.save("trkv_routes.json", items)
    return obj


def update_route(route_id: int, data: dict) -> Optional[dict]:
    items = data_store.load("trkv_routes.json")
    for i, r in enumerate(items):
        if r["id"] == route_id:
            items[i].update(data)
            data_store.save("trkv_routes.json", items)
            return items[i]
    return None


def delete_route(route_id: int) -> bool:
    items = data_store.load("trkv_routes.json")
    new_items = [r for r in items if r["id"] != route_id]
    if len(new_items) == len(items):
        return False
    data_store.save("trkv_routes.json", new_items)
    return True


# ─── 컨테이너 티어 ───────────────────────────────────────────────────

def get_all_container_tiers() -> list:
    items = data_store.load("container_tiers.json")
    return sorted(items, key=lambda x: (x.get("cont_type", ""), x.get("is_dg", False)))


def bulk_save_container_tiers(new_items: list) -> list:
    """[{cont_type, is_dg, tier_number}, ...] 일괄 저장 (upsert)"""
    items = data_store.load("container_tiers.json")
    results = []
    for new in new_items:
        existing = next(
            (x for x in items if x["cont_type"] == new["cont_type"] and x["is_dg"] == new["is_dg"]),
            None,
        )
        if existing:
            existing["tier_number"] = new.get("tier_number")
            results.append(existing)
        else:
            obj = {"id": data_store.next_id(items), **new}
            items.append(obj)
            results.append(obj)
    data_store.save("container_tiers.json", items)
    return results


def update_container_tier(tier_id: int, tier_number: Optional[int]) -> Optional[dict]:
    items = data_store.load("container_tiers.json")
    for i, t in enumerate(items):
        if t["id"] == tier_id:
            items[i]["tier_number"] = tier_number
            data_store.save("container_tiers.json", items)
            return items[i]
    return None


# ─── 핵심 요율 조회 ──────────────────────────────────────────────────

def get_trkv_expected(
    pickup_name: Optional[str],
    departure_name: Optional[str],
    dest_name: Optional[str],
    cont_type: Optional[str],
    dg_raw: Optional[str],
) -> Optional[float]:
    """
    TRKV 예상 금액 반환. 설정 누락 시 None 반환 → NO_RATE 처리.

    1. pickup_name / dest_name → 포트 매핑으로 부산신항/부산북항 해석
    2. (cont_type, dg_raw) → container_tiers.json 에서 tier_number 조회
    3. (pickup_port, departure_name, dest_port) → trkv_routes.json 조회 후 tier{N} 반환
    """
    # 1. 포트 해석
    pickup_port = resolve_port(pickup_name)
    dest_port   = resolve_port(dest_name)

    # 2. D/G 판단
    is_dg = str(dg_raw or "").strip().upper() == "X"

    # 3. 컨테이너 티어 조회
    ct = str(cont_type or "").strip()
    tiers = data_store.load("container_tiers.json")
    tier_row = next(
        (t for t in tiers if t["cont_type"] == ct and t["is_dg"] == is_dg),
        None,
    )
    if not tier_row or tier_row.get("tier_number") is None:
        return None

    tier_num = tier_row["tier_number"]  # 1~6

    # 4. 구간요율 조회
    dep = str(departure_name or "").strip()
    routes = data_store.load("trkv_routes.json")
    route = next(
        (r for r in routes
         if r["pickup_port"] == pickup_port
         and r["departure_name"] == dep
         and r["dest_port"] == dest_port),
        None,
    )
    if not route:
        return None

    # 5. 티어번호에 해당하는 단가 반환
    price = route.get(f"tier{tier_num}")
    return price  # None이면 NO_RATE

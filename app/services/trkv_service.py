from typing import Optional
from app import data_store


# ─── 포트명 매핑 ──────────────────────────────────────────────────────

def resolve_port(name: Optional[str]) -> Optional[str]:
    """엑셀 포트명 → 포트구분. 매핑 없으면 원본 그대로 반환."""
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


# ─── 출하지 매핑 ──────────────────────────────────────────────────────

def resolve_departure(name: Optional[str]) -> Optional[str]:
    """엑셀 출하지명 → 출하지코드. 매핑 없으면 원본 그대로 반환."""
    if not name:
        return name
    name = name.strip()
    items = data_store.load("departure_mappings.json")
    for m in items:
        if m["departure_name"] == name:
            return m["departure_code"]
    return name


def get_all_departure_mappings() -> list:
    return sorted(data_store.load("departure_mappings.json"), key=lambda x: x["id"])


def create_departure_mapping(departure_name: str, departure_code: str) -> dict:
    items = data_store.load("departure_mappings.json")
    if any(m["departure_name"] == departure_name.strip() for m in items):
        raise ValueError("이미 등록된 출하지명입니다.")
    obj = {
        "id": data_store.next_id(items),
        "departure_name": departure_name.strip(),
        "departure_code": departure_code.strip(),
    }
    items.append(obj)
    data_store.save("departure_mappings.json", items)
    return obj


def update_departure_mapping(mapping_id: int, departure_name: str, departure_code: str) -> Optional[dict]:
    items = data_store.load("departure_mappings.json")
    for i, m in enumerate(items):
        if m["id"] == mapping_id:
            items[i]["departure_name"] = departure_name.strip()
            items[i]["departure_code"] = departure_code.strip()
            data_store.save("departure_mappings.json", items)
            return items[i]
    return None


def delete_departure_mapping(mapping_id: int) -> bool:
    items = data_store.load("departure_mappings.json")
    new_items = [m for m in items if m["id"] != mapping_id]
    if len(new_items) == len(items):
        return False
    data_store.save("departure_mappings.json", new_items)
    return True


# ─── 구간요율 ─────────────────────────────────────────────────────────

def get_all_routes() -> list:
    items = data_store.load("trkv_routes.json")
    return sorted(items, key=lambda x: (x.get("pickup_port", ""), x.get("departure_code", x.get("departure_name", "")), x.get("dest_port", "")))


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

    1. pickup_name → resolve_port  → pickup_port
    2. departure_name → resolve_departure → departure_code
    3. dest_name  → resolve_port  → dest_port
    4. (cont_type, dg_raw) → container_tiers → tier_number
    5. (pickup_port, departure_code, dest_port) → trkv_routes → tier{N} 단가
    """
    # 1. 포트 해석
    pickup_port = resolve_port(pickup_name)
    dest_port   = resolve_port(dest_name)

    # 2. 출하지 코드 해석
    departure_code = resolve_departure(departure_name)

    # 3. D/G 판단
    is_dg = str(dg_raw or "").strip().upper() == "X"

    # 4. 컨테이너 티어 조회
    ct = str(cont_type or "").strip()
    tiers = data_store.load("container_tiers.json")
    tier_row = next(
        (t for t in tiers if t["cont_type"] == ct and t["is_dg"] == is_dg),
        None,
    )
    if not tier_row or tier_row.get("tier_number") is None:
        return None

    tier_num = tier_row["tier_number"]  # 1~6

    # 5. 구간요율 조회 (departure_code 기준, 구버전 departure_name 호환)
    dep = str(departure_code or "").strip()
    routes = data_store.load("trkv_routes.json")
    route = next(
        (r for r in routes
         if r.get("pickup_port") == pickup_port
         and r.get("departure_code", r.get("departure_name", "")) == dep
         and r.get("dest_port") == dest_port),
        None,
    )
    if not route:
        return None

    price = route.get(f"tier{tier_num}")
    return price

from typing import Optional
from app import data_store

TIERS = [1, 2, 3, 4, 5, 6]
TIER_FIELDS = [f"tier{t}" for t in TIERS]


def find_storage_rate(
    odcy_name: Optional[str],
    odcy_terminal_type: Optional[str],
    odcy_location: Optional[str],
    dest_port_type: Optional[str],
    dest_terminal_type: Optional[str],
    tier_number: Optional[int],
    om_a: Optional[str] = None,
) -> dict:
    """
    보관료/상하차료/셔틀비 요율 조회.
    OM-A(ODCY도착지명)를 첫 번째 키로 정확 매칭한 뒤,
    나머지 5키(ODCY명, odcy터미널구분, ODCY_위치, 포트구분, 터미널구분)로 세부 매칭.
    반환: {"storage_unit": ..., "handling_unit": ..., "shuttle_unit": ...}
    """
    items = data_store.load("storage_rates.json")

    def matches_sub(r: dict) -> bool:
        if odcy_name and r.get("odcy_name") and r["odcy_name"] != odcy_name:
            return False
        if odcy_terminal_type and r.get("odcy_terminal_type") and r["odcy_terminal_type"] != odcy_terminal_type:
            return False
        if odcy_location and r.get("odcy_location") and r["odcy_location"] != odcy_location:
            return False
        if dest_port_type and r.get("dest_port_type") and r["dest_port_type"] != dest_port_type:
            return False
        if dest_terminal_type and r.get("dest_terminal_type") and r["dest_terminal_type"] != dest_terminal_type:
            return False
        return True

    # 1단계: OM-A 정확 매칭 (첫 번째 키)
    candidates = []
    if om_a:
        candidates = [r for r in items if r.get("om_a") == om_a and matches_sub(r)]

    # 2단계: OM-A 매칭 결과 없으면 기존 5키로 폴백
    if not candidates:
        candidates = [r for r in items if matches_sub(r)]

    if not candidates:
        return {}

    def specificity(r: dict) -> int:
        score = 0
        if r.get("om_a") and om_a and r["om_a"] == om_a:
            score += 16
        if r.get("odcy_name"):           score += 4
        if r.get("odcy_terminal_type"):   score += 2
        if r.get("odcy_location"):        score += 2
        if r.get("dest_port_type"):       score += 1
        if r.get("dest_terminal_type"):   score += 1
        return score

    candidates.sort(key=specificity, reverse=True)
    best = candidates[0]

    # best의 행번호 찾기 (ID 순서 기준, 헤더 2행 → 데이터는 3행부터)
    rate_row_num = None
    for idx, r in enumerate(items, 3):  # 헤더 2행 → 데이터는 3행부터
        if r.get("id") == best.get("id"):
            rate_row_num = idx
            break

    t = tier_number if isinstance(tier_number, int) and tier_number in TIERS else 1

    # 티어별 필드 (storage_tier1) 우선, fallback으로 단일 단가 필드 (storage_unit)
    def _get_unit(prefix):
        val = best.get(f"{prefix}_tier{t}")
        if val is not None:
            return val
        # fallback: 구버전 단일 단가 필드
        return best.get(f"{prefix}_unit")

    return {
        "storage_unit":  _get_unit("storage"),
        "handling_unit": _get_unit("handling"),
        "shuttle_unit":  _get_unit("shuttle"),
        "rate_row_num":  rate_row_num,
    }


def get_all_storage_rates() -> list:
    return sorted(
        data_store.load("storage_rates.json"),
        key=lambda x: (x.get("odcy_name", ""), x.get("odcy_terminal_type", ""), x["id"]),
    )


def create_storage_rate(data: dict) -> dict:
    items = data_store.load("storage_rates.json")
    obj = {"id": data_store.next_id(items), **data}
    items.append(obj)
    data_store.save("storage_rates.json", items)
    return obj


def update_storage_rate(rate_id: int, data: dict) -> Optional[dict]:
    items = data_store.load("storage_rates.json")
    for i, r in enumerate(items):
        if r["id"] == rate_id:
            items[i].update(data)
            data_store.save("storage_rates.json", items)
            return items[i]
    return None


def delete_storage_rate(rate_id: int) -> bool:
    items = data_store.load("storage_rates.json")
    new_items = [r for r in items if r["id"] != rate_id]
    if len(new_items) == len(items):
        return False
    data_store.save("storage_rates.json", new_items)
    return True

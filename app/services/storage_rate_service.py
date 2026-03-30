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
) -> dict:
    """
    5개 키 (ODCY명, odcy터미널구분, ODCY_위치, 포트구분, 터미널구분) 기준으로
    보관료/상하차료/셔틀비 티어 요율 조회.
    더 구체적인(필드 수 많은) 레코드 우선 반환.
    반환: {"storage_unit": ..., "handling_unit": ..., "shuttle_unit": ...}
    """
    items = data_store.load("storage_rates.json")

    def matches(r: dict) -> bool:
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

    candidates = [r for r in items if matches(r)]
    if not candidates:
        return {}

    def specificity(r: dict) -> int:
        score = 0
        if r.get("odcy_name"):           score += 4
        if r.get("odcy_terminal_type"):   score += 2
        if r.get("odcy_location"):        score += 2
        if r.get("dest_port_type"):       score += 1
        if r.get("dest_terminal_type"):   score += 1
        return score

    candidates.sort(key=specificity, reverse=True)
    best = candidates[0]

    t = tier_number if isinstance(tier_number, int) and tier_number in TIERS else 1
    return {
        "storage_unit":  best.get(f"storage_tier{t}"),
        "handling_unit": best.get(f"handling_tier{t}"),
        "shuttle_unit":  best.get(f"shuttle_tier{t}"),
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

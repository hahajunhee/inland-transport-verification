from typing import Optional
from app import data_store


def find_storage_rate(odcy_name: Optional[str], zone_type: Optional[str]) -> Optional[dict]:
    """
    (ODCY명, 단지구분) 기준으로 보관료/상하차료 요율 조회.
    odcy_name 또는 zone_type 가 None/"" 이면 해당 필드 무시(any 매치).
    더 구체적인(필드 수 많은) 레코드 우선 반환.
    """
    items = data_store.load("storage_rates.json")

    def matches(r: dict) -> bool:
        if odcy_name and r.get("odcy_name") and r["odcy_name"] != odcy_name:
            return False
        if zone_type and r.get("zone_type") and r["zone_type"] != zone_type:
            return False
        return True

    candidates = [r for r in items if matches(r)]
    if not candidates:
        return None

    def specificity(r: dict) -> int:
        score = 0
        if r.get("odcy_name"):
            score += 1
        if r.get("zone_type"):
            score += 1
        return score

    candidates.sort(key=specificity, reverse=True)
    return candidates[0]


def get_all_storage_rates() -> list:
    return sorted(data_store.load("storage_rates.json"), key=lambda x: (x.get("odcy_name", ""), x.get("zone_type", ""), x["id"]))


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

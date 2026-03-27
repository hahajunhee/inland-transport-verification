from typing import Optional
from app import data_store


def find_rate(
    charge_type: str,
    pickup_code: Optional[str],
    odcy_code: Optional[str],
    dest_code: Optional[str],
    container_type: Optional[str],
) -> Optional[dict]:
    """
    charge_type 필수, 나머지 필드는 exact match 또는 None(any) 등록된 것과 매칭.
    더 구체적인 규칙(None 필드 수 적은 것) 우선 반환.
    """
    all_rates = data_store.load("transport_rates.json")

    def matches(r: dict) -> bool:
        if r.get("charge_type") != charge_type:
            return False
        if r.get("pickup_code") is not None and r.get("pickup_code") != pickup_code:
            return False
        if r.get("odcy_code") is not None and r.get("odcy_code") != odcy_code:
            return False
        if r.get("dest_code") is not None and r.get("dest_code") != dest_code:
            return False
        if r.get("container_type") is not None and r.get("container_type") != container_type:
            return False
        return True

    candidates = [r for r in all_rates if matches(r)]
    if not candidates:
        return None

    def specificity(r: dict) -> int:
        score = 0
        if r.get("pickup_code") is not None:
            score += 1
        if r.get("odcy_code") is not None:
            score += 1
        if r.get("dest_code") is not None:
            score += 1
        if r.get("container_type") is not None:
            score += 1
        return score

    candidates.sort(key=specificity, reverse=True)
    return candidates[0]


def get_all_rates(
    charge_type: Optional[str] = None,
    pickup_code: Optional[str] = None,
    dest_code: Optional[str] = None,
) -> list:
    items = data_store.load("transport_rates.json")
    if charge_type:
        items = [r for r in items if r.get("charge_type") == charge_type]
    if pickup_code:
        items = [r for r in items if r.get("pickup_code") == pickup_code]
    if dest_code:
        items = [r for r in items if r.get("dest_code") == dest_code]
    return sorted(items, key=lambda x: (x.get("charge_type", ""), x.get("id", 0)))


def create_rate(data: dict) -> dict:
    items = data_store.load("transport_rates.json")
    obj = {"id": data_store.next_id(items), **data}
    items.append(obj)
    data_store.save("transport_rates.json", items)
    return obj


def update_rate(rate_id: int, data: dict) -> Optional[dict]:
    items = data_store.load("transport_rates.json")
    for i, r in enumerate(items):
        if r["id"] == rate_id:
            for k, v in data.items():
                if v is not None:
                    items[i][k] = v
            data_store.save("transport_rates.json", items)
            return items[i]
    return None


def delete_rate(rate_id: int) -> bool:
    items = data_store.load("transport_rates.json")
    new_items = [r for r in items if r["id"] != rate_id]
    if len(new_items) == len(items):
        return False
    data_store.save("transport_rates.json", new_items)
    return True

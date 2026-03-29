from datetime import datetime
from app import data_store
from app.services.rate_service import find_rate
from app.services import trkv_service
from app.services.trkv_service import resolve_port, resolve_departure, resolve_odcy_name, resolve_zone_type, get_trkv_details
from app.services.storage_rate_service import find_storage_rate

TOLERANCE = 1.0  # 원 단위 허용 오차

CHARGES = [
    ("TRKV",   "trkv_actual",    "trkv_expected",    "trkv_diff",    "trkv_status"),
    ("보관료",  "storage_actual", "storage_expected", "storage_diff", "storage_status"),
    ("상하차료", "handling_actual", "handling_expected", "handling_diff", "handling_status"),
    ("셔틀비용", "shuttle_actual", "shuttle_expected",  "shuttle_diff",  "shuttle_status"),
]


def _verify_charge(charge_type, actual, pickup_code, odcy_code, dest_code, container_type,
                   pickup_name=None, departure_name=None, dest_name=None,
                   cont_type=None, dg_raw=None, quantity=1.0, weekend_holiday="",
                   odcy_name_resolved=None, zone_type=None):
    if charge_type == "TRKV":
        expected = trkv_service.get_trkv_expected(
            pickup_name, departure_name, dest_name, cont_type, dg_raw, quantity, weekend_holiday
        )
    elif charge_type in ("보관료", "상하차료"):
        rate = find_storage_rate(odcy_name_resolved, zone_type)
        if rate:
            expected = rate.get("storage_unit") if charge_type == "보관료" else rate.get("handling_unit")
        else:
            expected = None
    else:
        rate = find_rate(charge_type, pickup_code, odcy_code, dest_code, container_type)
        expected = rate.get("unit_price") if rate else None

    if expected is None:
        if actual == 0.0:
            return None, None, "SKIP"
        return None, None, "NO_RATE"
    diff = actual - expected
    status = "OK" if abs(diff) < TOLERANCE else "DIFF"
    return expected, diff, status


def run_verification(filename: str, rows: list) -> dict:
    sessions = data_store.load("verification_sessions.json")
    session_id = data_store.next_id(sessions)

    session = {
        "id": session_id,
        "filename": filename,
        "uploaded_at": datetime.now().isoformat(),
        "total_rows": len(rows),
        "trkv_pass": 0, "trkv_fail": 0, "trkv_no_rate": 0,
        "storage_pass": 0, "storage_fail": 0, "storage_no_rate": 0,
        "handling_pass": 0, "handling_fail": 0, "handling_no_rate": 0,
        "shuttle_pass": 0, "shuttle_fail": 0, "shuttle_no_rate": 0,
        "total_diff": 0.0,
    }

    prefix_map = {"TRKV": "trkv", "보관료": "storage", "상하차료": "handling", "셔틀비용": "shuttle"}
    total_diff = 0.0
    results = []
    result_id = 1

    for row in rows:
        pickup_code    = row.get("pickup_code")
        odcy_code      = row.get("odcy_code")
        dest_code      = row.get("dest_code")
        container_type = row.get("container_type")
        pickup_name    = row.get("pickup_name")
        departure_name = row.get("departure_name")
        dest_name      = row.get("dest_name")
        cont_type              = row.get("cont_type")
        dg_raw                 = row.get("dg_raw")
        quantity               = float(row.get("quantity") or 1.0)
        weekend_holiday        = str(row.get("weekend_holiday") or "").strip().upper()
        odcy_destination_name  = row.get("odcy_destination_name")
        odcy_name_resolved     = resolve_odcy_name(odcy_destination_name or row.get("odcy_name"))
        zone_type              = resolve_zone_type(dest_name)

        result = {
            "id": result_id,
            "session_id": session_id,
            "row_number": row.get("row_number", 0),
            "container_no": row.get("container_no"),
            "transport_date": row.get("transport_date"),
            "pickup_code": pickup_code,
            "pickup_name": pickup_name,
            "pickup_port_resolved": resolve_port(pickup_name),
            "odcy_code": odcy_code,
            "odcy_name": row.get("odcy_name"),
            "departure_name": departure_name,
            "departure_code_resolved": resolve_departure(departure_name),
            "dest_code": dest_code,
            "dest_name": dest_name,
            "dest_port_resolved": resolve_port(dest_name),
            "container_type": container_type,
            "dg_flag": row.get("dg_flag", False),
            "quantity": quantity,
            "weekend_holiday": weekend_holiday,
            "odcy_destination_name": odcy_destination_name,
            "odcy_name_resolved": odcy_name_resolved,
            "zone_type": zone_type,
        }

        # 티어번호 + 단가 조회 (운송 구간 정보에 표시용)
        trkv_details = get_trkv_details(
            pickup_name, departure_name, dest_name, cont_type, dg_raw, quantity, weekend_holiday
        )
        result["tier_number"]   = trkv_details.get("tier_number")
        result["trkv_unit_rate"] = trkv_details.get("unit_rate")

        result_id += 1

        statuses = []
        for (charge_type, actual_key, exp_key, diff_key, status_key) in CHARGES:
            actual = row.get(actual_key, 0.0)
            expected, diff, status = _verify_charge(
                charge_type, actual, pickup_code, odcy_code, dest_code, container_type,
                pickup_name=pickup_name, departure_name=departure_name, dest_name=dest_name,
                cont_type=cont_type, dg_raw=dg_raw, quantity=quantity,
                weekend_holiday=weekend_holiday,
                odcy_name_resolved=odcy_name_resolved, zone_type=zone_type,
            )
            result[actual_key] = actual
            result[exp_key] = expected
            result[diff_key] = diff
            result[status_key] = status
            statuses.append(status)

            prefix = prefix_map[charge_type]
            if status in ("OK", "SKIP"):
                session[f"{prefix}_pass"] += 1
            elif status == "DIFF":
                session[f"{prefix}_fail"] += 1
                total_diff += abs(diff or 0)
            elif status == "NO_RATE":
                session[f"{prefix}_no_rate"] += 1

        # 종합 상태
        if all(s in ("OK", "SKIP") for s in statuses):
            overall = "OK"
        elif "NO_RATE" in statuses:
            overall = "NO_RATE"
        else:
            overall = "DIFF"

        result["overall_status"] = overall
        results.append(result)

    session["total_diff"] = total_diff

    sessions.append(session)
    data_store.save("verification_sessions.json", sessions)
    data_store.save_results(session_id, results)
    return session

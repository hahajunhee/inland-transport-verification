from datetime import datetime, date
from app import data_store
from app.services.rate_service import find_rate
from app.services import trkv_service
from app.services.trkv_service import (
    resolve_port, resolve_port_terminal_type, resolve_departure,
    resolve_odcy_name, resolve_terminal_type, resolve_odcy_location,
    get_trkv_details, get_storage_tier_number,
)
from app.services.storage_rate_service import find_storage_rate

TOLERANCE = 1.0  # 원 단위 허용 오차

CHARGES = [
    ("TRKV",   "trkv_actual",    "trkv_expected",    "trkv_diff",    "trkv_status"),
    ("보관료",  "storage_actual", "storage_expected", "storage_diff", "storage_status"),
    ("상하차료", "handling_actual", "handling_expected", "handling_diff", "handling_status"),
    ("셔틀비용", "shuttle_actual", "shuttle_expected",  "shuttle_diff",  "shuttle_status"),
]


def _parse_date_value(val) -> date | None:
    """날짜 문자열 또는 datetime 객체를 date로 변환."""
    if val is None:
        return None
    if isinstance(val, date) and not isinstance(val, datetime):
        return val
    if isinstance(val, datetime):
        return val.date()
    s = str(val).strip()
    if not s or s in ("nan", "None", "NaT"):
        return None
    # 다양한 날짜 형식 시도
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y%m%d"):
        try:
            return datetime.strptime(s[:10], fmt).date()
        except ValueError:
            continue
    return None


FREE_TIME_LOCATIONS = {"부산신항", "KRPUSN"}
FREE_TIME_DAYS = 3

def _get_free_days(odcy_location: str) -> int:
    """FREE타임 적용 일수 반환."""
    if odcy_location and odcy_location.strip() in FREE_TIME_LOCATIONS:
        return FREE_TIME_DAYS
    return 0

FREE_TIER_NUMBERS = {1, 2}  # FREE 적용 대상 보관료 티어

def _calc_storage_days(odcy_in_date_str, odcy_out_date_str, odcy_location: str,
                       storage_tier_number: int | None = None) -> tuple[int | None, int | None, int]:
    """보관일수 계산. 반환: (raw_days, billable_days, free_days)
    - raw_days: 반출일 - 반입일 + 1 (순수 보관일수, 표시용)
    - billable_days: max(raw_days - free_days, 0) (보관료 계산용)
    - free_days: FREE 적용 일수
    FREE 조건: ODCY위치가 KRPUSN/부산신항 AND 보관료티어가 T1 또는 T2
    보관료 = 단가(일) × billable_days × quantity
    """
    in_dt = _parse_date_value(odcy_in_date_str)
    out_dt = _parse_date_value(odcy_out_date_str)
    if in_dt is None or out_dt is None:
        return None, None, 0
    raw_days = (out_dt - in_dt).days + 1  # 순수 보관일수

    # FREE 적용: ODCY위치가 KRPUSN/부산신항 AND 티어 T1/T2일 때만 3일 차감
    free_days = 0
    if (odcy_location and odcy_location.strip() in FREE_TIME_LOCATIONS
            and storage_tier_number in FREE_TIER_NUMBERS):
        free_days = FREE_TIME_DAYS

    billable_days = max(raw_days - free_days, 0)
    return raw_days, billable_days, free_days


def _verify_charge(charge_type, actual, pickup_code, odcy_code, dest_code, container_type,
                   pickup_name=None, departure_name=None, dest_name=None,
                   cont_type=None, dg_raw=None, quantity=1.0, weekend_holiday="",
                   odcy_name_resolved=None, odcy_terminal_type=None,
                   odcy_location=None, dest_port_type=None, dest_terminal_type=None,
                   storage_tier_number=None, storage_days=None):
    """반환: (expected, diff, status, rate_row, unit_rate)"""
    rate_row = None
    unit_rate = None  # 보관료 day당 단가
    if charge_type == "TRKV":
        expected = trkv_service.get_trkv_expected(
            pickup_name, departure_name, dest_name, cont_type, dg_raw, quantity, weekend_holiday
        )
    elif charge_type in ("보관료", "상하차료", "셔틀비용"):
        rate = find_storage_rate(
            odcy_name_resolved, odcy_terminal_type, odcy_location,
            dest_port_type, dest_terminal_type, storage_tier_number,
        )
        rate_row = rate.get("rate_row_num")
        if charge_type == "보관료":
            unit = rate.get("storage_unit")
        elif charge_type == "상하차료":
            unit = rate.get("handling_unit")
        else:
            unit = rate.get("shuttle_unit")

        if charge_type == "보관료":
            unit_rate = unit  # day당 보관단가 기록

        if unit is not None:
            if charge_type == "보관료":
                # 보관료: 단가 × FREE반영일수 × 수량
                if storage_days is not None and storage_days >= 0:
                    expected = unit * storage_days * quantity
                else:
                    expected = None
            else:
                # 상하차료/셔틀비: 단가 × 수량 (보관일수 무관)
                expected = unit * quantity
        else:
            expected = None
    else:
        rate = find_rate(charge_type, pickup_code, odcy_code, dest_code, container_type)
        expected = rate.get("unit_price") if rate else None

    if expected is None:
        if actual == 0.0:
            return None, None, "SKIP", rate_row, unit_rate
        return None, None, "NO_RATE", rate_row, unit_rate
    diff = expected - actual
    status = "OK" if abs(diff) < TOLERANCE else "DIFF"
    return expected, diff, status, rate_row, unit_rate


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

        # ODCY 매핑 해석 (5개 키 중 3개: odcy_name, odcy_terminal_type, odcy_location)
        odcy_name_resolved     = resolve_odcy_name(odcy_destination_name or row.get("odcy_name"))
        odcy_terminal_type     = resolve_terminal_type(odcy_destination_name)
        odcy_location          = resolve_odcy_location(odcy_destination_name)

        # 도착지 포트 매핑 해석 (5개 키 중 2개: dest_port_type, dest_terminal_type)
        dest_port_type         = resolve_port(dest_name)
        dest_terminal_type     = resolve_port_terminal_type(dest_name)

        # 보관료/상하차료/셔틀비 전용 컨테이너 티어
        storage_tier_number    = get_storage_tier_number(cont_type, dg_raw)

        # 보관일수 계산
        odcy_in_date  = row.get("odcy_in_date")
        odcy_out_date = row.get("odcy_out_date")
        raw_days, billable_days, free_days = _calc_storage_days(odcy_in_date, odcy_out_date, odcy_location, storage_tier_number)

        result = {
            "id": result_id,
            "session_id": session_id,
            "row_number": row.get("row_number", 0),
            "container_no": row.get("container_no"),
            "fwo_doc": row.get("fwo_doc"),
            "c_invoice_no": row.get("c_invoice_no"),
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
            # 구분값 정보 (5개 키)
            "odcy_terminal_type": odcy_terminal_type,
            "odcy_location": odcy_location,
            "dest_port_type": dest_port_type,
            "dest_terminal_type": dest_terminal_type,
            # 보관료 전용 티어 + 일수
            "storage_tier_number": storage_tier_number,
            "odcy_in_date": odcy_in_date,
            "odcy_out_date": odcy_out_date,
            "storage_days": raw_days,
            "billable_days": billable_days,
            "free_days": free_days,
        }

        # 티어번호 + 단가 조회 (TRKV 운송 구간 정보에 표시용)
        trkv_details = get_trkv_details(
            pickup_name, departure_name, dest_name, cont_type, dg_raw, quantity, weekend_holiday
        )
        tier_number = trkv_details.get("tier_number")
        result["tier_number"]    = tier_number
        result["trkv_unit_rate"] = trkv_details.get("unit_rate")
        result["trkv_rate_row"]  = trkv_details.get("route_row_num")

        result_id += 1

        # 직반입 판정: ODCY도착지명이 공란이면 터미널 직반입건
        is_direct_delivery = not odcy_destination_name or str(odcy_destination_name).strip() == ""

        statuses = []
        storage_rate_row = None
        for (charge_type, actual_key, exp_key, diff_key, status_key) in CHARGES:
            actual = row.get(actual_key, 0.0)

            # 직반입건: 보관료/상하차료/셔틀비용은 0원이 정상
            if is_direct_delivery and charge_type in ("보관료", "상하차료", "셔틀비용"):
                expected = 0.0
                diff = expected - actual
                status = "OK" if abs(diff) < TOLERANCE else "DIFF"
                rate_row = None
                unit_rate = None
            else:
                expected, diff, status, rate_row, unit_rate = _verify_charge(
                    charge_type, actual, pickup_code, odcy_code, dest_code, container_type,
                    pickup_name=pickup_name, departure_name=departure_name, dest_name=dest_name,
                    cont_type=cont_type, dg_raw=dg_raw, quantity=quantity,
                    weekend_holiday=weekend_holiday,
                    odcy_name_resolved=odcy_name_resolved,
                    odcy_terminal_type=odcy_terminal_type,
                    odcy_location=odcy_location,
                    dest_port_type=dest_port_type,
                    dest_terminal_type=dest_terminal_type,
                    storage_tier_number=storage_tier_number,
                    storage_days=billable_days,
                )
            result[actual_key] = actual
            result[exp_key] = expected
            result[diff_key] = diff
            result[status_key] = status
            statuses.append(status)
            if charge_type == "보관료":
                storage_rate_row = rate_row
                result["storage_unit_rate"] = unit_rate

            prefix = prefix_map[charge_type]
            if status in ("OK", "SKIP"):
                session[f"{prefix}_pass"] += 1
            elif status == "DIFF":
                session[f"{prefix}_fail"] += 1
                total_diff += abs(diff or 0)
            elif status == "NO_RATE":
                session[f"{prefix}_no_rate"] += 1

        result["storage_rate_row"] = storage_rate_row

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

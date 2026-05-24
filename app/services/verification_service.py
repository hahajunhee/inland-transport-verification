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

# OM-D 코드 → 도착포트 매핑
BUSN_SINPORT_OMD = {"KRPUSN", "PUSN16", "PUSN7"}

def _resolve_om_d(odcy_name_val: str | None) -> str | None:
    """상세 ODCY명 → ODCY매핑(odcy_mappings)의 OM-A(odcy_destination_name)에서 매칭하여 OM-D(odcy_location) 반환."""
    if not odcy_name_val:
        return None
    name = odcy_name_val.strip()
    if not name:
        return None
    items = data_store.load("odcy_mappings.json")
    for m in items:
        if m.get("odcy_destination_name") == name:
            loc = m.get("odcy_location") or ""
            return loc.strip() if loc.strip() else None
    return None

def _resolve_dest_port_by_omd(om_d: str | None) -> str | None:
    """OM-D 값으로 도착포트 매핑.
    KRPUSN, PUSN16, PUSN7 → 부산신항 / 그 외 → 부산북항.
    """
    if not om_d:
        return None
    code = om_d.strip()
    if not code:
        return None
    if code in BUSN_SINPORT_OMD:
        return "부산신항"
    return "부산북항"

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
                   storage_tier_number=None, storage_days=None,
                   trkv_dest_name=None, trkv_dest_port=None,
                   odcy_destination_name=None):
    """반환: (expected, diff, status, rate_row, unit_rate)"""
    rate_row = None
    unit_rate = None  # 보관료 day당 단가
    if charge_type == "TRKV":
        expected = trkv_service.get_trkv_expected(
            pickup_name, departure_name, trkv_dest_name, cont_type, dg_raw, quantity, weekend_holiday,
            dest_port_override=trkv_dest_port,
        )
    elif charge_type in ("보관료", "상하차료", "셔틀비용"):
        rate = find_storage_rate(
            odcy_name_resolved, odcy_terminal_type, odcy_location,
            dest_port_type, dest_terminal_type, storage_tier_number,
            om_a=odcy_destination_name,
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
    data_store.begin_cache()
    try:
        return _run_verification_core(filename, rows)
    finally:
        data_store.end_cache()


def _run_verification_core(filename: str, rows: list) -> dict:
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

    # ── 매핑 데이터 일괄 프리로드: dict 기반 O(1) 조회 ──────────────────
    _pm = data_store.load("port_mappings.json")
    _om = data_store.load("odcy_mappings.json")
    _dm = data_store.load("departure_mappings.json")
    _ct = data_store.load("container_tiers.json")
    _st = data_store.load("storage_container_tiers.json")
    _tr = data_store.load("trkv_routes.json")

    port_map = {m["excel_name"].strip(): m for m in _pm}
    odcy_map = {m["odcy_destination_name"].strip(): m for m in _om}
    dep_map = {m["departure_name"].strip(): m["departure_code"] for m in _dm}
    tier_map = {(t["cont_type"], t["is_dg"]): t.get("tier_number") for t in _ct}
    stier_map = {(t["cont_type"], t["is_dg"]): t.get("tier_number") for t in _st}
    route_map = {}
    route_rows = {}
    for _i, _r in enumerate(_tr, 2):
        _k = (_r.get("pickup_port"),
              _r.get("departure_code", _r.get("departure_name", "")),
              _r.get("dest_port"))
        if _k not in route_map:
            route_map[_k] = _r
            route_rows[_k] = _i

    # ── 인라인 조회 헬퍼 (dict O(1)) ────────────────────────────────────

    def _rp(name):
        """resolve_port — O(1)"""
        if not name:
            return name
        pm = port_map.get(name.strip())
        return pm["port_type"] if pm else name

    def _rpt(name):
        """resolve_port_terminal_type — O(1)"""
        if not name:
            return ""
        pm = port_map.get(name.strip())
        return (pm.get("terminal_type") or "") if pm else ""

    def _rd(name):
        """resolve_departure — O(1)"""
        if not name:
            return name
        return dep_map.get(name.strip(), name)

    def _ron(name):
        """resolve_odcy_name — O(1)"""
        if not name:
            return name
        m = odcy_map.get(name.strip())
        return m["odcy_name"] if m else name

    def _rtt(name):
        """resolve_terminal_type — O(1)"""
        if not name:
            return ""
        m = odcy_map.get(name.strip())
        return (m.get("odcy_terminal_type") or m.get("terminal_type") or "") if m else ""

    def _rol(name):
        """resolve_odcy_location — O(1)"""
        if not name:
            return ""
        m = odcy_map.get(name.strip())
        return (m.get("odcy_location") or "") if m else ""

    def _omd(name):
        """_resolve_om_d — O(1)"""
        if not name:
            return None
        n = name.strip()
        if not n:
            return None
        m = odcy_map.get(n)
        if m:
            loc = m.get("odcy_location") or ""
            return loc.strip() if loc.strip() else None
        return None

    def _gstn(ct, dg):
        """get_storage_tier_number — O(1)"""
        c = str(ct or "").strip()
        d = str(dg or "").strip().upper() == "X"
        return stier_map.get((c, d))

    def _trkv_lookup(pp, dp, dc, ct, dg, qty, wh):
        """get_trkv_details 인라인 — O(1) dict 조회"""
        r = {"tier_number": None, "unit_rate": None, "expected": None, "route_row_num": None}
        ig = str(dg or "").strip().upper() == "X"
        c = str(ct or "").strip()
        tn = tier_map.get((c, ig))
        if tn is None:
            return r
        r["tier_number"] = tn
        dk = str(dc or "").strip()
        rk = (pp, dk, dp)
        route = route_map.get(rk)
        if not route:
            return r
        r["route_row_num"] = route_rows.get(rk)
        price = route.get(f"tier{tn}")
        if price is None:
            return r
        r["unit_rate"] = price
        if str(wh or "").strip().upper() == "X":
            r["expected"] = round(price * 1.2 * qty, -2)
        else:
            r["expected"] = price * qty
        return r

    # ── 행별 검증 루프 ──────────────────────────────────────────────────

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

        # ODCY 매핑 해석 — O(1) dict 조회
        odcy_name_resolved     = _ron(odcy_destination_name or row.get("odcy_name"))
        odcy_terminal_type     = _rtt(odcy_destination_name)
        odcy_location          = _rol(odcy_destination_name)

        # OM-D — O(1) dict 조회
        odcy_name_val          = row.get("odcy_name")
        om_d                   = _omd(odcy_name_val)

        # TRKV용 도착지/포트
        trkv_dest_name         = om_d
        trkv_dest_port         = _resolve_dest_port_by_omd(om_d)
        if trkv_dest_port is None and (not odcy_destination_name or str(odcy_destination_name).strip() == ""):
            trkv_dest_port = _rp(dest_name)
        dest_port_type         = _rp(dest_name)
        dest_terminal_type     = _rpt(dest_name)

        # 보관료 티어 + 일수 — O(1)
        storage_tier_number    = _gstn(cont_type, dg_raw)
        odcy_in_date  = row.get("odcy_in_date")
        odcy_out_date = row.get("odcy_out_date")
        raw_days, billable_days, free_days = _calc_storage_days(odcy_in_date, odcy_out_date, odcy_location, storage_tier_number)

        # TRKV 구간 조회 — O(1) (인라인, 중복 제거)
        pickup_port = _rp(pickup_name)
        departure_code = _rd(departure_name)
        trkv_dp = trkv_dest_port if trkv_dest_port else _rp(dest_name)
        trkv_details = _trkv_lookup(pickup_port, trkv_dp, departure_code,
                                    cont_type, dg_raw, quantity, weekend_holiday)

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
            "pickup_port_resolved": pickup_port,
            "odcy_code": odcy_code,
            "odcy_name": row.get("odcy_name"),
            "departure_name": departure_name,
            "departure_code_resolved": departure_code,
            "dest_code": dest_code,
            "dest_name_original": dest_name,
            "dest_name": trkv_dest_name,
            "dest_port_resolved": trkv_dest_port,
            "container_type": container_type,
            "dg_flag": row.get("dg_flag", False),
            "quantity": quantity,
            "weekend_holiday": weekend_holiday,
            "om_d": om_d,
            "odcy_destination_name": odcy_destination_name,
            "odcy_name_resolved": odcy_name_resolved,
            "odcy_terminal_type": odcy_terminal_type,
            "odcy_location": odcy_location,
            "dest_port_type": dest_port_type,
            "dest_terminal_type": dest_terminal_type,
            "storage_tier_number": storage_tier_number,
            "odcy_in_date": odcy_in_date,
            "odcy_out_date": odcy_out_date,
            "storage_days": raw_days,
            "billable_days": billable_days,
            "free_days": free_days,
            # TRKV 구간 정보
            "tier_number": trkv_details.get("tier_number"),
            "trkv_unit_rate": trkv_details.get("unit_rate"),
            "trkv_rate_row": trkv_details.get("route_row_num"),
        }

        result_id += 1

        # 직반입 판정
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
            elif charge_type == "TRKV":
                # 이미 계산된 trkv_details 사용 (중복 조회 제거)
                expected = trkv_details.get("expected")
                rate_row = None
                unit_rate = None
                if expected is None:
                    if actual == 0.0:
                        diff, status = None, "SKIP"
                    else:
                        diff, status = None, "NO_RATE"
                else:
                    diff = expected - actual
                    status = "OK" if abs(diff) < TOLERANCE else "DIFF"
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
                    trkv_dest_name=trkv_dest_name,
                    trkv_dest_port=trkv_dest_port,
                    odcy_destination_name=odcy_destination_name,
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

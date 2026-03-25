from sqlalchemy.orm import Session
from app.models import VerificationSession, VerificationResult
from app.services.rate_service import find_rate
from app.services import trkv_service

TOLERANCE = 1.0  # 원 단위 허용 오차

CHARGES = [
    ("TRKV",   "trkv_actual",    "trkv_expected",    "trkv_diff",    "trkv_status"),
    ("보관료",  "storage_actual", "storage_expected", "storage_diff", "storage_status"),
    ("상하차료", "handling_actual", "handling_expected", "handling_diff", "handling_status"),
    ("셔틀비용", "shuttle_actual", "shuttle_expected",  "shuttle_diff",  "shuttle_status"),
]


def _verify_charge(db, charge_type, actual, pickup_code, odcy_code, dest_code, container_type,
                   pickup_name=None, departure_name=None, dest_name=None,
                   cont_type=None, dg_raw=None):
    if charge_type == "TRKV":
        # TRKV 전용 2단계 요율 로직: 구간 + 컨테이너 티어
        expected = trkv_service.get_trkv_expected(
            db, pickup_name, departure_name, dest_name, cont_type, dg_raw
        )
    else:
        rate = find_rate(db, charge_type, pickup_code, odcy_code, dest_code, container_type)
        expected = rate.unit_price if rate else None

    if expected is None:
        if actual == 0.0:
            return None, None, "SKIP"
        return None, None, "NO_RATE"
    diff = actual - expected
    status = "OK" if abs(diff) < TOLERANCE else "DIFF"
    return expected, diff, status


def run_verification(db: Session, filename: str, rows: list[dict]) -> VerificationSession:
    session = VerificationSession(filename=filename, total_rows=len(rows))
    db.add(session)
    db.flush()

    counters = {
        "trkv_pass": 0, "trkv_fail": 0, "trkv_no_rate": 0,
        "storage_pass": 0, "storage_fail": 0, "storage_no_rate": 0,
        "handling_pass": 0, "handling_fail": 0, "handling_no_rate": 0,
        "shuttle_pass": 0, "shuttle_fail": 0, "shuttle_no_rate": 0,
    }
    total_diff = 0.0

    results = []
    for row in rows:
        pickup_code = row.get("pickup_code")
        odcy_code = row.get("odcy_code")
        dest_code = row.get("dest_code")
        container_type = row.get("container_type")

        # TRKV 전용 추가 필드
        pickup_name    = row.get("pickup_name")
        departure_name = row.get("departure_name")
        dest_name      = row.get("dest_name")
        cont_type      = row.get("cont_type")   # 엑셀 원본 (22G1 등)
        dg_raw         = row.get("dg_raw")      # 엑셀 원본 ("X" or "")

        result_kwargs = dict(
            session_id=session.id,
            row_number=row.get("row_number", 0),
            container_no=row.get("container_no"),
            transport_date=row.get("transport_date"),
            pickup_code=pickup_code,
            pickup_name=pickup_name,
            odcy_code=odcy_code,
            odcy_name=row.get("odcy_name"),
            dest_code=dest_code,
            dest_name=dest_name,
            container_type=container_type,
            dg_flag=row.get("dg_flag", False),
        )

        statuses = []
        for (charge_type, actual_key, exp_key, diff_key, status_key) in CHARGES:
            actual = row.get(actual_key, 0.0)
            expected, diff, status = _verify_charge(
                db, charge_type, actual, pickup_code, odcy_code, dest_code, container_type,
                pickup_name=pickup_name, departure_name=departure_name, dest_name=dest_name,
                cont_type=cont_type, dg_raw=dg_raw,
            )
            result_kwargs[actual_key] = actual
            result_kwargs[exp_key] = expected
            result_kwargs[diff_key] = diff
            result_kwargs[status_key] = status
            statuses.append(status)

            # 카운터 집계 (charge_type → prefix 매핑)
            prefix_map = {"TRKV": "trkv", "보관료": "storage", "상하차료": "handling", "셔틀비용": "shuttle"}
            prefix = prefix_map[charge_type]
            if status == "OK" or status == "SKIP":
                counters[f"{prefix}_pass"] += 1
            elif status == "DIFF":
                counters[f"{prefix}_fail"] += 1
                total_diff += abs(diff or 0)
            elif status == "NO_RATE":
                counters[f"{prefix}_no_rate"] += 1

        # 종합 상태
        if all(s in ("OK", "SKIP") for s in statuses):
            overall = "OK"
        elif "NO_RATE" in statuses:
            overall = "NO_RATE"
        else:
            overall = "DIFF"

        result_kwargs["overall_status"] = overall
        results.append(VerificationResult(**result_kwargs))

    db.bulk_save_objects(results)

    # 세션 카운터 업데이트
    for k, v in counters.items():
        setattr(session, k, v)
    session.total_diff = total_diff
    db.commit()
    db.refresh(session)
    return session

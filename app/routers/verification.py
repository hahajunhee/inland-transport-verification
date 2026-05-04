from datetime import datetime
from fastapi import APIRouter, HTTPException, UploadFile, File
from fastapi.responses import StreamingResponse
from typing import List, Optional
from io import BytesIO
from urllib.parse import quote

from app import data_store
from app.services.excel_service import parse_settlement_excel, generate_results_excel, generate_fwo_charge_excel
from app.services.verification_service import run_verification
from app.services.trkv_service import resolve_port, resolve_departure
from app.services.storage_rate_service import find_storage_rate

router = APIRouter()


@router.post("/upload")
async def upload_and_verify(file: UploadFile = File(...)):
    if not file.filename.endswith((".xlsx", ".xls")):
        raise HTTPException(status_code=400, detail="엑셀 파일(.xlsx, .xls)만 업로드 가능합니다.")

    content = await file.read()
    try:
        rows = parse_settlement_excel(content)
    except ValueError as e:
        raise HTTPException(status_code=422, detail=str(e))

    if not rows:
        raise HTTPException(status_code=422, detail="유효한 데이터 행이 없습니다.")

    session = run_verification(file.filename, rows)
    return session


@router.get("/sessions")
def list_sessions():
    sessions = data_store.load("verification_sessions.json")
    return sorted(sessions, key=lambda x: x["id"], reverse=True)


@router.get("/sessions/{session_id}")
def get_session(session_id: int):
    sessions = data_store.load("verification_sessions.json")
    s = next((x for x in sessions if x["id"] == session_id), None)
    if not s:
        raise HTTPException(status_code=404, detail="세션을 찾을 수 없습니다.")
    return s


@router.get("/sessions/{session_id}/results")
def get_results(
    session_id: int,
    status_filter: Optional[str] = None,
    skip: int = 0,
    limit: int = 500,
):
    results = data_store.load_results(session_id)
    results = sorted(results, key=lambda x: x.get("row_number", 0))

    if status_filter and status_filter != "ALL":
        if status_filter == "DIFF_OR_NO_RATE":
            results = [r for r in results if r.get("overall_status") in ("DIFF", "NO_RATE")]
        else:
            results = [r for r in results if r.get("overall_status") == status_filter]

    return results[skip: skip + limit]


@router.get("/sessions/{session_id}/export")
def export_results(session_id: int):
    sessions = data_store.load("verification_sessions.json")
    session = next((x for x in sessions if x["id"] == session_id), None)
    if not session:
        raise HTTPException(status_code=404, detail="세션을 찾을 수 없습니다.")
    results = data_store.load_results(session_id)
    results = sorted(results, key=lambda x: x.get("row_number", 0))
    excel_bytes = generate_results_excel(results)
    filename = f"검증결과_{session_id}.xlsx"
    return StreamingResponse(
        BytesIO(excel_bytes),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename*=UTF-8''{quote(filename)}"},
    )


@router.get("/sessions/{session_id}/export-fwo-charge")
def export_fwo_charge(session_id: int):
    """DIFF 행을 FWO Charge 템플릿으로 변환한 엑셀 다운로드."""
    sessions = data_store.load("verification_sessions.json")
    session = next((x for x in sessions if x["id"] == session_id), None)
    if not session:
        raise HTTPException(status_code=404, detail="세션을 찾을 수 없습니다.")
    results = data_store.load_results(session_id)
    results = sorted(results, key=lambda x: x.get("row_number", 0))
    excel_bytes = generate_fwo_charge_excel(results)
    now_str = datetime.now().strftime("%Y%m%d%H%M%S")
    filename = f"결과-매출인보이스생성(차지)_{now_str}.xlsx"
    return StreamingResponse(
        BytesIO(excel_bytes),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename*=UTF-8''{quote(filename)}"},
    )


@router.post("/sessions/{session_id}/generate-missing-rates")
def generate_missing_rates(session_id: int):
    """NO_RATE인 건들에 대해 누락 요율을 자동 생성 (티어1~6 모두 0)."""
    results = data_store.load_results(session_id)
    if not results:
        raise HTTPException(status_code=404, detail="세션 결과를 찾을 수 없습니다.")

    now_str = datetime.now().strftime("%Y-%m-%d %H:%M")
    auto_memo = f"[자동생성] 요율검증에서 자동생성된 요율입니다. ({now_str})"
    zero_tiers = {f"tier{t}": 0.0 for t in range(1, 7)}

    # ── TRKV 누락 구간 수집 ──
    trkv_routes = data_store.load("trkv_routes.json")
    trkv_missing = set()  # (pickup_port, departure_code, dest_port)
    for r in results:
        if r.get("trkv_status") == "NO_RATE":
            pp = r.get("pickup_port_resolved")
            dc = r.get("departure_code_resolved")
            dp = r.get("dest_port_resolved")
            if pp and dc and dp:
                # 이미 존재하는지 확인
                exists = any(
                    rt.get("pickup_port") == pp
                    and rt.get("departure_code", rt.get("departure_name", "")) == dc
                    and rt.get("dest_port") == dp
                    for rt in trkv_routes
                )
                if not exists:
                    trkv_missing.add((pp, dc, dp))

    trkv_created = 0
    for pp, dc, dp in trkv_missing:
        obj = {
            "id": data_store.next_id(trkv_routes),
            "pickup_port": pp,
            "departure_code": dc,
            "dest_port": dp,
            **zero_tiers,
            "memo": auto_memo,
            "auto_generated": True,
        }
        trkv_routes.append(obj)
        trkv_created += 1
    if trkv_created:
        data_store.save("trkv_routes.json", trkv_routes)

    # ── 보관료/상하차료/셔틀비 누락 요율 수집 ──
    storage_rates = data_store.load("storage_rates.json")
    storage_zero_tiers = {}
    for prefix in ("storage", "handling", "shuttle"):
        for t in range(1, 7):
            storage_zero_tiers[f"{prefix}_tier{t}"] = 0.0

    storage_missing = set()  # (odcy_name, odcy_terminal_type, odcy_location, dest_port_type, dest_terminal_type)
    for r in results:
        if r.get("storage_status") == "NO_RATE" or r.get("handling_status") == "NO_RATE" or r.get("shuttle_status") == "NO_RATE":
            key = (
                r.get("odcy_name_resolved") or "",
                r.get("odcy_terminal_type") or "",
                r.get("odcy_location") or "",
                r.get("dest_port_type") or "",
                r.get("dest_terminal_type") or "",
            )
            if any(key):  # 최소 하나는 값이 있어야 의미 있음
                # 이미 존재하는지 확인 (정확히 같은 5개 키 조합)
                exists = any(
                    (sr.get("odcy_name") or "") == key[0]
                    and (sr.get("odcy_terminal_type") or "") == key[1]
                    and (sr.get("odcy_location") or "") == key[2]
                    and (sr.get("dest_port_type") or "") == key[3]
                    and (sr.get("dest_terminal_type") or "") == key[4]
                    for sr in storage_rates
                )
                if not exists:
                    storage_missing.add(key)

    storage_created = 0
    for odcy_name, oterm, oloc, dpt, dtt in storage_missing:
        obj = {
            "id": data_store.next_id(storage_rates),
            "odcy_name": odcy_name,
            "odcy_terminal_type": oterm,
            "odcy_location": oloc,
            "dest_port_type": dpt,
            "dest_terminal_type": dtt,
            **storage_zero_tiers,
            "memo": auto_memo,
            "auto_generated": True,
        }
        storage_rates.append(obj)
        storage_created += 1
    if storage_created:
        data_store.save("storage_rates.json", storage_rates)

    return {
        "trkv_created": trkv_created,
        "storage_created": storage_created,
        "message": f"TRKV {trkv_created}건, 보관료/상하차료/셔틀비 {storage_created}건 요율 생성 완료",
    }


@router.delete("/sessions/{session_id}")
def delete_session(session_id: int):
    sessions = data_store.load("verification_sessions.json")
    new_sessions = [x for x in sessions if x["id"] != session_id]
    if len(new_sessions) == len(sessions):
        raise HTTPException(status_code=404, detail="세션을 찾을 수 없습니다.")
    data_store.save("verification_sessions.json", new_sessions)
    data_store.delete_results(session_id)
    return {"ok": True}

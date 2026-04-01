from fastapi import APIRouter, HTTPException, UploadFile, File
from fastapi.responses import StreamingResponse
from typing import List, Optional
from io import BytesIO
from urllib.parse import quote

from app import data_store
from app.services.excel_service import parse_settlement_excel, generate_results_excel
from app.services.verification_service import run_verification

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


@router.delete("/sessions/{session_id}")
def delete_session(session_id: int):
    sessions = data_store.load("verification_sessions.json")
    new_sessions = [x for x in sessions if x["id"] != session_id]
    if len(new_sessions) == len(sessions):
        raise HTTPException(status_code=404, detail="세션을 찾을 수 없습니다.")
    data_store.save("verification_sessions.json", new_sessions)
    data_store.delete_results(session_id)
    return {"ok": True}

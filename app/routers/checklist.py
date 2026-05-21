"""
1단계 체크리스트 검증 API 라우터
"""
from io import BytesIO
from urllib.parse import quote
from fastapi import APIRouter, UploadFile, File, HTTPException
from fastapi.responses import StreamingResponse
from pydantic import BaseModel

from app.services.checklist_service import (
    parse_checklist_excel,
    run_checklist,
    run_stage7,
    get_session,
    list_sessions,
    generate_checklist_report,
    load_user_holidays,
    save_user_holidays,
    STAGE_NAMES,
)

router = APIRouter()


@router.post("/upload")
async def upload_checklist(file: UploadFile = File(...)):
    if not file.filename.endswith((".xlsx", ".xls")):
        raise HTTPException(400, "엑셀 파일(.xlsx, .xls)만 지원합니다.")
    try:
        content = await file.read()
        rows = parse_checklist_excel(content)
    except ValueError as e:
        raise HTTPException(400, str(e))
    except Exception as e:
        raise HTTPException(400, f"파일 파싱 실패: {e}")

    session = run_checklist(file.filename, rows)

    error_counts = {}
    for stage_num in ("1", "2", "3", "4", "5", "6", "8"):
        error_counts[stage_num] = len(session["errors"].get(stage_num, []))

    return {
        "session_id": session["id"],
        "filename": session["filename"],
        "total_rows": session["total_rows"],
        "error_counts": error_counts,
        "errors": session["errors"],
        "stage_names": STAGE_NAMES,
    }


class Stage7Request(BaseModel):
    container_numbers: list[str]


@router.post("/stage7/{session_id}")
async def check_stage7(session_id: int, body: Stage7Request):
    errors = run_stage7(session_id, body.container_numbers)
    return {
        "errors": errors,
        "count": len(errors),
        "total_containers": len([c for c in body.container_numbers if c.strip()]),
    }


@router.get("/sessions")
def get_sessions():
    return list_sessions()


# ─── 휴일 관리 ────────────────────────────────────────────
@router.get("/holidays")
def get_holidays():
    return load_user_holidays()


class HolidaysRequest(BaseModel):
    dates: list[str]


@router.post("/holidays")
def set_holidays(body: HolidaysRequest):
    save_user_holidays(body.dates)
    return {"count": len(body.dates), "dates": sorted(set(body.dates))}


@router.get("/download/{session_id}")
def download_report(session_id: int):
    session = get_session(session_id)
    if not session:
        raise HTTPException(404, "세션을 찾을 수 없습니다.")

    try:
        report = generate_checklist_report(session)
    except Exception as e:
        raise HTTPException(500, f"리포트 생성 실패: {e}")

    raw_name = session.get("filename", "report")
    filename = f"체크리스트검증_{raw_name}.xlsx"

    return StreamingResponse(
        BytesIO(report),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename*=UTF-8''{quote(filename)}"},
    )

"""
3단계 모비스 검증 API 라우터
"""
from io import BytesIO
from typing import Optional
from urllib.parse import quote
from fastapi import APIRouter, UploadFile, File, HTTPException, Request
from fastapi.responses import StreamingResponse

from app import data_store
from app.services.mobis_service import (
    parse_mobis_excel,
    run_mobis_verification,
    build_ref_lookup,
    generate_mobis_report,
)

router = APIRouter()


@router.get("/sessions")
def get_verification_sessions():
    """2단계 정산검증 세션 목록 반환 (참조용)."""
    sessions = data_store.load("verification_sessions.json")
    return sorted(sessions, key=lambda s: s.get("id", 0), reverse=True)


@router.post("/upload")
async def upload_mobis(file: UploadFile = File(...), session_id: Optional[int] = None):
    if not file.filename.endswith((".xlsx", ".xls")):
        raise HTTPException(400, "엑셀 파일(.xlsx, .xls)만 지원합니다.")
    try:
        content = await file.read()
        parsed = parse_mobis_excel(content)
    except ValueError as e:
        raise HTTPException(400, str(e))
    except Exception as e:
        raise HTTPException(400, f"파일 파싱 실패: {e}")

    # 2단계 검증결과 참조 빌드
    ref_lookup = None
    if session_id is not None:
        try:
            ref_lookup = build_ref_lookup(session_id)
        except Exception:
            pass  # 참조 실패 시 무시 (존재X 표시)

    result = run_mobis_verification(file.filename, parsed, ref_lookup)
    return result


@router.post("/download")
async def download_mobis_report_endpoint(request: Request):
    """검증 결과 JSON을 받아 엑셀 다운로드."""
    try:
        verification = await request.json()
    except Exception:
        raise HTTPException(400, "잘못된 요청입니다.")

    try:
        report = generate_mobis_report(verification)
    except Exception as e:
        raise HTTPException(500, f"리포트 생성 실패: {e}")

    raw_name = verification.get("filename", "결과")
    if "." in raw_name:
        raw_name = raw_name.rsplit(".", 1)[0]
    filename = f"모비스검증_{raw_name}.xlsx"

    return StreamingResponse(
        BytesIO(report),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename*=UTF-8''{quote(filename)}"},
    )

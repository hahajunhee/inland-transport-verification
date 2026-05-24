"""
3단계 모비스 검증 API 라우터
"""
from io import BytesIO
from urllib.parse import quote
from fastapi import APIRouter, UploadFile, File, HTTPException, Request
from fastapi.responses import StreamingResponse

from app.services.mobis_service import (
    parse_mobis_excel,
    run_mobis_verification,
    generate_mobis_report,
)

router = APIRouter()


@router.post("/upload")
async def upload_mobis(file: UploadFile = File(...)):
    if not file.filename.endswith((".xlsx", ".xls")):
        raise HTTPException(400, "엑셀 파일(.xlsx, .xls)만 지원합니다.")
    try:
        content = await file.read()
        parsed = parse_mobis_excel(content)
    except ValueError as e:
        raise HTTPException(400, str(e))
    except Exception as e:
        raise HTTPException(400, f"파일 파싱 실패: {e}")

    result = run_mobis_verification(file.filename, parsed)
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

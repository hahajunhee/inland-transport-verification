"""
3단계 모비스 검증 API 라우터
"""
from io import BytesIO
from urllib.parse import quote
from fastapi import APIRouter, UploadFile, File, HTTPException
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
async def download_mobis_report(file: UploadFile = File(...)):
    """검증 후 결과 엑셀 다운로드."""
    if not file.filename.endswith((".xlsx", ".xls")):
        raise HTTPException(400, "엑셀 파일(.xlsx, .xls)만 지원합니다.")
    try:
        content = await file.read()
        parsed = parse_mobis_excel(content)
    except ValueError as e:
        raise HTTPException(400, str(e))
    except Exception as e:
        raise HTTPException(400, f"파일 파싱 실패: {e}")

    verification = run_mobis_verification(file.filename, parsed)

    try:
        report = generate_mobis_report(verification)
    except Exception as e:
        raise HTTPException(500, f"리포트 생성 실패: {e}")

    raw_name = file.filename.rsplit(".", 1)[0] if "." in file.filename else file.filename
    filename = f"모비스검증_{raw_name}.xlsx"

    return StreamingResponse(
        BytesIO(report),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename*=UTF-8''{quote(filename)}"},
    )

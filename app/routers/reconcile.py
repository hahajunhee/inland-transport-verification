"""
4단계 정산대사 검증 API 라우터
- 여러 정산 엑셀(1·2단계 양식) 병합본 ↔ MOBIS 엑셀(3단계 양식) 금액 대사
"""
from io import BytesIO
from typing import List
from urllib.parse import quote

from fastapi import APIRouter, UploadFile, File, Form, HTTPException, Request
from fastapi.responses import StreamingResponse

from app.services.reconcile_service import (
    run_reconciliation, generate_reconcile_report, GUBUN_OPTIONS,
)

router = APIRouter()


def _check_ext(filename: str):
    if not filename.lower().endswith((".xlsx", ".xls")):
        raise HTTPException(400, f"엑셀 파일(.xlsx, .xls)만 지원합니다: {filename}")


@router.post("/upload")
async def upload_reconcile(
    settlement_files: List[UploadFile] = File(...),
    mobis_file: UploadFile = File(...),
    gubun: str = Form("MOBIS 산출"),
):
    if not settlement_files:
        raise HTTPException(400, "정산(합칠) 파일을 1개 이상 업로드하세요.")
    if gubun not in GUBUN_OPTIONS:
        gubun = "MOBIS 산출"

    settlement_bytes = []
    for f in settlement_files:
        _check_ext(f.filename)
        settlement_bytes.append(await f.read())

    _check_ext(mobis_file.filename)
    mobis_bytes = await mobis_file.read()

    try:
        result = run_reconciliation(
            settlement_bytes, mobis_bytes,
            gubun_filter=gubun, mobis_filename=mobis_file.filename,
        )
    except ValueError as e:
        raise HTTPException(400, str(e))
    except Exception as e:
        raise HTTPException(400, f"대사 처리 실패: {e}")

    return result


@router.post("/download")
async def download_reconcile(request: Request):
    """대사 결과 JSON → 엑셀 다운로드."""
    try:
        verification = await request.json()
    except Exception:
        raise HTTPException(400, "잘못된 요청입니다.")

    try:
        report = generate_reconcile_report(verification)
    except Exception as e:
        raise HTTPException(500, f"리포트 생성 실패: {e}")

    raw = verification.get("filename", "결과")
    if "." in raw:
        raw = raw.rsplit(".", 1)[0]
    filename = f"정산대사_{raw}.xlsx"

    return StreamingResponse(
        BytesIO(report),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename*=UTF-8''{quote(filename)}"},
    )

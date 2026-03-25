from fastapi import APIRouter, Depends, HTTPException, UploadFile, File
from fastapi.responses import StreamingResponse
from sqlalchemy.orm import Session
from typing import List
from io import BytesIO

from app.database import get_db
from app.schemas import SessionSummary, ResultRow
from app.models import VerificationSession, VerificationResult
from app.services.excel_service import parse_settlement_excel, generate_results_excel
from app.services.verification_service import run_verification

router = APIRouter()


@router.post("/upload", response_model=SessionSummary)
async def upload_and_verify(file: UploadFile = File(...), db: Session = Depends(get_db)):
    if not file.filename.endswith((".xlsx", ".xls")):
        raise HTTPException(status_code=400, detail="엑셀 파일(.xlsx, .xls)만 업로드 가능합니다.")

    content = await file.read()
    try:
        rows = parse_settlement_excel(content)
    except ValueError as e:
        raise HTTPException(status_code=422, detail=str(e))

    if not rows:
        raise HTTPException(status_code=422, detail="유효한 데이터 행이 없습니다.")

    session = run_verification(db, file.filename, rows)
    return session


@router.get("/sessions", response_model=List[SessionSummary])
def list_sessions(db: Session = Depends(get_db)):
    return db.query(VerificationSession).order_by(VerificationSession.id.desc()).all()


@router.get("/sessions/{session_id}", response_model=SessionSummary)
def get_session(session_id: int, db: Session = Depends(get_db)):
    s = db.query(VerificationSession).filter(VerificationSession.id == session_id).first()
    if not s:
        raise HTTPException(status_code=404, detail="세션을 찾을 수 없습니다.")
    return s


@router.get("/sessions/{session_id}/results", response_model=List[ResultRow])
def get_results(
    session_id: int,
    status_filter: str = None,
    skip: int = 0,
    limit: int = 500,
    db: Session = Depends(get_db),
):
    q = db.query(VerificationResult).filter(VerificationResult.session_id == session_id)
    if status_filter and status_filter != "ALL":
        if status_filter == "DIFF_OR_NO_RATE":
            q = q.filter(VerificationResult.overall_status.in_(["DIFF", "NO_RATE"]))
        else:
            q = q.filter(VerificationResult.overall_status == status_filter)
    return q.order_by(VerificationResult.row_number).offset(skip).limit(limit).all()


@router.get("/sessions/{session_id}/export")
def export_results(session_id: int, db: Session = Depends(get_db)):
    session = db.query(VerificationSession).filter(VerificationSession.id == session_id).first()
    if not session:
        raise HTTPException(status_code=404, detail="세션을 찾을 수 없습니다.")
    results = db.query(VerificationResult).filter(VerificationResult.session_id == session_id).order_by(VerificationResult.row_number).all()
    excel_bytes = generate_results_excel(results)
    filename = f"검증결과_{session_id}.xlsx"
    return StreamingResponse(
        BytesIO(excel_bytes),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


@router.delete("/sessions/{session_id}")
def delete_session(session_id: int, db: Session = Depends(get_db)):
    session = db.query(VerificationSession).filter(VerificationSession.id == session_id).first()
    if not session:
        raise HTTPException(status_code=404, detail="세션을 찾을 수 없습니다.")
    db.query(VerificationResult).filter(VerificationResult.session_id == session_id).delete()
    db.delete(session)
    db.commit()
    return {"ok": True}

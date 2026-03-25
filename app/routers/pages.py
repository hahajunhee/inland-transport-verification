from fastapi import APIRouter, Request, Depends
from fastapi.templating import Jinja2Templates
from sqlalchemy.orm import Session
import os

from app.database import get_db
from app.models import TransportRate, VerificationSession, TRKVRoute, TRKVContainerTier

router = APIRouter()
templates = Jinja2Templates(directory=os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), "templates"))


@router.get("/")
def index(request: Request, db: Session = Depends(get_db)):
    rate_count = db.query(TransportRate).count()
    session_count = db.query(VerificationSession).count()
    recent = db.query(VerificationSession).order_by(VerificationSession.id.desc()).limit(5).all()
    trkv_route_count = db.query(TRKVRoute).count()
    # 컨테이너 티어 8개 중 설정된 수
    trkv_tier_set = db.query(TRKVContainerTier).filter(TRKVContainerTier.tier_number.isnot(None)).count()
    return templates.TemplateResponse("index.html", {
        "request": request,
        "rate_count": rate_count,
        "session_count": session_count,
        "recent_sessions": recent,
        "trkv_route_count": trkv_route_count,
        "trkv_tier_set": trkv_tier_set,
    })


@router.get("/rates")
def rates_page(request: Request):
    return templates.TemplateResponse("rates.html", {"request": request})


@router.get("/verification")
def verification_page(request: Request):
    return templates.TemplateResponse("verification.html", {"request": request})


@router.get("/trkv")
def trkv_page(request: Request):
    return templates.TemplateResponse("trkv.html", {"request": request})

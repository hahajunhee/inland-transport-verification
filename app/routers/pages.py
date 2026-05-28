from fastapi import APIRouter, Request
from fastapi.templating import Jinja2Templates
import os

from app import data_store, __version__ as APP_VERSION

router = APIRouter()
templates = Jinja2Templates(directory=os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), "templates"))
# 모든 페이지에서 {{ version }} 으로 앱 버전 사용 가능 (base.html 의 <title>/favicon 옆에 표시)
templates.env.globals["version"] = APP_VERSION


@router.get("/")
def index(request: Request):
    rates = data_store.load("transport_rates.json")
    sessions = data_store.load("verification_sessions.json")
    routes = data_store.load("trkv_routes.json")
    tiers = data_store.load("container_tiers.json")

    rate_count = len(rates)
    session_count = len(sessions)
    recent = sorted(sessions, key=lambda x: x["id"], reverse=True)[:5]
    trkv_route_count = len(routes)
    trkv_tier_set = sum(1 for t in tiers if t.get("tier_number") is not None)

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


@router.get("/checklist")
def checklist_page(request: Request):
    return templates.TemplateResponse("checklist.html", {"request": request})


@router.get("/mobis")
def mobis_page(request: Request):
    return templates.TemplateResponse("mobis.html", {"request": request})


@router.get("/rate-register")
def rate_register_page(request: Request):
    return templates.TemplateResponse("rate_register.html", {"request": request})


# 하위호환: 기존 URL 유지 (요율등록 페이지로 리다이렉트)
@router.get("/trkv")
def trkv_page(request: Request):
    from fastapi.responses import RedirectResponse
    return RedirectResponse(url="/rate-register")


@router.get("/mapping")
def mapping_page(request: Request):
    from fastapi.responses import RedirectResponse
    return RedirectResponse(url="/rate-register")


@router.get("/storage-rates")
def storage_rates_page(request: Request):
    from fastapi.responses import RedirectResponse
    return RedirectResponse(url="/rate-register")

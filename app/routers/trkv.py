from io import BytesIO
from fastapi import APIRouter, Depends, HTTPException, UploadFile, File, Form
from fastapi.responses import StreamingResponse
from sqlalchemy.orm import Session
from pydantic import BaseModel
from typing import Optional
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.datavalidation import DataValidation

from app.database import get_db
from app.models import TRKVPortMapping, TRKVRoute
from app.services import trkv_service

router = APIRouter()


# ─── Pydantic 스키마 ─────────────────────────────────────────────────

class PortMappingCreate(BaseModel):
    excel_name: str
    port_type: str  # "부산신항" / "부산북항"


class RouteCreate(BaseModel):
    pickup_port: str
    departure_name: str
    dest_port: str
    tier1: Optional[float] = None
    tier2: Optional[float] = None
    tier3: Optional[float] = None
    tier4: Optional[float] = None
    tier5: Optional[float] = None
    tier6: Optional[float] = None
    memo: Optional[str] = None


class ContainerTierItem(BaseModel):
    cont_type: str   # "22G1" / "22R1" / "45G1" / "45R1"
    is_dg: bool
    tier_number: Optional[int] = None  # 1~6


class ContainerTierBulk(BaseModel):
    items: list[ContainerTierItem]


# ─── 포트명 매핑 ─────────────────────────────────────────────────────

@router.get("/port-mappings")
def list_port_mappings(db: Session = Depends(get_db)):
    items = trkv_service.get_all_port_mappings(db)
    return [{"id": i.id, "excel_name": i.excel_name, "port_type": i.port_type} for i in items]


@router.post("/port-mappings", status_code=201)
def add_port_mapping(body: PortMappingCreate, db: Session = Depends(get_db)):
    try:
        obj = trkv_service.create_port_mapping(db, body.excel_name, body.port_type)
    except Exception:
        raise HTTPException(status_code=400, detail="이미 등록된 포트명이거나 오류가 발생했습니다.")
    return {"id": obj.id, "excel_name": obj.excel_name, "port_type": obj.port_type}


@router.put("/port-mappings/{mapping_id}")
def edit_port_mapping(mapping_id: int, body: PortMappingCreate, db: Session = Depends(get_db)):
    obj = trkv_service.update_port_mapping(db, mapping_id, body.excel_name, body.port_type)
    if not obj:
        raise HTTPException(status_code=404, detail="포트 매핑을 찾을 수 없습니다.")
    return {"id": obj.id, "excel_name": obj.excel_name, "port_type": obj.port_type}


@router.delete("/port-mappings/{mapping_id}", status_code=204)
def remove_port_mapping(mapping_id: int, db: Session = Depends(get_db)):
    if not trkv_service.delete_port_mapping(db, mapping_id):
        raise HTTPException(status_code=404, detail="포트 매핑을 찾을 수 없습니다.")


# ─── 포트명 매핑 - 엑셀 템플릿 & 업로드 ──────────────────────────────

@router.get("/port-mappings/template")
def download_port_mappings_template():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "포트명 매핑"

    headers = ["엑셀 원본명", "포트 구분"]
    ws.append(headers)

    # 헤더 스타일
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="4472C4")
        cell.alignment = Alignment(horizontal="center")
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 15

    # 포트 구분 드롭다운 유효성 검사
    dv = DataValidation(type="list", formula1='"부산신항,부산북항"', allow_blank=False)
    ws.add_data_validation(dv)
    dv.sqref = "B2:B1000"

    # 예시 행
    ws.append(["부산신항BPTS", "부산신항"])
    ws.append(["부산북항BPNC", "부산북항"])

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename*=UTF-8''port_mappings_template.xlsx"},
    )


@router.post("/port-mappings/upload")
async def upload_port_mappings(
    file: UploadFile = File(...),
    mode: str = Form("append"),
    db: Session = Depends(get_db),
):
    content = await file.read()
    try:
        wb = openpyxl.load_workbook(BytesIO(content), data_only=True)
    except Exception:
        raise HTTPException(400, detail="올바른 xlsx 파일이 아닙니다.")

    ws = wb.active
    header = [cell.value for cell in ws[1]]
    col_map = {str(v).strip(): i for i, v in enumerate(header) if v is not None}

    if "엑셀 원본명" not in col_map or "포트 구분" not in col_map:
        raise HTTPException(400, detail="'엑셀 원본명', '포트 구분' 컬럼이 필요합니다.")

    if mode == "replace":
        db.query(TRKVPortMapping).delete()
        db.commit()

    success, failed = 0, []
    for i, row in enumerate(ws.iter_rows(min_row=2), start=2):
        try:
            excel_name = row[col_map["엑셀 원본명"]].value
            port_type = row[col_map["포트 구분"]].value
            if not excel_name and not port_type:
                continue
            if not excel_name or not port_type:
                failed.append({"row": i, "error": "엑셀 원본명, 포트 구분 모두 필수입니다."})
                continue
            excel_name = str(excel_name).strip()
            port_type = str(port_type).strip()
            if port_type not in ("부산신항", "부산북항"):
                failed.append({"row": i, "error": f"포트 구분 '{port_type}'은 부산신항 또는 부산북항이어야 합니다."})
                continue
            # Upsert
            existing = db.query(TRKVPortMapping).filter(TRKVPortMapping.excel_name == excel_name).first()
            if existing:
                existing.port_type = port_type
            else:
                db.add(TRKVPortMapping(excel_name=excel_name, port_type=port_type))
            success += 1
        except Exception as e:
            failed.append({"row": i, "error": str(e)})

    db.commit()
    return {"success": success, "failed": failed, "mode": mode}


# ─── 구간요율 ─────────────────────────────────────────────────────────

@router.get("/routes")
def list_routes(db: Session = Depends(get_db)):
    items = trkv_service.get_all_routes(db)
    return [
        {
            "id": r.id,
            "pickup_port": r.pickup_port,
            "departure_name": r.departure_name,
            "dest_port": r.dest_port,
            "tier1": r.tier1, "tier2": r.tier2, "tier3": r.tier3,
            "tier4": r.tier4, "tier5": r.tier5, "tier6": r.tier6,
            "memo": r.memo,
        }
        for r in items
    ]


@router.post("/routes", status_code=201)
def add_route(body: RouteCreate, db: Session = Depends(get_db)):
    obj = trkv_service.create_route(db, body.model_dump())
    return {"id": obj.id, **body.model_dump()}


@router.put("/routes/{route_id}")
def edit_route(route_id: int, body: RouteCreate, db: Session = Depends(get_db)):
    obj = trkv_service.update_route(db, route_id, body.model_dump())
    if not obj:
        raise HTTPException(status_code=404, detail="구간 요율을 찾을 수 없습니다.")
    return {"id": obj.id, "pickup_port": obj.pickup_port, "departure_name": obj.departure_name,
            "dest_port": obj.dest_port, "tier1": obj.tier1, "tier2": obj.tier2,
            "tier3": obj.tier3, "tier4": obj.tier4, "tier5": obj.tier5, "tier6": obj.tier6,
            "memo": obj.memo}


@router.delete("/routes/{route_id}", status_code=204)
def remove_route(route_id: int, db: Session = Depends(get_db)):
    if not trkv_service.delete_route(db, route_id):
        raise HTTPException(status_code=404, detail="구간 요율을 찾을 수 없습니다.")


# ─── 구간요율 - 엑셀 템플릿 & 업로드 ────────────────────────────────

@router.get("/routes/template")
def download_routes_template():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "구간별 요율"

    headers = ["픽업항", "출하지명", "도착항", "티어1", "티어2", "티어3", "티어4", "티어5", "티어6", "비고"]
    ws.append(headers)

    # 헤더 스타일
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="4472C4")
        cell.alignment = Alignment(horizontal="center")

    # 컬럼 너비
    widths = [15, 15, 15, 12, 12, 12, 12, 12, 12, 25]
    for col, w in enumerate(widths, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = w

    # 픽업항(A), 도착항(C) 드롭다운 유효성 검사
    dv_a = DataValidation(type="list", formula1='"부산신항,부산북항"', allow_blank=False)
    ws.add_data_validation(dv_a)
    dv_a.sqref = "A2:A1000"

    dv_c = DataValidation(type="list", formula1='"부산신항,부산북항"', allow_blank=False)
    ws.add_data_validation(dv_c)
    dv_c.sqref = "C2:C1000"

    # 예시 행
    ws.append(["부산신항", "아산", "부산북항", 100000, 110000, 120000, 130000, 140000, 150000, "예시 (등록 후 삭제)"])

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename*=UTF-8''trkv_routes_template.xlsx"},
    )


@router.post("/routes/upload")
async def upload_routes(
    file: UploadFile = File(...),
    mode: str = Form("append"),
    db: Session = Depends(get_db),
):
    content = await file.read()
    try:
        wb = openpyxl.load_workbook(BytesIO(content), data_only=True)
    except Exception:
        raise HTTPException(400, detail="올바른 xlsx 파일이 아닙니다.")

    ws = wb.active
    header = [cell.value for cell in ws[1]]
    col_map = {str(v).strip(): i for i, v in enumerate(header) if v is not None}

    required = ["픽업항", "출하지명", "도착항"]
    for c in required:
        if c not in col_map:
            raise HTTPException(400, detail=f"'{c}' 컬럼이 없습니다. 양식을 다운로드해 사용하세요.")

    def gv(row, name):
        idx = col_map.get(name)
        if idx is None:
            return None
        return row[idx].value

    def to_float(v):
        if v is None or str(v).strip() == "":
            return None
        try:
            return float(v)
        except Exception:
            return None

    if mode == "replace":
        db.query(TRKVRoute).delete()
        db.commit()

    success, failed = 0, []
    for i, row in enumerate(ws.iter_rows(min_row=2), start=2):
        try:
            pickup = gv(row, "픽업항")
            departure = gv(row, "출하지명")
            dest = gv(row, "도착항")
            # 완전히 빈 행 건너뜀
            if not pickup and not departure and not dest:
                continue
            if not pickup or not departure or not dest:
                failed.append({"row": i, "error": "픽업항, 출하지명, 도착항은 필수입니다."})
                continue

            data = {
                "pickup_port": str(pickup).strip(),
                "departure_name": str(departure).strip(),
                "dest_port": str(dest).strip(),
                "tier1": to_float(gv(row, "티어1")),
                "tier2": to_float(gv(row, "티어2")),
                "tier3": to_float(gv(row, "티어3")),
                "tier4": to_float(gv(row, "티어4")),
                "tier5": to_float(gv(row, "티어5")),
                "tier6": to_float(gv(row, "티어6")),
                "memo": str(gv(row, "비고") or "").strip() or None,
            }
            trkv_service.create_route(db, data)
            success += 1
        except Exception as e:
            failed.append({"row": i, "error": str(e)})

    return {"success": success, "failed": failed, "mode": mode}


# ─── 컨테이너 티어 ───────────────────────────────────────────────────

@router.get("/container-tiers")
def list_container_tiers(db: Session = Depends(get_db)):
    items = trkv_service.get_all_container_tiers(db)
    return [
        {"id": i.id, "cont_type": i.cont_type, "is_dg": i.is_dg, "tier_number": i.tier_number}
        for i in items
    ]


@router.post("/container-tiers/bulk")
def save_container_tiers(body: ContainerTierBulk, db: Session = Depends(get_db)):
    items = [i.model_dump() for i in body.items]
    result = trkv_service.bulk_save_container_tiers(db, items)
    return [
        {"id": r.id, "cont_type": r.cont_type, "is_dg": r.is_dg, "tier_number": r.tier_number}
        for r in result
    ]


@router.put("/container-tiers/{tier_id}")
def edit_container_tier(tier_id: int, tier_number: Optional[int] = None, db: Session = Depends(get_db)):
    obj = trkv_service.update_container_tier(db, tier_id, tier_number)
    if not obj:
        raise HTTPException(status_code=404, detail="컨테이너 티어를 찾을 수 없습니다.")
    return {"id": obj.id, "cont_type": obj.cont_type, "is_dg": obj.is_dg, "tier_number": obj.tier_number}

from io import BytesIO
from fastapi import APIRouter, HTTPException, UploadFile, File
from fastapi.responses import StreamingResponse
from pydantic import BaseModel
from typing import Optional
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.datavalidation import DataValidation

from app.services import trkv_service
from app import data_store

router = APIRouter()


# ─── Pydantic 스키마 ─────────────────────────────────────────────────

class PortMappingCreate(BaseModel):
    excel_name: str
    port_type: str


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
    cont_type: str
    is_dg: bool
    tier_number: Optional[int] = None


class ContainerTierBulk(BaseModel):
    items: list[ContainerTierItem]


# ─── 포트명 매핑 CRUD ─────────────────────────────────────────────────

@router.get("/port-mappings")
def list_port_mappings():
    return trkv_service.get_all_port_mappings()


@router.post("/port-mappings", status_code=201)
def add_port_mapping(body: PortMappingCreate):
    try:
        return trkv_service.create_port_mapping(body.excel_name, body.port_type)
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))


@router.put("/port-mappings/{mapping_id}")
def edit_port_mapping(mapping_id: int, body: PortMappingCreate):
    obj = trkv_service.update_port_mapping(mapping_id, body.excel_name, body.port_type)
    if not obj:
        raise HTTPException(status_code=404, detail="포트 매핑을 찾을 수 없습니다.")
    return obj


@router.delete("/port-mappings/{mapping_id}", status_code=204)
def remove_port_mapping(mapping_id: int):
    if not trkv_service.delete_port_mapping(mapping_id):
        raise HTTPException(status_code=404, detail="포트 매핑을 찾을 수 없습니다.")


# ─── 구간요율 CRUD ─────────────────────────────────────────────────────

@router.get("/routes")
def list_routes():
    return trkv_service.get_all_routes()


@router.post("/routes", status_code=201)
def add_route(body: RouteCreate):
    return trkv_service.create_route(body.model_dump())


@router.put("/routes/{route_id}")
def edit_route(route_id: int, body: RouteCreate):
    obj = trkv_service.update_route(route_id, body.model_dump())
    if not obj:
        raise HTTPException(status_code=404, detail="구간 요율을 찾을 수 없습니다.")
    return obj


@router.delete("/routes/{route_id}", status_code=204)
def remove_route(route_id: int):
    if not trkv_service.delete_route(route_id):
        raise HTTPException(status_code=404, detail="구간 요율을 찾을 수 없습니다.")


# ─── 컨테이너 티어 ───────────────────────────────────────────────────

@router.get("/container-tiers")
def list_container_tiers():
    return trkv_service.get_all_container_tiers()


@router.post("/container-tiers/bulk")
def save_container_tiers(body: ContainerTierBulk):
    return trkv_service.bulk_save_container_tiers([i.model_dump() for i in body.items])


@router.put("/container-tiers/{tier_id}")
def edit_container_tier(tier_id: int, tier_number: Optional[int] = None):
    obj = trkv_service.update_container_tier(tier_id, tier_number)
    if not obj:
        raise HTTPException(status_code=404, detail="컨테이너 티어를 찾을 수 없습니다.")
    return obj


# ─── 통합 엑셀 템플릿 (현재 등록된 데이터 포함) ───────────────────────

def _style_header(ws, headers: list, col_widths: list):
    ws.append(headers)
    hdr_fill = PatternFill("solid", fgColor="4472C4")
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = hdr_fill
        cell.alignment = Alignment(horizontal="center")
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w


@router.get("/template")
def download_unified_template():
    """현재 등록된 데이터를 포함한 통합 양식 다운로드 (전체 교체용)"""
    port_mappings = trkv_service.get_all_port_mappings()
    routes = trkv_service.get_all_routes()

    wb = openpyxl.Workbook()

    # ── Sheet 1: 포트명 매핑 ──────────────────────────────────────────
    ws_pm = wb.active
    ws_pm.title = "포트명 매핑"
    _style_header(ws_pm, ["엑셀 원본명", "포트 구분"], [30, 15])

    dv_pm = DataValidation(type="list", formula1='"부산신항,부산북항"', allow_blank=False)
    ws_pm.add_data_validation(dv_pm)
    dv_pm.sqref = "B2:B1000"

    for pm in port_mappings:
        ws_pm.append([pm["excel_name"], pm["port_type"]])

    if not port_mappings:
        ws_pm.append(["부산신항BPTS", "부산신항"])
        ws_pm.append(["부산북항BPNC", "부산북항"])

    # ── Sheet 2: TRKV 구간 요율 ──────────────────────────────────────
    ws_rt = wb.create_sheet("TRKV 구간 요율")
    _style_header(
        ws_rt,
        ["픽업항", "출하지명", "도착항", "티어1", "티어2", "티어3", "티어4", "티어5", "티어6", "비고"],
        [15, 15, 15, 12, 12, 12, 12, 12, 12, 25],
    )

    dv_a = DataValidation(type="list", formula1='"부산신항,부산북항"', allow_blank=False)
    ws_rt.add_data_validation(dv_a)
    dv_a.sqref = "A2:A1000"

    dv_c = DataValidation(type="list", formula1='"부산신항,부산북항"', allow_blank=False)
    ws_rt.add_data_validation(dv_c)
    dv_c.sqref = "C2:C1000"

    for r in routes:
        ws_rt.append([
            r.get("pickup_port"), r.get("departure_name"), r.get("dest_port"),
            r.get("tier1"), r.get("tier2"), r.get("tier3"),
            r.get("tier4"), r.get("tier5"), r.get("tier6"),
            r.get("memo"),
        ])

    if not routes:
        ws_rt.append(["부산신항", "아산", "부산북항", 100000, 110000, 120000, 130000, 140000, 150000, "예시 (등록 후 삭제)"])

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename*=UTF-8''trkv_template.xlsx"},
    )


# ─── 통합 업로드 (항상 전체 교체) ────────────────────────────────────

@router.post("/upload")
async def upload_unified(file: UploadFile = File(...)):
    """통합 업로드: 내용이 있는 시트를 전체 교체로 처리"""
    content = await file.read()
    try:
        wb = openpyxl.load_workbook(BytesIO(content), data_only=True)
    except Exception:
        raise HTTPException(400, detail="올바른 xlsx 파일이 아닙니다.")

    def to_float(v):
        if v is None or str(v).strip() == "":
            return None
        try:
            return float(v)
        except Exception:
            return None

    result = {}

    # ── 포트명 매핑 시트 ─────────────────────────────────────────────
    if "포트명 매핑" in wb.sheetnames:
        ws = wb["포트명 매핑"]
        header = [cell.value for cell in ws[1]]
        col_map = {str(v).strip(): i for i, v in enumerate(header) if v is not None}

        has_data = any(
            any(c.value is not None for c in row)
            for row in ws.iter_rows(min_row=2, max_row=ws.max_row)
        )

        if has_data and "엑셀 원본명" in col_map and "포트 구분" in col_map:
            # 전체 교체
            data_store.save("port_mappings.json", [])

            success, failed = 0, []
            new_items = []
            next_id = 1
            for i, row in enumerate(ws.iter_rows(min_row=2), start=2):
                excel_name = row[col_map["엑셀 원본명"]].value
                port_type  = row[col_map["포트 구분"]].value
                if not excel_name and not port_type:
                    continue
                if not excel_name or not port_type:
                    failed.append({"row": i, "error": "엑셀 원본명, 포트 구분 모두 필수입니다."})
                    continue
                excel_name = str(excel_name).strip()
                port_type  = str(port_type).strip()
                if port_type not in ("부산신항", "부산북항"):
                    failed.append({"row": i, "error": f"포트 구분 '{port_type}'은 부산신항 또는 부산북항이어야 합니다."})
                    continue
                new_items.append({"id": next_id, "excel_name": excel_name, "port_type": port_type})
                next_id += 1
                success += 1
            data_store.save("port_mappings.json", new_items)
            result["포트명 매핑"] = {"success": success, "failed": failed}

    # ── TRKV 구간 요율 시트 ──────────────────────────────────────────
    if "TRKV 구간 요율" in wb.sheetnames:
        ws = wb["TRKV 구간 요율"]
        header = [cell.value for cell in ws[1]]
        col_map = {str(v).strip(): i for i, v in enumerate(header) if v is not None}

        has_data = any(
            any(c.value is not None for c in row)
            for row in ws.iter_rows(min_row=2, max_row=ws.max_row)
        )

        if has_data and all(c in col_map for c in ["픽업항", "출하지명", "도착항"]):
            # 전체 교체
            data_store.save("trkv_routes.json", [])

            def gv(row, name):
                idx = col_map.get(name)
                return row[idx].value if idx is not None else None

            success, failed = 0, []
            new_routes = []
            next_id = 1
            for i, row in enumerate(ws.iter_rows(min_row=2), start=2):
                pickup   = gv(row, "픽업항")
                departure = gv(row, "출하지명")
                dest     = gv(row, "도착항")
                if not pickup and not departure and not dest:
                    continue
                if not pickup or not departure or not dest:
                    failed.append({"row": i, "error": "픽업항, 출하지명, 도착항은 필수입니다."})
                    continue
                data = {
                    "id": next_id,
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
                new_routes.append(data)
                next_id += 1
                success += 1
            data_store.save("trkv_routes.json", new_routes)
            result["TRKV 구간 요율"] = {"success": success, "failed": failed}

    if not result:
        raise HTTPException(400, detail="처리할 수 있는 시트가 없습니다. 통합 양식을 사용하세요.")

    return {"sheets": result}

"""
요율 데이터 백업(JSON 다운로드) / 복원(JSON 업로드) API
"""
import json
from io import BytesIO
from datetime import datetime

from fastapi import APIRouter, HTTPException, UploadFile, File
from fastapi.responses import StreamingResponse

from app import data_store

router = APIRouter()


@router.get("/backup")
def download_backup():
    """전체 요율 데이터를 JSON으로 다운로드"""
    data = {
        "backup_at": datetime.now().isoformat(),
        "version": 2,
        "transport_rates": data_store.load("transport_rates.json"),
        "trkv_port_mappings": data_store.load("port_mappings.json"),
        "trkv_routes": data_store.load("trkv_routes.json"),
        "trkv_container_tiers": data_store.load("container_tiers.json"),
    }
    content = json.dumps(data, ensure_ascii=False, indent=2).encode("utf-8")
    filename = f"transport_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
    return StreamingResponse(
        BytesIO(content),
        media_type="application/json",
        headers={"Content-Disposition": f"attachment; filename*=UTF-8''{filename}"},
    )


@router.post("/restore")
async def restore_backup(file: UploadFile = File(...)):
    """JSON 백업 파일을 업로드하여 요율 데이터 복원 (기존 데이터 전체 교체)"""
    content = await file.read()
    try:
        data = json.loads(content.decode("utf-8"))
    except Exception:
        raise HTTPException(400, detail="올바른 JSON 파일이 아닙니다.")

    if data.get("version") not in (1, 2):
        raise HTTPException(400, detail="지원하지 않는 백업 형식입니다.")

    counts = {}

    rates = data.get("transport_rates", [])
    data_store.save("transport_rates.json", rates)
    counts["transport_rates"] = len(rates)

    pms = data.get("trkv_port_mappings", [])
    data_store.save("port_mappings.json", pms)
    counts["trkv_port_mappings"] = len(pms)

    routes = data.get("trkv_routes", [])
    data_store.save("trkv_routes.json", routes)
    counts["trkv_routes"] = len(routes)

    tiers = data.get("trkv_container_tiers", [])
    data_store.save("container_tiers.json", tiers)
    counts["trkv_container_tiers"] = len(tiers)

    return {"status": "ok", "restored": counts}

"""
요율 데이터 백업(JSON 다운로드) / 복원(JSON 업로드) API
"""
import json
from io import BytesIO
from datetime import datetime

from fastapi import APIRouter, Depends, HTTPException, UploadFile, File
from fastapi.responses import StreamingResponse
from sqlalchemy.orm import Session

from app.database import get_db
from app.models import TransportRate, TRKVPortMapping, TRKVRoute, TRKVContainerTier

router = APIRouter()


def _serialize(obj) -> dict:
    """SQLAlchemy 모델 인스턴스를 dict로 변환 (id, created_at 제외)"""
    skip = {"id", "_sa_instance_state"}
    result = {}
    for col in obj.__table__.columns:
        if col.name in skip:
            continue
        val = getattr(obj, col.name)
        if hasattr(val, "isoformat"):
            val = val.isoformat()
        result[col.name] = val
    return result


@router.get("/backup")
def download_backup(db: Session = Depends(get_db)):
    """전체 요율 데이터를 JSON으로 다운로드"""
    data = {
        "backup_at": datetime.now().isoformat(),
        "version": 1,
        "transport_rates": [_serialize(r) for r in db.query(TransportRate).all()],
        "trkv_port_mappings": [_serialize(r) for r in db.query(TRKVPortMapping).all()],
        "trkv_routes": [_serialize(r) for r in db.query(TRKVRoute).all()],
        "trkv_container_tiers": [_serialize(r) for r in db.query(TRKVContainerTier).all()],
    }
    content = json.dumps(data, ensure_ascii=False, indent=2).encode("utf-8")
    filename = f"transport_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
    return StreamingResponse(
        BytesIO(content),
        media_type="application/json",
        headers={"Content-Disposition": f"attachment; filename*=UTF-8''{filename}"},
    )


@router.post("/restore")
async def restore_backup(
    file: UploadFile = File(...),
    db: Session = Depends(get_db),
):
    """JSON 백업 파일을 업로드하여 요율 데이터 복원 (기존 데이터 전체 교체)"""
    content = await file.read()
    try:
        data = json.loads(content.decode("utf-8"))
    except Exception:
        raise HTTPException(400, detail="올바른 JSON 파일이 아닙니다.")

    if data.get("version") != 1:
        raise HTTPException(400, detail="지원하지 않는 백업 형식입니다.")

    # 기존 데이터 삭제
    db.query(TRKVContainerTier).delete()
    db.query(TRKVRoute).delete()
    db.query(TRKVPortMapping).delete()
    db.query(TransportRate).delete()
    db.commit()

    counts = {}

    # TransportRate 복원
    rates = data.get("transport_rates", [])
    for row in rates:
        db.add(TransportRate(**{k: v for k, v in row.items() if hasattr(TransportRate, k)}))
    counts["transport_rates"] = len(rates)

    # TRKVPortMapping 복원
    pms = data.get("trkv_port_mappings", [])
    for row in pms:
        db.add(TRKVPortMapping(**{k: v for k, v in row.items() if hasattr(TRKVPortMapping, k)}))
    counts["trkv_port_mappings"] = len(pms)

    # TRKVRoute 복원
    routes = data.get("trkv_routes", [])
    for row in routes:
        db.add(TRKVRoute(**{k: v for k, v in row.items() if hasattr(TRKVRoute, k)}))
    counts["trkv_routes"] = len(routes)

    # TRKVContainerTier 복원
    tiers = data.get("trkv_container_tiers", [])
    for row in tiers:
        db.add(TRKVContainerTier(**{k: v for k, v in row.items() if hasattr(TRKVContainerTier, k)}))
    counts["trkv_container_tiers"] = len(tiers)

    db.commit()
    return {"status": "ok", "restored": counts}

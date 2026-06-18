import uvicorn
from contextlib import asynccontextmanager
from fastapi import FastAPI
from fastapi.staticfiles import StaticFiles
import os

from app import data_store
from app.routers import rates, verification, pages, trkv, backup, storage_rates, checklist, mobis, reconcile, mouselock


@asynccontextmanager
async def lifespan(app: FastAPI):
    # data/ 디렉토리 및 SQLite DB 초기화
    data_store.init_db()
    yield


app = FastAPI(title="내륙운송정산검증 시스템", lifespan=lifespan)

BASE_DIR = os.path.dirname(__file__)

app.mount("/static", StaticFiles(directory=os.path.join(BASE_DIR, "static")), name="static")

app.include_router(pages.router)
app.include_router(rates.router, prefix="/api/rates", tags=["rates"])
app.include_router(verification.router, prefix="/api/verification", tags=["verification"])
app.include_router(trkv.router, prefix="/api/trkv", tags=["trkv"])
app.include_router(backup.router, prefix="/api", tags=["backup"])
app.include_router(storage_rates.router, prefix="/api/storage-rates", tags=["storage-rates"])
app.include_router(checklist.router, prefix="/api/checklist", tags=["checklist"])
app.include_router(mobis.router, prefix="/api/mobis", tags=["mobis"])
app.include_router(reconcile.router, prefix="/api/reconcile", tags=["reconcile"])
app.include_router(mouselock.router, prefix="/api/using", tags=["using"])


if __name__ == "__main__":
    # reload=False: 핫리로드를 끄면 단일 프로세스로 실행됨
    # (reload 시 런처+워커 두 프로세스가 각각 입력 훅을 설치해 토글이 어긋나는 문제 방지)
    uvicorn.run(app, host="127.0.0.1", port=8000, reload=False)

import uvicorn
from contextlib import asynccontextmanager
from fastapi import FastAPI
from fastapi.staticfiles import StaticFiles
import os

from app.database import create_tables
from app.routers import rates, verification, pages, trkv, backup


@asynccontextmanager
async def lifespan(app: FastAPI):
    create_tables()
    yield


app = FastAPI(title="내륙운송정산검증 시스템", lifespan=lifespan)

BASE_DIR = os.path.dirname(__file__)

app.mount("/static", StaticFiles(directory=os.path.join(BASE_DIR, "static")), name="static")

app.include_router(pages.router)
app.include_router(rates.router, prefix="/api/rates", tags=["rates"])
app.include_router(verification.router, prefix="/api/verification", tags=["verification"])
app.include_router(trkv.router, prefix="/api/trkv", tags=["trkv"])
app.include_router(backup.router, prefix="/api", tags=["backup"])


if __name__ == "__main__":
    uvicorn.run("main:app", host="127.0.0.1", port=8000, reload=True)

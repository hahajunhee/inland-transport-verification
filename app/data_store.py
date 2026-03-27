"""
JSON 파일 기반 데이터 저장소 유틸리티
SQLite .db 파일 대신 JSON 파일로 데이터를 저장/관리합니다.
"""
import json
import threading
from pathlib import Path

DATA_DIR = Path(__file__).parent.parent / "data"
RESULTS_DIR = DATA_DIR / "results"

_lock = threading.Lock()


def _ensure_dirs():
    DATA_DIR.mkdir(exist_ok=True)
    RESULTS_DIR.mkdir(exist_ok=True)


def load(filename: str) -> list:
    """JSON 파일 로드. 파일이 없으면 빈 리스트 반환."""
    path = DATA_DIR / filename
    if not path.exists():
        return []
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def save(filename: str, data: list):
    """JSON 파일로 저장 (스레드 안전)."""
    _ensure_dirs()
    with _lock:
        with open(DATA_DIR / filename, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2, default=str)


def next_id(items: list) -> int:
    """현재 리스트에서 다음 ID 계산 (max + 1)."""
    if not items:
        return 1
    return max(item["id"] for item in items) + 1


def load_results(session_id: int) -> list:
    """세션별 검증 결과 로드."""
    path = RESULTS_DIR / f"session_{session_id}.json"
    if not path.exists():
        return []
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def save_results(session_id: int, results: list):
    """세션별 검증 결과 저장."""
    _ensure_dirs()
    with _lock:
        with open(RESULTS_DIR / f"session_{session_id}.json", "w", encoding="utf-8") as f:
            json.dump(results, f, ensure_ascii=False, indent=2, default=str)


def delete_results(session_id: int):
    """세션별 검증 결과 파일 삭제."""
    path = RESULTS_DIR / f"session_{session_id}.json"
    if path.exists():
        path.unlink()

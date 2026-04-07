"""
SQLite 기반 데이터 저장소 유틸리티
JSON 파일 대신 단일 .db 파일로 모든 데이터를 관리합니다.
"""
import json
import sqlite3
import threading
from pathlib import Path

DATA_DIR = Path(__file__).parent.parent / "data"
DB_PATH = DATA_DIR / "inland_transport.db"

_local = threading.local()
_write_lock = threading.Lock()


def _ensure_dirs():
    DATA_DIR.mkdir(exist_ok=True)


def _get_conn() -> sqlite3.Connection:
    """스레드별 SQLite 연결 반환 (WAL 모드)."""
    if not hasattr(_local, "conn") or _local.conn is None:
        _ensure_dirs()
        conn = sqlite3.connect(str(DB_PATH), check_same_thread=False)
        conn.execute("PRAGMA journal_mode=WAL")
        conn.execute("PRAGMA foreign_keys=ON")
        conn.row_factory = sqlite3.Row
        _local.conn = conn
    return _local.conn


# ─── 파일명 → 테이블명 매핑 ───────────────────────────────────────────

_FILE_TO_TABLE = {
    "transport_rates.json": "transport_rates",
    "storage_rates.json": "storage_rates",
    "trkv_routes.json": "trkv_routes",
    "container_tiers.json": "container_tiers",
    "storage_container_tiers.json": "storage_container_tiers",
    "port_mappings.json": "port_mappings",
    "odcy_mappings.json": "odcy_mappings",
    "departure_mappings.json": "departure_mappings",
    "verification_sessions.json": "verification_sessions",
}

# ─── 테이블 스키마 정의 ─────────────────────────────────────────────────

_SCHEMAS = {
    "transport_rates": [
        "id INTEGER PRIMARY KEY",
        "charge_type TEXT",
        "pickup_code TEXT",
        "odcy_code TEXT",
        "dest_code TEXT",
        "container_type TEXT",
        "unit_price REAL",
    ],
    "trkv_routes": [
        "id INTEGER PRIMARY KEY",
        "pickup_port TEXT",
        "departure_code TEXT",
        "departure_name TEXT",
        "dest_port TEXT",
        "tier1 REAL", "tier2 REAL", "tier3 REAL",
        "tier4 REAL", "tier5 REAL", "tier6 REAL",
        "memo TEXT DEFAULT ''",
        "auto_generated INTEGER DEFAULT 0",
        "sheet_name TEXT DEFAULT ''",
    ],
    "container_tiers": [
        "id INTEGER PRIMARY KEY",
        "cont_type TEXT NOT NULL",
        "is_dg INTEGER DEFAULT 0",
        "tier_number INTEGER",
    ],
    "storage_container_tiers": [
        "id INTEGER PRIMARY KEY",
        "cont_type TEXT NOT NULL",
        "is_dg INTEGER DEFAULT 0",
        "tier_number INTEGER",
    ],
    "port_mappings": [
        "id INTEGER PRIMARY KEY",
        "excel_name TEXT NOT NULL",
        "port_type TEXT",
        "terminal_type TEXT DEFAULT ''",
    ],
    "departure_mappings": [
        "id INTEGER PRIMARY KEY",
        "departure_name TEXT NOT NULL",
        "departure_code TEXT",
    ],
    "odcy_mappings": [
        "id INTEGER PRIMARY KEY",
        "odcy_destination_name TEXT NOT NULL",
        "odcy_name TEXT",
        "odcy_terminal_type TEXT DEFAULT ''",
        "odcy_location TEXT DEFAULT ''",
    ],
    "storage_rates": [
        "id INTEGER PRIMARY KEY",
        "odcy_name TEXT DEFAULT ''",
        "odcy_terminal_type TEXT DEFAULT ''",
        "odcy_location TEXT DEFAULT ''",
        "dest_port_type TEXT DEFAULT ''",
        "dest_terminal_type TEXT DEFAULT ''",
        "storage_tier1 REAL", "storage_tier2 REAL", "storage_tier3 REAL",
        "storage_tier4 REAL", "storage_tier5 REAL", "storage_tier6 REAL",
        "handling_tier1 REAL", "handling_tier2 REAL", "handling_tier3 REAL",
        "handling_tier4 REAL", "handling_tier5 REAL", "handling_tier6 REAL",
        "shuttle_tier1 REAL", "shuttle_tier2 REAL", "shuttle_tier3 REAL",
        "shuttle_tier4 REAL", "shuttle_tier5 REAL", "shuttle_tier6 REAL",
        "storage_unit REAL",
        "handling_unit REAL",
        "shuttle_unit REAL",
        "memo TEXT DEFAULT ''",
        "auto_generated INTEGER DEFAULT 0",
        "sheet_name TEXT DEFAULT ''",
    ],
    "verification_sessions": [
        "id INTEGER PRIMARY KEY",
        "filename TEXT",
        "uploaded_at TEXT",
        "total_rows INTEGER DEFAULT 0",
        "trkv_pass INTEGER DEFAULT 0",
        "trkv_fail INTEGER DEFAULT 0",
        "trkv_no_rate INTEGER DEFAULT 0",
        "storage_pass INTEGER DEFAULT 0",
        "storage_fail INTEGER DEFAULT 0",
        "storage_no_rate INTEGER DEFAULT 0",
        "handling_pass INTEGER DEFAULT 0",
        "handling_fail INTEGER DEFAULT 0",
        "handling_no_rate INTEGER DEFAULT 0",
        "shuttle_pass INTEGER DEFAULT 0",
        "shuttle_fail INTEGER DEFAULT 0",
        "shuttle_no_rate INTEGER DEFAULT 0",
        "total_diff REAL DEFAULT 0",
    ],
    "verification_results": [
        "id INTEGER PRIMARY KEY AUTOINCREMENT",
        "session_id INTEGER NOT NULL",
        "data_json TEXT NOT NULL",
    ],
}

# 각 테이블의 컬럼명 캐시 (타입 제외)
_COLUMN_NAMES = {}
for _tbl, _cols in _SCHEMAS.items():
    _COLUMN_NAMES[_tbl] = [col.split()[0] for col in _cols]


def init_db():
    """모든 테이블 생성 (존재하지 않으면)."""
    conn = _get_conn()
    for table_name, columns in _SCHEMAS.items():
        sql = f"CREATE TABLE IF NOT EXISTS {table_name} ({', '.join(columns)})"
        conn.execute(sql)
    # verification_results 인덱스
    conn.execute(
        "CREATE INDEX IF NOT EXISTS idx_results_session "
        "ON verification_results(session_id)"
    )
    conn.commit()


# ─── 범용 load / save ──────────────────────────────────────────────────

def _table_for(filename: str) -> str:
    tbl = _FILE_TO_TABLE.get(filename)
    if not tbl:
        raise ValueError(f"알 수 없는 데이터 파일: {filename}")
    return tbl


def _row_to_dict(row: sqlite3.Row) -> dict:
    """sqlite3.Row → dict 변환. None 값도 포함."""
    return dict(row)


def _prepare_row(table: str, item: dict) -> dict:
    """dict에서 해당 테이블 컬럼에 맞는 값만 추출. bool → int 변환."""
    cols = _COLUMN_NAMES[table]
    result = {}
    for col in cols:
        if col in item:
            val = item[col]
            # SQLite는 bool을 지원하지 않으므로 int로 변환
            if isinstance(val, bool):
                val = int(val)
            result[col] = val
    return result


def load(filename: str) -> list:
    """테이블 전체 로드. 파일명으로 테이블 매핑."""
    init_db()
    table = _table_for(filename)
    conn = _get_conn()
    rows = conn.execute(f"SELECT * FROM {table}").fetchall()
    result = []
    for row in rows:
        d = _row_to_dict(row)
        # is_dg 필드: int → bool 변환 (기존 호환성)
        if "is_dg" in d and d["is_dg"] is not None:
            d["is_dg"] = bool(d["is_dg"])
        # auto_generated 필드: int → bool 변환
        if "auto_generated" in d and d["auto_generated"] is not None:
            d["auto_generated"] = bool(d["auto_generated"])
        result.append(d)
    return result


def save(filename: str, data: list):
    """테이블 전체 교체 (스레드 안전)."""
    init_db()
    table = _table_for(filename)
    with _write_lock:
        conn = _get_conn()
        conn.execute(f"DELETE FROM {table}")
        if data:
            for item in data:
                prepared = _prepare_row(table, item)
                if not prepared:
                    continue
                cols = list(prepared.keys())
                placeholders = ", ".join(["?"] * len(cols))
                col_names = ", ".join(cols)
                conn.execute(
                    f"INSERT INTO {table} ({col_names}) VALUES ({placeholders})",
                    [prepared[c] for c in cols],
                )
        conn.commit()


def next_id(items: list) -> int:
    """현재 리스트에서 다음 ID 계산 (max + 1)."""
    if not items:
        return 1
    return max(item.get("id", 0) for item in items) + 1


# ─── 검증 결과 (세션별) ────────────────────────────────────────────────

def load_results(session_id: int) -> list:
    """세션별 검증 결과 로드."""
    init_db()
    conn = _get_conn()
    rows = conn.execute(
        "SELECT data_json FROM verification_results WHERE session_id = ? ORDER BY id",
        (session_id,),
    ).fetchall()
    return [json.loads(row["data_json"]) for row in rows]


def save_results(session_id: int, results: list):
    """세션별 검증 결과 저장."""
    init_db()
    with _write_lock:
        conn = _get_conn()
        conn.execute("DELETE FROM verification_results WHERE session_id = ?", (session_id,))
        for item in results:
            conn.execute(
                "INSERT INTO verification_results (session_id, data_json) VALUES (?, ?)",
                (session_id, json.dumps(item, ensure_ascii=False, default=str)),
            )
        conn.commit()


def delete_results(session_id: int):
    """세션별 검증 결과 삭제."""
    init_db()
    with _write_lock:
        conn = _get_conn()
        conn.execute("DELETE FROM verification_results WHERE session_id = ?", (session_id,))
        conn.commit()


# ─── 앱 시작 시 초기화 ─────────────────────────────────────────────────

init_db()

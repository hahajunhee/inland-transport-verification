"""
기존 SQLite DB(data/transport.db) → JSON 파일 마이그레이션 스크립트
최초 1회 실행 후 삭제해도 됩니다.

실행 방법:
    python migrate_db_to_json.py
"""
import sqlite3
import json
import pathlib

DATA_DIR = pathlib.Path(__file__).parent / "data"
DB_PATH  = DATA_DIR / "transport.db"

if not DB_PATH.exists():
    print("transport.db 파일이 없습니다. 마이그레이션이 필요하지 않습니다.")
    exit(0)

conn = sqlite3.connect(str(DB_PATH))
conn.row_factory = sqlite3.Row


def table_to_list(table_name: str) -> list:
    try:
        rows = conn.execute(f"SELECT * FROM {table_name}").fetchall()
        return [dict(row) for row in rows]
    except Exception as e:
        print(f"  [{table_name}] 테이블 없음 또는 오류: {e}")
        return []


def save_json(filename: str, data: list):
    path = DATA_DIR / filename
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2, default=str)
    print(f"  → {filename} ({len(data)}건 저장)")


print("=== SQLite → JSON 마이그레이션 시작 ===")

# transport_rates
rates = table_to_list("transport_rates")
save_json("transport_rates.json", rates)

# trkv_port_mappings
pms = table_to_list("trkv_port_mappings")
save_json("port_mappings.json", pms)

# trkv_routes
routes = table_to_list("trkv_routes")
save_json("trkv_routes.json", routes)

# trkv_container_tiers
tiers = table_to_list("trkv_container_tiers")
save_json("container_tiers.json", tiers)

conn.close()

print("\n=== 마이그레이션 완료 ===")
print("서버를 재시작하면 JSON 파일의 데이터가 적용됩니다.")
print("마이그레이션 후 이 파일(migrate_db_to_json.py)은 삭제해도 됩니다.")

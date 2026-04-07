"""
JSON 파일 → SQLite DB 마이그레이션 스크립트
기존 data/*.json 파일들을 inland_transport.db로 변환합니다.
실행: python migrate_json_to_db.py
"""
import json
import sys
from pathlib import Path

# 프로젝트 루트를 path에 추가
sys.path.insert(0, str(Path(__file__).parent))

from app import data_store

DATA_DIR = Path(__file__).parent / "data"
RESULTS_DIR = DATA_DIR / "results"

# JSON 파일명 목록 (data_store._FILE_TO_TABLE 과 동일)
JSON_FILES = [
    "transport_rates.json",
    "storage_rates.json",
    "trkv_routes.json",
    "container_tiers.json",
    "storage_container_tiers.json",
    "port_mappings.json",
    "odcy_mappings.json",
    "departure_mappings.json",
    "verification_sessions.json",
]


def load_json(filename: str) -> list:
    path = DATA_DIR / filename
    if not path.exists():
        return []
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def migrate():
    print("=" * 50)
    print("JSON → SQLite 마이그레이션 시작")
    print(f"DB 경로: {data_store.DB_PATH}")
    print("=" * 50)

    # 1. 테이블 초기화
    data_store.init_db()
    print("✓ 테이블 생성 완료")

    # 2. 일반 테이블 마이그레이션
    for filename in JSON_FILES:
        items = load_json(filename)
        if items:
            data_store.save(filename, items)
            print(f"✓ {filename} → {len(items)}건 마이그레이션")
        else:
            print(f"- {filename} → 데이터 없음 (스킵)")

    # 3. 검증 결과 (session_*.json) 마이그레이션
    if RESULTS_DIR.exists():
        result_files = sorted(RESULTS_DIR.glob("session_*.json"))
        for rf in result_files:
            session_id = int(rf.stem.split("_")[1])
            with open(rf, "r", encoding="utf-8") as f:
                results = json.load(f)
            if results:
                data_store.save_results(session_id, results)
                print(f"✓ {rf.name} → {len(results)}건 검증결과 마이그레이션")
            else:
                print(f"- {rf.name} → 데이터 없음 (스킵)")
    else:
        print("- results/ 디렉토리 없음 (스킵)")

    # 4. 검증: 각 테이블 행 수 확인
    print("\n" + "=" * 50)
    print("마이그레이션 검증")
    print("=" * 50)
    for filename in JSON_FILES:
        items = data_store.load(filename)
        table = data_store._FILE_TO_TABLE[filename]
        print(f"  {table}: {len(items)}건")

    conn = data_store._get_conn()
    count = conn.execute("SELECT COUNT(*) FROM verification_results").fetchone()[0]
    print(f"  verification_results: {count}건")

    print("\n✓ 마이그레이션 완료!")
    print(f"  DB 파일: {data_store.DB_PATH}")
    print(f"  DB 크기: {data_store.DB_PATH.stat().st_size / 1024:.1f} KB")


if __name__ == "__main__":
    migrate()

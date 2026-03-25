from sqlalchemy import Column, Integer, Float, Text, DateTime, ForeignKey, Boolean
from sqlalchemy.sql import func
from app.database import Base


class TransportRate(Base):
    __tablename__ = "transport_rates"

    id = Column(Integer, primary_key=True, autoincrement=True)
    charge_type = Column(Text, nullable=False)   # TRKV / 보관료 / 상하차료 / 셔틀비용
    pickup_code = Column(Text, nullable=True)     # 픽업지 코드 (None = any)
    odcy_code = Column(Text, nullable=True)       # ODCY 코드 (None = any)
    dest_code = Column(Text, nullable=True)       # 도착지 코드 (None = any)
    container_type = Column(Text, nullable=True)  # 컨테이너유형 (None = any)
    unit_price = Column(Float, nullable=False)    # 건당 단가 (원)
    memo = Column(Text, nullable=True)
    created_at = Column(DateTime, server_default=func.now())


class VerificationSession(Base):
    __tablename__ = "verification_sessions"

    id = Column(Integer, primary_key=True, autoincrement=True)
    filename = Column(Text, nullable=False)
    uploaded_at = Column(DateTime, server_default=func.now())
    total_rows = Column(Integer, default=0)

    trkv_pass = Column(Integer, default=0)
    trkv_fail = Column(Integer, default=0)
    trkv_no_rate = Column(Integer, default=0)

    storage_pass = Column(Integer, default=0)
    storage_fail = Column(Integer, default=0)
    storage_no_rate = Column(Integer, default=0)

    handling_pass = Column(Integer, default=0)
    handling_fail = Column(Integer, default=0)
    handling_no_rate = Column(Integer, default=0)

    shuttle_pass = Column(Integer, default=0)
    shuttle_fail = Column(Integer, default=0)
    shuttle_no_rate = Column(Integer, default=0)

    total_diff = Column(Float, default=0.0)


class VerificationResult(Base):
    __tablename__ = "verification_results"

    id = Column(Integer, primary_key=True, autoincrement=True)
    session_id = Column(Integer, ForeignKey("verification_sessions.id"), nullable=False)
    row_number = Column(Integer, nullable=False)

    container_no = Column(Text, nullable=True)
    transport_date = Column(Text, nullable=True)
    pickup_code = Column(Text, nullable=True)
    pickup_name = Column(Text, nullable=True)
    odcy_code = Column(Text, nullable=True)
    odcy_name = Column(Text, nullable=True)
    dest_code = Column(Text, nullable=True)
    dest_name = Column(Text, nullable=True)
    container_type = Column(Text, nullable=True)
    dg_flag = Column(Boolean, default=False)

    trkv_actual = Column(Float, nullable=True)
    trkv_expected = Column(Float, nullable=True)
    trkv_diff = Column(Float, nullable=True)
    trkv_status = Column(Text, nullable=True)   # OK / DIFF / NO_RATE / SKIP

    storage_actual = Column(Float, nullable=True)
    storage_expected = Column(Float, nullable=True)
    storage_diff = Column(Float, nullable=True)
    storage_status = Column(Text, nullable=True)

    handling_actual = Column(Float, nullable=True)
    handling_expected = Column(Float, nullable=True)
    handling_diff = Column(Float, nullable=True)
    handling_status = Column(Text, nullable=True)

    shuttle_actual = Column(Float, nullable=True)
    shuttle_expected = Column(Float, nullable=True)
    shuttle_diff = Column(Float, nullable=True)
    shuttle_status = Column(Text, nullable=True)

    overall_status = Column(Text, nullable=True)  # OK / DIFF / NO_RATE
    memo = Column(Text, nullable=True)


# ─── TRKV 전용 요율 테이블 ───────────────────────────────────────────

class TRKVPortMapping(Base):
    """엑셀에서 오는 픽업지명/도착지명 → 부산신항/부산북항 매핑"""
    __tablename__ = "trkv_port_mappings"

    id         = Column(Integer, primary_key=True, autoincrement=True)
    excel_name = Column(Text, nullable=False, unique=True)  # 엑셀 원본 명칭
    port_type  = Column(Text, nullable=False)               # "부산신항" 또는 "부산북항"


class TRKVRoute(Base):
    """운송구간(픽업항-출하지-도착항)별 티어1~6 단가"""
    __tablename__ = "trkv_routes"

    id             = Column(Integer, primary_key=True, autoincrement=True)
    pickup_port    = Column(Text, nullable=False)  # "부산신항" / "부산북항"
    departure_name = Column(Text, nullable=False)  # 출하지명 (엑셀 "출하지명" 컬럼 exact match)
    dest_port      = Column(Text, nullable=False)  # "부산신항" / "부산북항"
    tier1  = Column(Float, nullable=True)
    tier2  = Column(Float, nullable=True)
    tier3  = Column(Float, nullable=True)
    tier4  = Column(Float, nullable=True)
    tier5  = Column(Float, nullable=True)
    tier6  = Column(Float, nullable=True)
    memo   = Column(Text, nullable=True)
    created_at = Column(DateTime, server_default=func.now())


class TRKVContainerTier(Base):
    """Cont.Type + D/G여부 조합 → 티어번호(1~6) 매핑 (8개 고정 조합)"""
    __tablename__ = "trkv_container_tiers"

    id          = Column(Integer, primary_key=True, autoincrement=True)
    cont_type   = Column(Text, nullable=False)     # "22G1" / "22R1" / "45G1" / "45R1"
    is_dg       = Column(Boolean, nullable=False)  # True = D/G여부가 "X"
    tier_number = Column(Integer, nullable=True)   # 1 ~ 6 (미설정이면 NULL)

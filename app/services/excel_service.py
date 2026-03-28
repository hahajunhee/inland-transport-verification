import pandas as pd
import re
from io import BytesIO
from typing import Optional
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, numbers
from openpyxl.utils import get_column_letter

# ─── 컬럼 헤더 매핑 ───────────────────────────────────────────
COLUMN_MAP = {
    "D/G여부":                "dg_flag",
    "Cont.Category":          "cont_category",
    "Cont Size":              "cont_size",
    "Cont.Type":              "cont_type",
    "픽업지":                 "pickup_code",
    "픽업지명":               "pickup_name",
    "상세 ODCY":              "odcy_code",
    "상세 ODCY명":            "odcy_name",
    "출하지명":               "departure_name",
    "도착지":                 "dest_code",
    "도착지명":               "dest_name",
    "출하일":                 "transport_date",
    "Contrainer No.":         "container_no",
    "Quantity":               "quantity",
    "Weekend / Holiday":      "weekend_holiday",
    "Mobis 운임합계(매출)":   "trkv_actual",
    "ODCY 보관료":            "storage_actual",
    "ODCY 상하차료":          "handling_actual",
    "ODCY 셔틀료":            "shuttle_actual",
}


def _safe_float(value) -> float:
    """쉼표 포함 숫자 문자열도 파싱"""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).replace(",", "").strip()
    try:
        return float(s)
    except ValueError:
        return 0.0


def _derive_container_type(row: dict) -> str:
    """
    Cont.Category (DR/RF) + D/G여부 + Cont.Type(앞2자리) → 컨테이너 유형 문자열
    예) DR + 20 + DG=N → "20드라이"
        RF + 40 + DG=Y → "40리퍼위험물"
    """
    category = str(row.get("cont_category") or "").strip().upper()
    dg_raw = str(row.get("dg_flag") or "").strip().upper()
    is_dg = dg_raw in ("Y", "TRUE", "1", "예", "X")

    # 사이즈: Cont.Type 앞 두 자리 숫자 우선
    cont_type_str = str(row.get("cont_type") or "")
    size_match = re.match(r"(\d{2})", cont_type_str)
    if size_match:
        size_num = int(size_match.group(1))
        # ISO: 22=20ft, 42=40ft, 45=45ft, 25=25ft rare
        size_map = {22: "20", 42: "40", 45: "45", 25: "25"}
        size = size_map.get(size_num, str(size_num))
    else:
        # fallback: Cont Size 필드
        cont_size = str(row.get("cont_size") or "").strip().upper()
        if "20" in cont_size:
            size = "20"
        elif "45" in cont_size:
            size = "45"
        else:
            size = "40"

    type_str = "드라이" if category == "DR" else "리퍼" if category == "RF" else category
    dg_str = "위험물" if is_dg else ""
    return f"{size}{type_str}{dg_str}"


def parse_settlement_excel(file_bytes: bytes) -> list[dict]:
    """업로드된 엑셀 파일 파싱. 헤더행 자동 탐지."""
    df = pd.read_excel(BytesIO(file_bytes), header=None, dtype=str)

    # 헤더 행 탐지: "Cont.Category" 또는 "픽업지" 포함된 행
    header_row_idx = None
    for i, row in df.iterrows():
        row_values = [str(v) for v in row.values]
        if any(col in row_values for col in ["Cont.Category", "픽업지", "D/G여부"]):
            header_row_idx = i
            break

    if header_row_idx is None:
        raise ValueError("엑셀에서 헤더 행을 찾을 수 없습니다. 컬럼명(픽업지, Cont.Category 등)이 포함된 헤더 행이 필요합니다.")

    df.columns = df.iloc[header_row_idx]
    df = df.iloc[header_row_idx + 1:].reset_index(drop=True)

    # 필요 컬럼만 추출
    available = {col: mapped for col, mapped in COLUMN_MAP.items() if col in df.columns}
    missing_required = [col for col in ["픽업지", "도착지"] if col not in available]
    if missing_required:
        raise ValueError(f"필수 컬럼 없음: {missing_required}")

    rows = []
    for idx, raw_row in df.iterrows():
        row: dict = {}
        for excel_col, internal_key in available.items():
            row[internal_key] = raw_row.get(excel_col)

        # 빈 행 스킵 (컨테이너번호 또는 픽업지가 비어있으면)
        pickup = str(row.get("pickup_code") or "").strip()
        if not pickup or pickup in ("nan", "None"):
            continue

        # 금액 파싱
        for field in ("trkv_actual", "storage_actual", "handling_actual", "shuttle_actual"):
            row[field] = _safe_float(row.get(field))

        # 수량 파싱 (없으면 1.0 기본값)
        qty = _safe_float(row.get("quantity"))
        row["quantity"] = qty if qty > 0 else 1.0

        # 컨테이너 유형 도출
        row["container_type"] = _derive_container_type(row)

        # DG 플래그: 원본값 보존 후 bool 변환
        dg_raw = str(row.get("dg_flag") or "").strip().upper()
        row["dg_raw"] = dg_raw  # TRKV 서비스에서 "X" 판단용 원본값
        row["dg_flag"] = dg_raw in ("Y", "TRUE", "1", "예", "X")

        # 날짜 정규화
        td = str(row.get("transport_date") or "").strip()
        row["transport_date"] = td if td not in ("nan", "None", "") else None

        # 코드/이름 strip
        for field in ("pickup_code", "pickup_name", "odcy_code", "odcy_name", "dest_code", "dest_name", "container_no", "departure_name"):
            val = str(row.get(field) or "").strip()
            row[field] = val if val not in ("nan", "None", "") else None

        # Weekend / Holiday: "X" 여부만 필요하므로 대문자 strip
        wh = str(row.get("weekend_holiday") or "").strip().upper()
        row["weekend_holiday"] = "X" if wh == "X" else ""

        row["row_number"] = int(idx) + header_row_idx + 2  # 실제 엑셀 행 번호
        rows.append(row)

    return rows


# ─── 결과 엑셀 생성 ───────────────────────────────────────────

FILL_OK = PatternFill("solid", fgColor="C6EFCE")
FILL_DIFF = PatternFill("solid", fgColor="FFC7CE")
FILL_NO_RATE = PatternFill("solid", fgColor="FFEB9C")
FILL_SKIP = PatternFill("solid", fgColor="F2F2F2")
FILL_HEADER = PatternFill("solid", fgColor="1a73e8")

STATUS_FILL = {
    "OK": FILL_OK,
    "DIFF": FILL_DIFF,
    "NO_RATE": FILL_NO_RATE,
    "SKIP": FILL_SKIP,
}


def generate_results_excel(results: list) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "정산검증결과"

    headers = [
        "행번호", "컨테이너번호", "운송일자", "픽업지코드", "픽업지명",
        "ODCY코드", "ODCY명", "도착지코드", "도착지명", "컨테이너유형", "위험물", "수량", "주말/휴일", "티어번호",
        # TRKV
        "TRKV단가", "TRKV청구금액", "TRKV예상금액", "TRKV차이금액", "TRKV상태",
        # 보관료
        "보관료청구금액", "보관료예상금액", "보관료차이금액", "보관료상태",
        # 상하차료
        "상하차료청구금액", "상하차료예상금액", "상하차료차이금액", "상하차료상태",
        # 셔틀비용
        "셔틀료청구금액", "셔틀료예상금액", "셔틀료차이금액", "셔틀료상태",
        "종합상태",
    ]

    # 헤더 행
    ws.append(headers)
    for cell in ws[1]:
        cell.fill = FILL_HEADER
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center")

    money_cols = {headers.index(h) + 1 for h in headers if "금액" in h}

    for r in results:
        # dict 또는 ORM 객체 모두 지원
        g = (lambda k: r.get(k) if isinstance(r, dict) else getattr(r, k, None))
        row_data = [
            g("row_number"), g("container_no"), g("transport_date"),
            g("pickup_code"), g("pickup_name"),
            g("odcy_code"), g("odcy_name"),
            g("dest_code"), g("dest_name"),
            g("container_type"), "Y" if g("dg_flag") else "N", g("quantity"), g("weekend_holiday") or "", g("tier_number"),
            g("trkv_unit_rate"), g("trkv_actual"), g("trkv_expected"), g("trkv_diff"), g("trkv_status"),
            g("storage_actual"), g("storage_expected"), g("storage_diff"), g("storage_status"),
            g("handling_actual"), g("handling_expected"), g("handling_diff"), g("handling_status"),
            g("shuttle_actual"), g("shuttle_expected"), g("shuttle_diff"), g("shuttle_status"),
            g("overall_status"),
        ]
        ws.append(row_data)
        excel_row = ws.max_row

        # 행 색상 (overall_status 기준)
        fill = STATUS_FILL.get(g("overall_status") or "", FILL_SKIP)
        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=excel_row, column=col_idx)
            cell.fill = fill
            if col_idx in money_cols and cell.value is not None:
                cell.number_format = '#,##0'

    # 컬럼 너비 자동 조정
    for col_idx, col_cells in enumerate(ws.columns, 1):
        max_len = max((len(str(c.value or "")) for c in col_cells), default=8)
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 30)

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()

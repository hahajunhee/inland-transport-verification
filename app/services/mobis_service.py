"""
3단계 모비스 검증 서비스
엑셀 업로드 → 컨테이너번호+C/INV번호 중복제거 → 구분별 AE:AI 합계 비교 → 오류 판정
"""
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter


# ─── 열제목 매핑 ────────────────────────────────────────────────────────
# 엑셀 1:2행 병합 헤더에서 찾아야 할 컬럼명
_REQUIRED_HEADERS = ["컨테이너 번호", "C/INV\n번호", "구분"]
_COST_HEADERS = ["내륙운임", "보관료", "상하차료", "셔틀료", "대기료"]

# 구분 카테고리 (원본 텍스트 그대로)
CAT_GROVE_ODCY = "GROVE ODCY 전송"
CAT_GROVE = "GROVE 전송"
CAT_MOBIS = "MOBIS 산출"
_CATEGORIES = [CAT_GROVE_ODCY, CAT_GROVE, CAT_MOBIS]


def _safe_float(value) -> float:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).replace(",", "").strip()
    try:
        return float(s)
    except ValueError:
        return 0.0


def _find_header_columns(df: pd.DataFrame) -> dict:
    """
    1:2행 병합 헤더에서 필요 컬럼 위치를 찾는다.
    반환: {internal_key: column_index}
    """
    # 0행과 1행을 합쳐서 검색 (병합 셀은 한쪽에만 값이 있음)
    nrows = min(5, len(df))
    col_map = {}

    for ci in range(len(df.columns)):
        candidates = []
        for ri in range(nrows):
            val = str(df.iloc[ri, ci]).strip()
            if val and val != "nan":
                candidates.append(val)
        combined = "\n".join(candidates)

        # 정확 매칭 (줄바꿈 포함 비교)
        for header in _REQUIRED_HEADERS + _COST_HEADERS:
            norm_header = header.replace("\n", "").replace(" ", "")
            norm_combined = combined.replace("\n", "").replace(" ", "")
            if norm_header in norm_combined:
                col_map[header] = ci

        # 개별 행에서도 매칭 시도
        for val in candidates:
            norm_val = val.replace("\n", "").replace(" ", "")
            for header in _REQUIRED_HEADERS + _COST_HEADERS:
                norm_header = header.replace("\n", "").replace(" ", "")
                if norm_val == norm_header and header not in col_map:
                    col_map[header] = ci

    return col_map


def _find_data_start(df: pd.DataFrame, col_map: dict) -> int:
    """헤더 행 이후 실제 데이터 시작 행 인덱스 반환."""
    container_col = col_map.get("컨테이너 번호")
    gubun_col = col_map.get("구분")

    if container_col is None or gubun_col is None:
        return 2  # fallback

    for ri in range(len(df)):
        val = str(df.iloc[ri, gubun_col]).strip()
        # 구분 값이 3개 카테고리 중 하나이면 데이터 시작
        if val in _CATEGORIES:
            return ri
        # 또는 컨테이너 번호 패턴 (영문+숫자 11자리)
        cont_val = str(df.iloc[ri, container_col]).strip()
        if len(cont_val) >= 10 and cont_val[:4].isalpha():
            return ri

    return 2


def parse_mobis_excel(file_bytes: bytes) -> dict:
    """
    모비스 검증 엑셀 파싱.
    반환: {
        "rows": [...],
        "cost_headers": ["내륙운임", "보관료", ...],
        "col_map": {...}
    }
    """
    df = pd.read_excel(BytesIO(file_bytes), header=None, dtype=str)

    col_map = _find_header_columns(df)

    # 필수 컬럼 확인
    missing = []
    for h in _REQUIRED_HEADERS:
        if h not in col_map:
            missing.append(h)
    if missing:
        raise ValueError(f"필수 컬럼을 찾을 수 없습니다: {', '.join(missing)}")

    # 비용 컬럼 확인
    found_costs = [h for h in _COST_HEADERS if h in col_map]
    if not found_costs:
        raise ValueError("비용 컬럼(내륙운임, 보관료, 상하차료, 셔틀료, 대기료)을 찾을 수 없습니다.")

    data_start = _find_data_start(df, col_map)

    rows = []
    for ri in range(data_start, len(df)):
        container_no = str(df.iloc[ri, col_map["컨테이너 번호"]]).strip()
        c_inv = str(df.iloc[ri, col_map["C/INV\n번호"]]).strip()
        gubun = str(df.iloc[ri, col_map["구분"]]).strip()

        if container_no in ("nan", "None", "") and c_inv in ("nan", "None", ""):
            continue

        container_no = "" if container_no in ("nan", "None") else container_no
        c_inv = "" if c_inv in ("nan", "None") else c_inv
        gubun = "" if gubun in ("nan", "None") else gubun

        costs = {}
        for h in _COST_HEADERS:
            if h in col_map:
                costs[h] = _safe_float(df.iloc[ri, col_map[h]])
            else:
                costs[h] = 0.0

        rows.append({
            "container_no": container_no,
            "c_inv_no": c_inv,
            "gubun": gubun,
            "costs": costs,
            "cost_sum": sum(costs.values()),
            "row_number": ri + 1,  # 엑셀 행번호 (1-based)
        })

    return {
        "rows": rows,
        "cost_headers": found_costs,
    }


def run_mobis_verification(filename: str, parsed: dict) -> dict:
    """
    모비스 검증 실행.
    (컨테이너번호, C/INV번호) 중복 제거 → 구분별 합계 비교 → 오류 판정.
    """
    rows = parsed["rows"]
    cost_headers = parsed["cost_headers"]

    # (container_no, c_inv_no) 별로 그룹핑
    groups: dict[tuple, list] = {}
    group_order = []
    for row in rows:
        key = (row["container_no"], row["c_inv_no"])
        if key not in groups:
            groups[key] = []
            group_order.append(key)
        groups[key].append(row)

    results = []
    error_count = 0
    ok_count = 0

    for key in group_order:
        container_no, c_inv_no = key
        group_rows = groups[key]

        # 구분별 합계 계산
        cat_sums = {}
        cat_details = {}
        for cat in _CATEGORIES:
            matching = [r for r in group_rows if r["gubun"] == cat]
            if matching:
                total = sum(r["cost_sum"] for r in matching)
                detail = {}
                for h in cost_headers:
                    detail[h] = sum(r["costs"].get(h, 0.0) for r in matching)
                cat_sums[cat] = total
                cat_details[cat] = detail
            else:
                cat_sums[cat] = None
                cat_details[cat] = None

        grove_odcy_sum = cat_sums.get(CAT_GROVE_ODCY) or 0.0
        grove_sum = cat_sums.get(CAT_GROVE)
        mobis_sum = cat_sums.get(CAT_MOBIS)

        # 오류 판정
        errors = []
        # 1) GROVE ODCY 전송에 0이 아닌 값 존재
        if abs(grove_odcy_sum) >= 1:
            errors.append(f"GROVE ODCY 전송 합계 {grove_odcy_sum:,.0f}원 (0이 아님)")

        # 2) GROVE 전송 ≠ MOBIS 산출
        if grove_sum is not None and mobis_sum is not None:
            if abs(grove_sum - mobis_sum) >= 1:
                errors.append(f"GROVE 전송({grove_sum:,.0f}) ≠ MOBIS 산출({mobis_sum:,.0f})")
        elif grove_sum is not None or mobis_sum is not None:
            errors.append("GROVE 전송 또는 MOBIS 산출 데이터 누락")

        is_error = len(errors) > 0
        if is_error:
            error_count += 1
        else:
            ok_count += 1

        result = {
            "container_no": container_no,
            "c_inv_no": c_inv_no,
            "grove_odcy_sum": grove_odcy_sum,
            "grove_sum": grove_sum if grove_sum is not None else 0.0,
            "mobis_sum": mobis_sum if mobis_sum is not None else 0.0,
            "grove_odcy_detail": cat_details.get(CAT_GROVE_ODCY),
            "grove_detail": cat_details.get(CAT_GROVE),
            "mobis_detail": cat_details.get(CAT_MOBIS),
            "diff": (grove_sum or 0.0) - (mobis_sum or 0.0),
            "is_error": is_error,
            "error_reasons": errors,
            "status": "오류" if is_error else "정상",
        }
        results.append(result)

    return {
        "filename": filename,
        "total_groups": len(results),
        "error_count": error_count,
        "ok_count": ok_count,
        "cost_headers": cost_headers,
        "results": results,
    }


def generate_mobis_report(verification: dict) -> bytes:
    """모비스 검증 결과 엑셀 생성."""
    wb = Workbook()
    ws = wb.active
    ws.title = "모비스검증결과"

    cost_headers = verification.get("cost_headers", _COST_HEADERS)

    FILL_ERROR = PatternFill("solid", fgColor="FFC7CE")
    FILL_OK = PatternFill("solid", fgColor="FFFFFF")
    FILL_HEADER = PatternFill("solid", fgColor="4472C4")
    FILL_HEADER2 = PatternFill("solid", fgColor="D6E4F0")

    FONT_HEADER = Font(bold=True, color="FFFFFF", size=10)
    FONT_HEADER2 = Font(bold=True, color="1F4E79", size=10)
    FONT_ERROR = Font(color="9C0006", size=10)
    FONT_NORMAL = Font(size=10)

    headers = ["컨테이너 번호", "C/INV 번호",
               f"GROVE ODCY 전송\n합계", f"GROVE 전송\n합계", f"MOBIS 산출\n합계",
               "차이\n(GROVE-MOBIS)", "결과", "오류 사유"]

    # Row 1: 헤더
    ws.append(headers)
    for ci, cell in enumerate(ws[1], 1):
        cell.fill = FILL_HEADER
        cell.font = FONT_HEADER
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 36

    money_fmt = '#,##0'
    for r in verification["results"]:
        row_data = [
            r["container_no"],
            r["c_inv_no"],
            r["grove_odcy_sum"],
            r["grove_sum"],
            r["mobis_sum"],
            r["diff"],
            r["status"],
            " / ".join(r["error_reasons"]) if r["error_reasons"] else "",
        ]
        ws.append(row_data)
        excel_row = ws.max_row
        is_err = r["is_error"]
        for ci in range(1, len(row_data) + 1):
            cell = ws.cell(row=excel_row, column=ci)
            if is_err:
                cell.fill = FILL_ERROR
                cell.font = FONT_ERROR
            else:
                cell.fill = FILL_OK
                cell.font = FONT_NORMAL
            cell.alignment = Alignment(vertical="center")
            if ci in (3, 4, 5, 6):
                cell.number_format = money_fmt

    # 열 너비
    widths = [18, 18, 18, 18, 18, 18, 8, 40]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.freeze_panes = "A2"

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()

# -*- coding: utf-8 -*-
"""
거래관리.xlsm 자동 생성 스크립트
win32com.client을 사용하여 Excel을 직접 제어
"""
import os
import sys
import win32com.client as win32
import pythoncom

OUTPUT_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "거래관리.xlsm")

# ── 색상 상수 (Excel BGR int) ──────────────────────────────────────────────
CLR_HEADER_BG   = 0x2D5A8E   # 진청 (헤더 배경)
CLR_HEADER_FG   = 0xFFFFFF   # 흰색 글자
CLR_ALT_ROW     = 0xF2F2F2   # 연회색 (짝수행)
CLR_ACCENT      = 0x1F7A4D   # 초록 강조
CLR_WARN        = 0x0000CC   # 빨강 경고
CLR_SHEET_STMT  = 0xFFFFFF   # 명세서 탭 = 흰색 (print-only)

# ── 명세서 인쇄 최대 품목행 수 ────────────────────────────────────────────
MAX_ITEM_ROWS = 15

def rgb(r, g, b):
    """Excel은 BGR 정수"""
    return r + (g << 8) + (b << 16)

def set_header_row(ws, col_headers, row=1):
    """헤더행 텍스트 + 서식"""
    for i, h in enumerate(col_headers, 1):
        c = ws.Cells(row, i)
        c.Value = h
        c.Font.Bold = True
        c.Font.Color = CLR_HEADER_FG
        c.Interior.Color = CLR_HEADER_BG
        c.HorizontalAlignment = -4108  # xlCenter
    ws.Rows(row).RowHeight = 22

def freeze_row(ws, row=1):
    ws.Application.ActiveWindow.SplitRow = row
    ws.Application.ActiveWindow.FreezePanes = True

def border_range(rng):
    """얇은 선 테두리"""
    for idx in [7, 8, 9, 10]:  # xlEdgeLeft/Right/Top/Bottom
        rng.Borders(idx).LineStyle = 1
        rng.Borders(idx).Weight = 2
    for idx in [11, 12]:  # xlInsideVertical/Horizontal
        rng.Borders(idx).LineStyle = 1
        rng.Borders(idx).Weight = 2

def add_validation_list(ws, col_letter, values, start_row=2, max_row=10000):
    rng = ws.Range(f"{col_letter}{start_row}:{col_letter}{max_row}")
    rng.Validation.Delete()
    rng.Validation.Add(
        Type=3,          # xlValidateList
        AlertStyle=1,    # xlValidAlertStop
        Operator=1,
        Formula1=f'"{",".join(values)}"'
    )

def money_fmt(ws, col_letter, start_row=2, end_row=10000):
    ws.Range(f"{col_letter}{start_row}:{col_letter}{end_row}").NumberFormat = '#,##0'

def pct_fmt(ws, col_letter, start_row=2, end_row=10000):
    ws.Range(f"{col_letter}{start_row}:{col_letter}{end_row}").NumberFormat = '0.0%'

def date_fmt(ws, col_letter, start_row=2, end_row=10000):
    ws.Range(f"{col_letter}{start_row}:{col_letter}{end_row}").NumberFormat = 'YYYY-MM-DD'

def datetime_fmt(ws, col_letter, start_row=2, end_row=10000):
    ws.Range(f"{col_letter}{start_row}:{col_letter}{end_row}").NumberFormat = 'YYYY-MM-DD HH:MM'

def set_col_widths(ws, widths):
    """widths: list of (col_letter, width)"""
    for col_letter, w in widths:
        ws.Columns(col_letter).ColumnWidth = w

# ══════════════════════════════════════════════════════════════════════════════
# 1. 품목DB 시트
# ══════════════════════════════════════════════════════════════════════════════
def setup_품목DB(ws):
    ws.Name = "품목DB"
    ws.Tab.ColorIndex = 4  # 녹색

    headers = ["품목코드","품목명","규격","단위","단가","마진율","부가세여부","비고","등록일","수정일"]
    set_header_row(ws, headers)

    set_col_widths(ws, [
        ("A", 12), ("B", 16), ("C", 12), ("D", 7),
        ("E", 12), ("F", 8), ("G", 10), ("H", 20),
        ("I", 12), ("J", 12)
    ])

    money_fmt(ws, "E")
    pct_fmt(ws, "F")
    date_fmt(ws, "I")
    date_fmt(ws, "J")
    add_validation_list(ws, "D", ["kg","ea","box","g","개","묶음"])
    add_validation_list(ws, "G", ["Y","N"])

    # 초기 품목 데이터
    from datetime import date
    today = str(date.today())
    items = [
        ("ITM0001","돈가그리살","200g*4","kg",25000,0.12,"Y","",today,today),
        ("ITM0002","우진갈비살","2.5cm컷","kg",59000,0.15,"Y","",today,today),
        ("ITM0003","우갈비살","손질/채","kg",25000,0.12,"Y","",today,today),
        ("ITM0004","양꼬치","100g","ea",2100,0.18,"Y","",today,today),
    ]
    for r, row in enumerate(items, 2):
        for c, val in enumerate(row, 1):
            ws.Cells(r, c).Value = val

    ws.Activate()
    freeze_row(ws)

# ══════════════════════════════════════════════════════════════════════════════
# 2. 거래처DB 시트
# ══════════════════════════════════════════════════════════════════════════════
def setup_거래처DB(ws):
    ws.Name = "거래처DB"
    ws.Tab.ColorIndex = 8  # 하늘색

    headers = ["거래처코드","상호","성명","등록번호","주소","업태","업종","TEL","FAX","전일잔액","누적거래금액","등록일","수정일"]
    set_header_row(ws, headers)

    set_col_widths(ws, [
        ("A", 12), ("B", 16), ("C", 10), ("D", 14),
        ("E", 30), ("F", 12), ("G", 12), ("H", 14),
        ("I", 14), ("J", 14), ("K", 14), ("L", 12), ("M", 12)
    ])

    money_fmt(ws, "J")
    money_fmt(ws, "K")
    date_fmt(ws, "L")
    date_fmt(ws, "M")

    ws.Activate()
    freeze_row(ws)

# ══════════════════════════════════════════════════════════════════════════════
# 3. 거래헤더DB 시트
# ══════════════════════════════════════════════════════════════════════════════
def setup_거래헤더DB(ws):
    ws.Name = "거래헤더DB"
    ws.Tab.ColorIndex = 46  # 주황

    headers = ["거래ID","거래일자","거래처코드","상호","전일잔액","공급가액합","부가세합","합계금액","총액","금일입금액","금일잔액","출력담당자","인쇄여부","등록일시"]
    set_header_row(ws, headers)

    set_col_widths(ws, [
        ("A", 18), ("B", 12), ("C", 12), ("D", 16),
        ("E", 14), ("F", 14), ("G", 12), ("H", 14),
        ("I", 14), ("J", 14), ("K", 14), ("L", 12),
        ("M", 10), ("N", 16)
    ])

    for col in ["E","F","G","H","I","J","K"]:
        money_fmt(ws, col)
    date_fmt(ws, "B")
    datetime_fmt(ws, "N")
    add_validation_list(ws, "M", ["Y","N"])

    ws.Activate()
    freeze_row(ws)

# ══════════════════════════════════════════════════════════════════════════════
# 4. 거래상세DB 시트
# ══════════════════════════════════════════════════════════════════════════════
def setup_거래상세DB(ws):
    ws.Name = "거래상세DB"
    ws.Tab.ColorIndex = 44  # 연주황

    headers = ["상세ID","거래ID","순번","품목코드","품목명","규격","단위","수량","종량","단가","금액","부가세여부","부가세","이력번호","도축장","등록일시"]
    set_header_row(ws, headers)

    set_col_widths(ws, [
        ("A", 14), ("B", 18), ("C", 6), ("D", 12),
        ("E", 16), ("F", 12), ("G", 7), ("H", 8),
        ("I", 8), ("J", 12), ("K", 12), ("L", 10),
        ("M", 12), ("N", 16), ("O", 16), ("P", 16)
    ])

    money_fmt(ws, "J")
    money_fmt(ws, "K")
    money_fmt(ws, "M")
    ws.Range("H2:H10000").NumberFormat = '#,##0.##'
    ws.Range("I2:I10000").NumberFormat = '#,##0.##'
    add_validation_list(ws, "L", ["Y","N"])
    datetime_fmt(ws, "P")

    ws.Activate()
    freeze_row(ws)

# ══════════════════════════════════════════════════════════════════════════════
# 5. 거래명세서 시트  (인쇄 전용 서식)
# ══════════════════════════════════════════════════════════════════════════════
def setup_거래명세서(xl, ws):
    ws.Name = "거래명세서"
    ws.Tab.Color = rgb(220, 50, 50)  # 빨강 탭

    # 컬럼폭 설정 (A~K)
    widths = [("A",8),("B",16),("C",6),("D",8),("E",8),("F",12),("G",14),("H",12),("I",12),("J",14),("K",14)]
    set_col_widths(ws, widths)

    def mc(r1, c1, r2, c2):
        """셀 병합"""
        ws.Range(ws.Cells(r1,c1), ws.Cells(r2,c2)).Merge()

    def cv(r, c, v):
        ws.Cells(r, c).Value = v

    def bold(r, c):
        ws.Cells(r, c).Font.Bold = True

    def center(r, c):
        ws.Cells(r, c).HorizontalAlignment = -4108

    def vcenter(r, c):
        ws.Cells(r, c).VerticalAlignment = -4108

    def bg(r, c, color):
        ws.Cells(r, c).Interior.Color = color

    def fontsize(r, c, sz):
        ws.Cells(r, c).Font.Size = sz

    def named(r, c, name):
        ws.Cells(r, c).Name = name

    # ── Row 1: 제목 ──────────────────────────────────────────────────────
    mc(1,1,1,11)
    cv(1,1,"거  래  내  역  서")
    ws.Rows(1).RowHeight = 35
    ws.Cells(1,1).Font.Size = 20
    ws.Cells(1,1).Font.Bold = True
    center(1,1); vcenter(1,1)

    # ── Row 2: 날짜·담당자 ──────────────────────────────────────────────
    ws.Rows(2).RowHeight = 20
    mc(2,1,2,3); cv(2,1,"거래 년 월 일 :"); bold(2,1)
    mc(2,4,2,6); named(2,4,"ns_TxnDate")
    mc(2,7,2,8); cv(2,7,"출력담당자 :"); bold(2,7)
    mc(2,9,2,11); named(2,9,"ns_Operator")

    # ── Row 3: 공급자/공급받는자 레이블 ─────────────────────────────────
    ws.Rows(3).RowHeight = 20
    mc(3,1,3,5); cv(3,1,"  공  급  자  (인)"); bold(3,1)
    mc(3,6,3,11); cv(3,6,"  공  급  받  는  자"); bold(3,6)
    ws.Cells(3,1).Interior.Color = rgb(220,230,241)
    ws.Cells(3,6).Interior.Color = rgb(220,230,241)

    # 공급자·공급받는자 블록 (Row 4~9)
    labels_left  = ["등록번호","상호","성명","주소","업태","TEL"]
    labels_right = ["등록번호","상호","성명","주소","업태","TEL"]
    named_right  = ["ns_CustReg","ns_CustName","ns_CustContact","ns_CustAddr","ns_CustBizType","ns_CustTel"]
    # 고정 공급자 값 (모듈에서 채워넣을 placeholder — VBA에서 덮어씀)
    vals_left    = ["132-81-60911","굿푸드시스템","김수길","경기 구리시 동구릉로460번길 95","제조,도소매,서비스","031-555-6663"]

    for i, (ll, lr, nr, vl) in enumerate(zip(labels_left, labels_right, named_right, vals_left)):
        row = 4 + i
        ws.Rows(row).RowHeight = 18
        # 공급자
        mc(row,1,row,2); cv(row,1,f"  {ll}"); bold(row,1)
        mc(row,3,row,5); cv(row,3, vl)
        # 구분선
        ws.Cells(row,6).Value = ""
        # 공급받는자
        mc(row,6,row,7); cv(row,6,f"  {lr}"); bold(row,6)
        mc(row,8,row,11); named(row,8,nr)

    # Row 9: FAX 추가 행
    ws.Rows(9).RowHeight = 18
    mc(9,1,9,2); cv(9,1,"  FAX"); bold(9,1)
    mc(9,3,9,5); cv(9,3,"031-555-7774")
    mc(9,6,9,7); cv(9,6,"  FAX"); bold(9,6)
    mc(9,8,9,11); named(9,8,"ns_CustFax")

    # ── Row 10: 품목 테이블 헤더 ─────────────────────────────────────────
    ws.Rows(10).RowHeight = 22
    item_headers = ["품  목  명","단위","수량","종량","단가","금  액","부가세","이력번호","도축장"]
    col_map      = [     1,       2,    3,    4,    5,     6,      7,       8,       9]
    merge_map    = [(1,2),(3,3),(4,4),(5,5),(6,6),(7,8),(9,10),(11,11)]  # not used; manual below

    # 품목명(A~B), 단위(C), 수량(D), 종량(E), 단가(F), 금액(G~H), 부가세(I), 이력번호(J), 도축장(K)
    hdr_merges = [(10,1,10,2),(10,3,10,3),(10,4,10,4),(10,5,10,5),
                  (10,6,10,6),(10,7,10,8),(10,9,10,9),(10,10,10,10),(10,11,10,11)]
    hdr_vals   = ["품  목  명","단위","수량","종량","단가","금  액","부가세","이력번호","도축장"]
    for (r1,c1,r2,c2), v in zip(hdr_merges, hdr_vals):
        mc(r1,c1,r2,c2)
        ws.Cells(r1,c1).Value = v
        ws.Cells(r1,c1).Font.Bold = True
        ws.Cells(r1,c1).HorizontalAlignment = -4108
        ws.Cells(r1,c1).Interior.Color = rgb(198,224,180)

    # ── Rows 11~25: 품목 데이터 행 (15행) ───────────────────────────────
    ITEM_START_ROW = 11
    named(ITEM_START_ROW, 1, "ns_ItemStart")
    for r in range(ITEM_START_ROW, ITEM_START_ROW + MAX_ITEM_ROWS):
        ws.Rows(r).RowHeight = 18
        mc(r,1,r,2)   # 품목명
        mc(r,7,r,8)   # 금액
        if r % 2 == 0:
            ws.Range(ws.Cells(r,1), ws.Cells(r,11)).Interior.Color = rgb(242,242,242)
        for col_let, col_num in [("F",6),("G",7),("I",9)]:
            ws.Cells(r, col_num).NumberFormat = '#,##0'

    # ── Row 26: 총합계 행 ──────────────────────────────────────────────
    TOT_ROW = ITEM_START_ROW + MAX_ITEM_ROWS
    ws.Rows(TOT_ROW).RowHeight = 22
    mc(TOT_ROW,1,TOT_ROW,2); cv(TOT_ROW,1,"  총  합  계"); bold(TOT_ROW,1)
    ws.Cells(TOT_ROW,1).Interior.Color = rgb(198,224,180)
    named(TOT_ROW,4,"ns_SumQty")
    named(TOT_ROW,6,"ns_SumAmt"); ws.Cells(TOT_ROW,6).NumberFormat='#,##0'
    mc(TOT_ROW,7,TOT_ROW,8)
    named(TOT_ROW,7,"ns_SumVat"); ws.Cells(TOT_ROW,7).NumberFormat='#,##0'

    # ── Row 27: 잔액 요약 ─────────────────────────────────────────────
    FOOT1 = TOT_ROW + 1
    ws.Rows(FOOT1).RowHeight = 20
    mc(FOOT1,1,FOOT1,1); cv(FOOT1,1,"전 일 잔 액"); bold(FOOT1,1)
    named(FOOT1,2,"ns_PrevBal"); ws.Cells(FOOT1,2).NumberFormat='#,##0'
    mc(FOOT1,3,FOOT1,3); cv(FOOT1,3,"공 급 가 액"); bold(FOOT1,3)
    mc(FOOT1,4,FOOT1,5); named(FOOT1,4,"ns_Supply"); ws.Cells(FOOT1,4).NumberFormat='#,##0'
    mc(FOOT1,6,FOOT1,6); cv(FOOT1,6,"부 가 세"); bold(FOOT1,6)
    mc(FOOT1,7,FOOT1,8); named(FOOT1,7,"ns_VatTotal"); ws.Cells(FOOT1,7).NumberFormat='#,##0'
    mc(FOOT1,9,FOOT1,9); cv(FOOT1,9,"합 계 금 액"); bold(FOOT1,9)
    mc(FOOT1,10,FOOT1,11); named(FOOT1,10,"ns_InvTotal"); ws.Cells(FOOT1,10).NumberFormat='#,##0'

    FOOT2 = FOOT1 + 1
    ws.Rows(FOOT2).RowHeight = 20
    mc(FOOT2,1,FOOT2,1); cv(FOOT2,1,"총  액"); bold(FOOT2,1)
    mc(FOOT2,2,FOOT2,5); named(FOOT2,2,"ns_Grand"); ws.Cells(FOOT2,2).NumberFormat='#,##0'
    mc(FOOT2,6,FOOT2,6); cv(FOOT2,6,"금 일 입 금 액"); bold(FOOT2,6)
    mc(FOOT2,7,FOOT2,8); named(FOOT2,7,"ns_Payment"); ws.Cells(FOOT2,7).NumberFormat='#,##0'
    mc(FOOT2,9,FOOT2,9); cv(FOOT2,9,"금 일 잔 액"); bold(FOOT2,9)
    mc(FOOT2,10,FOOT2,11); named(FOOT2,10,"ns_TodayBal"); ws.Cells(FOOT2,10).NumberFormat='#,##0'

    # ── 전체 테두리 ───────────────────────────────────────────────────
    full_range = ws.Range(ws.Cells(1,1), ws.Cells(FOOT2,11))
    for idx in [7,8,9,10,11,12]:
        full_range.Borders(idx).LineStyle = 1
        full_range.Borders(idx).Weight = 2

    # ── 인쇄 설정 ─────────────────────────────────────────────────────
    ps = ws.PageSetup
    ps.PaperSize = 9           # xlPaperA4
    ps.Orientation = 1         # xlPortrait
    ps.Zoom = False
    ps.FitToPagesWide = 1
    ps.FitToPagesTall = 1
    ps.LeftMargin   = xl.InchesToPoints(0.59)
    ps.RightMargin  = xl.InchesToPoints(0.59)
    ps.TopMargin    = xl.InchesToPoints(0.79)
    ps.BottomMargin = xl.InchesToPoints(0.79)
    ws.PageSetup.PrintArea = ws.Range(ws.Cells(1,1), ws.Cells(FOOT2,11)).Address

    # 시트 보호 (VBA에서 해제하면서 작성)
    ws.Protect(Password="gf2024", UserInterfaceOnly=True)

# ══════════════════════════════════════════════════════════════════════════════
# 6. 대시보드 시트
# ══════════════════════════════════════════════════════════════════════════════
def setup_대시보드(ws):
    ws.Name = "대시보드"
    ws.Tab.Color = rgb(31, 122, 77)  # 진초록

    ws.Rows(1).RowHeight = 40
    ws.Range("A1:K1").Merge()
    ws.Cells(1,1).Value = "  거 래 관 리 시 스 템"
    ws.Cells(1,1).Font.Size = 24
    ws.Cells(1,1).Font.Bold = True
    ws.Cells(1,1).Font.Color = rgb(255,255,255)
    ws.Cells(1,1).Interior.Color = rgb(31,122,77)

    # 버튼 영역 Row 3
    ws.Rows(3).RowHeight = 14
    ws.Cells(3,1).Value = "[ 메 뉴 ]"
    ws.Cells(3,1).Font.Bold = True

    # 월별 요약표 헤더 Row 6
    ws.Rows(5).RowHeight = 14
    ws.Cells(5,1).Value = "▶ 월별 거래 요약"
    ws.Cells(5,1).Font.Bold = True

    monthly_headers = ["년월","거래건수","공급가액","부가세","합계금액","입금액","잔액합계"]
    for i, h in enumerate(monthly_headers, 1):
        c = ws.Cells(6, i)
        c.Value = h
        c.Font.Bold = True
        c.Interior.Color = rgb(31,122,77)
        c.Font.Color = rgb(255,255,255)
        c.HorizontalAlignment = -4108
    ws.Rows(6).RowHeight = 20

    # 거래처별 요약표 헤더 Row 6, Col 9~
    cust_headers = ["거래처명","이번달거래","공급가액합","잔액"]
    for i, h in enumerate(cust_headers, 9):
        c = ws.Cells(6, i)
        c.Value = h
        c.Font.Bold = True
        c.Interior.Color = rgb(31,122,77)
        c.Font.Color = rgb(255,255,255)
        c.HorizontalAlignment = -4108

    # 금액 서식
    for col in ["C","D","E","F","G"]:
        ws.Range(f"{col}7:{col}100").NumberFormat = '#,##0'
    for col in ["K","L","M"]:
        ws.Range(f"{col}7:{col}100").NumberFormat = '#,##0'

    set_col_widths(ws, [
        ("A",10),("B",10),("C",14),("D",12),("E",14),
        ("F",14),("G",14),("H",3),
        ("I",16),("J",12),("K",14),("L",14)
    ])

# ══════════════════════════════════════════════════════════════════════════════
# VBA 코드 삽입
# ══════════════════════════════════════════════════════════════════════════════
VBA_MODULES = {}

VBA_MODULES["modConstants"] = r"""
Option Explicit

' ── 공급자 고정정보 ──────────────────────────────────────────────────────────
Public Const SUPPLIER_NAME  As String = "굿푸드시스템"
Public Const SUPPLIER_REP   As String = "김수길"
Public Const SUPPLIER_REG   As String = "132-81-60911"
Public Const SUPPLIER_ADDR  As String = "경기 구리시 동구릉로460번길 95"
Public Const SUPPLIER_BIZ   As String = "제조,도소매,서비스"
Public Const SUPPLIER_TEL   As String = "031-555-6663"
Public Const SUPPLIER_FAX   As String = "031-555-7774"
Public Const SUPPLIER_BANK  As String = "하나 486-910008-32704"

' ── 시트명 ──────────────────────────────────────────────────────────────────
Public Const SH_ITEM_DB     As String = "품목DB"
Public Const SH_CUST_DB     As String = "거래처DB"
Public Const SH_TXN_HDR     As String = "거래헤더DB"
Public Const SH_TXN_DTL     As String = "거래상세DB"
Public Const SH_STATEMENT   As String = "거래명세서"
Public Const SH_DASHBOARD   As String = "대시보드"

' ── 기타 ────────────────────────────────────────────────────────────────────
Public Const MAX_ITEM_ROWS  As Integer = 15
Public Const VAT_RATE       As Double = 0.1
Public Const SHEET_PW       As String = "gf2024"
"""

VBA_MODULES["modUtils"] = r"""
Option Explicit

' ── 한국 금액 포맷 ───────────────────────────────────────────────────────────
Public Function FmtMoney(n As Long) As String
    FmtMoney = Format(n, "#,##0")
End Function

' ── 한국 날짜 표기 ───────────────────────────────────────────────────────────
Public Function FmtKorDate(d As Date) As String
    FmtKorDate = Format(d, "YYYY년 MM월 DD일")
End Function

' ── 날짜 문자열 파싱 (YYYY-MM-DD / YYYYMMDD / YYYY.MM.DD) ───────────────────
Public Function ParseDate(s As String) As Date
    Dim cleaned As String
    cleaned = Replace(Replace(Replace(s, "-", ""), ".", ""), "/", "")
    If Len(cleaned) = 8 Then
        ParseDate = DateSerial(CInt(Left(cleaned, 4)), CInt(Mid(cleaned, 5, 2)), CInt(Right(cleaned, 2)))
    Else
        ParseDate = CDate(s)
    End If
End Function

' ── 부가세 계산 (버림) ────────────────────────────────────────────────────────
Public Function CalcVAT(amount As Long, vatApply As String) As Long
    If UCase(vatApply) = "Y" Then
        CalcVAT = CLng(Int(amount * VAT_RATE))
    Else
        CalcVAT = 0
    End If
End Function

' ── 금액 계산 (수량 * 단가, 원 단위 반올림) ─────────────────────────────────
Public Function CalcAmount(qty As Double, unitPrice As Long) As Long
    CalcAmount = CLng(Round(qty * unitPrice, 0))
End Function

' ── 알림/에러/확인 ───────────────────────────────────────────────────────────
Public Sub ShowInfo(msg As String)
    MsgBox msg, vbInformation, "알림"
End Sub

Public Sub ShowError(msg As String)
    MsgBox msg, vbExclamation, "오류"
End Sub

Public Function Confirm(msg As String) As Boolean
    Confirm = (MsgBox(msg, vbYesNo + vbQuestion, "확인") = vbYes)
End Function

' ── 마지막 데이터 행 ─────────────────────────────────────────────────────────
Public Function GetLastRow(shName As String, Optional colIdx As Integer = 1) As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(shName)
    Dim lr As Long
    lr = ws.Cells(ws.Rows.Count, colIdx).End(xlUp).Row
    If lr < 1 Then lr = 1
    GetLastRow = lr
End Function

' ── 사업자번호 유효성 ────────────────────────────────────────────────────────
Public Function IsValidBizNo(s As String) As Boolean
    Dim digits As String
    digits = Replace(Replace(s, "-", ""), " ", "")
    If Len(digits) <> 10 Then IsValidBizNo = False: Exit Function
    Dim weights(0 To 9) As Integer
    Dim w() As Integer
    w = Array(1, 3, 7, 1, 3, 7, 1, 3, 5)
    Dim total As Integer
    total = 0
    Dim i As Integer
    For i = 0 To 8
        total = total + CInt(Mid(digits, i + 1, 1)) * w(i)
    Next i
    total = total + Int(CInt(Mid(digits, 9, 1)) * 5 / 10)
    IsValidBizNo = ((10 - (total Mod 10)) Mod 10 = CInt(Right(digits, 1)))
End Function
"""

VBA_MODULES["modDBHelper"] = r"""
Option Explicit

' ══════════════════════════════════════════════════════════════════════
' 사용자 정의 타입
' ══════════════════════════════════════════════════════════════════════
Public Type TItem
    ItemCode    As String
    ItemName    As String
    Spec        As String
    Unit        As String
    UnitPrice   As Long
    MarginRate  As Double
    VATApply    As String
    Remark      As String
End Type

Public Type TCustomer
    CustCode        As String
    CompanyName     As String
    ContactName     As String
    RegNumber       As String
    Address         As String
    BusinessType    As String
    BusinessCat     As String
    Tel             As String
    Fax             As String
    PrevBalance     As Long
End Type

Public Type TTxnHeader
    TxnID           As String
    TxnDate         As Date
    CustCode        As String
    CompanyName     As String
    PrevBalance     As Long
    SupplyTotal     As Long
    VATTotal        As Long
    InvoiceTotal    As Long
    GrandTotal      As Long
    PaymentToday    As Long
    TodayBalance    As Long
    PrintOperator   As String
    Printed         As String
End Type

Public Type TTxnDetail
    DetailID        As String
    TxnID           As String
    SeqNo           As Integer
    ItemCode        As String
    ItemName        As String
    Spec            As String
    Unit            As String
    Qty             As Double
    TotalWeight     As Double
    UnitPrice       As Long
    Amount          As Long
    VATApply        As String
    VAT             As Long
    TraceNo         As String
    SlaughterHouse  As String
End Type

' ══════════════════════════════════════════════════════════════════════
' ID 생성
' ══════════════════════════════════════════════════════════════════════
Public Function GenerateTxnID() As String
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_TXN_HDR)
    Dim lr As Long
    lr = GetLastRow(SH_TXN_HDR, 1)
    Dim seq As Long
    seq = lr  ' header row = 1, so first data row = 2 → seq = 1
    GenerateTxnID = "TXN" & Format(Now(), "YYYYMMDD") & Format(seq, "0000")
End Function

Public Function GenerateDetailID() As String
    Dim lr As Long
    lr = GetLastRow(SH_TXN_DTL, 1)
    GenerateDetailID = "DTL" & Format(lr, "00000")
End Function

Public Function GenerateItemCode() As String
    Dim lr As Long
    lr = GetLastRow(SH_ITEM_DB, 1)
    GenerateItemCode = "ITM" & Format(lr, "0000")
End Function

Public Function GenerateCustCode() As String
    Dim lr As Long
    lr = GetLastRow(SH_CUST_DB, 1)
    GenerateCustCode = "CST" & Format(lr, "0000")
End Function

' ══════════════════════════════════════════════════════════════════════
' 품목 조회
' ══════════════════════════════════════════════════════════════════════
Public Function GetItemByCode(itemCode As String) As TItem
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_ITEM_DB)
    Dim lr As Long: lr = GetLastRow(SH_ITEM_DB, 1)
    Dim r As Long
    For r = 2 To lr
        If ws.Cells(r, 1).Value = itemCode Then
            Dim t As TItem
            t.ItemCode   = ws.Cells(r, 1).Value
            t.ItemName   = ws.Cells(r, 2).Value
            t.Spec       = ws.Cells(r, 3).Value
            t.Unit       = ws.Cells(r, 4).Value
            t.UnitPrice  = CLng(ws.Cells(r, 5).Value)
            t.MarginRate = CDbl(ws.Cells(r, 6).Value)
            t.VATApply   = ws.Cells(r, 7).Value
            t.Remark     = ws.Cells(r, 8).Value
            GetItemByCode = t
            Exit Function
        End If
    Next r
End Function

Public Function GetItemList() As Variant
    ' Returns 2D array: (n, 0)=ItemCode, (n,1)=ItemName, (n,2)=Unit, (n,3)=UnitPrice, (n,4)=VATApply
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_ITEM_DB)
    Dim lr As Long: lr = GetLastRow(SH_ITEM_DB, 1)
    If lr < 2 Then GetItemList = Array(): Exit Function
    Dim arr() As Variant
    ReDim arr(0 To lr - 2, 0 To 4)
    Dim r As Long
    For r = 2 To lr
        arr(r - 2, 0) = ws.Cells(r, 1).Value
        arr(r - 2, 1) = ws.Cells(r, 2).Value
        arr(r - 2, 2) = ws.Cells(r, 4).Value
        arr(r - 2, 3) = CLng(ws.Cells(r, 5).Value)
        arr(r - 2, 4) = ws.Cells(r, 7).Value
    Next r
    GetItemList = arr
End Function

' ══════════════════════════════════════════════════════════════════════
' 거래처 조회
' ══════════════════════════════════════════════════════════════════════
Public Function GetCustomerByCode(custCode As String) As TCustomer
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_CUST_DB)
    Dim lr As Long: lr = GetLastRow(SH_CUST_DB, 1)
    Dim r As Long
    For r = 2 To lr
        If ws.Cells(r, 1).Value = custCode Then
            Dim t As TCustomer
            t.CustCode      = ws.Cells(r, 1).Value
            t.CompanyName   = ws.Cells(r, 2).Value
            t.ContactName   = ws.Cells(r, 3).Value
            t.RegNumber     = ws.Cells(r, 4).Value
            t.Address       = ws.Cells(r, 5).Value
            t.BusinessType  = ws.Cells(r, 6).Value
            t.BusinessCat   = ws.Cells(r, 7).Value
            t.Tel           = ws.Cells(r, 8).Value
            t.Fax           = ws.Cells(r, 9).Value
            t.PrevBalance   = CLng(ws.Cells(r, 10).Value)
            GetCustomerByCode = t
            Exit Function
        End If
    Next r
End Function

Public Function GetCustomerList() As Variant
    ' Returns 2D: (n,0)=CustCode, (n,1)=CompanyName
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_CUST_DB)
    Dim lr As Long: lr = GetLastRow(SH_CUST_DB, 1)
    If lr < 2 Then GetCustomerList = Array(): Exit Function
    Dim arr() As Variant
    ReDim arr(0 To lr - 2, 0 To 1)
    Dim r As Long
    For r = 2 To lr
        arr(r - 2, 0) = ws.Cells(r, 1).Value
        arr(r - 2, 1) = ws.Cells(r, 2).Value
    Next r
    GetCustomerList = arr
End Function

Public Function GetPrevBalance(custCode As String) As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_CUST_DB)
    Dim lr As Long: lr = GetLastRow(SH_CUST_DB, 1)
    Dim r As Long
    For r = 2 To lr
        If ws.Cells(r, 1).Value = custCode Then
            GetPrevBalance = CLng(ws.Cells(r, 10).Value)
            Exit Function
        End If
    Next r
    GetPrevBalance = 0
End Function

' ══════════════════════════════════════════════════════════════════════
' 거래 저장
' ══════════════════════════════════════════════════════════════════════
Public Sub SaveTransactionHeader(hdr As TTxnHeader)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_TXN_HDR)
    ws.Unprotect SHEET_PW
    Dim r As Long: r = GetLastRow(SH_TXN_HDR, 1) + 1
    ws.Cells(r, 1).Value  = hdr.TxnID
    ws.Cells(r, 2).Value  = hdr.TxnDate
    ws.Cells(r, 3).Value  = hdr.CustCode
    ws.Cells(r, 4).Value  = hdr.CompanyName
    ws.Cells(r, 5).Value  = hdr.PrevBalance
    ws.Cells(r, 6).Value  = hdr.SupplyTotal
    ws.Cells(r, 7).Value  = hdr.VATTotal
    ws.Cells(r, 8).Value  = hdr.InvoiceTotal
    ws.Cells(r, 9).Value  = hdr.GrandTotal
    ws.Cells(r, 10).Value = hdr.PaymentToday
    ws.Cells(r, 11).Value = hdr.TodayBalance
    ws.Cells(r, 12).Value = hdr.PrintOperator
    ws.Cells(r, 13).Value = "N"
    ws.Cells(r, 14).Value = Now()
    ws.Protect SHEET_PW
End Sub

Public Sub SaveTransactionDetail(dtl As TTxnDetail)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_TXN_DTL)
    ws.Unprotect SHEET_PW
    Dim r As Long: r = GetLastRow(SH_TXN_DTL, 1) + 1
    ws.Cells(r, 1).Value  = GenerateDetailID()
    ws.Cells(r, 2).Value  = dtl.TxnID
    ws.Cells(r, 3).Value  = dtl.SeqNo
    ws.Cells(r, 4).Value  = dtl.ItemCode
    ws.Cells(r, 5).Value  = dtl.ItemName
    ws.Cells(r, 6).Value  = dtl.Spec
    ws.Cells(r, 7).Value  = dtl.Unit
    ws.Cells(r, 8).Value  = dtl.Qty
    ws.Cells(r, 9).Value  = dtl.TotalWeight
    ws.Cells(r, 10).Value = dtl.UnitPrice
    ws.Cells(r, 11).Value = dtl.Amount
    ws.Cells(r, 12).Value = dtl.VATApply
    ws.Cells(r, 13).Value = dtl.VAT
    ws.Cells(r, 14).Value = dtl.TraceNo
    ws.Cells(r, 15).Value = dtl.SlaughterHouse
    ws.Cells(r, 16).Value = Now()
    ws.Protect SHEET_PW
End Sub

Public Sub UpdateCustomerBalance(custCode As String, newBalance As Long)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_CUST_DB)
    ws.Unprotect SHEET_PW
    Dim lr As Long: lr = GetLastRow(SH_CUST_DB, 1)
    Dim r As Long
    For r = 2 To lr
        If ws.Cells(r, 1).Value = custCode Then
            ws.Cells(r, 10).Value = newBalance
            ws.Cells(r, 13).Value = Now()
            ws.Protect SHEET_PW
            Exit Sub
        End If
    Next r
    ws.Protect SHEET_PW
End Sub

Public Function GetDetailsByTxnID(txnID As String) As Variant
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_TXN_DTL)
    Dim lr As Long: lr = GetLastRow(SH_TXN_DTL, 1)
    Dim rows() As Variant
    Dim cnt As Integer: cnt = 0

    Dim r As Long
    For r = 2 To lr
        If ws.Cells(r, 2).Value = txnID Then cnt = cnt + 1
    Next r

    If cnt = 0 Then GetDetailsByTxnID = Array(): Exit Function

    ReDim rows(0 To cnt - 1, 0 To 12)
    Dim idx As Integer: idx = 0
    For r = 2 To lr
        If ws.Cells(r, 2).Value = txnID Then
            rows(idx, 0)  = ws.Cells(r, 3).Value  ' SeqNo
            rows(idx, 1)  = ws.Cells(r, 5).Value  ' ItemName
            rows(idx, 2)  = ws.Cells(r, 6).Value  ' Spec
            rows(idx, 3)  = ws.Cells(r, 7).Value  ' Unit
            rows(idx, 4)  = ws.Cells(r, 8).Value  ' Qty
            rows(idx, 5)  = ws.Cells(r, 9).Value  ' TotalWeight
            rows(idx, 6)  = ws.Cells(r, 10).Value ' UnitPrice
            rows(idx, 7)  = ws.Cells(r, 11).Value ' Amount
            rows(idx, 8)  = ws.Cells(r, 12).Value ' VATApply
            rows(idx, 9)  = ws.Cells(r, 13).Value ' VAT
            rows(idx, 10) = ws.Cells(r, 14).Value ' TraceNo
            rows(idx, 11) = ws.Cells(r, 15).Value ' SlaughterHouse
            rows(idx, 12) = ws.Cells(r, 4).Value  ' ItemCode
            idx = idx + 1
        End If
    Next r
    GetDetailsByTxnID = rows
End Function

Public Function GetHeaderByTxnID(txnID As String) As TTxnHeader
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_TXN_HDR)
    Dim lr As Long: lr = GetLastRow(SH_TXN_HDR, 1)
    Dim r As Long
    For r = 2 To lr
        If ws.Cells(r, 1).Value = txnID Then
            Dim h As TTxnHeader
            h.TxnID         = ws.Cells(r, 1).Value
            h.TxnDate       = CDate(ws.Cells(r, 2).Value)
            h.CustCode      = ws.Cells(r, 3).Value
            h.CompanyName   = ws.Cells(r, 4).Value
            h.PrevBalance   = CLng(ws.Cells(r, 5).Value)
            h.SupplyTotal   = CLng(ws.Cells(r, 6).Value)
            h.VATTotal      = CLng(ws.Cells(r, 7).Value)
            h.InvoiceTotal  = CLng(ws.Cells(r, 8).Value)
            h.GrandTotal    = CLng(ws.Cells(r, 9).Value)
            h.PaymentToday  = CLng(ws.Cells(r, 10).Value)
            h.TodayBalance  = CLng(ws.Cells(r, 11).Value)
            h.PrintOperator = ws.Cells(r, 12).Value
            GetHeaderByTxnID = h
            Exit Function
        End If
    Next r
End Function
"""

VBA_MODULES["modStatement"] = r"""
Option Explicit

Public Sub GenerateStatement(txnID As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_STATEMENT)
    ws.Unprotect SHEET_PW

    ' 헤더 조회
    Dim hdr As TTxnHeader
    hdr = GetHeaderByTxnID(txnID)
    If hdr.TxnID = "" Then
        ShowError "거래ID를 찾을 수 없습니다: " & txnID
        ws.Protect SHEET_PW
        Exit Sub
    End If

    ' 거래처 조회
    Dim cust As TCustomer
    cust = GetCustomerByCode(hdr.CustCode)

    ' 품목 상세 조회
    Dim details As Variant
    details = GetDetailsByTxnID(txnID)

    Application.ScreenUpdating = False

    ' ── 공급자 고정 정보 (매번 갱신) ────────────────────────────────────
    ' (시트에 이미 고정값 있으나 상수로 재기록)

    ' ── 날짜 / 담당자 ──────────────────────────────────────────────────
    ws.Range("ns_TxnDate").Value  = FmtKorDate(hdr.TxnDate)
    ws.Range("ns_Operator").Value = hdr.PrintOperator

    ' ── 공급받는자 ─────────────────────────────────────────────────────
    ws.Range("ns_CustReg").Value      = cust.RegNumber
    ws.Range("ns_CustName").Value     = cust.CompanyName
    ws.Range("ns_CustContact").Value  = cust.ContactName
    ws.Range("ns_CustAddr").Value     = cust.Address
    ws.Range("ns_CustBizType").Value  = cust.BusinessType
    ws.Range("ns_CustTel").Value      = cust.Tel
    ws.Range("ns_CustFax").Value      = cust.Fax

    ' ── 품목 행 초기화 ─────────────────────────────────────────────────
    Dim startCell As Range
    Set startCell = ws.Range("ns_ItemStart")
    Dim startRow As Long: startRow = startCell.Row

    Dim i As Integer
    For i = 0 To MAX_ITEM_ROWS - 1
        Dim r As Long: r = startRow + i
        ws.Cells(r, 1).Value = ""  ' 품목명
        ws.Cells(r, 3).Value = ""  ' 수량
        ws.Cells(r, 4).Value = ""  ' 종량
        ws.Cells(r, 5).Value = ""  ' 단가
        ws.Cells(r, 6).Value = ""  ' 금액
        ws.Cells(r, 7).Value = ""  ' 부가세
        ws.Cells(r, 10).Value = "" ' 이력번호
        ws.Cells(r, 11).Value = "" ' 도축장
    Next i

    ' ── 품목 행 기록 ───────────────────────────────────────────────────
    Dim sumQty As Double: sumQty = 0
    Dim sumAmt As Long:   sumAmt = 0
    Dim sumVat As Long:   sumVat = 0

    If Not IsEmpty(details) And VarType(details) <> vbEmpty Then
        Dim nRows As Integer
        nRows = UBound(details, 1) + 1
        For i = 0 To nRows - 1
            r = startRow + i
            ws.Cells(r, 1).Value  = details(i, 1)   ' ItemName
            ws.Cells(r, 3).Value  = details(i, 4)   ' Qty
            ws.Cells(r, 4).Value  = details(i, 5)   ' TotalWeight
            ws.Cells(r, 5).Value  = details(i, 6)   ' UnitPrice
            ws.Cells(r, 6).Value  = details(i, 7)   ' Amount
            ws.Cells(r, 7).Value  = details(i, 9)   ' VAT
            ws.Cells(r, 10).Value = details(i, 10)  ' TraceNo
            ws.Cells(r, 11).Value = details(i, 11)  ' SlaughterHouse
            sumQty = sumQty + CDbl(details(i, 4))
            sumAmt = sumAmt + CLng(details(i, 7))
            sumVat = sumVat + CLng(details(i, 9))
        Next i
    End If

    ' ── 합계/잔액 행 ───────────────────────────────────────────────────
    ws.Range("ns_SumQty").Value   = sumQty
    ws.Range("ns_SumAmt").Value   = hdr.SupplyTotal
    ws.Range("ns_SumVat").Value   = hdr.VATTotal
    ws.Range("ns_PrevBal").Value  = hdr.PrevBalance
    ws.Range("ns_Supply").Value   = hdr.SupplyTotal
    ws.Range("ns_VatTotal").Value = hdr.VATTotal
    ws.Range("ns_InvTotal").Value = hdr.InvoiceTotal
    ws.Range("ns_Grand").Value    = hdr.GrandTotal
    ws.Range("ns_Payment").Value  = hdr.PaymentToday
    ws.Range("ns_TodayBal").Value = hdr.TodayBalance

    Application.ScreenUpdating = True
    ws.Protect SHEET_PW

    ' 인쇄여부 N→Y 업데이트
    MarkPrinted txnID
End Sub

Private Sub MarkPrinted(txnID As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_TXN_HDR)
    ws.Unprotect SHEET_PW
    Dim lr As Long: lr = GetLastRow(SH_TXN_HDR, 1)
    Dim r As Long
    For r = 2 To lr
        If ws.Cells(r, 1).Value = txnID Then
            ws.Cells(r, 13).Value = "Y"
            Exit For
        End If
    Next r
    ws.Protect SHEET_PW
End Sub

Public Sub PrintStatement()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_STATEMENT)
    ws.PrintOut Copies:=1, Collate:=True
End Sub

Public Sub ExportStatementToPDF(txnID As String, custName As String, txnDate As Date)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_STATEMENT)

    Dim outDir As String
    outDir = ThisWorkbook.Path & "\output"
    If Dir(outDir, vbDirectory) = "" Then MkDir outDir

    Dim fname As String
    fname = outDir & "\거래명세서_" & custName & "_" & Format(txnDate, "YYYYMMDD") & "_" & txnID & ".pdf"

    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=fname, Quality:=xlQualityStandard, _
        IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False

    ShowInfo "PDF 저장 완료:" & vbCrLf & fname
End Sub
"""

VBA_MODULES["modDashboard"] = r"""
Option Explicit

Public Sub UpdateDashboard()
    Application.ScreenUpdating = False

    Dim wsDash As Worksheet
    Set wsDash = ThisWorkbook.Sheets(SH_DASHBOARD)

    Dim wsHdr As Worksheet
    Set wsHdr = ThisWorkbook.Sheets(SH_TXN_HDR)

    Dim lr As Long: lr = GetLastRow(SH_TXN_HDR, 1)

    ' ── 기존 데이터 삭제 ──────────────────────────────────────────────
    wsDash.Range("A7:G100").ClearContents
    wsDash.Range("I7:L100").ClearContents

    If lr < 2 Then
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' ── 월별 집계 ─────────────────────────────────────────────────────
    Dim monthDict As Object
    Set monthDict = CreateObject("Scripting.Dictionary")

    Dim r As Long
    For r = 2 To lr
        Dim txnDate As Date
        On Error Resume Next
        txnDate = CDate(wsHdr.Cells(r, 2).Value)
        On Error GoTo 0
        If txnDate = 0 Then GoTo NextRow

        Dim ym As String
        ym = Format(txnDate, "YYYY-MM")

        Dim supply As Long: supply = CLng(wsHdr.Cells(r, 6).Value)
        Dim vat    As Long: vat    = CLng(wsHdr.Cells(r, 7).Value)
        Dim inv    As Long: inv    = CLng(wsHdr.Cells(r, 8).Value)
        Dim pay    As Long: pay    = CLng(wsHdr.Cells(r, 10).Value)
        Dim bal    As Long: bal    = CLng(wsHdr.Cells(r, 11).Value)

        If monthDict.Exists(ym) Then
            Dim arr As Variant: arr = monthDict(ym)
            arr(0) = arr(0) + 1
            arr(1) = arr(1) + supply
            arr(2) = arr(2) + vat
            arr(3) = arr(3) + inv
            arr(4) = arr(4) + pay
            arr(5) = bal   ' last known balance
            monthDict(ym) = arr
        Else
            monthDict.Add ym, Array(1, supply, vat, inv, pay, bal)
        End If
NextRow:
    Next r

    ' ── 월별 표 기록 ──────────────────────────────────────────────────
    Dim outRow As Long: outRow = 7
    Dim keys As Variant: keys = monthDict.Keys()
    ' 정렬 (버블 정렬)
    Dim ii As Integer, jj As Integer, tmp As Variant
    For ii = 0 To UBound(keys) - 1
        For jj = 0 To UBound(keys) - 1 - ii
            If keys(jj) > keys(jj + 1) Then
                tmp = keys(jj): keys(jj) = keys(jj + 1): keys(jj + 1) = tmp
            End If
        Next jj
    Next ii

    For ii = 0 To UBound(keys)
        Dim k As String: k = keys(ii)
        Dim v As Variant: v = monthDict(k)
        wsDash.Cells(outRow, 1).Value = k
        wsDash.Cells(outRow, 2).Value = v(0)
        wsDash.Cells(outRow, 3).Value = v(1)
        wsDash.Cells(outRow, 4).Value = v(2)
        wsDash.Cells(outRow, 5).Value = v(3)
        wsDash.Cells(outRow, 6).Value = v(4)
        wsDash.Cells(outRow, 7).Value = v(5)
        outRow = outRow + 1
    Next ii

    ' ── 거래처별 집계 (이번 달) ───────────────────────────────────────
    Dim thisMonth As String: thisMonth = Format(Now(), "YYYY-MM")
    Dim custDict As Object
    Set custDict = CreateObject("Scripting.Dictionary")

    For r = 2 To lr
        On Error Resume Next
        txnDate = CDate(wsHdr.Cells(r, 2).Value)
        On Error GoTo 0
        If Format(txnDate, "YYYY-MM") <> thisMonth Then GoTo NextRow2

        Dim custName As String: custName = wsHdr.Cells(r, 4).Value
        supply = CLng(wsHdr.Cells(r, 6).Value)
        bal = CLng(wsHdr.Cells(r, 11).Value)

        If custDict.Exists(custName) Then
            Dim ca As Variant: ca = custDict(custName)
            ca(0) = ca(0) + 1
            ca(1) = ca(1) + supply
            ca(2) = bal
            custDict(custName) = ca
        Else
            custDict.Add custName, Array(1, supply, bal)
        End If
NextRow2:
    Next r

    outRow = 7
    Dim ck As Variant
    For Each ck In custDict.Keys
        Dim cv As Variant: cv = custDict(ck)
        wsDash.Cells(outRow, 9).Value  = ck
        wsDash.Cells(outRow, 10).Value = cv(0)
        wsDash.Cells(outRow, 11).Value = cv(1)
        wsDash.Cells(outRow, 12).Value = cv(2)
        outRow = outRow + 1
    Next ck

    ' ── 차트 업데이트 ─────────────────────────────────────────────────
    Call RefreshCharts

    Application.ScreenUpdating = True
End Sub

Private Sub RefreshCharts()
    ' 월별 차트
    Dim wsDash As Worksheet
    Set wsDash = ThisWorkbook.Sheets(SH_DASHBOARD)

    Dim lastDataRow As Long
    lastDataRow = wsDash.Cells(wsDash.Rows.Count, 1).End(xlUp).Row
    If lastDataRow < 7 Then Exit Sub

    Dim cht As ChartObject
    On Error Resume Next
    Set cht = wsDash.ChartObjects("cht_Monthly")
    On Error GoTo 0

    If Not cht Is Nothing Then
        Dim monthRng As Range
        Set monthRng = wsDash.Range("A6:A" & lastDataRow & ",E6:E" & lastDataRow)
        cht.Chart.SetSourceData Source:=monthRng
    End If

    ' 거래처 파이차트
    Dim cht2 As ChartObject
    On Error Resume Next
    Set cht2 = wsDash.ChartObjects("cht_Customer")
    On Error GoTo 0

    Dim custLastRow As Long
    custLastRow = wsDash.Cells(wsDash.Rows.Count, 9).End(xlUp).Row

    If Not cht2 Is Nothing And custLastRow >= 7 Then
        Dim custRng As Range
        Set custRng = wsDash.Range("I6:I" & custLastRow & ",K6:K" & custLastRow)
        cht2.Chart.SetSourceData Source:=custRng
    End If
End Sub
"""

VBA_MODULES["modButtons"] = r"""
Option Explicit

Public Sub OpenNewTransaction()
    frmNewTransaction.Show
End Sub

Public Sub OpenProductMgmt()
    frmProductMgmt.Show
End Sub

Public Sub OpenCustomerMgmt()
    frmCustomerMgmt.Show
End Sub

Public Sub OpenHistory()
    frmTransactionHistory.Show
End Sub

Public Sub RefreshDashboard()
    Call UpdateDashboard
    ShowInfo "대시보드가 업데이트되었습니다."
End Sub
"""

# ══════════════════════════════════════════════════════════════════════════════
# UserForm FRM 코드
# ══════════════════════════════════════════════════════════════════════════════
FORM_CODE = {}

FORM_CODE["frmProductMgmt"] = r"""
Option Explicit

Private editMode As Boolean
Private editRow As Long

Private Sub UserForm_Initialize()
    Me.Caption = "품목 관리"
    RefreshList
    ClearFields
    editMode = False
End Sub

Private Sub RefreshList()
    lst_Products.Clear
    lst_Products.ColumnCount = 4
    lst_Products.ColumnWidths = "60;100;60;70"
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_ITEM_DB)
    Dim lr As Long: lr = GetLastRow(SH_ITEM_DB, 1)
    Dim r As Long
    For r = 2 To lr
        lst_Products.AddItem ws.Cells(r, 1).Value
        lst_Products.List(lst_Products.ListCount - 1, 1) = ws.Cells(r, 2).Value
        lst_Products.List(lst_Products.ListCount - 1, 2) = ws.Cells(r, 4).Value
        lst_Products.List(lst_Products.ListCount - 1, 3) = Format(CLng(ws.Cells(r, 5).Value), "#,##0")
    Next r
End Sub

Private Sub ClearFields()
    txt_ItemName.Value = ""
    txt_Spec.Value = ""
    cbo_Unit.Value = ""
    txt_UnitPrice.Value = ""
    txt_MarginRate.Value = ""
    opt_VAT_Y.Value = True
    txt_Remark.Value = ""
    editMode = False
    editRow = 0
    lbl_Mode.Caption = "[신규 입력]"
End Sub

Private Sub lst_Products_Click()
    If lst_Products.ListIndex < 0 Then Exit Sub
    Dim code As String
    code = lst_Products.List(lst_Products.ListIndex, 0)
    Dim item As TItem
    item = GetItemByCode(code)
    txt_ItemName.Value  = item.ItemName
    txt_Spec.Value      = item.Spec
    cbo_Unit.Value      = item.Unit
    txt_UnitPrice.Value = CStr(item.UnitPrice)
    txt_MarginRate.Value = CStr(item.MarginRate * 100)
    If item.VATApply = "Y" Then opt_VAT_Y.Value = True Else opt_VAT_N.Value = True
    txt_Remark.Value = item.Remark
    editMode = True
    ' find row
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SH_ITEM_DB)
    Dim lr As Long: lr = GetLastRow(SH_ITEM_DB, 1)
    Dim r As Long
    For r = 2 To lr
        If ws.Cells(r, 1).Value = code Then editRow = r: Exit For
    Next r
    lbl_Mode.Caption = "[수정 모드] " & item.ItemName
End Sub

Private Sub btn_New_Click()
    ClearFields
End Sub

Private Sub btn_Save_Click()
    If Trim(txt_ItemName.Value) = "" Then ShowError "품목명을 입력하세요.": Exit Sub
    If Trim(txt_UnitPrice.Value) = "" Then ShowError "단가를 입력하세요.": Exit Sub

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SH_ITEM_DB)
    ws.Unprotect SHEET_PW

    Dim vatVal As String
    If opt_VAT_Y.Value Then vatVal = "Y" Else vatVal = "N"

    If editMode And editRow > 1 Then
        ws.Cells(editRow, 2).Value = Trim(txt_ItemName.Value)
        ws.Cells(editRow, 3).Value = Trim(txt_Spec.Value)
        ws.Cells(editRow, 4).Value = cbo_Unit.Value
        ws.Cells(editRow, 5).Value = CLng(txt_UnitPrice.Value)
        ws.Cells(editRow, 6).Value = CDbl(txt_MarginRate.Value) / 100
        ws.Cells(editRow, 7).Value = vatVal
        ws.Cells(editRow, 8).Value = Trim(txt_Remark.Value)
        ws.Cells(editRow, 10).Value = Now()
    Else
        Dim r As Long: r = GetLastRow(SH_ITEM_DB, 1) + 1
        ws.Cells(r, 1).Value = GenerateItemCode()
        ws.Cells(r, 2).Value = Trim(txt_ItemName.Value)
        ws.Cells(r, 3).Value = Trim(txt_Spec.Value)
        ws.Cells(r, 4).Value = cbo_Unit.Value
        ws.Cells(r, 5).Value = CLng(txt_UnitPrice.Value)
        ws.Cells(r, 6).Value = CDbl(txt_MarginRate.Value) / 100
        ws.Cells(r, 7).Value = vatVal
        ws.Cells(r, 8).Value = Trim(txt_Remark.Value)
        ws.Cells(r, 9).Value = Now()
        ws.Cells(r, 10).Value = Now()
    End If

    ws.Protect SHEET_PW
    RefreshList
    ClearFields
    ShowInfo "저장 완료"
End Sub

Private Sub btn_Delete_Click()
    If Not editMode Or editRow <= 1 Then ShowError "삭제할 항목을 선택하세요.": Exit Sub
    If Not Confirm(lst_Products.List(lst_Products.ListIndex, 1) & " 품목을 삭제하시겠습니까?") Then Exit Sub
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SH_ITEM_DB)
    ws.Unprotect SHEET_PW
    ws.Rows(editRow).Delete
    ws.Protect SHEET_PW
    RefreshList
    ClearFields
    ShowInfo "삭제되었습니다."
End Sub

Private Sub btn_Close_Click()
    Unload Me
End Sub
"""

FORM_CODE["frmCustomerMgmt"] = r"""
Option Explicit

Private editMode As Boolean
Private editRow As Long

Private Sub UserForm_Initialize()
    Me.Caption = "거래처 관리"
    RefreshList
    ClearFields
End Sub

Private Sub RefreshList()
    lst_Customers.Clear
    lst_Customers.ColumnCount = 3
    lst_Customers.ColumnWidths = "60;120;90"
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SH_CUST_DB)
    Dim lr As Long: lr = GetLastRow(SH_CUST_DB, 1)
    Dim r As Long
    For r = 2 To lr
        lst_Customers.AddItem ws.Cells(r, 1).Value
        lst_Customers.List(lst_Customers.ListCount - 1, 1) = ws.Cells(r, 2).Value
        lst_Customers.List(lst_Customers.ListCount - 1, 2) = ws.Cells(r, 8).Value
    Next r
End Sub

Private Sub ClearFields()
    txt_CompanyName.Value = ""
    txt_ContactName.Value = ""
    txt_RegNumber.Value = ""
    txt_Address.Value = ""
    txt_BizType.Value = ""
    txt_BizCat.Value = ""
    txt_Tel.Value = ""
    txt_Fax.Value = ""
    txt_OpenBalance.Value = "0"
    editMode = False
    editRow = 0
    lbl_Mode.Caption = "[신규 입력]"
End Sub

Private Sub lst_Customers_Click()
    If lst_Customers.ListIndex < 0 Then Exit Sub
    Dim code As String
    code = lst_Customers.List(lst_Customers.ListIndex, 0)
    Dim cust As TCustomer
    cust = GetCustomerByCode(code)
    txt_CompanyName.Value = cust.CompanyName
    txt_ContactName.Value = cust.ContactName
    txt_RegNumber.Value   = cust.RegNumber
    txt_Address.Value     = cust.Address
    txt_BizType.Value     = cust.BusinessType
    txt_BizCat.Value      = cust.BusinessCat
    txt_Tel.Value         = cust.Tel
    txt_Fax.Value         = cust.Fax
    txt_OpenBalance.Value = CStr(cust.PrevBalance)
    editMode = True
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SH_CUST_DB)
    Dim lr As Long: lr = GetLastRow(SH_CUST_DB, 1)
    Dim r As Long
    For r = 2 To lr
        If ws.Cells(r, 1).Value = code Then editRow = r: Exit For
    Next r
    lbl_Mode.Caption = "[수정 모드] " & cust.CompanyName
End Sub

Private Sub btn_New_Click(): ClearFields: End Sub

Private Sub btn_Save_Click()
    If Trim(txt_CompanyName.Value) = "" Then ShowError "상호를 입력하세요.": Exit Sub

    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SH_CUST_DB)
    ws.Unprotect SHEET_PW

    If editMode And editRow > 1 Then
        ws.Cells(editRow, 2).Value  = Trim(txt_CompanyName.Value)
        ws.Cells(editRow, 3).Value  = Trim(txt_ContactName.Value)
        ws.Cells(editRow, 4).Value  = Trim(txt_RegNumber.Value)
        ws.Cells(editRow, 5).Value  = Trim(txt_Address.Value)
        ws.Cells(editRow, 6).Value  = Trim(txt_BizType.Value)
        ws.Cells(editRow, 7).Value  = Trim(txt_BizCat.Value)
        ws.Cells(editRow, 8).Value  = Trim(txt_Tel.Value)
        ws.Cells(editRow, 9).Value  = Trim(txt_Fax.Value)
        ws.Cells(editRow, 10).Value = CLng(txt_OpenBalance.Value)
        ws.Cells(editRow, 13).Value = Now()
    Else
        Dim r As Long: r = GetLastRow(SH_CUST_DB, 1) + 1
        ws.Cells(r, 1).Value  = GenerateCustCode()
        ws.Cells(r, 2).Value  = Trim(txt_CompanyName.Value)
        ws.Cells(r, 3).Value  = Trim(txt_ContactName.Value)
        ws.Cells(r, 4).Value  = Trim(txt_RegNumber.Value)
        ws.Cells(r, 5).Value  = Trim(txt_Address.Value)
        ws.Cells(r, 6).Value  = Trim(txt_BizType.Value)
        ws.Cells(r, 7).Value  = Trim(txt_BizCat.Value)
        ws.Cells(r, 8).Value  = Trim(txt_Tel.Value)
        ws.Cells(r, 9).Value  = Trim(txt_Fax.Value)
        ws.Cells(r, 10).Value = CLng(txt_OpenBalance.Value)
        ws.Cells(r, 11).Value = 0
        ws.Cells(r, 12).Value = Now()
        ws.Cells(r, 13).Value = Now()
    End If

    ws.Protect SHEET_PW
    RefreshList
    ClearFields
    ShowInfo "저장 완료"
End Sub

Private Sub btn_Delete_Click()
    If Not editMode Or editRow <= 1 Then ShowError "삭제할 항목을 선택하세요.": Exit Sub
    If Not Confirm(lst_Customers.List(lst_Customers.ListIndex, 1) & " 거래처를 삭제하시겠습니까?") Then Exit Sub
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SH_CUST_DB)
    ws.Unprotect SHEET_PW
    ws.Rows(editRow).Delete
    ws.Protect SHEET_PW
    RefreshList
    ClearFields
    ShowInfo "삭제되었습니다."
End Sub

Private Sub btn_Close_Click(): Unload Me: End Sub
"""

FORM_CODE["frmNewTransaction"] = r"""
Option Explicit

' 품목 행 컬렉션 (임시 저장)
Private Type TLineItem
    ItemCode    As String
    ItemName    As String
    Spec        As String
    Unit        As String
    Qty         As Double
    UnitPrice   As Long
    Amount      As Long
    VATApply    As String
    VAT         As Long
    TraceNo     As String
    Slaughter   As String
End Type

Private items(0 To 14) As TLineItem
Private itemCount As Integer
Private currentCustCode As String

Private Sub UserForm_Initialize()
    Me.Caption = "거래 입력"
    txt_TxnDate.Value = Format(Now(), "YYYY-MM-DD")
    txt_Payment.Value = "0"
    itemCount = 0

    ' 거래처 콤보박스 로드
    Dim data As Variant: data = GetCustomerList()
    cbo_Customer.Clear
    If Not IsEmpty(data) And VarType(data) <> vbEmpty Then
        Dim i As Integer
        For i = 0 To UBound(data, 1)
            cbo_Customer.AddItem data(i, 1)
            cbo_Customer.List(cbo_Customer.ListCount - 1, 1) = data(i, 0)
        Next i
    End If
    cbo_Customer.ColumnCount = 2
    cbo_Customer.ColumnWidths = "0;100"

    ' 품목 콤보박스 로드
    Dim idata As Variant: idata = GetItemList()
    cbo_Item.Clear
    If Not IsEmpty(idata) And VarType(idata) <> vbEmpty Then
        For i = 0 To UBound(idata, 1)
            cbo_Item.AddItem idata(i, 1)
            cbo_Item.List(cbo_Item.ListCount - 1, 1) = idata(i, 0)  ' ItemCode
            cbo_Item.List(cbo_Item.ListCount - 1, 2) = idata(i, 2)  ' Unit
            cbo_Item.List(cbo_Item.ListCount - 1, 3) = idata(i, 3)  ' UnitPrice
            cbo_Item.List(cbo_Item.ListCount - 1, 4) = idata(i, 4)  ' VATApply
        Next i
    End If
    cbo_Item.ColumnCount = 5
    cbo_Item.ColumnWidths = "100;0;0;0;0"

    RefreshGrid
    CalcTotals
End Sub

Private Sub cbo_Customer_Change()
    If cbo_Customer.ListIndex < 0 Then Exit Sub
    currentCustCode = cbo_Customer.List(cbo_Customer.ListIndex, 1)
    Dim bal As Long: bal = GetPrevBalance(currentCustCode)
    lbl_PrevBal.Caption = Format(bal, "#,##0") & " 원"
    CalcTotals
End Sub

Private Sub cbo_Item_Change()
    If cbo_Item.ListIndex < 0 Then Exit Sub
    lbl_Unit.Caption   = cbo_Item.List(cbo_Item.ListIndex, 2)
    lbl_UnitPrice.Caption = Format(CLng(cbo_Item.List(cbo_Item.ListIndex, 3)), "#,##0") & " 원"
    txt_Qty.SetFocus
End Sub

Private Sub txt_Qty_Change()
    UpdateAmountPreview
End Sub

Private Sub UpdateAmountPreview()
    If cbo_Item.ListIndex < 0 Then Exit Sub
    If Trim(txt_Qty.Value) = "" Then
        lbl_AmtPreview.Caption = "0 원"
        Exit Sub
    End If
    On Error Resume Next
    Dim qty As Double: qty = CDbl(txt_Qty.Value)
    Dim price As Long: price = CLng(cbo_Item.List(cbo_Item.ListIndex, 3))
    Dim amt As Long: amt = CalcAmount(qty, price)
    Dim vatApply As String: vatApply = cbo_Item.List(cbo_Item.ListIndex, 4)
    Dim vat As Long: vat = CalcVAT(amt, vatApply)
    lbl_AmtPreview.Caption = Format(amt, "#,##0") & " + VAT " & Format(vat, "#,##0") & " 원"
    On Error GoTo 0
End Sub

Private Sub btn_AddItem_Click()
    If cbo_Item.ListIndex < 0 Then ShowError "품목을 선택하세요.": Exit Sub
    If Trim(txt_Qty.Value) = "" Then ShowError "수량을 입력하세요.": Exit Sub

    Dim qty As Double
    On Error GoTo QtyErr
    qty = CDbl(txt_Qty.Value)
    If qty <= 0 Then ShowError "수량은 0보다 커야 합니다.": Exit Sub
    On Error GoTo 0

    If itemCount >= 15 Then ShowError "최대 15개 품목까지 입력 가능합니다.": Exit Sub

    Dim price As Long: price = CLng(cbo_Item.List(cbo_Item.ListIndex, 3))
    Dim vatApply As String: vatApply = cbo_Item.List(cbo_Item.ListIndex, 4)
    Dim amt As Long: amt = CalcAmount(qty, price)
    Dim vat As Long: vat = CalcVAT(amt, vatApply)

    items(itemCount).ItemCode  = cbo_Item.List(cbo_Item.ListIndex, 1)
    items(itemCount).ItemName  = cbo_Item.List(cbo_Item.ListIndex, 0)
    items(itemCount).Unit      = cbo_Item.List(cbo_Item.ListIndex, 2)
    items(itemCount).Qty       = qty
    items(itemCount).UnitPrice = price
    items(itemCount).Amount    = amt
    items(itemCount).VATApply  = vatApply
    items(itemCount).VAT       = vat
    items(itemCount).TraceNo   = Trim(txt_TraceNo.Value)
    items(itemCount).Slaughter = Trim(txt_Slaughter.Value)
    itemCount = itemCount + 1

    cbo_Item.ListIndex = -1
    txt_Qty.Value = ""
    txt_TraceNo.Value = ""
    txt_Slaughter.Value = ""
    lbl_AmtPreview.Caption = ""
    lbl_Unit.Caption = ""
    lbl_UnitPrice.Caption = ""

    RefreshGrid
    CalcTotals
    Exit Sub
QtyErr:
    ShowError "수량은 숫자로 입력하세요."
End Sub

Private Sub btn_DelItem_Click()
    If lst_Items.ListIndex < 0 Then ShowError "삭제할 항목을 선택하세요.": Exit Sub
    Dim idx As Integer: idx = lst_Items.ListIndex
    Dim i As Integer
    For i = idx To itemCount - 2
        items(i) = items(i + 1)
    Next i
    itemCount = itemCount - 1
    RefreshGrid
    CalcTotals
End Sub

Private Sub RefreshGrid()
    lst_Items.Clear
    lst_Items.ColumnCount = 7
    lst_Items.ColumnWidths = "20;100;40;50;60;60;50"
    Dim i As Integer
    For i = 0 To itemCount - 1
        lst_Items.AddItem CStr(i + 1)
        lst_Items.List(lst_Items.ListCount - 1, 1) = items(i).ItemName
        lst_Items.List(lst_Items.ListCount - 1, 2) = items(i).Unit
        lst_Items.List(lst_Items.ListCount - 1, 3) = Format(items(i).Qty, "#,##0.##")
        lst_Items.List(lst_Items.ListCount - 1, 4) = Format(items(i).UnitPrice, "#,##0")
        lst_Items.List(lst_Items.ListCount - 1, 5) = Format(items(i).Amount, "#,##0")
        lst_Items.List(lst_Items.ListCount - 1, 6) = Format(items(i).VAT, "#,##0")
    Next i
End Sub

Private Sub CalcTotals()
    Dim supplyTotal As Long: supplyTotal = 0
    Dim vatTotal    As Long: vatTotal    = 0
    Dim i As Integer
    For i = 0 To itemCount - 1
        supplyTotal = supplyTotal + items(i).Amount
        vatTotal    = vatTotal    + items(i).VAT
    Next i
    Dim invTotal As Long: invTotal = supplyTotal + vatTotal

    Dim prevBal As Long: prevBal = 0
    If currentCustCode <> "" Then prevBal = GetPrevBalance(currentCustCode)

    Dim grandTotal As Long: grandTotal = prevBal + invTotal

    Dim payment As Long: payment = 0
    On Error Resume Next
    payment = CLng(txt_Payment.Value)
    On Error GoTo 0

    Dim todayBal As Long: todayBal = grandTotal - payment

    lbl_SupplyTotal.Caption = Format(supplyTotal, "#,##0") & " 원"
    lbl_VATTotal.Caption    = Format(vatTotal,    "#,##0") & " 원"
    lbl_InvTotal.Caption    = Format(invTotal,    "#,##0") & " 원"
    lbl_Grand.Caption       = Format(grandTotal,  "#,##0") & " 원"
    lbl_TodayBal.Caption    = Format(todayBal,    "#,##0") & " 원"
End Sub

Private Sub txt_Payment_Change()
    CalcTotals
End Sub

Private Sub btn_Save_Click()
    If Not ValidateInput() Then Exit Sub
    SaveTransaction False
    ShowInfo "저장 완료!"
    Unload Me
End Sub

Private Sub btn_Print_Click()
    If Not ValidateInput() Then Exit Sub
    Dim txnID As String
    txnID = SaveTransaction(True)
    If txnID <> "" Then
        GenerateStatement txnID
        Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SH_STATEMENT)
        ws.Activate
        ShowInfo "거래명세서가 생성되었습니다." & vbCrLf & "인쇄하려면 Ctrl+P를 누르세요."
    End If
    Unload Me
End Sub

Private Sub btn_PDF_Click()
    If Not ValidateInput() Then Exit Sub
    Dim txnID As String
    txnID = SaveTransaction(True)
    If txnID <> "" Then
        Dim hdr As TTxnHeader: hdr = GetHeaderByTxnID(txnID)
        GenerateStatement txnID
        ExportStatementToPDF txnID, hdr.CompanyName, hdr.TxnDate
    End If
    Unload Me
End Sub

Private Function ValidateInput() As Boolean
    ValidateInput = False
    If cbo_Customer.ListIndex < 0 Then ShowError "거래처를 선택하세요.": Exit Function
    If itemCount = 0 Then ShowError "품목을 1개 이상 추가하세요.": Exit Function
    If Trim(txt_TxnDate.Value) = "" Then ShowError "거래일자를 입력하세요.": Exit Function
    ValidateInput = True
End Function

Private Function SaveTransaction(forPrint As Boolean) As String
    Dim txnID As String: txnID = GenerateTxnID()

    Dim supplyTotal As Long: supplyTotal = 0
    Dim vatTotal    As Long: vatTotal    = 0
    Dim i As Integer
    For i = 0 To itemCount - 1
        supplyTotal = supplyTotal + items(i).Amount
        vatTotal    = vatTotal    + items(i).VAT
    Next i
    Dim invTotal   As Long: invTotal   = supplyTotal + vatTotal
    Dim prevBal    As Long: prevBal    = GetPrevBalance(currentCustCode)
    Dim grandTotal As Long: grandTotal = prevBal + invTotal
    Dim payment    As Long
    On Error Resume Next: payment = CLng(txt_Payment.Value): On Error GoTo 0
    Dim todayBal   As Long: todayBal   = grandTotal - payment

    Dim hdr As TTxnHeader
    hdr.TxnID         = txnID
    hdr.TxnDate       = ParseDate(txt_TxnDate.Value)
    hdr.CustCode      = currentCustCode
    hdr.CompanyName   = cbo_Customer.List(cbo_Customer.ListIndex, 0)
    hdr.PrevBalance   = prevBal
    hdr.SupplyTotal   = supplyTotal
    hdr.VATTotal      = vatTotal
    hdr.InvoiceTotal  = invTotal
    hdr.GrandTotal    = grandTotal
    hdr.PaymentToday  = payment
    hdr.TodayBalance  = todayBal
    hdr.PrintOperator = Trim(txt_Operator.Value)
    hdr.Printed       = IIf(forPrint, "Y", "N")
    SaveTransactionHeader hdr

    For i = 0 To itemCount - 1
        Dim dtl As TTxnDetail
        dtl.TxnID          = txnID
        dtl.SeqNo          = i + 1
        dtl.ItemCode       = items(i).ItemCode
        dtl.ItemName       = items(i).ItemName
        dtl.Unit           = items(i).Unit
        dtl.Qty            = items(i).Qty
        dtl.TotalWeight    = items(i).Qty
        dtl.UnitPrice      = items(i).UnitPrice
        dtl.Amount         = items(i).Amount
        dtl.VATApply       = items(i).VATApply
        dtl.VAT            = items(i).VAT
        dtl.TraceNo        = items(i).TraceNo
        dtl.SlaughterHouse = items(i).Slaughter
        SaveTransactionDetail dtl
    Next i

    UpdateCustomerBalance currentCustCode, todayBal
    UpdateDashboard

    SaveTransaction = txnID
End Function

Private Sub btn_Cancel_Click(): Unload Me: End Sub
"""

FORM_CODE["frmTransactionHistory"] = r"""
Option Explicit

Private Sub UserForm_Initialize()
    Me.Caption = "거래 조회 / 재출력"
    txt_DateFrom.Value = Format(DateSerial(Year(Now()), Month(Now()), 1), "YYYY-MM-DD")
    txt_DateTo.Value   = Format(Now(), "YYYY-MM-DD")

    ' 거래처 필터 콤보
    cbo_FilterCust.AddItem "(전체)"
    Dim data As Variant: data = GetCustomerList()
    If Not IsEmpty(data) And VarType(data) <> vbEmpty Then
        Dim i As Integer
        For i = 0 To UBound(data, 1)
            cbo_FilterCust.AddItem data(i, 1)
            cbo_FilterCust.List(cbo_FilterCust.ListCount - 1, 1) = data(i, 0)
        Next i
    End If
    cbo_FilterCust.ListIndex = 0

    lst_Results.ColumnCount = 5
    lst_Results.ColumnWidths = "120;80;70;80;80"
    DoSearch
End Sub

Private Sub btn_Search_Click(): DoSearch: End Sub

Private Sub DoSearch()
    lst_Results.Clear
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SH_TXN_HDR)
    Dim lr As Long: lr = GetLastRow(SH_TXN_HDR, 1)

    Dim filterCode As String
    If cbo_FilterCust.ListIndex > 0 Then
        filterCode = cbo_FilterCust.List(cbo_FilterCust.ListIndex, 1)
    End If

    Dim dFrom As Date, dTo As Date
    On Error Resume Next
    dFrom = ParseDate(txt_DateFrom.Value)
    dTo   = ParseDate(txt_DateTo.Value)
    On Error GoTo 0

    Dim r As Long
    For r = 2 To lr
        Dim txnDate As Date
        On Error Resume Next: txnDate = CDate(ws.Cells(r, 2).Value): On Error GoTo 0
        If txnDate < dFrom Or txnDate > dTo Then GoTo Skip
        If filterCode <> "" And ws.Cells(r, 3).Value <> filterCode Then GoTo Skip

        lst_Results.AddItem ws.Cells(r, 1).Value  ' TxnID
        lst_Results.List(lst_Results.ListCount - 1, 1) = Format(txnDate, "YYYY-MM-DD")
        lst_Results.List(lst_Results.ListCount - 1, 2) = ws.Cells(r, 4).Value
        lst_Results.List(lst_Results.ListCount - 1, 3) = Format(CLng(ws.Cells(r, 8).Value), "#,##0")
        lst_Results.List(lst_Results.ListCount - 1, 4) = ws.Cells(r, 13).Value
Skip:
    Next r
End Sub

Private Sub btn_Reprint_Click()
    If lst_Results.ListIndex < 0 Then ShowError "재출력할 거래를 선택하세요.": Exit Sub
    Dim txnID As String: txnID = lst_Results.List(lst_Results.ListIndex, 0)
    GenerateStatement txnID
    ThisWorkbook.Sheets(SH_STATEMENT).Activate
    ShowInfo "거래명세서 생성 완료. Ctrl+P로 인쇄하세요."
    Unload Me
End Sub

Private Sub btn_PDF_Click()
    If lst_Results.ListIndex < 0 Then ShowError "PDF로 저장할 거래를 선택하세요.": Exit Sub
    Dim txnID As String: txnID = lst_Results.List(lst_Results.ListIndex, 0)
    Dim hdr As TTxnHeader: hdr = GetHeaderByTxnID(txnID)
    GenerateStatement txnID
    ExportStatementToPDF txnID, hdr.CompanyName, hdr.TxnDate
End Sub

Private Sub btn_Close_Click(): Unload Me: End Sub
"""

THISWORKBOOK_CODE = r"""
Option Explicit

Private Sub Workbook_Open()
    Application.UseSystemSeparators = True
    ThisWorkbook.Sheets(SH_DASHBOARD).Activate
    Call UpdateDashboard
End Sub
"""

# ══════════════════════════════════════════════════════════════════════════════
# VBA 코드를 .bas 파일로 저장 (COM 접근 불가 시 폴백)
# ══════════════════════════════════════════════════════════════════════════════
def _save_vba_as_files():
    vba_dir = os.path.join(os.path.dirname(OUTPUT_PATH), "vba_modules")
    os.makedirs(vba_dir, exist_ok=True)

    for name, code in VBA_MODULES.items():
        path = os.path.join(vba_dir, f"{name}.bas")
        with open(path, "w", encoding="utf-8-sig") as f:
            f.write(f"Attribute VB_Name = \"{name}\"\r\n")
            f.write(code)
        print(f"  저장: {path}")

    for name, code in FORM_CODE.items():
        path = os.path.join(vba_dir, f"{name}.bas")
        with open(path, "w", encoding="utf-8-sig") as f:
            f.write(f"Attribute VB_Name = \"{name}\"\r\n")
            f.write(code)
        print(f"  저장: {path}")

    # ThisWorkbook 코드
    path = os.path.join(vba_dir, "ThisWorkbook.bas")
    with open(path, "w", encoding="utf-8-sig") as f:
        f.write(THISWORKBOOK_CODE)
    print(f"  저장: {path}")

    # 임포트 안내 README
    readme = os.path.join(vba_dir, "IMPORT_GUIDE.txt")
    with open(readme, "w", encoding="utf-8-sig") as f:
        f.write("""VBA 코드 임포트 안내
====================

1. 거래관리.xlsm 파일을 Excel로 엽니다.
2. Alt+F11 을 눌러 VBA 편집기를 엽니다.
3. 아래 모듈들을 순서대로 임포트합니다:
   - 파일 → 파일 가져오기 → 각 .bas 파일 선택

임포트 순서:
  1. modConstants.bas
  2. modUtils.bas
  3. modDBHelper.bas
  4. modStatement.bas
  5. modDashboard.bas
  6. modButtons.bas

UserForm 임포트 (순서 무관):
  - frmProductMgmt.bas   → UserForm으로 추가 후 코드만 붙여넣기
  - frmCustomerMgmt.bas  → 동일
  - frmNewTransaction.bas → 동일
  - frmTransactionHistory.bas → 동일

UserForm 추가 방법:
  VBA 편집기 → 삽입 → 사용자 정의 폼
  폼 이름을 각각 frmProductMgmt 등으로 변경
  폼 더블클릭 → 코드 붙여넣기

4. ThisWorkbook 코드:
   프로젝트 탐색기에서 'ThisWorkbook' 더블클릭
   IMPORT_GUIDE.txt 하단의 코드 붙여넣기

5. 저장 후 Excel 재시작
""")
    print(f"\n[안내] VBA 임포트 가이드: {readme}")


# ══════════════════════════════════════════════════════════════════════════════
# 메인 빌더
# ══════════════════════════════════════════════════════════════════════════════
def build_workbook():
    import time
    pythoncom.CoInitialize()
    xl = win32.DispatchEx("Excel.Application")  # 새 인스턴스 강제
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.AutomationSecurity = 1   # msoAutomationSecurityLow (1=low, 2=byUI, 3=high)

    wb = xl.Workbooks.Add()

    # 기존 시트 이름 확보 (Sheet1 등)
    while wb.Sheets.Count < 6:
        wb.Sheets.Add(After=wb.Sheets(wb.Sheets.Count))

    print("시트 생성 중...")
    setup_대시보드(wb.Sheets(1))
    setup_품목DB(wb.Sheets(2))
    setup_거래처DB(wb.Sheets(3))
    setup_거래헤더DB(wb.Sheets(4))
    setup_거래상세DB(wb.Sheets(5))
    setup_거래명세서(xl, wb.Sheets(6))

    # 먼저 xlsm으로 저장 후 Excel 재시작하여 VBProject 접근
    print("저장 후 Excel 재시작...")
    wb.SaveAs(Filename=OUTPUT_PATH, FileFormat=52)
    wb.Close(SaveChanges=False)
    xl.Quit()
    time.sleep(3)

    # 완전히 새 Excel 인스턴스 시작
    xl = win32.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.AutomationSecurity = 1
    time.sleep(2)

    wb = xl.Workbooks.Open(OUTPUT_PATH)
    time.sleep(2)

    print("VBA 모듈 삽입 중...")
    try:
        vba = wb.VBProject
    except Exception as vba_err:
        # VBProject 접근 실패 시 .bas 파일로 대체 저장
        wb.Close(SaveChanges=False)
        xl.Quit()
        _save_vba_as_files()
        raise RuntimeError(
            "VBProject COM 접근 불가 - VBA 코드를 .bas 파일로 저장했습니다.\n"
            "Excel에서 Alt+F11 → 파일 → 파일 가져오기로 각 .bas 파일을 임포트하세요.\n"
            f"위치: {os.path.dirname(OUTPUT_PATH)}\\vba_modules\\"
        ) from vba_err

    for mod_name, code in VBA_MODULES.items():
        mod = vba.VBComponents.Add(1)  # vbext_ct_StdModule = 1
        mod.Name = mod_name
        mod.CodeModule.AddFromString(code)
        print(f"  모듈 추가: {mod_name}")

    print("UserForm 삽입 중...")
    for form_name, code in FORM_CODE.items():
        frm = vba.VBComponents.Add(3)  # vbext_ct_MSForm = 3
        frm.Name = form_name
        frm.CodeModule.AddFromString(code)
        print(f"  폼 추가: {form_name}")

    # ThisWorkbook 코드
    tb = vba.VBComponents("ThisWorkbook")
    tb.CodeModule.AddFromString(THISWORKBOOK_CODE)

    print("대시보드 버튼 추가 중...")
    wsDash = wb.Sheets("대시보드")
    btn_data = [
        ("거래 입력", "OpenNewTransaction",  2, 3, 120, 35),
        ("품목 관리", "OpenProductMgmt",     2, 4, 120, 35),
        ("거래처 관리","OpenCustomerMgmt",   2, 5, 120, 35),
        ("거래 조회", "OpenHistory",         2, 6, 120, 35),
        ("대시보드 새로고침","RefreshDashboard", 2, 7, 150, 35),
    ]
    for label, macro, col_start, row_num, w, h in btn_data:
        left_pos = wsDash.Cells(row_num, col_start).Left
        top_pos  = wsDash.Cells(row_num, col_start).Top
        btn = wsDash.Buttons().Add(left_pos, top_pos, w, h)
        btn.Caption = label
        btn.OnAction = macro
        btn.Font.Size = 11
        btn.Font.Bold = True

    # 월별 차트
    print("차트 생성 중...")
    wsDash.Activate()
    cht1 = wsDash.ChartObjects().Add(Left=wsDash.Cells(8,1).Left,
                                      Top=wsDash.Cells(30,1).Top,
                                      Width=400, Height=250)
    cht1.Name = "cht_Monthly"
    cht1.Chart.ChartType = 57  # xlColumnClustered
    cht1.Chart.HasTitle = True
    cht1.Chart.ChartTitle.Text = "월별 합계금액"

    # 거래처 파이 차트
    cht2 = wsDash.ChartObjects().Add(Left=wsDash.Cells(8,8).Left,
                                      Top=wsDash.Cells(30,1).Top,
                                      Width=350, Height=250)
    cht2.Name = "cht_Customer"
    cht2.Chart.ChartType = 5  # xlPie
    cht2.Chart.HasTitle = True
    cht2.Chart.ChartTitle.Text = "거래처별 공급가액 (이번달)"

    # 최종 저장
    print(f"저장 중: {OUTPUT_PATH}")
    wb.Save()
    print("완료!")

    xl.Visible = True
    xl.DisplayAlerts = True
    return wb, xl


def enable_vba_access():
    """Excel 시작 전에 레지스트리에 VBA 접근 권한 설정"""
    import winreg
    # Excel 버전 자동 감지 (16.0=2016/2019/365, 15.0=2013, 14.0=2010)
    for ver in ["16.0", "15.0", "14.0"]:
        key_path = rf"Software\Microsoft\Office\{ver}\Excel\Security"
        try:
            key = winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_SET_VALUE)
            winreg.SetValueEx(key, "AccessVBOM", 0, winreg.REG_DWORD, 1)
            winreg.CloseKey(key)
            print(f"  VBA 접근 권한 활성화 (Office {ver})")
        except Exception:
            pass


if __name__ == "__main__":
    import sys
    # stdout encoding 설정
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

    print("거래관리.xlsm 생성 시작...")
    print("레지스트리 VBA 접근 권한 설정 중...")
    enable_vba_access()

    # 기존 Excel 프로세스 종료
    import subprocess, time
    subprocess.run(["taskkill", "/F", "/IM", "EXCEL.EXE"], capture_output=True)
    time.sleep(2)

    try:
        wb, xl = build_workbook()
        print(f"\n[완료] 생성: {OUTPUT_PATH}")
        print("Excel 파일이 열려 있습니다. 확인 후 닫아주세요.")
    except Exception as e:
        print(f"\n[오류] {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

Attribute VB_Name = "modStatement"

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

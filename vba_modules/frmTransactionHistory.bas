Attribute VB_Name = "frmTransactionHistory"

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

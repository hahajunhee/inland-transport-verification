Attribute VB_Name = "frmNewTransaction"

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

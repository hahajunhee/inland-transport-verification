Attribute VB_Name = "modDBHelper"

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

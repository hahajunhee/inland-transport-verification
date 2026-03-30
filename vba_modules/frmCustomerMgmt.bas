Attribute VB_Name = "frmCustomerMgmt"

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

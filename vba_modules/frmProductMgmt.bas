Attribute VB_Name = "frmProductMgmt"

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

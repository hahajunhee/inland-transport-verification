Attribute VB_Name = "modBuildForms"
Option Explicit
' ======================================================================
' UserForm 컨트롤 레이아웃 생성 모듈
' InstallAll 실행 후 BuildAllForms 를 실행하세요
' ======================================================================

Public Sub BuildAllForms()
    BuildFormProduct
    BuildFormCustomer
    BuildFormNewTransaction
    BuildFormHistory
    MsgBox "폼 레이아웃 생성 완료!", vbInformation
End Sub

' ── 공통 헬퍼 ────────────────────────────────────────────────────────────────
Private Function AddLabel(frm As Object, cap As String, l As Single, t As Single, w As Single, h As Single) As Object
    Dim c As Object
    Set c = frm.Designer.Controls.Add("Forms.Label.1")
    c.Caption = cap: c.Left = l: c.Top = t: c.Width = w: c.Height = h
    c.Font.Size = 9
    Set AddLabel = c
End Function

Private Function AddTextBox(frm As Object, nm As String, l As Single, t As Single, w As Single, h As Single) As Object
    Dim c As Object
    Set c = frm.Designer.Controls.Add("Forms.TextBox.1")
    c.Name = nm: c.Left = l: c.Top = t: c.Width = w: c.Height = h
    c.Font.Size = 9
    Set AddTextBox = c
End Function

Private Function AddComboBox(frm As Object, nm As String, l As Single, t As Single, w As Single, h As Single) As Object
    Dim c As Object
    Set c = frm.Designer.Controls.Add("Forms.ComboBox.1")
    c.Name = nm: c.Left = l: c.Top = t: c.Width = w: c.Height = h
    c.Font.Size = 9
    Set AddComboBox = c
End Function

Private Function AddListBox(frm As Object, nm As String, l As Single, t As Single, w As Single, h As Single) As Object
    Dim c As Object
    Set c = frm.Designer.Controls.Add("Forms.ListBox.1")
    c.Name = nm: c.Left = l: c.Top = t: c.Width = w: c.Height = h
    c.Font.Size = 9
    Set AddListBox = c
End Function

Private Function AddButton(frm As Object, nm As String, cap As String, l As Single, t As Single, w As Single, h As Single) As Object
    Dim c As Object
    Set c = frm.Designer.Controls.Add("Forms.CommandButton.1")
    c.Name = nm: c.Caption = cap: c.Left = l: c.Top = t: c.Width = w: c.Height = h
    c.Font.Size = 9
    Set AddButton = c
End Function

Private Function AddOptionBtn(frm As Object, nm As String, cap As String, l As Single, t As Single, w As Single, h As Single) As Object
    Dim c As Object
    Set c = frm.Designer.Controls.Add("Forms.OptionButton.1")
    c.Name = nm: c.Caption = cap: c.Left = l: c.Top = t: c.Width = w: c.Height = h
    c.Font.Size = 9
    Set AddOptionBtn = c
End Function

' ══════════════════════════════════════════════════════════════════════════════
' 품목 관리 폼
' ══════════════════════════════════════════════════════════════════════════════
Private Sub BuildFormProduct()
    Dim frm As Object
    Set frm = ThisWorkbook.VBProject.VBComponents("frmProductMgmt")
    frm.Properties("Caption") = "품목 관리"
    frm.Properties("Width")   = 500
    frm.Properties("Height")  = 400

    ' 목록
    AddLabel frm, "[ 등록된 품목 ]", 6, 6, 200, 14
    Dim lst As Object: Set lst = AddListBox(frm, "lst_Products", 6, 22, 480, 120)
    lst.ColumnHeads = False

    ' 입력 레이블 + 텍스트박스
    AddLabel frm, "품목명 *", 6, 150, 50, 14
    AddTextBox frm, "txt_ItemName", 60, 148, 160, 18

    AddLabel frm, "규격", 6, 172, 50, 14
    AddTextBox frm, "txt_Spec", 60, 170, 100, 18

    AddLabel frm, "단위 *", 175, 150, 40, 14
    AddComboBox frm, "cbo_Unit", 218, 148, 70, 18

    AddLabel frm, "단가 *", 6, 194, 50, 14
    AddTextBox frm, "txt_UnitPrice", 60, 192, 100, 18

    AddLabel frm, "마진율(%)", 175, 172, 60, 14
    AddTextBox frm, "txt_MarginRate", 240, 170, 50, 18

    AddLabel frm, "부가세", 175, 194, 40, 14
    AddOptionBtn frm, "opt_VAT_Y", "과세", 218, 192, 50, 14
    AddOptionBtn frm, "opt_VAT_N", "면세", 272, 192, 50, 14

    AddLabel frm, "비고", 6, 216, 50, 14
    AddTextBox frm, "txt_Remark", 60, 214, 200, 18

    ' 모드 표시
    AddLabel frm, "[신규 입력]", 6, 238, 280, 14
    Dim lm As Object: Set lm = frm.Designer.Controls("lbl_Mode")
    ' lbl_Mode 이름 재설정 필요 - 마지막 추가된 레이블에 이름 설정
    frm.Designer.Controls(frm.Designer.Controls.Count - 1).Name = "lbl_Mode"

    ' 버튼
    AddButton frm, "btn_New",    "신규",   6, 258, 60, 22
    AddButton frm, "btn_Save",   "저장",  72, 258, 60, 22
    AddButton frm, "btn_Delete", "삭제", 138, 258, 60, 22
    AddButton frm, "btn_Close",  "닫기", 420, 258, 60, 22
End Sub

' ══════════════════════════════════════════════════════════════════════════════
' 거래처 관리 폼
' ══════════════════════════════════════════════════════════════════════════════
Private Sub BuildFormCustomer()
    Dim frm As Object
    Set frm = ThisWorkbook.VBProject.VBComponents("frmCustomerMgmt")
    frm.Properties("Caption") = "거래처 관리"
    frm.Properties("Width")   = 500
    frm.Properties("Height")  = 430

    AddLabel frm, "[ 등록된 거래처 ]", 6, 6, 200, 14
    AddListBox frm, "lst_Customers", 6, 22, 480, 110

    Dim fields As Variant
    fields = Array( _
        Array("상호 *", "txt_CompanyName", 60, 140),  _
        Array("성명", "txt_ContactName", 60, 160),    _
        Array("등록번호", "txt_RegNumber", 60, 180),  _
        Array("업태", "txt_BizType", 60, 200),        _
        Array("업종", "txt_BizCat", 60, 220),         _
        Array("TEL", "txt_Tel", 60, 240),             _
        Array("FAX", "txt_Fax", 60, 260),             _
        Array("개시잔액", "txt_OpenBalance", 60, 280) _
    )

    Dim i As Integer
    For i = 0 To UBound(fields)
        AddLabel frm, CStr(fields(i)(0)), 6, CSng(fields(i)(3)), 50, 14
        AddTextBox frm, CStr(fields(i)(1)), CSng(fields(i)(2)), CSng(fields(i)(3)) - 2, 180, 18
    Next i

    AddLabel frm, "주소", 6, 300, 50, 14
    AddTextBox frm, "txt_Address", 60, 298, 400, 18

    AddLabel frm, "[신규 입력]", 6, 322, 280, 14
    frm.Designer.Controls(frm.Designer.Controls.Count - 1).Name = "lbl_Mode"

    AddButton frm, "btn_New",    "신규",  6, 344, 60, 22
    AddButton frm, "btn_Save",   "저장", 72, 344, 60, 22
    AddButton frm, "btn_Delete", "삭제", 138, 344, 60, 22
    AddButton frm, "btn_Close",  "닫기", 420, 344, 60, 22
End Sub

' ══════════════════════════════════════════════════════════════════════════════
' 신규 거래 입력 폼
' ══════════════════════════════════════════════════════════════════════════════
Private Sub BuildFormNewTransaction()
    Dim frm As Object
    Set frm = ThisWorkbook.VBProject.VBComponents("frmNewTransaction")
    frm.Properties("Caption") = "거래 입력"
    frm.Properties("Width")   = 660
    frm.Properties("Height")  = 530

    ' 헤더 행
    AddLabel frm, "거래일자 *", 6, 8, 55, 14
    AddTextBox frm, "txt_TxnDate", 64, 6, 90, 18

    AddLabel frm, "거래처 *", 160, 8, 45, 14
    AddComboBox frm, "cbo_Customer", 210, 6, 170, 18

    AddLabel frm, "전일잔액 :", 390, 8, 60, 14
    AddLabel frm, "0 원", 454, 8, 100, 14
    frm.Designer.Controls(frm.Designer.Controls.Count - 1).Name = "lbl_PrevBal"

    AddLabel frm, "출력담당자", 6, 30, 60, 14
    AddTextBox frm, "txt_Operator", 70, 28, 100, 18

    ' 구분선 (Frame)
    Dim fr As Object
    Set fr = frm.Designer.Controls.Add("Forms.Frame.1")
    fr.Caption = "품목 추가": fr.Left = 6: fr.Top = 52: fr.Width = 640: fr.Height = 66

    AddLabel frm, "품목", 12, 68, 30, 14
    AddComboBox frm, "cbo_Item", 46, 66, 160, 18

    AddLabel frm, "단위:", 214, 68, 28, 14
    AddLabel frm, "-", 242, 68, 50, 14
    frm.Designer.Controls(frm.Designer.Controls.Count - 1).Name = "lbl_Unit"

    AddLabel frm, "단가:", 298, 68, 28, 14
    AddLabel frm, "0 원", 326, 68, 80, 14
    frm.Designer.Controls(frm.Designer.Controls.Count - 1).Name = "lbl_UnitPrice"

    AddLabel frm, "수량 *", 12, 90, 32, 14
    AddTextBox frm, "txt_Qty", 48, 88, 60, 18

    AddLabel frm, "이력번호", 116, 90, 50, 14
    AddTextBox frm, "txt_TraceNo", 170, 88, 100, 18

    AddLabel frm, "도축장", 278, 90, 40, 14
    AddTextBox frm, "txt_Slaughter", 322, 88, 100, 18

    AddButton frm, "btn_AddItem", "추가 ▶", 434, 72, 60, 22
    AddButton frm, "btn_DelItem", "◀ 삭제", 500, 72, 60, 22

    ' 미리보기
    AddLabel frm, "금액 미리보기:", 568, 72, 76, 14
    AddLabel frm, "", 568, 86, 80, 14
    frm.Designer.Controls(frm.Designer.Controls.Count - 1).Name = "lbl_AmtPreview"

    ' 품목 목록
    AddLabel frm, "[ 입력된 품목 ]", 6, 122, 120, 14
    AddListBox frm, "lst_Items", 6, 136, 640, 150

    ' 합계 행
    AddLabel frm, "공급가액합 :", 6, 292, 72, 14
    AddLabel frm, "0 원", 82, 292, 80, 14
    frm.Designer.Controls(frm.Designer.Controls.Count - 1).Name = "lbl_SupplyTotal"

    AddLabel frm, "부가세합 :", 170, 292, 58, 14
    AddLabel frm, "0 원", 232, 292, 70, 14
    frm.Designer.Controls(frm.Designer.Controls.Count - 1).Name = "lbl_VATTotal"

    AddLabel frm, "합계금액 :", 310, 292, 58, 14
    AddLabel frm, "0 원", 372, 292, 80, 14
    frm.Designer.Controls(frm.Designer.Controls.Count - 1).Name = "lbl_InvTotal"

    AddLabel frm, "총액 :", 460, 292, 36, 14
    AddLabel frm, "0 원", 500, 292, 100, 14
    frm.Designer.Controls(frm.Designer.Controls.Count - 1).Name = "lbl_Grand"

    ' 입금 행
    AddLabel frm, "금일입금액 *", 6, 316, 72, 14
    AddTextBox frm, "txt_Payment", 82, 314, 80, 18

    AddLabel frm, "금일잔액 :", 170, 316, 58, 14
    AddLabel frm, "0 원", 232, 316, 100, 14
    frm.Designer.Controls(frm.Designer.Controls.Count - 1).Name = "lbl_TodayBal"

    ' 버튼
    AddButton frm, "btn_Save",   "저장",       6, 344, 80, 26
    AddButton frm, "btn_Print",  "저장 및 출력", 92, 344, 100, 26
    AddButton frm, "btn_PDF",    "PDF 저장",   198, 344, 80, 26
    AddButton frm, "btn_Cancel", "취소",       570, 344, 70, 26
End Sub

' ══════════════════════════════════════════════════════════════════════════════
' 거래 조회 폼
' ══════════════════════════════════════════════════════════════════════════════
Private Sub BuildFormHistory()
    Dim frm As Object
    Set frm = ThisWorkbook.VBProject.VBComponents("frmTransactionHistory")
    frm.Properties("Caption") = "거래 조회 / 재출력"
    frm.Properties("Width")   = 580
    frm.Properties("Height")  = 450

    AddLabel frm, "거래처 :", 6, 8, 45, 14
    AddComboBox frm, "cbo_FilterCust", 56, 6, 160, 18

    AddLabel frm, "기간 :", 224, 8, 36, 14
    AddTextBox frm, "txt_DateFrom", 264, 6, 90, 18
    AddLabel frm, "~", 358, 8, 12, 14
    AddTextBox frm, "txt_DateTo", 374, 6, 90, 18

    AddButton frm, "btn_Search", "조회", 470, 4, 60, 22

    AddLabel frm, "[ 거래 내역 ] ※ 클릭 후 재출력/PDF 버튼 사용", 6, 30, 360, 14
    AddListBox frm, "lst_Results", 6, 46, 560, 300
    frm.Designer.Controls("lst_Results").ColumnHeads = False

    AddButton frm, "btn_Reprint", "재출력",   6, 354, 80, 26
    AddButton frm, "btn_PDF",     "PDF 저장", 92, 354, 80, 26
    AddButton frm, "btn_Close",   "닫기",    490, 354, 70, 26
End Sub

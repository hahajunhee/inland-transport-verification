Attribute VB_Name = "INSTALLER"
' ======================================================================
' 거래관리 VBA 자동설치 매크로
' ----------------------------------------------------------------------
' 사용법:
'   1. 거래관리.xlsm 을 Excel로 엽니다 (매크로 활성화)
'   2. Alt+F11 → VBA 편집기를 엽니다
'   3. 삽입 → 모듈 을 클릭합니다
'   4. 이 파일의 내용 전체를 붙여넣습니다 (Ctrl+A → Ctrl+C 후 VBE에서 Ctrl+V)
'   5. F5 (또는 실행 → 매크로 실행 → InstallAll) 을 클릭합니다
'   6. 설치 완료 후 이 모듈(INSTALLER)은 자동 삭제됩니다
' ======================================================================
Option Explicit

Sub InstallAll()
    Dim vba As Object
    Set vba = ThisWorkbook.VBProject

    Dim basDir As String
    basDir = ThisWorkbook.Path & "\vba_modules\"

    If Dir(basDir, vbDirectory) = "" Then
        MsgBox "vba_modules 폴더를 찾을 수 없습니다." & vbCrLf & _
               "거래관리.xlsm 과 같은 폴더에 vba_modules 폴더가 있어야 합니다.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False

    ' ── 기존 모듈 삭제 ────────────────────────────────────────────────
    Dim comp As Object
    Dim toDelete As Collection
    Set toDelete = New Collection
    For Each comp In vba.VBComponents
        Select Case comp.Name
            Case "modConstants", "modUtils", "modDBHelper", "modStatement", _
                 "modDashboard", "modButtons", _
                 "frmProductMgmt", "frmCustomerMgmt", _
                 "frmNewTransaction", "frmTransactionHistory"
                toDelete.Add comp
        End Select
    Next comp
    Dim c As Object
    For Each c In toDelete
        vba.VBComponents.Remove c
    Next c

    ' ── 표준 모듈 임포트 ──────────────────────────────────────────────
    Dim mods As Variant
    mods = Array("modConstants", "modUtils", "modDBHelper", "modStatement", "modDashboard", "modButtons", "modBuildForms")
    Dim m As Variant
    For Each m In mods
        Dim fPath As String
        fPath = basDir & m & ".bas"
        If Dir(fPath) <> "" Then
            vba.VBComponents.Import fPath
        Else
            MsgBox "파일 없음: " & fPath, vbExclamation
        End If
    Next m

    ' ── UserForm 생성 + 코드 삽입 ─────────────────────────────────────
    Dim forms As Variant
    forms = Array("frmProductMgmt", "frmCustomerMgmt", "frmNewTransaction", "frmTransactionHistory")
    Dim frm As Variant
    For Each frm In forms
        fPath = basDir & frm & ".bas"
        If Dir(fPath) <> "" Then
            ' UserForm 컴포넌트 추가
            Dim newForm As Object
            Set newForm = vba.VBComponents.Add(3) ' vbext_ct_MSForm
            newForm.Name = CStr(frm)

            ' 코드 읽어서 삽입
            Dim fNum As Integer
            fNum = FreeFile
            Open fPath For Input As #fNum
            Dim allCode As String
            allCode = ""
            Dim lineStr As String
            Dim firstLine As Boolean: firstLine = True
            Do While Not EOF(fNum)
                Line Input #fNum, lineStr
                ' Attribute VB_Name 줄 건너뜀
                If firstLine And InStr(lineStr, "Attribute VB_Name") > 0 Then
                    firstLine = False
                Else
                    allCode = allCode & lineStr & vbCrLf
                    firstLine = False
                End If
            Loop
            Close #fNum
            newForm.CodeModule.AddFromString allCode
        End If
    Next frm

    ' ── ThisWorkbook 코드 삽입 ────────────────────────────────────────
    fPath = basDir & "ThisWorkbook.bas"
    If Dir(fPath) <> "" Then
        Dim tbCode As String
        Dim fNum2 As Integer: fNum2 = FreeFile
        Open fPath For Input As #fNum2
        tbCode = ""
        Do While Not EOF(fNum2)
            Line Input #fNum2, lineStr
            tbCode = tbCode & lineStr & vbCrLf
        Loop
        Close #fNum2
        ' ThisWorkbook 기존 코드에 추가
        Dim tbComp As Object
        Set tbComp = vba.VBComponents("ThisWorkbook")
        Dim existingLines As Long
        existingLines = tbComp.CodeModule.CountOfLines
        If existingLines > 0 Then
            tbComp.CodeModule.DeleteLines 1, existingLines
        End If
        tbComp.CodeModule.AddFromString tbCode
    End If

    ' ── UserForm 레이아웃 생성 ───────────────────────────────────────
    Application.Run "BuildAllForms"

    ' ── 이 설치 모듈 자체 삭제 ────────────────────────────────────────
    On Error Resume Next
    vba.VBComponents.Remove vba.VBComponents("INSTALLER")
    On Error GoTo 0

    Application.ScreenUpdating = True
    ThisWorkbook.Save

    MsgBox "설치 완료!" & vbCrLf & vbCrLf & _
           "거래관리 시스템이 설치되었습니다." & vbCrLf & _
           "파일을 저장했습니다." & vbCrLf & vbCrLf & _
           "Excel을 재시작하면 대시보드가 표시됩니다.", vbInformation, "설치 완료"
End Sub

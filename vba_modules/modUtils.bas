Attribute VB_Name = "modUtils"

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

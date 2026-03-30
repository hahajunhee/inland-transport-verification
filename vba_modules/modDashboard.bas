Attribute VB_Name = "modDashboard"

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


Option Explicit

Private Sub Workbook_Open()
    Application.UseSystemSeparators = True
    ThisWorkbook.Sheets(SH_DASHBOARD).Activate
    Call UpdateDashboard
End Sub

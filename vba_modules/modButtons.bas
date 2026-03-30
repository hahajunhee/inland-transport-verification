Attribute VB_Name = "modButtons"

Option Explicit

Public Sub OpenNewTransaction()
    frmNewTransaction.Show
End Sub

Public Sub OpenProductMgmt()
    frmProductMgmt.Show
End Sub

Public Sub OpenCustomerMgmt()
    frmCustomerMgmt.Show
End Sub

Public Sub OpenHistory()
    frmTransactionHistory.Show
End Sub

Public Sub RefreshDashboard()
    Call UpdateDashboard
    ShowInfo "대시보드가 업데이트되었습니다."
End Sub

VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LongReportForm 
   Caption         =   "產生長表"
   ClientHeight    =   4755
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   3825
   OleObjectBlob   =   "LongReportForm.frx":0000
   StartUpPosition =   1  '所屬視窗中央
End
Attribute VB_Name = "LongReportForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmdGetDayReport_Click()

Call cmdGetLongReport

Unload Me

End Sub

Private Sub UserForm_Initialize()

With Sheets("日報資料庫")

    lr = .Cells(.Rows.Count, 2).End(xlUp).Row
    
    s_date = .Range("B2")

    If s_date = "" Then MsgBox "目前沒有資料!", vbCritical: End

    Me.tbosDate = s_date
    
    e_date = .Range("B" & lr)
    Me.tboeDate = e_date

End With

End Sub


VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DayReportForm 
   Caption         =   "產生日報"
   ClientHeight    =   5490
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   4335
   OleObjectBlob   =   "DayReportForm.frx":0000
   StartUpPosition =   1  '所屬視窗中央
End
Attribute VB_Name = "DayReportForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdGetDayReport_Click()

Call cmdButton.cmdGetDayReport

Unload Me

End Sub

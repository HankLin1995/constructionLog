VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DayReportForm 
   Caption         =   "���ͤ��"
   ClientHeight    =   5490
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   4335
   OleObjectBlob   =   "DayReportForm.frx":0000
   StartUpPosition =   1  '���ݵ�������
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

Private Sub UserForm_Initialize()

With Sheets("�����Ʈw")

    lr = .Cells(.Rows.Count, 2).End(xlUp).Row
    
    s_date = .Range("B2")

    If s_date = "" Then MsgBox "�ثe�S�����!", vbCritical: End

    Me.tbosDate = s_date
    
    e_date = .Range("B" & lr)
    Me.tboeDate = e_date

End With

End Sub

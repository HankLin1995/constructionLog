VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MLEForm 
   Caption         =   "工料設定"
   ClientHeight    =   4605
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   4440
   OleObjectBlob   =   "MLEForm.frx":0000
   StartUpPosition =   1  '所屬視窗中央
End
Attribute VB_Name = "MLEForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Private Sub cmdAdd_Click()

Dim MLEobj As New clsMLE

Call MLEobj.ReadData
Call MLEobj.sortData

Call MLEobj.setValidation_MLE

End Sub



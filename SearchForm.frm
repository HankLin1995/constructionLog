VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SearchForm 
   Caption         =   "報表編號查詢器"
   ClientHeight    =   1005
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   4635
   OleObjectBlob   =   "SearchForm.frx":0000
   StartUpPosition =   1  '所屬視窗中央
End
Attribute VB_Name = "SearchForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False








Private Sub CommandButton1_Click()

Application.ScreenUpdating = False

recCode = Me.ComboBox1.Value

Call cmdGetDataByCode(recCode)

Application.ScreenUpdating = True

Unload Me

End Sub

Private Sub UserForm_Initialize()

Dim obj As New clsRecord

Set coll = obj.getSameCode

For Each it In coll

Me.ComboBox1.AddItem it

Next

End Sub

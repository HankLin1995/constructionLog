VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TmpForm 
   Caption         =   "TmpForm"
   ClientHeight    =   1200
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   4590
   OleObjectBlob   =   "TmpForm.frx":0000
   StartUpPosition =   1  '©ÒÄÝµøµ¡¤¤¥¡
End
Attribute VB_Name = "TmpForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Private Sub CommandButton1_Click()

tmp_name = Me.ComboBox1.Value
tmp_type = Me.Label1.Caption

Call cmdGetDataByTmpName(tmp_type, tmp_name)

'Call clearData

'Call getDatabyTmp(tmp_name)

'Call setValidation

Unload Me

End Sub


Private Sub UserForm_Initialize()

Dim o As New clsTmp

Set coll = o.getCboItems '(Me.Label1.Caption)

For Each it In coll

    ComboBox1.AddItem it

Next

End Sub

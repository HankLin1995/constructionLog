VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ERRORForm 
   Caption         =   "系統回報"
   ClientHeight    =   4650
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   4455
   OleObjectBlob   =   "ERRORForm.frx":0000
   StartUpPosition =   1  '所屬視窗中央
End
Attribute VB_Name = "ERRORForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub cmdSubmit_Click()

'=====SignDetail==========

Dim o As New clsFetchURL

user_name = Me.tboName.Text
user_company = Me.tboJob.Text
user_mail = Me.tboMail.Text
msg = Me.tboMSG.Text

Dim o2 As New clsUserInformation
Mac = o2.getMacAddress

myURL_GAS = o.CreateURL("ERRORMSG", Mac, user_name, user_company, user_mail, msg)
o.ExecHTTP (myURL_GAS)

MsgBox "已發送成功，等候通知!", vbInformation

Unload Me
    
End Sub


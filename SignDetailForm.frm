VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SignDetailForm 
   Caption         =   "���v�|�����"
   ClientHeight    =   4185
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   3915
   OleObjectBlob   =   "SignDetailForm.frx":0000
   StartUpPosition =   1  '���ݵ�������
End
Attribute VB_Name = "SignDetailForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False









'Const URL As String = "https://hankecpay.000webhostapp.com"
Const URL As String = "https://php.hanksvba.com"

Private Sub cmdSubmit_Click()

'=====SignDetail==========

Dim o As New clsFetchURL

user_name = Me.tboName.Text
user_company = Me.tboJob.Text
user_mail = Me.tboMail.Text

myURL_GAS = o.CreateURL("SignDetail", user_name, user_company, user_mail)
o.ExecHTTP (myURL_GAS)

'=====GOTO ECPAY==========

Dim ui As New clsUserInformation

mac_add = ui.getMacAddress

myURL = URL & "?email=" & mac_add
ActiveWorkbook.FollowHyperlink myURL
MsgBox "�[�ȧ�����Э��s���}���n��~"
'ThisWorkbook.Close SaveChanges:=True

Unload Me
    
End Sub

Private Sub UserForm_Initialize()

Dim o As New clsUserInformation

tboMac = o.getMacAddress

End Sub

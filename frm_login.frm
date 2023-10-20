VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_login 
   Caption         =   "Login"
   ClientHeight    =   5310
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   4875
   OleObjectBlob   =   "frm_login.frx":0000
   StartUpPosition =   1  '所屬視窗中央
End
Attribute VB_Name = "frm_login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Private Sub CommandButton1_Click() '登入

Dim obj As New clsFetchURL

id = Me.tboID
password = Me.tboPASSWORD

Call obj.checkAccesByID(id, password)

Unload Me

End Sub

Private Sub CommandButton2_Click() '註冊

frm_signup.Show

End Sub

Private Sub CommandButton3_Click() '忘記密碼

Dim obj As New clsFetchURL

If Me.tboID = "" Then MsgBox "帳號不得為空值!!!", vbCritical: Exit Sub

obj.getPassword (Me.tboID)

End Sub

Private Sub UserForm_Terminate()

If Me.Label6 <> "Pass" Then ThisWorkbook.Close savechanges:=False

End Sub

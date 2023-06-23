VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_signup 
   Caption         =   "註冊工具"
   ClientHeight    =   7680
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11670
   OleObjectBlob   =   "frm_signup.frx":0000
   StartUpPosition =   1  '所屬視窗中央
End
Attribute VB_Name = "frm_signup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub CommandButton1_Click()

If Not tboID.Text Like "*@*" Then

    MsgBox "Mail格式很奇怪!", vbCritical
    Exit Sub
    
End If

If Len(tboCostNum.Text) <> 5 Then

    MsgBox "請輸入帳號末五碼!", vbCritical
    Exit Sub

End If

If Len(tboWG.Text) = 0 Then

    MsgBox "請完整填寫機關名稱!", vbCritical
    Exit Sub
    
End If

If Len(tboName.Text) = 0 Then

    MsgBox "請填寫註冊人的姓名!", vbCritical
    Exit Sub
    
End If

Dim obj As New clsFetchURL
Call obj.signup(tboID.Text, tboCostNum.Text, tboWG.Text, tboName.Text)

Unload frm_signup

End Sub



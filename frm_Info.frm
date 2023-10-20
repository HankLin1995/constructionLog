VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Info 
   Caption         =   "ConstructionLog"
   ClientHeight    =   6540
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11145
   OleObjectBlob   =   "frm_Info.frx":0000
   StartUpPosition =   1  '©ÒÄÝµøµ¡¤¤¥¡
End
Attribute VB_Name = "frm_Info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub CommandButton1_Click()
ERRORForm.Show
Unload Me
End Sub

Private Sub Image2_Click()
ActiveWorkbook.FollowHyperlink Address:="https://hankvba.blogspot.com/2018/03/autocad-vba.html", NewWindow:=True
End Sub

Private Sub Label12_Click()
ActiveWorkbook.FollowHyperlink Address:="https://youtu.be/nH2Y0yFCZHM", NewWindow:=True
End Sub

Private Sub Label14_Click()
ActiveWorkbook.FollowHyperlink Address:="https://drive.google.com/file/d/15YBr3PcAcV0MDOYtxsYRgYfFXP1RWCxX/view?usp=share_link", NewWindow:=True
End Sub

Private Sub Label15_Click()
ActiveWorkbook.FollowHyperlink Address:="https://creativecommons.org/licenses/by-nc/3.0/tw/legalcode", NewWindow:=True
End Sub

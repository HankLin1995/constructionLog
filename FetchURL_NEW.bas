Attribute VB_Name = "FetchURL_NEW"
Sub FetchURL_Main()

MAC_ADDRESS = getMacAddress

Debug.Print MAC_ADDRESS

If MAC_ADDRESS = "" Then
    Debug.Print "請確認本機是否有連上網路~!"
Else

    Status = AccessStatus(MAC_ADDRESS)

    myChoose = split(Status, ",")

    If myChoose(0) <> "PASS" Then

        MsgBox "您的版本已無法繼續使用!", vbCritical
        
        frm_EndQRCode.Show
    
    Else
    
        Call ShowDialoge(myChoose(3))
        Call ShowDialoge(myChoose(2))  ', myChoose(3))

        '==========1206Add for 按鈕控制 ============
        
        Dim o As New clsUserInformation
        o.showCmd

    End If

End If

End Sub

Sub ShowDialoge(ByVal s1 As String) ', ByVal s2 As String)

If s1 = "" Then Exit Sub

tmp = split(s1, ":")

frmMSG.Label1.caption = tmp(0) ' caption
frmMSG.TextBox1.Value = tmp(1) ' s1
frmMSG.Show

End Sub

Function AccessStatus(ByVal mac_add As String) As String

KEEPACCESS:

Dim o As New clsFetchURL
Dim bIsClientSigned As Boolean

myURL = o.CreateURL("Access", mac_add)
Status = o.ExecHTTP(myURL)

myChoose = split(Status, ",")

Select Case myChoose(0) 'Status

Case "PASS"

    Debug.Print "驗證通過!"
    
Case "NOT_FOUND"

    Debug.Print "找不到資料庫有你的本機序號"

    bIsClientSigned = IsClientSigned(mac_add) '進行註冊嘗試並回傳結果
    
    If bIsCliendSigned = False Then GoTo KEEPACCESS

Case "ARRIVED":

Case Else

    Debug.Print Status

End Select

AccessStatus = Status

End Function

Function IsClientSigned(ByVal mac_add As String) As Boolean

'進行註冊並回傳註冊狀態
'1.已註冊:True
'2.註冊通過:False

IsClientSigned = False

Dim o As New clsFetchURL

myURL = o.CreateURL("Sign", mac_add)

If o.ExecHTTP(myURL) = "signed" Then
    MsgBox "該電腦已經被註冊過了!", vbCritical
    IsClientSigned = True
Else

myURL = o.CreateURL("SignDetail", mac_add, , , myMail)
Call o.ExecHTTP(myURL)

End If

End Function

'======這裡確定不會動==========

Function getMacAddress()

Dim objVMI As Object
Dim vAdptr As Variant
Dim objAdptr As Object
'Dim adptrCnt As Long


Set objVMI = GetObject("winmgmts:\\" & "." & "\root\cimv2")
Set vAdptr = objVMI.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")

For Each objAdptr In vAdptr
    If Not IsNull(objAdptr.MACAddress) And IsArray(objAdptr.IPAddress) Then
        For adptrCnt = 0 To UBound(objAdptr.IPAddress)
        If Not objAdptr.IPAddress(adptrCnt) = "0.0.0.0" Then
            GetNetworkConnectionMACAddress = objAdptr.MACAddress
            Exit For
        End If
        Next
    End If
Next

getMacAddress = GetNetworkConnectionMACAddress

End Function




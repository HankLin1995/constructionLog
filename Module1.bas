Attribute VB_Name = "Module1"
'1.CheckIfSigned
'2.CheckPASS

Sub test_CheckIsRepeat()

Dim o As New clsPCCES

o.checkIsRepeat

End Sub

Sub test_cmdGetDayReport() 'ByVal sDate As String, ByVal eDate As String, ByVal mode As Byte)

Dim obj As New clsDayReport


'============getInformationbyForm===============

With DayReportForm

obj.StartDate = .tbosDate
obj.EndDate = .tboeDate

obj.print_mode = 3 'print_mode

End With

'==============================================

Call obj.getInterval(sr, er) '取得日期起迄的數字型態

'Set wb = Workbooks.Add

For r = er To sr Step -1

    'get the codes
    'ThisWorkbook.Activate
    
    Set coll_code = obj.getCodes(obj.workDate + r - 1) '根據日期取得日期Codes
    
    Debug.Print obj.workDate + r - 1
    
    If coll_code.count <> 0 Then
    
        For Each it In coll_code
        
            Debug.Print r & ":" & it
        
        Next
    
    End If
    
Next
    
'    If print_mode = 3 Or print_mode = 4 Then
'
'        Dim coll_code_new As New Collection
'
'        For Each it In coll_code
'
'            If Split(it, "-")(1) = "1" Then
'
'                coll_code_new.Add Split(it, "-")(0)
'
'            End If
'
'        Next
'
'        Set coll_code = coll_code_new
'
'    End If
'

End Sub

Sub test_0509()

Dim obj As New clsDayReport

r = 28

Debug.Print obj.workDate + r - 1

Set coll_code = obj.getCodes(obj.workDate + r - 1)

For Each it In coll_code

    Debug.Print r & ":" & it

Next

End Sub

Sub test_tranNum()

s = "(2)"

Dim o As New clsFunction

Debug.Print o.tranCharcter_NUM(s)

End Sub

Sub test_adddata()

Set rng = Range("H1")
region = Array("□有 □無", "■有 □無", "□有 ■無")

Call data_validation_from_array(rng, region)

End Sub

Sub data_validation_from_array(ByVal rng As Range, ByVal region)

'Dim region, product As Variant
'Dim region_range, product_range As Range

region = Array("North", "South", "East", "West")

region = Array("□有 □無", "■有 □無", "□有 ■無")

With rng.Validation
.Delete
.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=Join(region, ",")
.IgnoreBlank = True
.InCellDropdown = True
.InputTitle = ""
.ErrorTitle = "Error"
.InputMessage = ""
.ErrorMessage = "Please Provide a Valid Input"
.ShowInput = True
.ShowError = True
End With

End Sub

Sub test_AddValidation()

arr = Array("□有 □無", "■有 □無", "□有 ■無")

Range("G" & r + 5).Validation.Delete
Range("G" & r + 5).Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
Formula1:=Join(arr, ",")

End Sub

Sub t()
Sheets("授權").Visible = True
End Sub

Sub CheckIfSigned()

Dim o As New clsUserInformation
o.hideCmd

MAC_ADDRESS = o.getMacAddress

If MAC_ADDRESS = "" Then
    Debug.Print "請確認本機是否有連上網路~!"
Else
    IsSigned = o.checkIsExist(MAC_ADDRESS)

    If IsSigned Then
    
        Debug.Print MAC_ADDRESS & "已經被使用了!"
        
    Else
        
        Debug.Print "試用版"
        Call o.Login(MAC_ADDRESS)
        Call test_SignClient
        ThisWorkbook.Save
  
    End If
    
    Call test_Access(MAC_ADDRESS)

End If

End Sub


Sub test_SignClient()

Dim o As New clsFetchURL

myURL = o.CreateURL("Sign") ', "Hank", "YLIA", "apple@mm")

Debug.Print "sign PASS"

If o.ExecHTTP(myURL) = "signed" Then
    MsgBox "該電腦已經被註冊過了!", vbCritical
End If

End Sub

Sub test_Access(ByVal mac_add As String)

Dim ui As New clsUserInformation
Dim o As New clsFetchURL

myURL = o.CreateURL("Access")

Status = o.ExecHTTP(myURL)

Select Case Status

Case "PASS"
    Debug.Print "驗證通過!"
    ui.showCmd
    
Case "NOT_FOUND"
    MsgBox "找不到資料庫有你的本機序號，請將報告傳送至管理人員", vbCritical
    ERRORForm.Show
    
Case "ARRIVED":

    MsgBox "偵測到使用天數為0日，如果要使用請先進行授權!", vbInformation
    SignDetailForm.Show
Case Else

    MsgBox Status

End Select

End Sub

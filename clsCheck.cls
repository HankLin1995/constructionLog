VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public MainRowColl As Object

Private Sub Class_Initialize()

Dim o As New clsWriteData
Set MainRowColl = o.getMainRowColl

End Sub

Function checkIsValidDate(ByVal mydatestr As String)

IsValidDate = True

If Len(mydatestr) <> 7 Then GoTo ERRORHANDLE

Y = mid(mydatestr, 1, 3)
m = mid(mydatestr, 4, 2)
d = mid(mydatestr, 6, 2)

On Error GoTo ERRORHANDLE:

a = CDate(1911 + Y & "/" & m & "/" & d)

Exit Function

ERRORHANDLE:
MsgBox "日期格式不正確!", vbCritical: Sheets("日報填寫").Protect: End

End Function

Sub checkUseDays(ByVal sDate As Date, ByVal eDate As Date) '輸出日期限制

'checkUseDays = True 'new

usedays = 60

If eDate - sDate > usedays Then

    MsgBox "日期不能超過" & usedays & "天" & vbNewLine & "你的起始日期為" & sDate & vbNewLine & "結束日期最多只到" & sDate + 60
    
    Sheets("日報填寫").Protect
    
    End
    'checkUseDays = False 'new

End If

End Sub

Sub checkIsHaveDataInDates(ByVal sDate As Date, ByVal eDate As Date)

Dim recObj As New clsRecord
Dim funObj As New clsFunction

recObj.mode = 1
Set coll = recObj.getSameCode

For Each it In coll

    tmp = Split(it, "-")(0)
    
    myDate = funObj.tranDate(tmp)

    If myDate >= sDate And myDate <= eDate Then Exit Sub

Next

MsgBox sDate & "~" & eDate & "未含有任何資料!", vbCritical: Sheets("日報填寫").Protect: End

End Sub

Function checkInformation(ByVal recCode As String)

If recCode = "" Then MsgBox "未新增報表編號", vbCritical: Sheets("日報填寫").Protect: End

End Function

Function checkIsLocEmpty(ByVal myLoc As String)

If myLoc = "" Then MsgBox "未填寫工程地點", vbCritical: Sheets("日報填寫").Protect: End 'checkIsLocEmpty = InputBox("請填寫工程地點")

End Function

Function checkIsDataEmpty() 'only check 施工工項

With Sheets("日報填寫")

    For i = 1 To 4
    
        For r = MainRowColl(i) + 2 To MainRowColl(i + 1) - 1
    
            recItem = .Cells(r, 3)
            recAmount = .Cells(r, 5)
            
            If recItem <> "" Then
            
                If recAmount = "" Then msg = msg & recItem & "未填寫數量!" & vbNewLine
            
            End If
    
        Next
    
    Next

End With

If msg <> "" Then

    MsgBox msg, vbCritical: End

End If

End Function

Function checkIsDataUsed()

Dim coll As New Collection

With Sheets("日報填寫")

    For i = 1 To 4
    
        For r = MainRowColl(i) + 2 To MainRowColl(i + 1) - 1
        
            recKey = .Cells(r, 1)
            recItem = .Cells(r, 3)
            
            On Error GoTo ERRORHANDLE
            
            If recKey <> "" Then coll.Add recKey, recKey
        
        Next
    
    Next

End With

Exit Function

ERRORHANDLE:

MsgBox recItem & "被重複使用了!!", vbCritical: End

End Function

Function checkIsDataUndefine()

Dim coll As New Collection

With Sheets("日報填寫")
    
    For i = 1 To 4
    
        For r = MainRowColl(i) + 2 To MainRowColl(i + 1) - 1
        
            recItem = .Cells(r, 3)
            
            On Error GoTo ERRORHANDLE
            If recItem Like "*錯誤*" Then Debug.Print recItem
        
        Next
        
    Next

End With

Exit Function

ERRORHANDLE:
MsgBox "篩選欄位有誤!!", vbCritical: End

End Function

Function checkTmpNameExist(ByVal s As String)

checkTmpNameExist = True

With Sheets("範本資料庫")

Set rng = .Columns("B").Find(s)

If rng Is Nothing Then
    checkTmpNameExist = False
Else
    MsgBox "名稱重複或空白，請輸入別的範本名稱!", vbCritical: End
End If
End With

End Function



VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWriteData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private checkObj As New clsCheck
Private recObj As New clsRecord

Private shtData As Object
Private shtDB As Object
Public recCode As String
Public recDate As String
Public recLoc As String
Public test_mode As Boolean

Private MainRowColl As New Collection


Private Sub Class_Initialize()

Set shtData = Sheets("日報填寫")
Set shtDB = Sheets("日報資料庫")
shtData.Unprotect

Set MainRowColl = getMainRowColl

End Sub

Private Sub Class_Terminate(): shtData.Protect: End Sub

Sub readData()

Call recObj.dealExistData(recCode)

With shtData
    
    For i = 1 To 5
    
        sr = MainRowColl(i) + 2
        er = MainRowColl(i + 1) - 1

        For r = sr To er
        
            If i <> 5 Then '日報、人力、機具、材料
        
                If .Cells(r, 1) <> "" Then
    
                    recNum = .Cells(r, 2)
                    recItem = .Cells(r, 3)
                    recUnit = .Cells(r, 4)
                    recAmount = .Cells(r, 5)
    
                    arr = Array(recCode, recDate, recLoc, recNum, recItem, recUnit, recAmount)
    
                    Call keyInData(arr, i)
    
                End If
                
            Else '文件
    
                    recNum = .Cells(r, 2)
                    recItem = .Cells(r, 3)
                    recUnit = .Cells(r, 4)
                    recAmount = .Cells(r, 5)
    
                    arr = Array(recCode, recDate, recLoc, recNum, recItem, recUnit, recAmount)
    
                    Call keyInData(arr, i)
            
            End If
        
        Next
    
    Next
    
    'Call recObj.sortRecord
    
    .Activate

End With

End Sub

Function getDB_name(ByVal mode As Byte)

Select Case mode

Case 1: getDB_name = "日報資料庫"
Case 2: getDB_name = "材料資料庫"
Case 3: getDB_name = "人員資料庫"
Case 4: getDB_name = "機具資料庫"
Case 5: getDB_name = "文件資料庫"

End Select

End Function

Sub keyInData(arr, ByVal mode As Byte)

With Sheets(getDB_name(mode))

    .Unprotect
    
    lr = .Cells(.Rows.Count, 1).End(xlUp).Row + 1
    
    For i = 1 To 7
    
        .Cells(lr, i) = arr(i - 1)
    
    Next
    
    lr = .Cells(.Rows.Count, 1).End(xlUp).Row

    .Range("A2").Resize(lr, 8).Sort key1:=.Columns(1), order1:=xlAscending
    
    .Protect

End With

End Sub

Sub readInformation()

With shtData

    recCode = .Range("B1")
    recDate = .Range("B2")
    recLoc = .Range("B3")
    
    Call checkObj.checkIsLocEmpty(recLoc)

End With

End Sub

Sub getInformation()

With shtData
    
    If test_mode = False Then
        recDate = InputBox("新增填報日期(EX:1110316)")
    Else
        a = Format(WorksheetFunction.RandBetween(1, 12), "00")
        b = Format(WorksheetFunction.RandBetween(1, 28), "00")
        recDate = "111" & a & b
    End If
    
    Set checkObj.MainRowColl = getMainRowColl
    Call checkObj.checkIsValidDate(recDate)
    
    recCode = recDate & "-" & recObj.getCodeCount(recDate)
    .Unprotect
    .Range("B1") = recCode
    '.Range("B2") = recDate
    .Range("B2") = "'" & recDate
    .Protect
End With

If test_mode = True Then
    Debug.Print "新增完成!編號為" & recCode
Else
    MsgBox "新增完成!編號為" & recCode, vbInformation
End If
End Sub

Sub clearDataOne(ByVal mode As Byte)

With shtData

    If mode <> 5 Then

        For i = mode To mode
        
            sr = MainRowColl(i)
            er = MainRowColl(i + 1)
        
            cnt = er - sr - 2
            
            .Range("A" & sr + 2).Resize(cnt, 1).ClearContents
            .Range("E" & sr + 2).Resize(cnt, 1).ClearContents
        
        Next
    
    Else
    
        For j = 6 To 9
        
        .Range("E" & MainRowColl(5) + j).ClearContents
        
        Next
    
    End If
    
End With

End Sub

Sub clearDataAll()

With shtData
    .Unprotect
    For i = 1 To 4
    
        sr = MainRowColl(i)
        er = MainRowColl(i + 1)
    
        cnt = er - sr - 2
        
        .Range("A" & sr + 2).Resize(cnt, 1).ClearContents
        .Range("E" & sr + 2).Resize(cnt, 2).ClearContents
    
    Next
    
    For j = 6 To 9
    
    .Range("E" & MainRowColl(5) + j).ClearContents
    
    Next
    .Protect
End With

End Sub

Sub clearInformation()

With shtData

    .Unprotect
    
    .Range("B1:B2").ClearContents
    .Range("B3:C3").ClearContents
    
    .Protect

End With

End Sub

Sub reverseData(ByVal arr, ByVal i As Byte)

With shtData
    
    .Unprotect
    
    myMainRow = MainRowColl(i)
    
    lr = .Cells(myMainRow, 1).End(xlDown).Row
    'lr = .Cells(.Rows.Count, 1).End(xlUp).Row
    
    r = lr + 1
    
    .Range("B1") = arr(0) '重複工作
    .Range("B2") = arr(1) '重複工作
    .Range("B3") = arr(2) '重複工作
    
    .Range("E" & r) = arr(6)
    
    .Cells(r, 1) = arr(3) & ">" & arr(4)
    
    Debug.Print r & "~" & .Cells(r, 1)

    .Protect
    
End With

End Sub

Sub hideRng(ByVal mode As Byte, ByVal IsHidden As Boolean, Optional IsData As Boolean = False) 'ByVal mode)

ActiveWindow.SmallScroll Down:=-1000

    With Sheets("日報填寫")
    
    .Unprotect
    
    sr = MainRowColl(mode)
    er = MainRowColl(mode + 1) - 1
    
    'If IsReset = False Then
    
   '     If mode <> 5 Then .Range("A" & sr + 2).Select
    
   ' End If
    'If IsData = False Then
    
        For r = sr + 1 To er
            
            .Rows(r).Hidden = IsHidden
            
        Next
'
'    Else
'
'        For r = sr + 1 To er
'
'            If .Cells(r, 1) = "" Then .Rows(r).Hidden = IsHidden
'
'        Next
'
'    End If
    
    .Protect

End With

End Sub

Function getMainRowColl()

Dim coll As New Collection

With Sheets("日報填寫")

arr = Array("一、施工工項", "二、材料管理", "三、人員管理", "四、機具管理", "五、文件資料")

For Each it In arr

    Set rng = .Columns("A").Find(it)
    
    coll.Add rng.Row

Next

lr = .Cells(.Rows.Count, 2).End(xlUp).Row

coll.Add coll(5) + 10

End With

Set getMainRowColl = coll

End Function

Sub setValidation() 'for 文件資料

Dim myfun As New clsFunction

Set coll = getMainRowColl

r = coll(5)

With Sheets("日報填寫")

    .Unprotect
    
    For j = 2 To 5
    
        Set rng = .Range("E" & r + j)
        region = Array("□有 □無", "■有 □無", "□有 ■無")
        If j = 4 Then region = Array("□有 □■無 □無新進勞工", "■有 □無 □無新進勞工", "□有 □無 ■無新進勞工")
    
        Call myfun.data_validation_from_array(rng, region)

    Next

    .Protect

End With

Exit Sub

With Sheets("日報填寫")
    .Unprotect
    .Range("E" & r + 2).Validation.Delete
    .Range("E" & r + 2).Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
Formula1:="=日報填寫!$I" & r + 2 & ":$I" & r + 2 + 2

    .Range("E" & r + 3).Validation.Delete
    .Range("E" & r + 3).Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
Formula1:="=日報填寫!$I" & r + 2 & ":$I" & r + 2 + 2

    .Range("E" & r + 4).Validation.Delete
    .Range("E" & r + 4).Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
Formula1:="=日報填寫!$K" & r + 2 & ":$K" & r + 2 + 3

    .Range("E" & r + 5).Validation.Delete
    .Range("E" & r + 5).Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
Formula1:="=日報填寫!$I" & r + 2 & ":$I" & r + 2 + 2
    .Protect
End With

End Sub


Sub getWorkPlaceValidation()

Dim coll As New Collection
Dim myFunc As New clsMyfunction
Dim func As New clsFunction

With Sheets("標案設定")

    lr = .Cells(.Rows.Count, "D").End(xlUp).Row
    
    For r = 2 To lr

        coll.Add CStr(.Cells(r, "D"))

    Next

    arr = myFunc.tranColl2Array(coll)

End With

With Sheets("日報填寫")

    .Unprotect

    Call func.data_validation_from_array(.Range("B3"), arr)

    .Protect

End With

End Sub


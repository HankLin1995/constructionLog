Attribute VB_Name = "UnitTest"
Sub test_getDataByDate_second()

Dim o As New clsDayReport

o.print_mode = 4
recDate = #2/28/2022#
recCode = "1110228-2"

Call o.getDataByDate_second(recDate, recCode)

End Sub

Sub test_IsDataUsed()

Dim o As New clsCheck
o.checkIsDataUsed

End Sub

Sub test_getMoneyBYitem()

s = Range("A6") & ">" & Range("B6")

Debug.Print s

Dim o As New clsPCCES
Debug.Print o.getMoneyByItemKey(s)


End Sub

Sub test_getSumMoney()

Dim o As New clsPCCES

Debug.Print o.getSumMoney

End Sub

Sub test_batchAddData()

Call test_refreshDB

test_mode = True

cnt = InputBox("總共要新增幾筆?")

If Not IsNumeric(cnt) Then End

For i = 1 To cnt
    Call cmdButton.cmdCreateNewData(True)
    Call UnitTest.test_adddata
    Call cmdButton.cmdSaveData(, True)
Next

MsgBox "新增完成!!", vbInformation

End Sub

Sub test_checkReportRow()

Dim o As New clsDayReport

Call o.checkReportCount(3, 3)

End Sub

Sub test_getSRER()

Dim o As New clsDayReport

For i = 1 To 3

Call o.getReportSrEr(sr, er, i)

Debug.Print "sr=" & sr & ",er=" & er & ">mode=" & i

Next

End Sub

Sub test_getUsedItemByDateFromMode()

Dim o As New clsDayReport

Set coll = o.getUsedItemByDate("1110425", 3)

For Each it In coll

    Debug.Print it

Next

End Sub

Sub test_refreshDB()

Dim o As New clsPCCES
o.RefreshDB

Sheets("契約詳細表").Activate

End Sub

Sub test_checkDate()

testdata = "1111231"

Dim obj As New clsCheck

obj.checkIsValidDate (testdata)

Debug.Print "test_checkDate>>>PASS"

End Sub

Sub test_checkIsHaveDataInDates()

Dim myfun As New clsFunction

sd = myfun.tranDate("1110401")

ed = myfun.tranDate("1110402")

Dim obj As New clsCheck

Call obj.checkIsHaveDataInDates(sd, ed)

End Sub

Sub test_getDataByDate()

recDate = #4/5/2022#
recCode = "1110405-1"

Dim obj As New clsDayReport
Call obj.getDataByDate(recDate, recCode)

End Sub

Sub test_getmacaddress()

Dim o As New clsMacAddress
o.Login

End Sub

Sub test_adddata()

Dim o As New clsWriteData

o.clearDataAll

Set coll = o.getMainRowColl

For i = 1 To 4
    
    If i = 1 Then
    myRnd = WorksheetFunction.RandBetween(5, 10)
    Else
    myRnd = WorksheetFunction.RandBetween(1, 6)
    End If
    
    For j = 0 To myRnd

        If i = 1 Then
            test_item = getTestItem("N")
        ElseIf i = 2 Then
            test_item = getTestItem("M")
        ElseIf i = 3 Then
            test_item = getTestItem("L")
        ElseIf i = 4 Then
            test_item = getTestItem("E")
        End If
    
        r = coll(i) + 2
        
        With Sheets("日報填寫")
        
            Set rng = .Columns("A").Find(test_item)
            
            If rng Is Nothing Then
        
            .Range("B3") = "測試地點" & Format(Now(), "MMDDHHmm")

            .Cells(r + j, 1) = test_item
            .Cells(r + j, 5) = WorksheetFunction.RandBetween(1, 10)

            End If

        End With

    Next

Next

End Sub

Function getTestItem(ByVal property As String)

Select Case property

Case "N"

    With Sheets("契約詳細表")

        lr = .Cells(.Rows.Count, 1).End(xlUp).Row
        r = WorksheetFunction.RandBetween(1, lr)
        test_item = .Range("A" & r)
        test_note = .Range("G" & r)
        
        If test_note = "" Then getTestItem = test_item
    
    End With
    
Case "M", "L", "E"

    getTestItem = Sheets("工料設定").Range("A" & getRandRow(property))

End Select

End Function


Function getRandRow(ByVal mode As String)

Dim o As New clsMLE

With Sheets("工料設定")

    lr = .Cells(.Rows.Count, 1).End(xlUp).Row
    
    For r = 2 To lr
    
        If mid(.Cells(r, 2), 1, 1) = mode Then
        
            sr = r: Exit For
            
        End If
        
    Next
    
    cnt = o.countType(mode)
    
    er = sr + cnt - 1
    
    getRandRow = WorksheetFunction.RandBetween(sr, er)

End With

End Function


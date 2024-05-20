Attribute VB_Name = "cmdButton"

Sub cmdGetDayReport() 'ByVal sDate As String, ByVal eDate As String, ByVal mode As Byte)

Dim obj As New clsDayReport

'todo:
'1.getIntervalDate
'2.createWorkbook
'3.getData

Application.ScreenUpdating = False
Application.DisplayAlerts = False

'============getInformationbyForm===============

With DayReportForm

obj.StartDate = .tbosDate
obj.EndDate = .tboeDate

If .optMode1.Value = True Then
    print_mode = 1
ElseIf .optMode2.Value = True Then
    print_mode = 2
ElseIf .optMode3.Value = True Then
    print_mode = 3
ElseIf .optMode4.Value = True Then
    print_mode = 4
End If

obj.print_mode = print_mode

End With

'==============================================

Call obj.getInterval(sr, er) '取得日期起迄的數字型態

Set wb = Workbooks.Add

For r = er To sr Step -1

    'get the codes
    ThisWorkbook.Activate
    
    Set coll_code = obj.getCodes(obj.workDate + r - 1) '根據日期取得日期Codes
    
    For j = coll_code.Count To 1 Step -1
    
        code = coll_code(j)
        
        Debug.Print code
    
        ThisWorkbook.Activate
        
        If print_mode = 1 Or print_mode = 2 Then
        
            Call obj.getDataByDate(obj.workDate + r - 1, code) '取得該日日報
            Call obj.hideEmptyRow
            
            If print_mode = 1 Then Call obj.hideEmpyNum
        
        ElseIf print_mode = 3 Or print_mode = 4 Then
        
            'Call obj.getDataByDate_second(obj.workDate + r - 1, code)
        
        End If
        
        Call obj.outputData(wb, code)
        
    Next

Next

Application.DisplayAlerts = False

For Each sht In wb.Sheets
    If sht.Name Like "工作表*" Then sht.Delete
Next

Application.DisplayAlerts = True

'If wb.Sheets.Count > 1 Then wb.Sheets("工作表1").Delete

If print_mode = 3 Or print_mode = 4 Then

    For Each sht In wb.Sheets

        code = Split(sht.Name, "-")(0)
        Page = Split(sht.Name, "-")(1)
        
        Application.DisplayAlerts = False
            If Page = 1 Then
                sht.Name = code
            Else
                sht.Delete
            End If
        Application.DisplayAlerts = True
        

    Next

End If

Application.DisplayAlerts = True
Application.ScreenUpdating = True

ThisWorkbook.Sheets("日報填寫").Activate

End Sub

Sub cmdGetLongReport()

Dim obj As New clsLongReport

With LongReportForm

    obj.StartDate = .tbosDate
    obj.EndDate = .tboeDate
    
    If .optMode1.Value = True Then
        print_mode = 1
    ElseIf .optMode2.Value = True Then
        print_mode = 2
        targetMode = "M"
    ElseIf .optMode3.Value = True Then
        print_mode = 3
        targetMode = "L"
    ElseIf .optMode4.Value = True Then
        print_mode = 4
        targetMode = "E"
    End If

    obj.print_mode = print_mode

End With

obj.clearLongReport

If print_mode = 1 Then
    obj.getReportItemByPCCES
Else
    obj.getReportItemByMLE (targetMode)
End If

obj.KeyInLongReport
obj.SumReportAmount

Set wb = Workbooks.Add

If targetMode = "" Then

Call obj.outputData(wb, True)

Else

Call obj.outputData(wb, False)

End If

ThisWorkbook.Sheets("日報填寫").Activate

wb.Activate

End Sub

Sub cmdGetDataByTmpName(ByVal tmpType As String, ByVal tmpName As String) 'only for 施工工項

Dim dataObj As New clsWriteData
Call dataObj.hideRng(1, False)

dataObj.clearDataOne (CByte(tmpType))

Dim obj As New clsTmp

Call obj.getDatabyTmp(tmpType, tmpName)

End Sub

Sub cmdRecordTmp()

tmpType = InputBox("請輸入範本種類" & vbNewLine & "1.施工工項" & vbNewLine & "2.材料管理", , "1")
If Not (tmpType = "1" Or tmpType = "2") Then MsgBox "請輸入1或2", vbCritical: End

tmpName = InputBox("請輸入範本名稱")

Dim checkObj As New clsCheck
checkObj.checkTmpNameExist (tmpName)
checkObj.checkIsDataUndefine
checkObj.checkIsDataUsed

Dim obj As New clsTmp

Call obj.recordData(tmpType, tmpName)

End Sub

Sub cmdGetDataByCode(ByVal myCode As String)

'Application.ScreenUpdating = False

Dim obj As New clsWriteData

obj.clearDataAll

For i = 1 To 5
    Call obj.hideRng(i, False)
Next

Dim recObj As New clsRecord
recObj.getDatabyCode (myCode)

Dim pccesObj As New clsPCCES

pccesObj.setValidation

Dim MLEobj As New clsMLE
MLEobj.setValidation_MLE

'obj.getWorkPlaceValidation
'
'For i = 1 To 5
'    Call obj.hideRng(i, False, True)
'Next

'Application.ScreenUpdating = True

End Sub

Sub cmdDeleteData() '刪除

Application.ScreenUpdating = False

Dim obj As New clsWriteData

obj.clearDataAll

Call cmdSaveData("DeleteMode")

Application.ScreenUpdating = True

End Sub

Sub cmdSaveData(Optional mode As String, Optional test_mode As Boolean = False)

Application.ScreenUpdating = False

'todo:
'1.readInformation
'2.readData
'3.recData

Dim checkObj As New clsCheck
Dim obj As New clsWriteData

Set checkObj.MainRowColl = obj.getMainRowColl

If mode = "DeleteMode" Then
    checkObj.checkIsDataUndefine
    checkObj.checkIsDataUsed
Else
    checkObj.checkIsDataUndefine '是否資料不合法
    checkObj.checkIsDataEmpty '是否為空值(工項)
    checkObj.checkIsDataUsed '是否有重複(工項~MLE)
End If

obj.readInformation

Call checkObj.checkInformation(obj.recCode)

obj.ReadData

obj.clearInformation
obj.clearDataAll

If test_mode = False Then

    If mode = "" Then
        MsgBox "儲存完成!編號為" & obj.recCode, vbInformation
    Else
        MsgBox "編號為" & obj.recCode & "已作廢!", vbInformation
    End If

Else

    If mode = "" Then
        Debug.Print "儲存完成!編號為" & obj.recCode,
    Else
        Debug.Print "編號為" & obj.recCode & "已作廢!"
    End If


End If

Application.ScreenUpdating = True

End Sub

Sub cmdCreateNewData(Optional test_mode As Boolean = False)

Application.ScreenUpdating = False

'todo:
'1.getInformation
'2.waiting for keying data

Dim obj As New clsWriteData

obj.test_mode = test_mode

obj.clearInformation
obj.getInformation
obj.clearDataAll

'Call obj.hideRng(1, True)

For i = 1 To 5
    Call obj.hideRng(i, False)
Next

'Call obj.hideRng(1, False)

Dim pccesObj As New clsPCCES

msg = MsgBox("是否載入已經完成的項目?", vbYesNo + vbInformation)

If msg = vbYes Then

pccesObj.setValidation2

Else

pccesObj.setValidation

End If

Dim MLEobj As New clsMLE
MLEobj.setValidation_MLE

Call obj.setValidation
Call obj.getWorkPlaceValidation

Application.ScreenUpdating = True

End Sub

Sub cmdGetPCCES()

Dim obj As New clsPCCES

obj.getFileName
obj.clearPCCES_data
obj.getAllContents
obj.checkIsRepeat
obj.RefreshDB
obj.setValidation

Sheets("契約詳細表").Activate

End Sub

Sub showSearchForm()

SearchForm.Show

End Sub

Sub showTmpForm()

TmpForm.Show

End Sub

Sub showMLEForm()

MLEForm.Show (0)

End Sub

Sub showDayReportForm()

DayReportForm.Show (0)

'ThisWorkbook.Sheets("日報填寫").Activate

End Sub

Sub showLongReportForm()

LongReportForm.Show (0)

'ThisWorkbook.Sheets("日報填寫").Activate

End Sub

Sub cmdhideRng1()

Dim writeObj As New clsWriteData
Call writeObj.hideRng(1, True)

End Sub

Sub cmdhideRng2()

Dim writeObj As New clsWriteData
Call writeObj.hideRng(2, True)

End Sub

Sub cmdhideRng3()

Dim writeObj As New clsWriteData
Call writeObj.hideRng(3, True)

End Sub

Sub cmdhideRng4()

Dim writeObj As New clsWriteData
Call writeObj.hideRng(4, True)

End Sub

Sub cmdhideRng5()

Dim writeObj As New clsWriteData
Call writeObj.hideRng(5, True)

End Sub

Sub cmdOpenRng1()

Dim writeObj As New clsWriteData
Call writeObj.hideRng(1, False)

End Sub

Sub cmdOpenRng2()

Dim writeObj As New clsWriteData
Call writeObj.hideRng(2, False)

End Sub

Sub cmdOpenRng3()

Dim writeObj As New clsWriteData
Call writeObj.hideRng(3, False)

End Sub

Sub cmdOpenRng4()

Dim writeObj As New clsWriteData
Call writeObj.hideRng(4, False)

End Sub

Sub cmdOpenRng5()

Dim writeObj As New clsWriteData
Call writeObj.hideRng(5, False)

End Sub

Sub cmdGetProgByInter() '20230624

Call checkProgSetting
Set collProg = getProgColl

With Sheets("天氣設定")

    lr = .Cells(.Rows.Count, 4).End(xlUp).Row

    For r = 2 To lr
    
        myProg = .Cells(r, 4)
        
        If myProg = "" Then
        
            For i = 1 To collProg.Count
                
                tmp = Split(collProg(i), ":")
                
                If r <= CInt(tmp(0)) Then
                
                    r1 = Split(collProg(i - 1), ":")(0)
                    p1 = Split(collProg(i - 1), ":")(1)
                    
                    r2 = Split(collProg(i), ":")(0)
                    p2 = Split(collProg(i), ":")(1)
                    
                    newProg = Round(((r2 - r) * p1 + (r - r1) * p2) / (r2 - r1), 4)
                    
                    Exit For
                
                End If
            
            Next
            
            .Cells(r, 4) = newProg
        
        End If
        
    Next

End With

End Sub

'-----------FUNCTION-----------------

Sub checkProgSetting()

fixStartDate = Sheets("標案設定").Range("B3")
fixEndDate = Sheets("標案設定").Range("B4")

With Sheets("天氣設定")

    lr = .Cells(.Rows.Count, 1).End(xlUp).Row
    
    progStartDate = .Cells(2, 1)
    progStartProg = .Cells(2, 4)
    progEndDate = .Cells(lr, 1)
    progEndProg = .Cells(lr, 4)
    
    If progStartDate <> fixStartDate Then
    
        MsgBox ("開工日「" & progStartDate & "」，與標案設定開工日「" & fixStartDate & "」不一樣!"), vbCritical
        End
        
    End If
    
    If progEndDate <> fixEndDate Then
    
        MsgBox ("竣工日「" & progEndDate & "」，與標案設定竣工「" & fixEndDate & "」不一樣!"), vbCritical
        End
        
    End If
    
    If progStartProg = "" Then
    
        .Cells(2, 4) = 0
        MsgBox "系統自動於開工日補上0%", vbInformation
        
    End If
    
    If progEndProg <> 1 Then
    
        .Cells(lr, 4) = 1
        MsgBox "系統自動於竣工日補上100%", vbInformation
    
    End If
    
End With

End Sub

Function getProgColl()

Dim coll As New Collection

With Sheets("天氣設定")

    lr = .Cells(.Rows.Count, 1).End(xlUp).Row

    '------main-----------
    
    For r = 2 To lr
    
        mydate = .Cells(r, 1)
        myProg = .Cells(r, 4)
        
        If myProg <> "" Then
        
            coll.Add r & ":" & myProg
        
        End If
    
    Next
    
End With

Set getProgColl = coll

If coll.Count = 2 Then MsgBox ("建議在預定進度的欄位「D」填寫進度，內差成果才會比較準確!"), vbCritical

End Function

Sub cmdAddDefineWorkPlace()

myPlace = InputBox("請輸入自定義的工程地點:")

With Sheets("日報填寫")
    .Range("B3").Validation.Delete
    .Range("B3") = myPlace
End With

End Sub



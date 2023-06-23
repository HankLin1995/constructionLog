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
    
    For j = coll_code.count To 1 Step -1
    
        code = coll_code(j)
        
        Debug.Print code
    
        ThisWorkbook.Activate
        
        If print_mode = 1 Or print_mode = 2 Then
        
            Call obj.getDataByDate(obj.workDate + r - 1, code) '取得該日日報
            Call obj.hideEmptyRow
            
            If print_mode = 1 Then Call obj.hideEmpyNum
        
        ElseIf print_mode = 3 Or print_mode = 4 Then
        
            Call obj.getDataByDate_second(obj.workDate + r - 1, code)
        
        End If
        
        Call obj.outputData(wb, code)
        
    Next

Next

If wb.Sheets.count > 1 Then wb.Sheets("工作表1").Delete

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

obj.outputData (wb)

ThisWorkbook.Sheets("日報填寫").Activate

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

obj.readData

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
pccesObj.setValidation

Dim MLEobj As New clsMLE
MLEobj.setValidation_MLE

Call obj.setValidation

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


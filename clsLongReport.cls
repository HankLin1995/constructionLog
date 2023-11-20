VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLongReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public StartDate As String
Public EndDate As String
Public print_mode As Byte

Private recObj As New clsRecord
Private checkObj As New clsCheck
Private myFun As New clsFunction
Private writeObj As New clsWriteData

Function getCollRowByKey()

Dim coll As New Collection

With Sheets("長表格式")

lr = .Cells(.Rows.Count, 1).End(xlUp).Row

For r = 4 To lr

    recKey = .Cells(r, 1)
    recItem = .Cells(r, 2)
    
    Debug.Print recKey & ">" & recItem
    
    coll.Add r, recKey & ">" & recItem

Next

Set getCollRowByKey = coll

End With

End Function

Sub clearLongReport()

With Sheets("長表格式")

Set rng = .Cells.SpecialCells(xlCellTypeLastCell)

lr = rng.Row '.Cells(.Rows.count, 1).End(xlUp).Row
lc = rng.Column '.Cells(3, .Columns.count).End(xlToLeft).Column

.Range("A4").Resize(lr, lc).ClearContents
.Range("F1").Resize(3, lc).ClearContents

End With

End Sub

Sub getReportItemByPCCES()

With Sheets("契約詳細表")

lr = .Cells(.Rows.Count, 1).End(xlUp).Row
rr = 4
For r = 2 To lr

    pcces_num = .Cells(r, 2)
    pcces_item = .Cells(r, 3)
    pcces_unit = .Cells(r, 4)
    pcces_amount = .Cells(r, 5)
    pcces_cost = .Cells(r, 6)
    pcces_note = .Cells(r, 7)
    
    If pcces_note = "" Then
    
        With Sheets("長表格式")
        
            .Cells(rr, 1) = pcces_num
            .Cells(rr, 2) = pcces_item
            .Cells(rr, 3) = pcces_unit
            .Cells(rr, 4) = pcces_cost
            
            rr = rr + 1
        
        End With

    End If

Next

End With

End Sub

Sub getReportItemByMLE(ByVal targetMode As String)

With Sheets("工料設定")

lr = .Cells(.Rows.Count, 1).End(xlUp).Row
rr = 4

For r = 2 To lr

    pcces_num = .Cells(r, 2)
    pcces_item = .Cells(r, 3)
    pcces_unit = .Cells(r, 4)
    pcces_amount = .Cells(r, 5)
    pcces_cost = .Cells(r, 6)
    
    If mid(pcces_num, 1, 1) = targetMode Then
    
        With Sheets("長表格式")
        
            .Cells(rr, 1) = pcces_num
            .Cells(rr, 2) = pcces_item
            .Cells(rr, 3) = pcces_unit
            .Cells(rr, 4) = pcces_cost
            
            rr = rr + 1
        
        End With
    
    End If

Next

End With

End Sub

Function getIntervalCodes()

'StartDate = Val(InputBox("開始日期(EX:1110101)", , "1110101"))
Call checkObj.checkIsValidDate(StartDate)
'EndDate = Val(InputBox("結束日期(EX:1110131)", , "1110131"))
Call checkObj.checkIsValidDate(EndDate)

Call checkObj.checkUseDays(myFun.tranDate(StartDate), myFun.tranDate(EndDate))
Call checkObj.checkIsHaveDataInDates(myFun.tranDate(StartDate), myFun.tranDate(EndDate))

'workDate = Sheets("標案資料").Range("B3")

sDate = myFun.tranDate(StartDate)
eDate = myFun.tranDate(EndDate)

Set getIntervalCodes = recObj.getSameCodeByDates(sDate, eDate)

End Function

Sub KeyInLongReport() '依據Code去填寫值

Set coll_code = getIntervalCodes 'getStartDate&EndDate

Set coll_row = getCollRowByKey

With Sheets("長表格式")

    For Each code In coll_code

        c = 6 + i
        .Cells(1, c) = code
        .Cells(2, c) = myFun.tranDate(Split(code, "-")(0))
        .Cells(3, c) = recObj.getRecLoc(code)
        
        Call KeyInLongItems(code, c, coll_row)
        
        If .Cells(3, c) = "" Then
        
            .Columns(c).Delete
            i = i - 1
            
        End If
        
        i = i + 1
        
    Next
    
End With

End Sub

Private Sub KeyInLongItems(ByVal myCode As String, ByVal c, ByVal coll_row)

shtName = writeObj.getDB_name(print_mode)

With Sheets(shtName) '"日報資料庫")

lr = .Cells(.Rows.Count, 1).End(xlUp).Row

For r = 2 To lr

    recCode = .Cells(r, 1)
    recItem_Key = .Cells(r, 4)
    recItem = .Cells(r, 5)
    recAmount = .Cells(r, 7)
    recNote = .Cells(r, 8)

    If recCode = myCode And recNote = "" Then
        
        data_row = coll_row(recItem_Key & ">" & recItem)
    
        'Set Rng = Sheets("長表格式").Cells.Find(recItem)
        Sheets("長表格式").Cells(data_row, c) = recAmount
    
    End If

Next

End With

End Sub

Sub SumReportAmount() '長表

With Sheets("長表格式")

lr = .Cells(.Rows.Count, 1).End(xlUp).Row
lc = .Cells(1, .Columns.Count).End(xlToLeft).Column

For r = 4 To lr

    mySum = 0

    For c = 6 To lc
        mySum = mySum + .Cells(r, c)
    Next
    .Cells(r, 5) = mySum

Next

.Activate

End With

End Sub


Sub outputData(ByVal wb As Object, ByVal IsAsk As Boolean)

If IsAsk = False Then

cnt = 1

ElseIf IsAsk = True Then

cnt = InputBox("請輸入本次為第幾次估驗", , 1)

End If

ThisWorkbook.Activate
ThisWorkbook.Sheets("長表格式").Copy (wb.Sheets(1))

Set sht = wb.Sheets(1)

sht.Name = "第" & cnt & "估驗"
sht.Range("B1") = "(自費)第" & cnt & "次估驗"

Application.DisplayAlerts = False
If wb.Sheets.Count > 1 Then wb.Sheets("工作表1").Delete
Application.DisplayAlerts = True

End Sub

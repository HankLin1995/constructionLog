VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDayReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public StartDate As String
Public EndDate As String
Public print_mode As Byte
'Private sr As Integer
'Private er As Integer

'Private shtDReport As Object
Private myFun As New clsFunction
Private checkObj As New clsCheck
Private dataObj As New clsWriteData

Public workDate As Date

Sub getInterval(ByRef sr, ByRef er) '取得日期間的數字差

'StartDate = Val(InputBox("開始日期(EX:1110101)", , "1110101"))
Call checkObj.checkIsValidDate(StartDate)
'EndDate = Val(InputBox("結束日期(EX:1110131)", , "1110131"))
Call checkObj.checkIsValidDate(EndDate)

Call checkObj.checkUseDays(myFun.tranDate(StartDate), myFun.tranDate(EndDate))
Call checkObj.checkIsHaveDataInDates(myFun.tranDate(StartDate), myFun.tranDate(EndDate))

workDate = Sheets("標案設定").Range("B3")

sr = myFun.tranDate(StartDate) - workDate + 1
er = myFun.tranDate(EndDate) - workDate + 1

End Sub

Function getCodes(ByVal mydate As Date)

Dim coll As New Collection

'With ThisWorkbook.Sheets("日報資料庫")
With ThisWorkbook.Sheets("文件資料庫")

lr = .Cells(.Rows.Count, 1).End(xlUp).Row

For r = 2 To lr

    recCode = .Cells(r, 1)
    recNote = .Cells(r, 8)
    recDate = myFun.tranDate(.Cells(r, 2))
    If mydate = recDate And recNote = "" Then
    
        On Error Resume Next
        
            coll.Add recCode, recCode
        
        On Error GoTo 0
    
    End If

Next

End With

Set getCodes = coll

End Function

Sub outputData(ByVal wb As Object, ByVal code As String)

ThisWorkbook.Activate

If print_mode = 1 Or print_mode = 2 Then

    ThisWorkbook.Sheets("日報第一聯").Copy (wb.Sheets(1))

ElseIf print_mode = 3 Or print_mode = 4 Then

    ThisWorkbook.Sheets("日報第二聯").Copy (wb.Sheets(1))
    'code = Split(code, "-")(0)
    'Page = Split(code, "-")(1)

End If

Set sht = wb.Sheets(1)

With sht
    
    .Name = code
    .Columns("A:N").Select
    '.Range("A7").Activate
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    'If print_mode = 1 Then Call hideEmpyNum(sht)

End With

End Sub

Sub hideEmptyRow()

With Sheets("日報第一聯")

For i = 1 To 2

    Call getReportSrEr(sr, er, i)
    
    For r = sr + 2 To er - 1
        
        .Rows(r).Hidden = False
        
        If .Cells(r, "A") = "" Then .Rows(r).Hidden = True
    
    Next

Next

Call getReportSrEr(sr, er, 3)

For r = sr + 2 To er - 1

        .Rows(r).Hidden = False
        
        If .Cells(r, "A") = "" And .Cells(r, "H") = "" Then .Rows(r).Hidden = True

Next

End With

End Sub

Sub hideEmpyNum() '(ByVal sht As Object)

With Sheets("日報第一聯")

For i = 1 To 2

    Call getReportSrEr(sr, er, i)
    
    For r = sr + 2 To er - 1
        
        .Rows(r).Hidden = False
        
        If .Cells(r, "H") = 0 Then .Rows(r).Hidden = True
    
    Next

Next

End With

End Sub


Sub getDataByDate(ByVal recDate As Date, ByVal recCode As String) '第一聯累計工項不為0

Dim recObj As New clsRecord

'===== basic information=======

ThisWorkbook.Activate

With Sheets("日報第一聯")

.Range("B2") = recCode
.Range("K2") = recObj.getRecLoc(recCode)
.Range("K3") = recDate

.Rows.Hidden = False

End With

'====1-施工工項、2-材料管理==========

For i = 1 To 2 '施工工項、材料管理

Set coll_item = getUsedItemByDate(recCode, recDate, i)
Call checkReportCount(coll_item.Count, i)
Call KeyInItem(coll_item, recDate, recCode, i)

Next

'======3-人員、4-機具================

Set coll_L = getUsedItemByDate(recCode, recDate, 3)
Set coll_E = getUsedItemByDate(recCode, recDate, 4)

Call checkReportCount(WorksheetFunction.Max(coll_L.Count, coll_E.Count), 3)
Call KeyInItem_Merge(coll_L, coll_E, recDate, recCode, 3)

'=====5-文件============

Call KeyInItem_Paper(recCode)

End Sub

Sub KeyInItem_Paper(ByVal recCode As String)

With ThisWorkbook.Sheets("日報第一聯")

Call getReportSrEr(sr, er, 5)

.Range("F" & sr - 2) = getPaperItem(recCode, "P0001")
.Range("H" & sr + 2) = getPaperItem(recCode, "P0002")
.Range("J" & sr + 3) = getPaperItem(recCode, "P0003")
.Range("E" & sr + 4) = getPaperItem(recCode, "P0004")
.Range("C" & sr + 5) = getPaperItem(recCode, "P0005")
.Range("E" & sr + 7) = getPaperItem(recCode, "P0006")
.Range("E" & sr + 9) = getPaperItem(recCode, "P0007")
.Range("E" & sr + 11) = getPaperItem(recCode, "P0008")


Dim secondString As String
Dim outsideString As String

myString = .Range("E" & sr + 7)

Call SplitAndCombineParenthesesStrings(myString, secondString, outsideString)

.Range("E" & sr + 7) = myString ' outsideString
Call AddCommentToCell(.Range("E" & sr + 7), secondString)
.Range("E" & sr + 7) = Replace(.Range("E" & sr + 7), "[", "")
.Range("E" & sr + 7) = Replace(.Range("E" & sr + 7), "]", "")

End With

End Sub

Function getPaperItem(ByVal recCode As String, ByVal pcode As String)

With Sheets(dataObj.getDB_name(5))

    Set rng = .Columns("A").Find(recCode)
    
    If rng Is Nothing Then Exit Function
    
    sr = rng.Row
    lr = .Cells(.Rows.Count, 1).End(xlUp).Row
    
    For r = sr To lr
    
        recPcode = .Cells(r, 4)
        recNote = .Cells(r, "H")
        recItem = .Cells(r, "G")
    
        If recPcode = pcode And recNote = "" Then
            
            getPaperItem = recItem
            Exit Function
            
        End If
        
    Next

End With

End Function

Sub KeyInItem_Merge(ByVal coll_L, ByVal coll_E, ByVal recDate As Date, ByVal recCode As String, ByVal mode As Byte) ', ByVal recLoc As String)

With Sheets("日報第一聯")

Call getReportSrEr(sr, er, mode) '取得起始列、終點列

.Range("A" & sr + 2 & ":M" & er - 1).ClearContents

r = sr + 2

For Each it In coll_L

    If r = er Then Stop

    .Range("A" & r) = Split(it, ">")(1)
    .Range("A" & r).WrapText = True
    .Range("C" & r) = getUsedAmountByItem(it, recCode, mode)
    .Range("E" & r) = getSumAmountByItem(it, recDate, mode)
    .Range("A" & r & ":E" & r).Font.Name = "標楷體"

    r = r + 1
    
Next

r = sr + 2

For Each it In coll_E

    If r = er Then Stop

    .Range("H" & r) = Split(it, ">")(1)
    .Range("J" & r) = getUsedAmountByItem(it, recCode, mode + 1)
    .Range("L" & r) = getSumAmountByItem(it, recDate, mode + 1)
    .Range("H" & r & ":L" & r).Font.Name = "標楷體"
    
    r = r + 1
    
Next

End With

End Sub

Sub KeyInItem(ByVal coll_item, ByVal recDate As Date, ByVal recCode As String, ByVal mode As Byte) ', ByVal recLoc As String)

Dim o As New clsPCCES

With Sheets("日報第一聯")

Call getReportSrEr(sr, er, mode) '取得起始列、終點列

.Range("A" & sr + 2 & ":M" & er - 1).ClearContents

r = sr + 2

For Each it In coll_item

    If r = er Then Stop

    .Range("A" & r) = Split(it, ">")(1)
    
    If (Len(.Range("A" & r)) * 1.1) > 16.5 Then
    
    .Range("A" & r).RowHeight = Len(.Range("A" & r)) * 1.1
    
    Else
    
    .Range("A" & r).RowHeight = 16.5
    
    End If
    
    .Range("E" & r) = getUnitByItem(it, mode)
    .Range("F" & r) = getAmountByItem(it, mode)
    .Range("H" & r) = getUsedAmountByItem(it, recCode, mode)
    .Range("J" & r) = getSumAmountByItem(it, recDate, mode)
    .Range("A" & r & ":J" & r).Font.Name = "標楷體"
    
    If mode = 1 Then
    price = o.getMoneyByItemKey(it)
    amount = .Range("J" & r)
    use_money = use_money + price * amount

    End If

    .Range("A" & r).WrapText = True
    '.Range("A" & r).ShrinkToFit = True

    r = r + 1
    
Next

If mode = 1 Then

    progress = use_money / o.getSumMoney
    .Range("K8") = progress

End If

End With

End Sub

Function getSumAmountByItem(ByVal recItem As String, ByVal recDate As Date, ByVal mode As Byte)

With Sheets(dataObj.getDB_name(mode))

lr = .Cells(.Rows.Count, 1).End(xlUp).Row

For r = 2 To lr

    mydate = myFun.tranDate(.Cells(r, 2))
    myItem_key = .Cells(r, 4)
    myitem = .Cells(r, 5)
    myAmount = .Cells(r, 7)
    mynote = .Cells(r, 8)
    
    If mydate <= recDate And mynote = "" Then
    
        If recItem = myItem_key & ">" & myitem Then
        
            getSumAmountByItem = getSumAmountByItem + myAmount
            
        End If
    
    End If

Next

End With

End Function

Function getSumAmountByItem2(ByVal recItem As String, ByVal recDate_s As Date, ByVal recDate_e As Date, ByVal mode As Byte)

With Sheets(dataObj.getDB_name(mode))

lr = .Cells(.Rows.Count, 1).End(xlUp).Row

For r = 2 To lr

    mydate = myFun.tranDate(.Cells(r, 2))
    myItem_key = .Cells(r, 4)
    myitem = .Cells(r, 5)
    myAmount = .Cells(r, 7)
    mynote = .Cells(r, 8)
    
    If mydate <= recDate_e And mydate >= recDate_s And mynote = "" Then
    
        If recItem = myItem_key & ">" & myitem Then
        
            getSumAmountByItem2 = getSumAmountByItem2 + myAmount
            
        End If
    
    End If

Next

End With

End Function

Function getUsedAmountByItemWithDate(ByVal recItem As String, ByVal recDate As Date, ByVal mode As Byte)

With Sheets(dataObj.getDB_name(mode))

lr = .Cells(.Rows.Count, 1).End(xlUp).Row

For r = 2 To lr

    myCode = .Cells(r, 1)
    mydate = myFun.tranDate(.Cells(r, 2))
    myItem_key = .Cells(r, 4)
    myitem = .Cells(r, 5)
    myAmount = .Cells(r, 7)
    mynote = .Cells(r, 8)
    
    If recDate = mydate And mynote = "" Then
    
        If recItem = myItem_key & ">" & myitem Then

            getUsedAmountByItemWithDate = getUsedAmountByItemWithDate + myAmount
            
        End If
    
    End If

Next

End With

End Function

Function getUsedAmountByItem(ByVal recItem As String, ByVal recCode As String, ByVal mode As Byte) As Double ' ByVal recDate As Date)

With Sheets(dataObj.getDB_name(mode))

lr = .Cells(.Rows.Count, 1).End(xlUp).Row

For r = 2 To lr

    myCode = .Cells(r, 1)
    mydate = myFun.tranDate(.Cells(r, 2))
    myItem_key = .Cells(r, 4)
    myitem = .Cells(r, 5)
    myAmount = .Cells(r, 7)
    mynote = .Cells(r, 8)
    
        If recCode = myCode And mynote = "" Then
        
            If recItem = myItem_key & ">" & myitem Then
            
                getUsedAmountByItem = getUsedAmountByItem + myAmount
                
            End If
        
        End If

Next

End With

End Function

Function getUnitByItem(ByVal recItem, ByVal mode As Byte)

shtName = "工料設定"

If mode = 1 Then shtName = "契約詳細表"

With Sheets(shtName)

Set rng = .Cells.Find(what:=recItem)

r = rng.Row

getUnitByItem = .Range("D" & r)

End With

End Function

Function getAmountByItem(ByVal recItem, ByVal mode As Byte)

shtName = "工料設定"

If mode = 1 Then shtName = "契約詳細表"

With Sheets(shtName)

Set rng = .Cells.Find(what:=recItem)

r = rng.Row

getAmountByItem = .Range("E" & r)

End With

End Function

Function getNumByItem(ByVal recItem, ByVal mode As Byte)

shtName = "工料設定"

If mode = 1 Then shtName = "契約詳細表"

With Sheets(shtName)

Set rng = .Cells.Find(what:=recItem)

r = rng.Row

getNumByItem = .Range("B" & r)

End With

End Function

Function getReportSrEr(sr, er, ByVal mode As Byte)

Dim coll As New Collection

'set the target string to find in range("H2:H7") first

arr = ThisWorkbook.Sheets("標案設定").Range("H2:H7")

With ThisWorkbook.Sheets("日報第一聯")

For i = 1 To 6

Set rng = .Cells.Find(what:=arr(i, 1))
coll.Add rng.Row

Next

'Set rng = .Cells.Find(what:="一、依施工計畫書執行按圖施工概況（含約定之重要施工項目及完成數量等）：")
'coll.Add rng.Row
'
'Set rng = .Cells.Find(what:="二、工地材料管理概況（含約定之重要材料使用狀況及數量等）：")
'coll.Add rng.Row
'
'Set rng = .Cells.Find(what:="三、工地人員及機具管理（含約定之出工人數及機具使用情形及數量）：")
'coll.Add rng.Row
'
'Set rng = .Cells.Find(what:="四、本日施工項目是否有須依「營造業專業工程特定施工項目應置之技術士種類、比率或人數標準表」")
'coll.Add rng.Row
'
'Set rng = .Cells.Find(what:="五、工地職業安全衛生事項之督導、公共環境與安全之維護及其他工地行政事務：")
'coll.Add rng.Row
'
'Set rng = .Cells.Find(what:="六、施工取樣試驗紀錄：")
'coll.Add rng.Row

End With

sr = coll(mode)
er = coll(mode + 1)

End Function

Sub checkReportCount(ByVal item_count, ByVal mode As Byte)

With Sheets("日報第一聯")

.Activate

Call getReportSrEr(sr, er, mode) '取得標題之起迄列

total_count = (er - 1) - (sr + 2) + 1

needRow = item_count - total_count

'For j = 1 To total_count - 10
 '   .Rows(15).Delete
'Next

For i = 1 To needRow
    Sheets("日報第一聯").Rows(er - 1 & ":" & er - 1).Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.Insert Shift:=xlDown
Next

'If needRow < 0 And total_count > 11 Then Stop

End With

End Sub

Function getUsedItemByDate(ByVal myCode As String, ByVal mydate As Date, ByVal mode As Byte) ', ByVal mycode As String) '演算先取得項目
    
Dim coll_item As New Collection

With ThisWorkbook.Sheets(dataObj.getDB_name(mode)) '"日報資料庫")

lr = .Cells(.Rows.Count, 1).End(xlUp).Row

    For r = 2 To lr
    
        recCode = .Cells(r, 1)
        recDate = .Cells(r, 2)
        recNum = .Cells(r, 4)
        recItem = .Cells(r, 5)
        recNote = .Cells(r, "H")
        
        'If print_mode = 2 Or print_mode = 3 Then '累計
        If print_mode = 1 Or print_mode = 2 Or print_mode = 3 Then '累計
        
            If myFun.tranDate(recDate) <= mydate And recNote = "" Then
        
                recKey = recNum & ">" & recItem
            
                On Error Resume Next: coll_item.Add recKey, recKey: On Error GoTo 0
        
            End If
            
'        ElseIf print_mode = 1 Then '表單
'
'            If recCode = myCode And recNote = "" Then
'
'                recKey = recNum & ">" & recItem
'
'                On Error Resume Next: coll_item.Add recKey, recKey: On Error GoTo 0
'
'            End If
        
        End If
    
    Next

End With

'Call sortByPCCESRows(coll_item)

If mode <> 5 Then Set getUsedItemByDate = sortByPCCESRows(coll_item, mode) '做排序

End Function

Function sortByPCCESRows(ByVal coll_item, ByVal mode As Byte)

If mode = 1 Then
    shtName = "契約詳細表"
Else
    shtName = "工料設定"
End If
Dim coll As New Collection

With ThisWorkbook.Sheets("TMP")
    
    .Activate
    
    .Visible = True
    
    .Cells.Clear
    
    r = 1
    
    For Each it In coll_item
    
        .Cells(r, 1) = it
        
        Set rng = Sheets(shtName).Cells.Find(it)
        
        If rng Is Nothing Then MsgBox it & "找不到項目!"
        
        .Cells(r, 2) = rng.Row
        
        r = r + 1
    
    Next
    
    .Range("A1:B" & r).Sort key1:=.Columns(2), order1:=xlAscending
    
    For rr = 1 To r - 1
    
        recKey = .Cells(rr, 1)
        coll.Add recKey
        'tmp = Split(.Cells(rr, 1), ">")
        'coll.Add tmp(1)
    
    Next
    
    Set sortByPCCESRows = coll
    
    .Visible = False

End With

End Function

Sub getDataByDate_second(ByVal recDate_s As Date, ByVal recDate_e As Date, ByVal recCode As String)

Dim o As New clsPCCES

If print_mode = 3 Then

    Set coll = getUsedItemByDate(recCode, recDate, 1)

ElseIf print_mode = 4 Then

    Set coll = o.getAllItems

End If

'Dim coll2 As New Collection
'
'For Each it In coll_all
'
'    tmp = Split(it, "-")
'    myDate = tmp(0)
'    mySep = tmp(1)
'
'    If mySep = "1" Then coll2.Add myDate, myDate
'
'Next

Call checkRows(coll.Count)

With ThisWorkbook.Sheets("日報第二聯")

    .Range("K3") = recDate

    r = 8
    
    mode = 1
    
    For Each it In coll
    
        .Range("A" & r) = getNumByItem(it, mode)
        .Range("B" & r) = Split(it, ">")(1)
        .Range("G" & r) = getUnitByItem(it, mode)
        .Range("H" & r) = getAmountByItem(it, mode)
        '.Range("I" & r) = getUsedAmountByItemWithDate(it, recDate, mode)
        .Range("I" & r) = getSumAmountByItem2(it, recDate_s, recDate_e, mode)
        
        price = o.getMoneyByItemKey(it)
        
        .Range("J" & r) = price
        
        amount = .Range("K" & r)
        use_money = use_money + price * amount
        
        .Range("K" & r) = .Range("I" & r) * .Range("J" & r)
        
        r = r + 1

    Next
    
    progress = use_money / o.getSumMoney
    
    Set rng = .Cells.Find("完成百分率(註2)：")
    er = rng.Row
    
    .Range("K" & er) = progress
    
End With

End Sub

Sub checkRows(ByVal item_count)

With ThisWorkbook.Sheets("日報第二聯")

    .Activate

    sr = 8
    
    Set rng = .Cells.Find("完成百分率(註2)：")
    er = rng.Row

    .Range("A" & sr & ":M" & er - 1).ClearContents

    total_count = (er - 1) - (sr + 1) + 1
    
    needRow = item_count - total_count
    
    For i = 1 To needRow
        Sheets("日報第二聯").Rows(er - 1 & ":" & er - 1).Select
        Application.CutCopyMode = False
        Selection.Copy
        Selection.Insert Shift:=xlDown
    Next

End With

End Sub


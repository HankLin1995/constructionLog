VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPCCES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private wbname As String
Private wb As Object
Private myfun As New clsFunction

Private Sub Class_Initialize()

Application.ScreenUpdating = False

End Sub

Sub getFileName() '取得預算書內容

F = Application.GetOpenFilename

If F = "False" Then MsgBox "未取得檔案", vbCritical: End

tmp = Split(F, "\")

wbname = tmp(UBound(tmp))

Workbooks.Open (F)

Set wb = Workbooks(wbname)

End Sub

Sub clearPCCES_data()

With ThisWorkbook.Sheets("契約詳細表")

lr = .Cells(.Rows.Count, 1).End(xlUp).Row

If lr > 1 Then
    .Activate
    msg = MsgBox("你有舊資料存在，確定要刪除?", vbCritical + vbYesNo)
    
    If msg = vbYes Then
        .Range("A2").Resize(lr, 7).Clear
    Else
        MsgBox "取消操作", vbInformation
        wb.Close False
        End
    End If

End If

End With

End Sub

Function getShtContentName()

For Each sht In wb.Sheets

    If sht.Name Like "*詳細表" Then
    
        getShtContentName = sht.Name
    
    End If

Next

End Function

Sub getAllContents()

With wb.Sheets(getShtContentName)

    lr = .Cells(.Rows.Count, "F").End(xlUp).Row
    
    For r = 9 To lr
    
        item_num = myfun.tranCharcter_NUM(.Cells(r, 1))
        item_cont = .Cells(r, 2)
        item_unit = .Cells(r, 3)
        item_amount = .Cells(r, 4)
        item_cost = .Cells(r, 5)
        
        If item_num <> "" And item_amount <> 0 Then
        
            item_cont = getFullName(r)
        
            Call getData(item_num, item_cont, item_unit, item_amount, item_cost)
        
        End If
    
    Next

End With

MsgBox "取得PCCES預算書完成，請進行部分項目微調", vbInformation

wb.Close False

End Sub

Private Function getFullName(ByVal r As Integer)

Set wb = Workbooks(wbname)

With wb.Sheets(getShtContentName)

    If .Cells(r + 1, "F") <> "" Then
    
        getFullName = myfun.tranCharcter(.Cells(r, "B"))
        
    Else
    
        lr = .Cells(r, "F").End(xlDown).Row
        
        For rr = r To lr - 1
        
            s = s & myfun.tranCharcter(.Cells(rr, "B"))
            cnt = cnt + 1
        
        Next
    
        getFullName = s
        
    End If

End With

End Function

Private Sub getData(ByVal num As String, ByVal cont As String, ByVal unit As String, ByVal amount As String, ByVal cost As String)

    With ThisWorkbook.Sheets("契約詳細表")
    
        lr = .Cells(.Rows.Count, 1).End(xlUp).Row
        
        r = lr + 1
        
        .Cells(r, 2) = num
        .Cells(r, 3) = cont
        .Cells(r, 4) = unit
        .Cells(r, 5) = amount
        .Cells(r, 6) = cost
        
        .Cells(r, 1) = "=RC[1]&"">""&RC[2]" ' num & ">" & cont
        .Cells(r, 1).WrapText = False
        
        If unit = "" Then .Range("A" & r & ":F" & r).Interior.ColorIndex = 22
    
    End With

End Sub

Sub setValidation()

With Sheets("契約詳細表")

lr = .Cells(.Rows.Count, 1).End(xlUp).Row

'Debug.Print lr

End With

With Sheets("日報填寫")

.Unprotect

For r = 7 To 7 + 100 '只有取到第100欄

    .Range("A" & r).Validation.Delete

    .Range("A" & r).Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
Formula1:="=契約詳細表!$A" & 2 & ":$A" & lr

Next

.Protect
.Activate

End With

End Sub

Sub RefreshDB()

With Sheets("長表格式")

lr = .Cells(.Rows.Count, 1).End(xlUp).Row
lc = .Cells(1, .Columns.Count).End(xlToLeft).Column

.Range("A4").Resize(lr, lc).ClearContents
.Range("F1").Resize(3, lc).ClearContents

End With

With Sheets("範本資料庫")

lr = .Cells(.Rows.Count, 1).End(xlUp).Row
.Range("A2").Resize(lr, 8).ClearContents

End With

arr = Array("日報", "材料", "人員", "機具", "文件")
'arr = Array("材料", "人員", "機具", "文件")

For i = LBound(arr) To UBound(arr)

    With Sheets(arr(i) & "資料庫")
    
        .Unprotect
        
        lr = .Cells(.Rows.Count, 1).End(xlUp).Row
        .Range("A2").Resize(lr, 8).ClearContents
        
        .Protect

    End With

Next

End Sub

Function getSumMoney()

With Sheets("契約詳細表")

lr = .Cells(.Rows.Count, 1).End(xlUp).Row

Set rng = .Rows("1").Find("數量")

c = rng.Column

For r = 2 To lr
    
    amount = .Cells(r, c)
    price = .Cells(r, c + 1)
    note = .Cells(r, c + 2)
    
    If note = "" Then getSumMoney = getSumMoney + amount * price

Next

End With

End Function

Function getMoneyByItemKey(ByVal item_key)

With Sheets("契約詳細表")

r = .Columns("A").Find(item_key).Row
c = .Rows("1").Find("數量").Column
    
getMoneyByItemKey = .Cells(r, c + 1)

End With

End Function

Function getAllItems()

Dim coll As New Collection

With Sheets("契約詳細表")

lr = .Cells(.Rows.Count, 1).End(xlUp).Row

For r = 2 To lr
    
    item = .Cells(r, 1)
    note = .Cells(r, "G")
    If note = "" Then coll.Add item, item

Next

End With

Set getAllItems = coll

End Function

Sub checkIsRepeat()

Dim coll As New Collection

With Sheets("契約詳細表")

lr = .Cells(.Rows.Count, 1).End(xlUp).Row

For r = 2 To lr

    mykey = .Cells(r, 1)
    On Error GoTo ERRORHANDLE
    coll.Add mykey, mykey

Next

If s <> "" Then
    MsgBox s, vbCritical
    .Range("A2").Resize(lr, 7).Clear
Else
    Debug.Print "契約詳細表皆為唯一值!"
End If

Exit Sub

End With

ERRORHANDLE:       ' 錯誤處理用的程式碼
  s = s & "列數=" & r & ",項目=" & mykey & "已經被使用!請修正!" & vbNewLine
  Resume Next

End Sub

Function checkCollection(ByRef coll, ByVal coll_key As String)

checkCollection = coll(coll_key)

End Function

Private Sub Class_Terminate()

Application.ScreenUpdating = True

End Sub

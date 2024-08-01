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
Private myFun As New clsFunction

Private Sub Class_Initialize()

Application.ScreenUpdating = False

End Sub

Sub getFileName() '取得預算書內容

f = Application.GetOpenFilename

If f = "False" Then MsgBox "未取得檔案", vbCritical: End

tmp = Split(f, "\")

wbname = tmp(UBound(tmp))

Workbooks.Open (f)

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
    
        item_num = myFun.tranCharcter_NUM(.Cells(r, 1))
        item_cont = .Cells(r, 2)
        item_unit = .Cells(r, 3)
        item_amount = .Cells(r, 4)
        item_cost = .Cells(r, 5)
        
        '列出所有項目for第三階段調整
        
        If item_num <> "" Then 'And item_amount <> 0 Then
        
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
    
        getFullName = myFun.tranCharcter(.Cells(r, "B"))
        
    Else
    
        lr = .Cells(r, "F").End(xlDown).Row
        
        For rr = r To lr - 1
        
            If .Cells(rr, "F") <> "" Then
        
            s = s & myFun.tranCharcter(.Cells(rr, "B"))
            cnt = cnt + 1
            
            End If
        
        Next
    
        If s = "" Then s = myFun.tranCharcter(.Cells(r, "B"))
    
        getFullName = s
        
    End If

End With

End Function

Private Sub getData(ByVal num As String, ByVal cont As String, ByVal unit As String, ByVal amount As String, ByVal cost As String)

    With ThisWorkbook.Sheets("契約詳細表")
    
        lr = .Cells(.Rows.Count, 1).End(xlUp).Row
        
        r = lr + 1
        
        .Cells(r, 2) = "'" & num
        .Cells(r, 3) = cont
        .Cells(r, 4) = unit
        .Cells(r, 5) = amount
        .Cells(r, 6) = cost
        
        .Cells(r, 1) = "=RC[1]&"">""&RC[2]" ' num & ">" & cont
        .Cells(r, 1).WrapText = False
        
        If unit = "" Then .Range("A" & r & ":F" & r).Interior.ColorIndex = 22
    
    End With

End Sub

Sub setValidation2()

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

Sub setValidation()

    Dim customRange As String
    Dim wsContract As Worksheet
    Dim wsValidation As Worksheet
    Dim customRangeArray() As String
    Dim i As Long
    
    ' 獲取匹配的範圍
    customRange = CompareContractAndActualQuantities()
    
    ' 設定工作表變數
    Set wsContract = Sheets("契約詳細表")
    ' 你可以選擇將值放在一個新的工作表中或在契約詳細表中的空白區域
    Set wsValidation = Sheets("標案設定") ' 例如，可以將匹配值放在契約詳細表的某一空白列中
    
    wsValidation.Range("K1:K" & wsValidation.Cells(1, 1).End(xlDown).Row).ClearContents
    
    ' 將 customRange 拆分成單個地址
    customRangeArray = Split(customRange, ",")
    
    ' 將這些值放在 wsValidation 的某個空白區域，例如列 K
    For i = LBound(customRangeArray) To UBound(customRangeArray)
        wsValidation.Range("K" & i + 1).Value = wsContract.Range(customRangeArray(i)).Value
    Next i
    
    With Sheets("日報填寫")
    
        .Unprotect
    
        For r = 7 To 7 + 100 '只有取到第100欄
    
            .Range("A" & r).Validation.Delete
    
            .Range("A" & r).Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                Formula1:="=標案設定!$K$1:$K$" & UBound(customRangeArray) + 1
    
        Next
    
        .Protect
        .Activate
    
    End With

End Sub


Function CompareContractAndActualQuantities()

    Dim wsData As Worksheet
    Dim wsContract As Worksheet
    Dim dict As Object
    Dim contractDict As Object
    Dim rng As Range
    Dim cell As Range
    Dim key As Variant
    Dim i As Long
    Dim actualQty As Double
    Dim contractQty As Double
    Dim matchRanges As String
    
    ' 設定工作表變數
    Set wsData = Sheets("日報資料庫") ' 請根據你的實際工作表名稱調整
    Set wsContract = Sheets("契約詳細表") ' 請根據你的實際工作表名稱調整
    
    ' 建立 Dictionary 物件來儲存累積數量
    Set dict = CreateObject("Scripting.Dictionary")
    Set contractDict = CreateObject("Scripting.Dictionary")
    
    ' 迭代契約詳細表範圍，記錄契約數量
    For i = 2 To wsContract.Cells(wsContract.Rows.Count, "B").End(xlUp).Row
        key = wsContract.Cells(i, "B").Value
        contractQty = wsContract.Cells(i, "E").Value
        contractDict(key) = contractQty
    Next i
    
    ' 迭代日報資料庫範圍，統計 D 欄位的累積數量
    For i = 2 To wsData.Cells(wsData.Rows.Count, "D").End(xlUp).Row
        key = wsData.Cells(i, "D").Value
        If wsData.Cells(i, "H").Value = "" Then ' 檢查 H 欄位是否為空白
            If Not dict.exists(key) Then
                dict.Add key, wsData.Cells(i, "G").Value
            Else
                dict(key) = dict(key) + wsData.Cells(i, "G").Value
            End If
        End If
    Next i
    
    ' 比較契約數量與實作數量
    matchRanges = ""
    For i = 2 To wsContract.Cells(wsContract.Rows.Count, "B").End(xlUp).Row
        key = wsContract.Cells(i, "B").Value

        If dict.exists(key) Then
            actualQty = dict(key)
        Else
            actualQty = 0
        End If
        contractQty = contractDict(key)
        If contractQty > actualQty Then
            Debug.Print key & ">" & wsContract.Cells(i, "C").Value & ":" & contractQty & ">" & actualQty
            If matchRanges = "" Then
                matchRanges = wsContract.Cells(i, "A").Address
            Else
                matchRanges = matchRanges & "," & wsContract.Cells(i, "A").Address
            End If
        End If
    Next i
    
    ' 輸出匹配的範圍
    Debug.Print "匹配的範圍: " & matchRanges
    CompareContractAndActualQuantities = matchRanges
    
    ' 清理
    Set dict = Nothing
    Set contractDict = Nothing
    Set wsData = Nothing
    Set wsContract = Nothing

End Function

Sub RefreshDB()

With Sheets("長表格式")

.Unprotect

lr = .Cells(.Rows.Count, 1).End(xlUp).Row
lc = .Cells(1, .Columns.Count).End(xlToLeft).Column

.Range("A4").Resize(lr, lc).ClearContents
.Range("F1").Resize(3, lc).ClearContents

.Protect

End With

With Sheets("範本資料庫")

.Unprotect

lr = .Cells(.Rows.Count, 1).End(xlUp).Row
.Range("A2").Resize(lr, 8).ClearContents

.Protect

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

With Sheets("標案設定")

    lr = .Cells(.Rows.Count, "F").End(xlUp).Row

    For r = 2 To lr

        getSumMoney = getSumMoney - getContractMoneyByKeyItem(.Cells(r, "F"))
    
    Next

End With

End Function


Function getContractMoneyByKeyItem(ByVal s As String) As Double

With Sheets("契約詳細表")

Set rng = .Columns("A").Find(s)

If rng Is Nothing Then
    
    MsgBox "契約詳細表中找不到" & s & "!"

Else

    num = rng.Offset(0, 4)
    cost = rng.Offset(0, 5)
    getContractMoneyByKeyItem = num * cost

End If

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

With ThisWorkbook.Sheets("契約詳細表")

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

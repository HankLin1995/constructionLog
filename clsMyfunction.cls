VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMyfunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Function getBlankColl(ByVal sht As Object, ByVal col As Integer)

'20200729
'總表的線分段資料
'橫斷面的分段資料

Dim collBlank As New Collection

collBlank.Add 1

With sht

    lr = .Cells(.Rows.Count, col).End(xlUp).Row

    For r = 2 To lr
    
        If .Cells(r, col) = "" Then collBlank.Add r
    
    Next
    
    collBlank.Add lr + 1
    
End With

Set getBlankColl = collBlank

End Function

Sub ClearData(ByVal sht As Object, ByVal sr As Integer, ByVal col As Integer, ByVal xrange As Integer)

With sht

    lr = .Cells(.Rows.Count, col).End(xlUp).Row
    .Cells(sr, col).Resize(lr, xrange).ClearContents

End With

End Sub

Public Function BubbleSort_array(arr As Variant, Optional ByVal method As String = "Ascending")

'https://officeguide.cc/excel-vba-array-bubble-sort-function-tutorial-examples/
'https://stackoverflow.com/questions/3587662/how-do-i-sort-a-collection
'冒泡排序僅供少數資料使用(data<70)

    idxMin = LBound(arr)
    idxMax = UBound(arr)
  
    For i = idxMin To idxMax - 1
      For j = i + 1 To idxMax
      
        If method = "Ascending" Then '升排序
      
            If arr(i) > arr(j) Then ' 判斷條件
              temp = arr(i): arr(i) = arr(j): arr(j) = temp
            End If
        
        Else
        
            If arr(i) < arr(j) Then ' 判斷條件
              temp = arr(i): arr(i) = arr(j): arr(j) = temp
            End If
        
        End If
        
      Next j
    Next i
    
    BubbleSort_array = arr

End Function

Public Function BubbleSort_coll(ByVal coll As Collection, Optional ByVal method As String = "Ascending")

'rely on=>>
'   tranColl2Array
'   tranArray2Coll

    arr = tranColl2Array(coll)
    arr_sort = BubbleSort_array(arr, method)
    
    Set BubbleSort_coll = tranArray2Coll(arr_sort)

End Function

Public Function tranColls2Array(ParamArray colls() As Variant) ''ByVal colls As Collection)

'每一集合儲存成一維度

Dim arr()
Dim k As Integer

ReDim arr(0 To colls(0).Count - 1, UBound(colls))

M = 0

For Each coll In colls

    k = 0
    
    For Each it In coll
    
        arr(k, M) = it
        k = k + 1
    
    Next

    M = M + 1
    
Next

tranColls2Array = arr

End Function

Public Function tranColls2OneArray(ParamArray colls() As Variant) As Double()

'各集合依序儲存成同一維度 (常見於XY座標集合轉換成Vertices)

Dim arr() As Double
ReDim arr(0 To colls(0).Count * (UBound(colls) + 1) - 1)

For i = 1 To colls(0).Count

    For j = 0 To UBound(colls)
    
        arr(k) = colls(j)(i)
        k = k + 1
    
    Next

Next

tranColls2OneArray = arr

End Function

Public Function tranColl2Array(ByVal coll As Collection, Optional lboundcount As Integer = 0) As Variant

Dim arr()
Dim k As Integer

ReDim arr(0 + lboundcount To coll.Count - 1 + lboundcount)

For i = 1 To coll.Count
    arr(0 + lboundcount + k) = coll(i)
    k = k + 1
Next

tranColl2Array = arr

End Function

Public Function tranArray2Coll(ByVal arr As Variant) As Collection

Dim coll As New Collection

For i = LBound(arr) To UBound(arr)
    coll.Add arr(i)
Next

Set tranArray2Coll = coll

End Function

Public Sub showList(ByVal Data As Variant)

cnt = 1
If Not TypeName(Data) = "Collection" Then cnt = LBound(Data)

For Each it In Data
     Debug.Print cnt & " => " & it: cnt = cnt + 1
Next

End Sub

Function MergeSort_array(ByRef arr As Variant, Optional ByRef sortKeys As Variant = Null, Optional ByVal hasHeaders As Boolean = False)
'Sorts a one or two dimensional array.
'2 dimensional arrays can have their sort keys specified by passing
'the appropriate column number(s) as the sortKeys parameter.
'Function passes a reference so will mutate your original array.
'If this is not desirable you must pass a copy.
'
'Example uses:
' sortArray myArray - One-dimensional array
' sortArray myArray, 2 - Two-dimensional array, single sort key
' sortArray myArray, Array(2,3,1) - Two-dimensional array, multiple sort keys
' sortArray myArray, Array(2,3,1), True - Two-dimensional array, multiple sort keys with headers preserved

'https://www.ozgrid.com/forum/index.php?thread/1220720-vba-merge-sort/

Dim mid As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim X As Long
Dim Y As Long
Dim sortMode As Long
Dim arr1
Dim arr2
Dim head
Dim tmp

If UBound(arr) - LBound(arr) = 0 Then Exit Function

On Error Resume Next

i = UBound(arr, 2)
If Err.Number <> 0 Then
    sortMode = 1 'Not a 2D array
    If hasHeaders Then
        ReDim tmp(LBound(arr) To UBound(arr) - 1)
        ReDim head(1 To 1)
        For i = LBound(arr) To UBound(arr)
            If i = LBound(arr) Then
                head(1) = arr(LBound(arr))
            Else
                tmp(i - 1) = arr(i)
            End If
            
        Next i
        arr = tmp
    End If
Else
    sortMode = 2
    If hasHeaders Then
        ReDim tmp(LBound(arr) To (UBound(arr) - 1), LBound(arr, 2) To UBound(arr, 2))
        ReDim head(1 To 1, LBound(arr, 2) To UBound(arr, 2))
        For i = LBound(arr) To UBound(arr)
            For j = LBound(arr, 2) To UBound(arr, 2)
                If i = LBound(arr) Then
                    head(1, j) = arr(LBound(arr), j)
                Else
                    tmp(i - 1, j) = arr(i, j)
                End If
            Next j
        Next i
        arr = tmp
    End If
End If

On Error GoTo 0

If IsNumeric(sortKeys) Then
    sortKeys = Array(CLng(sortKeys))
ElseIf IsNull(sortKeys) Then
    sortKeys = Array(LBound(arr))
End If

Y = LBound(sortKeys)
mid = Int((UBound(arr) + IIf(LBound(arr) = 0, 1, 0)) / 2)

If mid < LBound(arr) Then mid = LBound(arr)

If sortMode = 1 Then

    ReDim arr1(LBound(arr) To mid - IIf(LBound(arr) = 0, 1, 0))
    ReDim arr2(LBound(arr) To UBound(arr) - mid)
    j = LBound(arr)
    For i = LBound(arr1) To UBound(arr1)
        arr1(i) = arr(j)
        j = j + 1
    Next i
    For i = LBound(arr2) To UBound(arr2)
        arr2(i) = arr(j)
        j = j + 1
    Next i
ElseIf sortMode = 2 Then

    ReDim arr1(LBound(arr) To mid - IIf(LBound(arr) = 0, 1, 0), LBound(arr, 2) To UBound(arr, 2))
    ReDim arr2(LBound(arr) To UBound(arr) - mid, LBound(arr, 2) To UBound(arr, 2))
    j = LBound(arr)
    For i = LBound(arr1) To UBound(arr1)
        For k = LBound(arr1, 2) To UBound(arr1, 2)
            arr1(i, k) = arr(j, k)
        Next k
        j = j + 1
    Next i
    
    For i = LBound(arr2) To UBound(arr2)
        For k = LBound(arr2, 2) To UBound(arr2, 2)
            arr2(i, k) = arr(j, k)
        Next k
     j = j + 1
    Next i
End If

MergeSort_array arr1, sortKeys 'I call myself!!!
MergeSort_array arr2, sortKeys 'I call myself again!!!
i = LBound(arr)
j = LBound(arr1)
k = LBound(arr2)

If sortMode = 1 Then
    While j <= UBound(arr1) And k <= UBound(arr2)
        If arr1(j) <= arr2(k) Then
        arr(i) = arr1(j)
        j = j + 1
        Else
        arr(i) = arr2(k)
        k = k + 1
        End If
        i = i + 1
    Wend
    While j <= UBound(arr1)
        arr(i) = arr1(j)
        j = j + 1
        i = i + 1
    Wend
    While k <= UBound(arr2)
        arr(i) = arr2(k)
        k = k + 1
        i = i + 1
    Wend
ElseIf sortMode = 2 Then
While j <= UBound(arr1) And k <= UBound(arr2)
    If arr1(j, sortKeys(Y)) < arr2(k, sortKeys(Y)) _
    Or (arr1(j, sortKeys(Y)) = arr2(k, sortKeys(Y)) And UBound(sortKeys) = Y) Then
    For X = LBound(arr1, 2) To UBound(arr1, 2)
        arr(i, X) = arr1(j, X)
        Next X
        j = j + 1
        Y = LBound(sortKeys)
    ElseIf arr1(j, sortKeys(Y)) > arr2(k, sortKeys(Y)) Then
        For X = LBound(arr2, 2) To UBound(arr2, 2)
            arr(i, X) = arr2(k, X)
        Next X
        k = k + 1
        Y = LBound(sortKeys)
    Else
        i = i - 1
         Y = Y + 1
    End If
    i = i + 1
Wend
While j <= UBound(arr1)
    For X = LBound(arr1, 2) To UBound(arr1, 2)
    arr(i, X) = arr1(j, X)
    Next X
    j = j + 1
    i = i + 1
Wend
While k <= UBound(arr2)
    For X = LBound(arr2, 2) To UBound(arr2, 2)
    arr(i, X) = arr2(k, X)
    Next X
    k = k + 1
    i = i + 1
Wend
End If
If hasHeaders Then
    If sortMode = 1 Then
        '1d
        ReDim tmp(LBound(tmp) To UBound(tmp) + 1)
        tmp(LBound(tmp)) = head(1)
        For i = LBound(arr) To UBound(arr)
            tmp(i + 1) = arr(i)
        Next i
    Else
        '2d
        ReDim tmp(LBound(tmp) To UBound(tmp) + 1, LBound(tmp, 2) To UBound(tmp, 2))
        For i = LBound(tmp) To UBound(tmp)
            For j = LBound(tmp, 2) To UBound(tmp, 2)
            If i = LBound(tmp) Then
                tmp(i, j) = head(1, j)
            Else
                tmp(i, j) = arr(i - 1, j)
            End If
            Next
        Next i
    End If
    arr = tmp
End If
End Function

Public Function ReverseColl(ByVal coll)

Dim coll_tmp As New Collection

For i = coll.Count To 1 Step -1

    coll_tmp.Add coll(i)

Next

Set ReverseColl = coll_tmp

End Function

Public Function combineArray(arr1 As Variant, arr2 As Variant)

Dim arr()

ReDim arr(LBound(arr1) To UBound(arr1) + UBound(arr2) + 1)

For Each a In arr1
    
    arr(cnt) = a
    cnt = cnt + 1
Next

For Each a In arr2

    arr(cnt) = a
    cnt = cnt + 1

Next

combineArray = arr

End Function

Function SortPTArray(ByVal arr)

Dim collX As New Collection
Dim collY As New Collection

For i = LBound(arr) To UBound(arr) Step 3

    collX.Add arr(i)
    collY.Add arr(i + 1)

Next

SortPTArray = tranColls2Array(collX, collY)

Call MergeSort_array(SortPTArray)

End Function


'Public Function getItemFromPTArray(ByVal arr, ByVal j As Byte, ByVal mode As String)
'
'Dim coll As New Collection
'
'For i = LBound(arr) To UBound(arr) Step 3
'
'    coll.Add arr(i + j)
'
'Next
'
'Set sort_coll = BubbleSort_coll(coll)
'
'If mode = "MAX" Then
'    getItemFromPTArray = sort_coll(sort_coll.Count)
'Else
'    getItemFromPTArray = sort_coll(1)
'End If
'
'End Function

Function tranDate(ByVal mydatestr) As Date

Y = mid(mydatestr, 1, 3)
M = mid(mydatestr, 4, 2)
d = mid(mydatestr, 6, 2)

tranDate = CDate(1911 + Y & "/" & M & "/" & d)

End Function

Function tranCharcter(ByVal s As String)

If s Like "*~*" Then

    For i = 1 To Len(s)
    
        mych = mid(s, i, 1)
        If mych = "~" Then mych = "-"
        new_ch = new_ch & mych
    
    Next
    
    tranCharcter = new_ch
    
Else

    tranCharcter = s

End If

End Function

Function tranCharcter_NUM(ByVal s As String)

s = Replace(s, "(", "[")
s = Replace(s, ")", "]")

tranCharcter_NUM = s

End Function

Function getRowsByUser(ByVal shtName As String, ByVal col_eng As String, ByVal key_value)
    
    Dim coll As New Collection
    Dim firstAddress As String
    
    With Sheets(shtName).Columns(col_eng)
        Set c = .Find(key_value, LookIn:=xlValues, LookAt:=xlWhole)
        If Not c Is Nothing Then
            firstAddress = c.Address
            Do
                coll.Add c.Row
                Set c = .FindNext(c)
                
            Loop Until c.Address = firstAddress
        End If
    End With
    
    Set getRowsByUser = coll
    
End Function

Function getRowsByUser2(ByVal shtName As String, ByVal key_value, ByVal row_head As Integer, ByVal col_value As String)
    
    Dim coll As New Collection
    Dim firstAddress As String
    
    Set rng = Sheets(shtName).Rows(row_head).Find(col_value)
    col_eng = ConvertToLetter(rng.Column)
        
    With Sheets(shtName).Columns(col_eng)
    
        Set c = .Find(key_value, LookIn:=xlValues, LookAt:=xlWhole)
        If Not c Is Nothing Then
            firstAddress = c.Address
            Do
                coll.Add c.Row
                Set c = .FindNext(c)
                
            Loop Until c.Address = firstAddress
        End If
    End With
    
    Set getRowsByUser2 = coll
    
End Function

Function getUniqueItems(ByVal shtName As String, ByVal sr As Integer, Optional ByVal find_col_eng As String, Optional find_col_keyValue As String = "")

Dim coll As New Collection

With Sheets(shtName)

    If find_col_eng = "" Then
        
        Set rng = .Rows(sr - 1).Find(find_col_keyValue)
        find_col_eng = ConvertToLetter(rng.Column)
        
    End If

    lr = .Cells(.Rows.Count, find_col_eng).End(xlUp).Row

    For r = sr To lr
    
        s = .Cells(r, find_col_eng)
    
        On Error Resume Next
        
        If s <> "" Then coll.Add CStr(s), CStr(s)
        
        On Error GoTo 0
    
    Next

End With

Set getUniqueItems = coll

End Function

Function getUniqueItemsToLastRow(ByVal shtName As String, ByVal sr As Integer, ByVal lr As Integer, Optional ByVal find_col_eng As String, Optional find_col_keyValue As String = "")

Dim coll As New Collection

With Sheets(shtName)

    If find_col_eng = "" Then
        
        Set rng = .Rows(sr - 1).Find(find_col_keyValue)
        find_col_eng = ConvertToLetter(rng.Column)
        
    End If

   ' lr = .Cells(.Rows.count, find_col_eng).End(xlUp).Row

    For r = sr To lr
    
        s = .Cells(r, find_col_eng)
    
        On Error Resume Next
        
        If s <> "" Then coll.Add CStr(s), CStr(s)
        
        On Error GoTo 0
    
    Next

End With

Set getUniqueItemsToLastRow = coll

End Function

Function getUniqueItemsInCollRows(ByVal shtName As String, ByVal coll_rows, ByVal row_head As Integer, Optional ByVal find_col_eng As String, Optional find_col_keyValue As String = "")

Dim coll As New Collection

With Sheets(shtName)

    If find_col_eng = "" Then
        
        Set rng = .Rows(row_head).Find(find_col_keyValue)
        find_col_eng = ConvertToLetter(rng.Column)
        
    End If

    For Each r In coll_rows
    
        s = .Cells(r, find_col_eng)
    
        On Error Resume Next
        
        If s <> "" Then coll.Add CStr(s), CStr(s)
        
        On Error GoTo 0
    
    Next

End With

Set getUniqueItemsInCollRows = coll

End Function

Sub AppendData(ByVal shtName As String, ByVal arr) ', Optional ByVal sc As Integer = 1)

With Sheets(shtName)

lr = .Cells(.Rows.Count, 1).End(xlUp).Row

r = lr + 1

For i = LBound(arr) To UBound(arr)

    .Cells(r, 1 + i) = arr(i)

Next

End With

End Sub

Function ch(ByVal i)

arr = Array("一", "二", "三", "四", "五", "六", "七", "八", "九") ', "參", "肆", "伍", "陸")

ch = arr(i - 1)

End Function

Function ConvertToLetter(iCol As Long) As String
   Dim a As Long
   Dim b As Long
   a = iCol
   ConvertToLetter = ""
   Do While iCol > 0
      a = Int((iCol - 1) / 26)
      b = (iCol - 1) Mod 26
      ConvertToLetter = Chr(b + 65) & ConvertToLetter
      iCol = a
   Loop
End Function

Function IsNumericWithPlusAndParentheses(inputString As String) As Boolean
    Dim i As Integer
    Dim char As String
    Dim validChars As String
    validChars = "0123456789+()上下.~、"
    IsNumericWithPlusAndParentheses = True ' 預設為符合條件

    For i = 1 To Len(inputString)
        char = mid(inputString, i, 1)
        If InStr(validChars, char) = 0 Then
            IsNumericWithPlusAndParentheses = False
            Exit Function ' 出現非法字符，退出循環
        End If
    Next i
End Function

Function TranLoc(ByVal Data As String) As Double

'樁號型態轉成可計算之樁號

tmp = split(Data, "+")

If UBound(tmp) = -1 Or Data = "" Then Exit Function ' TranLoc = CDbl(Data): Exit Function

tloc = tmp(0) '千位數
dloc = tmp(1)


If dloc Like "*(*" Then

    tmp2 = split(dloc, "(")

    If tmp2(0) Like "*.*" Then

        tmp3 = split(tmp2(0), ".")
        dloc = tmp3(0) + tmp3(1) / 10
    
    Else
    
        dloc = tmp2(0)
    
    End If
    
    If dloc > 1000 Then Exit Function
    
End If

For i = 1 To Len(tloc)

    loc_ch = mid(tloc, i, 1)
    If IsNumeric(loc_ch) Then ref = ref & loc_ch

Next

TranLoc = CDbl(ref) * 1000 + CDbl(dloc)
    
End Function

Function SplitAllLocs(ByVal rec_loc As String) '區間

Dim coll As New Collection

If rec_loc Like "*~*" Then

    tmp = split(rec_loc, "、")

    For i = LBound(tmp) To UBound(tmp)
    
        coll.Add tmp(i)
    
        loc_tmp = split(tmp(i), "~")

        For j = LBound(loc_tmp) To UBound(loc_tmp)

            my_loc = loc_tmp(j)
            
            If Not my_loc Like "0+000*" And TranLoc(my_loc) = 0 Then
            
                msg = MsgBox("【" & my_loc & "】樁號無法轉換...是否要繼續運行?", vbCritical + vbYesNo)
            
                If msg = vbNo Then End
            
            End If

        Next
        
    Next

End If

Set SplitAllLocs = coll

End Function

Function IsRecLocPass(ByVal rec_loc As String, ByVal item_loc As String) ', ByVal r As Integer)

IsRecLocPass = False

Dim myFunc As New clsMyfunction

tmp = split(rec_loc, "~")

rec_sloc = myFunc.TranLoc(tmp(0))
rec_eloc = myFunc.TranLoc(tmp(1))

tmp2 = split(item_loc, "~")

item_sloc = myFunc.TranLoc(tmp2(0))
item_eloc = myFunc.TranLoc(tmp2(1))

'If rec_sloc = 0 Then
'    err_prompt = "表單起點【" & tmp(0) & "】請再檢查。": Exit Function
'ElseIf rec_eloc = 0 Then
'    err_prompt = "表單起點【" & tmp(1) & "】請再檢查。": Exit Function
'ElseIf item_sloc = 0 Then
'    err_prompt = "紀錄起點【" & tmp2(0) & "】請再檢查。": Exit Function
'ElseIf item_eloc = 0 Then
'    err_prompt = "紀錄起點【" & tmp2(1) & "】請再檢查。": Exit Function
'End If


If item_sloc >= rec_sloc And item_sloc < rec_eloc Then

    err_prompt = "第" & r & "列衝突=>紀錄起點【" & tmp2(0) & "】已包含於本次填報【" & rec_loc & "】"

    If err_prompt <> "" Then p = p & err_prompt & vbNewLine

End If

If item_eloc > rec_sloc And item_eloc <= rec_eloc Then

    err_prompt = "第" & r & "列衝突=>紀錄終點【" & tmp2(1) & "】已包含於本次填報【" & rec_loc & "】"

    If err_prompt <> "" Then p = p & err_prompt & vbNewLine

End If

If rec_sloc >= item_sloc And rec_sloc < item_eloc Then

    err_prompt = "第" & r & "列衝突=>填報起點【" & tmp(0) & "】已包含於舊有紀錄【" & item_loc & "】"

    If err_prompt <> "" Then p = p & err_prompt & vbNewLine

End If

If rec_eloc > item_sloc And rec_eloc <= item_eloc Then

    err_prompt = "第" & r & "列衝突=>填報終點【" & tmp(1) & "】已包含於舊有紀錄【" & item_loc & "】"

    If err_prompt <> "" Then p = p & err_prompt & vbNewLine

End If

If p = "" Then IsRecLocPass = True

End Function

Function splitFileName_Check(ByVal filename As String, ByRef filename_ch, ByRef filename_eng)
    
    pt2 = InStrRev(filename, "[")

    filename_ch = mid(filename, 1, pt2 - 1)
    filename_eng = mid(filename, pt2 + 1, Len(filename) - pt2 - 1)

End Function

Function FileOpen(initialFilename As String, _
  Optional sDesc As String = "Excel (*.xls)", _
  Optional sFilter As String = "*.xls") As String
  With Application.FileDialog(msoFileDialogOpen)
    .ButtonName = "&Open"
    .initialFilename = initialFilename
    .Filters.Clear
    .Filters.Add sDesc, sFilter, 1
    .Title = "File Select"
    .AllowMultiSelect = False
    If .Show = -1 Then FileOpen = .SelectedItems(1)
  End With
End Function







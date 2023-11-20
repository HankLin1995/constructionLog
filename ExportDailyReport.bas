Attribute VB_Name = "ExportDailyReport"
'todo:
'1.取得該日內容
'2.依照該日內容組成日報數量格式

'by result

Sub cmdExportToDayReports()

'getResultWorkbook
Set wb = getResultWorkbook() 'ThisWorkbook.Path & "\第二階段\results.xls") ' Workbooks("Results.xls")

Set wb_new = Workbooks.Add

For Each sht In wb.Sheets

    If sht.Name Like "*-*" Then

    With sht
        
        rec_code = .Range("B2")
        rec_money = .Range("N1")
        rec_date = .Range("K3")
        weather_u = .Range("C3")
        weather_d = .Range("E3")
        con_name = .Range("D4")
        work_day = .Range("B5")
        work_day_extend = .Range("L5")
        work_day_start = .Range("D6")
        work_day_end = .Range("K6")
        pgs_design = .Range("D7")
        pgs_real = .Range("K7")
        items_str = getRecItemString(sht)
        
        r_data = .Cells.Find("六、施工取樣試驗紀錄：").Row
        
        test_str = .Range("E" & r_data) '75
        
        If Not .Range("E" & r_data).Comment Is Nothing Then
        
            test_to_second_str = .Range("E" & r_data).Comment.Text
            
            tmp = Split(test_to_second_str, ";")
            test_to_second_str = tmp(0)
            tmp2 = Split(tmp(1), "$")
            
            For i = LBound(tmp2) To UBound(tmp2)
            
                test_str = Replace(test_str, tmp2(i), "")
            
            Next
            
        Else
        
            test_to_second_str = ""
        
        End If
        
'
'        For i = LBound(tmp2) To UBound(tmp2)
'
'            test_str = Replace(test_str, tmp2(i), "")
'
'        Next
        
        safe_check = .Range("H" & r_data - 5)
        safe_str = .Range("C" & r_data - 2)
        import_str = .Range("E" & r_data + 4)
    
    End With
    
    With ThisWorkbook.Sheets("監造報表")

        .Range("B2") = rec_code 'rec_date - work_day_start + 1
        .Range("C3") = weather_u
        .Range("E3") = weather_d
        .Range("G3") = rec_date
        .Range("B4") = con_name
        .Range("B5") = work_day & "天"
        .Range("D5") = work_day_start
        .Range("F5") = work_day_end
        .Range("B7") = pgs_design
        .Range("F7") = pgs_real
        .Range("A10") = items_str
        .Range("A12") = test_to_second_str
        .Range("A14") = test_str
        .Range("A16") = getSafeCheck(safe_check)
        .Range("A17") = "（二）其他工地安全衛生督導事項：" & safe_str
        .Range("A19") = import_str
        
        '契約金額
        .Range("H6") = "原契約:" & Format(rec_money, "#,##0")
        '變更次數
        '變更後契約
        
        Call outputData(wb_new, rec_code) ' rec_date - work_day_start + 1)

    End With
    
    End If

Next

If wb_new.Sheets.Count > 1 Then

    Application.DisplayAlerts = False
    
    For Each sht In wb_new.Sheets
    'wb_new.Sheets("工作表1").Delete
        If sht.Name Like "工作表*" Then sht.Delete
    Next
    
    Application.DisplayAlerts = True
    
End If

wb.Close False

End Sub

Sub outputData(ByVal wb As Workbook, ByVal code As String)

    Dim lastSheet As Worksheet
    
    ThisWorkbook.Activate
    ThisWorkbook.Sheets("監造報表").Copy after:=wb.Sheets(wb.Sheets.Count)
    
    Set lastSheet = wb.Sheets(wb.Sheets.Count)
    
    lastSheet.Name = code
    lastSheet.Columns("A:H").Copy
    lastSheet.Columns("A:H").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    Application.CutCopyMode = False
End Sub


Function getSafeCheck(ByVal safe_check As String)

If safe_check = "■有 □無" Then
    s = "■完成□未完成"
Else
    s = "□完成■未完成"
End If

getSafeCheck = "（一）施工廠商施工前檢查事項辦理情形：" & s

End Function

Function getRecItemString(ByVal sht)

Dim coll As New Collection

With sht

    lr = .Cells.Find("二、工地材料管理概況（含約定之重要材料使用狀況及數量等）：").Row

    For r = 10 To lr - 1
    
        If .Rows(r).Hidden = False Then
            cnt = cnt + 1
            item_name = .Cells(r, "A")
            num_all = .Cells(r, "F")
            num_today = .Cells(r, "H")
            num_sum = .Cells(r, "J")
            
            s = cnt & ". " & item_name & ":" & Round(num_today / num_all, 2) & "% 累積" & Round(num_sum / num_all, 2) & "%"
            's = cnt & ". " & .Cells(r, "A") & ":" & .Cells(r, "H") & " 累積" & .Cells(r, "J")
        
            coll.Add s
        
        End If
    
    Next

End With

For Each it In coll

    p = p & it & vbCrLf

Next

getRecItemString = p

End Function

Function getResultWorkbook(Optional ByVal f As String) As Object  'Optional ByVal f As String) '取得預算書內容

MsgBox "請先選取施工日誌的第一聯日報檔案!", vbInformation

If f = "" Then f = Application.GetOpenFilename

If f = "False" Then MsgBox "未取得檔案", vbCritical: End

tmp = Split(f, "\")

wbname = tmp(UBound(tmp))

Workbooks.Open (f)

Set getResultWorkbook = Workbooks(wbname)

End Function

Attribute VB_Name = "tmp_code"
Sub getProgress(ByVal recDate_str As String)

Dim obj As New clsDayReport
Dim pccesObj As New clsPCCES
Dim myFunc As New clsMyfunction

recDate = recDate_str ' "1120910"
recCode = recDate & "-1"
mode = 1
obj.print_mode = 1

Set coll_item = obj.getUsedItemByDate(recCode, myFunc.tranDate(recDate), 1)

For Each it In coll_item
    
    sum_amount = obj.getSumAmountByItem(it, myFunc.tranDate(recDate), mode)
    
    price = pccesObj.getMoneyByItemKey(it)

    use_money = use_money + price * sum_amount

Next

progress = use_money / pccesObj.getSumMoney

Dim pgs_rec_date As String

With ThisWorkbook.Sheets("�Ѯ�]�w")

    lr = .Cells(.Rows.Count, 1).End(xlUp).Row
    
    For r = 2 To lr
    
        pgs_rec_date = .Cells(r, 1)
    
        If pgs_rec_date = CStr(myFunc.tranDate(recDate_str)) Then
        
            r_pgs = r: Exit For
    
        End If
    
    Next

    .Cells(r_pgs, "E") = progress

End With

End Sub

Sub checkUsedItems()

With Sheets("�����ԲӪ�")

    lr = .Cells(.Rows.Count, 1).End(xlUp).Row
    
    For r = 2 To lr
    
        'check �����Ʈw���e
    
        s = .Cells(r, 2)
    
        If IsItemUsed(s) = True Then
        
            Debug.Print s
        
        End If
        
    
    Next

End With

End Sub

Function IsItemUsed(ByVal item_key)

Dim f As New clsMyfunction

Set collRows = f.getRowsByUser2("�����Ʈw", item_key, 1, "����")

Set sht = Sheets("�����Ʈw")

For Each r In collRows

    If sht.Cells(r, "H") = "" Then
    
        IsItemUsed = True
        Exit Function
    
    End If

Next

IsItemUsed = False

End Function


Sub checkDB()

Dim coll As New Collection

With Sheets("�����Ʈw")

.Unprotect

    lr = .Cells(.Rows.Count, 1).End(xlUp).Row

    For r = 2 To lr
    
        cmt = .Cells(r, "H")
        
        If cmt = "" Then
        
            item_name = .Cells(r, "E")
            
            item_key = .Cells(r, "D")
            
            Set rng = Sheets("�����ԲӪ�").Columns("C").Find(item_name)
            Set rng2 = Sheets("�����ԲӪ�").Columns("B").Find(item_key)
            
            If rng Is Nothing Then
            
                Debug.Print item_name
                
                 .Cells(r, "E") = rng2.Offset(0, 1).Value 'correct name
                 
                 On Error Resume Next
                 
                 p = item_name & " -> " & .Cells(r, "E") & vbNewLine
                
                coll.Add p, p
                
                On Error GoTo 0
                
            End If
        
        End If
    
    
    Next

.Protect

End With

If coll.Count = 0 Then


Else

    For Each p In coll

        pp = pp & p & vbNewLine
    
    Next
    
    MsgBox "[�󥿤����Ʈw���e]" & vbNewLine & vbNewLine & pp

End If

End Sub

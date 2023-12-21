Attribute VB_Name = "tmp_code"
Sub checkUsedItems()

With Sheets("契約詳細表")

    lr = .Cells(.Rows.Count, 1).End(xlUp).Row
    
    For r = 2 To lr
    
        'check 日報資料庫內容
    
        s = .Cells(r, 2)
    
        If IsItemUsed(s) = True Then
        
            Debug.Print s
        
        End If
        
    
    Next

End With

End Sub

Function IsItemUsed(ByVal item_key)

Dim f As New clsMyfunction

Set collRows = f.getRowsByUser2("日報資料庫", item_key, 1, "項次")

Set sht = Sheets("日報資料庫")

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

With Sheets("日報資料庫")

.Unprotect

    lr = .Cells(.Rows.Count, 1).End(xlUp).Row

    For r = 2 To lr
    
        cmt = .Cells(r, "H")
        
        If cmt = "" Then
        
            item_name = .Cells(r, "E")
            
            item_key = .Cells(r, "D")
            
            Set rng = Sheets("契約詳細表").Columns("C").Find(item_name)
            Set rng2 = Sheets("契約詳細表").Columns("B").Find(item_key)
            
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
    
    MsgBox "[更正日報資料庫內容]" & vbNewLine & vbNewLine & pp

End If

End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private MainRowColl As New Collection
Public shtName As String

Sub recordData(ByVal tmpType As String, ByVal tmpName As String)

With Sheets("日報填寫")

sr = MainRowColl(CInt(tmpType)) + 2
lr = MainRowColl(CInt(tmpType) + 1) - 1 '.Cells(.Rows.Count, 1).End(xlUp).Row

For r = sr To lr

    rec_key = .Cells(r, 1)
    rec_num = .Cells(r, 2)
    rec_item = .Cells(r, 3)
    rec_unit = .Cells(r, 4)
    
    If rec_key <> "" Then
    
        arr = Array(tmpType, tmpName, rec_key, rec_num, rec_item, rec_unit)
        Call KeyInTmp(arr)
        
        msg = msg & rec_item & vbNewLine
    
    End If

Next

.Activate

End With

MsgBox ("儲存範本「" & tmpName & "」成功!!" & vbNewLine & vbNewLine & "內容包含：" & vbNewLine & msg), vbInformation

End Sub

Private Sub KeyInTmp(ByVal arr)

With Sheets("範本資料庫")

'.Unprotect

lr = .Cells(.Rows.Count, 1).End(xlUp).Row
r = lr + 1

For i = LBound(arr) To UBound(arr)

    .Cells(r, i + 1) = arr(i)

Next

'.Protect

End With

End Sub

Function getCboItems() 'ByVal tmp_type As String) ', ByVal tmp_name As String)

tmpType = InputBox("請輸入範本種類" & vbNewLine & "1.施工工項" & vbNewLine & "2.材料管理", , "1")

TmpForm.Label1 = tmpType

Dim coll As New Collection

With Sheets("範本資料庫")

lr = .Cells(.Rows.Count, 1).End(xlUp).Row

For r = 2 To lr

    If CStr(.Cells(r, 1)) = tmpType Then ' And CStr(.Cells(r, 2)) = tmp_name Then

        tmp_name = .Cells(r, 2)

        On Error Resume Next
        
        coll.Add tmp_name, CStr(tmp_name)
        
        On Error GoTo 0
    
    End If

Next

End With

Set getCboItems = coll

End Function

Sub getDatabyTmp(ByVal tmp_type As String, ByVal tmp_name As String)

With Sheets("範本資料庫")

'.Unprotect

lr = .Cells(.Rows.Count, 1).End(xlUp).Row

For r = 2 To lr

    If CStr(.Cells(r, 1)) = tmp_type And CStr(.Cells(r, 2)) = tmp_name Then
    
        mydata = .Cells(r, 3)
        
        With Sheets("日報填寫")
        
            'rr = .Cells(.Rows.Count, 1).End(xlUp).Row
            
            If tmp_type = "1" Then
                rr = .Cells(MainRowColl(1), 1).End(xlDown).Row
            Else
                rr = .Cells(MainRowColl(2), 1).End(xlDown).Row
            End If
            
            .Unprotect
            
            .Cells(rr + 1, 1).WrapText = False
            .Cells(rr + 1, 1) = mydata
        
            .Protect
        
        End With
    
    End If

Next

'.Protect

End With

End Sub

Private Sub Class_Initialize()

Dim o As New clsWriteData

Set MainRowColl = o.getMainRowColl

End Sub

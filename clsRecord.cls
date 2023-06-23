VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public mode As Byte

Private shtRec As Object
Private shtData As Object

Private recObj As New clsRecord
Private funObj As New clsFunction
Private dataObj As New clsWriteData

Private MainRowColl As New Collection

Private Sub Class_Initialize()

Set shtRec = Sheets("日報資料庫")
Set shtData = Sheets("日報填寫")

Set MainRowColl = dataObj.getMainRowColl

End Sub

Function getSameCode()

'If mode > 1 Then Set shtRec = Sheets(dataObj.getDB_name(mode))

Set shtRec = Sheets(dataObj.getDB_name(1))

Dim coll As New Collection

With shtRec

    .Unprotect

    lr = .Cells(.Rows.count, 1).End(xlUp).Row
    
    For r = 2 To lr
    
        recCode = .Cells(r, 1).Value
        recNote = .Cells(r, 8).Value

        If recNote = "" Then

        On Error Resume Next
        
        coll.Add recCode, recCode
        
        On Error GoTo 0
        
        End If
    
    Next
    
    .Protect

End With

Set getSameCode = coll

End Function

Function getSameCodeByDates(ByVal sDate As Date, ByVal eDate As Date)

Dim coll As New Collection

With shtRec

    .Unprotect

    lr = .Cells(.Rows.count, 1).End(xlUp).Row
    
    For r = 2 To lr
    
        recCode = .Cells(r, 1)
        recDate = funObj.tranDate(.Cells(r, 2))
        recNote = .Cells(r, 8)

        If recNote = "" And recDate >= sDate And recDate <= eDate Then

            On Error Resume Next
            
            coll.Add recCode, recCode
            
            On Error GoTo 0
        
        End If
    
    Next
    
    .Protect

End With

Set getSameCodeByDates = coll

End Function


Private Function checkCodeExist(ByVal code As String)

checkCodeExist = True

Set rng = shtRec.Cells.Find(code)

If rng Is Nothing Then checkCodeExist = False

End Function

Sub dealExistData(ByVal code As String) ', ByRef IsExist, ByRef IsDelete)

IsExist = checkCodeExist(code) '只有檢查日報資料庫

If IsExist = True Then

    msg = MsgBox("是否要作廢所有" & code & "?", vbYesNo)
    
    If msg = vbNo Then
    
        MsgBox "不作廢就無法用同一編號", vbCritical
        dataObj.clearInformation
        dataObj.clearDataAll
        
        Exit Sub

    End If

End If

For i = 1 To 5

    With Sheets(dataObj.getDB_name(i))
    
        .Unprotect
            
        lr = .Cells(.Rows.count, 1).End(xlUp).Row
        
        For r = 2 To lr
        
            If .Cells(r, 1) = code Then .Cells(r, "H") = "作廢"
        
        Next
    
        .Protect
    
    End With

Next

End Sub

Function getCodeCount(ByVal strDate As String)

'Dim obj As New clsRecord

Set coll = recObj.getSameCode

For Each it In coll

    tmp = Split(it, "-")
    
    If strDate = tmp(0) Then cnt = cnt + 1

Next

getCodeCount = cnt + 1

End Function

Sub getDatabyCode(ByVal myCode) '查詢

For i = 1 To 4

    With Sheets(dataObj.getDB_name(i)) 'shtRec
    
        .Unprotect
        
        lr = .Cells(.Rows.count, 1).End(xlUp).Row
        
        For r = 2 To lr
        
            If .Cells(r, 1) = myCode Then
            
                recDate = .Cells(r, 2)
                recLoc = .Cells(r, 3)
                recNum = .Cells(r, 4)
                recItem = .Cells(r, 5)
                recUnit = .Cells(r, 6)
                recAmount = .Cells(r, 7)
                recNote = .Cells(r, 8)
                
                If recNote = "" Then
            
                    arr = Array(myCode, recDate, recLoc, recNum, recItem, recUnit, recAmount)
            
                    Call dataObj.reverseData(arr, i)
            
                End If
            
            End If
        
        Next
        
        .Protect
    
    End With

Next

shtData.Activate

End Sub

'Sub sortRecord()
'
'With shtRec
'
'.Unprotect
'
'lr = .Cells(.Rows.Count, 1).End(xlUp).Row
'
'.Range("A2").Resize(lr, 8).Sort key1:=.Columns(1), order1:=xlAscending
'
'.Protect
'
'End With
'
'End Sub

Function getRecLoc(ByVal recCode As String)

With Sheets("日報資料庫")

lr = .Cells(.Rows.count, 1).End(xlUp).Row

For r = 2 To lr

    If .Cells(r, 1) = recCode And .Cells(r, 8) = "" Then
    
        getRecLoc = .Cells(r, 3): Exit Function
    
    End If

Next

End With

End Function


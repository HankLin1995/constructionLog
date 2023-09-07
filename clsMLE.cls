VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMLE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Sub ReadData()

Dim mytype As String
Dim item As String
Dim unit As String
Dim amount As String
Dim price As String

With MLEForm

    If .optM.Value = True Then
        mytype = "M" '材料
    ElseIf .optL.Value = True Then
        mytype = "L" '人力
    ElseIf .optE.Value = True Then
        mytype = "E" '機具
    End If

    If .tboItem = "" Then MsgBox "請填寫項目名稱", vbCritical: End
    If .tboUnit = "" Then MsgBox "請填寫單位", vbCritical: End

    item = .tboItem
    unit = .tboUnit
    amount = .tboAmount
    price = .tboPrice

    Call addMLEData(mytype, item, unit, amount, price)

End With

End Sub

Private Sub addMLEData(ByVal mytype As String, ByVal item As String, ByVal unit As String, ByVal amount As String, ByVal price As String)

With Sheets("工料設定")

    lr = .Cells(.Rows.Count, 1).End(xlUp).Row
    
    r = lr + 1
    
    .Cells(r, 2) = mytype & Format(countType(mytype) + 1, "0000")
    .Cells(r, 3) = item
    .Cells(r, 4) = unit
    .Cells(r, 5) = amount
    .Cells(r, 6) = price
    
    .Cells(r, 1) = .Cells(r, 2) & ">" & .Cells(r, 3)

    If .Cells(r, 5) = 0 Then .Cells(r, 5) = "-"

End With

End Sub

Sub sortData()

With Sheets("工料設定")

    lr = .Cells(.Rows.Count, 1).End(xlUp).Row
    
    .Range("A2").Resize(lr, 6).Sort key1:=.Columns(2), order1:=xlAscending

End With

End Sub

Function countType(ByVal mytype As String)

With Sheets("工料設定")

    lr = .Cells(.Rows.Count, 1).End(xlUp).Row
    
    For r = 2 To lr
    
        MLEType = .Cells(r, 2)
        
        If mid(MLEType, 1, 1) = mytype Then
        
            countType = countType + 1
        
        End If
    
    Next

End With

End Function

Sub setValidation_MLE()

Dim o As New clsWriteData

Set coll = o.getMainRowColl

m_formula = getValidationFormula("M")
l_formula = getValidationFormula("L")
e_formula = getValidationFormula("E")

m_row = coll(2)
l_row = coll(3)
e_row = coll(4)
p_row = coll(5)

With Sheets("日報填寫")

    .Unprotect

    For r = m_row + 2 To l_row - 1

        .Range("A" & r).Validation.Delete
        .Range("A" & r).Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
    Formula1:=m_formula
        
    Next
    
    For r = l_row + 2 To e_row - 1

        .Range("A" & r).Validation.Delete
        .Range("A" & r).Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
    Formula1:=l_formula
    
    Next
    
    For r = e_row + 2 To p_row - 1

        .Range("A" & r).Validation.Delete
        .Range("A" & r).Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
    Formula1:=e_formula
    
    Next
    
    .Protect

End With

End Sub

Private Function getValidationFormula(ByVal mode As String)

Dim coll As New Collection

With Sheets("工料設定")

    lr = .Cells(.Rows.Count, 1).End(xlUp).Row
    
    For r = 2 To lr
    
        If mid(.Cells(r, 2), 1, 1) = mode Then
        
            sr = r: Exit For
            
        End If
        
    Next
    
    cnt = countType(mode)
    
    getValidationFormula = "=工料設定!$A" & sr & ":$A" & sr + cnt - 1

End With

End Function


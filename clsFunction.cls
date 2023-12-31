VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Sub data_validation_from_array(ByVal rng As Range, ByVal region)

'Dim region, product As Variant
''Dim region_range, product_range As Range

'region = Array("North", "South", "East", "West")
'region = Array("□有 □無", "■有 □無", "□有 ■無")

With rng.Validation
.Delete
.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=Join(region, ",")
.IgnoreBlank = True
.InCellDropdown = True
.InputTitle = ""
.ErrorTitle = "Error"
.InputMessage = ""
.ErrorMessage = "Please Provide a Valid Input"
.ShowInput = True
.ShowError = True
End With

End Sub

Function tranDate(ByVal mydatestr) As Date

Y = mid(mydatestr, 1, 3)
M = mid(mydatestr, 4, 2)
D = mid(mydatestr, 6, 2)

tranDate = CDate(1911 + Y & "/" & M & "/" & D)

End Function

Function tranZHDate(ByVal mydate As Date) As String

tmp = Split(mydate, "/")

Y = tmp(0) - 1911
M = Format(tmp(1), "00")
D = Format(tmp(2), "00")

tranZHDate = Y & M & D

End Function

Function tranCharcter(ByVal s As String)

If s Like "*~*" Then

    For i = 1 To Len(s)
    
        ch = mid(s, i, 1)
        If ch = "~" Then ch = "-"
        new_ch = new_ch & ch
    
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

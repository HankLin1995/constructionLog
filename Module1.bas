Attribute VB_Name = "Module1"
'1.CheckIfSigned
'2.CheckPASS

Sub loadOldDatabase()

Set wb = Workbooks(getWbName())

'shtNameArr = Array("�����ԲӪ�", "�Ю׳]�w", "�u�Ƴ]�w", "�Ѯ�]�w", "�����Ʈw", "�d����Ʈw", "���Ƹ�Ʈw", "�����Ʈw", "�H����Ʈw", "����Ʈw")
shtNameArr = Array("�����Ʈw", "�d����Ʈw", "���Ƹ�Ʈw", "�����Ʈw", "�H����Ʈw", "����Ʈw")

For Each shtName In shtNameArr

    ThisWorkbook.Sheets(shtName).Unprotect

    With wb.Sheets(shtName)
    
        .Unprotect
        Set rng = .Cells.SpecialCells(xlCellTypeLastCell)
        .Range("A1").Resize(rng.Row, rng.Column).EntireRow.Copy ThisWorkbook.Sheets(shtName).Range("A1")
        .Protect
        
    End With
    
    ThisWorkbook.Sheets(shtName).Protect

Next

End Sub

Function getWbName()

Dim coll As New Collection

For Each wb In Workbooks

    If wb.Name <> ThisWorkbook.Name Then
        j = j + 1
        p = p & j & "." & wb.Name & vbNewLine
        coll.Add wb.Name
    End If

Next

mode = InputBox("�п�ܶפJ���ɮצW��:" & vbNewLine & p, , "1")

getWbName = coll(CInt(mode))

End Function

Sub checkSubOnAction()

For Each sht In Sheets

    For Each shp In sht.Shapes
    
        If shp.OnAction <> "" Then
        
            Debug.Print shp.OnAction
        
        End If
    
    Next

Next

End Sub


Function getSumByItemName(ByVal item_name_key As String, ByVal report_date As Date)

'item_name_key = "��.1>�g��u�@�A���g��-"
'report_date = Now()

tmp = split(item_name_key, ">")

item_index = tmp(0)
item_name = tmp(1)

Dim myFunc As New clsMyfunction
Dim func As New clsFunction

Set coll_rows = myFunc.getRowsByUser2("�����Ʈw", item_name, 1, "����")

'double check

With Sheets("�����Ʈw")

    For Each r In coll_rows
    
        If .Cells(r, "D") = tmp(0) And .Cells(r, "B") <> "" Then
            
            rec_date = func.tranDate(.Cells(r, "B"))
            rec_note = .Cells(r, "H")
            
            If rec_date <= report_date And rec_note = "" Then
            
            item_num = .Cells(r, "G")
            item_sum = item_sum + item_num
        
            End If
        
        Else
        
        End If
    
    Next

End With

With Sheets("�����ԲӪ�")

Set coll_rows_PCCES = myFunc.getRowsByUser2("�����ԲӪ�", item_name_key, 1, "�z�����")

For Each r_PCCES In coll_rows_PCCES

    con_sum = .Cells(r_PCCES, "E")

Next

End With

Debug.Print item_name & ":" & item_sum

getSumByItemName = con_sum - item_sum

End Function


Sub test_CheckIsRepeat()

Dim o As New clsPCCES

o.checkIsRepeat

End Sub

Sub test_cmdGetDayReport() 'ByVal sDate As String, ByVal eDate As String, ByVal mode As Byte)

Dim obj As New clsDayReport


'============getInformationbyForm===============

With DayReportForm

obj.StartDate = .tbosDate
obj.EndDate = .tboeDate

obj.print_mode = 3 'print_mode

End With

'==============================================

Call obj.getInterval(sr, er) '���o����_�����Ʀr���A

'Set wb = Workbooks.Add

For r = er To sr Step -1

    'get the codes
    'ThisWorkbook.Activate
    
    Set coll_code = obj.getCodes(obj.workDate + r - 1) '�ھڤ�����o���Codes
    
    Debug.Print obj.workDate + r - 1
    
    If coll_code.Count <> 0 Then
    
        For Each it In coll_code
        
            Debug.Print r & ":" & it
        
        Next
    
    End If
    
Next
    
'    If print_mode = 3 Or print_mode = 4 Then
'
'        Dim coll_code_new As New Collection
'
'        For Each it In coll_code
'
'            If Split(it, "-")(1) = "1" Then
'
'                coll_code_new.Add Split(it, "-")(0)
'
'            End If
'
'        Next
'
'        Set coll_code = coll_code_new
'
'    End If
'

End Sub

Sub test_0509()

Dim obj As New clsDayReport

r = 28

Debug.Print obj.workDate + r - 1

Set coll_code = obj.getCodes(obj.workDate + r - 1)

For Each it In coll_code

    Debug.Print r & ":" & it

Next

End Sub

Sub test_tranNum()

s = "(2)"

Dim o As New clsFunction

Debug.Print o.tranCharcter_NUM(s)

End Sub

Sub test_adddata()

Set rng = Range("B3")
region = Array("���� ���L", "���� ���L", "���� ���L")

Call data_validation_from_array(rng, region)

End Sub

Sub data_validation_from_array(ByVal rng As Range, ByVal region)

'Dim region, product As Variant
'Dim region_range, product_range As Range

region = Array("North", "South", "East", "West")

region = Array("���� ���L", "���� ���L", "���� ���L")

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

Sub test_AddValidation()

arr = Array("���� ���L", "���� ���L", "���� ���L")

Range("G" & r + 5).Validation.Delete
Range("G" & r + 5).Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
Formula1:=Join(arr, ",")

End Sub

Sub t()
Sheets("���v").Visible = True
End Sub

Sub CheckIfSigned()

Dim o As New clsUserInformation
o.hideCmd

MAC_ADDRESS = o.getMacAddress

If MAC_ADDRESS = "" Then
    Debug.Print "�нT�{�����O�_���s�W����~!"
Else
    IsSigned = o.checkIsExist(MAC_ADDRESS)

    If IsSigned Then
    
        Debug.Print MAC_ADDRESS & "�w�g�Q�ϥΤF!"
        
    Else
        
        Debug.Print "�եΪ�"
        Call o.Login(MAC_ADDRESS)
        Call test_SignClient
        ThisWorkbook.Save
  
    End If
    
    Call test_Access(MAC_ADDRESS)

End If

End Sub


Sub test_SignClient()

Dim o As New clsFetchURL

myURL = o.CreateURL("Sign") ', "Hank", "YLIA", "apple@mm")

Debug.Print "sign PASS"

If o.ExecHTTP(myURL) = "signed" Then
    MsgBox "�ӹq���w�g�Q���U�L�F!", vbCritical
End If

End Sub

Sub test_Access(ByVal mac_add As String)

Dim ui As New clsUserInformation
Dim o As New clsFetchURL

myURL = o.CreateURL("Access")

Status = o.ExecHTTP(myURL)

Select Case Status

Case "PASS"
    Debug.Print "���ҳq�L!"
    ui.showCmd
    
Case "NOT_FOUND"
    MsgBox "�䤣���Ʈw���A�������Ǹ��A�бN���i�ǰe�ܺ޲z�H��", vbCritical
    ERRORForm.Show
    
Case "ARRIVED":

    MsgBox "������ϥΤѼƬ�0��A�p�G�n�ϥνХ��i����v!", vbInformation
    SignDetailForm.Show
Case Else

    MsgBox Status

End Select

End Sub

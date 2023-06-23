Attribute VB_Name = "hardware"


Private Const Sep As String = ","

Private Type GUID_TYPE
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type
 
Private Declare PtrSafe Function CoCreateGuid Lib "ole32.dll" (GUID As GUID_TYPE) As LongPtr
Private Declare PtrSafe Function StringFromGUID2 Lib "ole32.dll" (GUID As GUID_TYPE, ByVal lpStrGuid As LongPtr, ByVal cbMax As Long) As LongPtr
 
Function CreateGuidString()
    Dim GUID As GUID_TYPE
    Dim strGuid As String
    Dim retValue As LongPtr
    
    Const guidLength As Long = 39 'registry GUID format with null terminator {xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx}
    
    retValue = CoCreateGuid(GUID)
    If retValue = 0 Then
        strGuid = String$(guidLength, vbNullChar)
        retValue = StringFromGUID2(GUID, StrPtr(strGuid), guidLength)
        If retValue = guidLength Then
            ' valid GUID as a string
            CreateGuidString = strGuid
        End If
    End If
End Function
 
Function GetGUID()
    Dim strGuid As String
    strGuid = CreateGuidString()
    
    strGuid = Replace(Replace(strGuid, "{", ""), "}", "")
    
    GetGUID = strGuid
End Function

Function getSystemID() '''取計算機ID

Dim idObj, id, inobj
Set idObj = GetObject("winmgmts:{impersonationLevel=impersonate}").InstancesOf("Win32_OperatingSystem")
For Each inobj In idObj
If inobj.SerialNumber <> "" Then 'SerialNumber 計算ID號
id = inobj.SerialNumber
End If
Next

getSystemID = id

End Function

Function getMacAddress()

Dim objVMI As Object
Dim vAdptr As Variant
Dim objAdptr As Object
'Dim adptrCnt As Long


Set objVMI = GetObject("winmgmts:\\" & "." & "\root\cimv2")
Set vAdptr = objVMI.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")

For Each objAdptr In vAdptr
    If Not IsNull(objAdptr.MACAddress) And IsArray(objAdptr.IPAddress) Then
        For adptrCnt = 0 To UBound(objAdptr.IPAddress)
        If Not objAdptr.IPAddress(adptrCnt) = "0.0.0.0" Then
            GetNetworkConnectionMACAddress = objAdptr.MACAddress
            Exit For
        End If
        Next
    End If
Next

getMacAddress = GetNetworkConnectionMACAddress

End Function



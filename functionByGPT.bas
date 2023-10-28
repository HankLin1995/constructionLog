Attribute VB_Name = "functionByGPT"
Sub test_SplitMultipleParenthesesStrings()

myString = "1�B������ޤΰt�q  2�B���j���ʪO 3�B���q�Խu 4�B�Ѫ�o���Q��e�m�@�~[4�B�q��u�{�I�u�d��<�X��>]"

Call SplitAndCombineParenthesesStrings(myString)

End Sub

Sub SplitAndCombineParenthesesStrings(ByVal originalString As String)
    Dim leftParenthesisPosition As Integer
    Dim rightParenthesisPosition As Integer
    Dim subString As String
    Dim outsideString As String
    Dim parenthesisContents() As String
    Dim i As Integer
    
    ' Initialize the array
    ReDim parenthesisContents(1 To 1)
    
    ' Initialize the outside string
    outsideString = ""
    
    ' Find the position of the left parenthesis
    leftParenthesisPosition = InStr(originalString, "[")
    
    ' Initialize the starting position
    Dim startPosition As Integer
    startPosition = 1

    ' Loop to process all parentheses
    Do While leftParenthesisPosition > 0
        ' Find the position of the right parenthesis
        rightParenthesisPosition = InStr(leftParenthesisPosition + 1, originalString, "]")
        
        If rightParenthesisPosition > 0 Then
            ' Extract the content within parentheses
            subString = mid(originalString, leftParenthesisPosition + 1, rightParenthesisPosition - leftParenthesisPosition - 1)
            
            ' Store the content within parentheses in the array
            parenthesisContents(UBound(parenthesisContents)) = subString
            ReDim Preserve parenthesisContents(1 To UBound(parenthesisContents) + 1)
            
            ' Update the outside string with content outside of parentheses
            outsideString = outsideString & mid(originalString, startPosition, leftParenthesisPosition - startPosition)
            
            ' Update the starting position
            startPosition = rightParenthesisPosition + 1
            
            ' Update the position of the left parenthesis
            leftParenthesisPosition = InStr(startPosition, originalString, "[")
        Else
            ' Exit the loop if a right parenthesis is not found
            Exit Do
        End If
    Loop
    
    ' Combine the outside string after the last parenthesis
    outsideString = outsideString & mid(originalString, startPosition)
    
    ' Output all contents within parentheses and the combined outside string
    For i = 1 To UBound(parenthesisContents) - 1
    
        tmp = split(parenthesisContents(i), "�B")
    
        'Debug.Print "�ĤG��:" & i & "�B" & tmp(1) 'parenthesisContents(i)
        
        secondString = secondString & i & "�B" & tmp(1)
        
    Next i
    
    Debug.Print "�ĤG��:" & secondString
    Debug.Print "��L: " & outsideString
    
End Sub


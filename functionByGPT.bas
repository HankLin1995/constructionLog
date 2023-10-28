Attribute VB_Name = "functionByGPT"
Sub test_SplitMultipleParenthesesStrings()

myString = "1、消防改管及配電  2、輕隔間封板 3、機電拉線 4、天花油漆噴塗前置作業[4、電氣工程施工查驗<合格>]"

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
    
        tmp = split(parenthesisContents(i), "、")
    
        'Debug.Print "第二項:" & i & "、" & tmp(1) 'parenthesisContents(i)
        
        secondString = secondString & i & "、" & tmp(1)
        
    Next i
    
    Debug.Print "第二項:" & secondString
    Debug.Print "其他: " & outsideString
    
End Sub


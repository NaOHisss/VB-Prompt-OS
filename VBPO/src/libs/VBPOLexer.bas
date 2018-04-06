Attribute VB_Name = "VBPOLexer"
Public Function SplitString(src_str As String, split_chars() As String, mul_split As Boolean, Optional start_pos As Long) As String()
Dim result() As String
ReDim result(1)
thisChar$ = ""
While True
    Dim word_pos As Long
workloop: thisChar = Mid(src_str, start_pos + 1, 1)
    start_pos = start_pos + 1
    If thisChar = "" Or thisChar = vbNullString Then
    GoTo ReturnRes
    End If
    If Contains(thisChar, split_chars) Then
        If result(word_pos) = "" Then
            GoTo workloop
        Else: word_pos = word_pos + 1
        ReDim Preserve result(word_pos)
        GoTo workloop
        End If
    End If
    result(word_pos) = result(word_pos) + thisChar
    GoTo workloop
Wend
ReturnRes:
ReDim Preserve result(word_pos)
SplitString = result
End Function
Public Function Contains(char As String, char_set() As String) As Boolean
For i = LBound(char_set) To UBound(char_set)
    If char = char_set(i) Then
        Contains = True
        Exit Function
    End If
Next i
Contains = False
End Function

Attribute VB_Name = "BaseIO"
Public Function read(path As String) As String
Dim str As String
Dim tmp As String
Open path For Input As #1
Do While Not EOF(1)
Line Input #1, tmp
tmp = Replace(tmp, "\n", vbCrLf)
str = str + tmp
Loop
Close #1
read = str
End Function
Public Function read_utf8(path As String) As String
Dim str As String
Dim tmp As String
Open path For Input As #1
Do While Not EOF(1)
Line Input #1, tmp
tmp = UTF8_Decode(tmp)
tmp = Replace(tmp, "\n", vbCrLf)
str = str + tmp
Loop
Close #1
read_utf8 = str
End Function
Public Function n_read(path As String) As String
Dim str As String
Dim tmp As String
Open path For Input As #1
Do While Not EOF(1)
Line Input #1, tmp
If str = "" Then
str = tmp
Else: str = str + vbCrLf + tmp
End If
Loop
Close #1
n_read = str
End Function
Public Function n_read_utf8(path As String) As String
Dim str As String
Dim tmp As String
Open path For Input As #1
Do While Not EOF(1)
Line Input #1, tmp
tmp = UTF8_Decode(tmp)
str = str + tmp
Loop
Close #1
n_read_utf8 = str
End Function
Public Sub writeTo(str As String, path As String)
Open path For Output As #1
Print #1, str
Close #1
End Sub
Public Sub appendTo(str As String, path As String)
Open path For Append As #1
Print #1, str
Close #1
End Sub
Public Function readEachLine(path As String) As String()
Dim str() As String
Dim tmp As String
Dim i As Long
i = 0
Open path For Input As #1
Do While Not EOF(1)
Line Input #1, tmp
tmp = Replace(tmp, "\n", vbCrLf)
ReDim Preserve str(0 To i)
str(i) = tmp
i = i + 1
Loop
Close #1
readEachLine = str
End Function
Public Function readEachLine_utf8(path As String) As String()
Dim tmp As String
Dim lines() As String
tmp = UTF8_Decode_File(path)
lines = Split(tmp, vbCrLf)
For i = 0 To UBound(lines)
lines(i) = Replace(lines(i), "\n", vbCrLf)
Next i
readEachLine_utf8 = lines
End Function
Public Function createFolder(fp As String) As Integer
Dim thisp As String
Dim tmp() As String
tmp = Split(fp, "\")
If UBound(tmp) = 0 Then GoTo l1
On Error GoTo Err
l2: If Dir(tmp(0), vbDirectory) = "" Then
    GoTo Err
    Else
    For i = 0 To UBound(tmp)
        If thisp = "" Then
        thisp = tmp(i)
        Else
        thisp = thisp + "\" + tmp(i)
        End If
    On Error Resume Next
    MkDir thisp
    Next i
    End If
createFolder = 1
Exit Function
l1: tmp = Split(fp, "/")
GoTo l2
Err: createFolder = 0
Exit Function
End Function
Public Function read_b(file As String) As Byte()
Dim b() As Byte
ReDim b(FileLen(file) - 1) As Byte
Open file For Binary As #1
Get #1, , b
Close #1
read_b = b
End Function
Public Sub write_b(file As String, bytes() As Byte)
Open file For Binary As #1
Put #1, , bytes
Close #1
End Sub
Public Function copy(s As String, d As String) As Integer
On Error GoTo Err
Open s For Binary As #1
Open d For Binary As #2
Dim b As Byte
Do While Not EOF(1)
Get #1, , b
Put #2, , b
Loop
Close #1
Close #2
copy = 1
Exit Function
Err: copy = 0
End Function
Public Function getLastName(p As String) As String
Dim tmp() As String
tmp = Split(p, "\")
If UBound(tmp) = 0 Then
GoTo l1
Else
l2: getLastName = tmp(UBound(tmp))
End If
Exit Function
l1: tmp = Split(p, "/")
GoTo l2
End Function
Public Function getExtendedName(file As String)
Dim arr() As String
arr = Split(getLastName(file), ".")
getExtendedName = arr(UBound(arr))
End Function

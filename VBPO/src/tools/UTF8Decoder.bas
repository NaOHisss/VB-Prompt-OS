Attribute VB_Name = "UTF8Decoder"
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByRef lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cbMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
Private Const CP_UTF8 = 65001
Private Const CP_ACP = 0

Public Function UTF8_Decode_File(ByVal path As String) As String
Dim file As String
file = UTF8_Decode(Get_String_From_File(path))
UTF8_Decode_File = Left(file, Len(file) - 1)
End Function
Public Function UTF8_Decode(ByVal UTF8Str As String) As String
If UTF8Str = "" Then
Exit Function
End If
Dim UTF8Bytes() As Byte
UTF8Bytes = UTF8Str
Dim UTF8Size As Long
Dim WideCharNum As Long
Dim strBuffer As String
Dim BufferSize As Long
Dim Num As Long
UTF8Size = UBound(UTF8Bytes) + 1
BufferSize = UTF8Size * 2
strBuffer = String$(BufferSize, vbNullChar)
Num = MultiByteToWideChar(CP_UTF8, 0, UTF8Bytes(0), UTF8Size, StrPtr(strBuffer), BufferSize)
UTF8_Decode = Mid(strBuffer, 2, Num)
End Function

Private Function Get_String_From_File(ByVal path As String) As String
Dim Res() As Byte
ReDim Res(FileLen(path) - 1) As Byte
Open path For Binary As #1
Get #1, , Res
Close #1
Get_String_From_File = Res
End Function

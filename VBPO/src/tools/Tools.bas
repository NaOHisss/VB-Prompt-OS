Attribute VB_Name = "Tools"
Public Declare Sub CopyMemory Lib "kernel132" Alias "RtlMoveMemory" (ptr_Des As Any, ptr_Src As Any, ByVal bLen As Long)
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Public Const EM_GETLINE = &HC4
Public Const EM_GETLINECOUNT = &HBA
Public Const EM_GETMODIFY = &HB8

Public Function getVersion() As String
getVersion = str(App.Major) + "." + str(App.Minor) + "." + str(App.Revision)
End Function

Function addForm(Name As String, Height As Integer, Width As Integer, Left As Integer, Top As Integer, Optional Visable As Boolean, Optional Icon As String, Optional showMode As Integer) As Long
Dim frm As Form
Set frm = New BaseForm
frm.Caption = Name
frm.Icon = LoadPicture(Icon)
frm.Height = Height
frm.Width = Width
frm.Left = Left
frm.Top = Top
If Visable Then
frm.Show showMode
End If
addForm = frm.hwnd
End Function
Public Function getExName(file As String) As String
On Error GoTo n
Dim tmp() As String
tmp = Split(file, ".")
If UBound(tmp) = 0 Then
GoTo n
End If
getExName = tmp(UBound(tmp))
Exit Function
n: getExName = ""
End Function
Public Function getEachText(txt As String) As String()
Dim t() As String
If txt = "" Then
ReDim t(0 To 0)
Else
t = Split(txt, vbCrLf)
End If
getEachText = t
End Function
Public Sub msgReBoot(msg As String)
Dim yn As Integer
yn = MsgBox(msg, 36, Strs.rebootvbpo)
If yn = vbYes Then
Call reboot
End If
End Sub
Public Sub reboot()
Shell App.path + "\" + App.EXEName + ".exe"
End
End Sub

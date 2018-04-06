Attribute VB_Name = "Options"
Public SysLan As String
Public showSuccess As Boolean

Public Sub loadOptions()
SysLan = read(App.path + "\system\options\language.data")
Select Case LCase(read(App.path + "\system\options\showsuccess.data"))
Case "true": showSuccess = True
Case "false": showSuccess = False
Case Else: writeTo "true", App.path + "\system\options\showsuccess.data"
showSuccess = True
End Select
End Sub

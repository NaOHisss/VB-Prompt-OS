Attribute VB_Name = "RunCmd"
Private cmd As String
Private cmdstrs() As String
Public pos As Long
Public para As String
Public tmp As String
Public splitchars(0 To 1) As String
Public Sub Run(sc As String)
ReSet
cmdstrs = SplitString(sc, splitchars, True)
cmd = getNextWord_R
para = getNextWord_R
RunCode
End Sub
Private Sub getNextWord()
On Error GoTo Err
pos = pos + 1
para = cmdstrs(pos)
Exit Sub
Err: para = ""
End Sub
Private Sub getWordAt(id As Integer)
On Error GoTo Err
para = cmdstrs(id)
Exit Sub
Err: para = ""
End Sub
Private Function getNextWord_R() As String
On Error GoTo Err
pos = pos + 1
tmp = cmdstrs(pos)
getNextWord_R = tmp
Exit Function
Err: tmp = "": getNextWord_R = ""
End Function
Private Function getWordAt_R(id As Integer) As String
On Error GoTo Err
tmp = cmdstrs(id)
getWordAt_R = tmp
Exit Function
Err: tmp = "": getWordAt_R = ""
End Function
Private Sub RunCode()
Select Case RunCmd.cmd
Case "":
Case "exec": Call cmd_exec
Case "help": Call cmd_help
Case "run": Call cmd_run
Case "reboot": Call cmd_reboot
Case "exit": Call cmd_exit
Case "fuck": Call cmd_fuck
Case "clear": Call cmd_clear
Case "about": Call cmd_about
Case "version": Call cmd_version
Case "time": Call cmd_time
Case "copy": Call cmd_copy
Case "delete": Call cmd_delete
Case "makedir": Call cmd_makedir
Case "hide": Call cmd_hide
Case "read": Call cmd_read
Case "write": Call cmd_write

Case Else: OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + Strs.err_unknowncommand + cmd
End Select

OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + vbCrLf
End Sub
Private Sub cmd_fuck()
Select Case LCase(para)
Case "naohisss":
OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + "You have no right to fuck me.Go and fuck yourself."
Case "author":
OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + "You have no right to fuck me.Go and fuck yourself."
Case "developer":
OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + "You have no right to fuck me.Go and fuck yourself."
Case "":
OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + "Input a person's name who you want to fuck."
Case "me":
OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + "No,it's not worth fucking you."
Case "i":
OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + "No,it's not worth fucking you."
Case "you":
OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + "Really?I'm surprised you want to fuck a Prompt OS.That's amazing,young man!"
Case "vbpo":
OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + "Really?I'm surprised you want to fuck a Prompt OS.That's amazing,young man!"
Case "this":
OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + "Really?I'm surprised you want to fuck a Prompt OS.That's amazing,young man!"
Case Else:
OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + para + " was fucked."
End Select
End Sub
'All functions here......
Private Sub cmd_exec()
If gerNextWord_R = "" Then
    On Error GoTo Err
    If Dir(para) = "" Then
        GoTo Err
    Else
    Shell RunCmd.para
    End If
    If showSuccess = True Then
        OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + Strs.suc_exec + para
    End If
Else: GoTo UP
End If
Exit Sub
UP: OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + Strs.err_unknownparameter + tmp
Exit Sub
Err: OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + Strs.err_exec + para
End Sub
Private Sub cmd_help()
If gerNextWord_R = "" Then
Select Case para
Case "": OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + Strs.cmd_help
Case "help": OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + Strs.cmd_help_help
Case "exec": OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + Strs.cmd_help_exec
Case "reboot": OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + Strs.cmd_help_reboot
Case "exit": OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + Strs.cmd_help_exit
Case "run": OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + Strs.cmd_help_run
Case "clear": OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + Strs.cmd_help_clear
Case "version": OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + Strs.cmd_help_version
Case "about": OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + Strs.cmd_help_about
Case "time": OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + Strs.cmd_help_time
Case "copy": OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + Strs.cmd_help_copy
Case "delete": OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + Strs.cmd_help_delete
Case "makedir": OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + Strs.cmd_help_makedir
Case "hide": OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + Strs.cmd_help_hide
Case "read": OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + Strs.cmd_help_read
Case "write": OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + Strs.cmd_help_write

Case Else: OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + Strs.err_help + para
End Select
Else: GoTo UP
End If
Exit Sub
UP: OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + Strs.err_unknownparameter + tmp
End Sub
Private Sub cmd_run()
If gerNextWord_R = "" Then
    Dim exname As String
    exname = LCase(getExtendedName(para))
    If exname = "vbpos" Then
        On Error GoTo Err2
        Dim code As String
        Open para For Input As #1
        Do While Not EOF(1)
        Line Input #1, code
        Run code
        Loop
        Close #1
    ElseIf tmp = "" Then GoTo Err2
    Else: GoTo Err1
    End If
Else: GoTo UP
End If
If showSuccess = True Then
OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + Strs.suc_run + para
End If
Exit Sub
UP: OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + Strs.err_unknownparameter + tmp
Exit Sub
Err1: OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + Strs.err_run_1 + para
Exit Sub
Err2: OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + Strs.err_run_2 + para
End Sub
Private Sub cmd_reboot()
On Error GoTo Err
If para = "" Then
    Call reboot
ElseIf IsNumeric(para) = False Then
    If para = "long" Then
        getNextWord
        If IsNumeric(para) = True Then
            Dim longnum As Long
            longnum = para
            getNextWord
            If IsNumeric(para) = True Then
                If getNextWord_R = "" Then
                    On Error GoTo Err
                    InputForm.reboot_mode = 1
                    InputForm.RebootClock.Interval = Int(para)
                    InputForm.RebootClock.Enabled = True
                Else: GoTo UP
                End If
            End If
        Else: GoTo UP
        End If
    Else
        GoTo UP
    End If
Else
    If getNextWord_R = "" Then
    InputForm.reboot_mode = 0
    InputForm.RebootClock.Interval = para
    InputForm.RebootClock.Enabled = True
    Else: GoTo UP
    End If
End If
If showSuccess = True Then
OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + Strs.suc_reboot + para
End If
Exit Sub
UP: OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + Strs.err_unknownparameter + tmp
Exit Sub
Err: OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + Strs.err_reboot + para
End Sub
Private Sub cmd_exit()
On Error GoTo Err
If para = "" Then
    End
ElseIf IsNumeric(para) = False Then
    If para = "long" Then
        getNextWord
        If IsNumeric(para) = True Then
            Dim longnum As Long
            longnum = para
            getNextWord
            If IsNumeric(para) = True Then
                If getNextWord_R = "" Then
                    On Error GoTo Err
                    InputForm.exit_mode = 1
                    InputForm.ExitClock.Interval = Int(para)
                    InputForm.ExitClock.Enabled = True
                Else: GoTo UP
                End If
            End If
        Else: GoTo UP
        End If
    Else
        GoTo UP
    End If
Else
    If getNextWord_R = "" Then
    InputForm.exit_mode = 0
    InputForm.ExitClock.Interval = para
    InputForm.ExitClock.Enabled = True
    Else: GoTo UP
    End If
End If
If showSuccess = True Then
OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + Strs.suc_exit + para
End If
Exit Sub
UP: OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + Strs.err_unknownparameter + tmp
Exit Sub
Err: OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + Strs.err_exit + para
End Sub
Private Sub cmd_clear()
If para = "" Then
    OutputForm.outtxtbox.Text = ""
Else:
    If getNextWord_R = "" Then
    OutputForm.outtxtbox.Text = Replace(OutputForm.outtxtbox.Text, para, "")
    Else: GoTo UP
    End If
End If
Exit Sub
UP: OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + Strs.err_unknownparameter + tmp
End Sub
Private Sub cmd_about()
If para = "" Then
    OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + Strs.cmd_about
Else
    Select Case LCase(para)
    Case "vbpo": OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + Strs.cmd_about_vbpo
    Case "phppo": OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + Strs.cmd_about_phppo
    Case "javapo": OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + Strs.cmd_about_javapo
    Case "hsppo": OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + Strs.cmd_about_hsppo
    Case "naohisss": OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + Strs.cmd_about_naohisss
    Case Else: GoTo UP
    End Select
End If
Exit Sub
UP: OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + Strs.err_unknownparameter + para
End Sub
Private Sub cmd_version()
If para = "" Then
    OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + Strs.cmd_version + getVersion
Else: GoTo UP
End If
Exit Sub
UP: OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + Strs.err_unknownparameter + para
End Sub
Private Sub cmd_time()
If getNextWord_R = "" Then
    Select Case para
    Case "":
    Dim t
    t = str(Now())
    OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + Strs.cmd_time + t + " " + Strs.weekday_ + str(Weekday(t))
    Case "year": OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + Strs.cmd_time_year + str(Year(Date))
    Case "date": OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + Strs.cmd_time_date + str(Date)
    Case "time": OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + Strs.cmd_time_time + str(Time())
    Case "weekday": OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + Strs.cmd_time_weekday + str(Weekday(Date))
    Case Else: GoTo UP
    End Select
Else: GoTo UP2
End If
Exit Sub
UP: OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + Strs.err_unknownparameter + para
Exit Sub
UP2: OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + Strs.err_unknownparameter + tmp
End Sub
Private Sub cmd_copy()
On Error GoTo Err
If getNextWord_R = "to" Then
    If getNextWord_R = "" Then
        GoTo Err
    Else
        Dim srcpath As String
        srcpath = tmp
        Dim fileName As String
        If getNextWord_R = "" Then
            fileName = getLastName(para)
        ElseIf tmp = "as" Then
            fileName = getNextWord_R
            If fileName = "" Then
                fileName = getLastName(para)
            End If
        Else: GoTo UP2
        End If
        If createFolder(srcpath) = 0 Then GoTo Err
        If Right(srcpath, 1) = "\" Or Right(srcpath, 1) = "/" Then
            copy para, srcpath + fileName
        Else
            copy para, srcpath + "\" + fileName
        End If
    End If
Else
GoTo Err
End If
If showSuccess = True Then
OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + Strs.suc_copy + para
End If
Exit Sub
Err: OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + Strs.err_copy + para
Exit Sub
UP2: OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + Strs.err_unknownparameter + tmp
End Sub
Private Sub cmd_delete()
If getNextWord_R = "" Then
    On Error GoTo Err
    Kill para
Else: GoTo UP2
End If
If showSuccess = True Then
OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + Strs.suc_delete + para
End If
Exit Sub
UP2: OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + Strs.err_unknownparameter + tmp
Exit Sub
Err: OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + Strs.err_delete + para
End Sub
Private Sub cmd_makedir()
If getNextWord_R = "" Then
    If createFolder(para) = 0 Then
        GoTo Err
    End If
Else: GoTo UP2
End If
If showSuccess = True Then
OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + Strs.suc_makedir + para
End If
Exit Sub
UP2: OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + Strs.err_unknownparameter + tmp
Exit Sub
Err: OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + Strs.err_makedir + para
End Sub
Private Sub cmd_hide()
On Error GoTo Err
Select Case getNextWord_R
Case "true":
If getNextWord_R = "" Then
SetAttr para, vbHidden
Else: GoTo UP2
End If
Case "false":
If getNextWord_R = "" Then
SetAttr para, vbNormal
Else: GoTo UP2
End If
Case Else: GoTo UP2
End Select
If showSuccess = True Then
OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + Strs.suc_hide + para
End If
Exit Sub
UP2: OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + Strs.err_unknownparameter + tmp
Exit Sub
Err: OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + Strs.err_hide + para
End Sub
Private Sub cmd_read()
Dim file As String
On Error GoTo Err
Select Case LCase(para)
Case "utf8":
    file = getNextWord_R
    If getNextWord_R = "" Then
        OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + UTF8_Decode_File(file) + vbCrLf
    Else: GoTo UP2
    End If
Case Else:
    file = para
    If getNextWord_R = "" Then
        OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + n_read(file) + vbCrLf
    Else: GoTo UP2
    End If
End Select
If showSuccess = True Then
OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + Strs.suc_read + file
End If
Exit Sub
UP: OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + Strs.err_unknownparameter + para
Exit Sub
UP2: OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + Strs.err_unknownparameter + tmp
Exit Sub
Err: OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + Strs.err_read + file
End Sub
Private Sub cmd_write()



Exit Sub
UP2: OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + Strs.err_unknownparameter + tmp
Exit Sub
Err: OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + Strs.err_write + file
End Sub
'All functions are above......
Public Sub ReSet()
RunCmd.cmd = "": RunCmd.para = "": RunCmd.pos = -1
End Sub

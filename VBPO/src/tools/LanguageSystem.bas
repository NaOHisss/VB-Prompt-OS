Attribute VB_Name = "LanguageSystem"
Public Sub loadLanguage(path As String)
Dim lang() As String
Select Case SysLan
Case "EN_US": lang = readEachLine(path)
Case Else
lang = readEachLine_utf8(path)
End Select
Dim tmp As String
For i = 0 To UBound(lang)
tmp = lang(i)
setText getHashTable(tmp)
Next i
End Sub
Private Sub setText(lan As HashTable)
Select Case lan.Key
Case "vbpromptos": InputForm.Caption = lan.Value + "-v" + getVersion
Case "output": OutputForm.Caption = lan.Value
Case "bootingvbpo": LogoForm.Caption = lan.Value
Case "rebootvbpo": Strs.rebootvbpo = lan.Value
Case "file": InputForm.m_file.Caption = lan.Value + "(&F)"
Case "new": InputForm.m_file_new.Caption = lan.Value
Case "open": InputForm.m_file_open.Caption = lan.Value
Case "save": InputForm.m_file_save.Caption = lan.Value
Case "saveas": InputForm.m_file_saveas.Caption = lan.Value
Case "exit": InputForm.m_file_exit.Caption = lan.Value
Case "run": InputForm.m_run.Caption = lan.Value + "(&R)"
Case "runafile": InputForm.m_run_runafile.Caption = lan.Value
Case "runinputcode": InputForm.m_run_runinputcode.Caption = lan.Value
Case "edit": InputForm.m_edit.Caption = lan.Value + "(&E)"
Case "undo": InputForm.m_edit_undo.Caption = lan.Value
Case "copy": InputForm.m_edit_copy.Caption = lan.Value
Case "paste": InputForm.m_edit_paste.Caption = lan.Value
Case "options": InputForm.m_options.Caption = lan.Value + "(&O)"
Case "showsuccess": InputForm.m_options_showsuccess.Caption = lan.Value
Case "language": InputForm.m_options_language.Caption = lan.Value
Case "help": InputForm.m_help.Caption = lan.Value + "(&H)"
Case "about": InputForm.m_about.Caption = lan.Value + "(&A)"
Case "english": InputForm.m_options_language_english.Caption = lan.Value
Case "chinesecn": InputForm.m_options_language_chinesecn.Caption = lan.Value
Case "chinesetw": InputForm.m_options_language_chinesetw.Caption = lan.Value
Case "japanese": InputForm.m_options_language_japanese.Caption = lan.Value
Case "weekday": Strs.weekday_ = lan.Value

Case "language_reboot": Strs.language_reboot = lan.Value
Case "cmd_help": Strs.cmd_help = lan.Value
Case "cmd_help_help": Strs.cmd_help_help = lan.Value
Case "cmd_help_exec": Strs.cmd_help_exec = lan.Value
Case "cmd_help_reboot": Strs.cmd_help_reboot = lan.Value
Case "cmd_help_exit": Strs.cmd_help_exit = lan.Value
Case "cmd_help_run": Strs.cmd_help_run = lan.Value
Case "cmd_help_clear": Strs.cmd_help_clear = lan.Value
Case "cmd_help_about": Strs.cmd_help_about = lan.Value
Case "cmd_help_version": Strs.cmd_help_version = lan.Value
Case "cmd_help_time": Strs.cmd_help_time = lan.Value
Case "cmd_help_copy": Strs.cmd_help_copy = lan.Value
Case "cmd_help_delete": Strs.cmd_help_delete = lan.Value
Case "cmd_help_makedir": Strs.cmd_help_makedir = lan.Value
Case "cmd_help_hide": Strs.cmd_help_hide = lan.Value
Case "cmd_help_read": Strs.cmd_help_read = lan.Value
Case "cmd_help_write": Strs.cmd_help_write = lan.Value

Case "cmd_about": Strs.cmd_about = lan.Value
Case "cmd_version": Strs.cmd_version = lan.Value
Case "cmd_about_vbpo": Strs.cmd_about_vbpo = lan.Value
Case "cmd_about_phppo": Strs.cmd_about_phppo = lan.Value
Case "cmd_about_javapo": Strs.cmd_about_javapo = lan.Value
Case "cmd_about_hsppo": Strs.cmd_about_hsppo = lan.Value
Case "cmd_about_naohisss": Strs.cmd_about_naohisss = lan.Value
Case "cmd_time": Strs.cmd_time = lan.Value
Case "cmd_time_time": Strs.cmd_time_time = lan.Value
Case "cmd_time_date": Strs.cmd_time_date = lan.Value
Case "cmd_time_year": Strs.cmd_time_year = lan.Value
Case "cmd_time_weekday": Strs.cmd_time_weekday = lan.Value

Case "err_help": Strs.err_help = lan.Value
Case "err_unknowncommand": Strs.err_unknowncommand = lan.Value
Case "err_unknownparameter": Strs.err_unknownparameter = lan.Value
Case "err_exec": Strs.err_exec = lan.Value
Case "err_run_1": Strs.err_run_1 = lan.Value
Case "err_run_2": Strs.err_run_2 = lan.Value
Case "err_reboot": Strs.err_reboot = lan.Value
Case "err_exit": Strs.err_exit = lan.Value
Case "err_copy": Strs.err_copy = lan.Value
Case "err_delete": Strs.err_delete = lan.Value
Case "err_makedir": Strs.err_makedir = lan.Value
Case "err_hide": Strs.err_hide = lan.Value
Case "err_read": Strs.err_read = lan.Value
Case "err_write": Strs.err_write = lan.Value

Case "suc_exec": Strs.suc_exec = lan.Value
Case "suc_run": Strs.suc_run = lan.Value
Case "suc_reboot": Strs.suc_reboot = lan.Value
Case "suc_exit": Strs.suc_exit = lan.Value
Case "suc_copy": Strs.suc_copy = lan.Value
Case "suc_delete": Strs.suc_delete = lan.Value
Case "suc_makedir": Strs.suc_makedir = lan.Value
Case "suc_hide": Strs.suc_hide = lan.Value
Case "suc_read": Strs.suc_read = lan.Value
Case "suc_write": Strs.suc_write = lan.Value

End Select
End Sub

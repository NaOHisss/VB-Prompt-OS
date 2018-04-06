VERSION 5.00
Begin VB.Form InputForm 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VB Prompt OS"
   ClientHeight    =   7665
   ClientLeft      =   45
   ClientTop       =   690
   ClientWidth     =   12855
   FillColor       =   &H80000018&
   BeginProperty Font 
      Name            =   "Consolas"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000018&
   Icon            =   "InputForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   12855
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.Timer RebootClock 
      Enabled         =   0   'False
      Left            =   12720
      Top             =   0
   End
   Begin VB.Timer ExitClock 
      Enabled         =   0   'False
      Left            =   12600
      Top             =   0
   End
   Begin VB.TextBox inputtxtbox 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000005&
      Height          =   7695
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   12855
   End
   Begin VB.Menu m_file 
      Caption         =   "File(&F)"
      Begin VB.Menu m_file_new 
         Caption         =   "New"
         Shortcut        =   ^N
      End
      Begin VB.Menu m_file_open 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu m_file_save 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu m_file_saveas 
         Caption         =   "Save as..."
         Shortcut        =   ^A
      End
      Begin VB.Menu m_file_exit 
         Caption         =   "Exit"
         Shortcut        =   +{DEL}
      End
   End
   Begin VB.Menu m_run 
      Caption         =   "Run(&R)"
      Begin VB.Menu m_run_runafile 
         Caption         =   "Run a file"
         Shortcut        =   {F4}
      End
      Begin VB.Menu m_run_runinputcode 
         Caption         =   "Run input code"
         Shortcut        =   ^R
      End
   End
   Begin VB.Menu m_edit 
      Caption         =   "Edit(&E)"
      Begin VB.Menu m_edit_undo 
         Caption         =   "Undo"
         Shortcut        =   ^U
      End
      Begin VB.Menu m_edit_copy 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu m_edit_paste 
         Caption         =   "Paste"
         Shortcut        =   ^P
      End
   End
   Begin VB.Menu m_options 
      Caption         =   "Options(&O)"
      Begin VB.Menu m_options_showsuccess 
         Caption         =   "Show success"
         Checked         =   -1  'True
      End
      Begin VB.Menu m_options_language 
         Caption         =   "Language"
         Begin VB.Menu m_options_language_english 
            Caption         =   "English"
            Checked         =   -1  'True
         End
         Begin VB.Menu m_options_language_chinesecn 
            Caption         =   "Chinese(Simplified)"
            Checked         =   -1  'True
         End
         Begin VB.Menu m_options_language_chinesetw 
            Caption         =   "Chinese(Traditional)"
            Checked         =   -1  'True
         End
         Begin VB.Menu m_options_language_japanese 
            Caption         =   "Japanese"
            Checked         =   -1  'True
         End
      End
   End
   Begin VB.Menu m_help 
      Caption         =   "Help(&H)"
   End
   Begin VB.Menu m_about 
      Caption         =   "About(&A)"
   End
End
Attribute VB_Name = "InputForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private nowTime_reboot As Long
Private nowTime_exit As Long
Private sTime_reboot As Long
Private sTime_exit As Long
Public exit_mode As Integer
Public reboot_mode As Integer

Private Sub ExitClock_Timer()
On Error GoTo Err
If exit_mode = 1 Then
If sTime_exit = 0 Then sTime_exit = para
If nowTime_exit > sTime_exit Then
End
End If
nowTime_exit = nowTime_exit + ExitClock.Interval
Else
End
End If
Exit Sub
Err: OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + Strs.err_exit + para
End Sub

Private Sub Form_Load()
Load OutputForm
If showSuccess = True Then
m_options_showsuccess.Checked = True
Else: m_options_showsuccess.Checked = False
End If
Select Case SysLan
Case "EN_US": m_options_language_english.Checked = True
m_options_language_chinesecn.Checked = False
m_options_language_chinesetw.Checked = False
m_options_language_japanese.Checked = False
Case "ZH_CN": m_options_language_chinesecn.Checked = True
m_options_language_english.Checked = False
m_options_language_chinesetw.Checked = False
m_options_language_japanese.Checked = False
Case "ZH_TW": m_options_language_chinesetw.Checked = True
m_options_language_english.Checked = False
m_options_language_chinesecn.Checked = False
m_options_language_japanese.Checked = False
Case "JP": m_options_language_japanese.Checked = True
m_options_language_english.Checked = False
m_options_language_chinesecn.Checked = False
m_options_language_chinesetw.Checked = False
End Select
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
End
End Sub

Private Sub m_file_exit_Click()
End
End Sub

Private Sub m_options_language_chinesecn_Click()
If m_options_language_chinesecn.Checked = True Then
Exit Sub
End If
m_options_language_english.Checked = False
m_options_language_chinesecn.Checked = True
m_options_language_chinesetw.Checked = False
m_options_language_japanese.Checked = False
writeTo "ZH_CN", App.path + "\system\options\language.data"
msgReBoot Strs.language_reboot
End Sub

Private Sub m_options_language_chinesetw_Click()
If m_options_language_chinesetw.Checked = True Then
Exit Sub
End If
m_options_language_english.Checked = False
m_options_language_chinesecn.Checked = False
m_options_language_chinesetw.Checked = True
m_options_language_japanese.Checked = False
writeTo "ZH_TW", App.path + "\system\options\language.data"
msgReBoot Strs.language_reboot
End Sub

Private Sub m_options_language_english_Click()
If m_options_language_english.Checked = True Then
Exit Sub
End If
m_options_language_english.Checked = True
m_options_language_chinesecn.Checked = False
m_options_language_chinesetw.Checked = False
m_options_language_japanese.Checked = False
writeTo "EN_US", App.path + "\system\options\language.data"
msgReBoot Strs.language_reboot
End Sub

Private Sub m_options_language_japanese_Click()
If m_options_language_japanese.Checked = True Then
Exit Sub
End If
m_options_language_english.Checked = False
m_options_language_chinesecn.Checked = False
m_options_language_chinesetw.Checked = False
m_options_language_japanese.Checked = True
writeTo "JP", App.path + "\system\options\language.data"
msgReBoot Strs.language_reboot
End Sub

Private Sub m_options_showsuccess_Click()
If m_options_showsuccess.Checked = True Then
m_options_showsuccess.Checked = False
writeTo "false", App.path + "\system\options\showsuccess.data"
showSuccess = False
Else: m_options_showsuccess.Checked = True
writeTo "true", App.path + "\system\options\showsuccess.data"
showSuccess = True
End If
End Sub

Private Sub m_run_runinputcode_Click()
OutputForm.Show
Dim c() As String
c = getEachText(inputtxtbox.Text)
For i = 0 To UBound(c)
Run c(i)
Next i
End Sub

Private Sub RebootClock_Timer()
On Error GoTo Err
If reboot_mode = 1 Then
    If sTime_reboot = 0 Then sTime_reboot = para
    If nowTime_reboot > sTime_reboot Then
    Call reboot
    End If
    nowTime_reboot = nowTime_reboot + RebootClock.Interval
Else
Call reboot
End If
Exit Sub
Err: OutputForm.outtxtbox.Text = OutputForm.outtxtbox.Text + Strs.err_reboot + para
End Sub

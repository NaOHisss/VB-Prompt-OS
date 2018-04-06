VERSION 5.00
Begin VB.Form LogoForm 
   BorderStyle     =   0  'None
   Caption         =   "Booting VBPO..."
   ClientHeight    =   3975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6240
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Consolas"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000018&
   Icon            =   "LogoForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5280
      Top             =   120
   End
   Begin VB.Timer Timer 
      Interval        =   2000
      Left            =   5760
      Top             =   120
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright 2017 PO Project,Naohisss.All rights reserved."
      ForeColor       =   &H80000018&
      Height          =   210
      Left            =   120
      TabIndex        =   1
      Top             =   3480
      Width           =   5775
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Visual Basic Prompt Operating System"
      ForeColor       =   &H80000018&
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   3120
      Width           =   3780
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   3975
      Left            =   0
      Picture         =   "LogoForm.frx":C84A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6255
   End
End
Attribute VB_Name = "LogoForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Timer1.Enabled = False
loadOptions
On Error GoTo loadeng
back: loadLanguage App.path + "\system\languages\" + SysLan + ".lang"
Exit Sub
loadeng: writeTo "EN_US", App.path + "\system\options\language.data"
GoTo back
End Sub

Private Sub Timer_Timer()
Timer.Enabled = False
LogoForm.Print ""
LogoForm.Print " Loading..."
Timer1.Enabled = True
Load InputForm
End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False
LogoForm.Print " Ready to start..."
InputForm.Show
Unload Me
End Sub

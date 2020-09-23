VERSION 5.00
Begin VB.Form FrmLock 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Enter Password"
   ClientHeight    =   405
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3765
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmLock.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   405
   ScaleWidth      =   3765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Tpass 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   0
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   0
      Width           =   3735
   End
End
Attribute VB_Name = "FrmLock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public SHORT_DET As Boolean
Public SHORT_DET_X As Boolean
Private Function Loader(Pass As String) As Boolean
If Pass = WHOPASS Then
    Loader = True
Else
    Loader = False
    MsgBox "Password Not Recognized.", vbCritical, "ERROR"
End If
Tpass.Text = ""
If SHORT_DET_X = False Then
    SHORT_DET = Loader
    Unload Me
Else
    If Loader = True Then
        SHORT_DET_X = False
        Unload Me
    End If
End If
End Function

Private Sub Form_Load()
Me.Caption = "[User:" & WHOLOG & "][Password Required:]"
End Sub

Private Sub Tpass_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Call Loader(Tpass.Text)
End Sub

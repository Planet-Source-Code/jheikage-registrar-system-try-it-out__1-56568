VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmuser 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "LAST USER"
   ClientHeight    =   1440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ISAPTECH.chameleonButton bload 
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   840
      Width           =   1455
      _extentx        =   2566
      _extenty        =   873
      und             =   0
      iname           =   "Tahoma"
      btype           =   14
      tx              =   "OK"
      enab            =   -1
      coltype         =   1
      focusr          =   -1
      bcol            =   15329769
      bcolo           =   15329769
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      umcol           =   -1
      soft            =   0
      picpos          =   0
      ngrey           =   0
      fx              =   0
      check           =   0
      value           =   0
   End
   Begin VB.TextBox Tuser 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   0
      Width           =   4695
   End
   Begin RichTextLib.RichTextBox rtbox 
      Height          =   855
      Left            =   -120
      TabIndex        =   0
      Top             =   1440
      Visible         =   0   'False
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   1508
      _Version        =   393217
      Enabled         =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmuser.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmuser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Bload_Click()
Unload Me
End Sub

Sub GetFileName(FIL As String)
rtbox.LoadFile FIL
Regon
End Sub
Sub Regon()
Dim i As Integer
Dim STRK As String, gik As String
Dim bb As Integer, com1 As String, com2 As String
bb = Len(rtbox.Text)
STRK = rtbox.Text
com1 = Chr(234)
com2 = Chr(34)
For i = 1 To bb
gik = Left(STRK, 1)
'MsgBox gik
STRK = Right(STRK, bb - i)
If gik = com1 Or gik = com2 Then
Else
Tuser.Text = Tuser.Text + gik
End If
Next
End Sub


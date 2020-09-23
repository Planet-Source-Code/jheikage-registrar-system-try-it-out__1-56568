VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FrmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About Registrar Information System"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5670
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ISAPTECH.chameleonButton CBOk 
      Height          =   615
      Left            =   3960
      TabIndex        =   3
      Top             =   2760
      Width           =   1575
      _extentx        =   2778
      _extenty        =   1085
      und             =   0   'False
      iname           =   "Tahoma"
      btype           =   14
      tx              =   "OK"
      enab            =   -1  'True
      coltype         =   2
      focusr          =   -1  'True
      bcol            =   12632256
      bcolo           =   8421504
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   0
      ngrey           =   0   'False
      fx              =   0
      check           =   0   'False
      value           =   0   'False
   End
   Begin RichTextLib.RichTextBox Rfil 
      Height          =   1575
      Left            =   1800
      TabIndex        =   2
      Top             =   960
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   2778
      _Version        =   393217
      BackColor       =   -2147483633
      BorderStyle     =   0
      Enabled         =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"FrmAbout.frx":57E2
   End
   Begin VB.Label LBLVersion 
      AutoSize        =   -1  'True
      Caption         =   "Version:"
      Height          =   240
      Left            =   1680
      TabIndex        =   1
      Top             =   480
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Registrar Information System"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   3630
   End
   Begin VB.Image Image1 
      Height          =   1320
      Left            =   120
      Picture         =   "FrmAbout.frx":585E
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1440
   End
End
Attribute VB_Name = "FrmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CBOk_Click()
Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo Handlers
Me.Rfil.LoadFile App.Path & "\Help\Cd._"
LBLVersion.Caption = App.Major & ":" & App.Minor & ":" & App.Revision
Exit Sub
Handlers:
    Rfil.Text = "Can't Locate file..."
End Sub

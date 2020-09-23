VERSION 5.00
Begin VB.Form FrmNamI 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1290
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4515
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   4515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox t3 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   395
      Left            =   720
      TabIndex        =   0
      Top             =   840
      Width           =   2775
   End
   Begin ISAPTECH.chameleonButton CBOK 
      Height          =   495
      Left            =   3600
      TabIndex        =   2
      Top             =   120
      Width           =   855
      _extentx        =   1508
      _extenty        =   873
      bold            =   0   'False
      ita             =   0   'False
      iname           =   "Tahoma"
      size            =   8.25
      und             =   0   'False
      btype           =   6
      tx              =   "OK"
      enab            =   -1  'True
      font            =   "FrmNamI.frx":0000
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   8421504
      bcolo           =   14737632
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   0
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin ISAPTECH.chameleonButton CBCANCEL 
      Height          =   495
      Left            =   3600
      TabIndex        =   3
      Top             =   720
      Width           =   855
      _extentx        =   1508
      _extenty        =   873
      bold            =   0   'False
      ita             =   0   'False
      iname           =   "Tahoma"
      size            =   8.25
      und             =   0   'False
      btype           =   6
      tx              =   "CANCEL"
      enab            =   -1  'True
      font            =   "FrmNamI.frx":002C
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   8421504
      bcolo           =   14737632
      fcol            =   0
      fcolo           =   0
      mcol            =   12632256
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   0
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "Search Option"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   3375
      Begin VB.OptionButton O2 
         Appearance      =   0  'Flat
         Caption         =   "Find By ID"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   1815
      End
      Begin VB.OptionButton O1 
         Appearance      =   0  'Flat
         Caption         =   "Find By Name"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   1815
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Value:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   555
   End
End
Attribute VB_Name = "FrmNamI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbCancel_Click()
Srch = ""
Unload Me
End Sub

Private Sub cbOK_Click()
With FrmInfoCNTR
Dim x As ListItem
If o1.Value = True Then     'Search by name
SRCHTYP = True
Else
SRCHTYP = False
End If
Srch = t3.Text
End With
Unload Me
End Sub

Function Trimer(Source As Object)
Source.Text = Trim(CONVS)
End Function

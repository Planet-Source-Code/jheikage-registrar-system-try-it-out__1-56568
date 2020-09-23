VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FrmSet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ISAP-MCNP Registrar"
   ClientHeight    =   7155
   ClientLeft      =   180
   ClientTop       =   2085
   ClientWidth     =   11745
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmSet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   11745
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrOut 
      Caption         =   "System Log Window"
      Height          =   1575
      Left            =   120
      TabIndex        =   25
      Top             =   5160
      Width           =   11535
      Begin RichTextLib.RichTextBox Routputbox 
         Height          =   1215
         Left            =   120
         TabIndex        =   26
         ToolTipText     =   "Out put Window"
         Top             =   240
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   2143
         _Version        =   393217
         BackColor       =   16777215
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"FrmSet.frx":57E2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame F1 
      Caption         =   "System Settings"
      Height          =   2775
      Left            =   120
      TabIndex        =   16
      Top             =   2280
      Width           =   3975
      Begin ISAPTECH.chameleonButton cbLock 
         Height          =   375
         Left            =   960
         TabIndex        =   24
         Top             =   2160
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   661
         BOLD            =   0   'False
         ITA             =   0   'False
         INAME           =   "Tahoma"
         SIZE            =   8.25
         UND             =   0   'False
         BTYPE           =   9
         TX              =   "Lock Application"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   0   'False
         BCOL            =   15329769
         BCOLO           =   15329769
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   2
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ISAPTECH.chameleonButton cbHelp 
         Height          =   615
         Left            =   840
         TabIndex        =   4
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   1085
         BOLD            =   0   'False
         ITA             =   0   'False
         INAME           =   "Tahoma"
         SIZE            =   8.25
         UND             =   0   'False
         BTYPE           =   9
         TX              =   "Help File"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   0   'False
         BCOL            =   15329769
         BCOLO           =   15329769
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ISAPTECH.chameleonButton CBconfig 
         Height          =   615
         Left            =   840
         TabIndex        =   3
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   1085
         BOLD            =   0   'False
         ITA             =   0   'False
         INAME           =   "Tahoma"
         SIZE            =   8.25
         UND             =   0   'False
         BTYPE           =   9
         TX              =   "Connection Configuration"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   0   'False
         BCOL            =   15329769
         BCOLO           =   15329769
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   240
         X2              =   3840
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         X1              =   240
         X2              =   3840
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Image Image9 
         Height          =   615
         Left            =   120
         Picture         =   "FrmSet.frx":586C
         Stretch         =   -1  'True
         Top             =   2040
         Width           =   735
      End
      Begin VB.Image Image7 
         Height          =   615
         Left            =   120
         Picture         =   "FrmSet.frx":5CAE
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   735
      End
      Begin VB.Image Image6 
         Height          =   615
         Left            =   0
         Picture         =   "FrmSet.frx":BF38
         Stretch         =   -1  'True
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "View Help File."
         Height          =   615
         Left            =   2280
         TabIndex        =   23
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label9 
         Caption         =   "Change your Connection Settings."
         Height          =   735
         Left            =   2280
         TabIndex        =   22
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame F4 
      Caption         =   "Remote User Commands (Admin and Clients)"
      Height          =   2775
      Left            =   4200
      TabIndex        =   15
      Top             =   2280
      Width           =   7455
      Begin ISAPTECH.chameleonButton CBINFOR 
         Height          =   615
         Left            =   960
         TabIndex        =   9
         Top             =   1800
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   1085
         BOLD            =   0   'False
         ITA             =   0   'False
         INAME           =   "Tahoma"
         SIZE            =   8.25
         UND             =   0   'False
         BTYPE           =   9
         TX              =   "INFORMATION CENTER"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   0   'False
         BCOL            =   8421504
         BCOLO           =   14737632
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ISAPTECH.chameleonButton CbTransfer 
         Height          =   615
         Left            =   960
         TabIndex        =   8
         Top             =   1080
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   1085
         BOLD            =   0   'False
         ITA             =   0   'False
         INAME           =   "Tahoma"
         SIZE            =   8.25
         UND             =   0   'False
         BTYPE           =   9
         TX              =   "DATA TRANSER"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   0   'False
         BCOL            =   8421504
         BCOLO           =   14737632
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ISAPTECH.chameleonButton CBSTUDENT 
         Height          =   615
         Left            =   960
         TabIndex        =   7
         Top             =   360
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   1085
         BOLD            =   0   'False
         ITA             =   0   'False
         INAME           =   "Tahoma"
         SIZE            =   8.25
         UND             =   0   'False
         BTYPE           =   9
         TX              =   "CLASSES/CS REPORTS"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   0   'False
         BCOL            =   8421504
         BCOLO           =   14737632
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Image Image5 
         Height          =   495
         Left            =   240
         Picture         =   "FrmSet.frx":1171A
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   615
      End
      Begin VB.Image Image4 
         Height          =   495
         Left            =   240
         Picture         =   "FrmSet.frx":16EFC
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   495
      End
      Begin VB.Image Image3 
         Height          =   615
         Left            =   120
         Picture         =   "FrmSet.frx":1757E
         Stretch         =   -1  'True
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Student Information Center. View Grades, Personal Information, Add Student and Create TOR."
         Height          =   735
         Left            =   2760
         TabIndex        =   21
         Top             =   1800
         Width           =   4575
      End
      Begin VB.Label Label7 
         Caption         =   "Transfer Data to and from the Database engine."
         Height          =   615
         Left            =   2760
         TabIndex        =   20
         Top             =   1080
         Width           =   4575
      End
      Begin VB.Label Label6 
         Caption         =   "Create Reports for Classes/ Control Sheets and Custom Reports."
         Height          =   615
         Left            =   2760
         TabIndex        =   19
         Top             =   360
         Width           =   4575
      End
   End
   Begin VB.Frame F3 
      Caption         =   "Administrator Commands (For Admin use only)"
      Height          =   2175
      Left            =   4200
      TabIndex        =   14
      Top             =   0
      Width           =   7455
      Begin ISAPTECH.chameleonButton cbSettings 
         Height          =   615
         Left            =   960
         TabIndex        =   6
         Top             =   1080
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   1085
         BOLD            =   0   'False
         ITA             =   0   'False
         INAME           =   "Tahoma"
         SIZE            =   8.25
         UND             =   0   'False
         BTYPE           =   9
         TX              =   "LIMITS"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   0   'False
         BCOL            =   8421504
         BCOLO           =   14737632
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ISAPTECH.chameleonButton CBCRS 
         Height          =   615
         Left            =   960
         TabIndex        =   5
         Top             =   360
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   1085
         BOLD            =   0   'False
         ITA             =   0   'False
         INAME           =   "Tahoma"
         SIZE            =   8.25
         UND             =   0   'False
         BTYPE           =   9
         TX              =   "COURSES AND CURRICULA"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   0   'False
         BCOL            =   8421504
         BCOLO           =   14737632
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Image Image1 
         Height          =   615
         Left            =   120
         Picture         =   "FrmSet.frx":181E4
         Stretch         =   -1  'True
         Top             =   1200
         Width           =   735
      End
      Begin VB.Image Image2 
         Height          =   615
         Left            =   120
         Picture         =   "FrmSet.frx":18F7A
         Stretch         =   -1  'True
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Set Limits for Classes and Sections."
         Height          =   615
         Left            =   2760
         TabIndex        =   18
         Top             =   1080
         Width           =   4575
      End
      Begin VB.Label Label2 
         Caption         =   "Add/Modify Course offered in the School and their Curricula."
         Height          =   615
         Left            =   2760
         TabIndex        =   17
         Top             =   360
         Width           =   4575
      End
   End
   Begin VB.Frame F2 
      Caption         =   "Database Security Connection Window"
      Height          =   2175
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   3975
      Begin ISAPTECH.chameleonButton CBCOn 
         Height          =   495
         Left            =   2640
         TabIndex        =   2
         Top             =   1440
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         BOLD            =   0   'False
         ITA             =   0   'False
         INAME           =   "Tahoma"
         SIZE            =   8.25
         UND             =   0   'False
         BTYPE           =   14
         TX              =   "Connect"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   8421504
         BCOLO           =   8421504
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox TDBPass 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1200
         PasswordChar    =   "*"
         TabIndex        =   1
         ToolTipText     =   "User Name"
         Top             =   960
         Width           =   2535
      End
      Begin VB.TextBox TDBUser 
         Height          =   375
         Left            =   1200
         TabIndex        =   0
         ToolTipText     =   "User Name"
         Top             =   480
         Width           =   2535
      End
      Begin VB.Image Image8 
         Height          =   615
         Left            =   240
         Picture         =   "FrmSet.frx":19CBA
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Password:"
         Height          =   240
         Left            =   120
         TabIndex        =   13
         Top             =   960
         Width           =   900
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "User Name:"
         Height          =   240
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   1020
      End
   End
   Begin MSComctlLib.StatusBar Sbar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   6780
      Width           =   11745
      _ExtentX        =   20717
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   3598
            MinWidth        =   1235
            Picture         =   "FrmSet.frx":1FF44
            Text            =   "Registrar System"
            TextSave        =   "Registrar System"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "6/25/2004"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14473
            Picture         =   "FrmSet.frx":25736
            Text            =   "System Administrator"
            TextSave        =   "System Administrator"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu MFile 
      Caption         =   "&File"
      Begin VB.Menu FConCon 
         Caption         =   "&Connection Configuration"
         Shortcut        =   ^C
      End
      Begin VB.Menu FLock 
         Caption         =   "&Lock"
         Shortcut        =   ^L
      End
      Begin VB.Menu FBreak 
         Caption         =   "-"
      End
      Begin VB.Menu FExit 
         Caption         =   "&Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu MPanels 
      Caption         =   "&View"
      Begin VB.Menu FDSC 
         Caption         =   "&Database Security Connection"
         Checked         =   -1  'True
      End
      Begin VB.Menu FSet 
         Caption         =   "&System Settings"
         Checked         =   -1  'True
      End
      Begin VB.Menu FSLW 
         Caption         =   "S&ystem Log Window"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu Mwindows 
      Caption         =   "&Windows"
      Begin VB.Menu WCrs 
         Caption         =   "Courses"
         Visible         =   0   'False
      End
      Begin VB.Menu WDT 
         Caption         =   "&Data Transfer"
         Visible         =   0   'False
      End
      Begin VB.Menu WSIC 
         Caption         =   "&Student Information Center"
         Visible         =   0   'False
      End
      Begin VB.Menu WCR 
         Caption         =   "C&lass Reports"
         Visible         =   0   'False
      End
      Begin VB.Menu WLmt 
         Caption         =   "&Limits"
         Visible         =   0   'False
      End
      Begin VB.Menu WBR 
         Caption         =   "-"
      End
      Begin VB.Menu WMain 
         Caption         =   "&Main Window"
      End
   End
   Begin VB.Menu Mhelp 
      Caption         =   "&Help"
      Begin VB.Menu FRISH 
         Caption         =   "&Registrar Information System Help"
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu HBr 
         Caption         =   "-"
      End
      Begin VB.Menu HAbout 
         Caption         =   "&About Registrar Information System"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "FrmSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CBCOn_Click()
If Trim(TDBUser.Text) = "" Or Trim(TDBPass.Text) = "" Then
MsgBox "Database User or password missing.", vbCritical, "Error Connecting"
TDBUser.SetFocus
SendKeys "{Home}+{END}"
Exit Sub
End If
SetModule.ConnectSQLSERVER TDBUser.Text, TDBPass.Text
End Sub

Private Sub CBconfig_Click()
FrmLock.Show 1
    If FrmLock.SHORT_DET = True Then FrmConfig.Show 1, Me
End Sub

Private Sub CBCRS_Click()
If isLogDb = True Then
    FrmLock.Show 1
    If FrmLock.SHORT_DET = True Then
        FrmCCUr.Show
        WCrs.Visible = True
    End If
End If
End Sub

Private Sub cbHelp_Click()
On Error GoTo ErbX
Dim strx As String
strx = App.Path & "\Help\UIH.bat"
Shell strx
Exit Sub
ErbX:
    ErrorTrap Err, "Help File"
End Sub

Private Sub CBINFOR_Click()
If isLogDb = True Then
FrmLock.Show 1
    If FrmLock.SHORT_DET = True Then
        FrmInfoCNTR.Show
        Me.WSIC.Visible = True
    End If
End If
End Sub

Private Sub cbLock_Click()
FrmLock.SHORT_DET_X = True
FrmLock.Show 1
End Sub

Private Sub cbSettings_Click()
If isLogDb = True Then
FrmLock.Show 1
    If FrmLock.SHORT_DET = True Then
        FrmOthers.Show
        WLmt.Visible = True
    End If
End If
End Sub

Private Sub CBSTUDENT_Click()
If isLogDb = True Then
FrmLock.Show 1
    If FrmLock.SHORT_DET = True Then
        FrmControls.Show
        WCR.Visible = True
    End If
End If
End Sub

Private Sub cbTransfer_Click()
If isLogDb = True Then
FrmLock.Show 1
    If FrmLock.SHORT_DET = True Then
        FrmTrans.Show
        Me.WDT.Visible = True
    End If
End If
End Sub

Private Sub FConCon_Click()
CBconfig_Click
End Sub

Private Sub FDSC_Click()
FDSC.Checked = Not FDSC.Checked
F2.Visible = FDSC.Checked
End Sub

Private Sub FExit_Click()
EndSys
End Sub

Private Sub FLock_Click()
cbLock_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
EndSys
End Sub


Private Sub FRISH_Click()
cbHelp_Click
End Sub

Private Sub FSet_Click()
FSet.Checked = Not FSet.Checked
F1.Visible = FSet.Checked
End Sub

Private Sub FSLW_Click()
FSLW.Checked = Not FSLW.Checked
FrOut.Visible = FSLW.Checked
End Sub


Private Sub HAbout_Click()
FrmAbout.Show 1
End Sub

Private Sub Routputbox_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 2 And Shift = 5 Then XXX.SetVal
End Sub

Private Sub WMain_Click()
FrmSet.Show
End Sub

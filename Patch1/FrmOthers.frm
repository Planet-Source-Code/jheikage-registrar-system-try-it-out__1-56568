VERSION 5.00
Begin VB.Form FrmOthers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Set Student Limits"
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6240
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmOthers.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   6240
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CBDel 
      Height          =   495
      Left            =   960
      Picture         =   "FrmOthers.frx":57E2
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Delete"
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton CBSave 
      Height          =   495
      Left            =   480
      Picture         =   "FrmOthers.frx":596C
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Save Record"
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton CBADD 
      Height          =   495
      Left            =   0
      Picture         =   "FrmOthers.frx":5FD6
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "New Record"
      Top             =   0
      Width           =   495
   End
   Begin VB.ComboBox CSYR 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2280
      TabIndex        =   4
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox TCLASS 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   1440
      Width           =   975
   End
   Begin VB.TextBox TSUBS 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   1080
      Width           =   975
   End
   Begin VB.ComboBox CSEM 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "FrmOthers.frx":6640
      Left            =   3600
      List            =   "FrmOthers.frx":664D
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   720
      Width           =   975
   End
   Begin ISAPTECH.chameleonButton CBF 
      Height          =   375
      Left            =   3600
      TabIndex        =   10
      Top             =   1440
      Width           =   375
      _extentx        =   450
      _extenty        =   661
      und             =   0
      iname           =   "Tahoma"
      btype           =   14
      tx              =   "|<"
      enab            =   -1
      coltype         =   1
      focusr          =   -1
      bcol            =   8421504
      bcolo           =   14737632
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
   Begin ISAPTECH.chameleonButton CBPR 
      Height          =   375
      Left            =   3960
      TabIndex        =   11
      Top             =   1440
      Width           =   375
      _extentx        =   450
      _extenty        =   661
      und             =   0
      iname           =   "Tahoma"
      btype           =   14
      tx              =   "<"
      enab            =   -1
      coltype         =   1
      focusr          =   -1
      bcol            =   8421504
      bcolo           =   14737632
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
   Begin ISAPTECH.chameleonButton CBNX 
      Height          =   375
      Left            =   5400
      TabIndex        =   12
      Top             =   1440
      Width           =   375
      _extentx        =   450
      _extenty        =   661
      und             =   0
      iname           =   "Tahoma"
      btype           =   14
      tx              =   ">"
      enab            =   -1
      coltype         =   1
      focusr          =   -1
      bcol            =   8421504
      bcolo           =   14737632
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
   Begin ISAPTECH.chameleonButton CBL 
      Height          =   375
      Left            =   5760
      TabIndex        =   13
      Top             =   1440
      Width           =   375
      _extentx        =   450
      _extenty        =   661
      und             =   0
      iname           =   "Tahoma"
      btype           =   14
      tx              =   ">|"
      enab            =   -1
      coltype         =   1
      focusr          =   -1
      bcol            =   8421504
      bcolo           =   14737632
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
   Begin VB.Label LBLREC 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Records"
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
      Height          =   375
      Left            =   4320
      TabIndex        =   14
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000001&
      Index           =   1
      X1              =   0
      X2              =   6240
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      BorderWidth     =   2
      Index           =   0
      X1              =   0
      X2              =   6240
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Student/Class Limit:"
      Height          =   240
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   1725
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Student/Subject Limit:"
      Height          =   240
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   1920
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "School Year/Semester:"
      Height          =   240
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   1995
   End
End
Attribute VB_Name = "FrmOthers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsMix As ADODB.Recordset
Dim Adding As Boolean

Private Sub CBADD_Click()
On Error GoTo ErrAdd
Adding = True
ClearAll
Ena False
CSYR.SetFocus
 Me.LBLREC.Caption = "Adding..."
Exit Sub
ErrAdd:
    ErrorTrap Err, "Adding Records"
    Adding = False
    Ena True
End Sub

Sub ClearAll()
CSYR.Text = ""
TCLASS.Text = ""
TSUBS.Text = ""
End Sub

Sub Ena(Bool As Boolean)
CBADD.Enabled = Bool
Me.CBF.Enabled = Bool
Me.CBL.Enabled = Bool
Me.CBNX.Enabled = Bool
Me.CBPR.Enabled = Bool
If Bool = True Then
    CBDEL.ToolTipText = "Delete"
Else
    CBDEL.ToolTipText = "Cancel"
End If
End Sub

Private Sub CBDel_Click()
On Error GoTo XP
Select Case UCase(CBDEL.ToolTipText)
Case UCase("Delete")
    If MsgBox("Delete Record?", vbYesNo + vbQuestion, "Delete") = vbYes Then
    RsMix.Delete
    End If
Case UCase("Cancel")
    RsMix.Cancel
End Select
    RsMix.Properties.Refresh
    If RsMix.RecordCount <> 0 Then
    CBL_Click
    Else
    ClearAll
    End If
    Navi
    Ena True
Exit Sub
XP:
    ErrorTrap Err, "Cancel/Delete Command"
    Me.Refresh
    Ena True
End Sub

Private Sub CBF_Click()
With RsMix
    .MoveFirst
    Navi
End With
End Sub

Private Sub CBL_Click()
With RsMix
    If .RecordCount = 0 Then Exit Sub
    .MoveLast
    Navi
End With
End Sub

Private Sub CBNX_Click()
With RsMix
    If .RecordCount = 0 Then Exit Sub
    .MoveNext
    If .EOF Then .MoveLast
    Navi
End With
End Sub

Private Sub CBPR_Click()
With RsMix
    If .RecordCount = 0 Then Exit Sub
    .MovePrevious
    If .BOF Then .MoveFirst
    Navi
End With
End Sub

Private Sub CBSave_Click()
On Error GoTo XP
Select Case Adding
    Case True
        RsMix.AddNew
End Select
Me.SetVal
RsMix.Update
RsMix.Properties.Refresh
CBL_Click
Ena True
Adding = False
Exit Sub
XP:
    ErrorTrap Err, "Save Changes"
    Me.Refresh
    Ena True
End Sub

Private Sub Form_Load()
SetRS
End Sub

Sub SetRS()
Set RsMix = Nothing
Set RsMix = New ADODB.Recordset
With RsMix
    .ActiveConnection = FrmInfoCNTR.ConX
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open "OTHERS"
    Navi
End With
End Sub

Sub Navi()
With RsMix
    If .RecordCount <> 0 Then
    CSYR.Text = .Fields("SCHOOLYEAR").Value
    CSEM.Text = .Fields("SEMESTER").Value
    TCLASS.Text = .Fields("STUDENT_PER_CLASS").Value
    TSUBS.Text = .Fields("STUDENT_PER_SUBJECT").Value
    Me.LBLREC.Caption = .AbsolutePosition & " of " & .RecordCount
    Else
    Me.LBLREC.Caption = "No Records"
    End If
End With
End Sub

Sub SetVal()
With RsMix
    .Fields("SCHOOLYEAR").Value = CSYR.Text
    .Fields("SEMESTER").Value = CSEM.Text
    .Fields("STUDENT_PER_CLASS").Value = TCLASS.Text
    .Fields("STUDENT_PER_SUBJECT").Value = TSUBS.Text
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
FrmSet.WLmt.Visible = False
End Sub

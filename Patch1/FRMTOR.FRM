VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FRMTOR 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transcript of Record"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9960
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FRMTOR.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   9960
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CDL 
      Left            =   960
      Top             =   4200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   240
      Top             =   4560
   End
   Begin VB.TextBox TDesc 
      Height          =   720
      Left            =   1440
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   24
      Top             =   4560
      Width           =   3855
   End
   Begin ISAPTECH.chameleonButton cbWrite 
      Height          =   615
      Left            =   8160
      TabIndex        =   23
      Top             =   5160
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1085
      BOLD            =   0   'False
      ITA             =   0   'False
      INAME           =   "Tahoma"
      SIZE            =   8.25
      UND             =   0   'False
      BTYPE           =   14
      TX              =   "Write TOR"
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
   Begin MSComctlLib.ListView LSYR 
      Height          =   1695
      Left            =   5640
      TabIndex        =   21
      Top             =   720
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   2990
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "cc"
      SmallIcons      =   "cc"
      ColHdrIcons     =   "cc"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "SCHOOL"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "SCHOOLYEAR"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "SEMESTER"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "COURSE"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.ComboBox CCOURSE 
      Height          =   360
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   4200
      Width           =   1455
   End
   Begin VB.TextBox TCred 
      Height          =   720
      Left            =   1440
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   19
      Top             =   3480
      Width           =   3855
   End
   Begin VB.TextBox TGrad 
      Height          =   360
      Left            =   1800
      TabIndex        =   18
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox TAdmission 
      Height          =   360
      Left            =   4080
      TabIndex        =   17
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox TGEN 
      Height          =   360
      Left            =   1440
      TabIndex        =   16
      Top             =   2400
      Width           =   855
   End
   Begin VB.TextBox TSO 
      Height          =   360
      Left            =   1440
      TabIndex        =   15
      Top             =   2040
      Width           =   3855
   End
   Begin VB.ComboBox CSCHOOL 
      Height          =   360
      ItemData        =   "FRMTOR.frx":57E2
      Left            =   1440
      List            =   "FRMTOR.frx":57EC
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   1680
      Width           =   3855
   End
   Begin VB.TextBox THS 
      Height          =   360
      Left            =   1440
      TabIndex        =   13
      Top             =   1320
      Width           =   3855
   End
   Begin VB.TextBox Tadd 
      Height          =   360
      Left            =   1440
      TabIndex        =   12
      Top             =   960
      Width           =   3855
   End
   Begin MSComctlLib.ListView LGRADES 
      Height          =   2535
      Left            =   5640
      TabIndex        =   22
      Top             =   2400
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   4471
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "cc"
      SmallIcons      =   "cc"
      ColHdrIcons     =   "cc"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "SUBJECT"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "DESCRIPTIVE TITLE"
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "UNITS"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "FINAL GRADE"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "RE-EXAM"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "REMARKS"
         Object.Width           =   1587
      EndProperty
   End
   Begin MSComctlLib.ImageList cc 
      Left            =   5520
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FRMTOR.frx":5848
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Course:"
      Height          =   240
      Left            =   120
      TabIndex        =   11
      Top             =   4200
      Width           =   675
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Admission Credentials:"
      Height          =   240
      Left            =   120
      TabIndex        =   10
      Top             =   3240
      Width           =   1965
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Date of Graduation:"
      Height          =   240
      Left            =   120
      TabIndex        =   9
      Top             =   2880
      Width           =   1680
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Date of Admission:"
      Height          =   240
      Left            =   2400
      TabIndex        =   8
      Top             =   2520
      Width           =   1620
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "HS Gen. Ave:"
      Height          =   240
      Left            =   120
      TabIndex        =   7
      Top             =   2520
      Width           =   1140
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "S.O. Number:"
      Height          =   240
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   1185
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "College/School:"
      Height          =   240
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   1350
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "High School:"
      Height          =   240
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   1080
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Address:"
      Height          =   240
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   765
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "GRADES"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5640
      TabIndex        =   2
      Top             =   0
      Width           =   4290
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00404040&
      X1              =   120
      X2              =   9850
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   120
      X2              =   9850
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label LBLNAME 
      AutoSize        =   -1  'True
      Caption         =   "NAME"
      Height          =   240
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "BASIC INFORMATION"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5595
   End
End
Attribute VB_Name = "FRMTOR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public TORID As String
Public TORNAME As String
Dim Arr(100) As String
Public RsInfo As ADODB.Recordset   'TBA_SPI
Public RsSY As ADODB.Recordset     'SCHOOLYEARS
Public RsGrades As ADODB.Recordset 'Grades

Sub LoadCourses()
Dim i As Integer
Set RsInfo = Nothing
Set RsInfo = New ADODB.Recordset
For i = 0 To 100
Arr(i) = ""
Next
With RsInfo
    .ActiveConnection = FrmInfoCNTR.ConX
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open "Select * From Courses Order by School, COURSE"
    
    Do Until .EOF
        CCOURSE.AddItem .Fields("Course").Value
        Arr(.AbsolutePosition - 1) = .Fields("DESCRIPTION").Value
        .MoveNext
    Loop
    CCOURSE.ListIndex = .RecordCount - 1
    TDesc.Text = Arr(.RecordCount - 1)
    .Close
End With
Set RsInfo = Nothing
End Sub

Sub LoadInfos()
'Load Student Records
Dim msg As String, Cred As String
Set RsInfo = Nothing
Set RsInfo = New ADODB.Recordset
msg = "Select Student,ADDRESS,HIGHSCHOOL,FORM138"
msg = msg & ",NCEE,DIPLOMA,TRANSCRIPT,HDISMISSAL,PERMIT"
msg = msg & " from TBA_SPI Where IDNO ='" & TORID
msg = msg & "' and STUDENT='" & TORNAME & "'"
With RsInfo
    .ActiveConnection = FrmInfoCNTR.ConX
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open msg
    If .RecordCount = 0 Then
    MsgBox "No Personal Information on " & TORNAME & ". Wizard is unloading.", vbInformation, "No Record"
    Exit Sub
    End If
    'Load if meron
    Tadd.Text = .Fields("ADDRESS").Value
    THS.Text = .Fields("HIGHSCHOOL").Value
    LBLNAME.Caption = .Fields("STUDENT").Value
    'Test Credentials
    Cred = ""
    If Trim(UCase(.Fields("Form138").Value)) = "YES" Then Cred = Cred & "Form 138, "
    If Trim(UCase(.Fields("NCEE").Value)) = "YES" Then Cred = Cred & "NCEE, "
    If Trim(UCase(.Fields("TRANSCRIPT").Value)) = "YES" Then Cred = Cred & "T.O.R., "
    If Trim(UCase(.Fields("HDismissal").Value)) = "YES" Then Cred = Cred & "Honorable Dismissal, "
    If Trim(UCase(.Fields("Permit").Value)) = "YES" Then Cred = Cred & "School Permit, "
    Cred = Left(Cred, Len(Cred) - 2)
    TCred.Text = Cred
    .Close
End With
Set RsInfo = Nothing
End Sub

Sub LoadSYS()
Dim msg As String, i As Long
Set RsSY = Nothing
Set RsSY = New ADODB.Recordset
msg = "Select SCHOOL, SCHOOLYEAR, SEMESTER, COURSE"
msg = msg & " from GRADING_SYS Where IDNO='" & TORID
msg = msg & "' GROUP BY SCHOOL,SCHOOLYEAR, SEMESTER,COURSE"
msg = msg & " ORDER BY SCHOOLYEAR, SEMESTER"
With RsSY
    .ActiveConnection = FrmInfoCNTR.ConX
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open msg
    LSYR.ListItems.Clear
    Do Until .EOF
        i = .AbsolutePosition
        LSYR.ListItems.Add i, , .Fields("SCHOOL").Value, 1, 1
        LSYR.ListItems(i).SubItems(1) = .Fields("SCHOOLYEAR").Value
        LSYR.ListItems(i).SubItems(2) = .Fields("SEMESTER").Value
        LSYR.ListItems(i).SubItems(3) = .Fields("COURSE").Value
        .MoveNext
    Loop
    .Close
End With
Set RsSY = Nothing
End Sub

Sub LoadGrades(SY As String, SEM As String)
Dim msg As String, i As Long
Set RsGrades = Nothing
Set RsGrades = New ADODB.Recordset
msg = "SELECT SUBJECT,SUBJECT_DESCRIPTION,UNITS,FINALS,REEXAM,REMARKS"
msg = msg & " From GRADING_SYS where IDNO='" & TORID & "' and SCHOOLYEAR='"
msg = msg & SY & "' and SEMESTER = '" & SEM & "' order by SUBJECT"
With RsGrades
    .ActiveConnection = FrmInfoCNTR.ConX
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open msg
    LGRADES.ListItems.Clear
    Do Until .EOF
        i = .AbsolutePosition
        LGRADES.ListItems.Add i, , .Fields("Subject").Value, 1, 1
        LGRADES.ListItems(i).SubItems(1) = .Fields("Subject_DESCRIPTION").Value
        LGRADES.ListItems(i).SubItems(2) = .Fields("UNITS").Value
        LGRADES.ListItems(i).SubItems(3) = .Fields("FINALS").Value
        LGRADES.ListItems(i).SubItems(4) = .Fields("REEXAM").Value
        LGRADES.ListItems(i).SubItems(5) = .Fields("REMARKS").Value
        .MoveNext
    Loop
End With
End Sub

Private Sub cbWrite_Click()
On Error GoTo ErrWrite
If CSCHOOL.Text = "" Then
MsgBox "Select School First.", vbInformation, "School Required"
Exit Sub
End If
Me.MousePointer = vbHourglass
Setup_Excel CSCHOOL.Text
Timer1.Enabled = 1
Me.MousePointer = vbDefault
MsgBox "Transcript for " & LBLNAME.Caption & " created ready for printing.", vbInformation, "Transcript Wizzard"
Exit Sub
ErrWrite:
    ErrorTrap Err, "Writing TOR"
    Me.Refresh
    Timer1.Enabled = 1
End Sub

Private Sub CCOURSE_LostFocus()
TDesc.Text = Arr(CCOURSE.ListIndex)
End Sub

Private Sub Form_Load()
Me.LoadCourses
End Sub

Private Sub LSYR_Click()
If LSYR.SelectedItem Is Nothing Then Exit Sub
LoadGrades LSYR.SelectedItem.SubItems(1), LSYR.SelectedItem.SubItems(2)
End Sub

Private Sub Timer1_Timer()
LoadInfos
LoadSYS
Timer1.Enabled = 0
End Sub

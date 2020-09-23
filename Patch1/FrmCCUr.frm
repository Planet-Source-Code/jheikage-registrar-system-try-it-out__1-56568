VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmCCUr 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Courses and Curricula"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   8775
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmCCUr.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   8775
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ix 
      Left            =   4080
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCCUr.frx":57E2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList IMg1 
      Left            =   5880
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   14
      ImageHeight     =   14
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCCUr.frx":AFD4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.ListBox LSubs 
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
      Height          =   1710
      Left            =   3720
      TabIndex        =   10
      Top             =   2880
      Width           =   3135
   End
   Begin MSComctlLib.ImageList Imgl 
      Left            =   4920
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   14
      ImageHeight     =   14
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCCUr.frx":107C6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ISAPTECH.chameleonButton CBADCOURSE 
      Height          =   735
      Left            =   6960
      TabIndex        =   3
      Top             =   120
      Width           =   1695
      _extentx        =   2990
      _extenty        =   1296
      und             =   0
      iname           =   "Tahoma"
      btype           =   14
      tx              =   "Add Course"
      enab            =   -1
      coltype         =   1
      focusr          =   -1
      bcol            =   4210752
      bcolo           =   12632256
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
   Begin MSComctlLib.ListView LCurri 
      Height          =   2415
      Left            =   3720
      TabIndex        =   0
      Top             =   240
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   4260
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "IMg1"
      SmallIcons      =   "IMg1"
      ColHdrIcons     =   "IMg1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Curriculum Info"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Total Subjects"
         Object.Width           =   2540
      EndProperty
   End
   Begin ISAPTECH.chameleonButton CBAdCur 
      Height          =   735
      Left            =   6960
      TabIndex        =   4
      Top             =   960
      Width           =   1695
      _extentx        =   2990
      _extenty        =   1296
      und             =   0
      iname           =   "Tahoma"
      btype           =   14
      tx              =   "Add Curriculum"
      enab            =   -1
      coltype         =   1
      focusr          =   -1
      bcol            =   4210752
      bcolo           =   12632256
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
   Begin ISAPTECH.chameleonButton CBDel 
      Height          =   735
      Left            =   6960
      TabIndex        =   5
      Top             =   1800
      Width           =   1695
      _extentx        =   2990
      _extenty        =   1296
      und             =   0
      iname           =   "Tahoma"
      btype           =   14
      tx              =   "Delete Course"
      enab            =   -1
      coltype         =   1
      focusr          =   -1
      bcol            =   4210752
      bcolo           =   12632256
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
   Begin ISAPTECH.chameleonButton CBPrint 
      Height          =   735
      Left            =   6960
      TabIndex        =   6
      Top             =   2640
      Width           =   1695
      _extentx        =   2990
      _extenty        =   1296
      und             =   0
      iname           =   "Tahoma"
      btype           =   14
      tx              =   "Print Report"
      enab            =   -1
      coltype         =   1
      focusr          =   -1
      bcol            =   4210752
      bcolo           =   12632256
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
   Begin MSComctlLib.ListView LCourse 
      DragIcon        =   "FrmCCUr.frx":15FB8
      Height          =   4695
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   8281
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ix"
      SmallIcons      =   "Imgl"
      ColHdrIcons     =   "Imgl"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Course"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "School"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Description"
         Object.Width           =   5292
      EndProperty
   End
   Begin VB.Label LBLUnits 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Units: 0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3720
      TabIndex        =   11
      Top             =   4560
      Width           =   1275
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SUBJECT LIST"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3720
      TabIndex        =   9
      Top             =   2640
      Width           =   3195
   End
   Begin VB.Label Label3 
      BackColor       =   &H00404040&
      BackStyle       =   0  'Transparent
      Caption         =   "Note: Curricula List format: SY|Yr|Sem"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   6960
      TabIndex        =   8
      Top             =   3480
      Width           =   1740
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CURRICULA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3720
      TabIndex        =   2
      Top             =   0
      Width           =   3180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "COURSES OFFERED"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   3660
   End
   Begin VB.Menu MCourses 
      Caption         =   "Courses"
      Visible         =   0   'False
      Begin VB.Menu CRefresh 
         Caption         =   "Refresh Records"
      End
      Begin VB.Menu CEdit 
         Caption         =   "Edit Record"
      End
      Begin VB.Menu CView 
         Caption         =   "View"
         Begin VB.Menu VIcons 
            Caption         =   "Icons"
         End
         Begin VB.Menu Vlist 
            Caption         =   "List"
         End
         Begin VB.Menu VlistI 
            Caption         =   "List Icons"
         End
         Begin VB.Menu Vdetails 
            Caption         =   "Details"
         End
      End
      Begin VB.Menu brk1 
         Caption         =   "-"
      End
      Begin VB.Menu CCount 
         Caption         =   "Count Records"
      End
   End
End
Attribute VB_Name = "FrmCCUr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public RecCourse As ADODB.Recordset
Public RecCurr As ADODB.Recordset
Public RecSubs As ADODB.Recordset
Sub LoadRecords()
On Error GoTo ErbX
Set RecCourse = Nothing
Set RecCourse = New ADODB.Recordset
With RecCourse
    .ActiveConnection = FrmInfoCNTR.ConX
    .LockType = adLockOptimistic
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .Open "SELECT * FROM COURSES Order by School"
    LCourse.Refresh
    LCourse.ListItems.Clear
    Do Until .EOF
        LCourse.ListItems.Add , , .Fields("Course").Value, 1, 1
        LCourse.ListItems.Item(LCourse.ListItems.Count).SubItems(1) = .Fields("School").Value
        LCourse.ListItems.Item(LCourse.ListItems.Count).SubItems(2) = .Fields("Description").Value
        .MoveNext
    Loop
End With
Exit Sub
ErbX:
    ErrorTrap Err, "Loading Courses"
    Me.Refresh
End Sub

Private Sub CBADCOURSE_Click()
frmCourse.IsAddingCourse = True
frmCourse.Show 1
LoadRecords
End Sub

Private Sub CBAdCur_Click()
With FrmCurricula
.EDCur = True
.LBLCOURSE.Caption = LCourse.SelectedItem.Text
.Show 1
End With
End Sub

Private Sub CBDel_Click()
'Delete All corresponding Curriculum
On Error GoTo ErbX
If LCourse.SelectedItem Is Nothing Then Exit Sub
Dim msg As String, Selected As String
Selected = LCourse.SelectedItem.Text
msg = "This action Will Delete all curriculum of " & Selected
msg = msg & ". Do you want to continue?"
If MsgBox(msg, vbCritical + vbYesNo, "Confirm Delete") = vbNo Then Exit Sub
Set RecCurr = Nothing
Set RecCurr = New ADODB.Recordset
msg = "Delete From Courses where Course='" & Selected & "'"
With RecCurr
    .ActiveConnection = FrmInfoCNTR.ConX
    .CursorType = adOpenDynamic
    .CursorLocation = adUseClient
    .LockType = adLockOptimistic
    .Open msg
End With
Set RecCourse = Nothing
Set RecCourse = New ADODB.Recordset
msg = "Delete From Curriculas Where Course='" & Selected & "'"
With RecCourse
    .ActiveConnection = FrmInfoCNTR.ConX
    .CursorType = adOpenDynamic
    .CursorLocation = adUseClient
    .LockType = adLockOptimistic
    .Open msg
End With
LoadRecords
LoadCurricula LCourse.SelectedItem.Text
Exit Sub
ErbX:
    ErrorTrap Err, "Delete Command"
End Sub

Private Sub CBPrint_Click()
Dim i  As Long
LCourse.Refresh
With FrmReport1.LCourse
    For i = 1 To LCourse.ListItems.Count
        .ListItems.Add , , LCourse.ListItems(i).Text
        .ListItems(i).SubItems(1) = LCourse.ListItems(i).SubItems(1)
        .ListItems(i).SubItems(2) = LCourse.ListItems(i).SubItems(2)
    Next
End With
FrmReport1.Show 1
End Sub

Private Sub CCount_Click()
Dim MCNP As Long, ISAP As Long, i As Long
Dim ISp As String
ISp = "INTERNATIONAL SCHOOL OF ASIA AND THE PACIFIC"
For MCNP = 1 To LCourse.ListItems.Count
    If UCase(LCourse.ListItems(MCNP).SubItems(1)) = ISp Then
    ISAP = ISAP + 1
    End If
Next
MCNP = LCourse.ListItems.Count - ISAP
MsgBox "Total Courses Offered in ISAP: " & ISAP _
    & vbNewLine & "Total Courses Offered in MCNP: " & _
        MCNP, vbInformation, "Courses"
End Sub

Private Sub CEdit_Click()
LCourse_DblClick
End Sub

Private Sub CRefresh_Click()
LoadRecords
End Sub

Private Sub Form_Load()
LoadRecords
End Sub
Sub LoadCurricula(Course As String)
On Error GoTo ErbX
Dim msg As String, Itm As String
Set RecCurr = Nothing
Set RecCurr = New ADODB.Recordset
msg = "Select Course, Schoolyear, Major,YearLevel,Semester,Count(SubjectCode) as Codes from Curriculas"
msg = msg & " Where Course = '" & Course & "' Group by"
msg = msg & " COURSE,SCHOOLYEAR, MAJOR,YearLevel,Semester order by Schoolyear, Yearlevel"
With RecCurr
    .ActiveConnection = FrmInfoCNTR.ConX
    .CursorType = adOpenDynamic
    .CursorLocation = adUseClient
    .LockType = adLockOptimistic
    .Open msg
    LCurri.ListItems.Clear
    Do Until .EOF
        If IsNull(.Fields("major").Value) Or _
            Trim(.Fields("Major").Value) = "" Then
            Itm = .Fields("Schoolyear").Value
        Else
            Itm = .Fields("Major").Value & "|" & .Fields("Schoolyear").Value
        End If
        Itm = Itm & "|" & .Fields("YearLevel").Value
        Itm = Itm & "|" & .Fields("Semester").Value
        LCurri.ListItems.Add , , Itm, 1, 1
        LCurri.ListItems.Item(LCurri.ListItems.Count).SubItems(1) = .Fields("CODES").Value & " subject/s"
        .MoveNext
    Loop
End With
Exit Sub
ErbX:
    ErrorTrap Err, "Loading Curricula"
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
FrmSet.WCrs.Visible = False
End Sub

Private Sub LCourse_Click()
If LCourse.SelectedItem Is Nothing Then Exit Sub
LCurri.ToolTipText = ""
LCourse.ToolTipText = ""
LSubs.Clear
LoadCurricula LCourse.SelectedItem.Text
LCourse.ToolTipText = LCourse.SelectedItem.SubItems(1)
End Sub

Private Sub LCourse_DblClick()
If LCourse.SelectedItem Is Nothing Then Exit Sub
frmCourse.IsAddingCourse = False
frmCourse.Show 1
LoadRecords
End Sub

Private Sub LCourse_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 2 And Shift = 0 Then
    Me.PopupMenu MCourses
    If LCourse.SelectedItem Is Nothing Then
        CEdit.Enabled = False
    Else
        CEdit.Enabled = True
    End If
End If
End Sub

Private Sub LCurri_Click()
'If LCurri.ListItems.Count = 0 Then
If LCurri.SelectedItem Is Nothing Then
LSubs.Clear
Exit Sub
End If
LCurri.ToolTipText = LCurri.SelectedItem.SubItems(1)
MethodX LCurri.SelectedItem.Text
End Sub

Private Sub LCurri_DblClick()
With FrmCurricula
.LBLCOURSE.Caption = LCourse.SelectedItem.Text
.AddNew False
.MethodSplit LCurri.SelectedItem.Text
.Show 1
End With
End Sub

Sub MethodX(strx As String)
Dim VRS, msg As String
Dim i1 As String, i2 As String, i3 As String, i4 As String
VRS = Split(strx, "|")
i1 = VRS(0)
i2 = VRS(1)
If UBound(VRS) > 2 Then
i4 = VRS(2)
i3 = VRS(3)
Else
i4 = ""
i3 = VRS(2)
End If
        msg = "Select * from Curriculas Where Course"
        msg = msg & "='" & LCourse.SelectedItem.Text & "' and"
        msg = msg & " YearLevel='" & i2 & "' and"
        msg = msg & " Schoolyear='" & i1 & "' and"
        msg = msg & " Semester='" & i3 & "'"
        If IsNull(i4) Then GoTo Loader
        If Len(Trim(i4)) <> 0 Then
        msg = msg & " and MAJOR='" & i4 & "'"
        End If
Loader:
LoadSubjects msg

End Sub

Sub LoadSubjects(msg As String)
Dim TotUn As Double
On Error GoTo ErbX
Set RecSubs = Nothing
Set RecSubs = New ADODB.Recordset
With RecSubs
    .ActiveConnection = FrmInfoCNTR.ConX
    .LockType = adLockOptimistic
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .Open msg
    LSubs.Clear
    Do Until .EOF
        LSubs.AddItem .Fields("Subjectcode").Value & " [Unit/s: " & .Fields("Units").Value & "]"
        TotUn = TotUn + Val(.Fields("Units").Value)
        .MoveNext
    Loop
End With
LBLUnits = "Total Units: " & TotUn
Exit Sub
ErbX:
    ErrorTrap Err, "Loading Subjects"
    Me.Refresh
End Sub

Private Sub Vdetails_Click()
LCourse.Arrange = lvwNone
LCourse.FlatScrollBar = True
LCourse.FullRowSelect = True
LCourse.View = lvwReport
End Sub

Private Sub VIcons_Click()
LCourse.Arrange = lvwAutoTop
LCourse.View = lvwIcon
End Sub

Private Sub Vlist_Click()
LCourse.Arrange = lvwAutoTop
LCourse.View = lvwList
End Sub

Private Sub VlistI_Click()
LCourse.Arrange = lvwAutoTop
LCourse.View = lvwSmallIcon
End Sub

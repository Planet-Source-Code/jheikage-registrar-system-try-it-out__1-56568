VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmReport1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Course/Curriculum Reports"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5715
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmReport1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   5715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   7223
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Courses"
      TabPicture(0)   =   "FrmReport1.frx":57E2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Line1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Line2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Line3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label3"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "o2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "CBOK"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "o1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Theader"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Curriculum"
      TabPicture(1)   =   "FrmReport1.frx":57FE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label6"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Line4"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label4"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Lb1"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label5"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "LCourse"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "CSY"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "CMJR"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "CBOK2"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).ControlCount=   9
      Begin VB.CommandButton CBOK2 
         Caption         =   "OK"
         Height          =   495
         Left            =   -70800
         TabIndex        =   15
         Top             =   3480
         Width           =   1215
      End
      Begin VB.ComboBox CMJR 
         Height          =   360
         Left            =   -71640
         TabIndex        =   13
         Text            =   "CMJR"
         Top             =   1680
         Width           =   1095
      End
      Begin VB.ComboBox CSY 
         Height          =   360
         Left            =   -71640
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox Theader 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   2640
         Width           =   5175
      End
      Begin VB.OptionButton o1 
         Appearance      =   0  'Flat
         Caption         =   "MCNP Courses"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   480
         TabIndex        =   3
         Top             =   1560
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.CommandButton CBOK 
         Caption         =   "OK"
         Height          =   495
         Left            =   4200
         TabIndex        =   2
         Top             =   3360
         Width           =   1215
      End
      Begin VB.OptionButton o2 
         Appearance      =   0  'Flat
         Caption         =   "ISAP Courses"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   480
         TabIndex        =   4
         Top             =   1800
         Width           =   1575
      End
      Begin MSComctlLib.ListView LCourse 
         Height          =   2415
         Left            =   -74880
         TabIndex        =   8
         Top             =   1560
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   4260
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FlatScrollBar   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Course List"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "School"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Description"
            Object.Width           =   5292
         EndProperty
      End
      Begin VB.Label Label5 
         Caption         =   "Major(Optional):"
         Height          =   255
         Left            =   -73080
         TabIndex        =   11
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Lb1 
         Caption         =   "Schoolyear:"
         Height          =   255
         Left            =   -73080
         TabIndex        =   10
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "List Of Courses"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   -74880
         TabIndex        =   9
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Enter Header Caption:"
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
         Left            =   240
         TabIndex        =   6
         Top             =   2280
         Width           =   2160
      End
      Begin VB.Line Line3 
         X1              =   240
         X2              =   5400
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Line Line2 
         X1              =   240
         X2              =   5400
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Select School Below:"
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
         Left            =   240
         TabIndex        =   5
         Top             =   1200
         Width           =   1995
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   5400
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label1 
         Caption         =   "Create A List of Courses Reports for ISAP or MCNP. Select the School you want to produce a report and Click OK."
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   5295
      End
      Begin VB.Line Line4 
         X1              =   -74880
         X2              =   -69600
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label Label6 
         Caption         =   $"FrmReport1.frx":581A
         Height          =   735
         Left            =   -74880
         TabIndex        =   14
         Top             =   480
         Width           =   5295
      End
   End
End
Attribute VB_Name = "FrmReport1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RecSel As ADODB.Recordset
Dim RecSel1 As ADODB.Recordset
Sub SetRecSel(msg As String)
Set RecSel = Nothing
Set RecSel = New ADODB.Recordset
With RecSel
    .ActiveConnection = FrmInfoCNTR.ConX
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open msg
End With
End Sub

Private Sub cbOK_Click()
Dim MXP As String, msg As String
msg = "Select * From Courses where School = '"
If o1.Value = False Then
MXP = "International School of Asia and the Pacific"
msg = msg & "International School of Asia and the pacific' order by Course"
Else
MXP = "Medical Colleges of Northern Philippines"
msg = msg & "Medical Colleges of Northern Philippines' order by Course"
End If
SetDenver msg, 1
With DR1
    .Sections("PageHeader").Controls("LblSchool").Caption = UCase(Trim(MXP))
    .Sections("PageHeader").Controls("Lblheader").Caption = Trim(UCase(Theader.Text))
    .Refresh
    .Show 1
End With
End Sub

Private Sub CBOK2_Click()
If CSY.ListIndex < 0 Then MsgBox "Select Schoolyear first.", vbInformation, "School Year": Exit Sub
If LCourse.SelectedItem Is Nothing Then MsgBox "Select Course First.", vbInformation, "Course": Exit Sub
Dim msg As String, MR As String, SY As String, CRS As String
Dim SCHL As String, DESC As String
If Trim(CMJR.Text) = "" Then CMJR.Text = "": MR = "" Else MR = "Major in " & Trim(CMJR.Text)
SY = CSY.Text: CRS = LCourse.SelectedItem.Text
DESC = LCourse.SelectedItem.SubItems(2)
SCHL = LCourse.SelectedItem.SubItems(1)
msg = "SHAPE {SELECT COURSE, MAJOR, SCHOOLYEAR, YEARLEVEL FROM CURRICULAs WHERE COURSE = '" & CRS & "' and Schoolyear='" & SY & "' GROUP BY COURSE, MAJOR, SCHOOLYEAR, YEARLEVEL ORDER BY YEARLEVEL}" & _
"AS cmdCurHead APPEND (( SHAPE {SELECT COURSE, MAJOR, SCHOOLYEAR, YEARLEVEL, SEMESTER, Sum(UNITS) as Total_Units FROM CURRICULAs WHERE COURSE = '" & CRS & "' and Schoolyear='" & SY & "' GROUP BY COURSE, " & _
"MAJOR, SCHOOLYEAR, YEARLEVEL, SEMESTER ORDER BY YEARLEVEL, SEMESTER}  AS SecHead APPEND ({SELECT * FROM CURRICULAS WHERE COURSE = '" & CRS & "' AND Schoolyear='" & SY & "' ORDER BY COURSE}  AS DetailCur " & _
"RELATE 'COURSE' TO 'COURSE','MAJOR' TO 'MAJOR','YEARLEVEL' TO 'YEARLEVEL','SCHOOLYEAR' TO 'SCHOOLYEAR','SEMESTER' TO 'SEMESTER') AS DetailCur) AS SecHead RELATE 'YEARLEVEL' TO 'YEARLEVEL','SCHOOLYEAR' TO 'SCHOOLYEAR') AS SecHead"
SetDenver msg, 2
With DR2
.Sections("pageHeader").Controls("LblSy").Caption = UCase("School year " & CSY.Text)
.Sections("pageHeader").Controls("LblSchool").Caption = UCase(LCourse.SelectedItem.SubItems(1))
.Sections("PageHeader").Controls("lblDesc").Caption = UCase(LCourse.SelectedItem.SubItems(2))
.Sections("pageHeader").Controls("LblMajor").Caption = UCase(MR)

.Refresh
DR2.Show 1
End With

End Sub
Sub LoadCurItem(Course As String)
Dim msg As String, Itm As String
Set RecSel1 = Nothing
Set RecSel1 = New ADODB.Recordset
msg = "Select Course, Schoolyear, Major from Curriculas"
msg = msg & " Where Course = '" & Course & "' Group by"
msg = msg & " COURSE,SCHOOLYEAR, MAJOR order by Schoolyear"
With RecSel1
    .ActiveConnection = FrmInfoCNTR.ConX
    .CursorType = adOpenDynamic
    .CursorLocation = adUseClient
    .LockType = adLockOptimistic
    .Open msg
    CSY.Clear
    CMJR.Clear
    Do Until .EOF
        If IsNull(.Fields("major").Value) = False Or _
            Trim(.Fields("Major").Value) <> "" Then
            CMJR.AddItem .Fields("Major").Value
        End If
        
        Itm = .Fields("Schoolyear").Value
        CSY.AddItem Itm
        .MoveNext
    Loop
End With
Exit Sub
ErbX:
    ErrorTrap Err, "Loading Curricula"
    Me.Refresh
End Sub


Private Sub LCourse_Click()
LoadCurItem LCourse.SelectedItem.Text
End Sub

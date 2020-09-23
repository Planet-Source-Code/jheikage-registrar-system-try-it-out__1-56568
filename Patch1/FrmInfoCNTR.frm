VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmInfoCNTR 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Student Information Center"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8685
   Icon            =   "FrmInfoCNTR.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   8685
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImgLst 
      Left            =   3360
      Top             =   4920
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
            Picture         =   "FrmInfoCNTR.frx":57E2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame FRM3 
      Appearance      =   0  'Flat
      Caption         =   "LOAD DATA"
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
      Height          =   1095
      Left            =   3960
      TabIndex        =   10
      Top             =   0
      Width           =   2895
      Begin ISAPTECH.chameleonButton CBNAV 
         Height          =   375
         Left            =   2040
         TabIndex        =   15
         Top             =   600
         Width           =   735
         _extentx        =   1296
         _extenty        =   661
         und             =   0
         iname           =   "Tahoma"
         btype           =   5
         tx              =   "GO"
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
         ItemData        =   "FrmInfoCNTR.frx":AFD4
         Left            =   1320
         List            =   "FrmInfoCNTR.frx":AFE1
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   600
         Width           =   735
      End
      Begin VB.ComboBox CSY 
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
         Left            =   1320
         TabIndex        =   11
         Text            =   "SCHOOL YEAR"
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "SEMESTER:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   825
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "SCHOOL YEAR:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1110
      End
   End
   Begin VB.Frame FRM1 
      Appearance      =   0  'Flat
      Caption         =   "BASIC INFORMATION"
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
      Height          =   1815
      Left            =   3960
      TabIndex        =   2
      Top             =   1080
      Width           =   2895
      Begin VB.Image IMGGRD 
         Height          =   855
         Left            =   1680
         MouseIcon       =   "FrmInfoCNTR.frx":AFF4
         MousePointer    =   99  'Custom
         Picture         =   "FrmInfoCNTR.frx":B146
         Stretch         =   -1  'True
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "COLLEGE RECORD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1560
         TabIndex        =   4
         Top             =   1200
         Width           =   1200
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "PERSONAL INFORMATION"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   1200
      End
      Begin VB.Image IMGPI 
         Height          =   855
         Left            =   240
         MouseIcon       =   "FrmInfoCNTR.frx":BE86
         MousePointer    =   99  'Custom
         Picture         =   "FrmInfoCNTR.frx":BFD8
         Stretch         =   -1  'True
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame FRM2 
      Appearance      =   0  'Flat
      Caption         =   "STUDENT INFO"
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
      Height          =   1455
      Left            =   3960
      TabIndex        =   5
      Top             =   2880
      Width           =   2895
      Begin VB.Label LBLCOURSE 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   9
         Top             =   1080
         Width           =   45
      End
      Begin VB.Label LBLNAME 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   45
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "COURSE/YR:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   930
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "NAME:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   480
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "DATABASE"
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
      Height          =   1335
      Left            =   3960
      TabIndex        =   17
      Top             =   4320
      Width           =   2895
      Begin VB.Label LBLSUBCON 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1800
         TabIndex        =   25
         Top             =   960
         Width           =   45
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "STUDENT RECORD/S:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   720
         Width           =   1560
      End
      Begin VB.Label LBLTOTAL 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1800
         TabIndex        =   19
         Top             =   480
         Width           =   45
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "TOTAL REQUESTED RECORD/S:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame FRM4 
      Appearance      =   0  'Flat
      Caption         =   "COMMANDS"
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
      Height          =   5655
      Left            =   6960
      TabIndex        =   16
      Top             =   0
      Width           =   1695
      Begin ISAPTECH.chameleonButton CBNEW 
         Height          =   615
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   1455
         _extentx        =   2566
         _extenty        =   1085
         und             =   0
         iname           =   "Tahoma"
         btype           =   14
         tx              =   "NEW STUDENT"
         enab            =   -1
         coltype         =   1
         focusr          =   -1
         bcol            =   8421504
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
      Begin ISAPTECH.chameleonButton CBSEARCH 
         Height          =   615
         Left            =   120
         TabIndex        =   21
         Top             =   1080
         Width           =   1455
         _extentx        =   2566
         _extenty        =   1085
         und             =   0
         iname           =   "Tahoma"
         btype           =   14
         tx              =   "SEARCH"
         enab            =   -1
         coltype         =   1
         focusr          =   -1
         bcol            =   8421504
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
      Begin ISAPTECH.chameleonButton CBDEL 
         Height          =   615
         Left            =   120
         TabIndex        =   22
         Top             =   1800
         Width           =   1455
         _extentx        =   2566
         _extenty        =   1085
         und             =   0
         iname           =   "Tahoma"
         btype           =   14
         tx              =   "DELETE STUDENT"
         enab            =   -1
         coltype         =   1
         focusr          =   -1
         bcol            =   8421504
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
      Begin ISAPTECH.chameleonButton CBREPORTS 
         Height          =   615
         Left            =   120
         TabIndex        =   23
         Top             =   2520
         Width           =   1455
         _extentx        =   2566
         _extenty        =   1085
         und             =   0
         iname           =   "Tahoma"
         btype           =   14
         tx              =   "REPORTS"
         enab            =   -1
         coltype         =   1
         focusr          =   -1
         bcol            =   8421504
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
      Begin ISAPTECH.chameleonButton CBList 
         Height          =   615
         Left            =   120
         TabIndex        =   26
         Top             =   3240
         Width           =   1455
         _extentx        =   2566
         _extenty        =   1085
         und             =   0
         iname           =   "Tahoma"
         btype           =   14
         tx              =   "SELECT STUDENT/S"
         enab            =   -1
         coltype         =   1
         focusr          =   -1
         bcol            =   8421504
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
      Begin VB.Frame FRM6 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1695
         Left            =   120
         TabIndex        =   27
         Top             =   3840
         Width           =   1455
         Begin VB.Label LBLNUML 
            AutoSize        =   -1  'True
            Caption         =   "000000000000"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   29
            Top             =   1080
            Width           =   1260
         End
         Begin VB.Label Label10 
            Caption         =   "NUMBER OF STUDENT/S LOADED BY THE SYSTEM:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   795
            Left            =   120
            TabIndex        =   28
            Top             =   240
            Width           =   1155
         End
      End
   End
   Begin MSComctlLib.ListView LDView 
      Height          =   5415
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   9551
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ImgLst"
      SmallIcons      =   "ImgLst"
      ColHdrIcons     =   "ImgLst"
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
         Text            =   "IDNO"
         Object.Width           =   2540
         ImageIndex      =   1
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "NAME"
         Object.Width           =   4410
         ImageIndex      =   1
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Sex"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ID/STUDENT LIST"
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
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3885
   End
   Begin VB.Menu Menu1 
      Caption         =   "Menu1"
      Visible         =   0   'False
      Begin VB.Menu MenuPrint 
         Caption         =   "&Print P.I."
      End
      Begin VB.Menu MenuView 
         Caption         =   "&View P.I."
      End
      Begin VB.Menu BR 
         Caption         =   "-"
      End
      Begin VB.Menu MenuTOR 
         Caption         =   "&View/Create TOR"
      End
   End
End
Attribute VB_Name = "FrmInfoCNTR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public WithEvents ConX As ADODB.Connection
Attribute ConX.VB_VarHelpID = -1
Public WithEvents ConRec As ADODB.Recordset
Attribute ConRec.VB_VarHelpID = -1
Dim NotDbl As Boolean   'You Must First CLick Go to continue
Dim SQLTOPRINT As String
Private Sub CBDel_Click()
Dim msg As String
msg = "This Action Will Delete all the records of " & LDVIEW.SelectedItem.SubItems(1)
msg = msg & " on your database. Do you want to continue?"
If MsgBox(msg, vbCritical + vbYesNo, "Delete Student Record") = vbYes Then
    'delete now
    DeleteSPI
    DeleteRecords
    MsgBox "Record permanently deleted from the Database.", vbInformation, "Delete Complete"
    LoadRecordsToList "Select * From TBA_SPI order by IDNO"
End If
End Sub

Sub DeleteSPI()
Set ConRec = New ADODB.Recordset
With ConRec
    .ActiveConnection = ConX
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open "Delete From TBA_SPI where IDNO='" & LDVIEW.SelectedItem.Text & "'"

End With
Set ConRec = Nothing
End Sub

Sub DeleteRecords()
Set ConRec = New ADODB.Recordset
With ConRec
    .ActiveConnection = ConX
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open "Delete From GRADING_SYS where IDNO='" & LDVIEW.SelectedItem.Text & "'"

End With
Set ConRec = Nothing
End Sub
Private Sub CBList_Click()
'Using Sqls
FrmNamI.Show 1
Dim SQL As String
If Trim(Srch) = "" Then Exit Sub
Select Case SRCHTYP
Case True
'Use by Name
'Tedious Work
Dim i As Long, SL
SQL = "Select * From TBA_SPI where STUDENT like '"
SL = Split(Srch, " ", , vbTextCompare)
For i = LBound(SL) To UBound(SL)
    SQL = SQL & SL(i) & "%"
Next
SQL = SQL & "' order by IDNO, Student"
Case False
'Use by ID
SQL = "Select * From Grading_Sys where IDNO like '" & Srch & "%' order by IDNO"
End Select
SQLTOPRINT = SQL
LoadRecordsToList SQL
End Sub

Private Sub CBNAV_Click()
If Trim(CSY.Text) = "" Or Trim(CSEM.Text) = "" Then Exit Sub
NotDbl = True
SqlStGrd = "Select * From Grading_Sys Where IDNO ='"
SqlStGrd = SqlStGrd & LDVIEW.SelectedItem.Text & "'"
SqlStGrd = SqlStGrd & " and Schoolyear = '" & CSY.Text
SqlStGrd = SqlStGrd & "' and Semester = '" & CSEM.Text & "' order by Subject"
Set ConRec = New ADODB.Recordset
With ConRec
    .ActiveConnection = ConX
    .LockType = adLockOptimistic
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .Open SqlStGrd
        If .RecordCount = 0 Then
            MsgBox "There are no Grading Records for this Student", vbInformation, "No Records"
            LBLNAME.Caption = LDVIEW.SelectedItem.SubItems(1)
            Me.LBLSUBCON.Caption = "0"
            .Close
            Exit Sub
        End If
        'Load years
        LBLSUBCON.Caption = .RecordCount
    .Close
End With
Set ConRec = Nothing
End Sub

Private Sub CBNEW_Click()
'Add new Personal Information for a new student
FrmSPI.AddOr = True
FrmSPI.Show 1
LoadRecordsToList "Select * From TBA_SPI order by IDNO"
End Sub

Private Sub CBREPORTS_Click()
'Loaddenver
Dim msg As String
msg = "Application is Ready to create reports." & vbNewLine & vbNewLine
msg = msg & "To print the selected file only click Ok, if you wish to print all "
msg = msg & "records loaded in the system, click cancel."
If MsgBox(msg, vbInformation + vbOKCancel, "Print Reports") = vbOK Then
'Single document
If LDVIEW.SelectedItem Is Nothing Then MsgBox "Select a record first.", vbInformation, "ERROR": Exit Sub
msg = "Select * From TBA_SPI where IDNO='" & LDVIEW.SelectedItem.Text & "'"
Else
msg = SQLTOPRINT
End If
SetDenver msg, 3
DR3.Refresh
DR3.Show 1
End Sub

Private Sub CBSearch_Click()
FrmNamI.Show 1

Dim x As ListItem, i As Long
With LDVIEW
If SRCHTYP = True Then  'Name
    For i = 0 To .ListItems.Count - 1
        If InStr(1, .ListItems.Item(i + 1).SubItems(1), Srch, vbTextCompare) <> 0 Then
        .ListItems(i + 1).EnsureVisible
        .ListItems(i + 1).Selected = True
        .SetFocus
        Exit Sub
        End If
    Next
Else
Set x = .FindItem(Srch, lvwText, , lvwPartial)
End If
If x Is Nothing Then
    MsgBox "Record No Match.", vbInformation, "No Match"
    Exit Sub
Else
SearchSkip:
    x.EnsureVisible
    x.Selected = True
End If
End With
LDVIEW.SetFocus
End Sub

Private Sub Form_Load()
Dim SQL As String
SQL = "SELECT * FROM TBA_SPI"
LoadRecordsToList SQL
SQLTOPRINT = SQL
End Sub

Private Sub Form_Unload(Cancel As Integer)
FrmSet.WSIC.Visible = False
End Sub

Private Sub IMGGRD_Click()
If NotDbl = False Then Exit Sub

FrmGRD.LBLID.Caption = Me.LDVIEW.SelectedItem.Text
FrmGRD.LBLNAME.Caption = Me.LDVIEW.SelectedItem.SubItems(1)
FrmGRD.LblSex.Caption = Me.LDVIEW.SelectedItem.SubItems(2)
FrmGRD.Show

FrmSet.CBINFOR.Enabled = False
Me.Hide
End Sub

Private Sub IMGPI_Click()
FrmSPI.AddOr = False
'Load for Editing
FrmSPI.Show 1
End Sub


Private Sub LDView_Click()
NotDbl = False
Me.LBLTOTAL.Caption = "0"
Me.LBLSUBCON.Caption = "0"
End Sub

Private Sub LDView_DblClick()
'Get Item
Dim msg As String
msg = "Select IDNO, STUDENT, SCHOOLYEAR,SEMESTER,COURSE,YEARLEVEL FROM grading_Sys "
msg = msg & " WHERE IDNO = '" & LDVIEW.SelectedItem.Text
msg = msg & "' GROUP By IDNO,STUDENT, SCHOOLYEAR, SEMESTER,COURSE, YEARLEVEL "
msg = msg & "Order by SCHOOLYEAR"
NotDbl = False
Set ConRec = New ADODB.Recordset
With ConRec
    .ActiveConnection = ConX
    .LockType = adLockOptimistic
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .Open msg
        If .RecordCount = 0 Then
            MsgBox "There are no Grading Records for this Student", vbInformation, "No Records"
            CSY.Clear
            LBLNAME.Caption = LDVIEW.SelectedItem.SubItems(1)
            Me.LBLTOTAL.Caption = "0"
            .Close
            Exit Sub
        End If
        'Load years
        loadSY
    .Close
Set ConRec = Nothing
End With
End Sub

Sub loadSY()
With ConRec
Dim i As Long
    .Properties.Refresh
    CSY.Clear
    .MoveFirst
    Do Until .EOF
        For i = 0 To CSY.ListCount - 1
        CSY.ListIndex = i
        If .Fields("Schoolyear").Value = CSY.Text Then
        GoTo NexRec
        End If
        Next
        CSY.AddItem .Fields("Schoolyear").Value
NexRec:
        .MoveNext
    Loop
    Me.LBLTOTAL.Caption = .RecordCount
    .MoveLast
    Me.LBLNAME.Caption = .Fields(1).Value
    Me.LBLCOURSE.Caption = Trim(.Fields("Course").Value) & " " & .Fields("YearLevel").Value
    CSY.ListIndex = CSY.ListCount - 1
End With

End Sub

Private Sub LDView_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 93 Then LDView_MouseDown 2, 0, 10, 10
End Sub

Private Sub LDView_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 2 Then
If LDVIEW.SelectedItem Is Nothing Then Exit Sub
 Me.PopupMenu Menu1
End If
End Sub

Private Sub MenuPrint_Click()
'Print Selected
SetDenver "Select * From TBA_SPI where IDNO='" & LDVIEW.SelectedItem.Text & "'", 3
DR3.Refresh
DR3.Show 1
End Sub

Private Sub MenuTOR_Click()
'Create TOR Here
With FRMTOR
    .TORID = LDVIEW.SelectedItem.Text
    .TORNAME = LDVIEW.SelectedItem.SubItems(1)
    .Show 1
End With
End Sub

Private Sub MenuView_Click()
IMGPI_Click
End Sub

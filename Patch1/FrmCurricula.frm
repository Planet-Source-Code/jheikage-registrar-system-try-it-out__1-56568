VERSION 5.00
Begin VB.Form FrmCurricula 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Frm Course"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   7080
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmCurricula.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   7080
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox LblMjr 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   1080
      TabIndex        =   1
      Top             =   360
      Width           =   3615
   End
   Begin ISAPTECH.chameleonButton CBL 
      Height          =   495
      Left            =   6600
      TabIndex        =   26
      Top             =   3120
      Width           =   375
      _extentx        =   661
      _extenty        =   873
      und             =   0
      iname           =   "Tahoma"
      btype           =   4
      tx              =   ">>"
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
   Begin ISAPTECH.chameleonButton CBN 
      Height          =   495
      Left            =   6240
      TabIndex        =   25
      Top             =   3120
      Width           =   375
      _extentx        =   661
      _extenty        =   873
      und             =   0
      iname           =   "Tahoma"
      btype           =   4
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
   Begin ISAPTECH.chameleonButton CBP 
      Height          =   495
      Left            =   5160
      TabIndex        =   24
      Top             =   3120
      Width           =   375
      _extentx        =   661
      _extenty        =   873
      und             =   0
      iname           =   "Tahoma"
      btype           =   4
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
   Begin ISAPTECH.chameleonButton CBB 
      Height          =   495
      Left            =   4800
      TabIndex        =   23
      Top             =   3120
      Width           =   375
      _extentx        =   661
      _extenty        =   873
      und             =   0
      iname           =   "Tahoma"
      btype           =   4
      tx              =   "<<"
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
   Begin ISAPTECH.chameleonButton CBSearch 
      Height          =   495
      Left            =   4800
      TabIndex        =   11
      Top             =   1920
      Width           =   2175
      _extentx        =   3836
      _extenty        =   873
      und             =   0
      iname           =   "Tahoma"
      btype           =   14
      tx              =   "Search"
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
   Begin ISAPTECH.chameleonButton CBDel 
      Height          =   495
      Left            =   4800
      TabIndex        =   10
      Top             =   1320
      Width           =   2175
      _extentx        =   3836
      _extenty        =   873
      und             =   0
      iname           =   "Tahoma"
      btype           =   14
      tx              =   "Delete"
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
   Begin ISAPTECH.chameleonButton CBUpdate 
      Height          =   495
      Left            =   4800
      TabIndex        =   9
      Top             =   720
      Width           =   2175
      _extentx        =   3836
      _extenty        =   873
      und             =   0
      iname           =   "Tahoma"
      btype           =   14
      tx              =   "Update"
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
   Begin ISAPTECH.chameleonButton CBNEW 
      Height          =   495
      Left            =   4800
      TabIndex        =   8
      Top             =   120
      Width           =   2175
      _extentx        =   3836
      _extenty        =   873
      und             =   0
      iname           =   "Tahoma"
      btype           =   14
      tx              =   "Add New"
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
   Begin VB.Frame FRM1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   120
      TabIndex        =   13
      Top             =   600
      Width           =   4575
      Begin VB.TextBox TSchlyr 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   3120
         TabIndex        =   3
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox CSEM 
         Appearance      =   0  'Flat
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
         ItemData        =   "FrmCurricula.frx":57E2
         Left            =   1080
         List            =   "FrmCurricula.frx":57EF
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   600
         Width           =   735
      End
      Begin VB.ComboBox CYR 
         Appearance      =   0  'Flat
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
         ItemData        =   "FrmCurricula.frx":5802
         Left            =   1080
         List            =   "FrmCurricula.frx":5815
         TabIndex        =   2
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Year Level:"
         Height          =   240
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   960
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Semester:"
         Height          =   240
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "School Year:"
         Height          =   240
         Left            =   1920
         TabIndex        =   14
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame FRM2 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   120
      TabIndex        =   18
      Top             =   1680
      Width           =   4575
      Begin VB.TextBox tunits 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   360
         Left            =   3720
         TabIndex        =   6
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox tdesc 
         Appearance      =   0  'Flat
         Height          =   840
         Left            =   1200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   720
         Width           =   3255
      End
      Begin VB.TextBox Tsubject 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1200
         TabIndex        =   5
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Units:"
         Height          =   240
         Left            =   3120
         TabIndex        =   21
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Description:"
         Height          =   240
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Width           =   1020
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Subject:"
         Height          =   240
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.Label LBLTOT 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0 of 0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   5520
      TabIndex        =   22
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "MAJOR:"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   360
      Width           =   855
   End
   Begin VB.Label LBLCOURSE 
      AutoSize        =   -1  'True
      Caption         =   "COURSE"
      Height          =   240
      Left            =   1080
      TabIndex        =   12
      Top             =   120
      Width           =   720
   End
   Begin VB.Label Label1 
      Caption         =   "COURSE:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "FrmCurricula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public RecCur As ADODB.Recordset
Public RecAdd As ADODB.Recordset
Dim Addix As Boolean
Public EDCur As Boolean

Private Sub CBB_Click()
With RecCur
    .MoveFirst
    Populate
    LBLTOT.Caption = .AbsolutePosition & " of " & .RecordCount
End With
End Sub

Private Sub CBDel_Click()
On Error GoTo ErbX
Dim msg As String
Select Case CBDel.Caption
    Case "Cancel"
    EnableAll
    Addix = False
    Case "Delete"
    If MsgBox("Do you want to delete this Subject?", vbYesNo + vbQuestion, "Delete") = vbNo Then Exit Sub
        msg = "Delete from Curriculas Where Course"
        msg = msg & "='" & LBLCOURSE.Caption & "' and"
        msg = msg & " YearLevel='" & CYR.Text & "' and"
        msg = msg & " Schoolyear='" & TSCHLYR.Text & "' and"
        msg = msg & " Semester='" & CSEM.Text & "' and"
        msg = msg & " SubjectCode='" & RecCur.Fields("Subjectcode").Value & "'"
        If IsNull(LblMjr.Text) Then GoTo Loader
        If Len(Trim(LblMjr.Text)) <> 0 Then
        msg = msg & " and MAJOR='" & LblMjr.Text & "'"
        End If
    
    Set RecAdd = Nothing
    Set RecAdd = New ADODB.Recordset
    With RecAdd
        .ActiveConnection = FrmInfoCNTR.ConX
        .CursorLocation = adUseClient
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .Open msg
    End With
    Set RecAdd = Nothing
End Select
        msg = "Select * from Curriculas Where Course"
        msg = msg & "='" & LBLCOURSE.Caption & "' and"
        msg = msg & " YearLevel='" & CYR.Text & "' and"
        msg = msg & " Schoolyear='" & TSCHLYR.Text & "' and"
        msg = msg & " Semester='" & CSEM.Text & "'"
        If IsNull(LblMjr.Text) Then GoTo Loader
        If Len(Trim(LblMjr.Text)) <> 0 Then
        msg = msg & " and MAJOR='" & LblMjr.Text & "'"
        End If
Loader:
LoadRecords msg
Exit Sub
ErbX:
    ErrorTrap Err, "Delete Record Command"
    Me.Refresh
End Sub

Private Sub CBL_Click()
With RecCur
    .MoveLast
    Populate
    LBLTOT.Caption = .AbsolutePosition & " of " & .RecordCount
End With

End Sub

Private Sub CBN_Click()
With RecCur
    .MoveNext
    If .EOF Then .MoveLast
    Populate
    LBLTOT.Caption = .AbsolutePosition & " of " & .RecordCount
End With

End Sub

Private Sub CBNEW_Click()
Addix = True
LBLTOT.Caption = "Adding..."
ClearAll
DisableAll
End Sub

Private Sub CBP_Click()
With RecCur
    .MovePrevious
    If .BOF Then .MoveFirst
    Populate
    LBLTOT.Caption = .AbsolutePosition & " of " & .RecordCount
End With
End Sub

Private Sub CBSearch_Click()
Dim StrSubject As String
StrSubject = InputBox("Enter Subject (part or whole word:", "Subject Search")
If Trim(StrSubject) = "" Then Exit Sub
With RecCur
    .Find "Subjectcode like '" & Trim(StrSubject) & "%'", , adSearchForward, 1
    If .EOF Then MsgBox "Search text not found.", vbInformation, "No Match": CBB_Click: Exit Sub
    Populate
    LBLTOT.Caption = RecCur.AbsolutePosition & " of " & RecCur.RecordCount
End With
End Sub

Private Sub CBupdate_Click()
On Error GoTo ErbX
If Trim(Tsubject.Text) = "" Or Trim(Tunits.Text) = "" Or _
    Trim(TDesc.Text) = "" Then
    MsgBox "Please Fill up all fields.", vbInformation, "Field Missing"
    Exit Sub
End If
If IsNumeric(Tunits.Text) = False Then
    MsgBox "Field Units must contain a Number.", vbInformation, "Field Missing"
    Tunits.SetFocus
    SendKeys "{HOME}+{END}"
    Exit Sub
End If
Dim msg As String, MsBox As String
Set RecAdd = Nothing
Set RecAdd = New ADODB.Recordset
With RecAdd
    .ActiveConnection = FrmInfoCNTR.ConX
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
Select Case Addix
    Case True       'Adding
        .Open "CURRICULAS"
        .AddNew
        MsBox = "Record Added."
    Case False      'Updating
        msg = "Select * from Curriculas Where Course"
        msg = msg & "='" & LBLCOURSE.Caption & "' and"
        msg = msg & " YearLevel='" & CYR.Text & "' and"
        msg = msg & " Schoolyear='" & TSCHLYR.Text & "' and"
        msg = msg & " Semester='" & CSEM.Text & "' and"
        msg = msg & " Subjectcode='" & RecCur.Fields("Subjectcode").Value & "'"
        If IsNull(LblMjr.Text) = False Or _
            Trim(LblMjr.Text) <> "" Then
        msg = msg & " and MAJOR='" & LblMjr.Text & "'"
        End If
        .Open msg
        MsBox = "Record Updated."
End Select
SetCurValue
.Update
.Properties.Refresh
.Close
        msg = "Select * from Curriculas Where Course"
        msg = msg & "='" & LBLCOURSE.Caption & "' and"
        msg = msg & " YearLevel='" & CYR.Text & "' and"
        msg = msg & " Schoolyear='" & TSCHLYR.Text & "' and"
        msg = msg & " Semester='" & CSEM.Text & "'"
        If IsNull(LblMjr.Text) Then GoTo Loader
        If Len(Trim(LblMjr.Text)) <> 0 Then
        msg = msg & " and MAJOR='" & LblMjr.Text & "'"
        End If
Loader:
LoadRecords msg
MsgBox MsBox, vbInformation, "Add Edit Command"
EnableAll
Addix = False
End With
Exit Sub
ErbX:
    ErrorTrap Err, "Add/Update Records"
    Me.Refresh
End Sub

Public Function AddNew(B As Boolean)
    
    'New Curriculum to Add
    TSCHLYR.Enabled = B
    CYR.Enabled = B
    CSEM.Enabled = B
    LblMjr.Enabled = B
    
    If B = True Then
        DisableSome
        CBNEW_Click
    End If
End Function

Sub DisableAll()
CBNEW.Enabled = 0
CBDel.Caption = "Cancel"
CBDel.Enabled = 1
CBSearch.Enabled = 0
CBupdate.Enabled = 1

CBB.Enabled = 0
CBP.Enabled = 0
CBN.Enabled = 0
CBL.Enabled = 0
End Sub

Sub EnableAll()
CBNEW.Enabled = 1
CBDel.Caption = "Delete"
CBupdate.Enabled = 1
CBDel.Enabled = 1
CBSearch.Enabled = 1

CBB.Enabled = 1
CBP.Enabled = 1
CBN.Enabled = 1
CBL.Enabled = 1
End Sub

Sub DisableSome()
'If no records
Addix = True
CBNEW.Enabled = 1
CBDel.Enabled = 0
CBupdate.Enabled = 0
CBSearch.Enabled = 0

CBB.Enabled = 0
CBP.Enabled = 0
CBN.Enabled = 0
CBL.Enabled = 0
End Sub

Sub LoadRecords(msg As String)
On Error GoTo ErbX
Set RecCur = Nothing
Set RecCur = New ADODB.Recordset
With RecCur
    .ActiveConnection = FrmInfoCNTR.ConX
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open msg
    If .RecordCount = 0 Then
        LBLTOT.Caption = "No Records"
        DisableSome
            TSCHLYR.Enabled = 1
            CYR.Enabled = 1
            CSEM.Enabled = 1
            LblMjr.Enabled = 1
            ClearAll
        Exit Sub
    Else
        LBLTOT.Caption = .AbsolutePosition & " of " & .RecordCount
        AddNew False
        Populate
    End If
End With
Exit Sub
ErbX:
    ErrorTrap Err, "Loading Records"
    Me.Refresh
End Sub

Sub SetCurValue()
On Error GoTo ErbX
With RecAdd
    .Fields("COURSE").Value = Trimers(LBLCOURSE, 1)
    .Fields("MAJOR").Value = Trimers(LblMjr, 2)
    .Fields("SCHOOLYEAR").Value = Trimers(TSCHLYR, 2)
    .Fields("Semester").Value = Trimers(CSEM, 2)
    .Fields("YEARLEVEL").Value = Trimers(CYR, 2)
    .Fields("SUBJECTCODE").Value = Trimers(Tsubject, 2)
    .Fields("UNITS").Value = Trimers(Tunits, 2)
    .Fields("DESCRIPTiON").Value = Trimers(TDesc, 2)
End With
Exit Sub
ErbX:
    ErrorTrap Err, "Setting Values"
    Me.Refresh
End Sub

Function Trimers(OBJ As Object, typ As Long) As String
If typ = 1 Then
OBJ.Caption = Trim(OBJ.Caption)
Trimers = OBJ.Caption
Else
OBJ.Text = Trim(OBJ.Text)
Trimers = OBJ.Text
End If
End Function

Sub ClearAll()
'TSchlyr.Text = ""
Tsubject.Text = ""
Tunits.Text = ""
TDesc.Text = ""
End Sub

Public Function MethodSplit(Str As String)
Dim VRS, msg As String
VRS = Split(Str, "|")
TSCHLYR.Text = VRS(0)
CYR.Text = VRS(1)
If UBound(VRS) > 2 Then
LblMjr.Text = VRS(2)
CSEM.Text = VRS(3)
Else
LblMjr.Text = ""
CSEM.Text = VRS(2)
End If
        msg = "Select * from Curriculas Where Course"
        msg = msg & "='" & LBLCOURSE.Caption & "' and"
        msg = msg & " YearLevel='" & CYR.Text & "' and"
        msg = msg & " Schoolyear='" & TSCHLYR.Text & "' and"
        msg = msg & " Semester='" & CSEM.Text & "'"
        If IsNull(LblMjr.Text) Then GoTo Loader
        If Len(Trim(LblMjr.Text)) <> 0 Then
        msg = msg & " and MAJOR='" & LblMjr.Text & "'"
        End If
Loader:
LoadRecords msg
End Function

Sub Populate()
With RecCur
    TSCHLYR.Text = .Fields("Schoolyear").Value
    Tsubject.Text = .Fields("subjectcode").Value
    Tunits.Text = .Fields("Units").Value
    TDesc.Text = .Fields("Description").Value
End With
End Sub

Private Sub Form_Load()
If EDCur = True Then
AddNew True
EDCur = False
Else
AddNew False
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
EDCur = False
End Sub

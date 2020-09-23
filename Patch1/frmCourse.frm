VERSION 5.00
Begin VB.Form frmCourse 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add/Edit Course"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   4800
   Icon            =   "frmCourse.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ISAPTECH.chameleonButton CBOK 
      Height          =   495
      Left            =   2400
      TabIndex        =   6
      Top             =   2280
      Width           =   1095
      _extentx        =   1931
      _extenty        =   873
      und             =   0   'False
      iname           =   "Tahoma"
      btype           =   14
      tx              =   "OK"
      enab            =   -1  'True
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   4210752
      bcolo           =   14737632
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
   Begin VB.ComboBox Cschl 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmCourse.frx":57E2
      Left            =   1560
      List            =   "frmCourse.frx":57EC
      TabIndex        =   5
      Text            =   "INTERNATIONAL SCHOOL OF ASIA AND THE PACIFIC"
      Top             =   1800
      Width           =   3135
   End
   Begin VB.TextBox Tdesc 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1560
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   600
      Width           =   3135
   End
   Begin VB.TextBox TCOURSE 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   3135
   End
   Begin ISAPTECH.chameleonButton CBCANCEL 
      Height          =   495
      Left            =   3600
      TabIndex        =   7
      Top             =   2280
      Width           =   1095
      _extentx        =   1931
      _extenty        =   873
      und             =   0   'False
      iname           =   "Tahoma"
      btype           =   14
      tx              =   "CANCEL"
      enab            =   -1  'True
      coltype         =   1
      focusr          =   -1  'True
      bcol            =   4210752
      bcolo           =   14737632
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
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "SCHOOL:"
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
      TabIndex        =   4
      Top             =   1800
      Width           =   795
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "DESCRIPTION:"
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
      TabIndex        =   2
      Top             =   600
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "COURSE CODE:"
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
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmCourse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public IsAddingCourse As Boolean    'if true then add
Public RsAddUpdate As ADODB.Recordset
Private Sub cbCancel_Click()
'Cancel Operation
IsAddingCourse = False
Unload Me
End Sub

Sub CheckIfAdd()
On Error GoTo ErbX
Dim MsBox As String
Set RsAddUpdate = Nothing
Set RsAddUpdate = New ADODB.Recordset
With RsAddUpdate
.ActiveConnection = FrmInfoCNTR.ConX
.CursorLocation = adUseClient
.CursorType = adOpenDynamic
.LockType = adLockOptimistic

Select Case IsAddingCourse
    Case True
            .Open "COURSES"
            .AddNew
            MsBox = "Record Added."
    Case False
        .Open "Select * From Courses where Course='" & FrmCCUr.LCourse.SelectedItem.Text & "'"
        MsBox = "Record Updated."
End Select
    SetCourseVal 1
    .Update
    .Properties.Refresh
    MsgBox MsBox, vbInformation, "Add Update Command"
    Exit Sub
End With
ErbX:
    ErrorTrap Err, "Add/Update Records"
    Me.Refresh
End Sub

Sub SetCourseVal(ByVal TypeX As Long)
With RsAddUpdate
Select Case TypeX
    Case 1      'GetNewvalues
        
            .Fields("Course").Value = Trim(TCOURSE.Text)
            .Fields("Description").Value = Trim(TDesc.Text)
            .Fields("School").Value = CSCHL.Text
        
    Case 2      'SetText
            TCOURSE.Text = .Fields("Course").Value
            TDesc.Text = .Fields("Description").Value
            CSCHL.Text = .Fields("School").Value
End Select
End With
End Sub

Private Sub CBOk_Click()

If Trim(TCOURSE.Text) = "" Or Trim(TDesc.Text) = "" Or _
    Trim(CSCHL.Text) = "" Then
    MsgBox "Please Fill up all fields.", vbInformation, "ERROR"
    TCOURSE.SetFocus
    Exit Sub
Else
    CheckIfAdd
End If
Unload Me

End Sub

Private Sub Form_Load()
If Me.IsAddingCourse = False Then       'Load Values
Loadvals
End If
End Sub

Sub Loadvals()
On Error GoTo ErbX
Set RsAddUpdate = Nothing
Set RsAddUpdate = New ADODB.Recordset
With RsAddUpdate
.ActiveConnection = FrmInfoCNTR.ConX
.CursorLocation = adUseClient
.CursorType = adOpenDynamic
.LockType = adLockOptimistic
.Open "Select * From Courses where Course='" & FrmCCUr.LCourse.SelectedItem.Text & "'"
SetCourseVal 2
End With
Exit Sub
ErbX:
    ErrorTrap Err, "Loading Records"
    Me.Refresh
End Sub

Sub CheckEntry()
'On Error GoTo erbx
Set RsAddUpdate = Nothing
Set RsAddUpdate = New ADODB.Recordset
With RsAddUpdate
If IsAddingCourse = True Then GoTo XPrS
If Trim$(TCOURSE.Text) = FrmCCUr.LCourse.SelectedItem.Text Then Exit Sub
XPrS:
.ActiveConnection = FrmInfoCNTR.ConX
.CursorLocation = adUseClient
.CursorType = adOpenDynamic
.LockType = adLockOptimistic
.Open "Select * From Courses where Course='" & Trim(TCOURSE.Text) & "'"
If .RecordCount <> 0 Then
    MsgBox "Course Code is already in used.", vbInformation, "Course"
    TCOURSE.SetFocus
    SendKeys "{Home}+{End}"
End If
End With
Exit Sub
ErbX:
    ErrorTrap Err, "Entry Checking"
    Me.Refresh
End Sub

Private Sub TDesc_GotFocus()
CheckEntry
End Sub

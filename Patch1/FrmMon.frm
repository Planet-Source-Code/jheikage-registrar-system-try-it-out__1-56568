VERSION 5.00
Begin VB.Form FrmMon 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Class/Subject Monitoring"
   ClientHeight    =   6450
   ClientLeft      =   8880
   ClientTop       =   735
   ClientWidth     =   3390
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   3390
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CBRec 
      Caption         =   "Count Class"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2040
      TabIndex        =   11
      Top             =   5760
      Width           =   1215
   End
   Begin VB.ListBox Lcls 
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
      Height          =   2670
      Left            =   120
      TabIndex        =   10
      Top             =   3000
      Width           =   3135
   End
   Begin VB.Frame FRM1 
      Appearance      =   0  'Flat
      Caption         =   "SUBJECT"
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "STUDENT ENROLLED:"
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
         Left            =   240
         TabIndex        =   8
         Top             =   1440
         Width           =   1665
      End
      Begin VB.Label LBLCOUNTSUB 
         AutoSize        =   -1  'True
         Caption         =   "COUNT"
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
         Left            =   1920
         TabIndex        =   5
         Top             =   1440
         Width           =   555
      End
      Begin VB.Label LBLSCHED 
         Caption         =   "SCHEDULE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   2685
      End
      Begin VB.Label LBLTEACHER 
         AutoSize        =   -1  'True
         Caption         =   "ACTIVE TEACHER"
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
         TabIndex        =   3
         Top             =   480
         Width           =   1275
      End
      Begin VB.Label LBLSUB 
         AutoSize        =   -1  'True
         Caption         =   "ACTIVE SUBJECT"
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
         TabIndex        =   2
         Top             =   240
         Width           =   1230
      End
   End
   Begin VB.Frame FRM2 
      Appearance      =   0  'Flat
      Caption         =   "CLASS"
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
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   3135
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "NUMBER OF STUDENT:"
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
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   1785
      End
      Begin VB.Label LBLCOUNTCLASS 
         AutoSize        =   -1  'True
         Caption         =   "COUNT"
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
         Left            =   2040
         TabIndex        =   7
         Top             =   480
         Width           =   555
      End
      Begin VB.Label LBLCLASS 
         AutoSize        =   -1  'True
         Caption         =   "ACTIVE CLASS"
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
         Top             =   240
         Width           =   1050
      End
   End
End
Attribute VB_Name = "FrmMon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetWindowsPos Lib "User32" Alias "SetWindowPos" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public RSSubjects As ADODB.Recordset
Attribute RSSubjects.VB_VarHelpID = -1
Public RSClassx As ADODB.Recordset
Attribute RSClassx.VB_VarHelpID = -1
Public RSClassy As ADODB.Recordset
Attribute RSClassy.VB_VarHelpID = -1
Public RSClass As ADODB.Recordset
Attribute RSClass.VB_VarHelpID = -1
Public RsRecount As ADODB.Recordset
Attribute RsRecount.VB_VarHelpID = -1
'Public Oparam(2) As ADODB.Parameter
Public OBJCom As ADODB.Command  'Recounting all records
Public ObjCom1 As ADODB.Command 'Count specific Class naman
Public ObjCom2 As ADODB.Command

Private Sub CBRec_Click()
Recount_AllClass
End Sub

Private Sub Form_Load()
'SetWindowsPos Me.hwnd, -1, 1, 1, 250, 350, &H10 Or &H40
End Sub

Function BaseCount(ix As String) As Double
Dim msg As String
Set RsRecount = Nothing
msg = "Select * From Classes where Section='" & ix & "';"
Set RsRecount = New ADODB.Recordset
Set ObjCom2 = New ADODB.Command
With ObjCom2
    .ActiveConnection = FrmInfoCNTR.ConX
    .CommandTimeout = 600
    .CommandType = adCmdText
    .CommandText = msg
    .Prepared = True
    Set RsRecount = .Execute
End With
With RsRecount
If IsNull(.Fields("Allowed").Value) Then
    BaseCount = 0
    Else
    BaseCount = .Fields("Allowed").Value
    End If
End With
End Function


Public Function CountTotal(ist As String) As Double
Dim MsgX As String
If FrmMon.RSClassy Is Nothing Then
Set FrmMon.RSClassy = New ADODB.Recordset
Else
FrmMon.RSClassy.Close
End If

MsgX = "SELECT IDNO" & _
" FROM Grading_Sys WHERE Grading_Sys.SCHOOLYEAR='" & FrmGRD.CSYR.Text & _
"'and  GRADING_SYS.Section='" & FrmGRD.CSECTION.Text & "' and  GRADING_SYS.Semester='" & FrmGRD.CSEM.Text & _
"'  GROUP BY GRADING_SYS.IDNO;"

Set OBJCom = New ADODB.Command
OBJCom.ActiveConnection = FrmInfoCNTR.ConX
OBJCom.CommandType = adCmdStoredProc
OBJCom.CommandText = "RE_COUNT"
OBJCom.CommandTimeout = 600
OBJCom.Parameters.Append OBJCom.CreateParameter("@SY", adVarChar, adParamInput, 10, FrmGRD.CSYR.Text)
OBJCom.Parameters.Append OBJCom.CreateParameter("@SEM", adVarChar, adParamInput, 3, FrmGRD.CSEM.Text)
OBJCom.Parameters.Append OBJCom.CreateParameter("@SECS", adVarChar, adParamInput, 10, ist)
Set RSClassy = OBJCom.Execute
RSClassy.Properties.Refresh
CountTotal = RSClassy.RecordCount
'    FrmMon.RSClassy.ActiveConnection = FrmInfoCNTR.ConX
'    FrmMon.RSClassy.CursorLocation = adUseClient
'    FrmMon.RSClassy.CursorType = adOpenDynamic
'    FrmMon.RSClassy.LockType = adLockOptimistic
'    FrmMon.RSClassy.Open MsgX, FrmInfoCNTR.ConX
'    CountTotal = FrmMon.RSClass.RecordCount

End Function


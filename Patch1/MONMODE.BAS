Attribute VB_Name = "MONMODE"
Public RsOthers As ADODB.Recordset
'ErrorMode! Something is wrong with sqlcodes
Public Sub Count_Subject(SY As String, SEM As String, TEACHER As String, SCHED As String, SUBJ As String)
'On Error Resume Next
Dim msg As String
With FrmGRD
    msg = "Select * From Grading_SYS where Schoolyear='"
    msg = msg & SY & "' and Semester = '" & SEM
    msg = msg & "' and Subject ='" & SUBJ & "'"
    msg = msg & " and Teacher = '" & TEACHER & "'"
    msg = msg & " and Schedule ='" & SCHED & "'"
    FrmMon.LBLSUB.Caption = SUBJ
    FrmMon.LBLTEACHER.Caption = TEACHER
    FrmMon.LBLSCHED.Caption = SCHED
    
End With
Set FrmMon.RSSubjects = Nothing
Set FrmMon.RSSubjects = New ADODB.Recordset
With FrmMon.RSSubjects
    .ActiveConnection = FrmInfoCNTR.ConX
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open msg, FrmInfoCNTR.ConX
    FrmMon.LBLCOUNTSUB.Caption = .RecordCount
    'Set Limits Here
    If .RecordCount >= 45 Then
        MsgBox "You are now having " & .RecordCount _
            & " records in this Subject.", vbInformation, "Enrolled"
    End If
    .Close
End With
Call Recount_AllClass
Set FrmMon.RSSubjects = Nothing
End Sub


Public Function Recount_AllClass()
Dim msg As String
Set FrmMon.RsRecount = Nothing
msg = "Select * From Classes Order by class, Year_Level"
Set FrmMon.RsRecount = New ADODB.Recordset
FrmMon.Lcls.Clear
With FrmMon.RsRecount
    .ActiveConnection = FrmInfoCNTR.ConX
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open msg
    Do Until .EOF
        FrmMon.Lcls.AddItem .Fields("Section").Value & " : " & FrmMon.CountTotal(.Fields("section").Value)
        .MoveNext
    Loop
End With
End Function

Public Function RetTotInSubs(Rs As ADODB.Recordset, SY As String, SEM As String, TEACHER As String, SCHED As String, SUBJ As String) As Integer
Dim msg As String
    msg = "Select * From Grading_SYS where Schoolyear='"
    msg = msg & SY & "' and Semester = '" & SEM
    msg = msg & "' and Subject ='" & SUBJ & "'"
    msg = msg & " and Teacher = '" & TEACHER & "'"
    msg = msg & " and Schedule ='" & SCHED & "'"
Set Rs = Nothing
Set Rs = New ADODB.Recordset
With Rs
    .ActiveConnection = FrmInfoCNTR.ConX
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open msg
    RetTotInSubs = .RecordCount
    .Close
End With
Set Rs = Nothing
End Function

Public Function GetTotinOthers(SY As String, SEM As String, ActCount As Integer) As Boolean
Dim msg As String
Set RsOthers = Nothing
Set RsOthers = New ADODB.Recordset
msg = "SELECT * FROM OTHERS where SCHOOLYEAR='" & _
    SY & "' and SEMESTER ='" & SEM & "'"
With RsOthers
    .ActiveConnection = FrmInfoCNTR.ConX
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open msg
    If .RecordCount = 0 Then
        MsgBox "SY " & SY & " of " & SEM & " does not have any limit for Subjects and Sections.", vbCritical, "ERROR"
        GetTotinOthers = False
        Exit Function
    End If
    Select Case ActCount
        Case Is = .Fields("STUDENT_PER_SUBJECT").Value
        MsgBox "Adding/Updating this record will exceed limit for Allowed students per subject. Action is Canceled.", vbInformation, "Limits"
        GetTotinOthers = False
        Exit Function
    End Select
    MsgBox ActCount & " Student\s are in this subject. Limit is " & .Fields("Student_PER_SUBJECT").Value & " Students.", vbInformation, "Limits"
    GetTotinOthers = True
End With
End Function

Public Function Count_Class(SY As String, SEM As String, SECS As String)
'On Error Resume Next
Dim msg As String, rcount As Double
Set FrmMon.RSClassx = Nothing
Set FrmMon.RSClassx = New ADODB.Recordset
msg = "Select SECTION,IDNO, SCHOOLYEAR, SEMESTER"
msg = msg & " From GRADING_SYS Where Schoolyear"
msg = msg & "='" & SY & "' and Semester='" & SEM & "'"
msg = msg & " and SECTION='" & SECS & "'"
msg = msg & " GROUP BY IDNO,SECTION,Schoolyear,Semester"
With FrmGRD
    FrmMon.LBLCLASS.Caption = SECS
End With
Set FrmMon.ObjCom1 = New ADODB.Command
With FrmMon.RSClassx
    .ActiveConnection = FrmInfoCNTR.ConX
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open msg
    '.CommandType = adCmdStoredProc
    '.CommandText = "Count_Class"
    '.Parameters.Append .CreateParameter("@SY", adVarChar, adParamInput, 10, FrmGRD.CSYR.Text)
    '.Parameters.Append .CreateParameter("@SEM", adVarChar, adParamInput, 3, FrmGRD.CSem.Text)
    '.Parameters.Append .CreateParameter("@SECS", adVarChar, adParamInput, 10, FrmGRD.CSECTION.Text)
    'Set FrmMon.RSClassx = .Execute
    'FrmMon.RSClassx.Properties.Refresh
    FrmMon.LBLCOUNTCLASS.Caption = FrmMon.RSClassx.RecordCount
End With
With FrmMon.RSClassx
    rcount = FrmMon.BaseCount(SECS)
        'If rcount = 0 Then
        '    If .RecordCount >= 45 Then
        '        MsgBox "You are now having " & .RecordCount _
        '        & " records in this Section.", vbInformation, "Enrolled"
        '    End If
        'Else
            If rcount <= .RecordCount Then
            MsgBox "You are now having " & .RecordCount _
            & " records in this " & SECS & ". This section allows " & rcount & " students.", vbInformation, "Enrolled"
            End If
    'End If

End With
Set FrmMon.RSClassx = Nothing
End Function

Public Function RetClass(Rs As ADODB.Recordset, SY As String, SEM As String, SECS As String) As Integer
msg = "Select SECTION,IDNO, SCHOOLYEAR, SEMESTER"
msg = msg & " From GRADING_SYS Where Schoolyear"
msg = msg & "='" & SY & "' and Semester='" & SEM & "'"
msg = msg & " and SECTION='" & SECS & "'"
msg = msg & " GROUP BY IDNO,SECTION,Schoolyear,Semester"
Set Rs = Nothing
Set Rs = New ADODB.Recordset
With Rs
    .ActiveConnection = FrmInfoCNTR.ConX
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open msg
    RetClass = .RecordCount
    .Close
End With
End Function

Public Function PermitClassAdd(SY As String, SEM As String, ActCount As Integer) As Boolean
Dim msg As String
Set RsOthers = Nothing
Set RsOthers = New ADODB.Recordset
msg = "SELECT * FROM OTHERS where SCHOOLYEAR='" & _
    SY & "' and SEMESTER ='" & SEM & "'"
With RsOthers
    .ActiveConnection = FrmInfoCNTR.ConX
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open msg
    If .RecordCount = 0 Then
        MsgBox "SY " & SY & " of " & SEM & " does not have any limit for Subjects and Sections.", vbCritical, "ERROR"
        PermitClassAdd = False
        Exit Function
    End If
    Select Case ActCount
        Case Is = .Fields("STUDENT_PER_CLASS").Value
        MsgBox "Adding/Updating this record will exceed limit for Allowed students per subject. Action is Canceled.", vbInformation, "Limits"
        PermitClassAdd = False
        Exit Function
    End Select
    MsgBox ActCount & " Student\s are in this SECTION. Limit is " & .Fields("Student_PER_CLASS").Value & " Students.", vbInformation, "Limits"
    PermitClassAdd = True
End With
End Function

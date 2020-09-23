Attribute VB_Name = "Functs"
Public Function ListSubs(GROUP As String, RS As ADODB.Recordset, OBJ As ComboBox)
Dim msg As String
msg = "Select Subject From Scheduling Where SectionKo like '%" & GROUP & "%' GROUP BY SUBJECT"
Set RS = Nothing
Set RS = New ADODB.Recordset
With RS
    .ActiveConnection = FrmInfoCNTR.ConX
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open msg
    OBJ.Clear
    Do Until .EOF
        OBJ.AddItem .Fields(0).Value
        .MoveNext
    Loop
End With
End Function


Public Function Get_Students(SY As String, SEM As String, SUBJECT As String, _
    SCHEDULE As String, TEACHER As String, RS As ADODB.Recordset, LC As ListView)
Dim msg As String, i As Long

msg = "Select * From Grading_SYS where SCHOOLYEAR='" & SY
msg = msg & "' and Semester = '" & SEM & "' and Subject = '"
msg = msg & SUBJECT & "' and Schedule='" & SCHEDULE & "' and Teacher='" & TEACHER & "' Order by SEX desc, IDNO, STUDENT, COURSE,YEARLEVEL"
Set RS = Nothing
Set RS = New ADODB.Recordset
With RS
    .ActiveConnection = FrmInfoCNTR.ConX
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open msg
'Load To ListView
    LC.ListItems.Clear
    Do Until .EOF
        i = .AbsolutePosition
        LC.ListItems.Add i, , .Fields("IDNO").Value, 1, 1
        CreateList LC, .Fields("STUDENT").Value, i, 1
        CreateList LC, .Fields("COURSE").Value, i, 2
        CreateList LC, .Fields("YEARLEVEL").Value, i, 3
        CreateList LC, .Fields("P1").Value, i, 4
        CreateList LC, .Fields("P2").Value, i, 5
        CreateList LC, .Fields("P3").Value, i, 6
        CreateList LC, .Fields("PRELIM").Value, i, 7
        CreateList LC, .Fields("m1").Value, i, 8
        CreateList LC, .Fields("m2").Value, i, 9
        CreateList LC, .Fields("m3").Value, i, 10
        CreateList LC, .Fields("MIDTERM").Value, i, 11
        CreateList LC, .Fields("s1").Value, i, 12
        CreateList LC, .Fields("s2").Value, i, 13
        CreateList LC, .Fields("s3").Value, i, 14
        CreateList LC, .Fields("SEMI").Value, i, 15
        CreateList LC, .Fields("F1").Value, i, 16
        CreateList LC, .Fields("F2").Value, i, 17
        CreateList LC, .Fields("F3").Value, i, 18
        CreateList LC, .Fields("FINALs").Value, i, 19
        CreateList LC, .Fields("REEXAM").Value, i, 20
        CreateList LC, .Fields("REMARKS").Value, i, 21
        .MoveNext
    Loop
End With
With LC
If .ListItems.Count = 0 Then Exit Function
    .ListItems(1).Selected = True
    .ListItems(1).EnsureVisible
    .SetFocus
End With
End Function

Function CreateList(LC As ListView, TEX As String, i As Long, x As Long)
LC.ListItems(i).SubItems(x) = TEX
End Function



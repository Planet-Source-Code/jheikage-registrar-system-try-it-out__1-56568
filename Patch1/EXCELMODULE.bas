Attribute VB_Name = "EXCELMODULE"
Dim ExclApp As Excel.Application
Dim ExcWork As Excel.Workbook
'CHECK GD
Public Function GetDrive(List As ListView)
Dim FileX As String, xso
FileX = "A:\Gradingdisk.Xls"
Set xso = CreateObject("Scripting.FileSystemObject")
If Not xso.fileexists(FileX) Then
    MsgBox "No GradingDisk file found in the Drive.", vbExclamation, "ERROR READING DISK"
    Set xso = Nothing
    Exit Function
End If
Set ExclApp = CreateObject("Excel.Application")
Set ExcWork = ExclApp.Workbooks.Open(FileX)
'Check Validity
If UCase(ExcWork.Application.Cells(1, 1)) = "IDNO" And _
    UCase(ExcWork.Application.Cells(1, 33)) = "REMARKS" _
    And UCase(ExcWork.Application.Cells(1, 34)) = "NOTE:" Then
    'Reading Begins
    FileRead List
Else
    MsgBox "Invalid GradingDisk.", vbCritical, "ERROR READING FILE"
    Set xso = Nothing
    ExcWork.Close False
    ExclApp.Workbooks.Close
    Exit Function
End If

Set xso = Nothing
ExcWork.Close False
ExclApp.Workbooks.Close
MsgBox "File Read Complete. Ready To udate Database.", vbInformation, "Read Complete"
Exit Function
End Function

Sub FileRead(LCONTROL As ListView)
Dim Pos As Long, x As Long, Y As Long
x = 1: Y = 1
With ExcWork.Application
    Do Until .Cells(x + 1, 1) = ""
    LCONTROL.ListItems.Add x, , .Cells(x + 1, 1), 1, 1
    'listSubitems
    For Y = 2 To 34
        If (Y) < 34 Then
        LCONTROL.ListItems(x).SubItems(Y - 1) = .Cells(x + 1, Y)
        End If
    Next
    x = x + 1
    Loop
End With
End Sub

'Sub ListSubs(LCONTROL As ListView, Pos As long, Row As long, Y As long)
'
'End Sub

'End Check
Public Function PrepareExport(Dialog As CommonDialog, List As ListView)
Dim xso
Dim i As Long, j As Long
Set xso = CreateObject("Scripting.FileSystemObject")
With Dialog
    .DialogTitle = "Export List To Excel"
    .Filter = "Excel Files(*.xls)|*.xls"
    .ShowSave
    If Trim(.FileName) = "" Then Exit Function
    If xso.fileexists(.FileName) Then
        MsgBox "File exist. Rename your file.", vbCritical, "ERROR"
        Exit Function
    End If
    xso.createTextFile .FileName, True
    Set ExclApp = CreateObject("Excel.Application")
    Set ExcWork = ExclApp.Workbooks.Open(.FileName)
End With
With List
    'i - Count Headers,j rows
    For i = 1 To List.ColumnHeaders.Count
    ExcWork.Application.Cells(1, i) = List.ColumnHeaders(i).Text
    Next
    For j = 1 To List.ListItems.Count
        For i = 1 To List.ColumnHeaders.Count
        ExcWork.Application.Cells(j + 1, 1) = List.ListItems(j).Text
        If i < List.ColumnHeaders.Count Then
        ExcWork.Application.Cells(j + 1, i + 1) = List.ListItems(j).SubItems(i)
        End If
        Next
    Next
    ExcWork.Save
    ExcWork.Saved = True
    ExclApp.Workbooks.Close
    Set xso = Nothing
    Set ExclApp = Nothing
    Set ExcWork = Nothing
End With
    MsgBox "Export Complete.", vbInformation, "Complte"
Exit Function
WriteError:
    ErrorTrap Err, "Exporting List"
    Set ExcWork = Nothing
    Set ExclApp = Nothing
    Set xso = Nothing
End Function

Public Function PrepareCS(Template As String, SQL As String, XP As CommonDialog)
'Set the Dialogbox
On Error GoTo WriteError
Dim xso, FSO, AppTemp As String, OPENFILE As String
With XP
If Right(App.Path, 1) = "\" Then
    AppTemp = App.Path & "Templates\" & Template
Else
    AppTemp = App.Path & "\Templates\" & Template
End If
    Set xso = CreateObject("Scripting.FileSystemObject")
    If Not xso.fileexists(AppTemp) Then
        MsgBox "Template do not exist.", vbCritical, "ERROR"
        Exit Function
    End If
    .DialogTitle = "Create Excel Control Sheet"
    .Filter = "Excel Files(*.xls)|*.xls"
    If Template <> "ControlSheet.xls" Then
        'Drive to A
        .InitDir = "A:\"
        .FileName = "A:\" & "GradingDisk.XLS"
    End If
    .ShowSave
    If Trim(.FileName) = "" Then Exit Function
    If xso.fileexists(.FileName) Then
        MsgBox "File exist. Rename your file.", vbCritical, "ERROR"
        Exit Function
    End If
    If Template <> "ControlSheet.xls" And UCase(.FileName) <> "A:\GRADINGDISK.XLS" Then
        MsgBox "You must create Grading Disk to a Removable Disk or do not change the name 'GRADINGDISK.XLS'.", vbCritical, "ERROR"
        Exit Function
    End If
    xso.copyfile AppTemp, .FileName, True
    Set ExclApp = CreateObject("Excel.Application")
    Set ExcWork = ExclApp.Workbooks.Open(.FileName)
    If Template = "ControlSheet.xls" Then
        CopyCSToEXCEL SQL
    Else
        
        CreateGD SQL
    End If
    ExcWork.Save
    ExcWork.Saved = True
    ExclApp.Workbooks.Close
    Set xso = Nothing
End With

MsgBox "File Transfer completed.", vbInformation, "Transfer Complete"
Exit Function
WriteError:
    ErrorTrap Err, "Writing Excel Files"
    Set xso = Nothing
    Set ExclApp = Nothing
    Set ExcWork = Nothing
End Function

Function CopyCSToEXCEL(SQL As String)   'Controlsheets
Dim SCHL As String
Set FrmControls.RsCS = Nothing
Set FrmControls.RsCS = New ADODB.Recordset
With FrmControls.RsCS
    .ActiveConnection = FrmInfoCNTR.ConX
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open SQL
    If .RecordCount = 0 Then
        MsgBox "There are no records in this Schedule.", vbInformation, "No Students"
        Exit Function
    End If
    SCHL = InputBox("SCHOOL:", "SCHOOL")
    writeHead SCHL
    Do Until .EOF
        WriteStudents .AbsolutePosition - 1, FrmControls.RsCS
        .MoveNext
    Loop
End With
End Function

Function CreateGD(SQL As String)
Set FrmTrans.RsTRANS = Nothing
Set FrmTrans.RsTRANS = New ADODB.Recordset
With FrmTrans.RsTRANS
    .ActiveConnection = FrmInfoCNTR.ConX
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open SQL
    If .RecordCount = 0 Then
        MsgBox "There are no records in this Schedule.", vbInformation, "No Students"
        Exit Function
    End If
    Do Until .EOF
    'create list
        WriteGD .AbsolutePosition, FrmTrans.RsTRANS
        .MoveNext
    Loop
End With
End Function

Sub writeHead(SCHL As String)   'For Creating Control Sheets
With ExcWork
    .Application.Cells(1, 1) = SCHL 'SCHOOL
    .Application.Cells(5, 2) = FrmControls.TSCHLYR.Text 'SY
    .Application.Cells(5, 4) = FrmControls.CSem.Text 'SEM
    .Application.Cells(7, 2) = FrmControls.CSubjects.Text 'SUBJECT
    .Application.Cells(7, 5) = FrmControls.Tunits.Text 'UNITS
    .Application.Cells(9, 2) = FrmControls.Cteacher.Text 'TEACHER
    .Application.Cells(10, 2) = FrmControls.Tsched.Text 'SCHED
    .Application.Cells(8, 3) = FrmControls.TSD.Text 'SCHED
End With
End Sub

Sub WriteStudents(Row As Long, rs As ADODB.Recordset)
With ExcWork
    .Application.Cells(12 + Row, 1) = rs.Fields("IDNO").Value
    .Application.Cells(12 + Row, 2) = rs.Fields("STUDENT").Value
    .Application.Cells(12 + Row, 3) = rs.Fields("COURSE").Value
    .Application.Cells(12 + Row, 4) = rs.Fields("YEARLEVEL").Value
    .Application.Cells(12 + Row, 5) = rs.Fields("SEX").Value
    .Application.Cells(12 + Row, 6) = rs.Fields("MAJOR").Value
    .Application.Cells(12 + Row, 7) = rs.Fields("Section").Value
    
    .Application.Cells(12 + Row, 8) = rs.Fields("p1").Value
    .Application.Cells(12 + Row, 9) = rs.Fields("p2").Value
    .Application.Cells(12 + Row, 10) = rs.Fields("p3").Value
    .Application.Cells(12 + Row, 11) = rs.Fields("PRELIM").Value
    
    .Application.Cells(12 + Row, 12) = rs.Fields("m1").Value
    .Application.Cells(12 + Row, 13) = rs.Fields("m2").Value
    .Application.Cells(12 + Row, 14) = rs.Fields("m3").Value
    .Application.Cells(12 + Row, 15) = rs.Fields("MIDTERM").Value
    
    .Application.Cells(12 + Row, 16) = rs.Fields("s1").Value
    .Application.Cells(12 + Row, 17) = rs.Fields("s2").Value
    .Application.Cells(12 + Row, 18) = rs.Fields("s3").Value
    .Application.Cells(12 + Row, 19) = rs.Fields("SEMI").Value
    
    .Application.Cells(12 + Row, 20) = rs.Fields("f1").Value
    .Application.Cells(12 + Row, 21) = rs.Fields("f2").Value
    .Application.Cells(12 + Row, 22) = rs.Fields("F3").Value
    .Application.Cells(12 + Row, 23) = rs.Fields("FINALS").Value
    
    .Application.Cells(12 + Row, 24) = rs.Fields("REEXAM").Value
    .Application.Cells(12 + Row, 25) = rs.Fields("REMARKS").Value
    
End With
End Sub

Sub WriteGD(Row As Long, rs As ADODB.Recordset)
With ExcWork
    .Application.Cells(1 + Row, 1) = rs.Fields("IDNO").Value
    .Application.Cells(1 + Row, 2) = rs.Fields("STUDENT").Value
    .Application.Cells(1 + Row, 3) = rs.Fields("SEX").Value
    .Application.Cells(1 + Row, 4) = rs.Fields("SCHOOLYEAR").Value
    .Application.Cells(1 + Row, 5) = rs.Fields("SEMESTER").Value
    .Application.Cells(1 + Row, 6) = rs.Fields("COURSE").Value
    .Application.Cells(1 + Row, 7) = rs.Fields("YEARLEVEL").Value
    
    .Application.Cells(1 + Row, 8) = rs.Fields("MAJOR").Value
    .Application.Cells(1 + Row, 9) = rs.Fields("SCHOOL").Value
    .Application.Cells(1 + Row, 10) = rs.Fields("SECTION").Value
    .Application.Cells(1 + Row, 11) = rs.Fields("SUBJECT").Value
    
    .Application.Cells(1 + Row, 12) = rs.Fields("UNITS").Value
    .Application.Cells(1 + Row, 13) = rs.Fields("SUBJECT_DESCRIPTION").Value
    .Application.Cells(1 + Row, 14) = rs.Fields("TEACHER").Value
    .Application.Cells(1 + Row, 15) = rs.Fields("SCHEDULE").Value
        
    .Application.Cells(1 + Row, 16) = rs.Fields("p1").Value
    .Application.Cells(1 + Row, 17) = rs.Fields("p2").Value
    .Application.Cells(1 + Row, 18) = rs.Fields("p3").Value
    .Application.Cells(1 + Row, 20) = rs.Fields("m1").Value
    
    .Application.Cells(1 + Row, 21) = rs.Fields("m2").Value
    .Application.Cells(1 + Row, 22) = rs.Fields("m3").Value
    .Application.Cells(1 + Row, 24) = rs.Fields("s1").Value
    .Application.Cells(1 + Row, 25) = rs.Fields("s2").Value
    .Application.Cells(1 + Row, 26) = rs.Fields("s3").Value
    .Application.Cells(1 + Row, 28) = rs.Fields("f1").Value
    .Application.Cells(1 + Row, 29) = rs.Fields("f2").Value
    .Application.Cells(1 + Row, 30) = rs.Fields("f3").Value
    
    .Application.Cells(1 + Row, 32) = rs.Fields("REEXAM").Value
    .Application.Cells(1 + Row, 33) = rs.Fields("REMARKS").Value
    
End With

End Sub

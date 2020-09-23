Attribute VB_Name = "SetModule"
'ID CONVERSION METHOD
'CREATING ID
'------------------
'SORT ALL SPI RECORDS BY ID
'NEW ID = LAST ID + 1
'NEW ID can BE EDITED
'Provider=SQLOLEDB.1;Persist Security Info=False;User ID=JHEI;Initial Catalog=REGSYS;Initial File Name=C:\MSSQL7\Data\REGSYS_Data.MDF
'Provider=SQLOLEDB.1;Persist Security Info=False;User ID=JHEI;Password=mouse;Initial Catalog=REGSYS;Initial File Name=C:\MSSQL7\Data\REGSYS_Data.MDF
'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\ISAPMCNP\Patch1\Dbase\RegSys2k2.mdb;Persist Security Info=False
Public Nigolx As New nigol.SparkClass
Public XXX As New TypeMan.CDLLXMEN
Public ComSel As ADODB.Command
Public ComIn As ADODB.Command
Public WHOLOG As String, WHOPASS As String
Public SqlStGrd As String
Public Srch As String, SRCHTYP As Boolean
Public Const P_Val_MAK As String = "1x"
Public Const P_Val_KAM As String = "2x"
Public STORED_PROC_MAT As String
Sub MaxOp()
    Dim File As String, gig As String, _
    enc As String, sync As String, maxs As Integer
    If Right(App.Path, 1) = "\" Then
    File = App.Path & "\inf\NIGOLFILE.FJV"
    Else
    File = App.Path & "\inf\" & "NIGOLFILE.FJV"
    End If
    gig = "[USER:" & WHOLOG & "]" & "[DATE:" & Date & "] [TIME:" & Time & "]"
    Dim i As Integer
    maxs = Len(gig)
    For i = 1 To maxs
        'Encript File
        sync = Left(gig, 1)
        gig = Right(gig, maxs - i)
        
        'MsgBox gig
        enc = enc & Chr(234) & sync
    Next
    frmuser.GetFileName File
    frmuser.Show 1
    
    Open File For Output As #1
        
        Write #1, enc
        
    Close #1
    
End Sub

Sub Main()
'ConnectSQLSERVER
Set Nigolx = New nigol.SparkClass
Set XXX = New TypeMan.CDLLXMEN
Nigolx.ShowForm
If Nigolx.LoginEnable(True) = True Then
    WHOLOG = Nigolx.XNAM
    WHOPASS = Nigolx.XPAS
    '<<>>'
    Types = "System Administrator"
    '<<>>'
    Dim FSO
    
    Set FSO = CreateObject("Scripting.FilesystemObject")
        If Not FSO.fileexists(App.Path & "\Inf\SQLDB.ini") Then
            MsgBox "Can't Continue loading Application." & vbNewLine & _
                "The File SQLDB.ini which is needed in loading the application is missing.", vbCritical, "ERROR Loading Applciation"
            WriteLog FrmSet.Routputbox, vbTab & "Initial File Loading - FAILED"
            EndSys
        End If
        
        If Not FSO.fileexists(App.Path & "\inf\Nigolfile.fjv") Then
            MsgBox "Can't Continue loading Application." & vbNewLine & _
                "The File FJV which is needed in loading the application is missing.", vbCritical, "ERROR Loading Applciation"
            WriteLog FrmSet.Routputbox, vbTab & "Security File Loading - FAILED"
            EndSys
        End If
        
    MaxOp
    isLogSys = False
    isLogDb = False
    FrmSet.Show
    SetStyle
    FrmSet.Routputbox.Text = GetStartDate
    Dim Mxg As String
    Mxg = vbTab & "System Security Login Name:" & WHOLOG & vbNewLine
    Mxg = Mxg & vbTab & "Application Type: " & Types & " - STARTED"
    'Write Messages----
    WriteLog FrmSet.Routputbox, vbTab & "Initial File Loading - Success"
    WriteLog FrmSet.Routputbox, vbTab & "Security File Loading - Success"
    WriteLog FrmSet.Routputbox, Mxg
    
    '-------------
    FrmSet.Sbar.Panels(3).Text = Types & " LOGGED USER:" & WHOLOG
    PerSecIn = RetriveIniValues("Persist Security info", "False")
    InCat = RetriveIniValues("Initial Catalog", "RegSys")
    InFiNa = RetriveIniValues("Initial FileName", "c:\MSSQL7\Data\RegSys_data.mdf")
    
End If
End Sub
Sub EndSys()
With FrmInfoCNTR
    If .ConX Is Nothing Then GoTo WRITEX
    If .ConX.State <> 0 Then .ConX.Close
    
End With
WRITEX:
FrmSet.Routputbox.SaveFile App.Path & "\AppLog.Doc"
End
End Sub

'-----------------------------------------------------
Sub ConnectSQLSERVER(User As String, Pass As String)
Dim msg As String
On Error GoTo ErrorX

        SetVal User, Pass
        
Exit Sub
ErrorX:
    msg = vbTab & "<<ERROR OCCURED!" & vbNewLine
    msg = msg & vbTab & "ERROR NUMBER:" & Err.Number
    msg = msg & vbNewLine & vbTab & "DESCRIPTION:"
    msg = msg & vbNewLine & vbTab & Err.Description
    msg = msg & vbNewLine & vbTab & "Source: SERVER CONNECTION..>>"
    WriteLog FrmSet.Routputbox, msg
    MsgBox Err.Description & _
    vbCrLf & vbCrLf & _
    "Database Access denied." _
    , vbCritical, "ERROR:" & Err.Number
    isLogDb = False
    If DENVER.SQLCON.State <> 0 Then DENVER.SQLCON.Close
    If Not FrmInfoCNTR.ConX Is Nothing Then
        If FrmInfoCNTR.ConX.State <> 0 Then FrmInfoCNTR.ConX.Close
    End If
End Sub

Sub SetVal(User As String, Pass As String)
Dim QB As String, XB As String
Dim xx As String
Dim msg As String

QB = "Provider=SQLOLEDB.1;Persist Security Info=" & PerSecIn & ";User ID=" & User & ";PASSWORD=" & Pass & ";Initial Catalog=" & InCat & ";Initial File Name=" & InFiNa & ";"
'MsgBox QB

With DENVER
    If .SQLCON.State <> 0 Then
        MsgBox "You are currently Connected.", vbInformation, "Connection"
        Exit Sub
    End If
msg = vbTab & "Database Log in using " & User & " as user name..."
msg = msg & vbNewLine & vbTab & "Password NOT SHOWN"
msg = msg & vbNewLine & vbTab & "Time: " & Time & " ..."
SysOutWin.WriteLog FrmSet.Routputbox, msg
    
    XB = _
    "Provider=MSDatashape.1;Data Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Dbase\RegSys.mdb;Persist Security Info=False"
    '"Provider=MSDatashape.1;data Provider=SQLOLEDB.1;" & "Persist Security Info=" & PerSecIn & ";User ID=" & User & ";Password=" & Pass & ";Initial Catalog=" & InCat & ";Initial File Name=" & InFiNa & ";"
    
    .SQLCON.Open XB

End With

    If FrmInfoCNTR.ConX Is Nothing Then
        Set FrmInfoCNTR.ConX = New ADODB.Connection
        
    Else
        If FrmInfoCNTR.ConX.State <> adStateClosed Then
            FrmInfoCNTR.ConX.Close
        End If
    End If

    FrmInfoCNTR.ConX.CursorLocation = adUseClient
    FrmInfoCNTR.ConX.Open XB '"Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=C:\ISAPMCNP\Patch1\Dbase\RegSys.mdb" 'Qb
    'MsgBox FrmInfoCNTR.ConX.State
    msg = vbTab & "Log in Success... " & vbNewLine
    msg = msg & vbTab & "Database Transaction Begin>>"
    WriteLog FrmSet.Routputbox, msg
    MsgBox "Database Connected.", vbInformation, "Connected"
    isLogDb = True
    Userx = User: PassX = Pass
End Sub
Sub LoadRecordsToList(SQL As String)
Dim i As Long

Set FrmInfoCNTR.ConRec = New ADODB.Recordset
With FrmInfoCNTR.ConRec
    .ActiveConnection = FrmInfoCNTR.ConX
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open SQL
End With

With FrmInfoCNTR.LDVIEW
    .ListItems.Clear
    Do Until FrmInfoCNTR.ConRec.EOF
        i = FrmInfoCNTR.ConRec.AbsolutePosition
        .ListItems.Add i, , FrmInfoCNTR.ConRec.Fields(0).Value, , 1
        .ListItems.Item(i).SubItems(1) = FrmInfoCNTR.ConRec.Fields(1).Value
        .ListItems(i).SubItems(2) = FrmInfoCNTR.ConRec.Fields("Sex").Value
        FrmInfoCNTR.ConRec.MoveNext
    Loop
    FrmInfoCNTR.LBLNUML.Caption = FrmInfoCNTR.ConRec.RecordCount
End With
Set FrmInfoCNTR.ConRec = Nothing
End Sub

Public Function ErrorTrap(Erb As ErrObject, Source As String)
Dim msg As String
Dim Xmsg As String
Xmsg = vbTab & "<<Error Occured!" & vbNewLine
Xmsg = Xmsg & vbTab & "Source: " & Source & vbNewLine
Xmsg = Xmsg & vbTab & "Reason: " & Err.Description & vbNewLine
Xmsg = Xmsg & vbTab & "Error End>>...Trap Success"
WriteLog FrmSet.Routputbox, Xmsg
msg = "Error Occured in the System at " & Source & "." & vbNewLine
msg = msg & "REASON:" & vbNewLine & vbNewLine
msg = msg & Erb.Description & vbNewLine & vbNewLine
msg = msg & "Click OK to refresh."
MsgBox msg, vbCritical, "ERROR " & Erb.Number
End Function


Public Sub SetDenver(msg As String, RecordSets As Long)
With DENVER
    Dim MsgX As String
    MsgX = "Provider=MSDataShape.1;Extended Properties='Initial File Name=" & InFiNa & "';Persist Security Info=" & PerSecIn & ";User ID=" & Userx & ";Password=" & PassX & ";Initial Catalog=" & InCat & ";Data Provider=SQLOLEDB.1"
    '"Provider=MSDataShape.1;Persist Security Info=False;Data Source=c:\ISAPMCNP\PATCH1\DBASE\REGSYS.mdb;Data Provider=MICROSOFT.JET.OLEDB.4.0"
    If .SQLCON.State <> 0 Then .SQLCON.Close
    .SQLCON.Open MsgX
    Select Case RecordSets
    Case 1
    If .rsCmdCourse.State <> 0 Then
    .rsCmdCourse.Close
    End If
    .rsCmdCourse.Open msg, .SQLCON
    Case 2
    If .rscmdCurHead.State <> 0 Then
    .rscmdCurHead.Close
    End If
    .rscmdCurHead.Open msg, .SQLCON
    Case 3  'Print PIS
    .rscmdPerInfo.Open msg, .SQLCON
    Case 4  'Print Class Report
    .rsClassReport.Open msg, .SQLCON
    Case 5  'Print SCR
    .rsHeadGrades.Open msg, .SQLCON
    End Select
End With
End Sub

Sub SetStyle()
With FrmSet
    Select Case XXX.ApplicationType
        Case P_Val_MAK
        Types = "Client Users"
        .CBCRS.Enabled = False
        .cbSettings.Enabled = False
        Case P_Val_KAM
        Types = "System Administrator"
        .CBCRS.Enabled = True
        .cbSettings.Enabled = True
    End Select
    STORED_PROC_MAT = XXX.ApplicationType
End With
End Sub

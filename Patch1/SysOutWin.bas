Attribute VB_Name = "SysOutWin"
Public PerSecIn As String
Public InCat As String
Public InFiNa As String
Public Types As String
Public Userx As String, PassX As String
Private Declare Function GetPrivateProfileString _
Lib "kernel32" Alias "GetPrivateProfileStringA" _
(ByVal lpApplicationName As String, _
ByVal lpKeyName As Any, ByVal lpDefault As String, _
ByVal lpReturnedString As String, _
ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Long

Public Function GetStartDate() As String
GetStartDate = "System Started>>" & vbNewLine
GetStartDate = GetStartDate & vbTab & "Date: " & Date & vbNewLine
GetStartDate = GetStartDate & vbTab & "Time: " & Time
End Function

Public Function WriteLog(Cntr As RichTextBox, msg As String)
Cntr.Text = Cntr.Text & vbNewLine & msg
End Function

Public Function RetriveIniValues(Key As String, Value As String) As String  'GetValues

Const maxs As Long = 500
Dim r As Long, RetStr As String * 500
Dim se As String
Dim x As String, rt As String
Dim Xne As String, ixp As Integer
se = "Database"
GetPrivateProfileString se, Key, Value, RetStr, maxs, App.Path & "\inf\SQLDB.ini"

For r = 1 To 500
x = Right(Left(RetStr, r), 1)
If r <> 500 Then
ixp = Asc(Right(Left(RetStr, r + 1), 1))
Else
ixp = 2
End If
If Asc(x) = 0 And ixp = 0 Then
    Exit For
Else
    rt = rt & x
End If
Next
RetriveIniValues = Trim(rt)
End Function


Public Function ReverIniValues(Key As String, Value As String)   'Set Values
WritePrivateProfileString "Database", Key, Value, App.Path & "\inf\SQLDB.ini"
End Function

VERSION 5.00
Begin VB.MDIForm MdiHold 
   BackColor       =   &H8000000C&
   Caption         =   "ISAP-MCNP Registrar"
   ClientHeight    =   5130
   ClientLeft      =   60
   ClientTop       =   600
   ClientWidth     =   6300
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Menu MBFile 
      Caption         =   "&FILE"
      Begin VB.Menu FLog 
         Caption         =   "Log In User"
         Shortcut        =   ^L
      End
      Begin VB.Menu FBREAK 
         Caption         =   "-"
      End
      Begin VB.Menu FExit 
         Caption         =   "Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu MBReg 
      Caption         =   "&Registrar"
      Begin VB.Menu RInfo 
         Caption         =   "Information Center"
         Shortcut        =   ^I
      End
      Begin VB.Menu RChange 
         Caption         =   "&Chage Accounts"
         Shortcut        =   ^C
      End
   End
   Begin VB.Menu Wilist 
      Caption         =   "&Windows"
      WindowList      =   -1  'True
   End
   Begin VB.Menu MHelp 
      Caption         =   "Help"
   End
End
Attribute VB_Name = "MdiHold"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub MDIForm_Unload(Cancel As Integer)
EndSys
End Sub

Private Sub RInfo_Click()
ConnectSQLSERVER
FrmInfoCNTR.Show
End Sub

VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmGRD 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "College Record"
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   2205
   ClientWidth     =   9780
   Icon            =   "FrmGRD.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   9780
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5040
      Top             =   120
   End
   Begin VB.Frame FRM6 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   6495
      Left            =   7320
      TabIndex        =   67
      Top             =   120
      Width           =   2295
      Begin VB.ComboBox CMovSem 
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
         ItemData        =   "FrmGRD.frx":57E2
         Left            =   120
         List            =   "FrmGRD.frx":57EF
         Style           =   2  'Dropdown List
         TabIndex        =   82
         Top             =   5040
         Width           =   1095
      End
      Begin ISAPTECH.chameleonButton CBMOVE 
         Height          =   1455
         Left            =   1320
         TabIndex        =   81
         Top             =   3960
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   2566
         BOLD            =   0   'False
         ITA             =   0   'False
         INAME           =   "Tahoma"
         SIZE            =   8.25
         UND             =   0   'False
         BTYPE           =   14
         TX              =   "Move to Selected SY and SEM"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   8421504
         BCOLO           =   14737632
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.ListBox LSYMove 
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
         Height          =   990
         Left            =   120
         TabIndex        =   80
         Top             =   3960
         Width           =   1095
      End
      Begin ISAPTECH.chameleonButton CBF 
         Height          =   375
         Left            =   120
         TabIndex        =   73
         Top             =   6000
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         BOLD            =   0   'False
         ITA             =   0   'False
         INAME           =   "Tahoma"
         SIZE            =   8.25
         UND             =   0   'False
         BTYPE           =   4
         TX              =   "|<"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   8421504
         BCOLO           =   14737632
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ISAPTECH.chameleonButton CBADD 
         Height          =   615
         Left            =   120
         TabIndex        =   68
         Top             =   240
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         BOLD            =   0   'False
         ITA             =   0   'False
         INAME           =   "Tahoma"
         SIZE            =   8.25
         UND             =   0   'False
         BTYPE           =   14
         TX              =   "Add Subject"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   8421504
         BCOLO           =   14737632
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ISAPTECH.chameleonButton CBUpdate 
         Height          =   615
         Left            =   120
         TabIndex        =   69
         Top             =   960
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         BOLD            =   0   'False
         ITA             =   0   'False
         INAME           =   "Tahoma"
         SIZE            =   8.25
         UND             =   0   'False
         BTYPE           =   14
         TX              =   "Update Subject"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   8421504
         BCOLO           =   14737632
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ISAPTECH.chameleonButton CBDel 
         Height          =   615
         Left            =   120
         TabIndex        =   70
         Top             =   1680
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         BOLD            =   0   'False
         ITA             =   0   'False
         INAME           =   "Tahoma"
         SIZE            =   8.25
         UND             =   0   'False
         BTYPE           =   14
         TX              =   "Delete Subject"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   8421504
         BCOLO           =   14737632
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ISAPTECH.chameleonButton CBSearch 
         Height          =   615
         Left            =   120
         TabIndex        =   71
         Top             =   2400
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         BOLD            =   0   'False
         ITA             =   0   'False
         INAME           =   "Tahoma"
         SIZE            =   8.25
         UND             =   0   'False
         BTYPE           =   14
         TX              =   "Search Subject"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   8421504
         BCOLO           =   14737632
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ISAPTECH.chameleonButton CbPrint 
         Height          =   615
         Left            =   120
         TabIndex        =   72
         Top             =   3120
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         BOLD            =   0   'False
         ITA             =   0   'False
         INAME           =   "Tahoma"
         SIZE            =   8.25
         UND             =   0   'False
         BTYPE           =   14
         TX              =   "Print College Record"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   8421504
         BCOLO           =   14737632
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ISAPTECH.chameleonButton CBPR 
         Height          =   375
         Left            =   360
         TabIndex        =   74
         Top             =   6000
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         BOLD            =   0   'False
         ITA             =   0   'False
         INAME           =   "Tahoma"
         SIZE            =   8.25
         UND             =   0   'False
         BTYPE           =   4
         TX              =   "<"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   8421504
         BCOLO           =   14737632
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ISAPTECH.chameleonButton CBNX 
         Height          =   375
         Left            =   1680
         TabIndex        =   75
         Top             =   6000
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         BOLD            =   0   'False
         ITA             =   0   'False
         INAME           =   "Tahoma"
         SIZE            =   8.25
         UND             =   0   'False
         BTYPE           =   4
         TX              =   ">"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   8421504
         BCOLO           =   14737632
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ISAPTECH.chameleonButton CBL 
         Height          =   375
         Left            =   1920
         TabIndex        =   76
         Top             =   6000
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   661
         BOLD            =   0   'False
         ITA             =   0   'False
         INAME           =   "Tahoma"
         SIZE            =   8.25
         UND             =   0   'False
         BTYPE           =   4
         TX              =   ">|"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   8421504
         BCOLO           =   14737632
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label LBLREC 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Records"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   600
         TabIndex        =   77
         Top             =   6000
         Width           =   1095
      End
   End
   Begin VB.Frame FRM4 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   6495
      Left            =   4560
      TabIndex        =   24
      Top             =   120
      Width           =   2655
      Begin TabDlg.SSTab TBGrades 
         Height          =   3015
         Left            =   120
         TabIndex        =   25
         Top             =   600
         Width           =   2340
         _ExtentX        =   4128
         _ExtentY        =   5318
         _Version        =   393216
         TabOrientation  =   1
         Style           =   1
         Tabs            =   4
         TabsPerRow      =   2
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "PRELIMS"
         TabPicture(0)   =   "FrmGRD.frx":5802
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label10"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label11"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label12"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label13"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Line1"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "tcs1"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "tq1"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "tt1"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "tave1"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "CBCom1"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).ControlCount=   10
         TabCaption(1)   =   "MIDTERMS"
         TabPicture(1)   =   "FrmGRD.frx":581E
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "TAVE2"
         Tab(1).Control(1)=   "TT2"
         Tab(1).Control(2)=   "TQ2"
         Tab(1).Control(3)=   "TCS2"
         Tab(1).Control(4)=   "CBCom2"
         Tab(1).Control(5)=   "Line2"
         Tab(1).Control(6)=   "Label17"
         Tab(1).Control(7)=   "Label16"
         Tab(1).Control(8)=   "Label15"
         Tab(1).Control(9)=   "Label14"
         Tab(1).ControlCount=   10
         TabCaption(2)   =   "SEMI-FINALS"
         TabPicture(2)   =   "FrmGRD.frx":583A
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "TAVE3"
         Tab(2).Control(1)=   "TT3"
         Tab(2).Control(2)=   "TQ3"
         Tab(2).Control(3)=   "TCS3"
         Tab(2).Control(4)=   "CBCom3"
         Tab(2).Control(5)=   "Line3"
         Tab(2).Control(6)=   "Label21"
         Tab(2).Control(7)=   "Label20"
         Tab(2).Control(8)=   "Label19"
         Tab(2).Control(9)=   "Label18"
         Tab(2).ControlCount=   10
         TabCaption(3)   =   "FINALS"
         TabPicture(3)   =   "FrmGRD.frx":5856
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "TAVE4"
         Tab(3).Control(1)=   "TT4"
         Tab(3).Control(2)=   "TQ4"
         Tab(3).Control(3)=   "TCS4"
         Tab(3).Control(4)=   "CBCom4"
         Tab(3).Control(5)=   "Line4"
         Tab(3).Control(6)=   "Label25"
         Tab(3).Control(7)=   "Label24"
         Tab(3).Control(8)=   "Label23"
         Tab(3).Control(9)=   "Label22"
         Tab(3).ControlCount=   10
         Begin VB.TextBox TAVE4 
            Enabled         =   0   'False
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
            Left            =   -74160
            TabIndex        =   59
            Top             =   1320
            Width           =   495
         End
         Begin VB.TextBox TT4 
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
            Left            =   -74160
            TabIndex        =   57
            Top             =   840
            Width           =   495
         End
         Begin VB.TextBox TQ4 
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
            Left            =   -74160
            TabIndex        =   55
            Top             =   480
            Width           =   495
         End
         Begin VB.TextBox TCS4 
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
            Left            =   -74160
            TabIndex        =   53
            Top             =   120
            Width           =   495
         End
         Begin VB.TextBox TAVE3 
            Enabled         =   0   'False
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
            Left            =   -74160
            TabIndex        =   50
            Top             =   1320
            Width           =   495
         End
         Begin VB.TextBox TT3 
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
            Left            =   -74160
            TabIndex        =   48
            Top             =   840
            Width           =   495
         End
         Begin VB.TextBox TQ3 
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
            Left            =   -74160
            TabIndex        =   46
            Top             =   480
            Width           =   495
         End
         Begin VB.TextBox TCS3 
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
            Left            =   -74160
            TabIndex        =   44
            Top             =   120
            Width           =   495
         End
         Begin VB.TextBox TAVE2 
            Enabled         =   0   'False
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
            Left            =   -74160
            TabIndex        =   41
            Top             =   1320
            Width           =   495
         End
         Begin VB.TextBox TT2 
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
            Left            =   -74160
            TabIndex        =   39
            Top             =   840
            Width           =   495
         End
         Begin VB.TextBox TQ2 
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
            Left            =   -74160
            TabIndex        =   37
            Top             =   480
            Width           =   495
         End
         Begin VB.TextBox TCS2 
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
            Left            =   -74160
            TabIndex        =   35
            Top             =   120
            Width           =   495
         End
         Begin ISAPTECH.chameleonButton CBCom1 
            Height          =   375
            Left            =   1200
            TabIndex        =   34
            Top             =   1800
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   661
            BOLD            =   0   'False
            ITA             =   0   'False
            INAME           =   "Tahoma"
            SIZE            =   8.25
            UND             =   0   'False
            BTYPE           =   5
            TX              =   "Compute"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   8421504
            BCOLO           =   14737632
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.TextBox tave1 
            Enabled         =   0   'False
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
            Left            =   840
            TabIndex        =   32
            Top             =   1320
            Width           =   495
         End
         Begin VB.TextBox tt1 
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
            Left            =   840
            TabIndex        =   30
            Top             =   840
            Width           =   495
         End
         Begin VB.TextBox tq1 
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
            Left            =   840
            TabIndex        =   28
            Top             =   480
            Width           =   495
         End
         Begin VB.TextBox tcs1 
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
            Left            =   840
            TabIndex        =   26
            Top             =   120
            Width           =   495
         End
         Begin ISAPTECH.chameleonButton CBCom2 
            Height          =   375
            Left            =   -73800
            TabIndex        =   43
            Top             =   1800
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   661
            BOLD            =   0   'False
            ITA             =   0   'False
            INAME           =   "Tahoma"
            SIZE            =   8.25
            UND             =   0   'False
            BTYPE           =   5
            TX              =   "Compute"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   8421504
            BCOLO           =   14737632
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ISAPTECH.chameleonButton CBCom3 
            Height          =   375
            Left            =   -73800
            TabIndex        =   52
            Top             =   1800
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   661
            BOLD            =   0   'False
            ITA             =   0   'False
            INAME           =   "Tahoma"
            SIZE            =   8.25
            UND             =   0   'False
            BTYPE           =   5
            TX              =   "Compute"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   8421504
            BCOLO           =   14737632
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ISAPTECH.chameleonButton CBCom4 
            Height          =   375
            Left            =   -73800
            TabIndex        =   61
            Top             =   1800
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   661
            BOLD            =   0   'False
            ITA             =   0   'False
            INAME           =   "Tahoma"
            SIZE            =   8.25
            UND             =   0   'False
            BTYPE           =   5
            TX              =   "Compute"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   8421504
            BCOLO           =   14737632
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Line Line4 
            X1              =   -73680
            X2              =   -74880
            Y1              =   1200
            Y2              =   1200
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "AVE:"
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
            Left            =   -74880
            TabIndex        =   60
            Top             =   1320
            Width           =   420
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "TT:"
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
            Left            =   -74880
            TabIndex        =   58
            Top             =   840
            Width           =   315
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "TQ:"
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
            Left            =   -74880
            TabIndex        =   56
            Top             =   480
            Width           =   330
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "CS:"
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
            Left            =   -74880
            TabIndex        =   54
            Top             =   120
            Width           =   315
         End
         Begin VB.Line Line3 
            X1              =   -73680
            X2              =   -74880
            Y1              =   1200
            Y2              =   1200
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "AVE:"
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
            Left            =   -74880
            TabIndex        =   51
            Top             =   1320
            Width           =   420
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "TT:"
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
            Left            =   -74880
            TabIndex        =   49
            Top             =   840
            Width           =   315
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "TQ:"
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
            Left            =   -74880
            TabIndex        =   47
            Top             =   480
            Width           =   330
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "CS:"
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
            Left            =   -74880
            TabIndex        =   45
            Top             =   120
            Width           =   315
         End
         Begin VB.Line Line2 
            X1              =   -73680
            X2              =   -74880
            Y1              =   1200
            Y2              =   1200
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "AVE:"
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
            Left            =   -74880
            TabIndex        =   42
            Top             =   1320
            Width           =   420
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "TT:"
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
            Left            =   -74880
            TabIndex        =   40
            Top             =   840
            Width           =   315
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "TQ:"
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
            Left            =   -74880
            TabIndex        =   38
            Top             =   480
            Width           =   330
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "CS:"
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
            Left            =   -74880
            TabIndex        =   36
            Top             =   120
            Width           =   315
         End
         Begin VB.Line Line1 
            X1              =   1320
            X2              =   120
            Y1              =   1200
            Y2              =   1200
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "AVE:"
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
            TabIndex        =   33
            Top             =   1320
            Width           =   420
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "TT:"
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
            TabIndex        =   31
            Top             =   840
            Width           =   315
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "TQ:"
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
            TabIndex        =   29
            Top             =   480
            Width           =   330
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "CS:"
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
            TabIndex        =   27
            Top             =   120
            Width           =   315
         End
      End
      Begin VB.Frame FRM5 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   120
         TabIndex        =   62
         Top             =   3600
         Width           =   2415
         Begin VB.ComboBox CREM 
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
            ItemData        =   "FrmGRD.frx":5872
            Left            =   1200
            List            =   "FrmGRD.frx":5885
            TabIndex        =   66
            Top             =   600
            Width           =   855
         End
         Begin VB.TextBox TRE 
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
            Left            =   1200
            TabIndex        =   65
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "REMARKS:"
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
            Left            =   240
            TabIndex        =   64
            Top             =   600
            Width           =   915
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "RE-EXAM:"
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
            Left            =   240
            TabIndex        =   63
            Top             =   240
            Width           =   870
         End
      End
      Begin VB.Label LBLUNITS 
         AutoSize        =   -1  'True
         Caption         =   "TOTAL UNITS: 0"
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
         TabIndex        =   85
         Top             =   6120
         Width           =   1425
      End
   End
   Begin VB.Frame FRM1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.Label LblSex 
         AutoSize        =   -1  'True
         Caption         =   "SEX"
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
         Left            =   3840
         TabIndex        =   86
         Top             =   120
         Width           =   345
      End
      Begin VB.Label LBLNAME 
         AutoSize        =   -1  'True
         Caption         =   "STUDENT NAME"
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
         Top             =   360
         Width           =   1380
      End
      Begin VB.Label LBLID 
         AutoSize        =   -1  'True
         Caption         =   "I.D. Number"
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
         TabIndex        =   1
         Top             =   120
         Width           =   1035
      End
   End
   Begin VB.Frame FRM2 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   4335
      Begin VB.ComboBox LblSchool 
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
         ItemData        =   "FrmGRD.frx":58A5
         Left            =   1440
         List            =   "FrmGRD.frx":58AF
         TabIndex        =   83
         Text            =   "INTERNATIONAL SCHOOL OF ASIA AND THE PACIFIC"
         Top             =   1680
         Width           =   2775
      End
      Begin VB.TextBox TMjr 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         TabIndex        =   12
         Top             =   1320
         Width           =   1935
      End
      Begin VB.ComboBox CYR 
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
         ItemData        =   "FrmGRD.frx":590B
         Left            =   2760
         List            =   "FrmGRD.frx":591E
         TabIndex        =   11
         Top             =   960
         Width           =   615
      End
      Begin VB.ComboBox CCOURSE 
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
         Left            =   1440
         TabIndex        =   10
         Top             =   960
         Width           =   1335
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
         ItemData        =   "FrmGRD.frx":5931
         Left            =   1440
         List            =   "FrmGRD.frx":593E
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   600
         Width           =   735
      End
      Begin VB.ComboBox CSYR 
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
         Left            =   1440
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label29 
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
         TabIndex        =   84
         Top             =   1680
         Width           =   795
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "MAJOR:"
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
         TabIndex        =   7
         Top             =   1320
         Width           =   675
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "COURSE/YR:"
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
         TabIndex        =   6
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "SEMESTER:"
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
         TabIndex        =   5
         Top             =   600
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "SCHOOL YEAR:"
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
         Top             =   240
         Width           =   1305
      End
   End
   Begin VB.Frame FRM3 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   120
      TabIndex        =   13
      Top             =   2880
      Width           =   4335
      Begin VB.ComboBox CSECTION 
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
         Left            =   1200
         TabIndex        =   78
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox TSCHED 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   23
         Top             =   2760
         Width           =   3855
      End
      Begin VB.ComboBox Cteacher 
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
         Left            =   1200
         TabIndex        =   22
         Top             =   2040
         Width           =   2895
      End
      Begin VB.ComboBox CSubjects 
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
         Left            =   1200
         TabIndex        =   21
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox TSD 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   20
         Top             =   1200
         Width           =   3855
      End
      Begin VB.TextBox Tunits 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3600
         TabIndex        =   19
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "SECTION:"
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
         Left            =   240
         TabIndex        =   79
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "SCHEDULE:"
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
         Left            =   240
         TabIndex        =   18
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "TEACHER:"
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
         Left            =   240
         TabIndex        =   17
         Top             =   2040
         Width           =   885
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "UNITS:"
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
         Left            =   2880
         TabIndex        =   16
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label6 
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
         Left            =   240
         TabIndex        =   15
         Top             =   960
         Width           =   1260
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "SUBJECT:"
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
         Left            =   240
         TabIndex        =   14
         Top             =   600
         Width           =   840
      End
   End
End
Attribute VB_Name = "FrmGRD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ISadd As Boolean
Dim RecSY As ADODB.Recordset
Dim RecCourse As ADODB.Recordset
Dim RecSched As ADODB.Recordset
Dim RecRec As ADODB.Recordset
Dim ConstRect As ADODB.Recordset
Dim RecUnits As ADODB.Recordset
Dim UseSched As Boolean, Adding As Boolean
Dim FirstLoad As Boolean
Dim Prev_Sub As String
Dim Prev_Sec As String

Sub Get_SYPresent() 'Onload
Set RecSY = New ADODB.Recordset
Dim msg As String, i As Long
msg = "Select IDNO, STUDENT, SCHOOLYEAR, SEMESTER,COURSE,YEARLEVEL FROM grading_Sys "
msg = msg & " WHERE IDNO = '" & FrmInfoCNTR.LDVIEW.SelectedItem.Text
msg = msg & "' GROUP By IDNO,STUDENT, SCHOOLYEAR, SEMESTER,COURSE,YEARLEVEL "
msg = msg & "Order by SCHOOLYEAR"
With RecSY
    .ActiveConnection = FrmInfoCNTR.ConX
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open msg
    CSYR.Clear
    LSYMove.Clear
    Do Until .EOF
        For i = 0 To CSYR.ListCount - 1
        CSYR.ListIndex = i
        If .Fields("Schoolyear").Value = CSYR.Text Then
        GoTo NexRec
        End If
        Next
        CSYR.AddItem .Fields(2).Value
        LSYMove.AddItem .Fields(2).Value
NexRec:
        .MoveNext
    Loop
    .Close
End With
    Set RecSY = Nothing
End Sub

Sub Get_Courses()   'Onload
Set RecCourse = Nothing
Set RecCourse = New ADODB.Recordset
Dim msg As String
msg = "Select * From Courses"
With RecCourse
    .ActiveConnection = FrmInfoCNTR.ConX
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open msg
    CCOURSE.Clear
    Do Until .EOF
        CCOURSE.AddItem .Fields(0).Value
        .MoveNext
    Loop
End With
End Sub

Sub SectionList()   'Onload
Set RecSched = Nothing
Set RecSched = New ADODB.Recordset
Dim msg As String
msg = "Select * From Classes "
msg = msg & "Where Class='" & CCOURSE.Text & "' and year_level='" & CYR.Text & "' order by Class,year_level"
With RecSched
    .ActiveConnection = FrmInfoCNTR.ConX
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open msg
    
    CSECTION.Clear
    If .RecordCount = 0 Then Exit Sub
    
    Dim i As String
    Do Until .EOF
        CSECTION.AddItem .Fields("Section").Value
        .MoveNext
    Loop
End With
Set RecSched = Nothing
End Sub

Sub SubjectList()   'Upon Leave of Sections
Set RecSched = New ADODB.Recordset
Dim msg As String
msg = "Select Subject,sectionko From Scheduling Where Sectionko like '%"
msg = msg & CSECTION.Text & "%' Group by Subject,Sectionko order by Subject"
With RecSched
    .ActiveConnection = FrmInfoCNTR.ConX
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open msg
    CSubjects.Clear
    Do Until .EOF
        CSubjects.AddItem .Fields("Subject").Value
        .MoveNext
    Loop
End With
Set RecSched = Nothing
End Sub

Sub SubjectDetails()    'Upon Leave in Subjects
'Generate Details before the schedule
Set RecRec = New ADODB.Recordset
Dim msg As String
msg = "Select * From Scheduling Where Subject = '"
msg = msg & CSubjects.Text & "' order by Scheduling.date"
With RecRec
    .ActiveConnection = FrmInfoCNTR.ConX
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open msg
    If .RecordCount = 0 Then Exit Sub
    Tunits.Text = .Fields("UNITS").Value
    Cteacher.Text = .Fields("teacher").Value
    TSD.Text = .Fields("SUBJECT_DESCRIPTION").Value
    TSCHED.Text = ""
    GetSched
End With
Set RecRec = Nothing
End Sub

Sub GetFirst()

Set RecSched = New ADODB.Recordset
Dim msg As String, i As Long, P1 As String, _
    P2 As String, P3 As String, P4 As String, _
    NL As String, Adx As String, strx As String
msg = "Select * From Scheduling Where Subject='"
msg = msg & CSubjects.Text & "' and Sectionko like '%"
msg = msg & CSECTION.Text & "%' order by Time_In"
With RecSched
    .ActiveConnection = FrmInfoCNTR.ConX
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open msg
    GetSched
End With
End Sub
Sub GetSched()

Set RecSched = New ADODB.Recordset
Dim msg As String, i As Long, P1 As String, _
    P2 As String, P3 As String, P4 As String, _
    NL As String, Adx As String, strx As String
msg = "Select * From Scheduling Where Subject='"
msg = msg & CSubjects.Text & "' and Sectionko like '%"
msg = msg & CSECTION.Text & "%'  order by Subject"
With RecSched
    .ActiveConnection = FrmInfoCNTR.ConX
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open msg
    If .RecordCount = 0 Then Exit Sub   'No Schedule
    
    NL = Left(.Fields("DATE").Value, 1)
        If NL = "T" Then
            If Left(.Fields("DATE").Value, 2) = "TH" Then
                NL = "TH"
            End If
        End If
    P1 = .Fields("Time_In").Value
    P2 = .Fields("Time_Out").Value
    P3 = NL
    Dim x As String
    Do Until .EOF
    If P1 = .Fields("Time_In").Value And P2 = .Fields("Time_Out").Value Then

        NL = Left(.Fields("DATE").Value, 1)
        If NL = "T" Then
            If Left(.Fields("DATE").Value, 2) = "TH" Then
                NL = "TH"
            End If
        End If
        If NL <> P3 Then
            P3 = NL
        End If
        'If Left(Adx, 1) = "," Or Left(Adx, 1) = " " Then
        '    Adx = Right(Adx, Len(Adx) - 1)
        'End If
        'If Adx = "" Then
        '    Adx = P3
        'Else
            Adx = Adx & P3
        'End If
    Else
        TSCHED.Text = TSCHED.Text & Adx & " " & P1 & " - " & P2 & " "
        NL = Left(.Fields("DATE").Value, 1)
        If NL = "T" Then
            If Left(.Fields("DATE").Value, 2) = "TH" Then
                NL = "TH"
            End If
        End If
        Adx = NL
    End If
    
    P3 = Left(.Fields("DATE").Value, 1)
    If P3 = "T" Then
        If Left(.Fields("DATE").Value, 2) = "TH" Then
            P3 = "TH"
        End If
    End If
        P1 = .Fields("Time_In").Value
        P2 = .Fields("Time_Out").Value
        
        .MoveNext
        If .EOF Then    'add last file
        TSCHED.Text = TSCHED.Text & Adx & " " & P1 & " - " & P2 & " "
        End If
    Loop
    If Right(TSCHED.Text, 1) = " " Then
        TSCHED.Text = Left(TSCHED.Text, Len(TSCHED.Text) - 1)
    End If
End With
Set RecSched = Nothing
End Sub

Private Sub CBADD_Click()
    Adding = True
    DisAble
    ClearAll
    CSECTION.SetFocus
End Sub

Sub LoadRecords()
Dim p(6) As String
'Get Previous Values
If FirstLoad = False Then        'Firstload, load the first record to count in
p(0) = CSYR.Text
p(1) = CSEM.Text
p(2) = CSECTION.Text
p(3) = CSubjects.Text
p(4) = Cteacher.Text
p(5) = TSCHED.Text
End If
Set ConstRect = Nothing
Set ConstRect = New ADODB.Recordset
Dim msg As String
msg = SqlStGrd
With ConstRect
    .ActiveConnection = FrmInfoCNTR.ConX
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open msg
    .Properties.Refresh
    LBLREC.Caption = .AbsolutePosition & " of " & .RecordCount
    If .RecordCount = 0 Then     'Clear All
    ClearAll
    'DisAble Some
    LBLREC.Caption = "No Records"
    CBDel.Enabled = 0
    CBupdate.Enabled = 0
    CBSearch.Enabled = 0
    Exit Sub
    Else
    Enable
    End If
    Me.Get_SYPresent
    Enable
    Popul
    'After Populating Check if First Load
    If FirstLoad = True Then
    p(0) = CSYR.Text
    p(1) = CSEM.Text
    p(2) = CSECTION.Text
    p(3) = CSubjects.Text
    p(4) = Cteacher.Text
    p(5) = TSCHED.Text
    FirstLoad = False
    End If
    Call Count_Subject(p(0), p(1), p(4), p(5), p(3))
    Call Count_Class(p(0), p(1), p(2))
    Call TotalUnits
End With
End Sub
Sub Popul()

With ConstRect
'If .RecordCount = 0 Then Exit Sub
CSYR.Text = .Fields("SchoolYear").Value
CSEM.Text = .Fields("Semester").Value
CCOURSE.Text = .Fields("Course").Value
CYR.Text = .Fields("yearlevel").Value

IsN TMjr, .Fields("MAJOR").Value
IsN LblSchool, .Fields("SCHOOL").Value
IsN CSECTION, .Fields("SECTION").Value
IsN CSubjects, .Fields("Subject").Value
IsN Tunits, .Fields("UNITS").Value
IsN TSD, .Fields("SUBJECT_DESCRIPTION").Value
IsN Cteacher, .Fields("Teacher").Value
IsN TSCHED, .Fields("SCHEDULE").Value
IsN TRE, .Fields("REEXAM").Value
IsN CREM, .Fields("REMARKS").Value
IsN LblSex, .Fields("Sex").Value

IsN tcs1, .Fields("P1").Value
IsN tq1, .Fields("p2").Value
IsN tt1, .Fields("p3").Value
IsN tave1, .Fields("Prelim").Value
    
IsN TCS2, .Fields("m1").Value
IsN TQ2, .Fields("m2").Value
IsN TT2, .Fields("m3").Value
IsN TAVE2, .Fields("midterm").Value

IsN TCS3, .Fields("s1").Value
IsN TQ3, .Fields("s2").Value
IsN TT3, .Fields("s3").Value
IsN TAVE3, .Fields("Semi").Value

IsN TCS4, .Fields("f1").Value
IsN TQ4, .Fields("f2").Value
IsN TT4, .Fields("f3").Value
IsN TAVE4, .Fields("Finals").Value

End With
End Sub
Function IsN(OBJ As Object, Value As Variant)
If IsNull(Value) = True Then
OBJ.Text = ""
Else
If OBJ.Name = "LBLID" Or _
    OBJ.Name = "LBLNAME" Or OBJ.Name = "LblSex" Then
OBJ.Caption = Value
Else
OBJ.Text = Value
End If
End If
End Function

Sub ClearAll()
Dim i As Long
CSubjects.Clear

Tunits.Text = ""
TSD.Text = ""
Cteacher.Text = ""
TSCHED.Text = ""
For i = 1 To 4
    Me.Controls("Tcs" & i).Text = ""
    Me.Controls("Tq" & i).Text = ""
    Me.Controls("Tt" & i).Text = ""
    Me.Controls("Tave" & i).Text = ""
Next
CREM.Text = ""
TRE.Text = ""

End Sub

Sub DisAble()   'Adding
CBADD.Enabled = 0
CBupdate.Enabled = 1
CBDel.Caption = "Cancel"
CBDel.Enabled = 1
CBSearch.Enabled = 0
CBPrint.Enabled = 0
CBF.Enabled = 0
CBF.Enabled = 0
CBPR.Enabled = 0
CBNX.Enabled = 0
CBL.Enabled = 0
CBMOVE.Enabled = 0
Me.LBLREC.Caption = "Adding..."
End Sub

Sub Enable()    'Update
CBADD.Enabled = 1
CBDel.Caption = "Delete Subject"
CBSearch.Enabled = 1
CBupdate.Enabled = 1
CBDel.Enabled = 1
CBPrint.Enabled = 1
CBF.Enabled = 1
CBF.Enabled = 1
CBPR.Enabled = 1
CBNX.Enabled = 1
CBL.Enabled = 1
CBMOVE.Enabled = 1
End Sub

Private Sub CBCom1_Click()
On Error GoTo ErrorX
Dim CS As Double, AQ As Double, TT As Double, AVE As Double, Sums As Double
    CS = tcs1.Text
    AQ = tq1.Text
    TT = tt1.Text
    AVE = (CS + AQ + TT) / 3
    AVE = Format(AVE, "##.##")
    tave1.Text = AVE
Exit Sub
ErrorX:
    MsgBox Err.Description, vbInformation, "Error: " & Err.Number
    Me.Refresh

End Sub

Private Sub CBCom2_Click()
On Error GoTo ErrorX
Dim CS As Double, AQ As Double, TT As Double, AVE As Double, Sums As Double
    CS = TCS2.Text
    AQ = TQ2.Text
    TT = TT2.Text
    AVE = (CS + AQ + TT) / 3
    Sums = ((AVE * 2) + tave1.Text) / 3
    Sums = Format(Sums, "##.##")
    TAVE2.Text = Sums
Exit Sub
ErrorX:
    MsgBox Err.Description, vbInformation, "Error: " & Err.Number
    Me.Refresh
End Sub

Private Sub CBCom3_Click()
On Error GoTo ErrorX
Dim CS As Double, AQ As Double, TT As Double, AVE As Double, Sums As Double
    CS = TCS3.Text
    AQ = TQ3.Text
    TT = TT3.Text
    AVE = (CS + AQ + TT) / 3
    Sums = ((AVE * 2) + TAVE2.Text) / 3
    Sums = Format(Sums, "##.##")
    TAVE3.Text = Sums
Exit Sub
ErrorX:
    MsgBox Err.Description, vbInformation, "Error: " & Err.Number
    Me.Refresh
End Sub

Private Sub CBCom4_Click()
On Error GoTo ErrorX
Dim CS As Double, AQ As Double, TT As Double, AVE As Double, Sums As Double
    CS = TCS4.Text
    AQ = TQ4.Text
    TT = TT4.Text
    AVE = (CS + AQ + TT) / 3
    Sums = ((AVE * 2) + TAVE3.Text) / 3
    Sums = Format(Sums, "##.##")
    TAVE4.Text = Sums
    CREM.Text = Format(Sums, "##")
Exit Sub
ErrorX:
    MsgBox Err.Description, vbInformation, "Error: " & Err.Number
    Me.Refresh
End Sub

Private Sub CBDel_Click()
On Error GoTo Erb
Select Case CBDel.Caption
    Case "Delete Subject"   'Delete Current Subject
    Dim msg As String
    With ConstRect
        If MsgBox("Are you sure you want to delete this record?", vbQuestion + vbYesNo, "DELETE") = vbNo Then Exit Sub
        msg = "Delete From Grading_Sys where "
        msg = msg & " IDNO = '" & LBLID.Caption
        msg = msg & "' and Schoolyear='" & .Fields("SCHOOLYEAR").Value
        msg = msg & "' and Semester='" & .Fields("Semester").Value
        msg = msg & "' and COURSE='" & .Fields("Course").Value
        msg = msg & "' and YEARLEVEL='" & .Fields("YEARLEVEL").Value
        msg = msg & "' and SUBJECT='" & .Fields("Subject").Value
        msg = msg & "'"
        If .State <> 0 Then
            .Close
        End If
            .CursorLocation = adUseClient
            .CursorType = adOpenDynamic
            .LockType = adLockOptimistic
            .Open msg
            LoadRecords
            MsgBox "Record Deleted, and can never be retreived.", vbInformation, "Delete Complete"
    End With
    Case "Cancel"           'Adding Active
    Enable
    Adding = False
    LoadRecords
End Select
Exit Sub
Erb:
    ErrorTrap Err, "Delete/Cancel Command"
    Me.Refresh
    LoadRecords
End Sub

Private Sub CBF_Click()
With ConstRect
    .MoveFirst
    Popul
    Me.LBLREC.Caption = .AbsolutePosition & " of " & .RecordCount
End With
End Sub

Private Sub CBL_Click()
With ConstRect
    .MoveLast
    Popul
    Me.LBLREC.Caption = .AbsolutePosition & " of " & .RecordCount
End With
End Sub

Private Sub CBMOVE_Click()
If CMovSem.Text = "" Or LSYMove.Text = "" Then Exit Sub
Dim x As String
x = LSYMove.Text
SqlStGrd = "Select * From Grading_Sys Where IDNO ='"
SqlStGrd = SqlStGrd & LBLID.Caption & "'"
SqlStGrd = SqlStGrd & " and Schoolyear = '" & Me.LSYMove.Text
SqlStGrd = SqlStGrd & "' and Semester = '" & CMovSem.Text & "' order by Subject"
LoadRecords
CSYR.Text = x
CSEM.Text = CMovSem.Text
End Sub

Private Sub CBNX_Click()
With ConstRect
    .MoveNext
    If .EOF Then .MoveLast
    Popul
    Me.LBLREC.Caption = .AbsolutePosition & " of " & .RecordCount
End With
End Sub

Private Sub CBPR_Click()
With ConstRect
    .MovePrevious
    If .BOF Then .MoveFirst
    Popul
    Me.LBLREC.Caption = .AbsolutePosition & " of " & .RecordCount
End With
End Sub

Private Sub CBPrint_Click()
'Print SCR
Dim msg As String
msg = "SHAPE {SELECT IDNO,STUDENT,SCHOOLYEAR,SEMESTER,"
msg = msg & "SCHOOL,COURSE,YEARLEVEL,MAJOR FROM "
msg = msg & "GRADING_SYS WHERE IDNO='" & Me.LBLID.Caption & "' GROUP BY SCHOOLYEAR,SEMESTER,"
msg = msg & "IDNO,STUDENT ,SCHOOL,COURSE,YEARLEVEL,MAJOR}"
msg = msg & "AS HeadGrades APPEND ({SELECT * FROM GRADING_SYS where IDNO='" & LBLID.Caption & "'}"
msg = msg & " AS SecHeadGrades RELATE 'IDNO' TO 'IDNO','STUDENT' TO "
msg = msg & "'STUDENT','SCHOOLYEAR' TO 'SCHOOLYEAR','SEMESTER' "
msg = msg & "TO 'SEMESTER','SCHOOL' TO 'SCHOOL','COURSE' TO "
msg = msg & "'COURSE','YEARLEVEL' TO 'YEARLEVEL') AS SecHeadGrades"
SetDenver msg, 5
With DR6
    Dim SCHL As String
    SCHL = InputBox("Enter School", "School")
    .Sections("pageHeader").Controls("LblSchool").Caption = Trim(UCase(SCHL))
    .Refresh
    .Show 1
End With
End Sub

Private Sub CBSearch_Click()
'Search Subject
Dim Itm As String, strx As String
Itm = InputBox("Enter Subject (Part or Whole Word):", "Search")
If Trim(Itm) = "" Then Exit Sub
With ConstRect
    Itm = Trim(Itm)
    'STRX = "Schoolyear = '" & .Fields("Schoolyear").Value
    'STRX = STRX & "' and Semester='" & .Fields("Semester").Value
    strx = "Subject like '" & Itm & "%'"
    .Find strx, , adSearchForward, 1
    If .EOF Then
        MsgBox "Record not found.", vbInformation, "Search"
        .MoveFirst
        LBLREC.Caption = .AbsolutePosition & " of " & .RecordCount
        Exit Sub
    End If
    'Load the value
    Popul
    LBLREC.Caption = .AbsolutePosition & " of " & .RecordCount
End With
End Sub

Private Sub CBupdate_Click()
On Error GoTo ErrorTrapx
Dim msg As String
With ConstRect
Select Case Adding
    Case True   'Adding Record
        .AddNew
        msg = "Subject Added to " & LBLNAME.Caption & "."
    Case False
        'Just Update
        If MsgBox("Update This Record?", vbQuestion + vbYesNo, "Update") = vbNo Then
        .Properties.Refresh
        Me.LoadRecords
        Enable
        Exit Sub
        End If
        msg = "Subject updated to " & LBLNAME.Caption & "."
End Select
SeTValues
.Update
.Properties.Refresh
SqlStGrd = "Select * From Grading_Sys Where IDNO ='"
SqlStGrd = SqlStGrd & LBLID.Caption & "'"
SqlStGrd = SqlStGrd & " and Schoolyear = '" & CSYR.Text & "' and School = '" & LblSchool.Text
SqlStGrd = SqlStGrd & "' and Semester = '" & CSEM.Text & "' order by Subject"
LoadRecords
Adding = False
Enable
MsgBox msg, vbInformation, "Add Edit Command"
End With
Exit Sub
ErrorTrapx:
    ErrorTrap Err, "Add Edit Command"
    Me.Refresh
    LoadRecords
    Enable
    Adding = False
End Sub

Sub SeTValues()
With ConstRect
.Fields("IDNO").Value = LBLID.Caption
.Fields("Student").Value = LBLNAME.Caption

.Fields("SchoolYear").Value = CSYR.Text
.Fields("Semester").Value = CSEM.Text
.Fields("Course").Value = CCOURSE.Text
.Fields("yearlevel").Value = CYR.Text
.Fields("Sex").Value = LblSex.Caption

.Fields("MAJOR").Value = TMjr.Text
.Fields("SCHOOL").Value = LblSchool.Text
.Fields("SECTION").Value = CSECTION.Text
.Fields("Subject").Value = CSubjects.Text
.Fields("UNITS").Value = Tunits.Text
.Fields("SUBJECT_DESCRIPTION").Value = TSD.Text
.Fields("Teacher").Value = Cteacher.Text
.Fields("SCHEDULE").Value = TSCHED.Text
.Fields("REEXAM").Value = TRE.Text
.Fields("REMARKS").Value = CREM.Text

.Fields("P1").Value = Val(tcs1.Text)
.Fields("p2").Value = Val(tq1.Text)
.Fields("p3").Value = Val(tt1.Text)
.Fields("Prelim").Value = Val(tave1.Text)
    
 .Fields("m1").Value = Val(TCS2.Text)
.Fields("m2").Value = Val(TQ2.Text)
.Fields("m3").Value = Val(TT2.Text)
.Fields("midterm").Value = Val(TAVE2.Text)

.Fields("s1").Value = Val(TCS3.Text)
.Fields("s2").Value = Val(TQ3.Text)
.Fields("s3").Value = Val(TT3.Text)
.Fields("Semi").Value = Val(TAVE3.Text)

.Fields("f1").Value = Val(TCS4.Text)
.Fields("f2").Value = Val(TQ4.Text)
.Fields("f3").Value = Val(TT4.Text)
.Fields("Finals").Value = Val(TAVE4.Text)

End With

End Sub
Private Sub CCOURSE_LostFocus()
With RecCourse
    .Find "course = '" & CCOURSE.Text & "'", , adSearchForward, 1
    If .EOF Then
    MsgBox "Course not found.", vbCritical, "Error"
    CBDel_Click
    Exit Sub
    End If
    LblSchool.Text = .Fields("School").Value
End With
End Sub


Private Sub CSECTION_GotFocus()
Dim XPR As String
XPR = CSECTION.Text
If UseSched = True Then SectionList
CSECTION.Text = XPR
Prev_Sec = XPR
End Sub

Private Sub CSECTION_LostFocus()
Dim i As Integer
If Prev_Sec <> CSECTION.Text Then
    i = MONMODE.RetClass(RecSched, CSYR.Text, CSEM.Text, CSECTION.Text)
    If MONMODE.PermitClassAdd(CSYR.Text, CSEM.Text, i) = False Then
        If Adding = True Then
            CBDel_Click
        Else
            CSECTION.Text = Prev_Sec
        End If
    End If
End If
End Sub

Private Sub CSubjects_GotFocus()
Dim XPR As String
XPR = CSubjects.Text
If UseSched = True Then SubjectList
CSubjects.Text = XPR
Prev_Sub = XPR
End Sub

Private Sub CSubjects_lostfocus()
Dim i As Integer
If UseSched = True Then SubjectDetails
If CSubjects.Text <> Prev_Sub Then
i = MONMODE.RetTotInSubs(FrmMon.RSSubjects, CSYR.Text, CSEM.Text, Cteacher.Text, TSCHED.Text, CSubjects.Text)
' If false Something is wrong, Get def if not adding else cancel

If MONMODE.GetTotinOthers(CSYR.Text, CSEM.Text, i) = False Then
    If Adding = True Then   'Cancel
        CBDel_Click
    Else
        CSubjects.Text = Prev_Sub
    End If
End If
End If

End Sub

Private Sub Form_Load()
FirstLoad = True
Adding = False
FrmMon.Show
Me.Get_Courses
Me.Get_SYPresent
'Me.SectionList
Call LoadRecords
With FrmInfoCNTR
CSYR.Text = .CSY.Text
CSEM.Text = .CSEM.Text
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload FrmMon
FrmSet.CBINFOR.Enabled = True
FrmInfoCNTR.Show
End Sub

Private Sub Timer1_Timer()
If MsgBox("Do you want to use the Scheduling Patch?", vbQuestion + vbYesNo, "Use Patch") = vbYes Then UseSched = True Else UseSched = False
Timer1.Enabled = False
End Sub

Private Sub TotalUnits()
    Dim msg As String
    msg = "Select SUM(Units) as TotalUnits from Grading_SYS where Schoolyear='"
    msg = msg & CSYR.Text & "' and Semester='" & CSEM.Text & "' and IDNO='"
    msg = msg & FrmInfoCNTR.LDVIEW.SelectedItem.Text & "'"
    Set RecUnits = Nothing
    Set RecUnits = New ADODB.Recordset
    With RecUnits
        .ActiveConnection = FrmInfoCNTR.ConX
        .CursorLocation = adUseClient
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
        .Open msg
        LBLUnits.Caption = "Total Units: " & .Fields(0).Value
        'RecUnits.Close
    End With
    Set RecUnits = Nothing
End Sub

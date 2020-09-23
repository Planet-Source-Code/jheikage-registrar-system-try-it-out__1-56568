VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FrmSPI 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Personal Information"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   7575
   Icon            =   "FrmSPI.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   7575
   StartUpPosition =   2  'CenterScreen
   Begin ISAPTECH.chameleonButton CBOK 
      Height          =   495
      Left            =   4920
      TabIndex        =   38
      Top             =   3720
      Width           =   1215
      _extentx        =   2143
      _extenty        =   873
      und             =   0
      iname           =   "Tahoma"
      btype           =   14
      tx              =   "OK"
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
   Begin ISAPTECH.chameleonButton CBCANCEL 
      Height          =   495
      Left            =   6240
      TabIndex        =   39
      Top             =   3720
      Width           =   1215
      _extentx        =   2143
      _extenty        =   873
      und             =   0
      iname           =   "Tahoma"
      btype           =   14
      tx              =   "CANCEL"
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
   Begin TabDlg.SSTab PerTabs 
      Height          =   3495
      Left            =   120
      TabIndex        =   37
      Top             =   120
      Width           =   7365
      _ExtentX        =   12991
      _ExtentY        =   6165
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Personal Profile"
      TabPicture(0)   =   "FrmSPI.frx":57E2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label6"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label7"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label8"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label9"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label10"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label11"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "TID"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "TNAM"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Tadd"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "TAge"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "LSex"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "TCS"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "TNAT"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "TBP"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "TCAdd"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "THW"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "MBDAY"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).ControlCount=   22
      TabCaption(1)   =   "Family Background"
      TabPicture(1)   =   "FrmSPI.frx":57FE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label12"
      Tab(1).Control(1)=   "Label13"
      Tab(1).Control(2)=   "Label14"
      Tab(1).Control(3)=   "Label15"
      Tab(1).Control(4)=   "Label16"
      Tab(1).Control(5)=   "Label17"
      Tab(1).Control(6)=   "Label18"
      Tab(1).Control(7)=   "Label19"
      Tab(1).Control(8)=   "Label20"
      Tab(1).Control(9)=   "Label38"
      Tab(1).Control(10)=   "Label39"
      Tab(1).Control(11)=   "TPADD"
      Tab(1).Control(12)=   "TF"
      Tab(1).Control(13)=   "TM"
      Tab(1).Control(14)=   "TOF"
      Tab(1).Control(15)=   "TOM"
      Tab(1).Control(16)=   "TG"
      Tab(1).Control(17)=   "TOG"
      Tab(1).Control(18)=   "TADDG"
      Tab(1).Control(19)=   "TEM"
      Tab(1).Control(20)=   "TPPH"
      Tab(1).Control(21)=   "TGPH"
      Tab(1).ControlCount=   22
      TabCaption(2)   =   "School/s Attended"
      TabPicture(2)   =   "FrmSPI.frx":581A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label21"
      Tab(2).Control(1)=   "Label22"
      Tab(2).Control(2)=   "Label23"
      Tab(2).Control(3)=   "Label24"
      Tab(2).Control(4)=   "Label25"
      Tab(2).Control(5)=   "Label26"
      Tab(2).Control(6)=   "Label27"
      Tab(2).Control(7)=   "Label28"
      Tab(2).Control(8)=   "Label30"
      Tab(2).Control(9)=   "TPR"
      Tab(2).Control(10)=   "TPY"
      Tab(2).Control(11)=   "TIN"
      Tab(2).Control(12)=   "TIY"
      Tab(2).Control(13)=   "THS"
      Tab(2).Control(14)=   "THY"
      Tab(2).Control(15)=   "TLCA"
      Tab(2).Control(16)=   "TCY"
      Tab(2).Control(17)=   "TLY"
      Tab(2).ControlCount=   18
      TabCaption(3)   =   "Credentials Presented"
      TabPicture(3)   =   "FrmSPI.frx":5836
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label29"
      Tab(3).Control(1)=   "Label31"
      Tab(3).Control(2)=   "Label32"
      Tab(3).Control(3)=   "Label33"
      Tab(3).Control(4)=   "Label34"
      Tab(3).Control(5)=   "Label35"
      Tab(3).Control(6)=   "Label36"
      Tab(3).Control(7)=   "LF"
      Tab(3).Control(8)=   "LN"
      Tab(3).Control(9)=   "LT"
      Tab(3).Control(10)=   "LD"
      Tab(3).Control(11)=   "LHD"
      Tab(3).Control(12)=   "LPF"
      Tab(3).ControlCount=   13
      Begin MSMask.MaskEdBox MBDAY 
         Height          =   375
         Left            =   1320
         TabIndex        =   7
         Top             =   1920
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "/"
      End
      Begin VB.TextBox TGPH 
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
         Left            =   -69120
         TabIndex        =   20
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox TPPH 
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
         Left            =   -69120
         TabIndex        =   16
         Top             =   1440
         Width           =   1335
      End
      Begin VB.ListBox LPF 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "FrmSPI.frx":5852
         Left            =   -69720
         List            =   "FrmSPI.frx":585C
         TabIndex        =   36
         Top             =   1440
         Width           =   855
      End
      Begin VB.ListBox LHD 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "FrmSPI.frx":5869
         Left            =   -69720
         List            =   "FrmSPI.frx":5873
         TabIndex        =   35
         Top             =   960
         Width           =   855
      End
      Begin VB.ListBox LD 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "FrmSPI.frx":5880
         Left            =   -69720
         List            =   "FrmSPI.frx":588A
         TabIndex        =   34
         Top             =   480
         Width           =   855
      End
      Begin VB.ListBox LT 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "FrmSPI.frx":5897
         Left            =   -73800
         List            =   "FrmSPI.frx":58A1
         TabIndex        =   33
         Top             =   1440
         Width           =   855
      End
      Begin VB.ListBox LN 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "FrmSPI.frx":58AE
         Left            =   -73800
         List            =   "FrmSPI.frx":58B8
         TabIndex        =   32
         Top             =   960
         Width           =   855
      End
      Begin VB.ListBox LF 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "FrmSPI.frx":58C5
         Left            =   -73800
         List            =   "FrmSPI.frx":58CF
         TabIndex        =   31
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox TLY 
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
         Left            =   -69120
         TabIndex        =   30
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox TCY 
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
         Left            =   -73440
         TabIndex        =   29
         Top             =   2400
         Width           =   2655
      End
      Begin VB.TextBox TLCA 
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
         Left            =   -72240
         TabIndex        =   28
         Top             =   1920
         Width           =   4455
      End
      Begin VB.TextBox THY 
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
         Left            =   -69120
         TabIndex        =   27
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox THS 
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
         Left            =   -73440
         TabIndex        =   26
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox TIY 
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
         Left            =   -69120
         TabIndex        =   25
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox TIN 
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
         Left            =   -73440
         TabIndex        =   24
         Top             =   960
         Width           =   2655
      End
      Begin VB.TextBox TPY 
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
         Left            =   -69120
         TabIndex        =   23
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox TPR 
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
         Left            =   -73440
         TabIndex        =   22
         Top             =   480
         Width           =   2655
      End
      Begin VB.TextBox TEM 
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
         Left            =   -70800
         TabIndex        =   21
         Top             =   2880
         Width           =   3015
      End
      Begin VB.TextBox TADDG 
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
         Left            =   -73800
         TabIndex        =   19
         Top             =   2400
         Width           =   3735
      End
      Begin VB.TextBox TOG 
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
         Left            =   -70080
         TabIndex        =   18
         Top             =   1920
         Width           =   2295
      End
      Begin VB.TextBox TG 
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
         Left            =   -73800
         TabIndex        =   17
         Top             =   1920
         Width           =   2535
      End
      Begin VB.TextBox TOM 
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
         Left            =   -70080
         TabIndex        =   14
         Top             =   960
         Width           =   2295
      End
      Begin VB.TextBox TOF 
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
         Left            =   -70080
         TabIndex        =   12
         Top             =   480
         Width           =   2295
      End
      Begin VB.TextBox TM 
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
         Left            =   -73800
         TabIndex        =   13
         Top             =   960
         Width           =   2535
      End
      Begin VB.TextBox TF 
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
         Left            =   -73800
         TabIndex        =   11
         Top             =   480
         Width           =   2535
      End
      Begin VB.TextBox THW 
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
         Left            =   3240
         TabIndex        =   10
         Top             =   2880
         Width           =   3975
      End
      Begin VB.TextBox TCAdd 
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
         Left            =   1320
         TabIndex        =   9
         Top             =   2400
         Width           =   5895
      End
      Begin VB.TextBox TBP 
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
         Left            =   4200
         TabIndex        =   8
         Top             =   1920
         Width           =   3015
      End
      Begin VB.TextBox TNAT 
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
         Left            =   6480
         TabIndex        =   6
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox TCS 
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
         Left            =   4440
         TabIndex        =   5
         Top             =   1440
         Width           =   975
      End
      Begin VB.ListBox LSex 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "FrmSPI.frx":58DC
         Left            =   2640
         List            =   "FrmSPI.frx":58E6
         TabIndex        =   4
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox TAge 
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
         Left            =   1320
         TabIndex        =   3
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox Tadd 
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
         Left            =   1320
         TabIndex        =   2
         Top             =   960
         Width           =   5895
      End
      Begin VB.TextBox TNAM 
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
         Left            =   3600
         TabIndex        =   1
         Top             =   480
         Width           =   3615
      End
      Begin VB.TextBox TID 
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
         Left            =   1320
         TabIndex        =   0
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox TPADD 
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
         Left            =   -73800
         TabIndex        =   15
         Top             =   1440
         Width           =   3735
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         Caption         =   "Phone:"
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
         Left            =   -69840
         TabIndex        =   78
         Top             =   2520
         Width           =   600
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         Caption         =   "Phone:"
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
         Left            =   -69840
         TabIndex        =   77
         Top             =   1560
         Width           =   600
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         Caption         =   "Please Select Fill up all parameters."
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
         TabIndex        =   75
         Top             =   3120
         Width           =   3075
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         Caption         =   "Permit From Last School Attended"
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
         Left            =   -72720
         TabIndex        =   74
         Top             =   1560
         Width           =   2925
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         Caption         =   "Honorable Dismissal"
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
         Left            =   -71520
         TabIndex        =   73
         Top             =   1080
         Width           =   1725
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "Diploma"
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
         Left            =   -70560
         TabIndex        =   72
         Top             =   600
         Width           =   690
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "Transcript"
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
         Left            =   -74760
         TabIndex        =   71
         Top             =   1560
         Width           =   870
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "NCEE"
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
         Left            =   -74400
         TabIndex        =   70
         Top             =   1080
         Width           =   450
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "Form 138"
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
         Left            =   -74760
         TabIndex        =   69
         Top             =   600
         Width           =   825
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "Year/s Attended:"
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
         Left            =   -70680
         TabIndex        =   68
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "Course/YR:"
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
         TabIndex        =   67
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "Last College/School Attended:"
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
         TabIndex        =   66
         Top             =   2040
         Width           =   2580
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Year/s Attended:"
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
         Left            =   -70680
         TabIndex        =   65
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "High School:"
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
         TabIndex        =   64
         Top             =   1560
         Width           =   1080
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Year/s Attended:"
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
         Left            =   -70680
         TabIndex        =   63
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Intermediate:"
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
         TabIndex        =   62
         Top             =   1080
         Width           =   1170
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Year/s Attended:"
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
         Left            =   -70680
         TabIndex        =   61
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Primary School:"
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
         Top             =   600
         Width           =   1365
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Name/Address of Employer if Working Student:"
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
         TabIndex        =   59
         Top             =   3000
         Width           =   4065
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Address:"
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
         Top             =   2520
         Width           =   765
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Occupation:"
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
         Left            =   -71160
         TabIndex        =   57
         Top             =   2040
         Width           =   1020
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Guardian:"
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
         Top             =   2040
         Width           =   840
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Address:"
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
         TabIndex        =   55
         Top             =   1560
         Width           =   765
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Occupation:"
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
         Left            =   -71160
         TabIndex        =   54
         Top             =   1080
         Width           =   1020
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Occupation:"
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
         Left            =   -71160
         TabIndex        =   53
         Top             =   600
         Width           =   1020
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Mother:"
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
         TabIndex        =   52
         Top             =   1080
         Width           =   675
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Father:"
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
         Top             =   600
         Width           =   630
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Name of Husband/Wife if Married:"
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
         TabIndex        =   50
         Top             =   3000
         Width           =   2940
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "City Address:"
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
         TabIndex        =   49
         Top             =   2520
         Width           =   1140
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Birth Place:"
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
         Left            =   3120
         TabIndex        =   48
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "*Birth Day:"
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
         TabIndex        =   47
         Top             =   2040
         Width           =   960
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Nationality:"
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
         Left            =   5520
         TabIndex        =   46
         Top             =   1560
         Width           =   960
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Civil Status:"
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
         Left            =   3360
         TabIndex        =   45
         Top             =   1560
         Width           =   1020
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Sex:"
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
         Left            =   2160
         TabIndex        =   44
         Top             =   1560
         Width           =   390
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "*Age:"
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
         TabIndex        =   43
         Top             =   1560
         Width           =   525
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Address:"
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
         TabIndex        =   42
         Top             =   1080
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "*STUDENT:"
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
         Left            =   2640
         TabIndex        =   41
         Top             =   600
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "*ID Number:"
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
         TabIndex        =   40
         Top             =   600
         Width           =   1110
      End
   End
   Begin VB.Label Label37 
      AutoSize        =   -1  'True
      Caption         =   "* - Required Fields. Must contain values."
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
      TabIndex        =   76
      Top             =   3720
      Width           =   3495
   End
End
Attribute VB_Name = "FrmSPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public AddOr As Boolean
Dim SQLSt As String
Dim ErrEn As Boolean
Private Sub cbCancel_Click()
Unload Me
End Sub

Private Sub CBOk_Click()
On Error GoTo ErroX
Dim msg As String
Set FrmInfoCNTR.ConRec = New ADODB.Recordset
With FrmInfoCNTR.ConRec
 If .State <> 0 Then .Close
        .ActiveConnection = FrmInfoCNTR.ConX
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
        .CursorType = adOpenDynamic
        
Select Case AddOr
Case True       'Adding Of Records
    Noentry
    If ErrEn = True Then Exit Sub
    
       .Open "TBA_SPI"
        .AddNew
        FillSqlMsg
        
        msg = "Record Added."
Case False      'Updating of Records
        AddOr = True
        .Open "Select * From TBA_SPI Where IDNO ='" & _
            FrmInfoCNTR.LDVIEW.SelectedItem.Text & "'"
        FillSqlMsg
        msg = "Record Updated."
End Select
        .Update
        .Properties.Refresh
End With
    MsgBox msg, vbInformation, "Add/Edit Command"
    Set FrmInfoCNTR.ConRec = Nothing
    Unload Me
    Exit Sub
ErroX:
    ErrorTrap Err, "Adding/Updating SPI"
    Set FrmInfoCNTR.ConRec = Nothing
    Unload Me
End Sub

Sub Noentry()
If Trim(TID.Text) = "" Or _
    Trim(TNAM.Text) = "" Or _
    IsNumeric(TAge.Text) = False Or _
    IsDate(MBDAY.Text) = False Then
    ErrEn = True    'Novalue
    MsgBox "Can't Continue in the said operation. Check your entries.", vbCritical, "ERROR"
Else
    ErrEn = False
End If
End Sub

Sub FillSqlMsg()
Dim Items(36) As String
Trimer TID, Items(0), 0
Trimer TNAM, Items(1), 1
Trimer Tadd, Items(2), 3
Trimer TAge, Items(3), 6
Trimer LSex, Items(4), 7
Trimer TCS, Items(5), 8
Trimer TNAT, Items(6), 9
Trimer MBDAY, Items(7), 4
Trimer TBP, Items(8), 5
Trimer TCAdd, Items(9), 2
Trimer THW, Items(10), 10

Trimer TF, Items(11), 11
Trimer TOF, Items(12), 12
Trimer TM, Items(13), 13
Trimer TOM, Items(14), 14
Trimer TPADD, Items(15), 15
Trimer TPPH, Items(16), 16
Trimer TG, Items(17), 17
Trimer TOG, Items(18), 18
Trimer Me.TADDG, Items(19), 19
Trimer Me.TGPH, Items(20), 20
Trimer Me.TEM, Items(21), 21

Trimer Me.TPR, Items(22), 22
Trimer Me.TPY, Items(23), 23
Trimer Me.TIN, Items(24), 24
Trimer Me.TIY, Items(25), 25
Trimer Me.THS, Items(26), 26
Trimer Me.THY, Items(27), 27
Trimer Me.TLCA, Items(28), 28
Trimer Me.TCY, Items(29), 29
Trimer Me.TLY, Items(30), 30

Trimer Me.LF, Items(31), 31
Trimer Me.LN, Items(32), 32
Trimer Me.LT, Items(33), 33
Trimer Me.LD, Items(34), 34
Trimer LHD, Items(35), 35
Trimer LPF, Items(36), 36

End Sub

Sub Trimer(ByVal OBJ As Object, STRVAL As String, i As Long)
STRVAL = Trim(OBJ.Text)
Select Case AddOr
Case True
With FrmInfoCNTR.ConRec
    .Fields(i).Value = STRVAL
End With
Case False
With FrmInfoCNTR.ConRec
    If IsNull(.Fields(i).Value) Then
    OBJ.Text = ""
    Else
        If i = 4 Then
        Dim OJ, ix As Long, MX As String
        OJ = Split(.Fields(i).Value, "/", , vbTextCompare)
        For ix = LBound(OJ) To UBound(OJ)
        If Len(OJ(ix)) = 1 Then 'Add Zero
        MX = MX & "0" & OJ(ix) & "/"
        Else
        MX = MX & OJ(ix) & "/"
        End If
        Next
        OBJ.Text = Left(MX, Len(MX) - 1)
        Else
        OBJ.Text = Trim(.Fields(i).Value)
        End If
    End If
End With
End Select
End Sub

Private Sub Form_Load()
Set FrmInfoCNTR.ConRec = New ADODB.Recordset
Select Case AddOr
Case True
'Addnew method

Case False
'Edit Method Load All to the fields
With FrmInfoCNTR.ConRec
        If .State <> 0 Then .Close
        .ActiveConnection = FrmInfoCNTR.ConX
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
        .CursorType = adOpenDynamic
        .Open "SELECT * FROM TBA_SPI WHERE IDNO = '" & FrmInfoCNTR.LDVIEW.SelectedItem.Text & "'"
End With
FillSqlMsg
End Select
Set FrmInfoCNTR.ConRec = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)

AddOr = False
End Sub

Function CheckEntry() As Boolean
Dim x As String
'if addor = True then    'Don't Check selected.
If FrmInfoCNTR.LDVIEW.SelectedItem Is Nothing Then GoTo CheckNotList
If Trim(TID.Text) = Trim(FrmInfoCNTR.LDVIEW.SelectedItem.Text) Then Exit Function
CheckNotList:
Set FrmInfoCNTR.ConRec = New ADODB.Recordset
With FrmInfoCNTR.ConRec
    .ActiveConnection = FrmInfoCNTR.ConX
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    
    If .State <> 0 Then
    .Close
    End If
    .Open "Select * FROM TBA_SPI"
    .Properties.Refresh
     .Find " IDNO = '" & Trim(TID.Text) & "'", , adSearchForward, 1
    If Not .EOF Then
    MsgBox "This id is in use", vbCritical, "Error"
    TID.SetFocus
    SendKeys "{Home}+{End}"
    End If
    .Close
End With
Set FrmInfoCNTR.ConRec = Nothing
End Function

Private Sub TID_LostFocus()
CheckEntry
End Sub

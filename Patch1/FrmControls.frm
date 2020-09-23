VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form FrmControls 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control Sheets\Classes\Courses Reports"
   ClientHeight    =   7890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8520
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmControls.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   8520
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CDL 
      Left            =   6480
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox TSCHLYR 
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
      Left            =   1320
      TabIndex        =   67
      Top             =   0
      Width           =   1335
   End
   Begin VB.ComboBox CSEM 
      Appearance      =   0  'Flat
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
      ItemData        =   "FrmControls.frx":57E2
      Left            =   3720
      List            =   "FrmControls.frx":57EF
      Style           =   2  'Dropdown List
      TabIndex        =   54
      Top             =   0
      Width           =   735
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7455
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   13150
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Control Sheets"
      TabPicture(0)   =   "FrmControls.frx":5802
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "LBLCount"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "TBGrades"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "LCONTROLSHEET"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "IXD"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "FRM3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "CBCreate"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "CBupdate"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "FRM5"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "CBGET"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Classes"
      TabPicture(1)   =   "FrmControls.frx":581E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "CsectionCLass"
      Tab(1).Control(1)=   "LViewClass"
      Tab(1).Control(2)=   "CBCLASSCREATE"
      Tab(1).Control(3)=   "LblCountX"
      Tab(1).Control(4)=   "Label1"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Course"
      TabPicture(2)   =   "FrmControls.frx":583A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "LblCourseCount"
      Tab(2).Control(1)=   "Label29"
      Tab(2).Control(2)=   "CBCRSE"
      Tab(2).Control(3)=   "LviewCourse"
      Tab(2).Control(4)=   "CCourses"
      Tab(2).Control(5)=   "CYRLEVEL"
      Tab(2).Control(6)=   "CBGO1"
      Tab(2).ControlCount=   7
      TabCaption(3)   =   "Custom Reports"
      TabPicture(3)   =   "FrmControls.frx":5856
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "LDVIEW"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Frame1"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Fr1"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "CBGEN"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "cdx"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "FR2"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "FR3"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "CBEXCEL"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).ControlCount=   8
      Begin ISAPTECH.chameleonButton CBEXCEL 
         Height          =   495
         Left            =   -68400
         TabIndex        =   98
         Top             =   2880
         Width           =   1575
         _extentx        =   2778
         _extenty        =   873
         und             =   0   'False
         iname           =   "Tahoma"
         btype           =   14
         tx              =   "Export to Excel"
         enab            =   -1  'True
         coltype         =   1
         focusr          =   -1  'True
         bcol            =   8421504
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
      Begin VB.Frame FR3 
         Appearance      =   0  'Flat
         Caption         =   " Parameters "
         ForeColor       =   &H80000008&
         Height          =   2295
         Left            =   -70440
         TabIndex        =   92
         Top             =   480
         Visible         =   0   'False
         Width           =   3615
         Begin VB.CheckBox CpAll3 
            Caption         =   "Report Both School"
            Height          =   375
            Left            =   840
            TabIndex        =   95
            Top             =   720
            Width           =   1935
         End
         Begin VB.ComboBox CSCHL 
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
            ItemData        =   "FrmControls.frx":5872
            Left            =   840
            List            =   "FrmControls.frx":587C
            TabIndex        =   93
            Top             =   360
            Width           =   2655
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "School:"
            Height          =   240
            Left            =   120
            TabIndex        =   94
            Top             =   360
            Width           =   645
         End
      End
      Begin VB.Frame FR2 
         Appearance      =   0  'Flat
         Caption         =   " Parameters "
         ForeColor       =   &H80000008&
         Height          =   2295
         Left            =   -70440
         TabIndex        =   87
         Top             =   480
         Visible         =   0   'False
         Width           =   3615
         Begin VB.CheckBox CpAll2 
            Caption         =   "Report all Course/YR"
            Height          =   375
            Left            =   1080
            TabIndex        =   91
            Top             =   720
            Width           =   2175
         End
         Begin VB.ComboBox CpCourse2 
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
            Left            =   1080
            TabIndex        =   89
            Top             =   360
            Width           =   1575
         End
         Begin VB.ComboBox CpYr1 
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
            ItemData        =   "FrmControls.frx":58D8
            Left            =   2640
            List            =   "FrmControls.frx":58EB
            Style           =   2  'Dropdown List
            TabIndex        =   88
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Course\YR:"
            Height          =   240
            Left            =   120
            TabIndex        =   90
            Top             =   360
            Width           =   975
         End
      End
      Begin MSComctlLib.ImageList cdx 
         Left            =   -74880
         Top             =   2880
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmControls.frx":58FE
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin ISAPTECH.chameleonButton CBGEN 
         Height          =   495
         Left            =   -70080
         TabIndex        =   96
         Top             =   2880
         Width           =   1575
         _extentx        =   2778
         _extenty        =   873
         und             =   0   'False
         iname           =   "Tahoma"
         btype           =   14
         tx              =   "&Generate Query"
         enab            =   -1  'True
         coltype         =   1
         focusr          =   -1  'True
         bcol            =   8421504
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
      Begin VB.Frame Fr1 
         Appearance      =   0  'Flat
         Caption         =   " Parameters "
         ForeColor       =   &H80000008&
         Height          =   2295
         Left            =   -70440
         TabIndex        =   83
         Top             =   480
         Width           =   3615
         Begin VB.CheckBox CpAll1 
            Caption         =   "Report all Courses"
            Height          =   375
            Left            =   960
            TabIndex        =   86
            Top             =   720
            Width           =   1935
         End
         Begin VB.ComboBox CpCourse1 
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
            Left            =   960
            TabIndex        =   84
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Course:"
            Height          =   240
            Left            =   240
            TabIndex        =   85
            Top             =   360
            Width           =   675
         End
      End
      Begin ISAPTECH.chameleonButton CBGET 
         Height          =   375
         Left            =   120
         TabIndex        =   77
         Top             =   2640
         Width           =   1575
         _extentx        =   2778
         _extenty        =   661
         und             =   0   'False
         iname           =   "Tahoma"
         btype           =   14
         tx              =   "GET STUDENTS"
         enab            =   -1  'True
         coltype         =   1
         focusr          =   -1  'True
         bcol            =   8421504
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
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         Caption         =   " Report Type "
         ForeColor       =   &H80000008&
         Height          =   2295
         Left            =   -74880
         TabIndex        =   76
         Top             =   480
         Width           =   4335
         Begin VB.OptionButton o5 
            Caption         =   "Count Student by School"
            Height          =   375
            Left            =   240
            TabIndex        =   82
            Top             =   1800
            Width           =   4000
         End
         Begin VB.OptionButton o4 
            Caption         =   "Count Student by Course and Sex"
            Height          =   375
            Left            =   240
            TabIndex        =   81
            Top             =   1440
            Width           =   4000
         End
         Begin VB.OptionButton o3 
            Caption         =   "Count Student by Course and Year by Sex"
            Height          =   375
            Left            =   240
            TabIndex        =   80
            Top             =   1080
            Width           =   4000
         End
         Begin VB.OptionButton o2 
            Caption         =   "Count Student by Course and Year"
            Height          =   375
            Left            =   240
            TabIndex        =   79
            Top             =   720
            Width           =   4000
         End
         Begin VB.OptionButton o1 
            Caption         =   "Count Student by Course"
            Height          =   375
            Left            =   240
            TabIndex        =   78
            Top             =   360
            Value           =   -1  'True
            Width           =   4000
         End
      End
      Begin ISAPTECH.chameleonButton CBGO1 
         Height          =   375
         Left            =   -71760
         TabIndex        =   75
         Top             =   360
         Width           =   495
         _extentx        =   873
         _extenty        =   661
         und             =   0   'False
         iname           =   "Tahoma"
         btype           =   4
         tx              =   "GO"
         enab            =   -1  'True
         coltype         =   1
         focusr          =   -1  'True
         bcol            =   12632256
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
      Begin VB.ComboBox CYRLEVEL 
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
         ItemData        =   "FrmControls.frx":B0F0
         Left            =   -72480
         List            =   "FrmControls.frx":B103
         Style           =   2  'Dropdown List
         TabIndex        =   74
         Top             =   360
         Width           =   615
      End
      Begin VB.ComboBox CCourses 
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
         Left            =   -74040
         TabIndex        =   69
         Top             =   360
         Width           =   1575
      End
      Begin VB.ComboBox CsectionCLass 
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
         Left            =   -73920
         TabIndex        =   63
         Top             =   480
         Width           =   1575
      End
      Begin VB.Frame FRM5 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   5640
         TabIndex        =   57
         Top             =   5640
         Width           =   2295
         Begin VB.TextBox TRE 
            Height          =   360
            Left            =   1200
            TabIndex        =   59
            Top             =   240
            Width           =   855
         End
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
            ItemData        =   "FrmControls.frx":B116
            Left            =   1200
            List            =   "FrmControls.frx":B129
            TabIndex        =   58
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "RE-EXAM:"
            Height          =   240
            Left            =   240
            TabIndex        =   61
            Top             =   240
            Width           =   870
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "REMARKS:"
            Height          =   240
            Left            =   240
            TabIndex        =   60
            Top             =   600
            Width           =   915
         End
      End
      Begin ISAPTECH.chameleonButton CBupdate 
         Height          =   375
         Left            =   6360
         TabIndex        =   51
         Top             =   6960
         Width           =   1695
         _extentx        =   2990
         _extenty        =   661
         und             =   0   'False
         iname           =   "Tahoma"
         btype           =   14
         tx              =   "UPDATE"
         enab            =   -1  'True
         coltype         =   1
         focusr          =   -1  'True
         bcol            =   8421504
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
      Begin ISAPTECH.chameleonButton CBCreate 
         Height          =   375
         Left            =   4560
         TabIndex        =   50
         Top             =   6960
         Width           =   1695
         _extentx        =   2990
         _extenty        =   661
         und             =   0   'False
         iname           =   "Tahoma"
         btype           =   14
         tx              =   "CREATE REPORT"
         enab            =   -1  'True
         coltype         =   1
         focusr          =   -1  'True
         bcol            =   8421504
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
      Begin VB.Frame FRM3 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   2175
         Left            =   120
         TabIndex        =   39
         Top             =   360
         Width           =   7815
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
            Left            =   1080
            TabIndex        =   52
            Top             =   240
            Width           =   1575
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
            Height          =   315
            Left            =   3480
            TabIndex        =   44
            Top             =   600
            Width           =   495
         End
         Begin VB.TextBox TSD 
            Height          =   795
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   43
            Top             =   1200
            Width           =   3855
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
            Left            =   1080
            TabIndex        =   42
            Top             =   600
            Width           =   1575
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
            Left            =   5040
            TabIndex        =   41
            Top             =   600
            Width           =   2655
         End
         Begin VB.TextBox TSCHED 
            Height          =   795
            Left            =   4200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   40
            Top             =   1200
            Width           =   3495
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "SECTION:"
            Height          =   240
            Left            =   120
            TabIndex        =   53
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "SUBJECT:"
            Height          =   240
            Left            =   120
            TabIndex        =   49
            Top             =   600
            Width           =   840
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "DESCRIPTION:"
            Height          =   240
            Left            =   120
            TabIndex        =   48
            Top             =   960
            Width           =   1260
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "UNITS:"
            Height          =   240
            Left            =   2760
            TabIndex        =   47
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "TEACHER:"
            Height          =   240
            Left            =   4080
            TabIndex        =   46
            Top             =   600
            Width           =   885
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "SCHEDULE:"
            Height          =   240
            Left            =   4200
            TabIndex        =   45
            Top             =   960
            Width           =   975
         End
      End
      Begin MSComctlLib.ImageList IXD 
         Left            =   120
         Top             =   4080
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmControls.frx":B149
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView LCONTROLSHEET 
         Height          =   3735
         Left            =   120
         TabIndex        =   1
         Top             =   3120
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   6588
         View            =   3
         Arrange         =   2
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "IXD"
         SmallIcons      =   "IXD"
         ColHdrIcons     =   "IXD"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   22
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "IDNO"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "NAME"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "COURSE"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "YR"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "CS"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "TQ"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "TT"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "PRELIMS"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "CS"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "TQ"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "TT"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "MIDTERMS"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   12
            Text            =   "CS"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "TQ"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "TT"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   15
            Text            =   "SEMIS"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   16
            Text            =   "CS"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   17
            Text            =   "TQ"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   18
            Text            =   "TT"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   19
            Text            =   "FINALS"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   20
            Text            =   "REEXAM"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   21
            Text            =   "REMARKS"
            Object.Width           =   2540
         EndProperty
      End
      Begin TabDlg.SSTab TBGrades 
         Height          =   3015
         Left            =   5640
         TabIndex        =   2
         Top             =   2640
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
         TabPicture(0)   =   "FrmControls.frx":B7DB
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Line1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label13"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label12"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label11"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Label10"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "CBCom1"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "tave1"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "tt1"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "tq1"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "tcs1"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).ControlCount=   10
         TabCaption(1)   =   "MIDTERMS"
         TabPicture(1)   =   "FrmControls.frx":B7F7
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "TCS2"
         Tab(1).Control(1)=   "TQ2"
         Tab(1).Control(2)=   "TT2"
         Tab(1).Control(3)=   "TAVE2"
         Tab(1).Control(4)=   "CBCom2"
         Tab(1).Control(5)=   "Label14"
         Tab(1).Control(6)=   "Label15"
         Tab(1).Control(7)=   "Label16"
         Tab(1).Control(8)=   "Label17"
         Tab(1).Control(9)=   "Line2"
         Tab(1).ControlCount=   10
         TabCaption(2)   =   "SEMI-FINALS"
         TabPicture(2)   =   "FrmControls.frx":B813
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "TCS3"
         Tab(2).Control(1)=   "TQ3"
         Tab(2).Control(2)=   "TT3"
         Tab(2).Control(3)=   "TAVE3"
         Tab(2).Control(4)=   "CBCom3"
         Tab(2).Control(5)=   "Label18"
         Tab(2).Control(6)=   "Label19"
         Tab(2).Control(7)=   "Label20"
         Tab(2).Control(8)=   "Label21"
         Tab(2).Control(9)=   "Line3"
         Tab(2).ControlCount=   10
         TabCaption(3)   =   "FINALS"
         TabPicture(3)   =   "FrmControls.frx":B82F
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "TCS4"
         Tab(3).Control(1)=   "TQ4"
         Tab(3).Control(2)=   "TT4"
         Tab(3).Control(3)=   "TAVE4"
         Tab(3).Control(4)=   "CBCom4"
         Tab(3).Control(5)=   "Label22"
         Tab(3).Control(6)=   "Label23"
         Tab(3).Control(7)=   "Label24"
         Tab(3).Control(8)=   "Label25"
         Tab(3).Control(9)=   "Line4"
         Tab(3).ControlCount=   10
         Begin VB.TextBox tcs1 
            Height          =   360
            Left            =   840
            TabIndex        =   19
            Top             =   120
            Width           =   495
         End
         Begin VB.TextBox tq1 
            Height          =   360
            Left            =   840
            TabIndex        =   18
            Top             =   480
            Width           =   495
         End
         Begin VB.TextBox tt1 
            Height          =   360
            Left            =   840
            TabIndex        =   17
            Top             =   840
            Width           =   495
         End
         Begin VB.TextBox tave1 
            Enabled         =   0   'False
            Height          =   360
            Left            =   840
            TabIndex        =   16
            Top             =   1320
            Width           =   495
         End
         Begin VB.TextBox TCS2 
            Height          =   360
            Left            =   -74160
            TabIndex        =   14
            Top             =   120
            Width           =   495
         End
         Begin VB.TextBox TQ2 
            Height          =   360
            Left            =   -74160
            TabIndex        =   13
            Top             =   480
            Width           =   495
         End
         Begin VB.TextBox TT2 
            Height          =   360
            Left            =   -74160
            TabIndex        =   12
            Top             =   840
            Width           =   495
         End
         Begin VB.TextBox TAVE2 
            Enabled         =   0   'False
            Height          =   360
            Left            =   -74160
            TabIndex        =   11
            Top             =   1320
            Width           =   495
         End
         Begin VB.TextBox TCS3 
            Height          =   360
            Left            =   -74160
            TabIndex        =   10
            Top             =   120
            Width           =   495
         End
         Begin VB.TextBox TQ3 
            Height          =   360
            Left            =   -74160
            TabIndex        =   9
            Top             =   480
            Width           =   495
         End
         Begin VB.TextBox TT3 
            Height          =   360
            Left            =   -74160
            TabIndex        =   8
            Top             =   840
            Width           =   495
         End
         Begin VB.TextBox TAVE3 
            Enabled         =   0   'False
            Height          =   360
            Left            =   -74160
            TabIndex        =   7
            Top             =   1320
            Width           =   495
         End
         Begin VB.TextBox TCS4 
            Height          =   360
            Left            =   -74160
            TabIndex        =   6
            Top             =   120
            Width           =   495
         End
         Begin VB.TextBox TQ4 
            Height          =   360
            Left            =   -74160
            TabIndex        =   5
            Top             =   480
            Width           =   495
         End
         Begin VB.TextBox TT4 
            Height          =   360
            Left            =   -74160
            TabIndex        =   4
            Top             =   840
            Width           =   495
         End
         Begin VB.TextBox TAVE4 
            Enabled         =   0   'False
            Height          =   360
            Left            =   -74160
            TabIndex        =   3
            Top             =   1320
            Width           =   495
         End
         Begin ISAPTECH.chameleonButton CBCom1 
            Height          =   375
            Left            =   1200
            TabIndex        =   15
            Top             =   1800
            Width           =   855
            _extentx        =   1508
            _extenty        =   661
            und             =   0
            iname           =   "Tahoma"
            btype           =   5
            tx              =   "Compute"
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
         Begin ISAPTECH.chameleonButton CBCom2 
            Height          =   375
            Left            =   -73800
            TabIndex        =   20
            Top             =   1800
            Width           =   855
            _extentx        =   1508
            _extenty        =   661
            und             =   0
            iname           =   "Tahoma"
            btype           =   5
            tx              =   "Compute"
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
         Begin ISAPTECH.chameleonButton CBCom3 
            Height          =   375
            Left            =   -73800
            TabIndex        =   21
            Top             =   1800
            Width           =   855
            _extentx        =   1508
            _extenty        =   661
            und             =   0
            iname           =   "Tahoma"
            btype           =   5
            tx              =   "Compute"
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
         Begin ISAPTECH.chameleonButton CBCom4 
            Height          =   375
            Left            =   -73800
            TabIndex        =   22
            Top             =   1800
            Width           =   855
            _extentx        =   1508
            _extenty        =   661
            und             =   0
            iname           =   "Tahoma"
            btype           =   5
            tx              =   "Compute"
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
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "CS:"
            Height          =   240
            Left            =   120
            TabIndex        =   38
            Top             =   120
            Width           =   315
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "TQ:"
            Height          =   240
            Left            =   120
            TabIndex        =   37
            Top             =   480
            Width           =   330
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "TT:"
            Height          =   240
            Left            =   120
            TabIndex        =   36
            Top             =   840
            Width           =   315
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "AVE:"
            Height          =   240
            Left            =   120
            TabIndex        =   35
            Top             =   1320
            Width           =   420
         End
         Begin VB.Line Line1 
            X1              =   1320
            X2              =   120
            Y1              =   1200
            Y2              =   1200
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "CS:"
            Height          =   240
            Left            =   -74880
            TabIndex        =   34
            Top             =   120
            Width           =   315
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "TQ:"
            Height          =   240
            Left            =   -74880
            TabIndex        =   33
            Top             =   480
            Width           =   330
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "TT:"
            Height          =   240
            Left            =   -74880
            TabIndex        =   32
            Top             =   840
            Width           =   315
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "AVE:"
            Height          =   240
            Left            =   -74880
            TabIndex        =   31
            Top             =   1320
            Width           =   420
         End
         Begin VB.Line Line2 
            X1              =   -73680
            X2              =   -74880
            Y1              =   1200
            Y2              =   1200
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "CS:"
            Height          =   240
            Left            =   -74880
            TabIndex        =   30
            Top             =   120
            Width           =   315
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "TQ:"
            Height          =   240
            Left            =   -74880
            TabIndex        =   29
            Top             =   480
            Width           =   330
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "TT:"
            Height          =   240
            Left            =   -74880
            TabIndex        =   28
            Top             =   840
            Width           =   315
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "AVE:"
            Height          =   240
            Left            =   -74880
            TabIndex        =   27
            Top             =   1320
            Width           =   420
         End
         Begin VB.Line Line3 
            X1              =   -73680
            X2              =   -74880
            Y1              =   1200
            Y2              =   1200
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "CS:"
            Height          =   240
            Left            =   -74880
            TabIndex        =   26
            Top             =   120
            Width           =   315
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "TQ:"
            Height          =   240
            Left            =   -74880
            TabIndex        =   25
            Top             =   480
            Width           =   330
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "TT:"
            Height          =   240
            Left            =   -74880
            TabIndex        =   24
            Top             =   840
            Width           =   315
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "AVE:"
            Height          =   240
            Left            =   -74880
            TabIndex        =   23
            Top             =   1320
            Width           =   420
         End
         Begin VB.Line Line4 
            X1              =   -73680
            X2              =   -74880
            Y1              =   1200
            Y2              =   1200
         End
      End
      Begin MSComctlLib.ListView LViewClass 
         Height          =   5775
         Left            =   -73920
         TabIndex        =   62
         Top             =   960
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   10186
         View            =   3
         Arrange         =   2
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "IXD"
         SmallIcons      =   "IXD"
         ColHdrIcons     =   "IXD"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "IDNO"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "NAME"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "COURSE"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "YR"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "SEX"
            Object.Width           =   2540
         EndProperty
      End
      Begin ISAPTECH.chameleonButton CBCLASSCREATE 
         Height          =   495
         Left            =   -68640
         TabIndex        =   68
         Top             =   6840
         Width           =   1695
         _extentx        =   2990
         _extenty        =   1085
         und             =   0
         iname           =   "Tahoma"
         btype           =   14
         tx              =   "CREATE REPORT"
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
      Begin MSComctlLib.ListView LviewCourse 
         Height          =   5895
         Left            =   -74040
         TabIndex        =   70
         Top             =   840
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   10398
         View            =   3
         Arrange         =   2
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "IXD"
         SmallIcons      =   "IXD"
         ColHdrIcons     =   "IXD"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "IDNO"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "NAME"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "SEX"
            Object.Width           =   2540
         EndProperty
      End
      Begin ISAPTECH.chameleonButton CBCRSE 
         Height          =   495
         Left            =   -68640
         TabIndex        =   71
         Top             =   6840
         Width           =   1695
         _extentx        =   2990
         _extenty        =   1085
         und             =   0
         iname           =   "Tahoma"
         btype           =   14
         tx              =   "CREATE REPORT"
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
      Begin MSComctlLib.ListView LDVIEW 
         Height          =   3855
         Left            =   -74880
         TabIndex        =   97
         Top             =   3480
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   6800
         View            =   3
         Arrange         =   2
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "cdx"
         SmallIcons      =   "cdx"
         ColHdrIcons     =   "cdx"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "Course\YR:"
         Height          =   240
         Left            =   -75000
         TabIndex        =   73
         Top             =   360
         Width           =   975
      End
      Begin VB.Label LblCourseCount 
         AutoSize        =   -1  'True
         Caption         =   "Nothing Selected"
         Height          =   240
         Left            =   -74040
         TabIndex        =   72
         Top             =   6720
         Width           =   1440
      End
      Begin VB.Label LblCountX 
         AutoSize        =   -1  'True
         Caption         =   "Nothing Selected"
         Height          =   240
         Left            =   -73920
         TabIndex        =   66
         Top             =   6720
         Width           =   1440
      End
      Begin VB.Label LBLCount 
         AutoSize        =   -1  'True
         Caption         =   "Nothing Selected"
         Height          =   240
         Left            =   240
         TabIndex        =   65
         Top             =   6960
         Width           =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "SECTION:"
         Height          =   240
         Left            =   -74880
         TabIndex        =   64
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "School Year:"
      Height          =   240
      Left            =   120
      TabIndex        =   56
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Semester:"
      Height          =   240
      Left            =   2760
      TabIndex        =   55
      Top             =   0
      Width           =   900
   End
End
Attribute VB_Name = "FrmControls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public RsCS As ADODB.Recordset
Public RsSecs As ADODB.Recordset
Public RsSubs As ADODB.Recordset
Public RsDetails As ADODB.Recordset
Public RSSCHED As ADODB.Recordset
Public RsCl As ADODB.Recordset
Dim RecCrse As ADODB.Recordset
Dim RsCrse1 As ADODB.Recordset

'Commands and Execution Recordsets
Dim OBJCom As ADODB.Command
Dim RsObj As ADODB.Recordset
Dim SQLPASTQUERY As String

Private Sub CBCLASSCREATE_Click()
'Create CLassreport
Dim msg As String, SCHL As String, SYR As String, SEMX As String
msg = "Select IDNO,STUDENT,COURSE, YEARLEVEL,SEX from GRADING_SYS Where "
msg = msg & "Schoolyear ='" & TSCHLYR.Text & "' and SEMESTER ='" & CSEM.Text
msg = msg & "' and Grading_Sys.Section='" & CsectionCLass.Text
msg = msg & "' GROUP BY IDNO, STUDENT,COURSE,YEARLEVEL,SEX ORDER BY SEX DESC, IDNO ASC"
SetDenver msg, 4
With DR4
    SCHL = InputBox("Enter School:", "SCHOOL")
    SYR = "SCHOOL YEAR " & TSCHLYR.Text
    Select Case CSEM.Text
        Case "1st"
            SEMX = "First Semester"
        Case "2nd"
            SEMX = "Second Semester"
        Case "Sum"
            SEMX = "Summer"
    End Select
    .Sections("PageHeader").Controls("LBLSCHOOL").Caption = UCase(SCHL)
    .Sections("PageHeader").Controls("LBLSY").Caption = SYR
    .Sections("PageHeader").Controls("LBLSEM").Caption = SEMX
    .Sections("PageHeader").Controls("LBLSECTION").Caption = UCase(CsectionCLass.Text)
End With
DR4.Refresh
DR4.Show 1
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

Private Sub CBCreate_Click()
Dim msg As String
msg = "Select * From Grading_SYS where SCHOOLYEAR='" & TSCHLYR.Text
msg = msg & "' and Semester = '" & CSEM.Text & "' and Subject = '"
msg = msg & CSubjects.Text & "' and Schedule='" & TSCHED.Text & "' Order by SEX DESC, IDNO ASC"
PrepareCS "ControlSheet.xls", msg, CDL
End Sub

Private Sub CBGEN_Click()
Dim msg As String
Set OBJCom = Nothing
Set RsObj = Nothing
Set OBJCom = New ADODB.Command
Set RsObj = New ADODB.Recordset
With OBJCom
    .ActiveConnection = FrmInfoCNTR.ConX
    .CommandType = adCmdStoredProc
    .CommandTimeout = 600
    If o1.Value = True Then     'Course Report
        'Check if CpAll1 = vbchecked
        If CpAll1.Value = vbChecked Then
            'GEt Op1_AllCourse
            .CommandText = "Op1_AllCourse"
            .Parameters.Append .CreateParameter("@SY", adVarChar, adParamInput, 10, TSCHLYR.Text)
            .Parameters.Append .CreateParameter("@SEM", adVarChar, adParamInput, 3, CSEM.Text)
        Else
            'Get Op1_SingleCourse
            .CommandText = "Op1_SingleCourse"
            .Parameters.Append .CreateParameter("SY", adVarChar, adParamInput, 10, TSCHLYR.Text)
            .Parameters.Append .CreateParameter("SEM", adVarChar, adParamInput, 3, CSEM.Text)
            .Parameters.Append .CreateParameter("CRS", adVarChar, adParamInput, 15, CpCourse1.Text)
        End If
    End If
    If o2.Value = True Then
    If CpAll2.Value = vbChecked Then
            'GEt Op2_AllYear
            .CommandText = "Op2_AllYear"
            .Parameters.Append .CreateParameter("SY", adVarChar, adParamInput, 10, TSCHLYR.Text)
            .Parameters.Append .CreateParameter("SEM", adVarChar, adParamInput, 3, CSEM.Text)
        Else
            'Get Op2_SingleYear
            .CommandText = "Op2_SingleYear"
            .Parameters.Append .CreateParameter("SY", adVarChar, adParamInput, 10, TSCHLYR.Text)
            .Parameters.Append .CreateParameter("SEM", adVarChar, adParamInput, 3, CSEM.Text)
            .Parameters.Append .CreateParameter("CRS", adVarChar, adParamInput, 15, CpCourse2.Text)
            .Parameters.Append .CreateParameter("YR", adVarChar, adParamInput, 15, CpYr1.Text)
        End If
    End If
    If o3.Value = True Then
    If CpAll2.Value = vbChecked Then
            'GEt Op3_AllYear
            .CommandText = "Op3_AllYear"
            .Parameters.Append .CreateParameter("SY", adVarChar, adParamInput, 10, TSCHLYR.Text)
            .Parameters.Append .CreateParameter("SEM", adVarChar, adParamInput, 3, CSEM.Text)
        Else
            'Get Op3_SingleYear
            .CommandText = "Op3_SingleYear"
            .Parameters.Append .CreateParameter("SY", adVarChar, adParamInput, 10, TSCHLYR.Text)
            .Parameters.Append .CreateParameter("SEM", adVarChar, adParamInput, 3, CSEM.Text)
            .Parameters.Append .CreateParameter("CRS", adVarChar, adParamInput, 15, CpCourse2.Text)
            .Parameters.Append .CreateParameter("YR", adVarChar, adParamInput, 15, CpYr1.Text)
        End If
    End If
    If o4.Value = True Then
    If CpAll1.Value = vbChecked Then
            'GEt Op4_AllCourse
            .CommandText = "Op4_AllCourse"
            .Parameters.Append .CreateParameter("SY", adVarChar, adParamInput, 10, TSCHLYR.Text)
            .Parameters.Append .CreateParameter("SEM", adVarChar, adParamInput, 3, CSEM.Text)
        Else
            'Get Op4_SingleCourse
            .CommandText = "Op4_SingleCourse"
            .Parameters.Append .CreateParameter("SY", adVarChar, adParamInput, 10, TSCHLYR.Text)
            .Parameters.Append .CreateParameter("SEM", adVarChar, adParamInput, 3, CSEM.Text)
            .Parameters.Append .CreateParameter("CRS", adVarChar, adParamInput, 15, CpCourse1.Text)
        End If
    End If
    If o5.Value = True Then
        If CpAll3.Value = vbChecked Then
            'GEt Op1_AllCourse
            .CommandText = "Opt5_both"
            .Parameters.Append .CreateParameter("SY", adVarChar, adParamInput, 10, TSCHLYR.Text)
            .Parameters.Append .CreateParameter("SEM", adVarChar, adParamInput, 3, CSEM.Text)
        Else
            'Get Op1_SingleCourse
            .CommandText = "Opt5_Single"
            .Parameters.Append .CreateParameter("SY", adVarChar, adParamInput, 10, TSCHLYR.Text)
            .Parameters.Append .CreateParameter("SEM", adVarChar, adParamInput, 3, CSEM.Text)
            .Parameters.Append .CreateParameter("SCHL", adVarChar, adParamInput, 100, CSCHL.Text)
        End If
    End If
    
    Set RsObj = .Execute
    LoadToView
    RsObj.Close
End With
End Sub

Sub LoadToView()
Dim i As Long, j As Long, fvAL As Long
With RsObj
    'Set Ldview
    
    LDVIEW.ColumnHeaders.Clear
    LDVIEW.ListItems.Clear
    For i = 1 To .Fields.Count
        LDVIEW.ColumnHeaders.Add i, , .Fields(i - 1).Name, 1440
    Next
    Do Until .EOF
        i = .AbsolutePosition
        LDVIEW.ListItems.Add i, , .Fields(0).Value, 1, 1
        For j = 0 To .Fields.Count
            If (j + 1) < .Fields.Count Then
            LDVIEW.ListItems(i).SubItems(j + 1) = .Fields(j + 1).Value
            End If
        Next
        .MoveNext
    Loop
End With
End Sub

Private Sub CBExcel_Click()
If LDVIEW.ListItems.Count <> 0 Then
Me.MousePointer = vbHourglass
EXCELMODULE.PrepareExport CDL, LDVIEW
Me.MousePointer = vbDefault
End If
End Sub


Private Sub CBGET_Click()
Set RsCS = New ADODB.Recordset
Get_Students TSCHLYR.Text, CSEM.Text, CSubjects.Text, TSCHED.Text, Cteacher.Text, RsCS, LCONTROLSHEET
LBLCount.Caption = "Total Students: " & LCONTROLSHEET.ListItems.Count
LCONTROLSHEET.SetFocus
LCONTROLSHEET_Click

End Sub

Private Sub CBGO1_Click()
'XSS
ListCourse
End Sub

Private Sub CBupdate_Click()
On Error GoTo ErrorX
If LCONTROLSHEET.SelectedItem Is Nothing Then Exit Sub
If MsgBox("Are you sure you want to update the selected record?", vbQuestion + vbYesNo, "Update Grades") = vbYes Then
UpdateSelected
End If
Exit Sub
ErrorX:
    ErrorTrap Err, "Updating Grades"
    Me.Refresh
    Set RsCS = New ADODB.Recordset
    Get_Students TSCHLYR.Text, CSEM.Text, CSubjects.Text, TSCHED.Text, Cteacher.Text, RsCS, LCONTROLSHEET
End Sub

Function UpdateSelected()
Dim msg As String
Set RsCS = Nothing
Set RsCS = New ADODB.Recordset
msg = "Select * From Grading_Sys WHERE SCHOOLYEAR='" & TSCHLYR.Text
msg = msg & "' and SEMESTER ='" & CSEM.Text & "' and IDNO='" & LCONTROLSHEET.SelectedItem.Text
msg = msg & "' and Subject ='" & CSubjects.Text & "'"
With RsCS
    .ActiveConnection = FrmInfoCNTR.ConX
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open msg
    .Fields("p1").Value = tcs1.Text
    .Fields("p2").Value = tq1.Text
    .Fields("p3").Value = tt1.Text
    .Fields("prelim").Value = tave1.Text
    .Fields("m1").Value = TCS2.Text
    .Fields("m2").Value = TQ2.Text
    .Fields("m3").Value = TT2.Text
    .Fields("midterm").Value = TAVE2.Text
    .Fields("s1").Value = TCS3.Text
    .Fields("s2").Value = TQ3.Text
    .Fields("s3").Value = TT3.Text
    .Fields("semi").Value = TAVE3.Text
    .Fields("f1").Value = TCS4.Text
    .Fields("f2").Value = TQ4.Text
    .Fields("f3").Value = TT4.Text
    .Fields("FInals").Value = TAVE4.Text
    .Fields("REEXAM").Value = TRE.Text
    .Fields("REMARKS").Value = CREM.Text
    .Update
    .Properties.Refresh
    .Close
    'Reload All
    Set RsCS = New ADODB.Recordset
    Get_Students TSCHLYR.Text, CSEM.Text, CSubjects.Text, TSCHED.Text, Cteacher.Text, RsCS, LCONTROLSHEET

End With
End Function

Private Sub cbCRSE_Click()
'Create CLassreport
Dim msg As String, SCHL As String, SYR As String, SEMX As String, XPR As String
msg = "Select IDNO,STUDENT,COURSE, YEARLEVEL,SEX from GRADING_SYS Where "
msg = msg & "Schoolyear ='" & TSCHLYR.Text & "' and SEMESTER ='" & CSEM.Text
msg = msg & "' and Course ='" & CCourses.Text & "' and YEARLEVEL='" & CYRLEVEL.Text
msg = msg & "' GROUP BY IDNO, STUDENT,COURSE,YEARLEVEL, SEX ORDER BY SEX DESC, IDNO ASC"
SetDenver msg, 4
With DR5
    SCHL = InputBox("Enter School:", "SCHOOL")
    SYR = "SCHOOL YEAR " & TSCHLYR.Text
    Select Case CSEM.Text
        Case "1st"
            SEMX = "First Semester"
        Case "2nd"
            SEMX = "Second Semester"
        Case "Sum"
            SEMX = "Summer"
    End Select
    XPR = UCase(CCourses.Text & " " & CYRLEVEL.Text)
    .Sections("PageHeader").Controls("LBLSCHOOL").Caption = UCase(SCHL)
    .Sections("PageHeader").Controls("LBLSY").Caption = SYR
    .Sections("PageHeader").Controls("LBLSEM").Caption = SEMX
    .Sections("PageHeader").Controls("LBLCRSE").Caption = XPR
End With
DR5.Refresh
DR5.Show 1
End Sub

Private Sub CSECTION_LostFocus()
'Call ListSubs(CSECTION.Text, RsSecs, CSubjects)
Me.Get_SubsKo
End Sub

Private Sub CsectionCLass_LostFocus()
'X
ListClass
End Sub

Sub ListClass()
Dim msg As String, i As Long
msg = "Select IDNO, STUDENT,COURSE,YEARLEVEL,SEX from Grading_SYS where Section='"
msg = msg & CsectionCLass.Text & "' and SchoolYear='" & TSCHLYR.Text & "' and Semester='"
msg = msg & CSEM.Text & "'GROUP by IDNO, STUDENT,COURSE, YEARLEVEL,SEX order by  SEX DESC,IDNO ASC"
Set RsCl = Nothing
Set RsCl = New ADODB.Recordset
With RsCl
    .ActiveConnection = FrmInfoCNTR.ConX
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open msg
    LViewClass.ListItems.Clear
    Do Until .EOF
        i = .AbsolutePosition
        LViewClass.ListItems.Add i, , .Fields("IDNO").Value, 1, 1
        LViewClass.ListItems(i).SubItems(1) = .Fields("STUDENT").Value
        LViewClass.ListItems(i).SubItems(2) = .Fields("COURSE").Value
        LViewClass.ListItems(i).SubItems(3) = .Fields("YEARLEVEL").Value
        LViewClass.ListItems(i).SubItems(4) = .Fields("SEX").Value
        .MoveNext
    Loop
    LblCountX.Caption = "Total Students: " & .RecordCount
End With

End Sub

Private Sub CSEM_LostFocus()
ProduceSectionList
End Sub

Private Sub CSubjects_lostfocus()
If MsgBox("System can generate details and schedule for this section and subject," & _
    " do you want to use this feature now?", vbQuestion + vbYesNo, "Use Schedules") = vbYes Then
'SubjectDetails
Me.Get_Det
Set RsCS = New ADODB.Recordset
Get_Students TSCHLYR.Text, CSEM.Text, CSubjects.Text, TSCHED.Text, Cteacher.Text, RsCS, LCONTROLSHEET
LBLCount.Caption = "Total Students: " & LCONTROLSHEET.ListItems.Count
LCONTROLSHEET.SetFocus
LCONTROLSHEET_Click
End If

End Sub

Private Sub Form_Load()
On Error GoTo Errbx:
Get_Courses
GetSY
Exit Sub
Errbx:
    ErrorTrap Err, "Form Loading"
End Sub

Sub ProduceSectionList()
Dim msg As String
msg = "Select SECTION from GRADING_SYS"
msg = msg & " WHERE SCHOOLYEAR='" & TSCHLYR.Text & "'"
msg = msg & " and SEMESTER = '" & CSEM.Text & "'"
msg = msg & " GROUP by SCHOOLYEAR, SEMESTER, SECTION Order by Section"

Set RsSecs = Nothing
Set RsSecs = New ADODB.Recordset
With RsSecs
    .ActiveConnection = FrmInfoCNTR.ConX
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open msg
    Me.CSECTION.Clear
    Me.CsectionCLass.Clear
    'Loadnow
    Do Until .EOF
        CSECTION.AddItem .Fields(0).Value
        CsectionCLass.AddItem .Fields(0).Value
        .MoveNext
    Loop
End With
End Sub


Sub GetSched()
Set RSSCHED = New ADODB.Recordset
Dim msg As String, i As Long, P1 As String, _
    P2 As String, P3 As String, P4 As String, _
    NL As String, Adx As String, strx As String
msg = "Select * From Scheduling Where Subject='"
msg = msg & CSubjects.Text & "' and Sectionko like '%"
msg = msg & CSECTION.Text & "%'  order by Subject"
With RSSCHED
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
        If Left(Adx, 1) = "," Or Left(Adx, 1) = " " Then
            Adx = Right(Adx, Len(Adx) - 1)
        End If
        If Adx = "" Then
            Adx = P3 & " "
        Else
            Adx = Adx & "," & P3 & " "
        End If
    Else
        TSCHED.Text = TSCHED.Text & Adx & P1 & " - " & P2 & " "
        Adx = Left(.Fields("DATE").Value, 1)
        If Adx = "T" Then
            If Left(.Fields("DATE").Value, 2) = "TH" Then
                Adx = "TH"
            End If
        End If
        Adx = Adx & " "
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
        TSCHED.Text = TSCHED.Text & Adx & P1 & " - " & P2 & " "
        End If
    Loop
    If Right(TSCHED.Text, 1) = " " Then
        TSCHED.Text = Left(TSCHED.Text, Len(TSCHED.Text) - 1)
    End If
End With
Set RSSCHED = Nothing
End Sub

Sub SubjectDetails()    'Upon Leave in Subjects
'Generate Details before the schedule
Set RsDetails = New ADODB.Recordset
Dim msg As String
msg = "Select * From Scheduling Where Subject = '"
msg = msg & CSubjects.Text & "' order by Scheduling.date"
With RsDetails
    .ActiveConnection = FrmInfoCNTR.ConX
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open msg
    If .RecordCount = 0 Then Exit Sub
    Tunits.Text = .Fields("UNITS").Value
    Cteacher.Text = .Fields("teacher").Value
    TSD.Text = .Fields("SUBJECT_DESCRIPTION").Value
    
    GetSched
End With
Set RsDetails = Nothing
End Sub


Private Sub Form_Unload(Cancel As Integer)
FrmSet.WCR.Visible = False
End Sub

Private Sub LCONTROLSHEET_Click()
If LCONTROLSHEET.SelectedItem Is Nothing Then Exit Sub
GradeLoad
End Sub

Sub GradeLoad()
With LCONTROLSHEET
TransferGrade tcs1, .SelectedItem.SubItems(4)
TransferGrade tq1, .SelectedItem.SubItems(5)
TransferGrade tt1, .SelectedItem.SubItems(6)
TransferGrade tave1, .SelectedItem.SubItems(7)
TransferGrade TCS2, .SelectedItem.SubItems(8)
TransferGrade TQ2, .SelectedItem.SubItems(9)
TransferGrade TT2, .SelectedItem.SubItems(10)
TransferGrade TAVE2, .SelectedItem.SubItems(11)
TransferGrade TCS3, .SelectedItem.SubItems(12)
TransferGrade TQ3, .SelectedItem.SubItems(13)
TransferGrade TT3, .SelectedItem.SubItems(14)
TransferGrade TAVE3, .SelectedItem.SubItems(15)
TransferGrade TCS4, .SelectedItem.SubItems(16)
TransferGrade TQ4, .SelectedItem.SubItems(17)
TransferGrade TT4, .SelectedItem.SubItems(18)
TransferGrade TAVE4, .SelectedItem.SubItems(19)
TransferGrade TRE, .SelectedItem.SubItems(20)
TransferGrade CREM, .SelectedItem.SubItems(21)
End With
End Sub

Function TransferGrade(OBJ As Object, txt As String)
OBJ.Text = txt
End Function

Sub GetSY()
Dim msg As String
msg = "SELECT SCHOOLYEAR From GRADING_SYS GROUP BY SCHOOLYEAR"
Set RsSecs = Nothing
Set RsSecs = New ADODB.Recordset
With RsSecs
    .ActiveConnection = FrmInfoCNTR.ConX
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open msg
    TSCHLYR.Clear
    'Loadnow
    Do Until .EOF
        TSCHLYR.AddItem .Fields(0).Value
        .MoveNext
    Loop
End With
End Sub

Sub Get_Courses()   'Onload
Set RecCrse = Nothing
Set RecCrse = New ADODB.Recordset
Dim msg As String
msg = "Select * From Courses"
With RecCrse
    .ActiveConnection = FrmInfoCNTR.ConX
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open msg
    CCourses.Clear
    CpCourse1.Clear
    CpCourse2.Clear
    Do Until .EOF
        CpCourse1.AddItem .Fields(0).Value
        CpCourse2.AddItem .Fields(0).Value
        CCourses.AddItem .Fields(0).Value
        .MoveNext
    Loop
End With

End Sub

Sub ListCourse()
Dim msg As String, i As Long
msg = "Select IDNO, STUDENT,SEX from Grading_SYS where COURSE='" & CCourses.Text & "' and YearLevel='"
msg = msg & CYRLEVEL.Text & "' and SchoolYear='" & TSCHLYR.Text & "' and Semester='"
msg = msg & CSEM.Text & "'GROUP by IDNO, STUDENT,COURSE, YEARLEVEL,SEX order by SEX DESC, IDNO ASC"
Set RsCrse1 = Nothing
Set RsCrse1 = New ADODB.Recordset
With RsCrse1
    .ActiveConnection = FrmInfoCNTR.ConX
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open msg
    LviewCourse.ListItems.Clear
    Do Until .EOF
        i = .AbsolutePosition
        LviewCourse.ListItems.Add i, , .Fields("IDNO").Value, 1, 1
        LviewCourse.ListItems(i).SubItems(1) = .Fields("STUDENT").Value
        LviewCourse.ListItems(i).SubItems(2) = .Fields("SEX").Value
        .MoveNext
    Loop
    Me.LblCourseCount.Caption = "Total Students: " & .RecordCount
End With

End Sub

Private Sub o1_Click()
'Fr1
Fr1.Visible = True
FR2.Visible = 0
FR3.Visible = 0
End Sub

Private Sub o2_Click()
'Fr2
FR2.Visible = True
FR3.Visible = 0
Fr1.Visible = 0
End Sub

Private Sub o3_Click()
'Fr2
FR2.Visible = True
FR3.Visible = 0
Fr1.Visible = 0
End Sub

Private Sub o4_Click()
'Fr1
Fr1.Visible = True
FR2.Visible = 0
FR3.Visible = 0
End Sub

Private Sub o5_Click()
'Fr3
FR3.Visible = True
FR2.Visible = 0
Fr1.Visible = 0
End Sub

Sub Get_SubsKo()
Dim msg As String
msg = "Select Subject from Grading_Sys WHERE "
msg = msg & " SCHOOLYEAR ='" & TSCHLYR.Text & "'"
msg = msg & " and SEMESTER='" & CSEM.Text & "'"
msg = msg & " and SECTION='" & CSECTION.Text & "'"
msg = msg & "GROUP by SUBJECT"
Set RsSecs = Nothing
Set RsSecs = New ADODB.Recordset
With RsSecs
    .ActiveConnection = FrmInfoCNTR.ConX
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open msg
    CSubjects.Clear
    Do Until .EOF
        CSubjects.AddItem .Fields(0).Value
        .MoveNext
    Loop
    .Close
End With
Set RsSecs = Nothing
End Sub

Sub Get_Det()
Dim msg  As String
Set RsDetails = Nothing
Set RsDetails = New ADODB.Recordset
msg = "SELECT TEACHER, UNITS, SUBJECT_DESCRIPTION, SCHEDULE From Grading_SYS"
msg = msg & " WHERE SCHOOLYEAR = '" & TSCHLYR.Text & "' and Semester='" & CSEM.Text & "'"
msg = msg & " and section='" & CSECTION.Text & "' and SUBJECT='"
msg = msg & CSubjects.Text & "' GROUP BY TEACHER,UNITS,SUBJECT_DESCRIPTION,SCHEDULE"
With RsDetails
.ActiveConnection = FrmInfoCNTR.ConX
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open msg
If .RecordCount = 0 Then Exit Sub
    Tunits.Text = .Fields(1).Value
    Cteacher.Text = .Fields(0).Value
    TSD.Text = .Fields(2).Value
    TSCHED.Text = .Fields(3).Value
    .Close
End With
Set RsDetails = Nothing
End Sub

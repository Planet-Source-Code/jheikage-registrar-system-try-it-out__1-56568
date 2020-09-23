VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmTrans 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transfer Data"
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   6720
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmTrans.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   6720
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   7095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   12515
      _Version        =   393216
      Style           =   1
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
      TabCaption(0)   =   "Subjects"
      TabPicture(0)   =   "FrmTrans.frx":57E2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Line1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Tsched"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cbbrow1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cbBegSubs"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "CDL"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Scheduling"
      TabPicture(1)   =   "FrmTrans.frx":57FE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label3"
      Tab(1).Control(1)=   "Line2"
      Tab(1).Control(2)=   "Label6"
      Tab(1).Control(3)=   "Label8"
      Tab(1).Control(4)=   "Line3"
      Tab(1).Control(5)=   "Line4"
      Tab(1).Control(6)=   "LbCntSched"
      Tab(1).Control(7)=   "lblSelInfo"
      Tab(1).Control(8)=   "Line5"
      Tab(1).Control(9)=   "TTransSched"
      Tab(1).Control(10)=   "cBsearch2"
      Tab(1).Control(11)=   "CbTransSched"
      Tab(1).ControlCount=   12
      TabCaption(2)   =   "Grades"
      TabPicture(2)   =   "FrmTrans.frx":581A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "SSTab2"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin TabDlg.SSTab SSTab2 
         Height          =   6495
         Left            =   -74880
         TabIndex        =   30
         Top             =   480
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   11456
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Create Grading Disk"
         TabPicture(0)   =   "FrmTrans.frx":5836
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "LBLCount"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label14"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label15"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "FRM3"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "lx"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "IXD"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "TSCHLYR"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "CSEMG"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "CBCREATEGD"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "CBGET"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).ControlCount=   10
         TabCaption(1)   =   "Get Grades From Disk"
         TabPicture(1)   =   "FrmTrans.frx":5852
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "CBGDCHECK"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "LSource"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "cbTransfer"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "LBLCntSRC"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "Label16"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).ControlCount=   5
         Begin ISAPTECH.chameleonButton CBGDCHECK 
            Height          =   495
            Left            =   -74880
            TabIndex        =   55
            Top             =   600
            Width           =   1575
            _extentx        =   2778
            _extenty        =   873
            und             =   0
            iname           =   "Tahoma"
            btype           =   14
            tx              =   "GET GD"
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
         Begin ISAPTECH.chameleonButton CBGET 
            Height          =   495
            Left            =   120
            TabIndex        =   54
            Top             =   3000
            Width           =   1575
            _extentx        =   2778
            _extenty        =   873
            und             =   0
            iname           =   "Tahoma"
            btype           =   14
            tx              =   "GET STUDENTS"
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
         Begin MSComctlLib.ListView LSource 
            Height          =   3855
            Left            =   -74880
            TabIndex        =   52
            Top             =   1680
            Width           =   6135
            _ExtentX        =   10821
            _ExtentY        =   6800
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            Icons           =   "IXD"
            SmallIcons      =   "IXD"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   33
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "IDNO"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "STUDENT"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "SEX"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "SCHOOLYEAR"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "SEMESTER"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "COURSE"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "YEARLEVEL"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "MAJOR"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Text            =   "SCHOOL"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   9
               Text            =   "SECTION"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   10
               Text            =   "SUBJECT"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   11
               Text            =   "UNITS"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   12
               Text            =   "SUBJECT DESCRIPTION"
               Object.Width           =   8819
            EndProperty
            BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   13
               Text            =   "TEACHER"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   14
               Text            =   "SCHEDULE"
               Object.Width           =   8819
            EndProperty
            BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   15
               Text            =   "CS"
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   16
               Text            =   "TQ"
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(18) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   17
               Text            =   "TT"
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(19) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   18
               Text            =   "PRELIM"
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(20) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   19
               Text            =   "CS"
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(21) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   20
               Text            =   "TQ"
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(22) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   21
               Text            =   "TT"
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(23) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   22
               Text            =   "MIDTERM"
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(24) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   23
               Text            =   "CS"
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(25) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   24
               Text            =   "TQ"
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(26) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   25
               Text            =   "TT"
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(27) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   26
               Text            =   "SEMI"
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(28) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   27
               Text            =   "CS"
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(29) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   28
               Text            =   "TQ"
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(30) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   29
               Text            =   "TT"
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(31) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   30
               Text            =   "FINALS"
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(32) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   31
               Text            =   "REEXAM"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(33) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   32
               Text            =   "REMARKS"
               Object.Width           =   2540
            EndProperty
         End
         Begin ISAPTECH.chameleonButton CBCREATEGD 
            Height          =   495
            Left            =   4800
            TabIndex        =   50
            Top             =   5760
            Width           =   1455
            _extentx        =   2566
            _extenty        =   873
            und             =   0
            iname           =   "Tahoma"
            btype           =   14
            tx              =   "Create GD"
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
         Begin VB.ComboBox CSEMG 
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
            ItemData        =   "FrmTrans.frx":586E
            Left            =   3720
            List            =   "FrmTrans.frx":587B
            Style           =   2  'Dropdown List
            TabIndex        =   47
            Top             =   480
            Width           =   735
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
            TabIndex        =   46
            Top             =   480
            Width           =   1335
         End
         Begin MSComctlLib.ImageList IXD 
            Left            =   2880
            Top             =   4320
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
                  Picture         =   "FrmTrans.frx":588E
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ListView lx 
            Height          =   2055
            Left            =   120
            TabIndex        =   44
            Top             =   3600
            Width           =   6135
            _ExtentX        =   10821
            _ExtentY        =   3625
            View            =   3
            Arrange         =   2
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            Icons           =   "IXD"
            SmallIcons      =   "IXD"
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
         Begin VB.Frame FRM3 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   2175
            Left            =   120
            TabIndex        =   31
            Top             =   720
            Width           =   6135
            Begin VB.TextBox TSched1 
               Height          =   795
               Left            =   3240
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   37
               Top             =   1200
               Width           =   2775
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
               Left            =   4080
               TabIndex        =   36
               Top             =   600
               Width           =   1935
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
               TabIndex        =   35
               Top             =   600
               Width           =   1575
            End
            Begin VB.TextBox TSD 
               Height          =   795
               Left            =   120
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   34
               Top             =   1200
               Width           =   2895
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
               Left            =   4080
               TabIndex        =   33
               Top             =   240
               Width           =   495
            End
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
               TabIndex        =   32
               Top             =   240
               Width           =   1575
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               Caption         =   "SCHEDULE:"
               Height          =   240
               Left            =   3240
               TabIndex        =   43
               Top             =   960
               Width           =   975
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               Caption         =   "TEACHER:"
               Height          =   240
               Left            =   3120
               TabIndex        =   42
               Top             =   600
               Width           =   885
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "UNITS:"
               Height          =   240
               Left            =   3360
               TabIndex        =   41
               Top             =   240
               Width           =   615
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "DESCRIPTION:"
               Height          =   240
               Left            =   120
               TabIndex        =   40
               Top             =   960
               Width           =   1260
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "SUBJECT:"
               Height          =   240
               Left            =   120
               TabIndex        =   39
               Top             =   600
               Width           =   840
            End
            Begin VB.Label Label28 
               AutoSize        =   -1  'True
               Caption         =   "SECTION:"
               Height          =   240
               Left            =   120
               TabIndex        =   38
               Top             =   240
               Width           =   855
            End
         End
         Begin ISAPTECH.chameleonButton cbTransfer 
            Height          =   495
            Left            =   -70320
            TabIndex        =   56
            Top             =   5640
            Width           =   1575
            _extentx        =   2778
            _extenty        =   873
            und             =   0
            iname           =   "Tahoma"
            btype           =   14
            tx              =   "Begin Transfer"
            enab            =   0
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
         Begin VB.Label LBLCntSRC 
            AutoSize        =   -1  'True
            Caption         =   "Nothing Selected"
            Height          =   240
            Left            =   -74880
            TabIndex        =   53
            Top             =   5520
            Width           =   1440
         End
         Begin VB.Label Label16 
            Caption         =   "List present on Selected Grading Disk:"
            Height          =   495
            Left            =   -74880
            TabIndex        =   51
            Top             =   1200
            Width           =   1935
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Semester:"
            Height          =   240
            Left            =   2760
            TabIndex        =   49
            Top             =   480
            Width           =   900
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "School Year:"
            Height          =   240
            Left            =   120
            TabIndex        =   48
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label LBLCount 
            AutoSize        =   -1  'True
            Caption         =   "Nothing Selected"
            Height          =   240
            Left            =   120
            TabIndex        =   45
            Top             =   5640
            Width           =   1440
         End
      End
      Begin MSComDlg.CommonDialog CDL 
         Left            =   6120
         Top             =   600
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         Filter          =   "Scheduling Project(*.mdb)|*.mdb"
         Flags           =   2
      End
      Begin VB.CommandButton CbTransSched 
         Caption         =   "Transfer Schedules"
         Enabled         =   0   'False
         Height          =   615
         Left            =   -69960
         TabIndex        =   22
         Top             =   4320
         Width           =   1575
      End
      Begin VB.CommandButton cBsearch2 
         Caption         =   "..."
         Height          =   375
         Left            =   -68880
         TabIndex        =   21
         Top             =   2400
         Width           =   495
      End
      Begin VB.TextBox TTransSched 
         Height          =   375
         Left            =   -72480
         TabIndex        =   20
         Top             =   2400
         Width           =   3615
      End
      Begin VB.CommandButton cbBegSubs 
         Caption         =   "Transfer Subjects"
         Enabled         =   0   'False
         Height          =   615
         Left            =   5040
         TabIndex        =   17
         Top             =   3120
         Width           =   1575
      End
      Begin VB.Frame Frame2 
         Caption         =   "Courses/Subjects Transferred in the Schedulign Project"
         Height          =   1815
         Left            =   0
         TabIndex        =   15
         Top             =   3720
         Width           =   6615
         Begin VB.ComboBox Cyear 
            Height          =   360
            ItemData        =   "FrmTrans.frx":5F20
            Left            =   2760
            List            =   "FrmTrans.frx":5F33
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   240
            Width           =   615
         End
         Begin VB.ListBox LCourses 
            Height          =   1500
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Width           =   2655
         End
         Begin VB.ListBox LSubs 
            Height          =   1500
            Left            =   3360
            TabIndex        =   16
            Top             =   240
            Width           =   3015
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Subjects to Transfer"
         Height          =   1335
         Left            =   0
         TabIndex        =   5
         Top             =   1680
         Width           =   6615
         Begin VB.CommandButton CBGo 
            Caption         =   "Subjects"
            Height          =   495
            Left            =   5520
            TabIndex        =   26
            ToolTipText     =   "Count Subjects"
            Top             =   720
            Width           =   975
         End
         Begin VB.ComboBox CSem 
            Height          =   360
            ItemData        =   "FrmTrans.frx":5F46
            Left            =   3120
            List            =   "FrmTrans.frx":5F53
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   720
            Width           =   615
         End
         Begin VB.ComboBox CCourses 
            Height          =   360
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   240
            Width           =   1095
         End
         Begin VB.ComboBox CYRL 
            Height          =   360
            ItemData        =   "FrmTrans.frx":5F66
            Left            =   1920
            List            =   "FrmTrans.frx":5F79
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   240
            Width           =   615
         End
         Begin VB.ComboBox CSY 
            Height          =   360
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   720
            Width           =   855
         End
         Begin VB.ComboBox CMJR 
            Height          =   360
            Left            =   4200
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   240
            Width           =   2295
         End
         Begin VB.Label LblTotSubs 
            AutoSize        =   -1  'True
            Caption         =   "Subject Count: 0"
            Height          =   240
            Left            =   3840
            TabIndex        =   25
            Top             =   720
            Width           =   1440
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Semester:"
            Height          =   240
            Left            =   2160
            TabIndex        =   14
            Top             =   720
            Width           =   900
         End
         Begin VB.Label Label4 
            Caption         =   "Course:"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Lb1 
            Caption         =   "Schoolyear:"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label5 
            Caption         =   "Major(Optional):"
            Height          =   255
            Left            =   2640
            TabIndex        =   8
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.CommandButton cbbrow1 
         Caption         =   "..."
         Height          =   375
         Left            =   6120
         TabIndex        =   4
         Top             =   1200
         Width           =   495
      End
      Begin VB.TextBox Tsched 
         Height          =   375
         Left            =   2640
         TabIndex        =   3
         Top             =   1200
         Width           =   3495
      End
      Begin VB.Line Line5 
         X1              =   -74880
         X2              =   -68400
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Label lblSelInfo 
         AutoSize        =   -1  'True
         Caption         =   "No Project Selected."
         Height          =   240
         Left            =   -74880
         TabIndex        =   29
         Top             =   3120
         Width           =   1740
      End
      Begin VB.Label LbCntSched 
         AutoSize        =   -1  'True
         Caption         =   "Your Scheduling System Contains 0 Records."
         Height          =   240
         Left            =   -74880
         TabIndex        =   28
         Top             =   5280
         Width           =   3840
      End
      Begin VB.Line Line4 
         X1              =   -74880
         X2              =   -68400
         Y1              =   5160
         Y2              =   5160
      End
      Begin VB.Line Line3 
         X1              =   -74880
         X2              =   -68400
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Label Label8 
         Caption         =   $"FrmTrans.frx":5F8C
         Height          =   735
         Left            =   -74880
         TabIndex        =   27
         Top             =   1200
         Width           =   6255
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Search Scheduling project:"
         Height          =   240
         Left            =   -75000
         TabIndex        =   19
         Top             =   2400
         Width           =   2325
      End
      Begin VB.Line Line2 
         X1              =   -74880
         X2              =   -68400
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label Label3 
         Caption         =   "Transfer Schedules to your Server from a Scheduling Project to be used for adding/editng records for this SY and Semester."
         Height          =   495
         Left            =   -74760
         TabIndex        =   18
         Top             =   480
         Width           =   6255
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Search Scheduling project:"
         Height          =   240
         Left            =   120
         TabIndex        =   2
         Top             =   1200
         Width           =   2325
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   6600
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label Label1 
         Caption         =   "Transfer subjects to a Scheduling Project for scheduling. This action requires a scheduling project."
         Height          =   495
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   6255
      End
   End
End
Attribute VB_Name = "FrmTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RecSAM As ADODB.Recordset
Dim RecSS As ADODB.Recordset
Dim WithEvents Schedu As ADODB.Connection
Attribute Schedu.VB_VarHelpID = -1
Dim WithEvents TransSched As ADODB.Connection
Attribute TransSched.VB_VarHelpID = -1
Dim RecSS1 As ADODB.Recordset
Dim RecDest As ADODB.Recordset
Dim RSSETS As ADODB.Recordset
Dim ErrCount As String
Dim TransFile As String
Dim RSSCHED As ADODB.Recordset
Dim RsDet As ADODB.Recordset
Public RsTRANS As ADODB.Recordset
Dim RsGrTr As ADODB.Recordset
Private Sub cbBegSubs_Click()
If MsgBox("Continue File Transfer Action?", vbYesNo + vbQuestion, "FileTransfer") = vbNo Then Exit Sub
TransferSubs
End Sub

Sub CountSchedule()
Set RecSS1 = Nothing
Set RecSS1 = New ADODB.Recordset
With RecSS1
    .ActiveConnection = FrmInfoCNTR.ConX
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open "Select * From Scheduling"
    Me.LbCntSched.Caption = "There are " & .RecordCount & " records present in your scheduling system."
    .Close
End With
End Sub

Private Sub cbbrow1_Click()
Dim xso
With CDL
    .DialogTitle = "Select Scheduling Project"
    .Filter = "Scheduling Project(*.mdb)|*.mdb"
    .ShowOpen
    If Trim(.FileName) = "" Then Exit Sub
    Set xso = CreateObject("Scripting.Filesystemobject")
    If xso.fileexists(.FileName) Then
        If TestBrowsed = False Then
        Me.Tsched.Text = ""
        Else
        Me.Tsched.Text = .FileName
        End If
        LoadListSubsFrmSCHD
    End If
    ErrCount = ""
End With
Set xso = Nothing
End Sub

Sub ConnectToSched()
If Schedu Is Nothing Then
Set Schedu = Nothing
Set Schedu = New ADODB.Connection
Else
    If Schedu.State <> 0 Then Schedu.Close
End If
With Schedu
    .ConnectionString = "Provider=MICROSOFT.JET.OLEDB.4.0;Persist Security Info=false;JET OLEDB:Database password=vip;Data source=" & CDL.FileName
    .CursorLocation = adUseClient
    .Open
End With
End Sub

Function TestBrowsed() As Boolean
Dim msg As String
On Error GoTo ErbX
ConnectToSched
Set RecSS = Nothing
Set RecSS = New ADODB.Recordset
With RecSS
    .ActiveConnection = Schedu
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open "Select * From ProjectInfo"
    msg = "Project Name: " & .Fields("PRJ").Value
    msg = msg & vbNewLine & "SY: " & .Fields("Schl_Year").Value
    MsgBox "Scheduling Project is valid." & vbNewLine & msg & vbNewLine & "Click ok to load records to memory.", vbInformation, "Valid File"
End With
Me.cbBegSubs.Enabled = 1
TestBrowsed = True
Exit Function
ErbX:
    TestBrowsed = False
    ErrorTrap Err, "Test Scheduling Project"
    Me.Refresh
    Me.cbBegSubs.Enabled = 0
    Set RecSS = Nothing
    LCourses.Clear
    If Schedu.State = 0 Then Exit Function
    Schedu.Close
End Function

Sub LoadListSubsFrmSCHD()
On Error GoTo ErbX
Dim msg As String, i As Long
Set RecSS = Nothing
Set RecSS = New ADODB.Recordset
msg = "Select Class From Subjects Group by Class"
With RecSS
    .ActiveConnection = Schedu
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open msg
    LCourses.Clear
    Do Until .EOF
        If LCourses.ListCount = 0 Then GoTo AddinX
        For i = 0 To LCourses.ListCount - 1
            LCourses.ListIndex = i
            If LCourses.Text = .Fields("Class").Value Then GoTo Mover
        Next
AddinX:
        LCourses.AddItem .Fields("Class").Value
Mover:
        .MoveNext
    Loop
End With
Exit Sub
ErbX:
    ErrorTrap Err, "Loading Subjects to Memory"
    Me.Refresh
End Sub

Sub LoadCourses()
Dim msg As String
Set RecSAM = Nothing
Set RecSAM = New ADODB.Recordset
msg = "Select * From Courses order by School"
With RecSAM
    .ActiveConnection = FrmInfoCNTR.ConX
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open msg
    Do Until .EOF
        CCourses.AddItem .Fields("Course").Value
        .MoveNext
    Loop
    .Close
    Set RecSAM = Nothing
    If CCourses.ListCount <> 0 Then CCourses.ListIndex = 0
    CYRL.ListIndex = 0
    CSEM.ListIndex = 0
End With
End Sub

Sub LoadSchlYr_and_MJR()
Dim msg As String, i As Long
Set RecSAM = Nothing
Set RecSAM = New ADODB.Recordset
msg = "Select Course, Schoolyear, Major,YearLevel"
msg = msg & " from Curriculas where Course = '" & CCourses.Text
msg = msg & "' and yearlevel ='" & CYRL.Text
msg = msg & "' GROUP by COurse, SchoolYear, Major, Yearlevel"
With RecSAM
    .ActiveConnection = FrmInfoCNTR.ConX
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open msg
    CSY.Clear
    CMJR.Clear
    Do Until .EOF
                CSY.AddItem .Fields("Schoolyear").Value
                CMJR.AddItem .Fields("Major").Value
NextRec1:
        .MoveNext
    Loop
End With

End Sub

Private Sub CBCREATEGD_Click()
Dim msg As String
msg = "Select * From Grading_SYS where SCHOOLYEAR='" & TSCHLYR.Text
msg = msg & "' and Semester = '" & CSEMG.Text & "' and Subject = '"
msg = msg & CSubjects.Text & "' and Schedule='" & TSched1.Text & "' Order by SEX desc, IDNO ASC"

EXCELMODULE.PrepareCS "Grading_Sys.xls", msg, CDL
End Sub

Private Sub CBGDCHECK_Click()
'XP
On Error GoTo ErrorCheck
GetDrive LSource
LBLCntSRC.Caption = "Total Students: " & LSource.ListItems.Count
Me.cbTransfer.Enabled = True
Exit Sub
ErrorCheck:
    ErrorTrap Err, "Checking Disk"
    cbTransfer.Enabled = False
End Sub

Private Sub CBGET_Click()
Set RSSETS = New ADODB.Recordset
Get_Students TSCHLYR.Text, CSEMG.Text, CSubjects.Text, TSched1.Text, Cteacher.Text, RSSETS, lx
LBLCount.Caption = "Total Students: " & lx.ListItems.Count
lx.SetFocus

End Sub

Private Sub CBGo_Click()
Dim msg As String, i As Long
Set RecSAM = Nothing
Set RecSAM = New ADODB.Recordset
msg = "Select Count(SubjectCode) as CountX, Course, Schoolyear, Major,YearLevel, Semester"
msg = msg & " from Curriculas where Course = '" & CCourses.Text
msg = msg & "' and yearlevel ='" & CYRL.Text
msg = msg & "' and Schoolyear ='" & CSY.Text
msg = msg & "' and Semester ='" & CSEM.Text
msg = msg & "' GROUP by COurse, SchoolYear, Major, Yearlevel,Semester"
With RecSAM
    .ActiveConnection = FrmInfoCNTR.ConX
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open msg
    If .RecordCount = 0 Then
    Me.LblTotSubs.Caption = "Total Subjects: 0"
    Else
    Me.LblTotSubs.Caption = "Total Subjects:" & .Fields("Countx").Value
    End If
    .Close
End With
End Sub

Private Sub cBsearch2_Click()

Dim xso
With CDL
    .DialogTitle = "Select Scheduling Project"
    .Filter = "Scheduling Project(*.mdb)|*.mdb"
    .ShowOpen
    If Trim(.FileName) = "" Then Exit Sub
    Set xso = CreateObject("Scripting.Filesystemobject")
    If xso.fileexists(.FileName) Then
        TTransSched.Text = .FileName
        CbTransSched.Enabled = True
        ConnectTransfer
    End If
End With

Set xso = Nothing
End Sub

Sub ConnectTransfer()
On Error GoTo ErbX
Dim FSO
Set FSO = CreateObject("Scripting.Filesystemobject")

If TransSched Is Nothing Then
Set TransSched = Nothing
Set TransSched = New ADODB.Connection
Else
If TransSched.State <> 0 Then TransSched.Close
If TTransSched.Text = "" Then
MsgBox "No Project Selected.", vbInformation
Exit Sub
End If
If Not FSO.fileexists(TTransSched.Text) Then Exit Sub
End If
TransSched.ConnectionString = "Provider=MICROSOFT.JET.OLEDB.4.0;Persist Security Info=false;JET OLEDB:Database password=vip;Data source=" & TTransSched.Text
TransSched.CursorLocation = adUseClient
TransSched.Open
TestTrans
Set FSO = Nothing
Exit Sub
ErbX:
    ErrorTrap Err, "Transfer Data"
    Set FSO = Nothing
    Me.Refresh
End Sub
Function TestTrans() As Boolean
Dim msg As String
On Error GoTo ErbX
Set RecSS1 = Nothing
Set RecSS1 = New ADODB.Recordset
With RecSS1
    .ActiveConnection = TransSched
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open "ProjectInfo"
    msg = "Project Name: " & .Fields("PRJ").Value
    msg = msg & vbNewLine & "SY: " & .Fields("Schl_Year").Value
    MsgBox "Scheduling Project is valid." & vbNewLine & msg & vbNewLine & "Click ok to load records to memory.", vbInformation, "Valid File"
    Me.CbTransSched.Enabled = 1
    
End With
TestTrans = True
GetTotal
Exit Function
ErbX:
    ErrorTrap Err, "Testing Project Validity"
    TestTrans = False
    Me.Refresh
    Me.TTransSched.Text = ""
    Me.CbTransSched.Enabled = 0
    If TransSched Is Nothing Then Exit Function
    TransSched.Close
    lblSelInfo.Caption = "No Project Selected."
End Function

Sub GetTotal()
Dim msg As String

Set RecSS1 = Nothing
Set RecSS1 = New ADODB.Recordset
With RecSS1
    .ActiveConnection = TransSched
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open "Select * From Scheduling"
    lblSelInfo.Caption = "There are " & .RecordCount & " schedules in this project."
End With
Set RecSS1 = Nothing

End Sub

Private Sub cbTransfer_Click()

Dim i As Long, msg As String
With LSource
For i = 1 To .ListItems.Count
    .ListItems(i).Selected = True
    msg = "Select * From Grading_Sys where " & _
        "IDNO='" & .ListItems(i).Text & "' and SCHOOLYEAR='" & _
        .ListItems(i).SubItems(3) & "' and SEMESTER='" & _
        .ListItems(i).SubItems(4) & "' and SUBJECT='" & _
        .ListItems(i).SubItems(10) & "' and TEACHER='" & _
        .ListItems(i).SubItems(13) & "' and Schedule='" & _
        .ListItems(i).SubItems(14) & "'"
        MsgBox msg
        TransferGrd msg, i
        
Next
End With
MsgBox "Transfer Complete. Check errors at System Output Window.", vbInformation, "Complete"

End Sub

Private Function TransferGrd(msg As String, i As Long)
On Error GoTo XPR
Set RsGrTr = Nothing
Set RsGrTr = New ADODB.Recordset
With RsGrTr
    .ActiveConnection = FrmInfoCNTR.ConX
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open msg
    If .RecordCount = 0 Then Exit Function
    .Fields("p1").Value = LSource.ListItems(i).SubItems(15)
    .Fields("p2").Value = LSource.ListItems(i).SubItems(16)
    .Fields("p3").Value = LSource.ListItems(i).SubItems(17)
    .Fields("prelim").Value = LSource.ListItems(i).SubItems(18)
    .Fields("m1").Value = LSource.ListItems(i).SubItems(19)
    .Fields("m2").Value = LSource.ListItems(i).SubItems(20)
    .Fields("m3").Value = LSource.ListItems(i).SubItems(21)
    .Fields("midterm").Value = LSource.ListItems(i).SubItems(22)
    .Fields("s1").Value = LSource.ListItems(i).SubItems(23)
    .Fields("s2").Value = LSource.ListItems(i).SubItems(24)
    .Fields("s3").Value = LSource.ListItems(i).SubItems(25)
    .Fields("semi").Value = LSource.ListItems(i).SubItems(26)
    .Fields("f1").Value = LSource.ListItems(i).SubItems(27)
    .Fields("f2").Value = LSource.ListItems(i).SubItems(28)
    .Fields("f3").Value = LSource.ListItems(i).SubItems(29)
    .Fields("finals").Value = LSource.ListItems(i).SubItems(30)
    .Fields("REEXAM").Value = LSource.ListItems(i).SubItems(31)
    .Fields("REMARKS").Value = LSource.ListItems(i).SubItems(32)
    .Update
    .Properties.Refresh
    .Close
End With
Exit Function
XPR:
WriteLog FrmSet.Routputbox, Err.Description
End Function

Private Sub CbTransSched_Click()
'Transfer Data
If MsgBox("Continue File Transfer Action?", vbYesNo + vbQuestion, "FileTransfer") = vbNo Then Exit Sub
Dim itsnum As Double
Set RecSS1 = Nothing
Set RecSS1 = New ADODB.Recordset
With RecSS1
    .ActiveConnection = TransSched
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open "Select * From Scheduling"
    DeleteCurrent
    DeleteClasses
    Dim Ival(11) As String, i As Long
    Do Until .EOF
    For i = 0 To 11
        If IsNull(.Fields(i).Value) Then
        Ival(i) = ""
        Else
        Ival(i) = .Fields(i).Value
        End If
    Next
        AddNewSched Ival(0), Ival(1), Ival(2), Ival(3), Ival(4), Ival(5), Ival(6), Ival(7), Ival(8), Ival(9), Ival(10), Ival(11)
        .MoveNext
        itsnum = Val(.AbsolutePosition) / Val(.RecordCount) * 100
        itsnum = Format(itsnum, "##.##")
        lblSelInfo.Caption = "Transferring Schedules " & itsnum & "%..."
    Loop
    .Close
End With
TransferAllowed
lblSelInfo.Caption = "Transfer Complete."
Me.CountSchedule
End Sub

Sub TransferAllowed()
Dim itsnum As Double
Set RecSS1 = Nothing
Set RecSS1 = New ADODB.Recordset
With RecSS1
    .ActiveConnection = TransSched
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open "Select * From Allowed"
    Do Until .EOF
        AddNewClass .Fields(0).Value, .Fields(1).Value, .Fields(2).Value, .Fields(3).Value
        .MoveNext
        itsnum = Abs(Val(.AbsolutePosition)) / Abs(Val(.RecordCount)) * 100
        itsnum = Format(itsnum, "##.##")
        lblSelInfo.Caption = "Transferring Schedules " & itsnum & "%..."
    Loop
End With
End Sub

Sub AddNewClass(i1 As String, i2 As String, i3 As String, i4 As String)

Set RecDest = Nothing
Set RecDest = New ADODB.Recordset
With RecDest
    .ActiveConnection = FrmInfoCNTR.ConX
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open "Select * From classes"
    .AddNew
    .Fields(0).Value = i1
    .Fields(1).Value = i2
    .Fields(2).Value = i3
    .Fields(3).Value = i4
    .Update
    .Properties.Refresh
    .Close
End With
Set RecDest = Nothing
End Sub

Sub DeleteCurrent()
Set RecDest = Nothing
Set RecDest = New ADODB.Recordset
With RecDest
    .ActiveConnection = FrmInfoCNTR.ConX
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open "Delete From Scheduling"
End With
End Sub

Sub DeleteClasses()
Set RecDest = Nothing
Set RecDest = New ADODB.Recordset
With RecDest
    .ActiveConnection = FrmInfoCNTR.ConX
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open "Delete From classes"
End With
End Sub

Sub AddNewSched(i1 As String, i2 As String, i3 As String, _
    i4 As String, I5 As String, I6 As String, I7 As String, _
    I8 As String, I9 As String, I10 As String, I11 As String, I12 As String)
Set RecDest = Nothing
Set RecDest = New ADODB.Recordset
With RecDest
    .ActiveConnection = FrmInfoCNTR.ConX
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open "Select * From Scheduling"
    .AddNew
    .Fields(0).Value = i1
    .Fields(1).Value = i2
    .Fields(2).Value = i3
    .Fields(3).Value = i4
    .Fields(4).Value = I5
    .Fields(5).Value = I6
    .Fields(6).Value = I7
    .Fields(7).Value = I8
    .Fields(8).Value = I9
    .Fields(9).Value = I10
    .Fields(10).Value = I11
    .Fields(11).Value = I12
    .Update
    .Properties.Refresh
    .Close
    Set RecDest = Nothing
End With
End Sub
    

Private Sub ccourses_LostFocus()
LoadSchlYr_and_MJR
End Sub

Private Sub CSECTION_LostFocus()
ListSubs CSECTION.Text, RSSETS, CSubjects
End Sub

Sub Get_Det()
Dim msg  As String
Set RsDet = Nothing
Set RsDet = New ADODB.Recordset
msg = "SELECT TEACHER, UNITS, SUBJECT_DESCRIPTION, SCHEDULE From Grading_SYS"
msg = msg & " WHERE SCHOOLYEAR = '" & TSCHLYR.Text & "' and Semester='" & CSEM.Text & "'"
msg = msg & " and section='" & CSECTION.Text & "' and SUBJECT='"
msg = msg & CSubjects.Text & "' GROUP BY TEACHER,UNITS,SUBJECT_DESCRIPTION,SCHEDULE"
With RsDet
.ActiveConnection = FrmInfoCNTR.ConX
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open msg
If .RecordCount = 0 Then Exit Sub
    Me.Tunits.Text = .Fields(1).Value
    Me.Cteacher.Text = .Fields(0).Value
    Me.TSD.Text = .Fields(2).Value
    Me.TSched1.Text = .Fields(3).Value
    .Close
End With
Set RsDet = Nothing
End Sub


Private Sub CSEMG_lostfocus()
Me.ProduceSectionList
End Sub

Private Sub CSubjects_lostfocus()
If MsgBox("System can generate details and schedule for this section and subject," & _
    " do you want to use this feature now?", vbQuestion + vbYesNo, "Use Schedules") = vbYes Then
Get_Det
Set RSSETS = New ADODB.Recordset
Get_Students TSCHLYR.Text, CSEMG.Text, CSubjects.Text, TSched1.Text, Cteacher.Text, RSSETS, lx
LBLCount.Caption = "Total Students: " & lx.ListItems.Count
lx.SetFocus
End If
End Sub

Private Sub Cyear_Click()
If LCourses.Text = "" Then Exit Sub     'Non is selected
Dim msg As String, i As Long
Set RecSS = Nothing
Set RecSS = New ADODB.Recordset
msg = "Select Class, Subject,Year_Level From subjects"
msg = msg & " where Class = '" & LCourses.Text
msg = msg & "' and year_level ='" & Cyear.Text & "'"
With RecSS
    .ActiveConnection = Schedu
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open msg
    LSubs.Clear
    Do Until .EOF
        LSubs.AddItem .Fields("Subject").Value
        .MoveNext
    Loop
End With
End Sub

Sub StoredSubs()

End Sub
Private Sub CYRL_LostFocus()
LoadSchlYr_and_MJR
End Sub

Private Sub Form_Load()
Me.GetSY
'Me.ProduceSectionList
LoadCourses
Me.CountSchedule
CheckMe
End Sub
Sub CheckMe()
If STORED_PROC_MAT = P_Val_MAK Then
Me.cbBegSubs.Enabled = False
Me.cbbrow1.Enabled = False
Me.CbTransSched.Enabled = False
Me.CBGo.Enabled = False
Me.cBsearch2.Enabled = False
End If
End Sub
Sub TransferSubs()
Dim msg As String, i As Long
Set RecSAM = Nothing
Set RecSAM = New ADODB.Recordset
msg = "Select * From Curriculas"
msg = msg & " where Course = '" & CCourses.Text
msg = msg & "' and yearlevel ='" & CYRL.Text
msg = msg & "' and SchoolYear='" & CSY.Text
msg = msg & "' and Semester='" & CSEM.Text & "'"
If Trim(CMJR.Text) <> "" Then
msg = msg & " and MAJOR='" & CMJR.Text & "'"
End If
With RecSAM
    .ActiveConnection = FrmInfoCNTR.ConX
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open msg
    Do Until .EOF
        AddNewSubs .Fields("SubjectCode").Value, .Fields("Description").Value, .Fields("Units").Value
        .MoveNext
    Loop
    'Prompt Errors
   LoadListSubsFrmSCHD
End With
End Sub

Sub AddNewSubs(i1 As String, i2 As String, i3 As String)
On Error GoTo ErbX
Set RecSS = Nothing
Set RecSS = New ADODB.Recordset
With RecSS
    .ActiveConnection = Schedu
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open "Select * From Subjects Subjects"
    .AddNew
    .Fields("Class").Value = CCourses.Text
    .Fields("Year_level").Value = CYRL.Text
    .Fields("Subject").Value = i1
    .Fields("subject_Description").Value = i2
    .Fields("Units").Value = i3
    .Update
    .Properties.Refresh
    .Close
End With
Exit Sub
ErbX:
MsgBox ErrCount & "Value: " & i1 & "-" & vbNewLine & Err.Description & vbNewLine, vbCritical, "ERROR"
Set RecSS = Nothing
End Sub

Sub ProduceSectionList()
Dim msg As String
msg = "Select SECTION from GRADING_SYS"
msg = msg & " WHERE SCHOOLYEAR = '" & Me.TSCHLYR.Text
msg = msg & "' and SEMESTER = '" & CSEMG.Text
msg = msg & "' Group By Section Order by Section"
Set RSSETS = Nothing
Set RSSETS = New ADODB.Recordset
With RSSETS
    .ActiveConnection = FrmInfoCNTR.ConX
    .CursorLocation = adUseClient
    .CursorType = adOpenDynamic
    .LockType = adLockOptimistic
    .Open msg
    Me.CSECTION.Clear
    'Loadnow
    Do Until .EOF
        CSECTION.AddItem .Fields(0).Value
        .MoveNext
    Loop
End With
End Sub

Sub GetSY()
Dim msg As String
msg = "SELECT SCHOOLYEAR From GRADING_SYS GROUP BY SCHOOLYEAR"
Set RSSETS = Nothing
Set RSSETS = New ADODB.Recordset
With RSSETS
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

'Sub SubjectDetails()    'Upon Leave in Subjects
'Generate Details before the schedule
'Set RSSETS = New ADODB.Recordset
'Dim msg As String
'msg = "Select * From Scheduling Where Subject = '"
'msg = msg & CSubjects.Text & "' order by Scheduling.date"
'With RSSETS
'    .ActiveConnection = FrmInfoCNTR.ConX
'    .CursorLocation = adUseClient
'    .CursorType = adOpenDynamic
'    .LockType = adLockOptimistic
'    .Open msg
'    If .RecordCount = 0 Then Exit Sub
'    Tunits.Text = .Fields("UNITS").Value
'    Cteacher.Text = .Fields("teacher").Value
'    TSD.Text = .Fields("SUBJECT_DESCRIPTION").Value
'    TSched1.Text = ""
'    GetSched
'End With
'Set RSSETS = Nothing
'End Sub

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
        TSched1.Text = TSched1.Text & Adx & P1 & " - " & P2 & " "
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
        TSched1.Text = TSched1.Text & Adx & P1 & " - " & P2 & " "
        End If
    Loop
    If Right(TSched1.Text, 1) = " " Then
        TSched1.Text = Left(TSched1.Text, Len(TSched1.Text) - 1)
    End If
End With
Set RSSCHED = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
FrmSet.WDT.Visible = False
End Sub

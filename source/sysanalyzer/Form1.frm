VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMain 
   Caption         =   "SysAnalyzer"
   ClientHeight    =   5400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10350
   Icon            =   "Form1.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "frmMain"
   ScaleHeight     =   5400
   ScaleWidth      =   10350
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   8760
      ScaleHeight     =   255
      ScaleWidth      =   1455
      TabIndex        =   32
      Top             =   5100
      Width           =   1455
      Begin VB.Label lblReport 
         Caption         =   "Report"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   0
         Width           =   615
      End
      Begin VB.Label lblTools 
         Caption         =   "Tools"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   840
         TabIndex        =   33
         Top             =   0
         Width           =   435
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5355
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10305
      _ExtentX        =   18177
      _ExtentY        =   9446
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      Tabs            =   7
      Tab             =   6
      TabsPerRow      =   10
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Running Processes"
      TabPicture(0)   =   "Form1.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblTimer"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblDisplay"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lvProcesses"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtProcess"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdAnalyze"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "tmrCountDown"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Open Ports"
      TabPicture(1)   =   "Form1.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lvPorts"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Process Dlls"
      TabPicture(2)   =   "Form1.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtDllPath"
      Tab(2).Control(1)=   "cmdDllProperties"
      Tab(2).Control(2)=   "cmdCopyDll"
      Tab(2).Control(3)=   "lvExplorer"
      Tab(2).Control(4)=   "lvIE"
      Tab(2).Control(5)=   "Label1(0)"
      Tab(2).Control(6)=   "Label1(1)"
      Tab(2).Control(7)=   "Label7"
      Tab(2).ControlCount=   8
      TabCaption(3)   =   "Loaded Drivers"
      TabPicture(3)   =   "Form1.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lvDrivers"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Reg Monitor"
      TabPicture(4)   =   "Form1.frx":04B2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "lvRegKeys"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Api Log"
      TabPicture(5)   =   "Form1.frx":04CE
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "cmdIgnoreApi"
      Tab(5).Control(1)=   "cmdApiDelete"
      Tab(5).Control(2)=   "txtAPIDelete"
      Tab(5).Control(3)=   "txtApiIgnore"
      Tab(5).Control(4)=   "lvAPILog"
      Tab(5).Control(5)=   "Label5"
      Tab(5).Control(6)=   "Label3(2)"
      Tab(5).ControlCount=   7
      TabCaption(6)   =   "Directory Watch Data"
      TabPicture(6)   =   "Form1.frx":04EA
      Tab(6).ControlEnabled=   -1  'True
      Tab(6).Control(0)=   "Label3(1)"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).Control(1)=   "Label3(0)"
      Tab(6).Control(1).Enabled=   0   'False
      Tab(6).Control(2)=   "lvDirWatch"
      Tab(6).Control(2).Enabled=   0   'False
      Tab(6).Control(3)=   "cmdSaveDirWatchFile"
      Tab(6).Control(3).Enabled=   0   'False
      Tab(6).Control(4)=   "cmdDelLike"
      Tab(6).Control(4).Enabled=   0   'False
      Tab(6).Control(5)=   "txtDeleteLike"
      Tab(6).Control(5).Enabled=   0   'False
      Tab(6).Control(6)=   "txtIgnore"
      Tab(6).Control(6).Enabled=   0   'False
      Tab(6).Control(7)=   "cmdDirWatch"
      Tab(6).Control(7).Enabled=   0   'False
      Tab(6).ControlCount=   8
      Begin VB.CommandButton cmdDirWatch 
         Caption         =   "Stop Monitor"
         Height          =   315
         Left            =   9000
         TabIndex        =   36
         Top             =   4500
         Width           =   1215
      End
      Begin VB.CommandButton cmdIgnoreApi 
         Caption         =   "Turn off Api Logging"
         Height          =   315
         Left            =   -66660
         TabIndex        =   35
         Top             =   4500
         Width           =   1935
      End
      Begin VB.Timer tmrCountDown 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   -65160
         Top             =   4560
      End
      Begin VB.TextBox txtIgnore 
         Height          =   315
         Left            =   720
         TabIndex        =   25
         Top             =   4140
         Width           =   9495
      End
      Begin VB.TextBox txtDeleteLike 
         Height          =   315
         Left            =   720
         TabIndex        =   24
         Top             =   4500
         Width           =   4755
      End
      Begin VB.CommandButton cmdDelLike 
         Caption         =   "Delete Lines Like"
         Height          =   315
         Left            =   5580
         TabIndex        =   23
         Top             =   4500
         Width           =   1575
      End
      Begin VB.CommandButton cmdSaveDirWatchFile 
         Caption         =   "Save Selected file"
         Height          =   315
         Left            =   7200
         TabIndex        =   22
         Top             =   4500
         Width           =   1575
      End
      Begin VB.CommandButton cmdApiDelete 
         Caption         =   "Delete Lines Like"
         Height          =   315
         Left            =   -69420
         TabIndex        =   18
         Top             =   4500
         Width           =   1575
      End
      Begin VB.TextBox txtAPIDelete 
         Height          =   315
         Left            =   -74340
         TabIndex        =   17
         Top             =   4500
         Width           =   4815
      End
      Begin VB.TextBox txtApiIgnore 
         Height          =   315
         Left            =   -74340
         TabIndex        =   16
         Top             =   4080
         Width           =   9555
      End
      Begin VB.TextBox txtDllPath 
         Height          =   315
         Left            =   -74400
         TabIndex        =   8
         Top             =   4500
         Width           =   6195
      End
      Begin VB.CommandButton cmdDllProperties 
         Caption         =   "Properties"
         Height          =   315
         Left            =   -68100
         TabIndex        =   7
         Top             =   4500
         Width           =   975
      End
      Begin VB.CommandButton cmdCopyDll 
         Caption         =   "Save Copy"
         Height          =   315
         Left            =   -66780
         TabIndex        =   6
         Top             =   4500
         Width           =   1035
      End
      Begin VB.CommandButton cmdAnalyze 
         Caption         =   "Analyze Process"
         Height          =   315
         Left            =   -72540
         TabIndex        =   5
         Top             =   4560
         Width           =   1875
      End
      Begin VB.TextBox txtProcess 
         Height          =   315
         Left            =   -73860
         TabIndex        =   3
         Top             =   4560
         Width           =   1215
      End
      Begin MSComctlLib.ListView lvPorts 
         Height          =   4755
         Left            =   -74940
         TabIndex        =   2
         Top             =   120
         Width           =   10155
         _ExtentX        =   17912
         _ExtentY        =   8387
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Port"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "PID"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Type"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Path"
            Object.Width           =   12347
         EndProperty
      End
      Begin MSComctlLib.ListView lvProcesses 
         Height          =   4455
         Left            =   -74940
         TabIndex        =   1
         Top             =   60
         Width           =   10155
         _ExtentX        =   17912
         _ExtentY        =   7858
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "PID"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ParentPID"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "User"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Path"
            Object.Width           =   10583
         EndProperty
      End
      Begin MSComctlLib.ListView lvExplorer 
         Height          =   1995
         Left            =   -74940
         TabIndex        =   9
         Top             =   180
         Width           =   10155
         _ExtentX        =   17912
         _ExtentY        =   3519
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "DLL Path"
            Object.Width           =   7937
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Company Name"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "File Description"
            Object.Width           =   7056
         EndProperty
      End
      Begin MSComctlLib.ListView lvIE 
         Height          =   1995
         Left            =   -74940
         TabIndex        =   10
         Top             =   2460
         Width           =   10155
         _ExtentX        =   17912
         _ExtentY        =   3519
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "DLL Path"
            Object.Width           =   7937
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Company Name"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "File Description"
            Object.Width           =   7056
         EndProperty
      End
      Begin MSComctlLib.ListView lvDrivers 
         Height          =   4815
         Left            =   -75000
         TabIndex        =   14
         Top             =   0
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   8493
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Driver File"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Company Name"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Description"
            Object.Width           =   6174
         EndProperty
      End
      Begin MSComctlLib.ListView lvRegKeys 
         Height          =   4815
         Left            =   -74940
         TabIndex        =   15
         Top             =   0
         Width           =   10155
         _ExtentX        =   17912
         _ExtentY        =   8493
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Registry Key"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Value"
            Object.Width           =   6174
         EndProperty
      End
      Begin MSComctlLib.ListView lvAPILog 
         Height          =   4035
         Left            =   -75000
         TabIndex        =   19
         Top             =   0
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   7117
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lvDirWatch 
         Height          =   4035
         Left            =   60
         TabIndex        =   26
         Top             =   0
         Width           =   10155
         _ExtentX        =   17912
         _ExtentY        =   7117
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Action"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "SIze"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "File"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lblDisplay 
         Caption         =   "Currently displaying  base snapshot"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -68760
         TabIndex        =   31
         Top             =   4620
         Width           =   3675
      End
      Begin VB.Label lblTimer 
         Caption         =   "Seconds Remaining: "
         Height          =   315
         Left            =   -70500
         TabIndex        =   30
         Top             =   4620
         Width           =   2955
      End
      Begin VB.Label Label3 
         Caption         =   "Ignore"
         Height          =   315
         Index           =   0
         Left            =   60
         TabIndex        =   28
         Top             =   4140
         Width           =   555
      End
      Begin VB.Label Label3 
         Caption         =   "Prune"
         Height          =   315
         Index           =   1
         Left            =   60
         TabIndex        =   27
         Top             =   4500
         Width           =   555
      End
      Begin VB.Label Label5 
         Caption         =   "Ignore"
         Height          =   315
         Left            =   -75000
         TabIndex        =   21
         Top             =   4140
         Width           =   555
      End
      Begin VB.Label Label3 
         Caption         =   "Prune"
         Height          =   315
         Index           =   2
         Left            =   -74940
         TabIndex        =   20
         Top             =   4560
         Width           =   555
      End
      Begin VB.Label Label1 
         Caption         =   "Explorer Dlls :"
         Height          =   255
         Index           =   0
         Left            =   -74940
         TabIndex        =   13
         Top             =   0
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "IE Dlls :"
         Height          =   255
         Index           =   1
         Left            =   -74940
         TabIndex        =   12
         Top             =   2220
         Width           =   1035
      End
      Begin VB.Label Label7 
         Caption         =   "Path"
         Height          =   255
         Left            =   -74940
         TabIndex        =   11
         Top             =   4560
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "Analyze PID"
         Height          =   315
         Left            =   -74880
         TabIndex        =   4
         Top             =   4560
         Width           =   1215
      End
   End
   Begin VB.OLE OLE1 
      Height          =   30
      Left            =   6720
      TabIndex        =   29
      Top             =   4800
      Width           =   75
   End
   Begin VB.Menu mnuProcessesPopup 
      Caption         =   "mnuProcessesPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuShowProcessDlls 
         Caption         =   "ShowDlls"
      End
      Begin VB.Menu mnuDumpProcess 
         Caption         =   "Dump"
      End
      Begin VB.Menu mnuKillProcess 
         Caption         =   "Kill"
      End
      Begin VB.Menu mnuProcessFileProps 
         Caption         =   "File Properties"
      End
   End
   Begin VB.Menu mnuDllsPopup 
      Caption         =   "mnuDllsPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuViewAllDllProps 
         Caption         =   "View All Properties"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuDumpDll 
         Caption         =   "Dump Module"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuCopyTo 
         Caption         =   "Copy To"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "mnuTools"
      Visible         =   0   'False
      Begin VB.Menu mnuSearch 
         Caption         =   "Search All Tabs"
      End
      Begin VB.Menu mnuCopySelected 
         Caption         =   "Copy All Selected Entries"
      End
      Begin VB.Menu mnuSpacer 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolItem 
         Caption         =   "Show Snapshot 1"
         Index           =   0
      End
      Begin VB.Menu mnuToolItem 
         Caption         =   "Show Snapshot 2"
         Index           =   1
      End
      Begin VB.Menu mnuToolItem 
         Caption         =   "Show Diff report"
         Index           =   2
      End
      Begin VB.Menu mnuToolItem 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuToolItem 
         Caption         =   "Take Snapshot 1"
         Index           =   4
      End
      Begin VB.Menu mnuToolItem 
         Caption         =   "Take Snapshot 2"
         Index           =   5
      End
      Begin VB.Menu mnuToolItem 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuToolItem 
         Caption         =   "Start Over"
         Index           =   7
      End
      Begin VB.Menu mnuToolItem 
         Caption         =   "Show Data Report"
         Index           =   8
         Visible         =   0   'False
      End
      Begin VB.Menu mnuKnownSpacer 
         Caption         =   "-"
      End
      Begin VB.Menu mnuKnownFiles 
         Caption         =   "Build Known File DB"
      End
      Begin VB.Menu mnuHideKnown 
         Caption         =   "Hide Known Files"
      End
      Begin VB.Menu mnuListUnknown 
         Caption         =   "Update Known Db"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'License:   GPL
'Copyright: 2005 iDefense a Verisign Company
'Site:      http://labs.idefense.com
'
'Author:    David Zimmer <david@idefense.com, dzzie@yahoo.com>
'
'         This program is free software; you can redistribute it and/or modify it
'         under the terms of the GNU General Public License as published by the Free
'         Software Foundation; either version 2 of the License, or (at your option)
'         any later version.
'
'         This program is distributed in the hope that it will be useful, but WITHOUT
'         ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or
'         FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for
'         more details.
'
'         You should have received a copy of the GNU General Public License along with
'         this program; if not, write to the Free Software Foundation, Inc., 59 Temple
'         Place, Suite 330, Boston, MA 02111-1307 USA

Dim WithEvents subclass As clsSubClass
Attribute subclass.VB_VarHelpID = -1

Dim liProc As ListItem
Dim liDirWatch As ListItem


Dim tickCount As Long
Dim seconds As Long

Public samplePath As String
Private ignoreAPILOG As Boolean
Private doCopy As Boolean

 

Dim lastViewMode As Integer

Sub Initalize()
    
    Set subclass = New clsSubClass
   
    subclass.AttachMessage Me.Hwnd, WM_COPYDATA            'for process_analyzer
    
    If Me.SSTab1.TabVisible(6) Then
        subclass.AttachMessage frmDirWatch.Hwnd, WM_COPYDATA
    End If
    
    If Me.SSTab1.TabVisible(5) Then
        subclass.AttachMessage frmApiLogger.Hwnd, WM_COPYDATA
    End If
    
    lvDirWatch.ColumnHeaders(3).Width = lvDirWatch.Width - 100 - lvDirWatch.ColumnHeaders(3).Left
    lvAPILog.ColumnHeaders(1).Width = lvAPILog.Width - 100
    
    txtIgnore = GetSetting(App.exename, "Settings", "txtIgnore", "\config\software , modified:, ")
    txtApiIgnore = GetSetting(App.exename, "Settings", "txtApiIgnore", "GetProcAddress, GetModuleHandle, ")

    lastViewMode = -1
    
End Sub

Sub StartCountDown(xSecs As Integer)
    
    seconds = xSecs
    lblTimer = seconds
    Me.Visible = True
    tmrCountDown.Enabled = True
    Unload frmWizard
    lastViewMode = 0
    
End Sub

Sub cmdDirWatch_Click()
    
    With cmdDirWatch
        If Len(.Tag) = 0 Then
            .Tag = "xx"
            DirWatchCtl False
            .Caption = "Start monitor"
        Else
            .Tag = ""
            DirWatchCtl True
            .Caption = "Stop monitor"
        End If
    End With
    
End Sub

Private Sub cmdIgnoreApi_Click()
    
    With cmdIgnoreApi
        If Not ignoreAPILOG Then
            .Caption = "Enable Api logging"
            ignoreAPILOG = True
        Else
            .Caption = "Disable Api Logging"
            ignoreAPILOG = False
        End If
    End With
    
End Sub

Private Sub Form_Load()
    If known.HideKnownInDisplays Then
        mnuHideKnown.Checked = True
        mnuListUnknown.Enabled = True
    End If
    
    Dim alv As ListView, i As Long
    For i = 0 To 6
        Set alv = GetActiveLV(i)
        alv.MultiSelect = True
        alv.HideSelection = False
    Next
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Me.Height = 5805
    Me.Width = 10470
End Sub

Private Sub lblReport_Click()
    ShowDataReport
End Sub

Private Sub lblTools_Click()
     PopupMenu mnuTools
End Sub



Private Sub mnuCopySelected_Click()
    
    Dim active_lv As ListView
    
    Dim i As Integer, tmp As String, match As Long, j As Long
    Dim li As ListItem, search As String, ret() As String
    
    For j = 0 To 6
        Set active_lv = GetActiveLV(j)
        For Each li In active_lv.ListItems
            If li.Selected Then
                tmp = li.Text & vbTab
                For i = 1 To active_lv.ColumnHeaders.Count - 1
                    tmp = tmp & li.SubItems(i) & vbTab
                Next
                li.Selected = True
                match = match + 1
                push ret(), active_lv.name & "> " & tmp
            End If
        Next
    Next
    
    If match > 0 Then
        frmReport.ShowList ret, , "selected_items.txt", False
    End If
    
End Sub

Private Sub mnuHideKnown_Click()
    mnuHideKnown.Checked = Not mnuHideKnown.Checked
    known.HideKnownInDisplays = mnuHideKnown.Checked
    mnuListUnknown.Enabled = mnuHideKnown.Checked
    If lastViewMode >= 0 Then
        mnuToolItem_Click lastViewMode
    End If
End Sub

Private Sub mnuKnownFiles_Click()
    frmKnownFiles.Show 1, Me
End Sub

 



Private Sub mnuSearch_Click()
    
    Dim active_lv As ListView
    
    Dim i As Integer, tmp As String, match As Long, j As Long
    Dim li As ListItem, search As String, ret() As String
    
    search = InputBox("Enter text to search for")
    If Len(search) = 0 Then Exit Sub
    
    For j = 0 To 6
        Set active_lv = GetActiveLV(j)
        For Each li In active_lv.ListItems
            tmp = li.Text & vbTab
            For i = 1 To active_lv.ColumnHeaders.Count - 1
                tmp = tmp & li.SubItems(i) & vbTab
            Next
            If InStr(1, tmp, search, vbTextCompare) > 0 Then
                li.Selected = True
                match = match + 1
                push ret(), active_lv.name & "> " & tmp
            Else
                li.Selected = False
            End If
        Next
    Next
    
    If match > 0 Then
        frmReport.ShowList ret, , "search_result.txt", False
    End If
    
End Sub

Function GetActiveLV(Optional index As Long = -1) As ListView

    Dim active_lv As ListView
    
    If index = -1 Then index = SSTab1.TabIndex
    
    Select Case index
        Case 0: Set active_lv = lvProcesses
        Case 1: Set active_lv = lvPorts
        Case 2: Set active_lv = lvExplorer ' , lvIE
        Case 3: Set active_lv = lvDrivers
        Case 4: Set active_lv = lvRegKeys
        Case 5: Set active_lv = lvAPILog
        Case 6: Set active_lv = lvDirWatch
    End Select
    
    Set GetActiveLV = active_lv
    
End Function

Private Sub tmrCountDown_Timer()
        
    tickCount = tickCount + 1
    If tickCount > seconds Then
        lblTimer.Visible = False
        tmrCountDown.Enabled = False
        'DirWatchCtl False
        'ignoreAPILOG = True
        diff.DoSnap2
        diff.ShowDiffReport
        lastViewMode = 2
        
        frmMain.lblDisplay = "Displaying Snapshot Diff report."
        If known.HideKnownInDisplays Then frmMain.lblDisplay = frmMain.lblDisplay & "  [HIDING TRUSTED FILES]"
        
        If lvProcesses.ListItems.Count < 1 Then
            MsgBox "No new processes detected look at the dlls or it may have exited", vbInformation
        ElseIf lvProcesses.ListItems.Count = 1 Then
            txtProcess = lvProcesses.ListItems(1).Text
            doCopy = True
            cmdAnalyze_Click
            ShowDataReport True
        Else
            MsgBox "Several new processes were detected. " & vbCrLf & vbCrLf & _
                   "Select one from the list and click Analyze " & vbCrLf & _
                   "Process or right click on it to view more options.", vbInformation
        End If
        
    Else
        lblTimer = (seconds - tickCount) & " Seconds remaining"
    End If
    
    
End Sub


Sub cmdAnalyze_Click()
    Dim p As String
    
    If IsIde() Then
        p = """" & App.path & "\..\..\proc_analyzer.exe"" " & txtProcess
    Else
        p = """" & App.path & "\proc_analyzer.exe"" " & txtProcess
    End If
    
    If doCopy Then 'automated from timer
        p = p & " /copy"
        doCopy = False
    End If
    
    On Error GoTo hell
    Shell p, vbNormalFocus
    
    Exit Sub
hell: MsgBox "Error in cmdAnalyze_Click: " & Err.Description, vbInformation
    
End Sub

Function GetClipboard() As String
    GetClipboard = Clipboard.GetText
End Function

Private Sub cmdCopyDll_Click()
    On Error Resume Next
    If Not fso.FileExists(txtDllPath) Then
        MsgBox "File not found"
        Exit Sub
    End If
    FileCopy txtDllPath, UserDeskTopFolder & "\"
    MsgBox "File saved to: " & UserDeskTopFolder, vbInformation
End Sub

Private Sub cmdDelLike_Click()
   
    Dim i As Long
    On Error Resume Next
    
top:
    For i = 1 To lvDirWatch.ListItems.Count
        If InStr(1, lvDirWatch.ListItems(i).Text, txtDeleteLike, vbTextCompare) > 0 Then
           lvDirWatch.ListItems.Remove i
           GoTo top
        End If
    Next
      
End Sub

Private Sub cmdDllProperties_Click()
    On Error Resume Next
    If Not fso.FileExists(txtDllPath) Then
        MsgBox "File not found"
        Exit Sub
    End If
    frmReport.ShowList QuickInfo(txtDllPath)
End Sub


Private Sub cmdSaveDirWatchFile_Click()
    If liDirWatch Is Nothing Then Exit Sub
    
    On Error Resume Next
    Dim f As String, d As String
    
    f = liDirWatch.SubItems(2)
    
    If Not fso.FileExists(f) Then
        MsgBox "File not found: " & f
    Else
        d = UserDeskTopFolder & "\" & fso.FileNameFromPath(f)
        FileCopy f, d
    End If
    
    If Err.Number <> 0 Then
        MsgBox "Error: " & Err.Description
    Else
        MsgBox FileLen(f) & " bytes saved successfully as: " & vbCrLf & vbCrLf & d, vbInformation
    End If
    
End Sub

Private Sub mnuListUnknown_Click()
    
    If lastViewMode < 0 Then Exit Sub
    
    Dim ret() As String
    Dim tmp As String
    
    push ret, GetAllText(lvProcesses, 3)
    push ret, GetAllText(lvPorts, 3)
    push ret, GetAllText(lvExplorer)
    push ret, GetAllText(lvIE)
    push ret, GetAllText(lvDrivers)
    
    tmp = Join(ret, vbCrLf)
    tmp = Replace(tmp, vbCrLf & vbCrLf, vbCrLf)
    ret = Split(tmp, vbCrLf)
    
    frmMarkKnown.loadFiles ret
    frmMarkKnown.Show 1, Me
    
End Sub

Sub ShowDataReport(Optional appendClipboard As Boolean = False)
    
    Dim ret() As String
    
    push ret, "Processes:"
    push ret, GetAllElements(lvProcesses)
    
    push ret, vbCrLf & "Ports:"
    push ret, GetAllElements(lvPorts)
    
    push ret, vbCrLf & "Explorer Dlls:"
    push ret, GetAllElements(lvExplorer)
    
    push ret, vbCrLf & "IE Dlls:"
    push ret, GetAllElements(lvIE)
    
    push ret, vbCrLf & "Loaded Drivers:"
    push ret, GetAllElements(lvDrivers)
    
    push ret, vbCrLf & "Monitored RegKeys"
    push ret, GetAllElements(lvRegKeys)
     
    If SSTab1.TabVisible(5) Then
        push ret, vbCrLf & "Kernel31 Api Log"
        push ret, GetAllElements(lvAPILog)
    End If
    
    If SSTab1.TabVisible(6) Then
        push ret, vbCrLf & "DirwatchData"
        push ret, GetAllElements(lvDirWatch)
    End If
        
    If appendClipboard Then
        push ret, Clipboard.GetText
    End If
    
    frmReport.ShowList Join(ret, vbCrLf)
    
End Sub

 



Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
     
    subclass.DetatchMessage frmDirWatch.Hwnd, WM_COPYDATA
    subclass.DetatchMessage frmApiLogger.Hwnd, WM_COPYDATA
    subclass.DetatchMessage Me.Hwnd, WM_COPYDATA
     
    Unload frmDirWatch
    Unload frmReport
    Unload frmApiLogger
    
End Sub
 
 
Private Sub lvDirWatch_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set liDirWatch = Item
End Sub

Private Sub lvExplorer_ItemClick(ByVal Item As MSComctlLib.ListItem)
    txtDllPath = Item.Text
End Sub
 
Private Sub lvIE_ItemClick(ByVal Item As MSComctlLib.ListItem)
    txtDllPath = Item.Text
End Sub
 


Private Sub mnuToolItem_Click(index As Integer)
    
    'show1, show2, diff, - , take1, take2, - , startover, report
    
    Dim c As String
    
    With diff
        Select Case index
            Case 0: .ShowBaseSnap
            Case 1: .ShowSnap2
            Case 2: .ShowDiffReport
            Case 4: .DoSnap1: .ShowBaseSnap
            Case 5: .DoSnap2: .ShowSnap2
            Case 7:
                    
                    If MsgBox("All current data will be lost continue?", vbExclamation + vbYesNo) = vbNo Then
                        Exit Sub
                    End If
                    
                    On Error Resume Next
                    Shell App.path & "\" & App.exename & ".exe", vbNormalFocus
                    Unload Me
                    
            Case 8: ShowDataReport
            
        End Select
    End With
    
    Select Case index
        Case 0: c = "Showing base snapshot"
        Case 1: c = "Showing snapshot 2"
        Case 2: c = "Showing snapshot diff"
        Case 4: c = "Showing fresh base snap"
        Case 5: c = "Showing fresh snap2"
    End Select
    
    If lastViewMode <= 5 Then
        lastViewMode = index
    Else
        lastViewMode = -1
    End If
    
    If known.HideKnownInDisplays Then c = c & "  [HIDING TRUSTED FILES]"
    frmMain.lblDisplay = c
    
End Sub

 


Private Sub subclass_MessageReceived(Hwnd As Long, wMsg As Long, wParam As Long, lParam As Long, Cancel As Boolean)
    Dim msg As String
    Dim tmp
    Dim li As ListItem
                
    If wMsg = WM_COPYDATA Then
        If RecieveTextMessage(lParam, msg) Then
            If Hwnd = Me.Hwnd Then
            
                If msg = "analyzer_report" Then
                    frmReport.Text1 = frmReport.Text1 & vbCrLf & vbCrLf & _
                                        "Proc_Analyzer Results: " & vbCrLf & _
                                        String(50, "-") & vbCrLf & _
                                        Clipboard.GetText
                    
                    If Not frmReport.Visible Then frmReport.Visible = True
                End If
                
            ElseIf Hwnd = frmApiLogger.Hwnd Then
            
                If ignoreAPILOG Then Exit Sub
                If AnyOfTheseInstr(msg, txtApiIgnore) Then Exit Sub
                If KeyExistsInCollection(cApiData, msg) Then Exit Sub
                On Error Resume Next
                cApiData.Add msg, msg
                lvAPILog.ListItems.Add , , msg
                
            ElseIf wParam = 0 Then 'analyzer report
                
            Else
               
                If AnyOfTheseInstr(msg, txtIgnore) Then Exit Sub
                If KeyExistsInCollection(cLogData, msg) Then Exit Sub
                On Error Resume Next
                cLogData.Add msg, msg
                tmp = Split(msg, ":", 2)
                Set li = lvDirWatch.ListItems.Add(, , tmp(0))
                li.SubItems(2) = Replace(Replace(Trim(tmp(1)), "\\", "\"), Chr(0), Empty)
                li.SubItems(1) = Hex(FileLen(li.SubItems(2)))
                
            End If
        End If
    End If
    
End Sub


 


Private Function RecieveTextMessage(lParam As Long, msg As String) As Boolean
   
    Dim CopyData As COPYDATASTRUCT
    Dim Buffer(1 To 2048) As Byte
    Dim Temp As String
    
    msg = Empty
    
    CopyMemory CopyData, ByVal lParam, Len(CopyData)
    
    If CopyData.dwFlag = 3 Then
        CopyMemory Buffer(1), ByVal CopyData.lpData, CopyData.cbSize
        Temp = StrConv(Buffer, vbUnicode)
        Temp = Left$(Temp, InStr(1, Temp, Chr$(0)) - 1)
        'heres where we work with the intercepted message
        msg = Temp
        RecieveTextMessage = True
    End If
    
End Function
 





Private Sub lvProcesses_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuProcessesPopup
End Sub

Sub lvProcesses_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set liProc = Item
    txtProcess = Item.Tag
End Sub
 


Private Sub mnuDumpProcess_Click()
    If liProc Is Nothing Then Exit Sub

    Dim pth As String
    pth = InputBox("Enter path to save dump as", , "C:\dmp.bin")
    If Len(pth) = 0 Then Exit Sub

    Dim cmod As CModule
    Dim col As Collection
    Dim pid As Long

    pid = CLng(liProc.Tag)
    Set col = diff.CProc.GetProcessModules(pid)
    Set cmod = col(1)

    Call diff.CProc.DumpProcessMemory(pid, cmod.base, cmod.size, pth)

End Sub

Private Sub mnuKillProcess_Click()
    On Error Resume Next
    If liProc Is Nothing Then Exit Sub
    If diff.CProc.TerminateProces(CLng(liProc.Tag)) Then
        lvProcesses.ListItems.Remove liProc.index
        MsgBox "Process Killed", vbInformation
    Else
        MsgBox "Unable to kill Process", vbInformation
    End If
End Sub

Private Sub mnuProcessFileProps_Click()
    
    If liProc Is Nothing Then Exit Sub
    
    Dim path As String
    Dim fsize As String

    On Error Resume Next

    path = liProc.SubItems(3)
    fsize = "FileSize: " & FileLen(path) & vbCrLf & String(70, "-") & vbCrLf

    path = QuickInfo(path)

    frmReport.ShowList fsize & path


End Sub


Private Sub mnuShowProcessDlls_Click()
    If liProc Is Nothing Then Exit Sub

    On Error Resume Next
    
    Dim col As Collection, n, list
    Set col = diff.CProc.GetProcessModules(CLng(liProc.Tag))

    For Each n In col
        list = list & n & vbCrLf
    Next

    frmReport.ShowList list

End Sub



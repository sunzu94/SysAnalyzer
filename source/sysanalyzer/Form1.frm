VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMain 
   Caption         =   "SysAnalyzer"
   ClientHeight    =   5400
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   10830
   Icon            =   "Form1.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "frmMain"
   ScaleHeight     =   5400
   ScaleWidth      =   10830
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox fraTools 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   8940
      ScaleHeight     =   255
      ScaleWidth      =   1275
      TabIndex        =   12
      Top             =   5100
      Width           =   1275
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
         Left            =   720
         TabIndex        =   14
         Top             =   0
         Width           =   435
      End
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
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   615
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5355
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10365
      _ExtentX        =   18283
      _ExtentY        =   9446
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      Tabs            =   7
      TabsPerRow      =   10
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Running Processes"
      TabPicture(0)   =   "Form1.frx":5C12
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lvProcesses"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraProc"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Open Ports"
      TabPicture(1)   =   "Form1.frx":5C2E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lvPorts"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Process Dlls"
      TabPicture(2)   =   "Form1.frx":5C4A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblIEDlls"
      Tab(2).Control(1)=   "Label1(0)"
      Tab(2).Control(2)=   "lvIE"
      Tab(2).Control(3)=   "lvExplorer"
      Tab(2).Control(4)=   "fraDlls"
      Tab(2).Control(5)=   "splitterDlls"
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "Loaded Drivers"
      TabPicture(3)   =   "Form1.frx":5C66
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lvDrivers"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Reg Monitor"
      TabPicture(4)   =   "Form1.frx":5C82
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "lvRegKeys"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Api Log"
      TabPicture(5)   =   "Form1.frx":5C9E
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "fraAPILog"
      Tab(5).Control(1)=   "lvAPILog"
      Tab(5).ControlCount=   2
      TabCaption(6)   =   "Directory Watch Data"
      TabPicture(6)   =   "Form1.frx":5CBA
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "fraDirWatch"
      Tab(6).Control(1)=   "lvDirWatch"
      Tab(6).ControlCount=   2
      Begin VB.Frame fraAPILog 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   915
         Left            =   -74940
         TabIndex        =   36
         Top             =   4020
         Width           =   10215
         Begin VB.TextBox txtApiIgnore 
            Height          =   315
            Left            =   660
            TabIndex        =   40
            Top             =   0
            Width           =   9555
         End
         Begin VB.TextBox txtAPIDelete 
            Height          =   315
            Left            =   660
            TabIndex        =   39
            Top             =   420
            Width           =   4815
         End
         Begin VB.CommandButton cmdApiDelete 
            Caption         =   "Delete Lines Like"
            Height          =   315
            Left            =   5580
            TabIndex        =   38
            Top             =   420
            Width           =   1575
         End
         Begin VB.CommandButton cmdIgnoreApi 
            Caption         =   "Turn off Api Logging"
            Height          =   315
            Left            =   8280
            TabIndex        =   37
            Top             =   420
            Width           =   1935
         End
         Begin VB.Label Label3 
            Caption         =   "Prune"
            Height          =   315
            Index           =   2
            Left            =   60
            TabIndex        =   42
            Top             =   480
            Width           =   555
         End
         Begin VB.Label Label5 
            Caption         =   "Ignore"
            Height          =   315
            Left            =   0
            TabIndex        =   41
            Top             =   60
            Width           =   555
         End
      End
      Begin VB.Frame splitterDlls 
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         Height          =   75
         Left            =   -74940
         MousePointer    =   7  'Size N S
         TabIndex        =   35
         Top             =   2160
         Width           =   10155
      End
      Begin VB.Frame fraDirWatch 
         BorderStyle     =   0  'None
         Height          =   795
         Left            =   -74940
         TabIndex        =   26
         Top             =   4080
         Width           =   10215
         Begin VB.CommandButton cmdSaveDirWatchFile 
            Caption         =   "Save Selected files"
            Height          =   315
            Left            =   7140
            TabIndex        =   31
            Top             =   360
            Width           =   1575
         End
         Begin VB.CommandButton cmdDelLike 
            Caption         =   "Delete Lines Like"
            Height          =   315
            Left            =   5520
            TabIndex        =   30
            Top             =   360
            Width           =   1575
         End
         Begin VB.TextBox txtDeleteLike 
            Height          =   315
            Left            =   660
            TabIndex        =   29
            Top             =   360
            Width           =   4755
         End
         Begin VB.TextBox txtIgnore 
            Height          =   315
            Left            =   660
            TabIndex        =   28
            Top             =   0
            Width           =   9495
         End
         Begin VB.CommandButton cmdDirWatch 
            Caption         =   "Stop Monitor"
            Height          =   315
            Left            =   8940
            TabIndex        =   27
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label3 
            Caption         =   "Prune"
            Height          =   315
            Index           =   1
            Left            =   0
            TabIndex        =   33
            Top             =   360
            Width           =   555
         End
         Begin VB.Label Label3 
            Caption         =   "Ignore"
            Height          =   315
            Index           =   0
            Left            =   0
            TabIndex        =   32
            Top             =   0
            Width           =   555
         End
      End
      Begin VB.Frame fraDlls 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   -74880
         TabIndex        =   21
         Top             =   4500
         Width           =   10035
         Begin VB.CommandButton cmdCopyDll 
            Caption         =   "Save Copy"
            Height          =   315
            Left            =   8160
            TabIndex        =   24
            Top             =   0
            Width           =   1035
         End
         Begin VB.CommandButton cmdDllProperties 
            Caption         =   "Properties"
            Height          =   315
            Left            =   6840
            TabIndex        =   23
            Top             =   0
            Width           =   975
         End
         Begin VB.TextBox txtDllPath 
            Height          =   315
            Left            =   540
            TabIndex        =   22
            Top             =   0
            Width           =   6195
         End
         Begin VB.Label Label7 
            Caption         =   "Path"
            Height          =   255
            Left            =   0
            TabIndex        =   25
            Top             =   60
            Width           =   495
         End
      End
      Begin VB.Frame fraProc 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   120
         TabIndex        =   15
         Top             =   4320
         Width           =   10035
         Begin VB.TextBox txtProcess 
            Height          =   315
            Left            =   1020
            TabIndex        =   17
            Top             =   300
            Width           =   1215
         End
         Begin VB.CommandButton cmdAnalyze 
            Caption         =   "Analyze Process"
            Height          =   315
            Left            =   2340
            TabIndex        =   16
            Top             =   300
            Width           =   1875
         End
         Begin VB.Timer tmrCountDown 
            Enabled         =   0   'False
            Interval        =   1000
            Left            =   9600
            Top             =   180
         End
         Begin MSComctlLib.ProgressBar pb 
            Height          =   255
            Left            =   0
            TabIndex        =   34
            Top             =   0
            Width           =   10155
            _ExtentX        =   17912
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   1
         End
         Begin VB.Label lblDisplay 
            Caption         =   "Currently displaying  base snapshot"
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   6240
            TabIndex        =   18
            Top             =   360
            Width           =   3675
         End
         Begin VB.Label Label4 
            Caption         =   "Analyze PID"
            Height          =   315
            Left            =   0
            TabIndex        =   20
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label lblTimer 
            Caption         =   "Seconds Remaining: "
            Height          =   315
            Left            =   4380
            TabIndex        =   19
            Top             =   360
            Width           =   2955
         End
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
         Height          =   4215
         Left            =   60
         TabIndex        =   1
         Top             =   60
         Width           =   10155
         _ExtentX        =   17912
         _ExtentY        =   7435
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
         TabIndex        =   3
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
         Height          =   1875
         Left            =   -74940
         TabIndex        =   4
         Top             =   2580
         Width           =   10155
         _ExtentX        =   17912
         _ExtentY        =   3307
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
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
         TabIndex        =   7
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
         TabIndex        =   8
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
         Height          =   3975
         Left            =   -75000
         TabIndex        =   9
         Top             =   0
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   7011
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
         Left            =   -74940
         TabIndex        =   10
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
            Text            =   "Size"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "File"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "Explorer Dlls :"
         Height          =   255
         Index           =   0
         Left            =   -74940
         TabIndex        =   6
         Top             =   0
         Width           =   1035
      End
      Begin VB.Label lblIEDlls 
         Caption         =   "IE Dlls :"
         Height          =   255
         Left            =   -74940
         TabIndex        =   5
         Top             =   2340
         Width           =   1035
      End
   End
   Begin VB.OLE OLE1 
      Height          =   30
      Left            =   6720
      TabIndex        =   11
      Top             =   4800
      Width           =   75
   End
   Begin VB.Menu mnuProcessesPopup 
      Caption         =   "mnuProcessesPopup"
      Begin VB.Menu mnuShowProcessDlls 
         Caption         =   "ShowDlls"
      End
      Begin VB.Menu mnuShowMemoryMap 
         Caption         =   "Memory Map"
      End
      Begin VB.Menu mnuScanProcForStealthInjects 
         Caption         =   "RWE Mem Scan"
      End
      Begin VB.Menu mnuDumpProcess 
         Caption         =   "Dump"
      End
      Begin VB.Menu mnuKillProcess 
         Caption         =   "Kill"
      End
      Begin VB.Menu mnuSpacer1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLaunchStrings 
         Caption         =   "Strings"
      End
      Begin VB.Menu mnuCopyProcessPath 
         Caption         =   "Copy File Path"
      End
      Begin VB.Menu mnuProcessFileProps 
         Caption         =   "File Properties"
      End
      Begin VB.Menu mnuSaveToAnalysisFolder 
         Caption         =   "Save to Analysis Folder"
      End
   End
   Begin VB.Menu mnuDllsPopup 
      Caption         =   "mnuDllsPopup"
      Begin VB.Menu mnuAddSelectedDllsToKnown 
         Caption         =   "Add Selected To Known DB"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuViewAllDllProps 
         Caption         =   "View All Properties"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDumpDll 
         Caption         =   "Dump Module"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCopyTo 
         Caption         =   "Copy To"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "mnuTools"
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
      Begin VB.Menu mnuReportViewer 
         Caption         =   "Report Viewer"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuManualTools 
         Caption         =   "Manual Tools"
         Begin VB.Menu mnuScanForUnknownMods 
            Caption         =   "Scan Procs for Unknown Dlls"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuScanProcsForDll 
            Caption         =   "Scan Procs For Dll"
         End
         Begin VB.Menu mnuStealthInjScan 
            Caption         =   "RWE Mem Scan"
         End
      End
   End
   Begin VB.Menu mnuDriversPopup 
      Caption         =   "mnuDriversPopup"
      Begin VB.Menu mnuSaveDriver 
         Caption         =   "Save File"
      End
      Begin VB.Menu mnuAddSelDrivertoKnownDB 
         Caption         =   "Add Selected To Known DB"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuRegMonitor 
      Caption         =   "mnuRegMonitor"
      Begin VB.Menu mnuRegMonCopyLine 
         Caption         =   "Copy Selected Line"
      End
      Begin VB.Menu mnuRegMonCopyTable 
         Caption         =   "Copy Entire Table"
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
'
'todo (11/2013):
'      make sure CreateProcess Apilogger hook is crash free..should switch over to CreateProcessInternalW, have it somewhere..
'      TEST WITH WIN7 AND WIN7X64
'      ability to runas another user? (explorer injection, screen lockers etc)
'      make sure IE process running at start of test?
'      Cstrings - be able to set min match leng

'X      show known db status on wizard, allow to build from there.
'X      speed up delay in ShowBaseSnapshot before malware launch if known db active
'X      remove or build in clsAdoKit --> remove annoying msgbox if fail, and screen pointer change
'X      create a verbose runlog with debug info in case of crash...
'X      analyze multiple processes if found
'X      apparent bug in some known files lookups? - was windows update patches being detected.
'X      integrate virustotal results with report viewer? - tricky because of junk files and delay required, plus all the network traffic..nevermind
'X      auto run RWE memory scan on new processes
'X      break reporting into System level and process level
'X      have dirwatch auto save files (code already in dirwatch.exe)
'X      savereport can not overwrite old reports...
'X      progress bar for snapshots, and show main form while loading..
'X      build analyzze process into main codebase, each process to seperate sub folders
'X      pull network data in from sniffhit (fix 127.0.0.1 bug)
'X      way to add files to unknown db from right click on lvExplorer/lvIEdlls/drivers
'X      splitter for lvExplorer/lvIEdlls
'X      import CStrings.filter, append filter results to bottom?
'X      import ShellExt.More external apps menu capabilities to reportviewer.tv.mnupopup
'X      resize code for apilog window

Private Capturing As Boolean
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Dim WithEvents subclass As CSubclass2
Attribute subclass.VB_VarHelpID = -1

Dim liProc As ListItem
Dim liDirWatch As ListItem
Dim liDriver As ListItem
Dim liRegMon As ListItem

Dim tickCount As Long
Dim seconds As Long

Public samplePath As String
Private ignoreAPILOG As Boolean
Private doCopy As Boolean
Private user_desktop As String

Dim lastViewMode As Integer

Property Let Display(msg)
    On Error Resume Next
    lblDisplay.Caption = msg
    lblDisplay.Refresh
    DoEvents
    debugLog msg
End Property

Sub Initalize()
    
    user_desktop = UserDeskTopFolder()
    
    Set subclass = New CSubclass2
   
    'subclass.AttachMessage Me.hwnd, WM_COPYDATA            'for process_analyzer
    
    If Me.SSTab1.TabVisible(6) Then
        If Not isIde() Then 'already debugged
            subclass.AttachMessage frmDirWatch.hWnd, WM_COPYDATA
        End If
    End If
    
    If Me.SSTab1.TabVisible(5) Then
        If Not isIde() Then 'already debugged
            subclass.AttachMessage frmApiLogger.hWnd, WM_COPYDATA
        End If
    End If
    
    lvDirWatch.ColumnHeaders(3).Width = lvDirWatch.Width - 100 - lvDirWatch.ColumnHeaders(3).Left
    lvAPILog.ColumnHeaders(1).Width = lvAPILog.Width - 100
    
    txtIgnore = GetSetting(App.exename, "Settings", "txtIgnore", "\config\software , modified:, ")
    txtApiIgnore = GetSetting(App.exename, "Settings", "txtApiIgnore", "GetProcAddress, GetModuleHandle, ")

    lastViewMode = -1
    debugLog "frmMain.Initilized"
    
End Sub

Sub StartCountDown(xSecs As Integer)
    
    seconds = xSecs
    lblTimer = seconds & " Seconds remaining"
    debugLog lblTimer.Caption
    
    Me.Visible = True
    tmrCountDown.Enabled = True
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
    
    On Error Resume Next
    
    mnuDriversPopup.Visible = False
    mnuProcessesPopup.Visible = False
    mnuDllsPopup.Visible = False
    mnuTools.Visible = False
    mnuRegMonitor.Visible = False
    
    If known.Loaded And known.Ready Then
        mnuAddSelectedDllsToKnown.Enabled = True
        mnuScanForUnknownMods.Enabled = True
        mnuAddSelDrivertoKnownDB.Enabled = True
    End If
    
    If known.HideKnownInDisplays Then
        mnuHideKnown.Checked = True
        mnuListUnknown.Enabled = True 'this gets all displayed dlls automatically, only makes sense with hideknown enabled..
    End If
    
    Dim alv As ListView, i As Long
    For i = 0 To 6
        Set alv = GetActiveLV(i)
        alv.MultiSelect = True
        alv.HideSelection = False
    Next
    
    RestoreFormSizeAnPosition Me
    splitterDlls.top = SSTab1.Height / 2
    Splitter_DoMove
    
    On Error Resume Next
    SSTab1.TabIndex = 1
    debugLog "frmMain_Load"
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next 'took me 7 yrs but i finally added form resize code!
    'Me.Height = 5925
    'Me.Width = 10470
      
    Dim o As Object
    Dim lv As ListView
    SSTab1.Width = Me.Width - SSTab1.Left - 100
    SSTab1.Height = Me.Height - SSTab1.top - 500
    For Each o In Me.Controls
        If TypeName(o) = "ListView" Then
            Set lv = o
            lv.Width = SSTab1.Width - 200
            lv.ColumnHeaders(lv.ColumnHeaders.count).Width = lv.Width - lv.ColumnHeaders(lv.ColumnHeaders.count).Left - 200
            If lv.name = "lvPorts" Or lv.name = "lvDrivers" Or lv.name = "lvRegKeys" Then
                With lv
                    .Height = SSTab1.Height - .top - 500
                End With
            End If
        End If
    Next
    
    fraTools.top = SSTab1.Height - 200
    fraTools.Left = SSTab1.Width - 200 - fraTools.Width
    
    With lvProcesses
        .Height = SSTab1.Height - .top - fraProc.Height - 500
        fraProc.top = .top + .Height + 100
        fraProc.Width = .Width
        pb.Width = .Width
    End With
    
    With lvIE
        .Height = SSTab1.Height - .top - fraDlls.Height - 500
        fraDlls.top = .top + .Height + 100
    End With
    
    With lvDirWatch
        .Height = SSTab1.Height - .top - fraDirWatch.Height - 500
        fraDirWatch.top = .top + .Height + 100
    End With
    
    With lvAPILog
        .Height = SSTab1.Height - .top - fraAPILog.Height - 500
        fraAPILog.top = .top + .Height + 100
    End With
    
    splitterDlls.Width = lvIE.Width
    
    Me.Refresh
End Sub

Private Sub lblReport_Click()
    ShowDataReport
End Sub

Private Sub lblTools_Click()
     PopupMenu mnuTools
End Sub

Private Sub lvDrivers_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set liDriver = Item
End Sub

Private Sub lvDrivers_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = 2 Then PopupMenu mnuDriversPopup
End Sub

Private Sub lvExplorer_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = 2 Then PopupMenu mnuDllsPopup
End Sub

Private Sub lvIE_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = 2 Then PopupMenu mnuDllsPopup
End Sub

Private Sub lvRegKeys_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set liRegMon = Item
End Sub

Private Sub lvRegKeys_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = 2 Then PopupMenu mnuRegMonitor
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show 1, Me
End Sub

Private Sub mnuAddSelDrivertoKnownDB_Click()
    Dim ret() As String
    Dim tmp As String
    
    push ret, GetAllText(lvDrivers, , True)
    
    tmp = Join(ret, vbCrLf)
    tmp = Replace(tmp, vbCrLf & vbCrLf, vbCrLf)
    ret = Split(tmp, vbCrLf)
    
    frmMarkKnown.loadFiles ret
    frmMarkKnown.Show 1, Me
    
End Sub

Private Sub mnuAddSelectedDllsToKnown_Click()

    Dim ret() As String
    Dim tmp As String
    
    push ret, GetAllText(lvExplorer, , True)
    push ret, GetAllText(lvIE, , True)
    
    tmp = Join(ret, vbCrLf)
    tmp = Replace(tmp, vbCrLf & vbCrLf, vbCrLf)
    ret = Split(tmp, vbCrLf)
    
    frmMarkKnown.loadFiles ret
    frmMarkKnown.Show 1, Me
    
End Sub

Private Sub mnuCopySelected_Click()
    
    Dim active_lv As ListView
    
    Dim i As Integer, tmp As String, match As Long, j As Long
    Dim li As ListItem, Search As String, ret() As String
    
    For j = 0 To 6
        Set active_lv = GetActiveLV(j)
        For Each li In active_lv.ListItems
            If li.Selected Then
                tmp = li.Text & vbTab
                For i = 1 To active_lv.ColumnHeaders.count - 1
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
    On Error Resume Next
    frmKnownFiles.Show 1, Me
End Sub

 



Private Sub mnuLaunchStrings_Click()
    If liProc Is Nothing Then Exit Sub
    Dim f As String
    On Error Resume Next
    f = liProc.SubItems(3)
    LaunchStrings f, True
End Sub

Private Sub mnuRegMonCopyLine_Click()
    If liRegMon Is Nothing Then Exit Sub
    On Error Resume Next
    Dim tmp As String
    tmp = liRegMon.Text & vbCrLf & vbTab & liRegMon.SubItems(1)
    Clipboard.Clear
    Clipboard.SetText tmp
End Sub

Private Sub mnuRegMonCopyTable_Click()
    Dim tmp As String
    tmp = GetAllElements(lvRegKeys)
    Clipboard.Clear
    Clipboard.SetText tmp
End Sub

Private Sub mnuReportViewer_Click()
    frmReportViewer.OpenAnalysisFolder UserDeskTopFolder
End Sub

Private Sub mnuSaveDriver_Click()
    If liDriver Is Nothing Then Exit Sub
    Dim p As String
    On Error Resume Next
    p = liDriver.Text
    If Not fso.FileExists(p) Then
        MsgBox "Could not find: " & p
        Exit Sub
    End If
    FileCopy p, UserDeskTopFolder & "\" & fso.FileNameFromPath(p)
    If Err.Number <> 0 Then
        MsgBox "Failed to save file " & Err.Description, vbExclamation
    Else
        MsgBox "File saved to desktop analysis folder", vbInformation
    End If
End Sub

Private Sub mnuSaveToAnalysisFolder_Click()
    If liProc Is Nothing Then Exit Sub
    Dim f As String, f2 As String
    On Error Resume Next
    
    f = liProc.SubItems(3)
    If Not fso.FileExists(f) Then
        MsgBox "File not found: " & f
    Else
        f2 = UserDeskTopFolder & "\" & fso.FileNameFromPath(f)
        If fso.FileExists(f2) Then Kill f2
        fso.Copy f, UserDeskTopFolder
        If Not fso.FileExists(f2) Then
            MsgBox "File copy failed..."
        Else
            MsgBox "File copied!"
        End If
    End If
    
End Sub

Private Sub mnuScanForUnknownMods_Click()
    Dim cp As New CProcessInfo
    Dim c As Collection
    Dim m As Collection
    Dim p As CProcess
    Dim cm As CModule
    Dim ret As String
    Dim i As Long
    Dim hit As Boolean
    Dim tmp As String
    Dim tmp2 As String
    Dim mm As matchModes
    
    'On Error Resume Next
    
    If Not known.Loaded Then
        MsgBox "Known database is not loaded..", vbInformation
        Exit Sub
    End If
    
    'ado.OpenConnection
    lblDisplay.Caption = "Starting scan..."
    
    i = 0
    
    Set c = cp.GetRunningProcesses()
    For Each p In c
        If p.pid <> 0 And p.pid <> 4 Then
            lblDisplay.Caption = "Scanning " & i & "/" & c.count
            Set m = cp.GetProcessModules(p.pid)
            If Not m Is Nothing And m.count > 0 Then
                tmp = "pid: " & p.pid & " " & p.path
                hit = False
                tmp2 = Empty
                For Each cm In m
                    mm = known.isFileKnown(cm.path)
                    If mm <> exact_match Then
                       tmp2 = tmp2 & vbCrLf & vbTab & IIf(mm = not_found, "Unknown Mod: ", "Hash Changed: ") & cm.path
                       hit = True
                    End If
                Next
                If hit Then ret = ret & tmp & tmp2 & vbCrLf & vbCrLf
            End If
            i = i + 1
            DoEvents
            lblDisplay.Refresh
        End If
    Next
    
    lblDisplay.Caption = ""
    'ado.CloseConnection
    
    Const header = "This list may also include files were locked at the time the database was created and could not be hashed for that reason."
    
    If Len(ret) > 0 Then
        frmReport.ShowList vbCrLf & header & vbCrLf & vbCrLf & Replace(ret, Chr(0), Empty)
    Else
        MsgBox "No unknown modules found in any process..."
    End If
    
    
End Sub

Private Sub mnuScanProcForStealthInjects_Click()
    If liProc Is Nothing Then Exit Sub
    Dim pid As Long
    pid = CLng(liProc.Tag)
    If diff.CProc.x64.IsProcess_x64(pid) <> r_32bit Then
        MsgBox x64Error, vbInformation
        Exit Sub
    End If
    frmInjectionScan.FindStealthInjections pid, liProc.SubItems(3)
End Sub

Private Sub mnuScanProcsForDll_Click()
    Dim cp As New CProcessInfo
    Dim c As Collection
    Dim m As Collection
    Dim p As CProcess
    Dim cm As CModule
    Dim ret As String
    Dim i As Long
    Dim hit As Boolean
    Dim tmp As String
    Dim tmp2 As String
    Dim mm As matchModes
    
    'On Error Resume Next
    
    Dim find As String
    find = InputBox("Enter string fragment of what to look for in dll name or path.")
    If Len(find) = 0 Then Exit Sub
    
    lblDisplay.Caption = "Starting scan..."
    
    i = 0
    
    Set c = cp.GetRunningProcesses()
    For Each p In c
        If p.pid <> 0 And p.pid <> 4 Then
            lblDisplay.Caption = "Scanning " & i & "/" & c.count
            lblDisplay.Refresh
            DoEvents
            Set m = cp.GetProcessModules(p.pid)
            If Not m Is Nothing Then
                If m.count > 0 Then
                    tmp = "pid: " & p.pid & " " & p.path
                    hit = False
                    tmp2 = Empty
                    For Each cm In m
                        If InStr(1, cm.path, find, vbTextCompare) > 0 Then
                           tmp2 = tmp2 & vbTab & Hex(cm.Base) & vbTab & cm.path & vbCrLf
                           hit = True
                        End If
                    Next
                    If hit Then ret = ret & tmp & tmp2
                End If
            End If
            i = i + 1
        End If
    Next
    
    lblDisplay.Caption = ""
    
    If Len(ret) > 0 Then
        frmReport.ShowList vbCrLf & Replace(ret, Chr(0), Empty)
    Else
        MsgBox "No modules found in any process matching your criteria"
    End If
    

End Sub

Private Sub mnuSearch_Click()
    
    Dim active_lv As ListView
    
    Dim i As Integer, tmp As String, match As Long, j As Long
    Dim li As ListItem, Search As String, ret() As String
    
    Search = InputBox("Enter text to search for")
    If Len(Search) = 0 Then Exit Sub
    
    For j = 0 To 6
        Set active_lv = GetActiveLV(j)
        For Each li In active_lv.ListItems
            tmp = li.Text & vbTab
            For i = 1 To active_lv.ColumnHeaders.count - 1
                tmp = tmp & li.SubItems(i) & vbTab
            Next
            If InStr(1, tmp, Search, vbTextCompare) > 0 Then
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

Function GetActiveLV(Optional Index As Long = -1) As ListView

    Dim active_lv As ListView
    
    If Index = -1 Then Index = SSTab1.TabIndex
    
    Select Case Index
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

Private Sub mnuShowMemoryMap_Click()
    If liProc Is Nothing Then Exit Sub
    Dim pid As Long
    pid = CLng(liProc.Tag)
    If diff.CProc.x64.IsProcess_x64(pid) <> r_32bit Then
        MsgBox x64Error, vbInformation
        Exit Sub
    End If
    frmMemoryMap.ShowMemoryMap pid
End Sub

Private Sub mnuStealthInjScan_Click()
    frmInjectionScan.StealthInjectionScan
End Sub


Private Sub tmrCountDown_Timer()
    
    On Error Resume Next
    Dim li As ListItem
    Dim ret() As String
    
    tickCount = tickCount + 1
    If tickCount > seconds Then
        lblTimer.Visible = False
        tmrCountDown.Enabled = False
    
        diff.DoSnap2
        diff.ShowDiffReport
        lastViewMode = 2
        
        frmMain.lblDisplay = "Displaying Snapshot Diff report."
        If known.HideKnownInDisplays Then frmMain.lblDisplay = frmMain.lblDisplay & "  [HIDING TRUSTED FILES]"
        
        frmAnalyzeProcess.ClearList
        
        For Each li In lvProcesses.ListItems
            frmAnalyzeProcess.AnalyzeProcess CLng(li.Text)
        Next
        
        debugLog "AnalyzeKnownProcessesforRWE(" & ProcessesToRWEScan & ")"
        frmAnalyzeProcess.AnalyzeKnownProcessesforRWE ProcessesToRWEScan '"explorer.exe,iexplore.exe,"
        Unload frmAnalyzeProcess
        
        ret() = GetSystemDataReport()
        
        If lvProcesses.ListItems.count < 1 Then
            ret(0) = ret(0) & "No new processes detected look at the dlls or it may have exited" & vbCrLf & vbCrLf
        End If
        
        fso.writeFile UserDeskTopFolder & "\Report_" & Format(Now(), "h.nam/pm") & ".txt", Join(ret, vbCrLf)
        frmReportViewer.OpenAnalysisFolder UserDeskTopFolder
        
        
    Else
        lblTimer = (seconds - tickCount) & " Seconds remaining"
    End If
    
    
End Sub

Sub cmdAnalyze_Click()
    
    On Error GoTo hell
    frmAnalyzeProcess.AnalyzeProcess CLng(txtProcess)
    Unload frmAnalyzeProcess
    
    frmReportViewer.OpenAnalysisFolder UserDeskTopFolder
    
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
    For i = 1 To lvDirWatch.ListItems.count
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
    
    On Error Resume Next
    Dim f As String, d As String
    
    Dim li As ListItem
    Dim tmp() As String
    Dim pFolder As String
    
    pFolder = UserDeskTopFolder & "\DirWatch"
    If Not fso.FolderExists(pFolder) Then MkDir pFolder
    
    For Each li In lvDirWatch.ListItems
    
        If li.Selected Then
        
            f = li.SubItems(2)
            Err.Clear
            
            If Not fso.FileExists(f) Then
                push tmp(), "File not found: " & f
            Else
                d = pFolder & "\" & fso.FileNameFromPath(f)
                FileCopy f, d
            End If
            
            If Err.Number <> 0 Then
                push tmp, "Error saving file: " & Err.Description
            Else
                push tmp(), FileLen(f) & " bytes saved successfully as: " & d
            End If
        
        End If
        
    Next
    
    frmReport.ShowList tmp
    
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

Function GetSystemDataReport(Optional appendClipboard As Boolean = False) As String()

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
    
    GetSystemDataReport = ret()
    
End Function

Sub ShowDataReport(Optional appendClipboard As Boolean = False)
    
    Dim ret() As String
    ret() = GetSystemDataReport(appendClipboard)
    frmReport.ShowList Join(ret, vbCrLf)
    
End Sub

 



Private Sub Form_Unload(Cancel As Integer)

    On Error Resume Next
    Dim f
    
    SaveFormSizeAnPosition Me
    tmrCountDown.Enabled = False
    'subclass.DetatchMessage Me.hwnd, WM_COPYDATA   'for process analyzer
     
    If Me.SSTab1.TabVisible(6) Then
        If Not isIde() Then
            subclass.DetatchMessage frmDirWatch.hWnd, WM_COPYDATA
        End If
    End If
    
    If Me.SSTab1.TabVisible(5) Then
        If Not isIde() Then
            subclass.DetatchMessage frmApiLogger.hWnd, WM_COPYDATA
        End If
    End If
    
    Set subclass = Nothing
    
    For Each f In Forms
        Unload f
    Next
    
    diff.shutDown = True

    Set fso = Nothing
    Set dlg = Nothing
    Set hash = Nothing
    Set diff = Nothing
    Set known = Nothing
    Set ado = Nothing
    Set apiDataManager = Nothing
    
    Unload Me
    End
    
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
 


Public Sub mnuToolItem_Click(Index As Integer)
    
    'show1, show2, diff, - , take1, take2, - , startover, report
    
    Dim c As String
    
    With diff
        Select Case Index
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
                    If isIde() Then
                        Shell App.path & "\..\..\sysanalyzer.exe" & " """ & samplePath & """", vbNormalFocus
                    Else
                        Shell App.path & "\sysanalyzer.exe" & " """ & samplePath & """", vbNormalFocus
                    End If
                    Unload Me
                    
            Case 8: ShowDataReport
            
        End Select
    End With
    
    Select Case Index
        Case 0: c = "Showing base snapshot"
        Case 1: c = "Showing snapshot 2"
        Case 2: c = "Showing snapshot diff"
        Case 4: c = "Showing fresh base snap"
        Case 5: c = "Showing fresh snap2"
    End Select
    
    If lastViewMode <= 5 Then
        lastViewMode = Index
    Else
        lastViewMode = -1
    End If
    
    If known.HideKnownInDisplays Then c = c & "  [HIDING TRUSTED FILES]"
    frmMain.lblDisplay = c
    
End Sub

 


Private Sub subclass_MessageReceived(hWnd As Long, wMsg As Long, wParam As Long, lParam As Long, Cancel As Boolean)
    Dim msg As String
    Dim tmp
    Dim li As ListItem
    Dim size As Long
    Dim saved As Boolean
    
    On Error Resume Next
    
    If wMsg = WM_COPYDATA Then
        If RecieveTextMessage(lParam, msg) Then
            If hWnd = Me.hWnd Then
            
'                If msg = "analyzer_report" Then
'                    frmReport.Text1 = frmReport.Text1 & vbCrLf & vbCrLf & _
'                                        "Proc_Analyzer Results: " & vbCrLf & _
'                                        String(50, "-") & vbCrLf & _
'                                        Clipboard.GetText
'
'                    If Not frmReport.Visible Then frmReport.Visible = True
'                End If
                
            ElseIf hWnd = frmApiLogger.hWnd Then
            
                apiDataManager.HandleApiMessage msg '5.18.12
                If ignoreAPILOG Then Exit Sub
                If AnyOfTheseInstr(msg, txtApiIgnore) Then Exit Sub
                If KeyExistsInCollection(cApiData, msg) Then Exit Sub 'some antispam..
                On Error Resume Next
                cApiData.Add msg, msg
                lvAPILog.ListItems.Add , , msg
                
            ElseIf wParam = 0 Then 'analyzer report
                
            Else
                'directory watch info coming in...
                msg = Trim(msg)
                If InStr(1, msg, "C:\\") > 0 Then msg = Replace(msg, "\\", "\")
                If InStr(1, msg, Chr(0)) > 0 Then msg = Replace(msg, Chr(0), Empty) 'standardize data first
                
                'hardcoded filters
                If LCase(Right(msg, 4)) = ".lnk" Then Exit Sub
                If InStr(1, msg, "C:\iDEFENSE\SysAnalyzer", vbTextCompare) > 0 Then Exit Sub
                If InStr(1, msg, user_desktop, vbTextCompare) > 0 Then Exit Sub
                If InStr(1, msg, "\Prefetch\") > 0 Then Exit Sub
                If InStr(1, msg, "NTUSER.DAT") > 0 Then Exit Sub
                If InStr(1, msg, "C:\LOG.TXT") > 0 Then Exit Sub
                If InStr(1, msg, "\Config\SYSTEM.LOG", vbTextCompare) > 0 Then Exit Sub
                If InStr(msg, "git_shell_ext_debug.txt") > 0 Then Exit Sub
                If InStr(msg, "desktop.ini") > 0 Then Exit Sub
                
                If AnyOfTheseInstr(msg, txtIgnore) Then Exit Sub 'user filters
                If KeyExistsInCollection(cLogData, msg) Then Exit Sub 'antispam
                                    
                On Error Resume Next 'logging
                cLogData.Add msg, msg
                tmp = Split(msg, ":", 2) 'format=  action:file
                tmp(1) = Trim(tmp(1))
                
                If fso.FileExists(CStr(tmp(1))) Then 'auto save file to [analysis]\DirWatch\
                    size = FileLen(CStr(tmp(1)))
                    If InStr(1, tmp(0), "mod", vbTextCompare) > 0 Then
                        saved = SafeFileCopy(CStr(tmp(1)), "DirWatch")
                    End If
                End If
                    
                Set li = lvDirWatch.ListItems.Add(, , tmp(0))
                If size > 0 Then li.SubItems(1) = Hex(size) & IIf(saved, " +", "")
                li.SubItems(2) = Trim(tmp(1))
                li.EnsureVisible
                
            End If
        End If
    End If
    
End Sub

Function SafeFileCopy(org As String, subfolder As String) As Boolean
    On Error Resume Next
    Dim p As String, i As Long, f As String
    Dim tmp
    Dim size As Long
    
    i = 1
    p = UserDeskTopFolder & "\" & subfolder & "\"
    If Not fso.FolderExists(p) Then fso.buildPath p
    
    size = FileLen(org)
    If size = 0 Then Exit Function
    
    f = fso.FileNameFromPath(org)
    
    tmp = f
    While fso.FileExists(p & "\" & tmp)
        tmp = f & "_" & i
        i = i + 1
    Wend
    
    Err.Clear
    FileCopy org, p & "\" & tmp
    
    If Err.Number = 0 Then
        'Debug.Print "Auto Saved: " & org & " -> " & p & "\" & tmp
        SafeFileCopy = True
    Else
         Debug.Print Err.Description & " : org: " & org & " -> " & p & "\" & tmp
    End If
    
End Function

 


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
 





Private Sub lvProcesses_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = 2 Then PopupMenu mnuProcessesPopup
End Sub

Sub lvProcesses_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set liProc = Item
    txtProcess = Item.Tag
End Sub
 
Private Sub mnuDumpProcess_Click()
    If liProc Is Nothing Then Exit Sub

    'MsgBox dlg.SaveDialog(AllFiles) 'threadlocks for unknown reason...

    Dim pid As Long

    pid = CLng(liProc.Tag)
    If diff.CProc.x64.IsProcess_x64(pid) <> r_32bit Then
        MsgBox x64Error, vbInformation
        Exit Sub
    End If
    
    Dim pth As String
    pth = fso.FileNameFromPath(liProc.SubItems(3)) & ".dmp"
    'pth = InputBox("Enter path to dump file as:", , UserDeskTopFolder & "\" & pth)
    pth = frmDlg.SaveDialog(AllFiles, UserDeskTopFolder, "Save Dump as", , Me, pth)
    If Len(pth) = 0 Then Exit Sub

    diff.CProc.DumpProcess pid, pth 'x64 safe...
    
'    Dim cmod As CModule
'    Dim col As Collection
'
'    Set col = diff.CProc.GetProcessModules(pid)
'    Set cmod = col(1)
'
'
'    Call diff.CProc.DumpProcessMemory(pid, cmod.Base, cmod.size, pth)

End Sub

Private Sub mnuCopyProcessPath_Click()
    On Error Resume Next
    If liProc Is Nothing Then Exit Sub
    Dim pth As String
    pth = liProc.SubItems(3)
    Clipboard.Clear
    Clipboard.SetText pth
End Sub

Private Sub mnuKillProcess_Click()
    On Error Resume Next
    If liProc Is Nothing Then Exit Sub
    If diff.CProc.TerminateProces(CLng(liProc.Tag)) Then
        lvProcesses.ListItems.Remove liProc.Index
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
    Dim pid As Long
    pid = CLng(liProc.Tag)
    
    frmMemoryMap.ShowDlls pid
    
    'Dim col As Collection, n, list
    'Set col = diff.CProc.GetProcessModules(CLng(liProc.Tag))
    '
    'For Each n In col
    '    list = list & n & vbCrLf
    'Next
    '
    'frmReport.ShowList list

    
    
    
End Sub

Private Sub lvdrivers_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    LV_ColumnSort lvDrivers, ColumnHeader
End Sub

Private Sub lvPorts_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    LV_ColumnSort lvPorts, ColumnHeader
End Sub

 Private Sub lvProcesses_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    LV_ColumnSort lvProcesses, ColumnHeader
End Sub


'splitter code
'------------------------------------------------
Private Sub splitterDlls_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    Dim a1&

    If Button = 1 Then 'The mouse is down
        If Capturing = False Then
            splitterDlls.ZOrder
            SetCapture splitterDlls.hWnd
            Capturing = True
        End If
        With splitterDlls
            a1 = .top + y
            If MoveOk(a1) Then
                .top = a1
            End If
        End With
    End If
End Sub

Private Sub splitterDlls_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Capturing Then
        ReleaseCapture
        Capturing = False
        Splitter_DoMove
    End If
End Sub


Private Sub Splitter_DoMove()
    On Error Resume Next
    Const buf = 30
    lvIE.top = splitterDlls.top + splitterDlls.Height + buf + lblIEDlls.Height + buf
    lblIEDlls.top = lvIE.top - buf - lblIEDlls.Height
    lvExplorer.Height = splitterDlls.top - buf - lvExplorer.top
    Form_Resize
End Sub


Private Function MoveOk(y&) As Boolean  'Put in any limiters you desire
    MoveOk = False
    If y > 2000 And y < SSTab1.Height - 3000 Then
        MoveOk = True
    End If
End Function

'------------------------------------------------
'end splitter code



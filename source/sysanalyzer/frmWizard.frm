VERSION 5.00
Begin VB.Form frmWizard 
   BackColor       =   &H005A5963&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SysAnalyzer Configuration Wizard"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   720
   ClientWidth     =   8955
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   8955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   8460
      TabIndex        =   24
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox txtArgs 
      Height          =   285
      Left            =   4350
      TabIndex        =   18
      Top             =   570
      Width           =   3975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H005A5963&
      Caption         =   " Options "
      ForeColor       =   &H00E0E0E0&
      Height          =   2865
      Left            =   3840
      TabIndex        =   7
      Top             =   1290
      Width           =   5025
      Begin VB.TextBox txtRWEScan 
         Height          =   315
         Left            =   1440
         TabIndex        =   20
         Text            =   "explorer.exe,iexplore.exe"
         Top             =   2400
         Width           =   3435
      End
      Begin VB.ComboBox cboIp 
         Height          =   315
         Left            =   1140
         TabIndex        =   16
         Top             =   1890
         Width           =   2475
      End
      Begin VB.TextBox txtInterface 
         Height          =   285
         Left            =   2610
         TabIndex        =   13
         Text            =   "1"
         Top             =   1590
         Width           =   405
      End
      Begin VB.CheckBox chkPacketCapture 
         BackColor       =   &H005A5963&
         Caption         =   "Full Packet Capture"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   480
         TabIndex        =   11
         Top             =   1320
         Width           =   1755
      End
      Begin VB.CheckBox chkNetworkAnalyzer 
         BackColor       =   &H005A5963&
         Caption         =   "Use SniffHit"
         ForeColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   480
         TabIndex        =   10
         Top             =   240
         Width           =   2775
      End
      Begin VB.CheckBox chkApiLog 
         BackColor       =   &H005A5963&
         Caption         =   "Use Api Logger"
         ForeColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   480
         TabIndex        =   9
         Top             =   570
         Width           =   1455
      End
      Begin VB.CheckBox chkWatchDirs 
         BackColor       =   &H005A5963&
         Caption         =   "Use Directory Watcher"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   480
         TabIndex        =   8
         Top             =   960
         Width           =   2835
      End
      Begin VB.Label Label3 
         BackColor       =   &H005A5963&
         Caption         =   "RWE Scan:"
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   420
         TabIndex        =   19
         Top             =   2460
         Width           =   915
      End
      Begin VB.Label lblip 
         BackColor       =   &H005A5963&
         Caption         =   "IP"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   840
         TabIndex        =   15
         Top             =   1950
         Width           =   285
      End
      Begin VB.Label lblLaunchTcpDump 
         BackColor       =   &H005A5963&
         Caption         =   "launch now"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   2820
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   14
         Top             =   1320
         Width           =   915
      End
      Begin VB.Label lblInterfaces 
         BackColor       =   &H005A5963&
         Caption         =   "Interface Index: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   1
         Left            =   1380
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   12
         Top             =   1620
         Width           =   1245
      End
   End
   Begin VB.TextBox txtDelay 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4320
      TabIndex        =   5
      Text            =   "3"
      Top             =   930
      Width           =   555
   End
   Begin VB.Timer tmrDelayShell 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   2820
      Top             =   2580
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   375
      Left            =   7740
      TabIndex        =   3
      Top             =   4320
      Width           =   1155
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   8460
      TabIndex        =   2
      Top             =   210
      Width           =   375
   End
   Begin VB.TextBox txtBinary 
      Height          =   285
      Left            =   4320
      OLEDropMode     =   1  'Manual
      TabIndex        =   1
      Top             =   180
      Width           =   4005
   End
   Begin VB.Label lblAdmin 
      BackColor       =   &H005A5963&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   345
      Left            =   270
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   29
      Top             =   4365
      Width           =   5625
   End
   Begin VB.Label lblDisplay 
      BackColor       =   &H005A5963&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   1290
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   28
      Top             =   3990
      Width           =   2295
   End
   Begin VB.Label cmdTools 
      BackColor       =   &H005A5963&
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
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   150
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   27
      Top             =   3990
      Width           =   675
   End
   Begin VB.Label cmdAbout 
      BackColor       =   &H005A5963&
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   120
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   26
      Top             =   3300
      Width           =   675
   End
   Begin VB.Label cmdReadme 
      BackColor       =   &H005A5963&
      Caption         =   "Help file"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   120
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   25
      Top             =   3660
      Width           =   675
   End
   Begin VB.Label lblKnown 
      BackColor       =   &H005A5963&
      Caption         =   "lblKnown"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   6360
      TabIndex        =   23
      Top             =   1020
      Width           =   975
   End
   Begin VB.Label lblBuildKnownFileDB 
      BackColor       =   &H005A5963&
      Caption         =   "build now"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   7680
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   22
      Top             =   1020
      Width           =   675
   End
   Begin VB.Label Label1 
      BackColor       =   &H005A5963&
      Caption         =   "Known file DB :"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   1
      Left            =   5160
      TabIndex        =   21
      Top             =   1020
      Width           =   1155
   End
   Begin VB.Label Label1 
      BackColor       =   &H005A5963&
      Caption         =   "Arguments"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   2
      Left            =   3360
      TabIndex        =   17
      Top             =   630
      Width           =   915
   End
   Begin VB.Label lblSkip 
      BackColor       =   &H005A5963&
      Caption         =   "Skip"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   6180
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   6
      Top             =   4380
      Width           =   435
   End
   Begin VB.Image Image1 
      Height          =   2970
      Left            =   0
      Picture         =   "frmWizard.frx":0000
      Top             =   0
      Width           =   3210
   End
   Begin VB.Label Label2 
      BackColor       =   &H005A5963&
      Caption         =   "Delay (secs)"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   3360
      TabIndex        =   4
      Top             =   990
      Width           =   975
   End
   Begin VB.Label lblBinary 
      BackColor       =   &H005A5963&
      Caption         =   "Executable:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   3360
      TabIndex        =   0
      Top             =   240
      Width           =   915
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Begin VB.Menu mnuScanForDll 
         Caption         =   "Scan Processes for DLL"
      End
      Begin VB.Menu mnuScanForUnknownMods 
         Caption         =   "Scan Procs for Unknown Dlls"
      End
      Begin VB.Menu mnuSpacer1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRWEScanAll 
         Caption         =   "RWE Memory Scan All"
      End
      Begin VB.Menu mnuRWEScanSingle 
         Caption         =   "RWE Memory Scan One"
      End
      Begin VB.Menu mnuSpacer2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReportViewer 
         Caption         =   "Open Saved Analysis"
      End
      Begin VB.Menu mnuKillAllLike 
         Caption         =   "Kill All Like"
      End
      Begin VB.Menu mnuSpacer3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExternal 
         Caption         =   "External"
         Begin VB.Menu mnuExt 
            Caption         =   "Sniffhit"
            Index           =   0
         End
         Begin VB.Menu mnuExt 
            Caption         =   "ProcWatch"
            Index           =   1
         End
         Begin VB.Menu mnuExt 
            Caption         =   "Api Logger"
            Index           =   2
         End
         Begin VB.Menu mnuExt 
            Caption         =   "DirWatch"
            Index           =   3
         End
         Begin VB.Menu mnuExt 
            Caption         =   "Command Prompt"
            Index           =   4
         End
      End
   End
End
Attribute VB_Name = "frmWizard"
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

Private Type config
    version As Integer
    sniffer As Byte
    apilog As Byte
    dirwatch As Byte
    delay As Long
    tcpdump As Byte
    interface As Byte
End Type
 
Private cfg As config
Private cfgFile As String
Private procWatch As String

Private going_toMainUI As Boolean
Private Declare Function GetCurrentProcessId Lib "kernel32.dll" () As Long

Private Sub cmdAbout_Click()
    frmAbout.Show 1, Me
End Sub

Private Sub cmdTools_Click()
    PopupMenu mnuPopup
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveConfig
    If Len(txtRWEScan) > 0 Then SaveMySetting "txtRWEScan", txtRWEScan.Text
    Dim f As Form
    If Not going_toMainUI Then
        For Each f In Forms
            Unload f
        Next
    End If
End Sub

Private Sub lblBinary_Click()
    MsgBox "SysAnalyzer can launch any file type which has a registered shell extension such as doc,pdf,html as well as standard executable extensions such as exe, pif,com, scr etc. Built in support is also included for launching dlls through the use of a helper application. You can also use this textbox to launch the parent application, and use the arguments box to load the specific malicious file", vbInformation
End Sub

Private Sub lblBuildKnownFileDB_Click()
    
    On Error Resume Next
    
    If Not known.Ready Then
        MsgBox "Known file database not found?", vbInformation
        Exit Sub
    End If
    
    frmKnownFiles.Show 1, Me
    
End Sub

Private Sub lblInterfaces_Click(Index As Integer)
    On Error Resume Next
    Dim f As String
    If isIde() Then
        f = App.path & "\..\..\win_dump.exe"
    Else
        f = App.path & "\win_dump.exe"
    End If
    
    Shell "cmd /k echo. && """ & f & """ -D && echo. && echo *** Use the interface index from the above list *** && echo.  ", vbNormalFocus
    
End Sub

Private Sub lblLaunchTcpDump_Click()
    launchtcpdump
End Sub

Private Sub lblSkip_Click()
    
   
    frmMain.Initalize
    frmMain.SSTab1.TabVisible(6) = True 'False
    frmMain.cmdDirWatch_Click
    frmMain.SSTab1.TabVisible(5) = False
    frmMain.lblTimer.Visible = False
    frmMain.Visible = True
    Me.Visible = False
    frmMain.mnuToolItem_Click 4 'take base snapshot..
    frmMain.lblDisplay = "Displaying Base Snapshot"
    
    going_toMainUI = True
    Unload Me
    
End Sub

Private Sub mnuExt_Click(Index As Integer)
    Dim ext(), f As String, ff As String
    
    ext = Array("sniff_hit", "proc_watch", "api_logger", "dirwatch_ui", "cmd")
    
    If isIde() Then
        f = App.path & "\..\..\" & ext(Index) & ".exe"
    Else
        f = App.path & "\" & ext(Index) & ".exe"
    End If
    
    If Not fso.FileExists(f) Then
        ff = Environ("windir") & "\system32\" & ext(Index) & ".exe"
        If fso.FileExists(ff) Then
            f = ff
        Else
            MsgBox "File not found: " & f, vbInformation
            Exit Sub
        End If
    End If
    
    On Error Resume Next
    Shell f, vbNormalFocus
    
End Sub

Private Sub mnuKillAllLike_Click()
    
    Dim c As Collection
    Dim p As CProcess
    Dim match As String
    Dim count As Long
    Dim myPid As Long
    
    match = InputBox("Enter parocess name match string to kill off")
    If Len(match) = 0 Then Exit Sub
    
    myPid = GetCurrentProcessId()
    Set c = diff.CProc.GetRunningProcesses()
    
    For Each p In c
        If InStr(1, p.path, match, vbTextCompare) > 0 And p.pid <> myPid Then
            diff.CProc.TerminateProces p.pid
            count = count + 1
        End If
    Next
    
    MsgBox count & " processes terminated", vbInformation
    
End Sub

Private Sub mnuReportViewer_Click()
    Dim f As String
    f = dlg.FolderDialog(, Me.hwnd)
    If Len(f) > 0 Then
        frmReportViewer.OpenAnalysisFolder f
    End If
End Sub

Private Sub mnuRWEScanAll_Click()
    frmInjectionScan.StealthInjectionScan
End Sub

Private Sub mnuRWEScanSingle_Click()
    
    Dim p As CProcess
    Set p = diff.CProc.SelectProcess()
    
    If p Is Nothing Then Exit Sub
    
    If diff.CProc.x64.IsProcess_x64(p.pid) <> r_32bit Then
        MsgBox x64Error, vbInformation
        Exit Sub
    End If
    
    frmInjectionScan.FindStealthInjections p.pid, p.path
    
End Sub

Private Sub mnuScanForDll_Click()
    ScanProcsForDll lblDisplay
End Sub

Private Sub mnuScanForUnknownMods_Click()
    ScanForUnknownMods lblDisplay
End Sub

Private Sub txtBinary_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    txtBinary = data.files(1)
End Sub

Sub SetConfigDefaults()
    With cfg
            .version = 2
            .apilog = 0
            .delay = 30
            .dirwatch = 1
            .sniffer = 1
            .interface = 1
            .tcpdump = 1
    End With
End Sub

Sub LoadConfig()
        
    If Not fso.FileExists(cfgFile) Then
        SetConfigDefaults
        SaveConfig
    Else
        Dim f As Long
        f = FreeFile
        Open cfgFile For Binary As f
        Get f, , cfg
        Close f
        If cfg.version <> 2 Then
            SetConfigDefaults
            SaveConfig
        End If
    End If
    
    With cfg
        chkApiLog.Value = .apilog
        chkNetworkAnalyzer = .sniffer
        chkWatchDirs = .dirwatch
        txtDelay = .delay
        txtInterface = .interface
        chkPacketCapture.Value = .tcpdump
    End With
    
End Sub

Sub SaveConfig()
        
    On Error Resume Next
    
    If Len(txtDelay) = 0 Or Not IsNumeric(txtDelay) Then txtDelay = 30
            
    With cfg
        .apilog = chkApiLog.Value
        .sniffer = chkNetworkAnalyzer
        .dirwatch = chkWatchDirs
        .delay = CLng(txtDelay)
        .interface = CByte(txtInterface)
        .tcpdump = chkPacketCapture.Value
    End With
    
    Dim f As Long
    f = FreeFile
    Open cfgFile For Binary As f
    Put f, , cfg
    Close f
        
End Sub


Private Sub cmdReadme_Click()
    
    Dim r As String
    r = App.path & IIf(isIde(), "\..\..", "") & "\SysAnalyzer_help.chm"
    
    If FileExists(r) Then
        ShellExecute 0, "open", r, "", "", 1
    Else
        MsgBox "Readme not found!" & vbCrLf & vbCrLf & r
    End If
    
    
End Sub

 
Private Sub cmdBrowse_Click(Index As Integer)
    Dim x
    x = dlg.OpenDialog(AllFiles, , "Open file for analysis", Me.hwnd)
    If Len(x) = 0 Then Exit Sub
    If Index = 0 Then
        txtBinary = x
    Else
        txtArgs = x
    End If
End Sub


Private Sub Form_Load()
      
    On Error GoTo hell
    
    Dim c As Collection
    Dim ip
    
    If IsVistaPlus() Then
        If Not IsProcessElevated() Then
            If Not MsgBox("Can I elevate to administrator?", vbYesNo) = vbYes Then
                If Not IsUserAnAdministrator() Then
                    lblAdmin.Caption = "This tool really requires admin privledges"
                Else
                    RunElevated App.path & "\sysanalyzer.exe", essSW_SHOW
                    End
                End If
            End If
        End If
    End If
    
    mnuPopup.Visible = False
    
    START_TIME = Now
    DebugLogFile = UserDeskTopFolder & "\debug.log"
    If fso.FileExists(DebugLogFile) Then fso.DeleteFile DebugLogFile
    fso.writeFile DebugLogFile, "-------[ SysAnalyzer v" & App.Major & "." & App.Minor & "." & App.Revision & "  " & START_TIME & " ]-------" & vbCrLf
    
    mnuScanForUnknownMods.Enabled = False
    
    If Not known.Ready Then
        lblKnown.Caption = "Not found"
    ElseIf known.Loaded Then
        lblKnown.Caption = "Loaded"
        mnuScanForUnknownMods.Enabled = True
    Else
        lblKnown.Caption = "Empty"
    End If
    
    cfgFile = App.path & "\cfg.dat"
    networkAnalyzer = App.path & IIf(isIde(), "\..\..", Empty) & "\sniff_hit.exe"
    procWatch = App.path & IIf(isIde(), "\..\..", Empty) & "\proc_watch.exe"
    tcpdump = App.path & IIf(isIde(), "\..\..", Empty) & "\win_dump.exe"
    txtRWEScan = GetMySetting("txtRWEScan", "explorer.exe,iexplore.exe,")
    
    Set c = AvailableInterfaces()
    For Each ip In c
        If ip <> "127.0.0.1" Then
            cboIp.AddItem ip
        End If
    Next
            
    If cboIp.ListCount <> 0 Then  'no active interfaces ?
        cboIp.ListIndex = 0
    End If
    
    cboIp.Visible = IIf(cboIp.ListCount > 1, True, False) 'try to keep config as easy as we can for them...
    lblip.Visible = cboIp.Visible


    'watchDirs.Add CStr(Environ("TEMP"))
    'watchDirs.Add CStr(Environ("WINDIR"))
    'watchDirs.Add CStr("C:\Program Files")
    watchDirs.Add CStr("C:\")
    
    Set cApiData = New Collection
    Set cLogData = New Collection
    
    LoadConfig

    If cboIp.ListCount = 0 Then  'no active interfaces ?
        chkPacketCapture.Enabled = False
        chkPacketCapture.Value = 0
        chkNetworkAnalyzer.Value = 0
        chkNetworkAnalyzer.Enabled = False
    End If
    
    If Len(Command) > 0 Then
        Dim cmd As String
        cmd = Trim(Replace(Command, """", Empty))
        If fso.FileExists(cmd) Then
            txtBinary = cmd
            'TODO auto run exe with settings if /launch
        End If
    End If
    
    If Len(txtBinary) = 0 And isIde() Then
        txtBinary = App.path & "\..\..\safe_test1.exe"
    End If

    
    Me.Icon = frmMain.Icon
    
Exit Sub
hell:
        MsgBox Err.Description
End Sub

Sub cmdStart_Click()
        
    On Error Resume Next
    
    ProcessesToRWEScan = txtRWEScan
    
    If chkPacketCapture.Value = 1 Then
        If Not IsNumeric(txtInterface.Text) Or txtInterface.Text = 0 Then
            MsgBox "Interface for tcpdump must be numeric and non-zero", vbInformation
            Exit Sub
        End If
    End If
    
    If Len(txtBinary) = 0 Then
        MsgBox "You must first set a binary to launch or choose skip to goto the main interface and analyze the system manually.", vbInformation
        Exit Sub
    End If
    
    If Not FileExists(txtBinary) Then
        MsgBox "Binary not found: " & txtBinary
        Exit Sub
    End If
    
    Dim cx As New Cx64
    If cx.isExe_x64(txtBinary) = r_64bit And chkApiLog.Value = 1 Then
        MsgBox "ApiLogger option is not yet compatiable with x64 targets", vbInformation
        chkApiLog.Value = 0
        Exit Sub
    End If
    
    If Len(txtDelay) = 0 Or Not IsNumeric(txtDelay) Then
        MsgBox "Invalid Delay Set defaulting to 30 seconds", vbInformation
        txtDelay = 30
    End If
        
    If chkNetworkAnalyzer.Value = 1 Then
        If Not isNetworkAnalyzerRunning() Then
            If fso.FileExists(networkAnalyzer) Then
                Shell """" & networkAnalyzer & """ /start /log """ & UserDeskTopFolder & """", vbMinimizedNoFocus
            Else
                MsgBox "Missing: " & networkAnalyzer
            End If
        End If
    End If
        
    If chkPacketCapture.Value = 1 Then launchtcpdump
    
    'must be last external process to launch as it monitors others...
    If fso.FileExists(procWatch) Then
        procWatchPID = Shell(procWatch & " /log=" & UserDeskTopFolder & "\ProcWatch.txt", vbMinimizedNoFocus)
    End If

    Dim baseName As String 'save a copy of the main malware executable for analysis folder..
    Dim saveAs As String
    
    baseName = "sample_" & fso.FileNameFromPath(txtBinary)
    If Len(baseName) = 0 Then baseName = "sample"
    saveAs = UserDeskTopFolder & "\" & baseName & "_"
    If Not fso.FileExists(saveAs) Then FileCopy txtBinary, saveAs
    
    going_toMainUI = True
    frmMain.Initalize
    
    frmMain.lblTimer = txtDelay & " Seconds remaining"
    frmMain.Visible = True
    Me.Visible = False
    
    diff.DoSnap1
    frmMain.Display = "Loading base snapshot."
    diff.ShowBaseSnap True   'only loads 2 tabs no known db lookup to eliminate delays..
    frmMain.Display = "Preparing to launch malware."
    tmrDelayShell.Enabled = True
    
Exit Sub
hell:
    MsgBox Err.Description
    
End Sub

Private Function launchtcpdump()
 ' http://www.winpcap.org/windump/docs/manual.htm
    '  -p not promiscious but not shortcut for ether host {local-hw-addr}
    '  -q quiet (less output on cmdline)
    '  -U write packets to file as received (not buffered)
    '  -i [interface index]
    '  -w [file path]
    '  -l show activity to stdout during capture..
    '  -s 0 capture entire packet do not truncate..
    
    On Error Resume Next
    Dim args As String
    Dim f As String
    Dim i As Long
    Dim c As Collection
    Dim ip As String
    
    i = 1
    
    If Not IsNumeric(txtInterface.Text) Or txtInterface.Text = 0 Then
        MsgBox "Interface for tcpdump must be numeric and non-zero", vbInformation
        Exit Function
    End If
    
    If fso.FileExists(tcpdump) Then
                
        f = UserDeskTopFolder() & "\capture.pcap"
        If fso.FileExists(f) Then
            While fso.FileExists(f)
                i = i + 1
                f = UserDeskTopFolder() & "\capture_" & i & ".pcap"
                If i = 100 Then Exit Function 'wtf?
            Wend
        End If
        
        args = " -w ""[PATH]"" -q -U -l -s 0 -i " & txtInterface & " ip src [IP] or ip dst [IP]"
        args = Replace(args, "[PATH]", f)
        args = Replace(args, "[IP]", cboIp.Text)
        args = "cmd /k """ & """" & tcpdump & """" & args & """"  'takes to long to initilize showing up in snapshots?
        'args = tcpdump & """" & args & """"
        
        Clipboard.Clear
        Clipboard.SetText args
        Shell args, vbMinimizedNoFocus
        Sleep 500
        
    Else
        MsgBox "Missing: " & tcpdump
    End If
    
End Function

Private Sub tmrDelayShell_Timer()

    tmrDelayShell.Enabled = False
    On Error GoTo hell
    
    If chkWatchDirs.Value = 1 Then
        DirWatchCtl True
    Else
        frmMain.SSTab1.TabVisible(6) = False
    End If
    
    frmMain.Display = "Launching malware..."
    
    If chkApiLog.Value = 1 Then
        Dim exe As String
            
        If VBA.Left(txtBinary, 4) = "pid:" Then
            exe = Replace(txtBinary, "pid:", Empty)
        ElseIf LCase(VBA.Right(txtBinary, 4)) = ".dll" Then
            exe = App.path & "\loadlib.exe """ & txtBinary & """"
        Else
            exe = txtBinary
        End If
        
        Dim dll As String
        
        If isIde Then
            dll = App.path & "\..\..\api_log.dll"
        Else
            dll = App.path & "\api_log.dll"
        End If
        
        If Not FileExists(dll) Then
            MsgBox "Could not locate Apilogger Dll?" & vbCrLf & vbCrLf & dll
            Exit Sub
        End If
        
        Dim tmp() As String
        
        debugLog "Starting process with api_log dll"
        StartProcessWithDLL exe & " " & txtArgs, dll, tmp()
    Else
        frmMain.SSTab1.TabVisible(5) = False
        If LCase(VBA.Right(txtBinary, 4)) = ".dll" Then
            debugLog "Starting dll with loadlib.exe"
            Shell App.path & "\loadlib.exe """ & txtBinary & """"
        Else
            debugLog "Starting malware directly"
            
            Dim ret As Long
            Dim args As String
             
            args = txtArgs
            If InStr(args, " ") > 0 Then args = """" & args & """"
             
            'this will handle word docs, pdfs, cpls, htms etc, not just exes
            'as long as a handler is registered for extension and extension is
            'correct for file type.
            ret = ShellExecute(0, "open", """" & txtBinary & """", args, fso.GetParentFolder(txtBinary), 1)
            If ret <= 32 Then
                debugLog "ShellExecute failed, trying VB Shell command.."
2               Shell txtBinary & " " & txtArgs, vbNormalFocus
            End If
             
             
        End If
    End If
    
    'test code
    'If isIde() And InStr(txtBinary, "safe_test") > 0 Then Shell "notepad.exe" 'for multiprocess testing..
    
    frmMain.Display = "Malware launched."
    frmMain.samplePath = txtBinary
    frmMain.StartCountDown CInt(txtDelay)
    Unload Me
    
Exit Sub
hell:
    If Erl = 2 Then
        'I could also fall back on using ShellExecute(open,cmdline) here..I should latter though..
        MsgBox "There was an error launching the malware directly. This could be due to an unknown file extension. ShellExecute did not know how to launch it." & _
               vbCrLf & vbCrLf & "For files such as these which can not be launched directly, you can use the parent application as the , and the malware as the argument." & _
               vbCrLf & vbCrLf & "The count down has not been initiated. You can now manually launch the file, and then after a period of time choose Tools->Take Snapshot 2" & _
               " and then choose Tools->Show Diff Report.", vbInformation
    Else
        MsgBox Err.Description
    End If
    
End Sub


Sub cmdStop_Click()

    On Error Resume Next
    diff.DoSnap2
    diff.ShowDiffReport
    
End Sub





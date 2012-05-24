VERSION 5.00
Begin VB.Form frmWizard 
   BackColor       =   &H005A5963&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SysAnalyzer Configuration Wizard"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtArgs 
      Height          =   285
      Left            =   4350
      TabIndex        =   20
      Top             =   570
      Width           =   3975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H005A5963&
      Caption         =   " Options "
      ForeColor       =   &H00E0E0E0&
      Height          =   2325
      Left            =   3840
      TabIndex        =   8
      Top             =   1290
      Width           =   5025
      Begin VB.ComboBox cboIp 
         Height          =   315
         Left            =   1140
         TabIndex        =   17
         Top             =   1890
         Width           =   2475
      End
      Begin VB.TextBox txtInterface 
         Height          =   285
         Left            =   2610
         TabIndex        =   14
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
         TabIndex        =   12
         Top             =   1320
         Width           =   1755
      End
      Begin VB.CheckBox chkNetworkAnalyzer 
         BackColor       =   &H005A5963&
         Caption         =   "Use SniffHit"
         ForeColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   480
         TabIndex        =   11
         Top             =   240
         Width           =   2775
      End
      Begin VB.CheckBox chkApiLog 
         BackColor       =   &H005A5963&
         Caption         =   "Use Api Logger"
         ForeColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   480
         TabIndex        =   10
         Top             =   570
         Width           =   1455
      End
      Begin VB.CheckBox chkWatchDirs 
         BackColor       =   &H005A5963&
         Caption         =   "Use Directory Watcher"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   480
         TabIndex        =   9
         Top             =   960
         Width           =   2835
      End
      Begin VB.Label lblip 
         BackColor       =   &H005A5963&
         Caption         =   "IP"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   840
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   13
         Top             =   1620
         Width           =   1245
      End
   End
   Begin VB.TextBox txtDelay 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4320
      TabIndex        =   6
      Text            =   "3"
      Top             =   930
      Width           =   555
   End
   Begin VB.Timer tmrDelayShell 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   90
      Top             =   3000
   End
   Begin VB.CommandButton cmdReadme 
      Caption         =   "Help"
      Height          =   375
      Left            =   3870
      TabIndex        =   4
      Top             =   3810
      Width           =   1155
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   375
      Left            =   7710
      TabIndex        =   3
      Top             =   3810
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
      Left            =   8430
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
   Begin VB.Label Label1 
      BackColor       =   &H005A5963&
      Caption         =   "Arguments"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   2
      Left            =   3390
      TabIndex        =   19
      Top             =   630
      Width           =   915
   End
   Begin VB.Label Label1 
      BackColor       =   &H005A5963&
      ForeColor       =   &H00E0E0E0&
      Height          =   1485
      Index           =   1
      Left            =   0
      TabIndex        =   18
      Top             =   2970
      Width           =   3255
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
      Left            =   6150
      TabIndex        =   7
      Top             =   3870
      Width           =   435
   End
   Begin VB.Image Image1 
      Height          =   4320
      Left            =   0
      Picture         =   "frmWizard.frx":0000
      Top             =   0
      Width           =   3240
   End
   Begin VB.Label Label2 
      BackColor       =   &H005A5963&
      Caption         =   "Delay (secs)"
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   3360
      TabIndex        =   5
      Top             =   990
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H005A5963&
      Caption         =   "Executable: "
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   0
      Left            =   3360
      TabIndex        =   0
      Top             =   240
      Width           =   915
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

Private going_toMainUI As Boolean

Private Sub Form_Unload(Cancel As Integer)
    SaveConfig
    Dim f As Form
    If Not going_toMainUI Then
        For Each f In Forms
            Unload f
        Next
    End If
End Sub

Private Sub lblInterfaces_Click(Index As Integer)
    On Error Resume Next
    Dim f As String
    If IsIde() Then
        f = App.path & "\..\..\windump.exe"
    Else
        f = App.path & "\windump.exe"
    End If
    
    Shell "cmd /k echo. && """ & f & """ -D && echo. && echo *** Use the interface index from the above list *** && echo.  ", vbNormalFocus
    
End Sub

Private Sub lblLaunchTcpDump_Click()
    launchtcpdump
End Sub

Private Sub lblSkip_Click()
    
    With frmMain
        .Initalize
        .SSTab1.TabVisible(6) = True 'False
        .cmdDirWatch_Click
        .SSTab1.TabVisible(5) = False
        .lblTimer.Visible = False
        .lblDisplay = "Use the tools menu to manually proceed"
        .Visible = True
    End With
    
    going_toMainUI = True
    
    Unload Me
    
End Sub

Private Sub txtBinary_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
    On Error Resume Next
    txtBinary = Data.files(1)
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
    r = App.path & "\SysAnalyzer_help.chm"
    
    If FileExists(r) Then
        'frmReport.ShowList ReadFile(r)
        ShellExecute 0, "open", r, "", "", 1
    Else
        MsgBox "Readme not found!" & vbCrLf & vbCrLf & r
    End If
    
    
End Sub

 
Private Sub cmdBrowse_Click()
    Dim x
    x = dlg.OpenDialog(exeFiles, , "Open file for analysis")
    If Len(x) = 0 Then Exit Sub
    txtBinary = x
End Sub


Private Sub Form_Load()
            
    On Error GoTo hell
    
    Dim c As Collection
    Dim ip
    
    'txtBinary = "D:\work_data\SysAnalyzer_2\examples\safe_test1.exe"
    
    cfgFile = App.path & "\cfg.dat"
    networkAnalyzer = App.path & "\sniff_hit.exe"
    tcpdump = App.path & "\windump.exe"
    
    If Not fso.FileExists(tcpdump) Then
        tcpdump = App.path & "\..\..\windump.exe"
    End If
    
    Set c = AvailableInterfaces()
    For Each ip In c
        If ip <> "127.0.0.1" Then
            cboIp.AddItem ip
        End If
    Next
            
    If cboIp.ListCount = 0 Then  'no active interfaces ?
        chkPacketCapture.Enabled = False
        chkPacketCapture.Value = 0
        chkNetworkAnalyzer.Value = 0
        chkNetworkAnalyzer.Enabled = False
    Else
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

    If Len(Command) > 0 Then
        Dim cmd As String
        cmd = Trim(Replace(Command, """", Empty))
        If fso.FileExists(cmd) Then
            txtBinary = cmd
            'TODO auto run exe with settings
        End If
    End If

    
    Me.Icon = frmMain.Icon
    
Exit Sub
hell:
        MsgBox Err.Description
End Sub

Sub cmdStart_Click()
        
    On Error Resume Next
    
    If chkPacketCapture.Value = 1 Then
        If Not IsNumeric(txtInterface.Text) Or txtInterface.Text = 0 Then
            MsgBox "Interface for tcpdump must be numeric and non-zero", vbInformation
            Exit Sub
        End If
    End If
    
    If Not FileExists(txtBinary) Then
        MsgBox "Binary not found: " & txtBinary
        Exit Sub
    End If
    
    If Len(txtDelay) = 0 Or Not IsNumeric(txtDelay) Then
        MsgBox "Invalid Delay Set defaulting to 30 seconds", vbInformation
        txtDelay = 30
    End If
        
    If chkNetworkAnalyzer.Value = 1 Then
        If Not isNetworkAnalyzerRunning() Then
            If fso.FileExists(networkAnalyzer) Then
                Shell """" & networkAnalyzer & """ /start", vbMinimizedNoFocus
            Else
                MsgBox "Missing: " & networkAnalyzer
            End If
        End If
    End If
    
    If chkPacketCapture.Value = 1 Then launchtcpdump
    
    going_toMainUI = True
    frmMain.Initalize
    
    diff.DoSnap1
    diff.ShowBaseSnap
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
        
        If IsIde Then
            dll = App.path & "\..\..\api_log.dll"
        Else
            dll = App.path & "\api_log.dll"
        End If
        
        If Not FileExists(dll) Then
            MsgBox "Could not locate Apilogger Dll?" & vbCrLf & vbCrLf & dll
            Exit Sub
        End If
        
        Dim tmp() As String
        
        StartProcessWithDLL exe & " " & txtArgs, dll, tmp()
    Else
        frmMain.SSTab1.TabVisible(5) = False
        If LCase(VBA.Right(txtBinary, 4)) = ".dll" Then
            Shell App.path & "\loadlib.exe """ & txtBinary & """"
        Else
            Shell txtBinary & " " & txtArgs
        End If
    End If
    
    frmMain.samplePath = txtBinary
    frmMain.StartCountDown CInt(txtDelay)
    Unload Me
    
Exit Sub
hell:
    MsgBox Err.Description
End Sub


Sub cmdStop_Click()

    On Error Resume Next
    diff.DoSnap2
    diff.ShowDiffReport
    
End Sub





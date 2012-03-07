VERSION 5.00
Begin VB.Form frmWizard 
   BackColor       =   &H005A5963&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "SysAnalyzer Configuration Wizard"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7770
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   7770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H005A5963&
      Caption         =   " Options "
      ForeColor       =   &H00E0E0E0&
      Height          =   1455
      Left            =   3840
      TabIndex        =   8
      Top             =   900
      Width           =   3795
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
         Top             =   600
         Width           =   2775
      End
      Begin VB.CheckBox chkWatchDirs 
         BackColor       =   &H005A5963&
         Caption         =   "Use Directory Watcher"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   480
         TabIndex        =   9
         Top             =   1020
         Width           =   2835
      End
   End
   Begin VB.TextBox txtDelay 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4320
      TabIndex        =   6
      Text            =   "3"
      Top             =   540
      Width           =   555
   End
   Begin VB.Timer tmrDelayShell 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   7200
      Top             =   3120
   End
   Begin VB.CommandButton cmdReadme 
      Caption         =   "Help"
      Height          =   375
      Left            =   3300
      TabIndex        =   4
      Top             =   2520
      Width           =   1155
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   375
      Left            =   6540
      TabIndex        =   3
      Top             =   2520
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
      Left            =   7320
      TabIndex        =   2
      Top             =   180
      Width           =   375
   End
   Begin VB.TextBox txtBinary 
      Height          =   285
      Left            =   4320
      OLEDropMode     =   1  'Manual
      TabIndex        =   1
      Top             =   180
      Width           =   2895
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
      Left            =   5280
      TabIndex        =   7
      Top             =   2640
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
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H005A5963&
      Caption         =   "Executable: "
      ForeColor       =   &H00E0E0E0&
      Height          =   255
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
    sniffer As Byte
    apilog As Byte
    dirwatch As Byte
    delay As Long
End Type
 
Private cfg As config
Private cfgFile As String

Private Sub Form_Unload(Cancel As Integer)
    SaveConfig
End Sub

Private Sub lblSkip_Click()
    
    With frmMain
        .Initalize
        .SSTab1.TabVisible(6) = True 'False
        .cmdDirWatch_Click
        .SSTab1.TabVisible(5) = False
        .lblTimer.Visible = False
        .lblDisplay = "Use the tools menu to manually proceede"
        .Visible = True
    End With
    
    Unload Me
    
End Sub

Private Sub txtBinary_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
    On Error Resume Next
    txtBinary = Data.files(1)
End Sub

Sub LoadConfig()
    
    If Not fso.FileExists(cfgFile) Then
        With cfg
            .apilog = 0
            .delay = 30
            .dirwatch = 1
            .sniffer = 1
        End With
        SaveConfig
    Else
        Dim f As Long
        f = FreeFile
        Open cfgFile For Binary As f
        Get f, , cfg
        Close f
    End If
    
    With cfg
        chkApiLog.value = .apilog
        chkNetworkAnalyzer = .sniffer
        chkWatchDirs = .dirwatch
        txtDelay = .delay
    End With
    
End Sub

Sub SaveConfig()
        
    If Len(txtDelay) = 0 Or Not IsNumeric(txtDelay) Then txtDelay = 30
            
    With cfg
        .apilog = chkApiLog.value
        .sniffer = chkNetworkAnalyzer
        .dirwatch = chkWatchDirs
        .delay = CLng(txtDelay)
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
 
    'txtBinary = "D:\work_data\SysAnalyzer_2\examples\safe_test1.exe"
    
    cfgFile = App.path & "\cfg.dat"
    networkAnalyzer = App.path & "\sniff_hit.exe"
    
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
            
    If Not FileExists(txtBinary) Then
        MsgBox "Binary not found: " & txtBinary
        Exit Sub
    End If
    
    If Len(txtDelay) = 0 Or Not IsNumeric(txtDelay) Then
        MsgBox "Invalid Delay Set defaulting to 30 seconds", vbInformation
        txtDelay = 30
    End If
        
    If chkNetworkAnalyzer.value = 1 Then
        If Not isNetworkAnalyzerRunning() Then
            If fso.FileExists(networkAnalyzer) Then
                Shell """" & networkAnalyzer & """ /start", vbNormalNoFocus
            Else
                MsgBox "Missing: " & networkAnalyzer
            End If
        End If
    End If
            
    frmMain.Initalize
    
    diff.DoSnap1
    diff.ShowBaseSnap
    tmrDelayShell.Enabled = True
    
Exit Sub
hell:
    MsgBox Err.Description
    
End Sub


Private Sub tmrDelayShell_Timer()
    tmrDelayShell.Enabled = False
    On Error GoTo hell
    
    If chkWatchDirs.value = 1 Then
        DirWatchCtl True
    Else
        frmMain.SSTab1.TabVisible(6) = False
    End If
    
    If chkApiLog.value = 1 Then
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
        
        StartProcessWithDLL exe, dll, tmp()
    Else
        frmMain.SSTab1.TabVisible(5) = False
        If LCase(VBA.Right(txtBinary, 4)) = ".dll" Then
            Shell App.path & "\loadlib.exe """ & txtBinary & """"
        Else
            Shell txtBinary
        End If
    End If
    
    frmMain.samplePath = txtBinary
    frmMain.StartCountDown CInt(txtDelay)
    
Exit Sub
hell:
    MsgBox Err.Description
End Sub


Sub cmdStop_Click()

    On Error Resume Next
    diff.DoSnap2
    diff.ShowDiffReport
    
End Sub





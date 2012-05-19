VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "ApiLogger"
   ClientHeight    =   7665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10275
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   7665
   ScaleWidth      =   10275
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   315
      Left            =   4920
      TabIndex        =   33
      Top             =   30
      Width           =   615
   End
   Begin VB.CommandButton cmdSelectProcess 
      Caption         =   "PID"
      Height          =   315
      Left            =   5580
      TabIndex        =   32
      Top             =   30
      Width           =   555
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Resume"
      Height          =   375
      Left            =   2460
      TabIndex        =   31
      Top             =   3630
      Width           =   1425
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Suspend"
      Height          =   375
      Left            =   1110
      TabIndex        =   30
      Top             =   3630
      Width           =   1305
   End
   Begin VB.CommandButton cmdTerminate 
      Caption         =   "Terminate"
      Height          =   375
      Left            =   3930
      TabIndex        =   29
      Top             =   3630
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Re-Apply"
      Height          =   315
      Left            =   6180
      TabIndex        =   21
      Top             =   1440
      Width           =   1305
   End
   Begin VB.Frame Frame1 
      Caption         =   " Api Startup Logging Options "
      Height          =   3435
      Left            =   7590
      TabIndex        =   19
      Top             =   60
      Width           =   2565
      Begin VB.CheckBox Check3 
         Caption         =   "Capture UrlDownload*"
         Enabled         =   0   'False
         Height          =   285
         Left            =   180
         TabIndex        =   28
         Top             =   2910
         Width           =   1965
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Capture send/recv bufs"
         Enabled         =   0   'False
         Height          =   255
         Left            =   180
         TabIndex        =   27
         Top             =   2550
         Width           =   1995
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Advance Time Checks"
         Enabled         =   0   'False
         Height          =   315
         Left            =   180
         TabIndex        =   26
         Top             =   2160
         Width           =   2145
      End
      Begin VB.CheckBox chkAdvanceGetTick 
         Caption         =   "Advance GetTickCount"
         Height          =   315
         Left            =   180
         TabIndex        =   25
         Top             =   1770
         Width           =   2205
      End
      Begin VB.CheckBox chkBlockOpenProcess 
         Caption         =   "Block OpenProcess"
         Height          =   345
         Left            =   180
         TabIndex        =   24
         Top             =   1380
         Width           =   1935
      End
      Begin VB.CheckBox chkNoRegistry 
         Caption         =   "No Registry Hooks"
         Height          =   285
         Left            =   180
         TabIndex        =   23
         Top             =   1020
         Width           =   1845
      End
      Begin VB.CheckBox chkNoGetProc 
         Caption         =   "No GetProcAddress"
         Height          =   285
         Left            =   180
         TabIndex        =   22
         Top             =   660
         Width           =   1725
      End
      Begin VB.CheckBox chkIgnoreSleep 
         Caption         =   "Ignore Long Sleeps"
         Height          =   315
         Left            =   180
         TabIndex        =   20
         Top             =   300
         Width           =   2055
      End
   End
   Begin VB.TextBox txtArgs 
      Height          =   315
      Left            =   960
      TabIndex        =   18
      Top             =   360
      Width           =   5145
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Copy"
      Height          =   375
      Left            =   8970
      TabIndex        =   16
      Top             =   3630
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      Height          =   375
      Left            =   7710
      TabIndex        =   15
      Top             =   3630
      Width           =   1095
   End
   Begin VB.TextBox txtIgnore 
      Height          =   315
      Left            =   960
      TabIndex        =   14
      Top             =   1410
      Width           =   5115
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Stop Logging"
      Height          =   375
      Left            =   5700
      TabIndex        =   12
      Top             =   3630
      Width           =   1905
   End
   Begin VB.CommandButton cmdContinue 
      Caption         =   "Continue"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6180
      TabIndex        =   11
      Top             =   1080
      Width           =   1305
   End
   Begin VB.TextBox txtDumpAt 
      Height          =   285
      Left            =   960
      TabIndex        =   9
      Top             =   1050
      Width           =   5145
   End
   Begin VB.ListBox List2 
      Height          =   1425
      Left            =   30
      TabIndex        =   8
      Top             =   2010
      Width           =   7455
   End
   Begin VB.TextBox txtDll 
      Height          =   285
      Left            =   960
      OLEDropMode     =   1  'Manual
      TabIndex        =   6
      Top             =   720
      Width           =   5145
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Inject && Log"
      Height          =   315
      Left            =   6180
      TabIndex        =   3
      Top             =   30
      Width           =   1335
   End
   Begin VB.TextBox txtPacked 
      Height          =   315
      Left            =   960
      OLEDropMode     =   1  'Manual
      TabIndex        =   2
      Top             =   0
      Width           =   3885
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3420
      Left            =   0
      TabIndex        =   0
      Top             =   4140
      Width           =   10155
   End
   Begin VB.Label Label7 
      Caption         =   "Args"
      Height          =   285
      Left            =   0
      TabIndex        =   17
      Top             =   360
      Width           =   945
   End
   Begin VB.Label Label6 
      Caption         =   "CSV Ignore"
      Height          =   255
      Left            =   30
      TabIndex        =   13
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Freeze At"
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   1110
      Width           =   915
   End
   Begin VB.Label Label4 
      Caption         =   "Injection Details"
      Height          =   315
      Left            =   0
      TabIndex        =   7
      Top             =   1770
      Width           =   1755
   End
   Begin VB.Label Label3 
      Caption         =   "Inject DLL"
      Height          =   315
      Left            =   0
      TabIndex        =   5
      Top             =   750
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "API Call Log"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   3870
      Width           =   1035
   End
   Begin VB.Label Label1 
      Caption         =   "Executable"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   60
      Width           =   975
   End
End
Attribute VB_Name = "Form2"
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

Private Type PROCESS_INFORMATION
   hProcess As Long
   hThread As Long
   dwProcessId As Long
   dwThreadId As Long
End Type

Private Type STARTUPINFO
        cb As Long
        lpReserved As String
        lpDesktop As String
        lpTitle As String
        dwX As Long
        dwY As Long
        dwXSize As Long
        dwYSize As Long
        dwXCountChars As Long
        dwYCountChars As Long
        dwFillAttribute As Long
        dwFlags As Long
        wShowWindow As Integer
        cbReserved2 As Integer
        lpReserved2 As Long
        hStdInput As Long
        hStdOutput As Long
        hStdError As Long
End Type

Private Enum ProcessAccessTypes
    PROCESS_TERMINATE = (&H1)
    PROCESS_CREATE_THREAD = (&H2)
    PROCESS_SET_SESSIONID = (&H4)
    PROCESS_VM_OPERATION = (&H8)
    PROCESS_VM_READ = (&H10)
    PROCESS_VM_WRITE = (&H20)
    PROCESS_DUP_HANDLE = (&H40)
    PROCESS_CREATE_PROCESS = (&H80)
    PROCESS_SET_QUOTA = (&H100)
    PROCESS_SET_INFORMATION = (&H200)
    PROCESS_QUERY_INFORMATION = (&H400)
    STANDARD_RIGHTS_REQUIRED = &HF0000
    SYNCHRONIZE = &H100000
    PROCESS_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF)
End Enum

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Long, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function CreateRemoteThread Lib "kernel32" (ByVal ProcessHandle As Long, lpThreadAttributes As Long, ByVal dwStackSize As Long, ByVal lpStartAddress As Any, ByVal lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadID As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function VirtualAllocEx Lib "kernel32" (ByVal hProcess As Long, lpAddress As Any, ByVal dwSize As Long, ByVal fAllocType As Long, FlProtect As Long) As Long
Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDriectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function ResumeThread Lib "kernel32" (ByVal hThread As Long) As Long
Private Declare Sub DebugBreak Lib "kernel32" ()
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function NtSuspendProcess Lib "ntdll.dll" (ByVal hProc As Long) As Long
Private Declare Function NtResumeProcess Lib "ntdll.dll" (ByVal hProc As Long) As Long



'I used my subclass library for simplicity, you can use whatever sub
'class technique or inline code you desire...
Dim WithEvents sc As CSubclass2
Attribute sc.VB_VarHelpID = -1
Dim dlg As New clsCmnDlg

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Const WM_COPYDATA = &H4A
Private Const WM_DISPLAY_TEXT = 3

Private Type COPYDATASTRUCT
    dwFlag As Long
    cbSize As Long
    lpData As Long
End Type

Dim pi As New CProcessInfo

Dim noLog As Boolean
Dim readyToReturn As Boolean
Dim ignored() As String
Dim getTickIncrements As Long
Dim g_hProc As Long            'used latter for ReadProcessmemory calls on send/recv bufs

'todo: parse incoming api to: handles -> process/file/socket mapping..,
'                             capture downloads
'                             capture send/recv bufs
'                             switch list to listview to capture more like bufs in .tag

Function AryIsEmpty(ary) As Boolean
  On Error GoTo oops
    Dim i As Long
    i = UBound(ary)  '<- throws error if not initalized
    AryIsEmpty = False
  Exit Function
oops: AryIsEmpty = True
End Function

Function ignoreit(v) As Boolean
    Dim i As Long
    
    If AryIsEmpty(ignored) Then Exit Function
    
    For i = 0 To UBound(ignored)
        If Len(Trim(ignored(i))) > 0 Then
            If InStr(1, v, ignored(i), vbTextCompare) Then
                ignoreit = True
                Exit Function
            End If
        End If
    Next
    
End Function

Private Sub cmdBrowse_Click()
    Dim f As String
    f = dlg.OpenDialog(AllFiles, , "Open Executable to monitor", Me.hwnd)
    If Len(f) = 0 Then Exit Sub
    txtPacked = f
End Sub

Private Sub cmdContinue_Click()
    readyToReturn = True
End Sub

Private Sub cmdSelectProcess_Click()
    Dim cp As CProcess
    Set cp = frmListProcess.SelectProcess(pi.GetRunningProcesses)
    If Not cp Is Nothing Then
        txtPacked = "pid:" & cp.pid
    End If
End Sub

Private Sub cmdStart_Click()
        
    Dim exe As String
    
    List1.Clear
    List2.Clear
    Erase ignored
    
    If Len(txtIgnore) > 0 Then
        ignored = Split(txtIgnore, ",")
    End If
    
    If VBA.Left(txtPacked, 4) = "pid:" Then
        exe = Replace(txtPacked, "pid:", Empty)
    Else
        If Not FileExists(txtPacked) Then
            MsgBox "Executable not found"
            Exit Sub
        End If
        exe = txtPacked
    End If
    
    If Not FileExists(txtDll) Then
        MsgBox "Dll To inject not found"
        Exit Sub
    End If
    
    If Len(txtArgs) > 0 Then exe = exe & " " & txtArgs
    StartProcessWithDLL exe, txtDll
    
End Sub

Private Sub cmdTerminate_Click()
    List2.AddItem "TerminateProcess = " & TerminateProcess(g_hProc, 1)
End Sub

Private Sub Command1_Click()
    
    If InStr(Command1.Caption, "Stop") > 0 Then
        noLog = True
        Command1.Caption = "Resume Logging"
    Else
        noLog = False
        Command1.Caption = "Stop Logging"
    End If
    
End Sub

Private Sub Command2_Click()
    List1.Clear
End Sub

Private Sub Command3_Click()
    On Error Resume Next
    Dim i As Long, t
    For i = 0 To List1.ListCount
        t = t & List1.List(i) & vbCrLf
    Next
    Clipboard.Clear
    Clipboard.SetText t
    MsgBox Len(t) & " bytes copied"
End Sub

Private Sub Command4_Click()
    Dim i As Long
    Erase ignored
    If Len(txtIgnore) > 0 Then
        ignored = Split(txtIgnore, ",")
    End If
    For i = List1.ListCount To 0 Step -1
        If ignoreit(List1.List(i)) Then
            List1.RemoveItem i
        End If
    Next
End Sub

Private Sub Command5_Click()
    List2.AddItem "NtSuspendProcess = " & NtSuspendProcess(g_hProc)
End Sub

Private Sub Command6_Click()
    List2.AddItem "NtResumeProcess = " & NtResumeProcess(g_hProc)
End Sub

Private Sub Form_Load()
    Set sc = New CSubclass2
    
    sc.AttachMessage Me.hwnd, WM_COPYDATA
     
    Dim defaultdll, defaultexe
    
    If isIde() Then defaultexe = App.path & "\..\..\..\safe_test1.exe"
    defaultdll = App.path & IIf(isIde(), "\..\..\..", "") & "\api_log.dll"
    If FileExists(defaultdll) Then txtDll = defaultdll
    If FileExists(defaultexe) Then txtPacked = defaultexe
    
    txtIgnore = GetMySetting("Ignore", "")
    
    If Len(Command) > 0 Then
        txtPacked = Replace(Command, """", Empty)
    End If
    
End Sub

Function isIde() As Boolean
    On Error Resume Next
    Debug.Print 1 / 0
    isIde = (Err.Number <> 0)
End Function

Private Sub Form_Resize()
    On Error Resume Next
    List1.Width = Me.Width - List1.Left - 200
    List1.Height = Me.Height - List1.Top - 400
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveMySetting "Ignore", txtIgnore
End Sub

Private Sub sc_MessageReceived(hwnd As Long, wMsg As Long, wParam As Long, lParam As Long, Cancel As Boolean)
    If wMsg = WM_COPYDATA Then RecieveTextMessage lParam
End Sub

Private Sub HandleConfig(msg As String)
    On Error GoTo hell
    Dim cmd
    
    cmd = Split(msg, ":")
    
    If InStr(1, msg, "gettickvalue", vbTextCompare) < 1 Then
        Debug.Print msg
    End If
    
    Select Case LCase(cmd(1))
        Case "nosleep": If chkIgnoreSleep.value = 1 Then sc.OverRideRetVal 1
        Case "noregistry": If chkNoRegistry.value = 1 Then sc.OverRideRetVal 1
        Case "nogetproc": If chkNoGetProc.value = 1 Then sc.OverRideRetVal 1
        Case "querygettick": If chkAdvanceGetTick.value = 1 Then sc.OverRideRetVal 1
        Case "blockopenprocess": If chkBlockOpenProcess.value = 1 Then sc.OverRideRetVal 1
        
        Case "gettickvalue":
        
                    If getTickIncrements = 0 Then
                        sc.OverRideRetVal GetTickCount()
                    Else
1                        sc.OverRideRetVal GetTickCount() + (getTickIncrements * &H10000)
                    End If
                    
                    getTickIncrements = getTickIncrements + 1
                    
    End Select
        
hell:

    If Erl = 1 Then
        getTickIncrements = 0
    End If
    
End Sub

Private Sub RecieveTextMessage(lParam As Long)
   
    Dim CopyData As COPYDATASTRUCT
    Dim Buffer(1 To 2048) As Byte
    Dim temp As String
    Dim hProcess As Long
    Dim writeLen As Long
    Dim ret As Long
    Dim hThread As Long
    
    CopyMemory CopyData, ByVal lParam, Len(CopyData)
    
    If CopyData.dwFlag = 3 Then
        CopyMemory Buffer(1), ByVal CopyData.lpData, CopyData.cbSize
        temp = StrConv(Buffer, vbUnicode)
        temp = Left$(temp, InStr(1, temp, Chr$(0)) - 1)
        
        If VBA.Left(temp, 10) = "***config:" Then
            HandleConfig temp
            Exit Sub
        End If
        
        'heres where we work with the intercepted message
        If Not noLog Then
            
            If ignoreit(temp) Then Exit Sub
            
            List1.AddItem temp
            List1.ListIndex = List1.ListCount - 1
        
            If Len(txtDumpAt) > 0 Then
                If InStr(1, temp, txtDumpAt, vbTextCompare) > 0 Then
                    'sendMessage is a blocking call so we will sit here till user hits continue
                    cmdContinue.Enabled = True
                    readyToReturn = False
                    While Not readyToReturn
                        DoEvents
                        Sleep 60
                    Wend
                    cmdContinue.Enabled = False
                End If
            End If
            
        End If
        
    End If
    
End Sub



Public Function StartProcessWithDLL(exePath As String, dllPath As String) As Long

    Dim hProcess As Long
    Dim lpfnLoadLib As Long
    Dim ret As Long
    Dim lpdllPath As Long
    Dim pi As PROCESS_INFORMATION
    Dim si As STARTUPINFO
    Dim hThread As Long
    Dim writeLen As Long
    Dim b() As Byte
    Dim buflen As Long
    
    Const PAGE_READWRITE = 4
    Const CREATE_SUSPENDED = &H4
    Const MEM_COMMIT = &H1000
    
    b() = StrConv(dllPath & Chr(0), vbFromUnicode)
    buflen = UBound(b) + 1
    
    With List2
        .Clear
        
        If IsNumeric(exePath) Then
            hProcess = OpenProcess(PROCESS_ALL_ACCESS, False, CLng(exePath))
            g_hProc = hProcess
            .AddItem "Opening PID: " & exePath & " Process Handle=" & hProcess
        Else
            ret = CreateProcess(0&, exePath, 0&, 0&, 1&, CREATE_SUSPENDED, 0&, 0&, si, pi)
            .AddItem "Create Process Suspended: " & ret & IIf(ret = 0, " Failed", " PID: " & pi.dwProcessId)
            
            hProcess = OpenProcess(PROCESS_ALL_ACCESS, False, pi.dwProcessId)
            g_hProc = hProcess
            .AddItem "OpenProcess Handle=" & hProcess
        End If
                    
        lpdllPath = VirtualAllocEx(hProcess, ByVal 0, buflen, MEM_COMMIT, ByVal PAGE_READWRITE)
        .AddItem "Remote Allocation base: " & Hex(lpdllPath)
            
        ret = WriteProcessMemory(hProcess, ByVal lpdllPath, b(0), buflen, writeLen)
        .AddItem "WriteProcessMemory=" & ret & " BufLen=" & buflen & " Bytes Written: " & writeLen
                
        lpfnLoadLib = GetProcAddress(GetModuleHandle("kernel32.dll"), "LoadLibraryA")
        .AddItem "LoadLibraryA = " & Hex(lpfnLoadLib)
        
        'DebugBreak
        ret = CreateRemoteThread(hProcess, ByVal 0, 0, lpfnLoadLib, lpdllPath, 0, hThread)
        .AddItem "CreateRemoteThread = " & ret & " ThreadID: " & Hex(hThread)
                
        Sleep 900
        
        If Not IsNumeric(exePath) Then ResumeThread pi.hThread
        
    End With

End Function







Function FileExists(path) As Boolean
  If Len(path) = 0 Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True
End Function
 
 

Private Sub txtDll_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single)
    txtDll = Data.Files(1)
End Sub

Private Sub txtPacked_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, Y As Single, State As Integer)
    txtPacked = Data.Files(1)
End Sub



Function IsHex(it) As Long
    On Error GoTo out
      IsHex = CLng("&H" & it)
    Exit Function
out:  IsHex = 0
End Function



Function GetMySetting(key, def)
    GetMySetting = GetSetting(App.EXEName, "General", key, def)
End Function

Sub SaveMySetting(key, value)
    SaveSetting App.EXEName, "General", key, value
End Sub


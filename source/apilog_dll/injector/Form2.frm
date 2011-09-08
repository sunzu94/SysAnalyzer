VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "ApiLogger"
   ClientHeight    =   6630
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7515
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   6630
   ScaleWidth      =   7515
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Copy"
      Height          =   255
      Left            =   6120
      TabIndex        =   16
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      Height          =   255
      Left            =   4560
      TabIndex        =   15
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox txtIgnore 
      Height          =   315
      Left            =   960
      TabIndex        =   14
      Top             =   960
      Width           =   6255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Stop Logging"
      Height          =   315
      Left            =   6240
      TabIndex        =   12
      Top             =   660
      Width           =   1335
   End
   Begin VB.CommandButton cmdContinue 
      Caption         =   "Continue"
      Enabled         =   0   'False
      Height          =   315
      Left            =   3000
      TabIndex        =   11
      Top             =   660
      Width           =   975
   End
   Begin VB.TextBox txtDumpAt 
      Height          =   285
      Left            =   960
      TabIndex        =   9
      Top             =   660
      Width           =   1875
   End
   Begin VB.ListBox List2 
      Height          =   1230
      Left            =   -60
      TabIndex        =   8
      Top             =   1560
      Width           =   7455
   End
   Begin VB.TextBox txtDll 
      Height          =   285
      Left            =   960
      OLEDropMode     =   1  'Manual
      TabIndex        =   6
      Top             =   360
      Width           =   5295
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Inject && Log"
      Height          =   315
      Left            =   6420
      TabIndex        =   3
      Top             =   300
      Width           =   1095
   End
   Begin VB.TextBox txtPacked 
      Height          =   315
      Left            =   960
      OLEDropMode     =   1  'Manual
      TabIndex        =   2
      Top             =   0
      Width           =   5295
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
      Top             =   3180
      Width           =   7455
   End
   Begin VB.Label Label6 
      Caption         =   "Ignore"
      Height          =   255
      Left            =   0
      TabIndex        =   13
      Top             =   1020
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Freeze At"
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   720
      Width           =   915
   End
   Begin VB.Label Label4 
      Caption         =   "Details"
      Height          =   315
      Left            =   0
      TabIndex        =   7
      Top             =   1320
      Width           =   915
   End
   Begin VB.Label Label3 
      Caption         =   "Inject DLL"
      Height          =   315
      Left            =   0
      TabIndex        =   5
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "API Call Log"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   2880
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


'I used my subclass library for simplicity, you can use whatever sub
'class technique or inline code you desire...
Dim WithEvents sc As clsSubClass
Attribute sc.VB_VarHelpID = -1

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Const WM_COPYDATA = &H4A
Private Const WM_DISPLAY_TEXT = 3

Private Type COPYDATASTRUCT
    dwFlag As Long
    cbSize As Long
    lpData As Long
End Type

Dim noLog As Boolean
Dim readyToReturn As Boolean
Dim ignored() As String



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
        If InStr(1, v, ignored(i), vbTextCompare) Then
            ignoreit = True
            Exit Function
        End If
    Next
    
End Function

Private Sub cmdContinue_Click()
    readyToReturn = True
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
    
    StartProcessWithDLL exe, txtDll
    
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

Private Sub Form_Load()
    Set sc = New clsSubClass
    
    sc.AttachMessage Me.hwnd, WM_COPYDATA
     
    Dim defaultdll, defaultexe
    
    'defaultexe = App.path & "\pespined.exe"
    defaultdll = App.path & "\api_log.dll"
    If FileExists(defaultdll) Then txtDll = defaultdll
    If FileExists(defaultexe) Then txtPacked = defaultexe
    
    txtIgnore = GetMySetting("Ignore", "")
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveMySetting "Ignore", txtIgnore
End Sub

Private Sub sc_MessageReceived(hwnd As Long, wMsg As Long, wParam As Long, lParam As Long, Cancel As Boolean)
    If wMsg = WM_COPYDATA Then RecieveTextMessage lParam
End Sub


Private Sub RecieveTextMessage(lParam As Long)
   
    Dim CopyData As COPYDATASTRUCT
    Dim Buffer(1 To 2048) As Byte
    Dim Temp As String
    Dim hProcess As Long
    Dim writeLen As Long
    Dim ret As Long
    Dim hThread As Long
    
    CopyMemory CopyData, ByVal lParam, Len(CopyData)
    
    If CopyData.dwFlag = 3 Then
        CopyMemory Buffer(1), ByVal CopyData.lpData, CopyData.cbSize
        Temp = StrConv(Buffer, vbUnicode)
        Temp = Left$(Temp, InStr(1, Temp, Chr$(0)) - 1)
        'heres where we work with the intercepted message
        If Not noLog Then
            
            If ignoreit(Temp) Then Exit Sub
            
            List1.AddItem Temp
            List1.ListIndex = List1.ListCount - 1
        
            If Len(txtDumpAt) > 0 Then
                If InStr(1, Temp, txtDumpAt, vbTextCompare) > 0 Then
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
            .AddItem "Opening PID: " & exePath & " Process Handle=" & hProcess
        Else
            ret = CreateProcess(0&, exePath, 0&, 0&, 1&, CREATE_SUSPENDED, 0&, 0&, si, pi)
            .AddItem "Create Process Suspended: " & ret & IIf(ret = 0, " Failed", " PID: " & pi.dwProcessId)
            
            hProcess = OpenProcess(PROCESS_ALL_ACCESS, False, pi.dwProcessId)
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


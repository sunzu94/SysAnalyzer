VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cx64"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long '()

'min build Windows XP
Private Declare Sub GetNativeSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)

'min build xp sp2
Private Declare Function IsWow64Process Lib "kernel32" (ByVal hProcess As Long, ByRef Wow64Process As Long) As Long

Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function GetTokenInformation Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal TokenInformationClass As Integer, TokenInformation As Any, ByVal TokenInformationLength As Long, ReturnLength As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32.dll" () As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
Private Declare Function GetLastError Lib "kernel32.dll" () As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, ByRef NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, ByRef PreviousState As TOKEN_PRIVILEGES, ByRef ReturnLength As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long

Private Type LUID
    LowPart As Long
    HighPart As Long
End Type

Private Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    TheLuid As LUID
    Attributes As Long
End Type

Private Type SYSTEM_INFO
    wProcessorArchitecture As Integer
    wReserved As Integer
    dwPageSize As Long
    lpMinimumApplicationAddress As Long
    lpMaximumApplicationAddress As Long
    dwActiveProcessorMask As Long
    dwNumberOrfProcessors As Long
    dwProcessorType As Long
    dwAllocationGranularity As Long
    wProcessorLevel As Integer
    wProcessorRevision As Integer
End Type

Enum results
    r_32bit = 0
    r_64bit = 1
    r_other = 2
    r_error = 3
End Enum

Private m_WindowsIs64Bit As Boolean
Private m_SeDebugEnabled As Boolean

Property Get isWindows64Bit() As Boolean
    isWindows64Bit = m_WindowsIs64Bit
End Property

Property Get isSeDebugEnabled() As Boolean
    isSeDebugEnabled = m_SeDebugEnabled
End Property

Private Sub Class_Initialize()
    m_SeDebugEnabled = GetSeDebug()   'some x64 bit processes require this access to check..
    m_WindowsIs64Bit = Is64BitProcessor()
End Sub

Private Function Is64BitProcessor() As Boolean

    Const PROCESSOR_ARCHITECTURE_AMD64 As Integer = 9
    Const PROCESSOR_ARCHITECTURE_IA64 As Integer = 6
    Dim si As SYSTEM_INFO
    
    Dim lpfn As Long

    lpfn = GetProcAddress(GetModuleHandle("kernel32"), "GetNativeSystemInfo")

    If lpfn = 0 Then Exit Function
      
    GetNativeSystemInfo si
   
    Is64BitProcessor = (si.wProcessorArchitecture = PROCESSOR_ARCHITECTURE_AMD64 _
        Or _
            si.wProcessorArchitecture = PROCESSOR_ARCHITECTURE_IA64)
            
End Function


Function IsProcess_x64(pid As Long) As results

    Dim ret As Long
    Dim lpfn As Long
    Dim h As Long
    
    Const PROCESS_QUERY_INFORMATION = (&H400)
    
    If Not m_WindowsIs64Bit Then
        IsProcess_x64 = r_32bit
        Exit Function
    End If
    
    lpfn = GetProcAddress(GetModuleHandle("kernel32"), "IsWow64Process")

    If lpfn = 0 Then
        IsProcess_x64 = r_error
        Exit Function
    End If
    
    h = OpenProcess(PROCESS_QUERY_INFORMATION, False, pid)
    
    If h = 0 Then
        IsProcess_x64 = r_error
        Exit Function
    End If

    IsWow64Process h, ret
    CloseHandle h
    
    If ret = 1 Then
        IsProcess_x64 = r_64bit
    Else
        IsProcess_x64 = r_32bit
    End If
    
End Function


Function isExe_x64(fpath As String) As results
    On Error GoTo hell
    Dim h As Long
    Dim b As Byte
    Dim b2 As Byte
    Dim lng As Long
    Dim i As Integer
    
    Const AMD64 = &H8664
    Const I386 = &H14C
    Const IA64 = &H200
    
    h = FreeFile
    Open fpath For Binary Access Read As h
    Get h, , b
    Get h, , b2
    
    If b <> Asc("M") Then GoTo hell
    If b2 <> Asc("Z") Then GoTo hell
    
    Get h, 61, lng 'DOSHeader.e_lfanew raw offset
    
    Get h, lng + 1, b
    Get h, , b2
    
    If b <> Asc("P") Then GoTo hell  'IMAGE_NT_HEADERS.Signature
    If b2 <> Asc("E") Then GoTo hell
    
    Get h, lng + 5, i
    
    'Debug.Print Hex(i)
    
    If i = AMD64 Or i = IA64 Then
        isExe_x64 = r_64bit
    ElseIf i = I386 Then
        isExe_x64 = r_32bit
    Else
        isExe_x64 = r_other
    End If
    
    Close h

    Exit Function
hell:
    On Error Resume Next
    Close h
    isExe_x64 = r_error
    Err.Clear
End Function


Private Function GetSeDebug() As Boolean
    Dim hToken As Long, hProcess As Long, lret As Long
    Dim tkp As TOKEN_PRIVILEGES
    
    Const TOKEN_ADJUST_PRIVILEGES As Long = &H20
    Const TOKEN_QUERY As Long = &H8

    hProcess = GetCurrentProcess()
    OpenProcessToken hProcess, TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, hToken
    LookupPrivilegeValue "", "SeDebugPrivilege", tkp.TheLuid

    tkp.PrivilegeCount = 1
    tkp.Attributes = 2 'SE_PRIVILEGE_ENABLED
    
    If AdjustTokenPrivileges(hToken, False, tkp, Len(tkp), tkp, lret) = 0 Then Exit Function
        
    If GetLastError = 0 Then GetSeDebug = True

End Function


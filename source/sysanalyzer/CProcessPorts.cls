VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CProcessPorts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
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
'
'BUG FIX:  XP apparently keeps a second copy of all port handles with pid 4?
'          had to reset stored values to latter pid to get real one?
'
'BUG:      Does not list ports opened by SYSTEM under XP?
 
Private Type LUID
    LowPart As Long
    HighPart As Long
End Type

Private Type LUID_AND_ATTRIBUTES
    pLuid As LUID
    Attributes As Long
End Type

Private Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    TheLuid As LUID
    Attributes As Long
End Type

Private Type SYSTEM_HANDLE_TABLE_ENTRY_INFO
    UniqueProcessId  As Integer
    CreatorBackTraceIndex  As Integer
    ObjectTypeIndex As Byte
    HandleAttributes As Byte
    HandleValue As Integer
    Object  As Long
    GrantedAccess As Long
End Type

Private Type SYSTEM_HANDLE_INFORMATION
    NumberOfHandles As Long
    Handles() As SYSTEM_HANDLE_TABLE_ENTRY_INFO
End Type

Private Type OBJECT_NAME_PRIVATE
    Length As Integer
    MaximumLength As Integer
    Buffer As Long
    ObjName(23) As Byte
End Type

Private Type TDI_CONNECTION_INFO
    State As Long
    Event As Long
    TransmittedTsdus As Long
    ReceivedTsdus As Long
    TransmissionErrors As Long
    ReceiveErrors As Long
    ThroughputLow As Long
    ThroughputHigh As Long
    DelayLow As Long
    DelayHigh As Long
    SendBufferSize As Long
    ReceiveBufferSize As Long
    Unreliable As Boolean
End Type

Private Type TDI_CONNECTION_INFORMATION
    UserDataLength As Long
    UserData As Long
    OptionsLength As Long
    Options As Long
    RemoteAddressLength As Long
    RemoteAddress As Long
End Type

Private Type IO_STATUS_BLOCK
    Status As Long
    Information As Long
End Type

Private Declare Function NtQueryObject Lib "ntdll.dll" (ByVal ObjectHandle As Long, ByVal ObjectInformationClass As Long, ObjectInformation As OBJECT_NAME_PRIVATE, ByVal Length As Long, ResultLength As Long) As Long
Private Declare Function NtDeviceIoControlFile Lib "ntdll.dll" (ByVal FileHandle As Long, ByVal pEvent As Long, ApcRoutine As Long, ApcContext As Long, IoStatusBlock As IO_STATUS_BLOCK, ByVal IoControlCode As Long, InputBuffer As TDI_CONNECTION_INFORMATION, ByVal InputBufferLength As Long, OutputBuffer As TDI_CONNECTION_INFO, ByVal OutputBufferLength As Long) As Long
Private Declare Function NtQuerySystemInformation Lib "ntdll.dll" (ByVal dwInfoType As Long, ByRef lpStructure As Byte, ByVal dwSize As Long, dwReserved As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32.dll" () As Long
Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function DuplicateHandle Lib "kernel32" (ByVal hSourceProcessHandle As Long, ByVal hSourceHandle As Long, ByVal hTargetProcessHandle As Long, lpTargetHandle As Long, ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwOptions As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub ToStruc Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef struc As Any, ByRef buf As Byte, ByVal Length As Long)
Private Declare Function EnumProcessModules Lib "PSAPI.DLL" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Declare Function ntohs Lib "ws2_32.dll" (ByVal netshort As Integer) As Integer
Private Declare Function GetModuleFileNameExA Lib "PSAPI.DLL" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, ByRef NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, ByRef PreviousState As TOKEN_PRIVILEGES, ByRef ReturnLength As Long) As Long
Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, ByRef TokenHandle As Long) As Long
Private Declare Function GetLastError Lib "kernel32.dll" () As Long

Private ports As Collection 'of cprocessport

'Function CompareSnapShots(snap1 As Collection, snap2 As Collection) As Collection
'    MsgBox "todo"
'End Function

'there is a bug in here under XP and XPsp2 now CRRRASH :(
Function TakePortProcSnapshot() As Collection ' of CProcessPort
    Dim ret As Long, hCount As Long, l As Long
    Dim shi As SYSTEM_HANDLE_INFORMATION
    Dim buf() As Byte
            
    Set ports = New Collection
    
    If Not GetSeDebug() Then MsgBox "Couldnot obtain SeDebugPrivledge, you should run this app as administrator", vbInformation
    
    l = 20
    ReDim buf(l)
    
    If NtQuerySystemInformation(16, buf(0), ByVal l, 0) = 0 Then Exit Function
    
    ToStruc hCount, buf(0), 4
    If hCount < 1 Then Exit Function
    l = hCount * 16 + 4
    ReDim buf(0 To l)

    NtQuerySystemInformation 16, buf(0), ByVal l, 0

    ReDim shi.Handles(hCount)
    ToStruc shi.Handles(0), buf(4), l - 4
    
    Dim tmp() As String
    
    For l = 0 To hCount - 1
       ExamineHandle shi.Handles(l).UniqueProcessId, shi.Handles(l).HandleValue
    Next

    Call SortPorts
    Set TakePortProcSnapshot = ports
 
    
End Function




Private Sub ExamineHandle(mpid As Integer, hCurrent As Integer)

    Dim hPort As Long, Port As Long, l As Long, hProc As Long, key As String
    Dim onp As OBJECT_NAME_PRIVATE, proc As CProcessPort
    Dim pid As Integer
    
    
    Const OBJNAMEINFO = 1
    Const PROCESS_DUP_HANDLE = &H40
    Const STANDARD_RIGHTS_ALL = &H1F0000
    Const GENERIC_ALL = &H10000000
    
    pid = mpid
    If pid < 1 Then Exit Sub
    
    hProc = OpenProcess(PROCESS_DUP_HANDLE, 0&, pid)
    If hProc < 1 Then Exit Sub

    If DuplicateHandle(hProc, hCurrent, -1, hPort, STANDARD_RIGHTS_ALL Or GENERIC_ALL, 0, 0) Then
        If NtQueryObject(hPort, OBJNAMEINFO, onp, LenB(onp), l) = 0 Then
            If InStr(onp.ObjName, "Device") > 0 Then
                    Port = GetPort(hPort)
                    If Port Then
                        Set proc = New CProcessPort
                        proc.pid = mpid
                        proc.Port = Port
                        proc.path = GetProcessPath(pid)
                        proc.isTCP = IIf(InStr(onp.ObjName, "Tcp"), True, False)
                        key = IIf(proc.isTCP, "TCPPORT:", "UDPPORT:") & Port
                        If Not KeyExistsInCollection(ports, key) Then
                            ports.Add proc, key
                        Else
                            If ports(key).pid = 4 Then 'XP pid 4 keeps copy of all so we have to overwrite
                                ports.Remove key
                                ports.Add proc, key
                            End If
                        End If
                    End If
            End If

        End If

    End If
    
    CloseHandle hPort
    CloseHandle hProc

End Sub

Private Function GetPort(hPort As Long) As Long

    Dim tci As TDI_CONNECTION_INFO
    Dim tcInfo As TDI_CONNECTION_INFORMATION
    Dim io As IO_STATUS_BLOCK
    Dim ret As Long
    On Error Resume Next
    
    tcInfo.RemoteAddressLength = 3
    
    If hPort < 1 Then Exit Function
    
    fso.AppendFile "c:\log.txt", "hport=" & hPort
    
    'XP is barfing here at controlfile sometimes
    ret = NtDeviceIoControlFile(hPort, 0, 0, 0, io, &H210012, tcInfo, Len(tcInfo), tci, Len(tci))
    
    If ret <> 0 Then Exit Function
    GetPort = ntohs(UnsignedToInteger(tci.ReceivedTsdus And 65535))
    If GetPort < 0 Then GetPort = GetPort + 65536
    
End Function

Private Function GetProcessPath(pid As Integer) As String
    Dim hProc As Long
    Dim hMods() As Long, cbAlloc As Long, ret As Long, retMax As Long
    Dim sPath As String
    Const PROCESS_VM_READ = (&H10)
    Const PROCESS_QUERY_INFORMATION = (&H400)
    hProc = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, False, pid)
    
    If hProc <> 0 Then
        cbAlloc = 2000
        ReDim hMods(cbAlloc)
        ret = EnumProcessModules(hProc, hMods(0), (cbAlloc * 4), retMax)
        
        sPath = Space$(2000)
        ret = GetModuleFileNameExA(hProc, hMods(0), sPath, 1260)
        If ret > 0 Then
            GetProcessPath = Left$(sPath, ret)
            If Len(GetProcessPath) = 0 Then GetProcessPath = "SYSTEM"
        End If
        Call CloseHandle(hProc)
       
    End If
    
    
End Function

Private Function KeyExistsInCollection(c As Collection, val As String) As Boolean
    On Error GoTo nope
    Dim t
    Set t = c(val)
    KeyExistsInCollection = True
 Exit Function
nope: KeyExistsInCollection = False
End Function

Private Function UnsignedToInteger(value As Long) As Integer
    If value <= 32767 Then UnsignedToInteger = value _
    Else UnsignedToInteger = value - 65536
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
    
    AdjustTokenPrivileges hToken, False, tkp, Len(tkp), tkp, lret
        
    If GetLastError = 0 Then GetSeDebug = True

End Function

Private Sub SortPorts()
    Dim c1 As New Collection
    Dim c2 As New Collection
    Dim p As CProcessPort
    Dim p2 As CProcessPort
    
    Dim i As Long
    On Error Resume Next
    
    'first sort by port number (udp and tcp both at once)
    For Each p In ports
        If c1.Count < 1 Then
            c1.Add p
        Else
            i = 0
            For Each p2 In c1
                i = i + 1
                If p.Port < p2.Port Then
                    c1.Add p, , i
                    Exit For
                End If
                If i = c1.Count Then c1.Add p
            Next
        End If
    Next
    
    Set ports = New Collection
   
    'add all tcp ports to final collection, store sorted udp in tmp collection
    For Each p In c1
        If Not p.isTCP Then c2.Add p Else ports.Add p, "TCPPORT:" & p.Port
    Next
     
    'add sorted udp ports to end of final collection
    For Each p In c2
        ports.Add p, "UDPPORT:" & p.Port
    Next
    
End Sub

Private Sub Class_Initialize()
    GetSeDebug
End Sub

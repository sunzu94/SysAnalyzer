VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSysDiff"
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

Public CProc As New CProcessInfo
Dim CPorts As New CProcessPorts
Dim CReg As New CRegDiff

Dim Processes1 As Collection
Dim Processes2 As Collection
Dim ChangedProcesses As Collection

Dim Ports1 As Collection
Dim ports2 As Collection
Dim ChangedPorts As Collection

Dim Services1 As New Collection  'unused
Dim Services2 As New Collection  'unused
Dim ChangedServices As Collection

Dim ExplorerDlls1 As Collection
Dim Explorerdlls2 As Collection
Dim ChangedExplorerDlls As Collection

Dim IEDlls1 As Collection
Dim IEDlls2 As Collection
Dim ChangedIEDlls As Collection

Dim Drivers1 As Collection
Dim Drivers2 As Collection
Dim ChangedDrivers As Collection

Dim Mutex1 As Collection
Dim Mutex2 As Collection
Dim ChangedMutexes As Collection

Dim Tasks1 As Collection
Dim Tasks2 As Collection
Dim ChangedTasks As Collection

Dim liProc As ListItem
Dim liExplorer As ListItem
Dim liIE As ListItem
Dim explorerPopup As Boolean

Public shutDown As Boolean


Sub ShowBaseSnap(Optional FirstTabOnly = False)
    If Processes1 Is Nothing Then
        MsgBox "You must first take the base snapshot", vbInformation
        Exit Sub
    End If
    ShowSnapShot Processes1, Ports1, Services1, ExplorerDlls1, IEDlls1, Mutex1, Tasks1, False, Drivers1, CReg.Snap1, FirstTabOnly
End Sub

Sub ShowSnap2()
    If Processes2 Is Nothing Then
        MsgBox "You must first take second snapshot", vbInformation
        Exit Sub
    End If
    ShowSnapShot Processes2, ports2, Services2, Explorerdlls2, IEDlls2, Mutex2, Tasks2, False, Drivers2, CReg.Snap2
End Sub


Sub DoSnap1()

    On Error Resume Next
    
    Dim col As Collection
    
    frmMain.lvProcesses.Clear
    frmMain.lvPorts.ListItems.Clear
    frmMain.lvExplorer.ListItems.Clear
    frmMain.lvIE.ListItems.Clear
    frmMain.lvDrivers.Clear
    frmMain.lvRegKeys.ListItems.Clear
    frmMain.lvTasks.ListItems.Clear
    frmMain.lvMutex.Clear
    
    frmMain.Display = "Taking SnapShot 1"
    frmMain.pb.max = 7
    frmMain.pb.value = 0
    
    Sleep 100
    DoEvents
    
    frmMain.Display = "Enumerating processes..."
    Set Processes1 = CProc.GetRunningProcesses()
    frmMain.pb.value = frmMain.pb.value + 1
    If shutDown Then Exit Sub
    
    Sleep 100
    DoEvents
    frmMain.Display = "Enumerating open ports..."
    Set Ports1 = CPorts.TakePortProcSnapshot()
    frmMain.pb.value = frmMain.pb.value + 1
    If shutDown Then Exit Sub
    
    Sleep 100
    DoEvents
    frmMain.Display = "Enumerating loaded drivers..."
    Set Drivers1 = CProc.EnumDrivers()
    frmMain.pb.value = frmMain.pb.value + 1
    If shutDown Then Exit Sub
     
    Sleep 100
    DoEvents
    frmMain.Display = "Enumerating mutexes..."
    Set Mutex1 = CProc.EnumMutexes()
    frmMain.pb.value = frmMain.pb.value + 1
    If shutDown Then Exit Sub
    
    Sleep 100
    DoEvents
    frmMain.Display = "Enumerating tasks..."
    Set Tasks1 = CProc.EnumTasks()
    frmMain.pb.value = frmMain.pb.value + 1
    If shutDown Then Exit Sub
    
    frmMain.Display = "Enumerating registry keys..."
    DoEvents
    CReg.Reset
    CReg.DoSnap1
    frmMain.pb.value = frmMain.pb.value + 1
    If shutDown Then Exit Sub
    
    Dim explorerPID As Long
    Dim IEPID As Long
    
    explorerPID = FindPidForExeName("explorer.exe", Processes1)
    IEPID = FindPidForExeName("iexplore.exe", Processes1)
     
    frmMain.Display = "Scanning explorer loaded dlls..."
    DoEvents
    If explorerPID = 0 Then
        Set ExplorerDlls1 = New Collection
    Else
        Set ExplorerDlls1 = CProc.GetProcessModules(explorerPID)
    End If
    frmMain.pb.value = frmMain.pb.value + 1
    If shutDown Then Exit Sub
    
    frmMain.Display = "Scanning IE loaded dlls..."
    DoEvents
    If IEPID = 0 Then
        Set IEDlls1 = New Collection
    Else
        Set IEDlls1 = CProc.GetProcessModules(IEPID)
    End If
    
    frmMain.pb.value = 0
    frmMain.Display = Empty
    
End Sub

Sub DoSnap2()

    'XP can get weird crashs in here seems to be around the
    'CprocessPorts::NtDeviceIoControlFile.GetPort.NtDeviceIoControlFile
    'somehow these sleeps seem to stop it for me..buggyyyyyyy :_(
    
    On Error Resume Next
    
    frmMain.pb.max = 8
    frmMain.pb.value = 0
    frmMain.Display = "Taking Snapshot 2"
    
    Sleep 100
    DoEvents
    frmMain.Display = "Enumerating processes..."
    Set Processes2 = CProc.GetRunningProcesses()
    frmMain.pb.value = frmMain.pb.value + 1
    If shutDown Then Exit Sub
    
    Sleep 100
    DoEvents
    frmMain.Display = "Enumerating open ports..."
    Set ports2 = CPorts.TakePortProcSnapshot()
    frmMain.pb.value = frmMain.pb.value + 1
    If shutDown Then Exit Sub
    
    Sleep 100
    DoEvents
    frmMain.Display = "Enumerating loaded drivers..."
    Set Drivers2 = CProc.EnumDrivers()
    frmMain.pb.value = frmMain.pb.value + 1
    If shutDown Then Exit Sub
     
    Sleep 100
    DoEvents
    frmMain.Display = "Enumerating mutexes..."
    Set Mutex2 = CProc.EnumMutexes()
    frmMain.pb.value = frmMain.pb.value + 1
    If shutDown Then Exit Sub
 
    Sleep 100
    DoEvents
    frmMain.Display = "Enumerating tasks..."
    Set Tasks2 = CProc.EnumTasks()
    frmMain.pb.value = frmMain.pb.value + 1
    If shutDown Then Exit Sub

    frmMain.Display = "Enumerating registry keys..."
    DoEvents
    
    CReg.DoSnap2
    If shutDown Then Exit Sub
    frmMain.pb.value = frmMain.pb.value + 1
    
    Dim explorerPID As Long
    Dim IEPID As Long
   
    explorerPID = FindPidForExeName("explorer.exe", Processes2)
    IEPID = FindPidForExeName("iexplore.exe", Processes2)
    
    frmMain.Display = "Scanning explorer loaded dlls..."
    DoEvents
    
    If explorerPID = 0 Then
        Set Explorerdlls2 = New Collection
    Else
        Set Explorerdlls2 = CProc.GetProcessModules(explorerPID)
    End If
    
    If shutDown Then Exit Sub
    frmMain.pb.value = frmMain.pb.value + 1
    
    frmMain.Display = "Scanning IE loaded dlls..."
    DoEvents
    
    If IEPID = 0 Then
        Set IEDlls2 = New Collection
    Else
        Set IEDlls2 = CProc.GetProcessModules(IEPID)
    End If
   
    frmMain.pb.value = 0
    
End Sub

Function FindPidForExeName(exename As String, col As Collection) As Long
    Dim proc As CProcess
    For Each proc In col
        If InStr(1, proc.path, exename, vbTextCompare) > 0 Then
            FindPidForExeName = proc.pid
            Exit Function
        End If
    Next
End Function


Sub ShowDiffReport()
        
    On Error Resume Next
    
    If Processes1 Is Nothing Then
        MsgBox "Take base Snapshot first", vbInformation
        Exit Sub
    End If
    
    If Processes2 Is Nothing Then
        MsgBox "Take second Snapshot first", vbInformation
        Exit Sub
    End If
    
    frmMain.Display = "Showing Diff report"
    
    Set ChangedProcesses = New Collection
    Set ChangedPorts = New Collection
    Set ChangedServices = New Collection
    Set ChangedExplorerDlls = New Collection
    Set ChangedIEDlls = New Collection
    Set ChangedDrivers = New Collection
    Set ChangedMutexes = New Collection
    Set ChangedTasks = New Collection
    
    Dim t As CTaskElem
    Dim m As CMutexElem
    Dim proc As CProcess
    Dim Port As CProcessPort
    Dim cmod As CModule
    Dim li As ListItem
    Dim spec As String
    Dim drv As Variant
    
    Dim r() As String
    
    For Each proc In Processes2
        If Not objKeyExistsInCollection(Processes1, "pid:" & proc.pid) Then
            ChangedProcesses.Add proc
        End If
        If shutDown Then Exit Sub
    Next
    
    For Each Port In ports2
        spec = IIf(Port.isTCP, "TCPPORT:", "UDPPORT:") & Port.Port
        If Not objKeyExistsInCollection(Ports1, spec) Then
            ChangedPorts.Add Port
        End If
        If shutDown Then Exit Sub
    Next
    
    For Each cmod In Explorerdlls2
        If Not objKeyExistsInCollection(ExplorerDlls1, cmod.path) Then
            ChangedExplorerDlls.Add cmod
        End If
        If shutDown Then Exit Sub
    Next
   
    For Each cmod In IEDlls2
        If Not objKeyExistsInCollection(IEDlls1, cmod.path) Then
            ChangedIEDlls.Add cmod
        End If
        If shutDown Then Exit Sub
    Next
    
    For Each drv In Drivers2
        If Not KeyExistsInCollection(Drivers1, CStr(drv)) Then
            ChangedDrivers.Add drv
        End If
        If shutDown Then Exit Sub
    Next
        
    Dim c1 As Collection
    Dim c2 As Collection
    
    'because we delete elements from them in the diff to track..
    Set c1 = CloneMutexCollection(Mutex1)
    Set c2 = CloneMutexCollection(Mutex2)
    
    For Each m In c2
        If objKeyExistsInCollection(c1, m.getKey()) Then
            c1.Remove m.getKey()
            c2.Remove m.getKey()
            Set m = Nothing
        Else
            m.isNew = True
            ChangedMutexes.Add m, m.getKey()
        End If
    Next
    
    'these original mutexes no longer exist (does this even matter?)
    'For Each m In c1
    '    ChangedMutexes.Add m, m.getKey()
    'Next
        
    For Each t In Tasks2
        If Not objKeyExistsInCollection(Tasks1, t.hashCode) Then
            ChangedTasks.Add t, t.hashCode
        End If
    Next

    ShowSnapShot ChangedProcesses, ChangedPorts, ChangedServices, ChangedExplorerDlls, ChangedIEDlls, ChangedMutexes, ChangedTasks, True, ChangedDrivers, CReg.GetDifferences


End Sub


Private Sub ShowSnapShot(procs As Collection, ports As Collection, servcs As Collection, expdlls As Collection, _
                        iedlls As Collection, mutexes As Collection, tasks As Collection, Optional isCompare As Boolean = False, Optional drivers As Collection, _
                        Optional regSnap As Collection, Optional FirstTabOnly = False)
   
    Dim proc As CProcess
    Dim Port As CProcessPort
    'Dim servc As CService
    Dim n As CModule
    Dim d
    Dim mm As matchModes
    Dim cme As CMutexElem
    
    Dim li As ListItem
    Dim fp As FILEPROPERTIE
    Dim root As String
    
    Const my_orange As Long = vbRed
    Const my_green As Long = &HC00000
    
    On Error Resume Next
    'If known.Loaded And known.Ready Then ado.OpenConnection
    
    frmMain.Display = "Showing ShapShot"
    frmMain.pb.value = 0
    frmMain.pb.max = 8
    
    frmMain.lvProcesses.Clear
    frmMain.lvPorts.ListItems.Clear
    'frmMain.lvServices.ListItems.Clear
    frmMain.lvExplorer.ListItems.Clear
    frmMain.lvIE.ListItems.Clear
    frmMain.lvDrivers.Clear
    frmMain.lvRegKeys.ListItems.Clear
    frmMain.lvMutex.Clear
    frmMain.lvTasks.ListItems.Clear
    
    frmMain.Display = "Loading processes.."
    For Each proc In procs
        mm = not_found
        If known.Ready And known.Loaded And Not FirstTabOnly Then
            mm = known.isFileKnown(proc.fullpath) 'CProc.GetProcessPath(proc.pid))
            If known.HideKnownInDisplays And mm = exact_match Then GoTo next_proc
        End If
        'Set li = frmMain.lvProcesses.ListItems.Add(, , pad(proc.pid))
        Set li = frmMain.lvProcesses.AddItem(pad(proc.pid))
        li.Tag = proc.pid
        li.subItems(1) = pad(proc.ParentPID)
        li.subItems(2) = proc.User & IIf(proc.is64Bit, "  *64 ", Empty)
        li.subItems(3) = proc.fullpath ' CProc.GetProcessPath(proc.pid)
        li.ToolTipText = proc.CmdLine
        If mm = exact_match Then SetLiColor li, my_green
        If mm = hash_changed Then SetLiColor li, my_orange
next_proc:
        DoEvents
        If shutDown Then Exit Sub
    Next
    frmMain.pb.value = frmMain.pb.value + 1
    
    frmMain.Display = "Loading open ports.."
    For Each Port In ports
        mm = not_found
        If known.Ready And known.Loaded And Not FirstTabOnly Then
            mm = known.isFileKnown(Port.path)
            If known.HideKnownInDisplays And mm = exact_match Then GoTo next_port
        End If
        Set li = frmMain.lvPorts.ListItems.Add(, , pad(Port.Port))
        li.subItems(1) = pad(Port.pid)
        li.subItems(2) = IIf(Port.isTCP, "TCP", "UDP")
        li.subItems(3) = Port.path
        If mm = exact_match Then SetLiColor li, my_green
        If mm = hash_changed Then SetLiColor li, my_orange
next_port:
        DoEvents
        If shutDown Then Exit Sub
    Next

    If FirstTabOnly Then
        frmMain.pb.value = 0
        Exit Sub
    End If
    
    frmMain.pb.value = frmMain.pb.value + 1
    'For Each servc In servcs
    '    Set li = frmMain.lvServices.ListItems.Add(, , servc.State)
    '    li.SubItems(1) = servc.DisplayName
    '    li.SubItems(2) = servc.ServiceName
    'Next
   
    frmMain.Display = "Loading explorer dlls.."
    For Each n In expdlls
        mm = not_found
        If known.Ready And known.Loaded Then
            mm = known.isFileKnown(n)
            If known.HideKnownInDisplays And mm = exact_match Then GoTo next_expdll
        End If
        Set li = frmMain.lvExplorer.ListItems.Add(, , n)
        fp = FileInfo(n)
        li.subItems(1) = fp.CompanyName
        li.subItems(2) = fp.FileDescription
        If mm = exact_match Then SetLiColor li, my_green
        If mm = hash_changed Then SetLiColor li, my_orange
next_expdll:
        DoEvents
        If shutDown Then Exit Sub
    Next
   
    frmMain.pb.value = frmMain.pb.value + 1
       
    frmMain.Display = "Loading ie dlls.."
    For Each n In iedlls
        mm = not_found
        If known.Ready And known.Loaded Then
            mm = known.isFileKnown(n)
            If known.HideKnownInDisplays And mm = exact_match Then GoTo next_iedll
        End If
        Set li = frmMain.lvIE.ListItems.Add(, , n)
        fp = FileInfo(n)
        li.subItems(1) = fp.CompanyName
        li.subItems(2) = fp.FileDescription
        If mm = exact_match Then SetLiColor li, my_green
        If mm = hash_changed Then SetLiColor li, my_orange
next_iedll:
        DoEvents
        If shutDown Then Exit Sub
    Next
    
    frmMain.pb.value = frmMain.pb.value + 1
        
    frmMain.Display = "Loading drivers.."
    For Each d In drivers
        root = Environ("WINDIR") & "\"
        If Len(root) = 1 Then
            root = "C:\Winnt\"
            If Not fso.FileExists(root) Then root = "C:\Windows\"
        End If
        d = Replace(d, "\SystemRoot\", root)
        d = Replace(d, "\??\", Empty)
        If Not FileExists(d) Then
            If VBA.Left(d, 2) = "\W" Then d = "c:" & d
            If InStr(1, d, "c:\", vbTextCompare) < 1 Then
                If FileExists(root & d) Then
                    d = root & d
                ElseIf FileExists(root & "system32\" & d) Then
                    d = root & "system32\" & d
                ElseIf FileExists(root & "system32\drivers\" & d) Then
                    d = root & "system32\drivers\" & d
                ElseIf FileExists(root & "system\" & d) Then
                    d = root & "system\" & d
                End If
            End If
        End If
        
        mm = not_found
        If known.Ready And known.Loaded Then
            mm = known.isFileKnown(d)
            If known.HideKnownInDisplays And mm = exact_match Then GoTo next_driver
        End If
        
        Set li = frmMain.lvDrivers.AddItem(d)
        
        'this can cause crash
        If FileExists(d) Then
            'debugLog "ShowSnapShot.Drivers.FileInfo(" & d & ")"
            fp = FileInfo(d)
            li.subItems(1) = fp.CompanyName
            li.subItems(2) = fp.FileDescription
        End If
        
        If mm = exact_match Then SetLiColor li, my_green
        If mm = hash_changed Then SetLiColor li, my_orange
        
next_driver:
        DoEvents
        If shutDown Then Exit Sub
        
    Next
    
    frmMain.pb.value = frmMain.pb.value + 1
    frmMain.Display = "Loading registry entries.."
    
    For Each d In regSnap
        Dim s As String
        s = CReg.GetValForKey(d)
        Set li = frmMain.lvRegKeys.ListItems.Add(, , d)
        li.subItems(1) = s
        If shutDown Then Exit Sub
    Next
    
    frmMain.pb.value = frmMain.pb.value + 1
    frmMain.Display = "Loading mutexes.."
    
    For Each cme In mutexes
        Set li = frmMain.lvMutex.AddItem(cme.pid)
        If isCompare Then
            li.subItems(1) = IIf(cme.isNew, "", "- ") & cme.name
        Else
            li.subItems(1) = cme.name
        End If
        If shutDown Then Exit Sub
    Next
    
    frmMain.pb.value = frmMain.pb.value + 1
    frmMain.Display = "Loading tasks.."
    
    Dim cte As CTaskElem
    For Each cte In tasks
        Set li = frmMain.lvTasks.ListItems.Add(, , cte.name)
        li.subItems(1) = cte.exe
        Set li.Tag = cte
        If shutDown Then Exit Sub
    Next
    
    frmMain.pb.value = frmMain.pb.value + 1
        
    If iedlls.count = 0 Then frmMain.lvIE.ListItems.Add , , IIf(isCompare, "No changes Found", "IE PID Not found")
    If expdlls.count = 0 Then frmMain.lvExplorer.ListItems.Add , , IIf(isCompare, "No changes Found", "Explorer PID Not found?")
    
    frmMain.pb.value = 0
    'If known.Loaded And known.Ready Then ado.CloseConnection
End Sub

'for listview sorting...
Private Function pad(v, Optional l As Long = 5)
    On Error GoTo hell
    Dim x As Long
    x = Len(v)
    If x < l Then
        pad = String(l - x, " ") & v
    Else
hell:
        pad = v
    End If
End Function

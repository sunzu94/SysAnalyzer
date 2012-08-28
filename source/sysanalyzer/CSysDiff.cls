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

Dim liProc As ListItem
Dim liExplorer As ListItem
Dim liIE As ListItem
Dim explorerPopup As Boolean

Sub ShowBaseSnap()
    If Processes1 Is Nothing Then
        MsgBox "You must first take the base snapshot", vbInformation
        Exit Sub
    End If
    ShowSnapShot Processes1, Ports1, Services1, ExplorerDlls1, IEDlls1, False, Drivers1, CReg.Snap1
End Sub

Sub ShowSnap2()
    If Processes2 Is Nothing Then
        MsgBox "You must first take second snapshot", vbInformation
        Exit Sub
    End If
    ShowSnapShot Processes2, ports2, Services2, Explorerdlls2, IEDlls2, False, Drivers2, CReg.Snap2
End Sub


Sub DoSnap1()

    Dim col As Collection
    
    frmMain.lvProcesses.ListItems.Clear
    frmMain.lvPorts.ListItems.Clear
    frmMain.lvExplorer.ListItems.Clear
    frmMain.lvIE.ListItems.Clear
    frmMain.lvDrivers.ListItems.Clear
    frmMain.lvRegKeys.ListItems.Clear
    
    Sleep 100
    DoEvents
    Set Processes1 = CProc.GetRunningProcesses()
    
    Sleep 100
    DoEvents
    Set Ports1 = CPorts.TakePortProcSnapshot()
    
    Sleep 100
    DoEvents
    Set Drivers1 = CProc.EnumDrivers()
    
    CReg.Reset
    CReg.DoSnap1
    
    Dim explorerPID As Long
    Dim IEPID As Long
    
    explorerPID = FindPidForExeName("explorer.exe", Processes1)
    IEPID = FindPidForExeName("iexplore.exe", Processes1)
     
    If explorerPID = 0 Then
        Set ExplorerDlls1 = New Collection
    Else
        Set ExplorerDlls1 = CProc.GetProcessModules(explorerPID)
    End If
    
    If IEPID = 0 Then
        Set IEDlls1 = New Collection
    Else
        Set IEDlls1 = CProc.GetProcessModules(IEPID)
    End If
    
End Sub

Sub DoSnap2()

    'XP can get weird crashs in here seems to be around the
    'CprocessPorts::NtDeviceIoControlFile.GetPort.NtDeviceIoControlFile
    'somehow these sleeps seem to stop it for me..buggyyyyyyy :_(
    
    Sleep 100
    DoEvents
    Set Processes2 = CProc.GetRunningProcesses()
    
    Sleep 100
    DoEvents
    Set ports2 = CPorts.TakePortProcSnapshot()
    
    Sleep 100
    DoEvents
    Set Drivers2 = CProc.EnumDrivers()
    
    CReg.DoSnap2
    
    Dim explorerPID As Long
    Dim IEPID As Long
   
    explorerPID = FindPidForExeName("explorer.exe", Processes2)
    
    IEPID = FindPidForExeName("iexplore.exe", Processes2)
    
    If explorerPID = 0 Then
        Set Explorerdlls2 = New Collection
    Else
        Set Explorerdlls2 = CProc.GetProcessModules(explorerPID)
    End If
    
    If IEPID = 0 Then
        Set IEDlls2 = New Collection
    Else
        Set IEDlls2 = CProc.GetProcessModules(IEPID)
    End If
   
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
    
    If Processes1 Is Nothing Then
        MsgBox "Take base Snapshot first", vbInformation
        Exit Sub
    End If
    
    If Processes2 Is Nothing Then
        MsgBox "Take second Snapshot first", vbInformation
        Exit Sub
    End If
    
    Set ChangedProcesses = New Collection
    Set ChangedPorts = New Collection
    Set ChangedServices = New Collection
    Set ChangedExplorerDlls = New Collection
    Set ChangedIEDlls = New Collection
    Set ChangedDrivers = New Collection
    
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
    Next
    
    For Each Port In ports2
        spec = IIf(Port.isTCP, "TCPPORT:", "UDPPORT:") & Port.Port
        If Not objKeyExistsInCollection(Ports1, spec) Then
            ChangedPorts.Add Port
        End If
    Next
    
    For Each cmod In Explorerdlls2
        If Not objKeyExistsInCollection(ExplorerDlls1, cmod.path) Then
            ChangedExplorerDlls.Add cmod
        End If
    Next
   
    For Each cmod In IEDlls2
        If Not objKeyExistsInCollection(IEDlls1, cmod.path) Then
            ChangedIEDlls.Add cmod
        End If
    Next
    
    For Each drv In Drivers2
        If Not KeyExistsInCollection(Drivers1, CStr(drv)) Then
            ChangedDrivers.Add drv
        End If
    Next
        
     ShowSnapShot ChangedProcesses, ChangedPorts, ChangedServices, ChangedExplorerDlls, ChangedIEDlls, True, ChangedDrivers, CReg.GetDifferences


End Sub


Private Sub ShowSnapShot(procs As Collection, ports As Collection, servcs As Collection, expdlls As Collection, iedlls As Collection, Optional isCompare As Boolean = False, Optional drivers As Collection, Optional regSnap As Collection)
   
    Dim proc As CProcess
    Dim Port As CProcessPort
    'Dim servc As CService
    Dim n As CModule
    Dim d
    Dim mm As matchModes
    
    Dim li As ListItem
    Dim fp As FILEPROPERTIE
    Dim root As String
    
    Const my_orange As Long = vbRed
    Const my_green As Long = &HC00000
    
    On Error Resume Next
    ado.OpenConnection
    
    frmMain.lvProcesses.ListItems.Clear
    frmMain.lvPorts.ListItems.Clear
    'frmMain.lvServices.ListItems.Clear
    frmMain.lvExplorer.ListItems.Clear
    frmMain.lvIE.ListItems.Clear
    frmMain.lvDrivers.ListItems.Clear
    frmMain.lvRegKeys.ListItems.Clear
    
    For Each proc In procs
        mm = not_found
        If known.Ready And known.Loaded Then
            mm = known.isFileKnown(proc.fullpath) 'CProc.GetProcessPath(proc.pid))
            If known.HideKnownInDisplays And mm = exact_match Then GoTo next_proc
        End If
        Set li = frmMain.lvProcesses.ListItems.Add(, , pad(proc.pid))
        li.Tag = proc.pid
        li.SubItems(1) = pad(proc.ParentPID)
        li.SubItems(2) = proc.User & IIf(proc.is64Bit, "  *64 ", Empty)
        li.SubItems(3) = proc.fullpath ' CProc.GetProcessPath(proc.pid)
        li.ToolTipText = proc.CmdLine
        If mm = exact_match Then SetLiColor li, my_green
        If mm = hash_changed Then SetLiColor li, my_orange
next_proc:
        DoEvents
    Next

    For Each Port In ports
        mm = not_found
        If known.Ready And known.Loaded Then
            mm = known.isFileKnown(Port.path)
            If known.HideKnownInDisplays And mm = exact_match Then GoTo next_port
        End If
        Set li = frmMain.lvPorts.ListItems.Add(, , pad(Port.Port))
        li.SubItems(1) = pad(Port.pid)
        li.SubItems(2) = IIf(Port.isTCP, "TCP", "UDP")
        li.SubItems(3) = Port.path
        If mm = exact_match Then SetLiColor li, my_green
        If mm = hash_changed Then SetLiColor li, my_orange
next_port:
        DoEvents
    Next

    'For Each servc In servcs
    '    Set li = frmMain.lvServices.ListItems.Add(, , servc.State)
    '    li.SubItems(1) = servc.DisplayName
    '    li.SubItems(2) = servc.ServiceName
    'Next
   
    For Each n In expdlls
        mm = not_found
        If known.Ready And known.Loaded Then
            mm = known.isFileKnown(n)
            If known.HideKnownInDisplays And mm = exact_match Then GoTo next_expdll
        End If
        Set li = frmMain.lvExplorer.ListItems.Add(, , n)
        fp = FileInfo(n)
        li.SubItems(1) = fp.CompanyName
        li.SubItems(2) = fp.FileDescription
        If mm = exact_match Then SetLiColor li, my_green
        If mm = hash_changed Then SetLiColor li, my_orange
next_expdll:
        DoEvents
    Next
   
    For Each n In iedlls
        mm = not_found
        If known.Ready And known.Loaded Then
            mm = known.isFileKnown(n)
            If known.HideKnownInDisplays And mm = exact_match Then GoTo next_iedll
        End If
        Set li = frmMain.lvIE.ListItems.Add(, , n)
        fp = FileInfo(n)
        li.SubItems(1) = fp.CompanyName
        li.SubItems(2) = fp.FileDescription
        If mm = exact_match Then SetLiColor li, my_green
        If mm = hash_changed Then SetLiColor li, my_orange
next_iedll:
        DoEvents
    Next
    
    For Each d In drivers
        root = Environ("WINDIR") & "\"
        If Len(root) = 1 Then
            root = "C:\Winnt\"
            If Not fso.FileExists(root) Then root = "C:\Windows\"
        End If
        d = Replace(d, "\SystemRoot\", root)
        d = Replace(d, "\??\", Empty)
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
        
        mm = not_found
        If known.Ready And known.Loaded Then
            mm = known.isFileKnown(d)
            If known.HideKnownInDisplays And mm = exact_match Then GoTo next_driver
        End If
        
        Set li = frmMain.lvDrivers.ListItems.Add(, , d)
        
        'this can cause crash
        If FileExists(d) Then
            fp = FileInfo(d)
            li.SubItems(1) = fp.CompanyName
            li.SubItems(2) = fp.FileDescription
        End If
        
        If mm = exact_match Then SetLiColor li, my_green
        If mm = hash_changed Then SetLiColor li, my_orange
        
next_driver:
        DoEvents
        
    Next
    
    
    For Each d In regSnap
        Dim s As String
        s = CReg.GetValForKey(d)
        Set li = frmMain.lvRegKeys.ListItems.Add(, , d)
        li.SubItems(1) = s
    Next
    
    If iedlls.count = 0 Then frmMain.lvIE.ListItems.Add , , IIf(isCompare, "No changes Found", "IE PID Not found")
    If expdlls.count = 0 Then frmMain.lvExplorer.ListItems.Add , , IIf(isCompare, "No changes Found", "Explorer PID Not found?")
    
    ado.CloseConnection
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

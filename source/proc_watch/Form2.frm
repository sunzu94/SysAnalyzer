VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form2 
   Caption         =   "Monitoring for new Processes"
   ClientHeight    =   4365
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   6900
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   4365
   ScaleWidth      =   6900
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1200
      Left            =   6360
      Top             =   3900
   End
   Begin MSComctlLib.ListView lvProc 
      Height          =   4245
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   7488
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
         Text            =   "Start"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "End"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "pid"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Path"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Menu mnuProcess 
      Caption         =   "Process"
      Visible         =   0   'False
      Begin VB.Menu mnuCopyList 
         Caption         =   "Copy List"
      End
      Begin VB.Menu mnuClearPidList 
         Caption         =   "Clear"
      End
      Begin VB.Menu mnuSpacer 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSuspend 
         Caption         =   "Suspend"
      End
      Begin VB.Menu mnuResume 
         Caption         =   "Resume"
      End
      Begin VB.Menu mnuTerminate 
         Caption         =   "Terminate"
      End
      Begin VB.Menu mnuKillAll 
         Caption         =   "Kill All"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function NtSuspendProcess Lib "ntdll.dll" (ByVal hProc As Long) As Long
Private Declare Function NtResumeProcess Lib "ntdll.dll" (ByVal hProc As Long) As Long


Public dlg As New clsCmnDlg
Public fso As New CFileSystem2
Dim cpi As New CProcessInfo
Dim procs As Collection 'of CProcess
Dim selli As ListItem

Private Sub Form_Load()
    Set procs = cpi.GetRunningProcesses()
    Timer1.Enabled = True
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    With lvProc
        .Width = Me.Width - 500
        .Height = Me.Height - 500
        .ColumnHeaders(.ColumnHeaders.Count).Width = .Width - .ColumnHeaders(.ColumnHeaders.Count).Left - 200
    End With
End Sub

Private Sub lvProc_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set selli = Item
End Sub

Private Sub lvProc_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu mnuProcess
End Sub

Private Sub mnuClearPidList_Click()
    lvProc.ListItems.Clear
End Sub

Private Sub mnuCopyList_Click()
    Screen.MousePointer = vbHourglass
    Clipboard.Clear
    Clipboard.SetText GetAllElements(lvProc)
    Screen.MousePointer = vbNormal
End Sub

Private Sub mnuKillAll_Click()
    Dim p As CProcess
    Dim li As ListItem
    On Error Resume Next
    For Each li In lvProc.ListItems
        Set p = li.Tag
        cpi.TerminateProces p.pid
    Next
End Sub

Private Sub mnuResume_Click()
    Dim p As CProcess
    On Error Resume Next
    If selli Is Nothing Then Exit Sub
    Set p = selli.Tag
    NtResumeProcess p.pid
End Sub

Private Sub mnuSuspend_Click()
    Dim p As CProcess
    On Error Resume Next
    If selli Is Nothing Then Exit Sub
    Set p = selli.Tag
    NtSuspendProcess p.pid
End Sub

Private Sub mnuTerminate_Click()
    Dim p As CProcess
    On Error Resume Next
    If selli Is Nothing Then Exit Sub
    Set p = selli.Tag
    cpi.TerminateProces p.pid
End Sub

Private Sub Timer1_Timer()

    Dim c As Collection
    Dim p As CProcess
    Dim li As ListItem
    
    Set c = cpi.GetRunningProcesses()
    For Each p In c
        If Not ProcessExistedAtStart(p) Then
            AddPid p
        End If
    Next
    
    For Each li In lvProc.ListItems
        Set p = li.Tag
        If li.ForeColor <> vbRed Then
            If Not ProcessStillExists(p, c) Then
                li.SubItems(1) = Time
                SetLiColor li, vbRed
                RemoveProcess p
                li.Selected = False
            End If
        End If
    Next
    
    'log when a base process quit? naghh not right now to confusing for output..
    
End Sub

Sub SetLiColor(li As ListItem, newcolor As Long)
    Dim f As ListSubItem
'    On Error Resume Next
    li.ForeColor = newcolor
    For Each f In li.ListSubItems
        f.ForeColor = newcolor
    Next
End Sub

Function RemoveProcess(p As CProcess)
    
    Dim p2 As CProcess
    Dim i As Long
    For i = 1 To procs.Count
        Set p2 = procs(i)
        If p2.pid = p.pid And p2.path = p.path Then
            procs.Remove i
            Exit Function
        End If
    Next
    
End Function

Function ProcessStillExists(p As CProcess, c As Collection) As Boolean
    
    Dim p2 As CProcess
    For Each p2 In c
        If p2.pid = p.pid And p2.path = p.path Then
            ProcessStillExists = True
            Exit Function
        End If
    Next
    
End Function

Function ProcessExistedAtStart(p As CProcess) As Boolean
    
    Dim p2 As CProcess
    For Each p2 In procs
        If p2.pid = p.pid And p2.path = p.path Then
            ProcessExistedAtStart = True
            Exit Function
        End If
    Next
    
End Function

Function AddPid(p As CProcess)
    Dim li As ListItem
    Dim pid As Long
    On Error Resume Next
   
    Set li = lvProc.ListItems.Add(, , Time)
    li.SubItems(2) = Hex(p.pid)
    li.SubItems(3) = cpi.GetProcessPath(p.pid)
    Set li.Tag = p
    procs.Add p
    
End Function


'Private Sub cmdSave_Click()
'    On Error Resume Next
'    Dim i As Long, t, f As String
'    Dim li As ListItem
'
'    f = dlg.SaveDialog(textFiles, , , , Me.hwnd)
'    If Len(f) = 0 Then Exit Sub
'
'    For Each li In lv.ListItems
'        t = t & "pid: " & li.Text & " - " & li.SubItems(1) & vbCrLf
'    Next
'
'    fso.WriteFile f, t
'
'End Sub

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

Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    Dim x As Integer
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init:     ReDim ary(0): ary(0) = value
End Sub

Function GetAllElements(lv As ListView) As String
    Dim ret() As String, i As Integer, tmp As String
    Dim li As ListItem

    For i = 1 To lv.ColumnHeaders.Count
        tmp = tmp & lv.ColumnHeaders(i).Text & vbTab
    Next

    push ret, tmp
    push ret, String(50, "-")

    For Each li In lv.ListItems
        tmp = li.Text & vbTab
        For i = 1 To lv.ColumnHeaders.Count - 1
            tmp = tmp & li.SubItems(i) & vbTab
        Next
        push ret, tmp
    Next

    GetAllElements = Join(ret, vbCrLf)

End Function

VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInjectionScan 
   Caption         =   "32bit process Injection Scan"
   ClientHeight    =   3720
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11970
   LinkTopic       =   "Form1"
   ScaleHeight     =   3720
   ScaleWidth      =   11970
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   60
      TabIndex        =   2
      Top             =   3060
      Width           =   11835
      Begin VB.TextBox Text1 
         Height          =   345
         Left            =   750
         TabIndex        =   8
         Top             =   30
         Width           =   4995
      End
      Begin VB.CommandButton cmdAbort 
         Caption         =   "Abort"
         Height          =   405
         Left            =   9720
         TabIndex        =   7
         Top             =   0
         Width           =   1005
      End
      Begin VB.CommandButton cmdRescan 
         Caption         =   "Rescan"
         Height          =   405
         Left            =   10800
         TabIndex        =   6
         Top             =   0
         Width           =   1035
      End
      Begin VB.CommandButton cmdNextProc 
         Caption         =   "Next Proc"
         Height          =   405
         Left            =   8430
         TabIndex        =   5
         Top             =   0
         Width           =   1185
      End
      Begin VB.TextBox txtMinEntropy 
         Height          =   345
         Left            =   7650
         TabIndex        =   4
         Text            =   "50"
         Top             =   30
         Width           =   465
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Remove if entropy <"
         Height          =   405
         Left            =   5820
         TabIndex        =   3
         Top             =   30
         Width           =   1785
      End
      Begin VB.Label Label1 
         Caption         =   "Process:"
         Height          =   255
         Left            =   0
         TabIndex        =   9
         Top             =   60
         Width           =   705
      End
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   255
      Left            =   30
      TabIndex        =   0
      Top             =   90
      Width           =   11925
      _ExtentX        =   21034
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.ListView lv 
      Height          =   2595
      Left            =   0
      TabIndex        =   1
      Top             =   420
      Width           =   11925
      _ExtentX        =   21034
      _ExtentY        =   4577
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "pid"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Base"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Size"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Protect"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Type"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Module"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Entropy"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Begin VB.Menu mnuView 
         Caption         =   "View"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuSaveAll 
         Caption         =   "Save All"
      End
      Begin VB.Menu mnuSearchMem 
         Caption         =   "Search"
      End
      Begin VB.Menu mnuStrings 
         Caption         =   "Strings"
      End
      Begin VB.Menu mnuViewMemoryMap 
         Caption         =   "Memory Map"
      End
   End
End
Attribute VB_Name = "frmInjectionScan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim abort As Boolean
Dim pi As New CProcessInfo
Dim selli As ListItem
Dim nextProc As Boolean
Dim totalScanned As Long
Dim totalRWEFound As Long
Dim multiscanMode As Boolean

'todo: user config list of common target processes and only scan selected processes to speed up?

Private Sub cmdAbort_Click()
    abort = True
End Sub

Function StealthInjectionScan()
        
    Dim cp As CProcess
    Dim c As Collection
    
    On Error Resume Next
    
    Me.Visible = True
    Set c = pi.GetRunningProcesses()
    pb.max = c.count
    pb.value = 1
    abort = False
    totalScanned = 0
    totalRWEFound = 0
    multiscanMode = True
    
    For Each cp In c
        Me.Caption = "Scanning " & pb.value & "/" & c.count & "  Found: " & lv.ListItems.count & " Processing: " & cp.path & " TotalRWEFound: " & totalRWEFound & " Total Allocs Scanned: " & totalScanned
       
        'If diff.CProc.x64.IsProcess_x64(cp.pid) = r_32bit Then
            FindStealthInjections cp.pid, pi.GetProcessPath(cp.pid)
        'End If
        
        DoEvents
        Sleep 20
        pb.value = pb.value + 1
        If abort Then Exit For
    Next
    
    multiscanMode = False
    pb.value = 0
    Me.Caption = "Found " & lv.ListItems.count & " allocations"
    
        
End Function


Sub FindStealthInjections(pid As Long, pName As String)
    
    Dim c As Collection
    Dim cMem As CMemory
    Dim li As ListItem
    Dim modules As Long
    Dim execSections As Long
    Dim mm As matchModes
    Dim knownModules As Long
    Dim s As String
    Dim entropy As Long
    Dim minEntropy As Long
    
    On Error Resume Next
    Me.Visible = True
    minEntropy = CLng(txtMinEntropy)
    
    'If diff.CProc.x64.IsProcess_x64(pid) <> r_32bit Then
    '    MsgBox x64Error, vbInformation
    '    Exit Sub
    'End If
    
    If Err.Number <> 0 Then
        minEntropy = 50
        txtMinEntropy = 50
        Err.Clear
    End If
    
    abort = False
    nextProc = False
    Set c = pi.GetMemoryMap(pid)

    If multiscanMode = False Then
        pb.max = c.count
        pb.value = 0
    End If
    
    'todo: replace(chr(0) in readmem, if it shrinks by % then its just junk?
    For Each cMem In c
        If abort Then Exit Sub
        If nextProc Then Exit Sub
        totalScanned = totalScanned + 1
        
        If multiscanMode = False Then
            pb.value = pb.value + 1
            Me.Caption = "Scanning " & pb.value & "/" & c.count & "  Found: " & lv.ListItems.count & " Total Allocs Scanned: " & totalScanned
        End If
         
        If cMem.Protection = PAGE_EXECUTE_READWRITE And cMem.MemType <> MEM_IMAGE Then
            
            totalRWEFound = totalRWEFound + 1
            entropy = -1
            
            If cMem.size < &H10000 Then 'some level of DoS protection against huge allocations (Dridex)...
                s = pi.ReadMemory2(cMem.pid, cMem.Base, cMem.size) 'doesnt add that much time
                entropy = CalculateEntropy(s)
                s = Empty
            
                'If chkMinEntropy.Value = 1 Then
                '    If entropy < minEntropy Then GoTo nextOne
                'End If
            End If
            
            Set li = lv.ListItems.Add(, , pid)
            li.SubItems(1) = cMem.BaseAsHexString
            li.SubItems(2) = Hex(cMem.size)
            li.SubItems(3) = cMem.MemTypeAsString()
            li.SubItems(4) = cMem.ProtectionAsString()
            li.SubItems(5) = pName
            
            If VBA.Left(pi.ReadMemory2(cMem.pid, cMem.Base, 2), 2) = "MZ" Then
                SetLiColor li, vbRed
            End If

            Set li.Tag = cMem
            li.SubItems(6) = entropy
        End If
        
nextone:
        DoEvents
        Sleep 5
    Next
    
    If multiscanMode = False Then pb.value = 0
    
End Sub




Private Sub cmdNextProc_Click()
    nextProc = True
End Sub

Private Sub cmdRescan_Click()
    lv.ListItems.Clear
    StealthInjectionScan
End Sub

Private Sub Form_Load()

     mnuPopup.Visible = False
     Me.Icon = frmMain.Icon
     lv.ColumnHeaders(6).Width = lv.Width - lv.ColumnHeaders(6).Left - 350 - lv.ColumnHeaders(7).Width
     
     If isIde() Then
        LoadLibrary "zlib.dll"
     End If
     
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Frame1.top = Me.Height - 400 - Frame1.Height - (TitleBarHeight(Me) - 255)
    lv.Height = Frame1.top - 600
    lv.Width = Me.Width - lv.Left - 200
    pb.Width = lv.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
    abort = True
End Sub

Private Sub lv_DblClick()
    mnuView_Click
End Sub

Private Sub mnuSave_Click()
    If selli Is Nothing Then Exit Sub
    Dim f As String
    Dim pid As Long
    On Error Resume Next
    pid = CLng(selli.Text)
    'f = InputBox("Save file as: ", , UserDeskTopFolder & "\" & pid & "_" & selli.SubItems(1) & ".mem")
    f = frmDlg.SaveDialog(AllFiles, UserDeskTopFolder, "Save As:", , Me, pid & "_" & selli.SubItems(1) & ".mem")
    If Len(f) = 0 Then Exit Sub
    If pi.DumpMemory(pid, selli.SubItems(1), selli.SubItems(2), f) Then
        MsgBox "File successfully saved"
    Else
        MsgBox "Error saving file: " & Err.Description
    End If
End Sub

Private Sub mnuSaveAll_Click()

    If lv.ListItems.count = 0 Then Exit Sub
    
    Dim f As String
    Dim pid As Long
    On Error Resume Next
    Dim li As ListItem
    Dim e As String
    Dim i As Long
    Dim saveDir As String
    
    saveDir = UserDeskTopFolder
    
    For Each li In lv.ListItems
        pid = CLng(li.Text)
        f = saveDir & "\" & pid & "_" & li.SubItems(1) & ".mem"
        If pi.DumpMemory(pid, li.SubItems(1), li.SubItems(2), f) Then
            i = i + 1
        Else
            e = e & "Error dumping " & f & vbCrLf
        End If
    Next
    
    If Len(e) = 0 Then
          MsgBox "All Files successfully saved to " & vbCrLf & vbCrLf & saveDir
    Else
          MsgBox i & " of " & lv.ListItems.count & " items saved " & vbCrLf & vbCrLf & e
    End If
        
End Sub

Private Sub mnuSearchMem_Click()
    Dim li As ListItem
    
    Dim s As String
    Dim s2 As String
    Dim ret As String
    Dim a As Long
    Dim b As Long
    Dim cMem As CMemory
    Dim m As String
    
    If lv.ListItems.count = 0 Then
        MsgBox "Nothing to search"
        Exit Sub
    End If
    
    s = InputBox("Enter string to search for:")
    If Len(s) = 0 Then Exit Sub
    
    s2 = StrConv(s, vbUnicode, LANG_US)
    pb.max = lv.ListItems.count
    pb.value = 0
    abort = False
    
    For Each li In lv.ListItems
        If abort = True Then Exit For
        li.Selected = True
        li.EnsureVisible
        Set cMem = li.Tag
        DoEvents
        lv.Refresh
        m = pi.ReadMemory2(cMem.pid, cMem.Base, cMem.size)
        a = InStr(1, m, s, vbTextCompare)
        b = InStr(1, m, s2, vbTextCompare)
        If a > 0 Then ret = ret & "pid: " & li.Text & " base: " & li.SubItems(1) & " offset: " & cMem.Base & "+" & a & " ASCII " & li.SubItems(5) & vbCrLf
        If b > 0 Then ret = ret & "pid: " & li.Text & " base: " & li.SubItems(1) & " offset: " & cMem.Base & "+" & b & " UNICODE " & li.SubItems(5) & vbCrLf
        pb.value = pb.value + 1
    Next
            
    pb.value = 0
    
    If Len(ret) > 0 Then
        frmReport.ShowList ret
    Else
        MsgBox "Specified string not found (ASCII or UNICODE)", vbInformation
    End If
    
End Sub

Private Sub mnuStrings_Click()
    If selli Is Nothing Then Exit Sub
    Dim f As String
    Dim pid As Long
    On Error Resume Next
    pid = CLng(selli.Text)
    f = fso.GetFreeFileName(Environ("temp"), ".bin")
    If pi.DumpMemory(pid, selli.SubItems(1), selli.SubItems(2), f) Then
        LaunchStrings f, True
    Else
        MsgBox "Error saving file: " & Err.Description
    End If
End Sub

Private Sub mnuView_Click()
    If selli Is Nothing Then Exit Sub
    Dim s As String
    Dim pid As Long
    Dim Base 'As Long
    On Error Resume Next
    Base = selli.SubItems(1)
    pid = CLng(selli.Text)
    s = pi.ReadMemory2(pid, Base, CLng("&h" & selli.SubItems(2)))
    If Len(s) = 0 Then
        MsgBox "Failed to readmemory?"
        Exit Sub
    End If
    Dim f As New rhexed.CHexEditor
    f.Editor.AdjustBaseOffset = Base
    f.Editor.LoadString s
End Sub


Private Sub lv_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Text1 = Item.SubItems(5)
    Set selli = Item
End Sub

Private Sub lv_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu mnuPopup
End Sub

Private Sub mnuViewMemoryMap_Click()
    If selli Is Nothing Then Exit Sub
    Dim pid As Long
    On Error Resume Next
    pid = CLng(selli.Text)
    If pid <> 0 Then
        frmMemoryMap.ShowMemoryMap pid
    End If
End Sub

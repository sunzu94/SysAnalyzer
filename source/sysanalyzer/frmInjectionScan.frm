VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmInjectionScan 
   Caption         =   "Injection Scan"
   ClientHeight    =   3585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11970
   LinkTopic       =   "Form1"
   ScaleHeight     =   3585
   ScaleWidth      =   11970
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAbort 
      Caption         =   "Abort"
      Height          =   405
      Left            =   10920
      TabIndex        =   4
      Top             =   3030
      Width           =   1005
   End
   Begin VB.TextBox Text1 
      Height          =   345
      Left            =   840
      TabIndex        =   3
      Top             =   3060
      Width           =   9975
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
      NumItems        =   6
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
   End
   Begin VB.Label Label1 
      Caption         =   "Process:"
      Height          =   255
      Left            =   90
      TabIndex        =   2
      Top             =   3090
      Width           =   705
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuView 
         Caption         =   "View"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
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
    
    For Each cp In c
        Me.Caption = "Scanning " & pb.value & "/" & c.count & "..."
        FindStealthInjections cp.pid, pi.GetProcessPath(cp.pid)
        DoEvents
        pb.value = pb.value + 1
        If abort Then Exit For
    Next
    
    pb.value = 0
    
    
        
End Function


Private Sub FindStealthInjections(pid As Long, pName As String)
    
    Dim c As Collection
    Dim cMem As CMemory
    Dim li As ListItem
    Dim modules As Long
    Dim execSections As Long
    Dim mm As matchModes
    Dim knownModules As Long
    
    On Error Resume Next
    Set c = pi.GetMemoryMap(pid)

    'todo: replace(chr(0) in readmem, if it shrinks by % then its just junk?
    For Each cMem In c
        If abort Then Exit Sub
        If cMem.Protection = PAGE_EXECUTE_READWRITE And cMem.MemType <> MEM_IMAGE Then
            Set li = lv.ListItems.Add(, , pid)
            li.SubItems(1) = Hex(cMem.base)
            li.SubItems(2) = Hex(cMem.Size)
            li.SubItems(3) = cMem.MemTypeAsString()
            li.SubItems(4) = cMem.ProtectionAsString()
            li.SubItems(5) = pName
            If VBA.Left(pi.ReadMemory(cMem.pid, cMem.base, 2), 2) = "MZ" Then
                SetLiColor li, vbRed
            End If
        End If
    Next
    
End Sub

Private Sub Form_Load()
     lv.ColumnHeaders(6).Width = lv.Width - lv.ColumnHeaders(6).Left - 350
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
    f = InputBox("Save file as: ", , UserDeskTopFolder & "\" & pid & "_" & selli.SubItems(1) & ".mem")
    If Len(f) = 0 Then Exit Sub
    If pi.DumpProcessMemory(pid, CLng("&h" & selli.SubItems(1)), CLng("&h" & selli.SubItems(2)), f) Then
        MsgBox "File successfully saved"
    Else
        MsgBox "Error saving file: " & Err.Description
    End If
End Sub

Private Sub mnuView_Click()
    If selli Is Nothing Then Exit Sub
    Dim s As String
    Dim pid As Long
    On Error Resume Next
    pid = CLng(selli.Text)
    s = pi.ReadMemory(pid, CLng("&h" & selli.SubItems(1)), CLng("&h" & selli.SubItems(2)))
    If Len(s) = 0 Then
        MsgBox "Failed to readmemory?"
        Exit Sub
    End If
    frmReport.ShowList HexDump(s), False, selli.SubItems(1) & ".mem", False
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

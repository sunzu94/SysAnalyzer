VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMemoryMap 
   Caption         =   "Memory Map"
   ClientHeight    =   6165
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11655
   LinkTopic       =   "Form1"
   ScaleHeight     =   6165
   ScaleWidth      =   11655
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   1035
      Left            =   90
      TabIndex        =   1
      Top             =   5070
      Width           =   11475
   End
   Begin MSComctlLib.ListView lv 
      Height          =   4935
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   11505
      _ExtentX        =   20294
      _ExtentY        =   8705
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Base"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Size"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Protect"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Type"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Module"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuViewMemory 
         Caption         =   "View"
      End
      Begin VB.Menu mnuSaveMemory 
         Caption         =   "Save"
      End
   End
End
Attribute VB_Name = "frmMemoryMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pi As New CProcessInfo
Dim active_pid As Long
Dim selli As ListItem

Public Sub ShowMemoryMap(pid As Long)
    
    Dim c As Collection
    Dim cMem As CMemory
    Dim li As ListItem
    Dim modules As Long
    Dim execSections As Long
    Dim mm As matchModes
    Dim knownModules As Long
    
    On Error Resume Next
    
    active_pid = pid
    
    Me.Show
    List1.AddItem "Loading memory map for pid: " & pid
    
    Set c = pi.GetMemoryMap(pid)
    
    If c.count = 0 Then
        List1.AddItem "Failed to load memory map!"
        Exit Sub
    End If
    
    lv.ListItems.Clear
    For Each cMem In c
        Set li = lv.ListItems.Add(, , Hex(cMem.base))
        li.SubItems(1) = Hex(cMem.Size)
        li.SubItems(2) = cMem.MemTypeAsString()
        li.SubItems(3) = cMem.ProtectionAsString()
        li.SubItems(4) = cMem.ModuleName
        
        If known.Loaded And known.Ready Then
            mm = known.isFileKnown(cMem.ModuleName)
            li.ListSubItems(4).ForeColor = IIf(mm = exact_match, vbGreen, vbRed)
            knownModules = knownModules + 1
        End If
            
        If Len(cMem.ModuleName) > 0 Then modules = modules + 1
        
        If cMem.Protection = PAGE_EXECUTE_READWRITE Or _
            cMem.Protection = PAGE_EXECUTE_READ Or _
            cMem.Protection = PAGE_EXECUTE_WRITECOPY _
        Then
            If cMem.Protection = PAGE_EXECUTE_READWRITE And cMem.MemType <> MEM_IMAGE Then
                SetLiColor li, vbRed
                If VBA.Left(pi.ReadMemory(cMem.pid, cMem.base, 2), 2) = "MZ" Then
                    List1.AddItem Hex(cMem.base) & " is RWE but not part of an image (CONFIRMED INJECTION)"
                Else
                    List1.AddItem Hex(cMem.base) & " is RWE but not part of an image..possible injection"
                End If
            Else
                SetLiColor li, vbBlue
                execSections = execSections + 1
                'List1.AddItem Hex(cMem.Base) & " is " & cMem.ProtectionAsString(true)
            End If
        End If
    Next
    
    List1.AddItem "Found " & modules & " modules and " & execSections & " executable sections"
    
    If known.Loaded And known.Ready Then List1.AddItem knownModules & " known modules found"
    
    
End Sub

Private Sub Form_Load()
    lv.ColumnHeaders(5).Width = lv.Width - lv.ColumnHeaders(5).Left - 350
End Sub

Private Sub lv_DblClick()
    mnuViewMemory_Click
End Sub

Private Sub lv_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set selli = Item
End Sub

Private Sub lv_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu mnuPopup
End Sub

Private Sub mnuSaveMemory_Click()
    If selli Is Nothing Then Exit Sub
    Dim f As String
    On Error Resume Next
    f = InputBox("Save file as: ", , UserDeskTopFolder & "\" & selli.Text & ".mem")
    If Len(f) = 0 Then Exit Sub
    If pi.DumpProcessMemory(active_pid, CLng("&h" & selli.Text), CLng("&h" & selli.SubItems(1)), f) Then
        MsgBox "File successfully saved"
    Else
        MsgBox "Error saving file: " & Err.Description
    End If
End Sub

Private Sub mnuViewMemory_Click()
    If selli Is Nothing Then Exit Sub
    Dim s As String
    On Error Resume Next
    s = pi.ReadMemory(active_pid, CLng("&h" & selli.Text), CLng("&h" & selli.SubItems(1)))
    If Len(s) = 0 Then
        List1.AddItem "Failed to readmemory?"
        Exit Sub
    End If
    frmReport.ShowList HexDump(s), False, selli.Text & ".mem", False
End Sub

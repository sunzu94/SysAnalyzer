VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   Caption         =   "SysAnalyzer"
   ClientHeight    =   6090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10230
   Icon            =   "Form1.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "frmMain"
   ScaleHeight     =   6090
   ScaleWidth      =   10230
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin MSComctlLib.ListView lv 
      Height          =   5295
      Left            =   60
      TabIndex        =   8
      Top             =   750
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   9340
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      OLEDropMode     =   1
      Checkboxes      =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      OLEDropMode     =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Watch Dirs"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   1950
      TabIndex        =   7
      Top             =   360
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   315
      Left            =   3690
      TabIndex        =   6
      Top             =   360
      Width           =   1155
   End
   Begin VB.CommandButton cmdCopyList 
      Caption         =   "Copy List"
      Height          =   315
      Left            =   4980
      TabIndex        =   5
      Top             =   360
      Width           =   1335
   End
   Begin VB.CommandButton cmdSaveDirWatchFile 
      Caption         =   "Save Selected file"
      Height          =   315
      Left            =   6420
      TabIndex        =   2
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox txtIgnore 
      Height          =   315
      Left            =   660
      TabIndex        =   1
      Top             =   0
      Width           =   9495
   End
   Begin VB.CommandButton cmdDirWatch 
      Caption         =   "Start Monitor"
      Height          =   315
      Left            =   8940
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
   Begin MSComctlLib.ListView lvDirWatch 
      Height          =   5355
      Left            =   1800
      TabIndex        =   3
      Top             =   720
      Width           =   8325
      _ExtentX        =   14684
      _ExtentY        =   9446
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Action"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "File"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "(Add folders with right click or drag/drop)"
      Height          =   315
      Left            =   210
      TabIndex        =   10
      Top             =   540
      Width           =   3465
   End
   Begin VB.Label Label1 
      Caption         =   "Check Folders to watch and start monitor"
      Height          =   225
      Left            =   90
      TabIndex        =   9
      Top             =   330
      Width           =   3555
   End
   Begin VB.Label Label3 
      Caption         =   "Ignore"
      Height          =   315
      Index           =   0
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   555
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuAddFolder 
         Caption         =   "Add Folder"
      End
      Begin VB.Menu mnuCheckAll 
         Caption         =   "Check All"
      End
      Begin VB.Menu mnuClearAll 
         Caption         =   "Clear All"
      End
   End
End
Attribute VB_Name = "frmMain"
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

Dim WithEvents subclass As CSubclass2
Attribute subclass.VB_VarHelpID = -1

Dim liDirWatch As ListItem
Dim dlg As New clsCmnDlg2
Dim fso As New CFileSystem2

Sub Initalize()
    
    Set subclass = New CSubclass2
   
    subclass.AttachMessage frmDirWatch.hwnd, WM_COPYDATA
    
    lvDirWatch.ColumnHeaders(2).Width = lvDirWatch.Width - 100 - lvDirWatch.ColumnHeaders(1).Width
    
    txtIgnore = GetSetting(App.EXEName, "Settings", "txtIgnore", "\config\software , modified:, ")

    
End Sub

Private Sub cmdDelLike_Click()

End Sub

Private Sub cmdClear_Click()
    lvDirWatch.ListItems.Clear
End Sub

Private Sub cmdCopyList_Click()
    Clipboard.Clear
    Clipboard.SetText GetAllElements(lvDirWatch)
    MsgBox "Copy complete", vbInformation
End Sub

Private Sub cmdDirWatch_Click()
    
    Dim li As ListItem
    
    With cmdDirWatch
        If Len(.Tag) > 0 Then
            .Tag = ""
            lv.Enabled = True
            DirWatchCtl False
            .Caption = "Start monitor"
        Else
            Set watchDirs = New Collection
            For Each li In lv.ListItems
                If li.Checked = True Then
                    watchDirs.Add li.Tag
                End If
            Next
            .Tag = "xx"
            lv.Enabled = False
            DirWatchCtl True
            .Caption = "Stop monitor"
        End If
    End With
    
End Sub



Private Sub Form_Load()
  
    Dim i As Long
    Dim li As ListItem
    On Error Resume Next
    Dim tmp
    
    Me.Visible = True
    Initalize
    
    For i = 0 To Drive1.ListCount - 1
        tmp = Split(Drive1.List(i), ":")
        Set li = lv.ListItems.Add(, , Drive1.List(i))
        li.Tag = tmp(0) & ":\"
    Next
    
    lv.ListItems(1).Checked = True
    
    Set cApiData = New Collection
    Set cLogData = New Collection

    'DirWatchCtl True

    
End Sub















Private Sub cmdSaveDirWatchFile_Click()
    If liDirWatch Is Nothing Then Exit Sub
    
    On Error Resume Next
    Dim f As String, d As String
    
    f = liDirWatch.SubItems(1)
    
    If Not fso.FileExists(f) Then
        MsgBox "File not found: " & f
    Else
        ' f, UserDeskTopFolder & "\"
        d = dlg.SaveDialog(AllFiles, , , , , fso.FileNameFromPath(f))
        FileCopy f, d
    End If
    
    If Err.Number <> 0 Then
        MsgBox "Error: " & Err.Description
    Else
        MsgBox FileLen(f) & " bytes saved successfully!", vbInformation
    End If
    
End Sub


 



Private Sub Form_Resize()
    On Error Resume Next
    lvDirWatch.Width = Me.Width - lvDirWatch.Left - 200
    With lvDirWatch
        .Height = Me.Height - .Top - 500
        lv.Height = .Height
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
     
    DirWatchCtl False
    subclass.DetatchMessage frmDirWatch.hwnd, WM_COPYDATA
    Unload frmDirWatch
    
End Sub
 
 
Private Sub lv_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu mnuPopup
End Sub

Private Sub lv_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    Dim f As String, li As ListItem
    
    f = Data.Files(1)
    If fso.FolderExists(f) Then
        For Each li In lv.ListItems
            If li.Tag = f Then
                MsgBox "This folder is already listed..", vbInformation
                Exit Sub
            End If
        Next
        
        Set li = lv.ListItems.Add(, , f)
        li.Tag = f
        li.Checked = True
    Else
        MsgBox "You can only drop folders on here to add them", vbInformation
    End If
    
End Sub

Private Sub lvDirWatch_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set liDirWatch = Item
End Sub


Private Sub mnuAddFolder_Click()
    Dim f As String
    Dim li As ListItem
    
    f = dlg.FolderDialog(, Me.hwnd)
    If Len(f) = 0 Then Exit Sub
    
    For Each li In lv.ListItems
        If li.Tag = f Then
            MsgBox "This folder is already listed..", vbInformation
            Exit Sub
        End If
    Next
    
    Set li = lv.ListItems.Add(, , f)
    li.Tag = f
    li.Checked = True
    
End Sub

Private Sub mnuCheckAll_Click()
    Dim li As ListItem
    For Each li In lv.ListItems
        li.Checked = True
    Next
End Sub

Private Sub mnuClearAll_Click()
    Dim li As ListItem
    For Each li In lv.ListItems
        li.Checked = False
    Next
End Sub

Private Sub subclass_MessageReceived(hwnd As Long, wMsg As Long, wParam As Long, lParam As Long, Cancel As Boolean)
    Dim msg As String
    Dim li As ListItem
    Dim tmp
    
    If wMsg = WM_COPYDATA Then
        If RecieveTextMessage(lParam, msg) Then
                
                If InStr(msg, "NTUSER.DAT") > 0 Then Exit Sub
                If InStr(msg, "\Prefetch\") > 0 Then Exit Sub
                If Right(msg, 4) = ".lnk" Then Exit Sub
                
                If AnyOfTheseInstr(msg, txtIgnore) Then Exit Sub
                If KeyExistsInCollection(cLogData, msg) Then Exit Sub
                On Error Resume Next
                cLogData.Add msg, msg
                If InStr(msg, ":") > 0 And VBA.Left(msg, 8) <> "Watching" Then
                    tmp = Split(msg, ":", 2)
                    tmp(0) = VBA.Left(tmp(0), 3) & " - " & Format(Now, "h:m:s")
                    Set li = lvDirWatch.ListItems.Add(, , tmp(0))
                    li.SubItems(1) = Replace(Replace(Trim(tmp(1)), "\\", "\"), Chr(0), Empty)
                Else
                    Set li = lvDirWatch.ListItems.Add(, , Format(Now, "h:m:s"))
                    li.SubItems(1) = Replace(Replace(Trim(msg), "\\", "\"), Chr(0), Empty)
                End If
                
        End If
    End If
    
End Sub


 


Private Function RecieveTextMessage(lParam As Long, msg As String) As Boolean
   
    Dim CopyData As COPYDATASTRUCT
    Dim Buffer(1 To 2048) As Byte
    Dim Temp As String
    
    msg = Empty
    
    CopyMemory CopyData, ByVal lParam, Len(CopyData)
    
    If CopyData.dwFlag = 3 Then
        CopyMemory Buffer(1), ByVal CopyData.lpData, CopyData.cbSize
        Temp = StrConv(Buffer, vbUnicode)
        Temp = Left$(Temp, InStr(1, Temp, Chr$(0)) - 1)
        'heres where we work with the intercepted message
        msg = Temp
        RecieveTextMessage = True
    End If
    
End Function
 








VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CKnownFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim fso2 As New CFileSystem2

Public Ready As Boolean
Public Loaded As Boolean
Public abort As Boolean
Public HideKnownInDisplays As Boolean

Event PercentDone(i As Long)
Event CurrentDirectory(folder As String, max As Long, filt As String)

Enum matchModes
    exact_match = 0
    not_found = 1
    hash_changed = 2
End Enum
    
Function TotalRecords() As Long
    On Error Resume Next
    ado.CloseConnection
    TotalRecords = CLng(ado.ReturnRecordSet("Select count(autoid) as cnt from tblFiles")!cnt)
    If TotalRecords > 0 Then Loaded = True
End Function

Function ClearDB()
    ado.Execute "Delete * from tblFiles"
End Function

Function BuildDB()
    Dim tmp() As String
    Dim f
    Dim win As String
    Dim max
    Dim i
    Dim pcent
    
    abort = False
    win = Environ("WinDir")
    If Not fso.FolderExists(win) Then
        MsgBox "Windir env variable set sorry"
        Exit Function
    End If
    
    tmp = fso.GetFolderFiles(win, "*.dll")
    addDir tmp, win, "*.dll"
    If abort Then Exit Function
    
    tmp = fso.GetFolderFiles(win, "*.exe")
    addDir tmp, win, "*.exe"
    If abort Then Exit Function
    
    win = win & "System32"
    tmp = fso.GetFolderFiles(win, "*.dll")
    addDir tmp, win, "*.dll"
    If abort Then Exit Function

    tmp = fso.GetFolderFiles(win, "*.exe")
    addDir tmp, win, "*.exe"
    If abort Then Exit Function

    tmp = fso.GetFolderFiles(win, "*.acm")
    addDir tmp, win, "*.acm"
    If abort Then Exit Function

    tmp = fso.GetFolderFiles(win, "*.drv")
    addDir tmp, win, "*.drv"
    If abort Then Exit Function

    tmp = fso.GetFolderFiles(win, "*.ime")
    addDir tmp, win, "*.ime"
    If abort Then Exit Function

    win = win & "\drivers"
    tmp = fso.GetFolderFiles(win)
    addDir tmp, win, "*.*"
    If abort Then Exit Function

    win = "c:\program files"
    tmp = fso2.GetFolderFiles(win, "*.exe", , True)
    addDir tmp, win, "*.exe"
    If abort Then Exit Function

    tmp = fso2.GetFolderFiles(win, "*.dll", , True)
    addDir tmp, win, "*.dll"
    
End Function

Private Function addDir(ary, d, filt As String)
    Dim f, i As Long, pcent As Long, max As Long
    On Error Resume Next
    
    i = 0
    max = UBound(ary)
    
    RaiseEvent CurrentDirectory(CStr(d), max, filt)
    
    For Each f In ary
        AddKnown CStr(f)
        If i Mod 10 = 0 Then
            'pcent = (i \ max) * 100
            RaiseEvent PercentDone(i)
        End If
        DoEvents
        If abort Then Exit Function
        DoEvents
        i = i + 1
    Next

End Function

Function recordExists(fpath As String, fhash As String) As Boolean
    
    Dim rs As Recordset
    
    On Error GoTo hell
    
    Set rs = ado.ReturnRecordSet("Select * from tblFiles where path='" & fpath & "' and hash='" & fhash & "'")
     
    If rs.EOF And rs.BOF Then GoTo hell
    
    recordExists = True
    
hell:

End Function


Function isFileKnown(fp) As matchModes
    
    Dim rs As Recordset
    Dim fhash As String
    Dim fpath As String
    
    On Error GoTo hell
    
    fpath = CStr(fp)
    
    If ado.ActiveCon.State = 0 Then ado.ActiveCon.Open
    
    Set rs = ado.ActiveCon.Execute("Select * from tblFiles where path like '" & fpath & "'")
    
    
    If rs.EOF And rs.BOF Then GoTo hell
    
    fhash = hash.HashFile(fpath, MD5, HexFormat)
    
    If rs("hash") <> fhash Then
        isFileKnown = hash_changed
        Exit Function
    End If
    
    isFileKnown = exact_match
    Exit Function
    
hell:
    isFileKnown = not_found
    
End Function

Function AddKnown(fpath As String)

    'tblFIles: autoid;name;path;version;trusted,hash
    On Error Resume Next
    
    If Not fso.FileExists(fpath) Then Exit Function
        
    Dim name, version, fhash
    
    'If InStr(fpath, "modem") > 0 Then Stop
    
    fpath = Replace(fpath, "\\", "\")
    name = fso.FileNameFromPath(fpath)
    version = FileInfo(fpath).FileVersion
    fhash = hash.HashFile(fpath, MD5, HexFormat)
    
    If Len(version) = 0 Then version = "None"
    
    If Not recordExists(fpath, CStr(fhash)) Then
        ado.Insert "tblFiles", "name,path,version,trusted,hash", name, fpath, version, 1, fhash
    End If
    
    If Err.Number > 0 Then
        Debug.Print "Error AddKnown file=" & fpath & " desc= " & Err.Description
    End If
    
End Function

Private Sub Class_Initialize()
    
    Dim db As String
    
    HideKnownInDisplays = IIf(GetSetting("sysanalyzer", "settings", "HideKnwon", 0) = 0, False, True)
    
    If IsIde Then
        db = App.path & "\..\..\known_files.mdb"
    Else
        db = App.path & "\known_files.mdb"
    End If
    
    If Not fso.FileExists(db) Then
        Ready = False
        MsgBox "Could not locate database file"
    Else
        Ready = True
        ado.BuildConnectionString Access, db
        If TotalRecords > 0 Then Loaded = True
    End If

End Sub

 
Private Sub Class_Terminate()
    SaveSetting "sysanalyzer", "settings", "HideKnwon", IIf(HideKnownInDisplays, 1, 0)
End Sub

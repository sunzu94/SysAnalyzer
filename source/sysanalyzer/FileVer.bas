Attribute VB_Name = "FileProps"
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

'Used in several projects do not change interface!

Public Declare Function StartWatch Lib "dir_watch.dll" (ByVal dirPath As String) As Long
Public Declare Function CloseWatch Lib "dir_watch.dll" (ByVal threadID As Long) As Long

Public Declare Function IDEStartWatch Lib "./../../dir_watch.dll" Alias "StartWatch" (ByVal dirPath As String) As Long
Public Declare Function IDECloseWatch Lib "./../../dir_watch.dll" Alias "CloseWatch" (ByVal threadID As Long) As Long

Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Global fso As New clsFileSystem
Global dlg As New clsCmnDlg2 'comdlg threadlocks on main form?! even MS one does..
Global hash As New CWinHash

#If isSysanalyzer = 1 Then
    Global diff As New CSysDiff
    Global ado As New clsAdoKit
    Global known As New CKnownFile
    Global apiDataManager As New CApiDataManager
#End If

Public Const x64Error = "This feature is only currently available for 32 bit processes."

Global ProcessesToRWEScan As String
Global tcpdump As String
Global networkAnalyzer As String
Global watchIDs() As Long
Global watchDirs As New Collection
Global cApiData As New Collection
Global cLogData As New Collection
Global DebugLogFile As String
Global START_TIME As Date
Global procWatchPID As Long

Global Const LANG_US = &H409
Private Const HWND_NOTOPMOST = -2
Private Const HWND_TOPMOST = -1
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32" (ByVal hWndOwner As Long, ByVal nFolder As Long, pidl As Long) As Long
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv As Long)
Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long
Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal path As String, ByVal cbBytes As Long) As Long
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, ByVal Source As Long, ByVal Length As Long)
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Public Type FILEPROPERTIE
    CompanyName As String
    FileDescription As String
    FileVersion As String
    InternalName As String
    LegalCopyright As String
    OrigionalFileName As String
    ProductName As String
    ProductVersion As String
    LanguageID As String
End Type

Public Type COPYDATASTRUCT
    dwFlag As Long
    cbSize As Long
    lpData As Long
End Type

Public Const WM_COPYDATA = &H4A
Public Const WM_DISPLAY_TEXT = 3

Private Const LANG_BULGARIAN = &H2
Private Const LANG_CHINESE = &H4
Private Const LANG_CROATIAN = &H1A
Private Const LANG_CZECH = &H5
Private Const LANG_DANISH = &H6
Private Const LANG_DUTCH = &H13
Private Const LANG_ENGLISH = &H9
Private Const LANG_FINNISH = &HB
Private Const LANG_FRENCH = &HC
Private Const LANG_GERMAN = &H7
Private Const LANG_GREEK = &H8
Private Const LANG_HUNGARIAN = &HE
Private Const LANG_ICELANDIC = &HF
Private Const LANG_ITALIAN = &H10
Private Const LANG_JAPANESE = &H11
Private Const LANG_KOREAN = &H12
Private Const LANG_NEUTRAL = &H0
Private Const LANG_NORWEGIAN = &H14
Private Const LANG_POLISH = &H15
Private Const LANG_PORTUGUESE = &H16
Private Const LANG_ROMANIAN = &H18
Private Const LANG_RUSSIAN = &H19
Private Const LANG_SLOVAK = &H1B
Private Const LANG_SLOVENIAN = &H24
Private Const LANG_SPANISH = &HA
Private Const LANG_SWEDISH = &H1D
Private Const LANG_TURKISH = &H1F

Private Type sockaddr_in
    sin_family As Integer
    sin_port As Integer
    sin_addr As Long
    sin_zero As String * 8
End Type

Private Type sockaddr_gen
    AddressIn As sockaddr_in
    filler(0 To 7) As Byte
End Type

Private Type INTERFACE_INFO
    iiFlags  As Long
    iiAddress As sockaddr_gen
    iiBroadcastAddress As sockaddr_gen
    iiNetmask As sockaddr_gen
End Type

Private Type INTERFACEINFO
    iInfo(0 To 7) As INTERFACE_INFO
End Type

Private Const WSADESCRIPTION_LEN As Long = 256
Private Const WSASYS_STATUS_LEN  As Long = 128

Private Type WSAData
    wVersion As Integer
    wHighVersion As Integer
    szDescription As String * WSADESCRIPTION_LEN
    szSystemStatus As String * WSASYS_STATUS_LEN
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type

Private Declare Function socket Lib "ws2_32.dll" (ByVal af As Long, ByVal s_type As Long, ByVal Protocol As Long) As Long
Private Declare Function closesocket Lib "ws2_32.dll" (ByVal s As Long) As Long
Private Declare Function WSAIoctl Lib "ws2_32.dll" (ByVal s As Long, ByVal dwIoControlCode As Long, lpvInBuffer As Any, ByVal cbInBuffer As Long, lpvOutBuffer As Any, ByVal cbOutBuffer As Long, lpcbBytesReturned As Long, lpOverlapped As Long, lpCompletionRoutine As Long) As Long
Private Declare Sub CopyMemory2 Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, ByVal pSrc As Long, ByVal ByteLen As Long)
Private Declare Function WSAStartup Lib "ws2_32.dll" (ByVal wVR As Long, lpWSAD As WSAData) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Public Sub AlwaysOnTop(f As Form, Optional SetOnTop As Boolean = True)
    Dim lflag As Long, tx As Long, ty As Long
     
    If isIde() Then Exit Sub 'we dont need this in our way when were debugging...
    
    tx = Screen.TwipsPerPixelX
    ty = Screen.TwipsPerPixelY
    
    lflag = IIf(SetOnTop, HWND_TOPMOST, HWND_NOTOPMOST)
     
    SetWindowPos f.hwnd, lflag, f.Left / tx, f.top / ty, f.Width / tx, f.Height / ty, SWP_NOACTIVATE Or SWP_SHOWWINDOW
    
End Sub

Public Sub debugLog(ByVal msg)
    On Error Resume Next
    Dim timestamp As String
    timestamp = Format(DateDiff("s", START_TIME, Now), "0###s > ")
    If InStr(msg, vbCrLf) > 0 Then msg = Replace(msg, vbCrLf, vbCrLf & vbTab) & vbCrLf
    fso.AppendFile DebugLogFile, timestamp & msg
End Sub

Public Function GetShortName(sFile As String) As String
    Dim sShortFile As String * 67
    Dim lResult As Long
    
    'the path must actually exist to get the short path name !!
    If Not fso.FileExists(sFile) Then 'fso.WriteFile sFile, ""
        GetShortName = sFile
        Exit Function
    End If
        
    'Make a call to the GetShortPathName API
    lResult = GetShortPathName(sFile, sShortFile, _
    Len(sShortFile))

    'Trim out unused characters from the string.
    GetShortName = Left$(sShortFile, lResult)
    
    If Len(GetShortName) = 0 Then GetShortName = sFile

End Function


Public Sub LV_ColumnSort(ListViewControl As ListView, Column As ColumnHeader)
     On Error Resume Next
    With ListViewControl
       If .SortKey <> Column.Index - 1 Then
             .SortKey = Column.Index - 1
             .SortOrder = lvwAscending
       Else
             If .SortOrder = lvwAscending Then
              .SortOrder = lvwDescending
             Else
              .SortOrder = lvwAscending
             End If
       End If
       .Sorted = -1
    End With
End Sub

Function pHex(x)
    y = Hex(x)
    While Len(y) < 8
        y = "0" & y
    Wend
    pHex = y
End Function

'todo: try zlib compressibility as another entropy check...
Function CalculateEntropy(ByVal s As String) As Integer 'very basic...
    On Error Resume Next
    If Len(s) = 0 Then Exit Function
    Dim a As Long, b As Long
    a = Len(s)
    's = Replace(s, Chr(0), Empty)
    s = SimpleCompress(s)
    b = Len(s)
    CalculateEntropy = ((b / a) * 100)
End Function


Function LaunchStrings(data As String, Optional isPath As Boolean = False)

    Dim b() As Byte
    Dim f As String
    Dim exe As String
    Dim h As Long
    
    On Error Resume Next
    
    exe = App.path & IIf(isIde(), "\..\..", "") & "\shellext.exe"
    If Not fso.FileExists(exe) Then
        MsgBox "Could not launch strings shellext not found", vbInformation
        Exit Function
    End If
    
    If isPath Then
        If fso.FileExists(data) Then
            f = data
        Else
            MsgBox "Can not launch strings, File not found: " & data, vbInformation
        End If
    Else
        b() = StrConv(dataOrPath, vbFromUnicode, LANG_US)
        f = fso.GetFreeFileName(Environ("temp"), ".bin")
        h = FreeFile
    End If
    
    Open f For Binary As h
    Put h, , b()
    Close h
    
    Shell exe & " """ & f & """ /peek"

End Function

Function LaunchExternalHexViewer(data As String, Optional isPath As Boolean = False, Optional base As String = Empty)

    Dim b() As Byte
    Dim f As String
    Dim exe As String
    Dim h As Long
    
    On Error Resume Next
    
    If Len(base) > 0 Then base = "/base=" & base
    
    exe = App.path & IIf(isIde(), "\..\..", "") & "\shellext.exe"
    If Not fso.FileExists(exe) Then
        MsgBox "Could not launch strings shellext not found", vbInformation
        Exit Function
    End If
    
    If isPath Then
        If fso.FileExists(data) Then
            f = data
        Else
            MsgBox "Can not launch strings, File not found: " & data, vbInformation
            Exit Function
        End If
    Else
        b() = StrConv(data, vbFromUnicode, LANG_US)
        f = fso.GetFreeFileName(Environ("temp"), ".bin")
        h = FreeFile
    End If
    
    Open f For Binary As h
    Put h, , b()
    Close h
    
    Shell exe & " """ & f & """" & IIf(Len(base) > 0, " " & base, "") & " /hexv"

End Function

Sub SaveMySetting(key, Value)
    SaveSetting "iDefense", App.exename, key, Value
End Sub

Function GetMySetting(key, def)
    GetMySetting = GetSetting("iDefense", App.exename, key, def)
End Function

Sub SaveFormSizeAnPosition(f As Form)
    On Error Resume Next
    Dim s As String
    If f.WindowState <> 0 Then Exit Sub 'vbnormal
    s = f.Left & "," & f.top & "," & f.Width & "," & f.Height
    SaveMySetting f.name & "_pos", s
End Sub

Function occuranceCount(haystack, match) As Long
    On Error Resume Next
    Dim tmp
    tmp = Split(haystack, match, , vbTextCompare)
    occuranceCount = UBound(tmp)
    If Err.Number <> 0 Then occuranceCount = 0
End Function

Sub RestoreFormSizeAnPosition(f As Form)

    On Error GoTo hell
    Dim s
    
    s = GetMySetting(f.name & "_pos", "")
    
    If Len(s) = 0 Then Exit Sub
    If occuranceCount(s, ",") <> 3 Then Exit Sub
    
    s = Split(s, ",")
    f.Left = s(0)
    f.top = s(1)
    f.Width = s(2)
    f.Height = s(3)
    
    Exit Sub
hell:
End Sub

Function HexDump(ByVal str, Optional hexOnly = 0, Optional offset As Long = 0) As String
    Dim s() As String, chars As String, tmp As String
    On Error Resume Next
    Dim ary() As Byte
   
    str = " " & str
    ary = StrConv(str, vbFromUnicode, LANG_US)
    
    chars = "   "
    For i = 1 To UBound(ary)
        tt = Hex(ary(i))
        If Len(tt) = 1 Then tt = "0" & tt
        tmp = tmp & tt & " "
        x = ary(i)
        'chars = chars & IIf((x > 32 And x < 127) Or x > 191, Chr(x), ".") 'x > 191 causes \x0 problems on non us systems... asc(chr(x)) = 0
        chars = chars & IIf((x > 32 And x < 127), Chr(x), ".")
        If i > 1 And i Mod 16 = 0 Then
            h = Hex(offset)
            While Len(h) < 6: h = "0" & h: Wend
            If hexOnly = 0 Then
                push s, h & "   " & tmp & chars
            Else
                push s, tmp
            End If
            offset = offset + 16
            tmp = Empty
            chars = "   "
        End If
    Next
    'if read length was not mod 16=0 then
    'we have part of line to account for
    If tmp <> Empty Then
        If hexOnly = 0 Then
            h = Hex(offset)
            While Len(h) < 6: h = "0" & h: Wend
            h = h & "   " & tmp
            While Len(h) <= 56: h = h & " ": Wend
            push s, h & chars
        Else
            push s, tmp
        End If
    End If
    
    HexDump = Join(s, vbCrLf)
    
    If hexOnly <> 0 Then
        HexDump = Replace(HexDump, " ", "")
        HexDump = Replace(HexDump, vbCrLf, "")
    End If
    
End Function


Function AvailableInterfaces() As Collection
  
    Dim hSocket As Long, size As Long, count As Integer
    Dim i As Integer, lngIp As Long, ip(3) As Byte
    Dim sIp As String
    Dim ret As New Collection
    Dim buf As INTERFACEINFO
    Dim WSAInfo As WSAData
    
    Const SIO_GET_INTERFACE_LIST As Long = &H4004747F
    Const INVALID_SOCKET As Long = 0
    Const SOCKET_ERROR As Long = -1
    Const AF_INET As Long = 2

    On Error GoTo failed
    Set AvailableInterfaces = ret
      
    WSAStartup &H202, WSAInfo
    hSocket = socket(AF_INET, 1, 0)
    If hSocket = INVALID_SOCKET Then Exit Function
    If WSAIoctl(hSocket, SIO_GET_INTERFACE_LIST, ByVal 0, 0, buf, 1024, size, ByVal 0, ByVal 0) Then GoTo failed
    
    count = CInt(size / 76) - 1
     
    For i = 0 To count
        lngIp = buf.iInfo(i).iiAddress.AddressIn.sin_addr
        CopyMemory2 ByVal VarPtr(ip(0)), VarPtr(lngIp), 4
        sIp = ip(0) & "." & ip(1) & "." & ip(2) & "." & ip(3)
        ret.Add sIp, sIp
    Next i
    
failed:
    closesocket hSocket
    
End Function

Sub DirWatchCtl(enable As Boolean)
    Dim i As Integer, d
     
    If enable Then
        Erase watchIDs
        For Each d In watchDirs
            If Len(d) > 0 Then
                If isIde Then
                    push watchIDs(), IDEStartWatch(d)
                Else
                    push watchIDs(), StartWatch(d)
                End If
                DoEvents
                Sleep 20
            End If
        Next
    Else
        If Not AryIsEmpty(watchIDs) Then
            For i = 0 To UBound(watchIDs)
                If isIde() Then
                    IDECloseWatch watchIDs(i)
                Else
                    CloseWatch watchIDs(i)
                End If
            Next
        End If
    End If
   
End Sub

Function QuickInfo(fileName As String)
    Dim f As FILEPROPERTIE
    
    f = FileInfo(fileName)
    
    QuickInfo = "CompanyName      " & f.CompanyName & vbCrLf & _
                "FileDescription  " & f.FileDescription & vbCrLf & _
                "FileVersion      " & f.FileVersion & vbCrLf & _
                "InternalName     " & f.InternalName & vbCrLf & _
                "LegalCopyright   " & f.LegalCopyright & vbCrLf & _
                "OriginalFilename " & f.OrigionalFileName & vbCrLf & _
                "ProductName      " & f.ProductName & vbCrLf & _
                "ProductVersion   " & FileInfo.ProductVersion
                

End Function

Public Function FileInfo(Optional ByVal PathWithFilename As String) As FILEPROPERTIE
    ' return file-properties of given file  (EXE , DLL , OCX)
    'http://support.microsoft.com/default.aspx?scid=kb;en-us;160042
    
    If Len(PathWithFilename) = 0 Then
        Exit Function
    End If
    
    Dim lngBufferlen As Long
    Dim lngDummy As Long
    Dim lngRc As Long
    Dim lngVerPointer As Long
    Dim lngHexNumber As Long
    Dim bytBuffer() As Byte
    Dim bytBuff() As Byte
    Dim strBuffer As String
    Dim strLangCharset As String
    Dim strVersionInfo(7) As String
    Dim strTemp As String
    Dim intTemp As Integer
           
    ReDim bytBuff(500)
    
    ' size
    lngBufferlen = GetFileVersionInfoSize(PathWithFilename, lngDummy)
    If lngBufferlen > 0 Then
    
       ReDim bytBuffer(lngBufferlen)
       lngRc = GetFileVersionInfo(PathWithFilename, 0&, lngBufferlen, bytBuffer(0))
       
       If lngRc <> 0 Then
          lngRc = VerQueryValue(bytBuffer(0), "\VarFileInfo\Translation", lngVerPointer, lngBufferlen)
          If lngRc <> 0 Then
             'lngVerPointer is a pointer to four 4 bytes of Hex number,
             'first two bytes are language id, and last two bytes are code
             'page. However, strLangCharset needs a  string of
             '4 hex digits, the first two characters correspond to the
             'language id and last two the last two character correspond
             'to the code page id.
             MoveMemory bytBuff(0), lngVerPointer, lngBufferlen
             lngHexNumber = bytBuff(2) + bytBuff(3) * &H100 + bytBuff(0) * &H10000 + bytBuff(1) * &H1000000
             strLangCharset = Hex(lngHexNumber)
             'now we change the order of the language id and code page
             'and convert it into a string representation.
             'For example, it may look like 040904E4
             'Or to pull it all apart:
             '04------        = SUBLANG_ENGLISH_USA
             '--09----        = LANG_ENGLISH
             ' ----04E4 = 1252 = Codepage for Windows:Multilingual
             Do While Len(strLangCharset) < 8
                 strLangCharset = "0" & strLangCharset
             Loop
             
             If Mid(strLangCharset, 2, 2) = LANG_ENGLISH Then
               strLangCharset2 = "English (US)"
             End If

             If Mid(strLangCharset, 2, 2) = LANG_BULGARIAN Then strLangCharset2 = "Bulgarian"
             If Mid(strLangCharset, 2, 2) = LANG_FRENCH Then strLangCharset2 = "French"
             If Mid(strLangCharset, 2, 2) = LANG_NEUTRAL Then strLangCharset2 = "Neutral"

             Do While Len(strLangCharset) < 8
                 strLangCharset = "0" & strLangCharset
             Loop

             ' assign propertienames
             strVersionInfo(0) = "CompanyName"
             strVersionInfo(1) = "FileDescription"
             strVersionInfo(2) = "FileVersion"
             strVersionInfo(3) = "InternalName"
             strVersionInfo(4) = "LegalCopyright"
             strVersionInfo(5) = "OriginalFileName"
             strVersionInfo(6) = "ProductName"
             strVersionInfo(7) = "ProductVersion"
             
             Dim n As Long
             
             ' loop and get fileproperties
             For intTemp = 0 To 7
                strBuffer = String$(800, 0)
                strTemp = "\StringFileInfo\" & strLangCharset & "\" & strVersionInfo(intTemp)
                lngRc = VerQueryValue(bytBuffer(0), strTemp, lngVerPointer, lngBufferlen)
                If lngRc <> 0 Then
                   ' get and format data
                   lstrcpy strBuffer, lngVerPointer
                   n = InStr(strBuffer, Chr(0)) - 1
                   If n > 0 Then
                        strBuffer = Mid$(strBuffer, 1, n)
                        strVersionInfo(intTemp) = strBuffer
                   End If
                 Else
                   ' property not found
                   strVersionInfo(intTemp) = ""
                End If
             Next intTemp
             
          End If
       End If
    End If
    
    ' assign array to user-defined-type
    FileInfo.CompanyName = strVersionInfo(0)
    FileInfo.FileDescription = strVersionInfo(1)
    FileInfo.FileVersion = strVersionInfo(2)
    FileInfo.InternalName = strVersionInfo(3)
    FileInfo.LegalCopyright = strVersionInfo(4)
    FileInfo.OrigionalFileName = strVersionInfo(5)
    FileInfo.ProductName = strVersionInfo(6)
    FileInfo.ProductVersion = strVersionInfo(7)
    FileInfo.LanguageID = strLangCharset2
    
End Function




Sub push(ary, Value) 'this modifies parent ary object
    On Error GoTo init
    Dim x As Integer
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = Value
    Exit Sub
init:     ReDim ary(0): ary(0) = Value
End Sub


Function objKeyExistsInCollection(c As Collection, val As String) As Boolean
    On Error GoTo nope
    Dim t
    Set t = c(val)
    Set t = Nothing
    objKeyExistsInCollection = True
 Exit Function
nope: objKeyExistsInCollection = False
End Function



Function KeyExistsInCollection(c As Collection, val As String) As Boolean
    On Error GoTo nope
    Dim t
    t = c(val)
    KeyExistsInCollection = True
 Exit Function
nope: KeyExistsInCollection = False
End Function

Function FileExists(path) As Boolean
  On Error Resume Next
  If Len(path) = 0 Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True
End Function



Function AryIsEmpty(ary) As Boolean
  On Error GoTo oops
    i = UBound(ary)  '<- throws error if not initalized
    AryIsEmpty = False
  Exit Function
oops: AryIsEmpty = True
End Function


Function GetAllElements(lv As ListView) As String
    Dim ret() As String, i As Integer, tmp As String
    Dim li As ListItem

    For i = 1 To lv.ColumnHeaders.count
        tmp = tmp & lv.ColumnHeaders(i).Text & vbTab
    Next

    push ret, tmp
    push ret, String(50, "-")

    For Each li In lv.ListItems
        tmp = li.Text & vbTab
        For i = 1 To lv.ColumnHeaders.count - 1
            tmp = tmp & li.SubItems(i) & vbTab
        Next
        push ret, tmp
    Next

    GetAllElements = Join(ret, vbCrLf)

End Function

Function GetAllText(lv As ListView, Optional subItemRow As Long = 0, Optional selectedOnly As Boolean = False) As String
    Dim i As Long
    Dim tmp As String, x As String
    
    For i = 1 To lv.ListItems.count
        If subItemRow = 0 Then
            x = lv.ListItems(i).Text
            If selectedOnly And Not lv.ListItems(i).Selected Then x = Empty
            If Len(x) > 0 Then
                tmp = tmp & x & vbCrLf
            End If
        Else
            x = lv.ListItems(i).SubItems(subItemRow)
            If selectedOnly And Not lv.ListItems(i).Selected Then x = Empty
            If Len(x) > 0 Then
                tmp = tmp & x & vbCrLf
            End If
        End If
    Next
    
    GetAllText = tmp
End Function


Function ReadFile(fileName)
Dim f, Temp
  f = FreeFile
  Temp = ""
   Open fileName For Binary As #f        ' Open file.(can be text or image)
     Temp = Input(FileLen(fileName), #f) ' Get entire Files data
   Close #f
   ReadFile = Temp
End Function




Function AnyOfTheseInstr(ByVal sIn, sCmp) As Boolean
    Dim tmp() As String, i As Integer
    tmp() = Split(sCmp, ",")
    sIn = LCase(sIn)
    For i = 0 To UBound(tmp)
        tmp(i) = LCase(Trim(tmp(i)))
        If Len(tmp(i)) > 0 And InStr(1, sIn, tmp(i), vbTextCompare) > 0 Then
            AnyOfTheseInstr = True
            Exit Function
        End If
    Next
End Function


Public Function UserDeskTopFolder() As String
    Dim idl As Long
    Dim p As String
    Const MAX_PATH As Long = 260
      
      p = String(MAX_PATH, Chr(0))
      If SHGetSpecialFolderLocation(0, 0, idl) <> 0 Then Exit Function
      SHGetPathFromIDList idl, p
      
      UserDeskTopFolder = Left(p, InStr(p, Chr(0)) - 1)
      CoTaskMemFree idl
        
      UserDeskTopFolder = UserDeskTopFolder & "\analysis"
      
      If Not fso.FolderExists(UserDeskTopFolder) Then
            fso.CreateFolder UserDeskTopFolder
      End If
  
End Function
 
Sub ScanProcsForDll(Optional lblDisplay As Label = Nothing)
    Dim cp As New CProcessInfo
    Dim c As Collection
    Dim m As Collection
    Dim p As CProcess
    Dim cm As CModule
    Dim ret As String
    Dim i As Long
    Dim hit As Boolean
    Dim tmp As String
    Dim tmp2 As String
    Dim mm As matchModes
    
    'On Error Resume Next
    
    Dim find As String
    find = InputBox("Enter string fragment of what to look for in dll name or path.")
    If Len(find) = 0 Then Exit Sub
    
    If Not lblDisplay Is Nothing Then lblDisplay.Caption = "Starting scan..."
    
    i = 0
    
    Set c = cp.GetRunningProcesses()
    For Each p In c
        If p.pid <> 0 And p.pid <> 4 Then
            If Not lblDisplay Is Nothing Then
                lblDisplay.Caption = "Scanning " & i & "/" & c.count
                lblDisplay.Refresh
            End If
            DoEvents
            Set m = cp.GetProcessModules(p.pid)
            If Not m Is Nothing Then
                If m.count > 0 Then
                    tmp = "pid: " & p.pid & " " & p.path
                    hit = False
                    tmp2 = Empty
                    For Each cm In m
                        If InStr(1, cm.path, find, vbTextCompare) > 0 Then
                           tmp2 = tmp2 & vbTab & Hex(cm.base) & vbTab & cm.path & vbCrLf
                           hit = True
                        End If
                    Next
                    If hit Then ret = ret & tmp & tmp2
                End If
            End If
            i = i + 1
        End If
    Next
    
    If Not lblDisplay Is Nothing Then lblDisplay.Caption = ""
    
    If Len(ret) > 0 Then
        frmReport.ShowList vbCrLf & Replace(ret, Chr(0), Empty)
    Else
        MsgBox "No modules found in any process matching your criteria"
    End If
    

End Sub

Sub ScanForUnknownMods(Optional lbl As Label = Nothing)
    Dim cp As New CProcessInfo
    Dim c As Collection
    Dim m As Collection
    Dim p As CProcess
    Dim cm As CModule
    Dim ret As String
    Dim i As Long
    Dim hit As Boolean
    Dim tmp As String
    Dim tmp2 As String
    Dim mm As matchModes
    
    'On Error Resume Next
    
    If Not known.Loaded Then
        MsgBox "Known database is not loaded..", vbInformation
        Exit Sub
    End If
    
    'ado.OpenConnection
    If Not lbl Is Nothing Then lbl.Caption = "Starting scan..."
    
    i = 0
    
    Set c = cp.GetRunningProcesses()
    For Each p In c
        If p.pid <> 0 And p.pid <> 4 Then
            If Not lbl Is Nothing Then lbl.Caption = "Scanning " & i & "/" & c.count
            Set m = cp.GetProcessModules(p.pid)
            If Not m Is Nothing And m.count > 0 Then
                tmp = "pid: " & p.pid & " " & p.path
                hit = False
                tmp2 = Empty
                For Each cm In m
                    mm = known.isFileKnown(cm.path)
                    If mm <> exact_match Then
                       tmp2 = tmp2 & vbCrLf & vbTab & IIf(mm = not_found, "Unknown Mod: ", "Hash Changed: ") & cm.path
                       hit = True
                    End If
                Next
                If hit Then ret = ret & tmp & tmp2 & vbCrLf & vbCrLf
            End If
            i = i + 1
            DoEvents
            If Not lbl Is Nothing Then lbl.Refresh
        End If
    Next
    
    If Not lbl Is Nothing Then lbl.Caption = ""
    'ado.CloseConnection
    
    Const header = "This list may also include files were locked at the time the database was created and could not be hashed for that reason."
    
    If Len(ret) > 0 Then
        frmReport.ShowList vbCrLf & header & vbCrLf & vbCrLf & Replace(ret, Chr(0), Empty)
    Else
        MsgBox "No unknown modules found in any process..."
    End If
    
    
End Sub

Function isIde() As Boolean
    On Error GoTo hell
    Debug.Print 1 \ 0
Exit Function
hell: isIde = True
End Function

Public Function MD5File(f As String) As String
    MD5File = hash.HashFile(f)
End Function


Function isNetworkAnalyzerRunning() As Boolean

    Const vbIDEClassName = "ThunderFormDC"
    Const vbEXEClassName = "ThunderRT6FormDC"
    Const vbWindowCaption = "Packet Sniffer"
    
    Dim hServer As Long
    
    hServer = FindWindow(vbIDEClassName, vbWindowCaption)
    If hServer = 0 Then hServer = FindWindow(vbEXEClassName, vbWindowCaption)
    If hServer <> 0 Then isNetworkAnalyzerRunning = True
    
End Function

Sub SetLiColor(li As ListItem, newcolor As Long)
    Dim f As ListSubItem
'    On Error Resume Next
    li.ForeColor = newcolor
    For Each f In li.ListSubItems
        f.ForeColor = newcolor
    Next
End Sub

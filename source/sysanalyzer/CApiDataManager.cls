VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CApiDataManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim handles As New Collection 'of CApiHandle
'handles numbers are unique (pertype? or across all?) but can be reused once closed...

Private Function AddHandle(data, h, rtype As rTypes)
    Dim ch As CApiHandle
    If HandleExists(h, rtype) Then Exit Function
    Set ch = New CApiHandle
    ch.Resource = data
    ch.handle = h
    ch.rtype = rtype
    handles.Add ch
End Function

Private Function HandleExists(h, rtype As rTypes) As Boolean
    Dim ch As CApiHandle
    For Each ch In handles
        If ch.handle = h And ch.rtype = rtype Then
            HandleExists = True
            Exit Function
        End If
    Next
End Function

Private Function RemoveHandle(h, rtype As rTypes)
    Dim i As Long
    Dim cw As CApiHandle
    For i = 1 To handles.count
        Set cw = handles(i)
        If cw.handle = h And cw.rtype = rtype Then
            Set cw = Nothing
            handles.Remove i
        End If
    Next
End Function


Function HandleApiMessage(msg As String)
    'on CloseHandle - copy files to /analysis
    'move WriteProcessMemory injections from c:\ to /analysis/injects
    'keep readable report of files/ports/connections ? data send/received?
    'extract buffers from process for send/recv/wpm
End Function

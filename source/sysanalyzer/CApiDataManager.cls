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
Dim Handles As New Collection 'of CApiHandle
'handles numbers are unique (pertype? or across all?) but can be reused once closed...

Enum captureTypes
    ct_file = 0
    ct_WriteProcMem = 1
    ct_send = 2
    ct_recv = 3
End Enum

Event CaptureMade(t As captureTypes, data As String)

Private Function AddHandle(data, h, rtype As rTypes)
    Dim ch As CApiHandle
    If HandleExists(h, rtype) Then Exit Function
    Set ch = New CApiHandle
    ch.Resource = data
    ch.handle = h
    ch.rtype = rtype
    Handles.Add ch
End Function

Private Function HandleExists(h, rtype As rTypes) As Boolean
    Dim ch As CApiHandle
    For Each ch In Handles
        If ch.handle = h And ch.rtype = rtype Then
            HandleExists = True
            Exit Function
        End If
    Next
End Function

Private Function RemoveHandle(h, rtype As rTypes)
    Dim i As Long
    Dim cw As CApiHandle
    For i = 1 To Handles.Count
        Set cw = Handles(i)
        If cw.handle = h And cw.rtype = rtype Then
            Set cw = Nothing
            Handles.Remove i
        End If
    Next
End Function


Function HandleApiMessage(msg)
    'on CloseHandle - copy files to /analysis
    'move WriteProcessMemory injections from c:\ to /analysis/injects
    'keep readable report of files/ports/connections ? data send/received?
    'extract buffers from process for send/recv/wpm
    
    'dataformat: pid,retAddr message
    
    
    
End Function

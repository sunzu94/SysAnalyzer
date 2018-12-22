VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Caption As String
Public Description As String
Public DisplayName As String
Public Name As String
Public PathName As String
Public ServiceType As String
Public Started As Boolean
Public StartMode As String
Public StartName As String
Public State As String
Public Status As String
Public SystemName As String
Public ProcessId As Long

Sub LoadSelf(objItem)

    On Error Resume Next
    
    Caption = objItem.Caption
    Description = objItem.Description
    DisplayName = objItem.DisplayName
    Name = objItem.Name
    PathName = objItem.PathName
    ServiceType = objItem.ServiceType
    Started = objItem.Started
    StartMode = objItem.StartMode
    StartName = objItem.StartName
    State = objItem.State
    Status = objItem.Status
    SystemName = objItem.SystemName
    ProcessId = objItem.ProcessId
    
End Sub

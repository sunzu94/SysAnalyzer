VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CColItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Name As String
Public Data As String
Public HashCode As String

Function GenHashCode() As String
    HashCode = hash.HashString(Name & Data)
    GenHashCode = HashCode
End Function

Sub AppendData(x, Optional seperator = vbCrLf)
    If Len(Data) = 0 Then
        Data = x
    Else
        Data = Data & seperator & x
    End If
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
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

'Used in several projects do not change interface!
Public pid As Long
Public path As String
Public cmdLine As String
Public ParentPID As Long
Public User As String
Public base As Long
Public size As Long
Public is64Bit As Boolean
Public domain As String
Public fullpath As String
Public SessionId As Long

Public PkgFullName
Public PkgName
Public PkgPublisher
Public PkgErr
Private pkgQueried As Boolean

Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

'win8+ only
Private Declare Function GetPackageFullName Lib "kernel32" (ByVal hProcess As Long, ByRef packageFullNameLength As Long, ByRef PackageFullName As Byte) As Long
'buffer contains PACKAGE_ID header, PWSTR pointers point to unicode strings held in remainder of buffer following header
Private Declare Function GetPackageId Lib "kernel32" (ByVal hProcess As Long, ByRef bufferLength As Long, ByRef buffer As Byte) As Long

Private Type PACKAGE_ID
  reserved As Long
  processorArchitecture  As Long
  revision As Integer
  build As Integer
  minor As Integer
  major As Integer
  name  As Long
  publisher  As Long
  resourceId  As Long
  publisherId  As Long
End Type

Property Get isAppStorePkg() As Boolean

    Dim hProcess As Long
    Dim length As Long, rc As Long, buf() As Byte
    Dim pack As PACKAGE_ID
    Dim base As Long
    
    Const PROCESS_QUERY_LIMITED_INFORMATION = &H1000
    Const ERROR_INSUFFICIENT_BUFFER = 122 '(0x7A)
    Const APPMODEL_ERROR_NO_PACKAGE = 15700
      
    On Error Resume Next
    If pkgQueried Then Exit Property
    
    pkgQueried = True
    
    If osInfo.WinVer < ev_Win8 Then Exit Property
    
    hProcess = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
    If hProcess = 0 Then
        PkgErr = "error opening pid"
        Exit Property
    End If
    
    rc = GetPackageFullName(hProcess, length, 0)
    
    If rc <> ERROR_INSUFFICIENT_BUFFER Then
        If rc = APPMODEL_ERROR_NO_PACKAGE Then
            PkgErr = "this pid is not a windows store app"
        Else
            PkgErr = "error GetPackageFullName returned " & rc
        End If
        CloseHandle hProcess
        Exit Property
    End If
        
    ReDim buf((length + 2) * 2)
    rc = GetPackageFullName(hProcess, length, buf(0))
    PkgFullName = GetUniString(buf, 0)

    rc = GetPackageId(hProcess, length, 0)
    ReDim buf(length)
    base = VarPtr(buf(0))
    
    rc = GetPackageId(hProcess, length, buf(0))
    
    CopyMemory ByVal VarPtr(pack), ByVal base, LenB(pack)
    
    PkgName = GetUniString(buf, pack.name - base)
    PkgPublisher = GetUniString(buf, pack.publisher - base)
    
    CloseHandle hProcess
    isAppStorePkg = True
   
End Property

Private Function GetUniString(buf() As Byte, ByVal offset As Long) As String
    
    Dim tmp() As Byte
    Dim sz As Long
    Dim i As Long
    
    If offset < 0 Or offset > UBound(buf) Then Exit Function
    
    ReDim tmp(UBound(buf))
    
    For i = offset To UBound(buf)
        If buf(i) = 0 And buf(i + 1) = 0 Then Exit For
        If buf(i) <> 0 Then
            tmp(sz) = buf(i)
            sz = sz + 1
        End If
    Next
    
    If sz > 0 Then
        ReDim Preserve tmp(sz - 1)
        GetUniString = StrConv(tmp, vbUnicode)
    End If
        
End Function



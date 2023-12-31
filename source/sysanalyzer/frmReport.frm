VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmReport 
   Caption         =   "List Data"
   ClientHeight    =   6285
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12075
   LinkTopic       =   "Form1"
   ScaleHeight     =   6285
   ScaleWidth      =   12075
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox Text1 
      Height          =   5865
      Left            =   45
      TabIndex        =   3
      Top             =   315
      Width           =   11940
      _ExtentX        =   21061
      _ExtentY        =   10345
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmReport.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   60
      Left            =   5040
      TabIndex        =   2
      Top             =   1260
      Width           =   60
      _ExtentX        =   106
      _ExtentY        =   106
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmReport.frx":007C
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save "
      Height          =   255
      Left            =   1740
      TabIndex        =   1
      Top             =   0
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Copy"
      Height          =   255
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   1305
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
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

Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal _
    hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx _
    As Long, ByVal cy As Long, ByVal wFlags As Long)
    
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2

Private filesaveName As String
    
Function ShowList(list, Optional modal As Boolean = False, Optional saveName As String, Optional topMost As Boolean = True)
    
    filesaveName = saveName
    
    If topMost Then
        SetWindowPos Me.hwnd, HWND_TOPMOST, Me.Left / 15, _
            Me.Top / 15, Me.Width / 15, _
            Me.Height / 15, 0
    End If
    
    If IsArray(list) Then
        Text1.Text = Join(list, vbCrLf)
    Else
        Text1.Text = list
    End If
    
    On Error Resume Next
    
    If modal Then
        Me.Show 1
    Else
        Me.Show
        Me.ZOrder 0
    End If
    
End Function

 
 

Private Sub Command2_Click()
    Dim Base As String
    On Error Resume Next
    
    If Len(LOGFILEEXT) = 0 Then LOGFILEEXT = ".txt"
    
    If Len(filesaveName) = 0 Then
        Base = fso.GetBaseName(frmMain.samplePath)
        Base = UserDeskTopFolder & "\" & Base & "_" & Format(Now(), "h.nam/pm") & LOGFILEEXT
    Else
        Base = UserDeskTopFolder & "\" & filesaveName
    End If
    
    fso.writeFile Base, Text1.Text
    If Err.Number = 0 Then MsgBox "Saved Successfully as:" & vbCrLf & vbCrLf & Base
    
End Sub

Private Sub Command1_Click()
    Clipboard.Clear
    Clipboard.SetText Text1.Text
End Sub

Private Sub Form_Load()
    
    On Error Resume Next
    Me.Icon = frmMain.Icon
    RestoreFormSizeAnPosition Me
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Text1.Height = Me.Height - Text1.Top - 550
    Text1.Width = Me.Width - Text1.Left - 200
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveFormSizeAnPosition Me
End Sub

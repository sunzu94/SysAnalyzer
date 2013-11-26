VERSION 5.00
Begin VB.Form frmAbout 
   Caption         =   "About SysAnalyzer"
   ClientHeight    =   5475
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10830
   LinkTopic       =   "Form1"
   ScaleHeight     =   5475
   ScaleWidth      =   10830
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   9540
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   1095
      ScaleWidth      =   1155
      TabIndex        =   1
      Top             =   180
      Width           =   1155
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   60
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   10695
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    On Error Resume Next
    Me.Icon = frmMain.Icon
    
    Text1 = vbCrLf & "SysAnalyzer - Version:" & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & _
            "" & vbCrLf & _
            "License:   GPL" & vbCrLf & _
            "Copyright: 2005 iDefense a Verisign Company" & vbCrLf & _
            "Site:      http://labs.idefense.com" & vbCrLf & _
            "Author:    David Zimmer <dzzie@yahoo.com>" & vbCrLf & _
            "" & vbCrLf & _
            "SysAnalyzer is an automated malware analysis solution originally released" & vbCrLf & _
            "under GPL while I was working at iDefense in 2005. It is not a sandbox." & vbCrLf & _
            "It works by analyzing before and after infection system states and live" & vbCrLf & _
            "logging." & vbCrLf & _
            " " & vbCrLf & _
            "iDefense no longer maintains a downloads section of their website," & vbCrLf & _
            "so I have picked the project back up again and continue to do maintenance" & vbCrLf & _
            "and updates on it in my spare time." & vbCrLf & _
            " " & vbCrLf & _
            "As time permits I am slowly working to bring in updated support" & vbCrLf & _
            "for Win7 and x64 systems. It has been primarily developed on WinXP machines" & vbCrLf & _
            " " & vbCrLf & _
            "Source download: https://github.com/dzzie/SysAnalyzer"
            
            
End Sub

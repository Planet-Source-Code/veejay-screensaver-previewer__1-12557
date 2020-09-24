VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Screensaver Previewer by VeeJay"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   3750
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbNames 
      Height          =   315
      Left            =   480
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2880
      Width           =   2775
   End
   Begin VB.ListBox lstPath 
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   2760
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.PictureBox picScreen 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   480
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   2535
      ScaleWidth      =   2760
      TabIndex        =   0
      Top             =   120
      Width           =   2760
      Begin VB.PictureBox picPreview 
         Height          =   1725
         Left            =   240
         ScaleHeight     =   1665
         ScaleWidth      =   2250
         TabIndex        =   1
         Top             =   240
         Width           =   2310
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' This code was written by John Zimmerman and updated by me
' Most of the credit should go to John, but his code generated
' a lot of errors, so i modified it to be usable.
' Sorry about the lack of comments as this code was put together
' in great haste
'
'
' E-Mail: veejay@post.com
' Website: http://www.crosswinds.net/~nicksveejay
' Enjoy the code
'
'
' Later
' VeeJay

Private scrProcess As Long

Private Sub cmbNames_Click()
    lstPath.ListIndex = cmbNames.ListIndex
    TerminateProcess scrProcess, 0
    scrProcess = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell((lstPath.List(lstPath.ListIndex) & " /p" & picPreview.hWnd), vbHide))
End Sub

Private Sub Form_Load()
    FindFiles GetWinDir, False
    FindFiles GetSysDir, False
    MsgBox "I hope you like my first submission", , "Introduction"
    cmbNames.ListIndex = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    TerminateProcess scrProcess, 0
    If MsgBox("Do you think this could be useful?", vbYesNo, "Trash?") = vbYes Then
        MsgBox "Thank you. May you use this code with caution.", , "Careful"
        Unload Me
    Else
        MsgBox "I did try. And remember, this was my first submission to PSC.", , "Oh, well..."
        Unload Me
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    TerminateProcess scrProcess, 0
End Sub

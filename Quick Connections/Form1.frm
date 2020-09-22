VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Quick Connection - Creature"
   ClientHeight    =   810
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5085
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   810
   ScaleWidth      =   5085
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox CAC 
      BackColor       =   &H00000000&
      Caption         =   "Close After Connection"
      ForeColor       =   &H00C0C000&
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   515
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.CommandButton Hlp 
      Caption         =   "?"
      Height          =   315
      Left            =   4560
      TabIndex        =   4
      Top             =   480
      Width           =   375
   End
   Begin VB.CommandButton CanIt 
      Caption         =   "C&ancel"
      Height          =   315
      Left            =   3480
      TabIndex        =   3
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton Cnt 
      Caption         =   "&Connect"
      Height          =   315
      Left            =   2160
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox Cn 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Computer Name:"
      ForeColor       =   &H00C0C000&
      Height          =   195
      Left            =   120
      LinkItem        =   "cn"
      TabIndex        =   0
      Top             =   120
      Width           =   1185
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'////////////////This program was designed by me, Salis\\\\\\\\\\\\\\\\\\\
'It is freeware but please give credit if you decide to send out
'I don't care if you do decide to share with others but again pleas give credit

Public sLine As String ' This will store the desired Computer Name

Private Sub CanIt_Click()
    End 'Ends the program
End Sub

Private Sub Cn_Change()
    sLine = "mstsc /v:" & Cn.Text & " /console" ' As soon as you start to type something
    'the command line is saved to sLine and will later be called by the shell function
End Sub

Private Sub Cn_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Cnt_Click 'Hit Enter and run the connection
End Sub

Private Sub Cnt_Click()
On Error Resume Next
    If Left(Cn.Text, 2) = "\\" Then Cn.Text = Mid(Cn.Text, 3) 'If some one types \\
    'the command line will not understand and freek out, returning an error
    'this single line removes the \\
Dim Shl
    Shl = Shell(sLine, 4) 'This line will call the sLine and run it command.com
    If CAC.Value = 1 Then End 'If the Close After connection is checked the program
    'will end after the connection is made wether or not connection is successful.
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    End 'This Sub is really not needed but it makes sure the program completly ends
End Sub

Private Sub Hlp_Click()
MsgBox "Created by Salis: Creature - CelticNight" & vbNewLine & vbNewLine & _
    "A small remote connection program" & vbNewLine & _
    "Just type the computer name and hit Enter. Target computer MUST have" & vbNewLine & _
    "remote connections turned on otherwise this will not work."
    'Same with this Sub. It's just for the Help button. I ask that you do leave this in as credit to me.
End Sub

VERSION 5.00
Begin VB.Form frmYahtzee 
   BackColor       =   &H80000001&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2880
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6255
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   6255
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2880
      Left            =   0
      Picture         =   "Form2.frx":0000
      ScaleHeight     =   2880
      ScaleWidth      =   6255
      TabIndex        =   0
      Top             =   0
      Width           =   6255
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   600
      Top             =   1140
   End
End
Attribute VB_Name = "frmYahtzee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'***********************************************************************
'This application was explicitly developed for
'PSC(Planet Source Code) Users as an Open Source Project.
'This code is the property of it's author.
'
'If you compile this application you may not redistribute it.
'However, you may use any of this code in you're own application(s).
'
'Alex Smoljanovic, Salex Software (c) 2001-2003
'salex_software@shaw.ca
'***********************************************************************


Private Sub Form_Activate()
 TransWin Me.hwnd 'see TransWin for more info...
End Sub

Private Sub Form_Load()
On Error Resume Next 'on the event of an error resume execution on the next line of this procedure
TransPrep Me.hwnd 'see TransPrep function for more info...
 Timer1.Enabled = True 'enable the timer control
  If PlaySounds = True Then 'if PlaySounds flag evaluates to true then...
   If Not (Dir(App.Path & "\snd\sirens.snd") = "") Then
   'if the file exists then...
    PlaySound App.Path & "\snd\sirens.snd", frmMain.hwnd, SND_FILENAME Or SND_ASYNC
    'play the sound, flag SND_FILENAME and SND_ASYNC specifies that the sound file specified is a file path,
    'and specifies that the function should perform the operation asynchronously rather than synchronously
   End If
  End If
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 TransWin Me.hwnd, False 'see TransWin for more info...
End Sub

Private Sub Timer1_Timer()
 Unload Me 'unload this dialog...
End Sub

VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSndOpt 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Sound Options"
   ClientHeight    =   2340
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4125
   ControlBox      =   0   'False
   HelpContextID   =   300
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   4125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3180
      TabIndex        =   6
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   1920
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      Caption         =   "&Play other sounds (Dice, Sirens, .ect)"
      Height          =   315
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Width           =   3315
   End
   Begin VB.Frame Frame2 
      Caption         =   "Other Sounds"
      Height          =   795
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   3915
   End
   Begin VB.Frame Frame1 
      Caption         =   "Voice Options"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      Begin MSComctlLib.Slider Slider1 
         Height          =   255
         Left            =   780
         TabIndex        =   1
         Top             =   360
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   450
         _Version        =   393216
         Max             =   65535
      End
      Begin VB.Label Label1 
         Caption         =   "Volume:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmSndOpt"
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

Dim memvl& 'dimensionalize memvl as long data type
'NOTE: Object DDS(Direct Speech Synthesis)'s maximum
'volume is the range of an unsigned short integer(65535) or so I beleive.
'For more information on
'Learnout & Haupsie's TruVoice Speech Synthesization engine refer
'to the Speech Software Development Kit(SDK) available for free from
'MSDN online(http://msdn.microsoft.com)

Private Sub Command1_Click()
 PlaySounds = Check1.Value
 'update PlaySounds flag
  Unload Me 'unload this dialog
End Sub

Private Sub Command2_Click()
 frmMain.dss.VolumeLeft = memvl
  frmMain.dss.VolumeRight = memvl
  'update frmMain's object dss(DirectSS Class[Direct Speech Synthesis])'s volume properties
   Unload Me 'unload this dialog
End Sub

Private Sub Form_Activate()
 TransWin Me.hwnd, True 'see TransWin for more info...
End Sub

Private Sub Form_Load()
On Error Resume Next 'on the event of an error resume execution on the next line of this procedure
 memvl = frmMain.dss.VolumeLeft 'store the current volume information
  Slider1.Max = frmMain.dss.MaxVolumeLeft
  'set the slider controls max property
   Slider1.Value = frmMain.dss.VolumeLeft
   'update the volume slider control
  If PlaySounds = False Then
   Check1.Value = 0
  Else
   Check1.Value = 1
  End If
  'update UI
   TransPrep Me.hwnd 'see TransPrep for more info...
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 TransWin Me.hwnd, False 'see TransWin for more info...
End Sub

Private Sub Slider1_Change()
 frmMain.dss.VolumeLeft = Slider1.Value
  frmMain.dss.VolumeRight = Slider1.Value
  'update object dss's volume properties
   Say "Testing, testing, 1 2 3", True 'see Say for more info...
End Sub

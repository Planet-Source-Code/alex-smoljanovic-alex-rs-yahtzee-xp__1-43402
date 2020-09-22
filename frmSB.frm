VERSION 5.00
Begin VB.Form frmSB 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Incompatible Operating System"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6315
   ControlBox      =   0   'False
   Icon            =   "frmSB.frx":0000
   LinkTopic       =   "Form5"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   6315
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1665
      Left            =   4380
      Picture         =   "frmSB.frx":000C
      ScaleHeight     =   1665
      ScaleWidth      =   1875
      TabIndex        =   7
      Top             =   420
      Width           =   1875
   End
   Begin VB.CommandButton cmdIgnore 
      Caption         =   "&Ignore"
      Height          =   375
      Left            =   5460
      TabIndex        =   6
      ToolTipText     =   "Ignore error message and continue to load"
      Top             =   2160
      Width           =   795
   End
   Begin VB.CommandButton cmdAbort 
      Caption         =   "&Abort"
      Default         =   -1  'True
      Height          =   375
      Left            =   4560
      TabIndex        =   5
      ToolTipText     =   "Close the program"
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label lblOS 
      BackStyle       =   0  'Transparent
      Caption         =   "%OS"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   300
      TabIndex        =   4
      Top             =   1560
      Width           =   4035
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Your Operating System:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   60
      TabIndex        =   3
      Top             =   1260
      Width           =   2070
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2. Windows XP (NT 5.1)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   300
      TabIndex        =   2
      Top             =   720
      Width           =   2025
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1. Windows 2000 (NT 5.0)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   300
      TabIndex        =   1
      Top             =   480
      Width           =   2220
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "This application is only compatible with the following operating systems:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   180
      Width           =   6255
   End
End
Attribute VB_Name = "frmSB"
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


Private Sub Form_Load()
Dim OSDesc$: RetOSInf OSDesc
'dimensionalize OSDesc as string data type
'call RetOSInf to initialize OSDesc with the description of the operating system
'see RetOSInf for more info...
 lblOS.Caption = OSDesc$
 'update the label's caption property
End Sub

Private Sub cmdAbort_Click()
 Unload Me 'unload this dialog
End Sub

Private Sub cmdIgnore_Click()
On Error GoTo errh 'on the event of an error jump to label errh(error handle)
 Me.Hide 'call this form's Hide method to set this dialogs visibility to hidden
  LoadHighScores 'see LoadHighScores for more info...
   If Not (Dir(App.Path & "\YahtzeeXP.HLP") = "") Then
    App.HelpFile = App.Path & "\YahtzeeXP.HLP"
   Else
    If Dir(App.Path & "\YahtzeeXP.HLP") = "" Then
     If MsgBox("Can't locate help file." & vbCrLf & vbCrLf & "If you have moved the help file associated with this program(Yahtzee XP) please return it to the installation directory and click Retry, otherwise click Cancel", vbCritical + vbRetryCancel, "Help File Missing") = vbRetry Then
      If Not (Dir(App.Path & "\YahtzeeXP.HLP") = "") Then
       App.HelpFile = App.Path & "\YahtzeeXP.HLP"
        MsgBox "Help file found." & vbCrLf & """" & App.HelpFile & """", vbInformation, "Help File"
      Else
       If Not (Dir(App.Path & "YahtzeeXP.HLP") = "") Then
        App.HelpFile = App.Path & "\YahtzeeXP.HLP"
         MsgBox "Help file found." & vbCrLf & """" & App.HelpFile & """", vbInformation, "Help File"
       End If
      End If
     End If
    Else
     App.HelpFile = App.Path & "\YahtzeeXP.HLP"
    End If
   End If
   'determine if the help file exists, if so, set this applications help file filepath
    Load frmMain 'load frmMain into memory(initialize dialog)
     frmMain.Show 'call frmMain's Show method to show the dialog
      Unload Me 'unload this dialog
       Exit Sub 'exit this procedure
errh: 'label errh
 MsgBox "The following error occured:" & vbCrLf & Err.Description, "Error " & Err.Number, vbCritical
 'inform user of the error which occured
  Unload Me
  'unload this dialog
End Sub


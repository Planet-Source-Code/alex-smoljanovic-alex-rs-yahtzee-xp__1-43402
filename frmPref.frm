VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmPref 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Preferences"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4755
   LinkTopic       =   "Form6"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   4755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   2700
      TabIndex        =   9
      Top             =   2700
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Default         =   -1  'True
      Height          =   375
      Left            =   3660
      TabIndex        =   8
      Top             =   2700
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "Settings"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   720
      TabIndex        =   2
      Top             =   780
      Width           =   3675
      Begin MSComctlLib.Slider sldSpeed 
         Height          =   255
         Left            =   960
         TabIndex        =   3
         Top             =   420
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   450
         _Version        =   393216
         Min             =   1
         Max             =   30
         SelStart        =   15
         Value           =   15
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fast"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Left            =   3060
         TabIndex        =   6
         Top             =   720
         Width           =   285
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Slow"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Left            =   1020
         TabIndex        =   5
         Top             =   720
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Speed:"
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
         Left            =   240
         TabIndex        =   4
         Top             =   420
         Width           =   615
      End
   End
   Begin VB.CheckBox chkWinTrans 
      Caption         =   "Window Transitions"
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
      Left            =   360
      TabIndex        =   1
      Top             =   480
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Caption         =   "Graphical User Interface Options"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4515
      Begin VB.CheckBox chkGraphTrans 
         Caption         =   "Other graphical transitions"
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
         Left            =   240
         TabIndex        =   7
         Top             =   2040
         Width           =   2715
      End
   End
End
Attribute VB_Name = "frmPref"
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


Private Sub chkWinTrans_Click()
 sldSpeed.Enabled = CBool(chkWinTrans.Value)
 'conditionally enable the sldSpeed slider control
End Sub

Private Sub cmdCancel_Click()
 Unload Me 'unload this dialog
End Sub

Private Sub cmdOk_Click()
 wndTransSpeed = sldSpeed.Value: SaveSetting "YAH", "Pref", "wndTransS", Str$(wndTransSpeed)
  wndTrans = chkWinTrans.Value: SaveSetting "YAH", "Pref", "wndTrans", CB(wndTrans)
   GraphTrans = chkGraphTrans.Value: SaveSetting "YAH", "Pref", "GraphTrans", CB(GraphTrans)
   'load the users preferences from the registry
    Unload Me 'unload this dialog
End Sub

Private Function CB(Bv As Boolean) As String
'typical BoolToStr conversion...
 If Bv = True Then CB = "1" Else CB = "0"
End Function

Private Sub Form_Activate()
 TransWin Me.hwnd 'see TransWin for more info...
End Sub

Private Sub Form_Load()
On Error Resume Next 'on the event of an error resume execution of this procedure on the next line
 sldSpeed.Value = wndTransSpeed 'update slider control
  If wndTrans = True Then chkWinTrans.Value = 1: sldSpeed.Enabled = True Else chkWinTrans.Value = 0: sldSpeed.Enabled = False
  'update control properties
   If GraphTrans = True Then chkGraphTrans.Value = 1 Else chkGraphTrans.Value = 0
   '...
    TransPrep Me.hwnd 'see TransPrep function for more info...
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 TransWin Me.hwnd, False 'see TransPrep function for more info...
End Sub

VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2655
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4500
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picPB 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   60
      ScaleHeight     =   225
      ScaleWidth      =   4365
      TabIndex        =   0
      Top             =   2340
      Width           =   4395
      Begin VB.Shape pbShape 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFFFF&
         DrawMode        =   6  'Mask Pen Not
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   0
         Top             =   0
         Width           =   50
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Loading Yahtzee XP..."
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
         Left            =   1260
         TabIndex        =   1
         Top             =   0
         Width           =   1860
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1860
      Left            =   0
      Picture         =   "frmSplash.frx":0000
      ScaleHeight     =   1860
      ScaleWidth      =   4410
      TabIndex        =   4
      Top             =   420
      Width           =   4410
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alex Smoljanovic"
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
         Left            =   3210
         TabIndex        =   11
         Top             =   960
         Width           =   1170
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         Caption         =   "NR"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   120
         TabIndex        =   10
         Top             =   1380
         Width           =   4245
      End
      Begin VB.Label lblOrg 
         BackStyle       =   0  'Transparent
         Caption         =   "NR"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   9
         Top             =   1560
         Width           =   4260
      End
      Begin VB.Label lblRegInf 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Registered to"
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
         Left            =   120
         TabIndex        =   8
         Top             =   1200
         Width           =   930
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DigiScene Studios© 2003"
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
         Left            =   1980
         TabIndex        =   7
         Top             =   780
         Width           =   2400
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Yahtzee XP© Salex Software© 2003"
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
         Left            =   1200
         TabIndex        =   6
         Top             =   600
         Width           =   3195
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Yahtzee XP for Windows XP"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1740
         TabIndex        =   5
         Top             =   360
         Width           =   2640
      End
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   0
      Picture         =   "frmSplash.frx":391C
      ScaleHeight     =   315
      ScaleWidth      =   2835
      TabIndex        =   2
      Top             =   0
      Width           =   2835
      Begin VB.Label lblVer 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1.0.0.26"
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
         Left            =   1200
         TabIndex        =   3
         Top             =   0
         Width           =   585
      End
   End
End
Attribute VB_Name = "frmSplash"
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

Private Declare Function SetWindowPos Lib "user32" _
(ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
ByVal x As Long, ByVal y As Long, ByVal cx As Long, _
ByVal cy As Long, ByVal wFlags As Long) As Long

Const SWP_NOMOVE = 2
Const SWP_NOSIZE = 1
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2



Private Sub Form_Initialize()
Dim x&: x = InitCommonControls
'Call InitCommonControls to request that windows check the version of common controls
'specified in this applications manifest resource, if that version specified is the version which utilizes the
'Windows XP UxTheme library and its ThemeData structures then that version of common controls will be used
'This is only neccessary because the version of Common Controls which uses the UxTheme theme can't be distributed on other operating systems
',so instead a manifest resource file is included either as a compiled resource or as an external file to specify which common controls library windows is to utilize
End Sub

Private Sub Form_Load()
On Error Resume Next '...
 SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
 'update this window position, NOMOVE and NOSIZE flags are
 'specified so that the window's X and Y coordinates and rectangular
 'dimensions are not modified, only its Z-axis or Z-Order is changed...
  lblName.Caption = GetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion", "RegisteredOwner")
  lblOrg.Caption = GetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion", "RegisteredOrganization")
  'update labels lblName and lblOrg's caption properties with the registrant information of the operating system
   If lblName.Caption = vbNullString Or lblOrg.Caption = vbNullString Then
    lblName.Visible = False
    lblOrg.Visible = False
     lblRegInf.Caption = "Can't retreive registrant information"
     lblRegInf.ForeColor = vbRed
     lblRegInf.FontBold = True
   End If '...
    lblVer.Caption = App.Major & "." & App.Minor & "." & App.Revision
    'update version label's caption property
     Me.Show 'call this dialogs Show method to show this dialog
      DoEvents 'yield execution to other asynchronously processing procedures
       pbShape.Width = picPB.Width / 5 'update the shape controls rectangular dimensions
        lblstatus.Caption = "Loading Previous High Scores..." '...
          pbShape.Width = picPB.Width / 4 '...
           LoadHighScores 'see LoadHighScores for more info...
            lblstatus.Caption = "Loading YahzteeXP..." '...
             wndTransSpeed = Int(GetSetting("YAH", "Pref", "wndTransS", "20"))
              wndTrans = Int(GetSetting("YAH", "Pref", "wndTrans", "1"))
               GraphTrans = Int(GetSetting("YAH", "Pref", "GraphTrans", "1"))
               'retreieve user preferences from the registry(these registry keys are user specified(HKEY_CURRENT_USER))
                pbShape.Width = picPB.Width / 3 '...
         
        
        lblstatus.Caption = "Locating HelpFile..."
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
          'determine if this applications help file exists, and set this applications help file file path accordingly
           pbShape.Width = picPB.Width / 2 '...
            lblstatus.Caption = "Starting YahzteeXP..." '...
              pbShape.Width = picPB.Width '...
               Load frmMain 'load frmMain into memory(initialize the dialog)
                frmMain.Show 'call frmMain's show method to show the dialog
End Sub


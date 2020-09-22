VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6180
   ControlBox      =   0   'False
   HelpContextID   =   700
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   6180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1025
      Left            =   60
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   16
      Text            =   "frmAbout.frx":000C
      Top             =   2400
      Width           =   6075
   End
   Begin VB.PictureBox Picture5 
      BorderStyle     =   0  'None
      Height          =   525
      Left            =   -240
      Picture         =   "frmAbout.frx":0012
      ScaleHeight     =   525
      ScaleWidth      =   1890
      TabIndex        =   14
      Top             =   3600
      Width           =   1890
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   60
      Picture         =   "frmAbout.frx":4DFE
      ScaleHeight     =   495
      ScaleWidth      =   2175
      TabIndex        =   5
      Top             =   0
      Width           =   2175
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Freeware License"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   180
         Left            =   60
         TabIndex        =   7
         Top             =   240
         Width           =   1305
      End
      Begin VB.Label lblversion 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1.0.0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Left            =   1260
         TabIndex        =   6
         Top             =   0
         Width           =   315
      End
   End
   Begin VB.PictureBox Picture3 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   585
      Left            =   3900
      Picture         =   "frmAbout.frx":755E
      ScaleHeight     =   585
      ScaleWidth      =   2280
      TabIndex        =   4
      Top             =   0
      Width           =   2280
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   765
      Left            =   0
      ScaleHeight     =   765
      ScaleWidth      =   6165
      TabIndex        =   2
      Top             =   3450
      Width           =   6165
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmAbout.frx":D242
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   720
         Left            =   1680
         TabIndex        =   3
         Top             =   60
         Width           =   4455
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   60
      Picture         =   "frmAbout.frx":D352
      ScaleHeight     =   915
      ScaleWidth      =   6075
      TabIndex        =   15
      Top             =   480
      Width           =   6075
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "32-bit Yahtzee program which uses the Learnout and Hauspie TruVoice speech synthesization engine produced by Microsoft."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   60
         TabIndex        =   18
         Top             =   180
         Width           =   5940
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "YahtzeeXP is designed specifically for 64-bit Windows NT 5.1 (Windows XP)"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   60
         TabIndex        =   17
         Top             =   540
         Width           =   5985
      End
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
      Left            =   60
      TabIndex        =   13
      Top             =   1440
      Width           =   930
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
      Left            =   60
      TabIndex        =   12
      Top             =   1800
      Width           =   3720
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
      Left            =   60
      TabIndex        =   11
      Top             =   1620
      Width           =   3525
   End
   Begin VB.Label Label9 
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
      Left            =   3420
      TabIndex        =   10
      Top             =   1440
      Width           =   2715
   End
   Begin VB.Label Label4 
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
      Left            =   3720
      TabIndex        =   9
      Top             =   1620
      Width           =   2400
   End
   Begin VB.Label Label2 
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
      Left            =   3690
      TabIndex        =   8
      Top             =   1800
      Width           =   2430
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Salex Software"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   150
      Left            =   4500
      MouseIcon       =   "frmAbout.frx":11A1E
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   2220
      Width           =   1605
   End
   Begin VB.Label lblDateInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Compiled by Alex Smoljanovic"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   150
      Left            =   60
      TabIndex        =   0
      Top             =   2220
      Width           =   2070
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

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

Private Sub Form_Click()
 Unload Me 'unload this dialog from memory
End Sub

Private Sub Form_Load()
On Error Resume Next
'on the even of an error resume execution on the next line of this procedure
 Text1.Text = "Yahtzee XP" & vbCrLf & "Program version:" & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & vbCrLf & _
 "Yahtzee XP is the Windows XP version of the Yahtzee Dice game." & vbCrLf & vbCrLf & "To review the End User License Agreement please read the text file named 'Eula.txt' in the installation directory of this program or view it in the Help file by pressing F1." & vbCrLf & vbCrLf & "To send Salex Software© feedback on this program(Yahtzee XP) please do so via electronic mail to the address salex_software@shaw.ca, please be sure to also include the version of the program you have acquired." & vbCrLf & vbCrLf & "Thank you," & vbCrLf & "Alex Smoljanovic"
 'update Textbox Text1's text property
  
  lblversion.Caption = App.Major & "." & App.Minor & "." & App.Revision
  'update the version label's caption with the current version information
  lblName.Caption = GetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion", "RegisteredOwner")
  lblOrg.Caption = GetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion", "RegisteredOrganization")
  'update the registrar information... (These registry keys represent the user information used to register the operating system)
  
   If lblName.Caption = vbNullString Or lblOrg.Caption = vbNullString Then
    lblName.Visible = False
    lblOrg.Visible = False
     lblRegInf.Caption = "Can't retreive registrant information"
     lblRegInf.ForeColor = vbRed
     lblRegInf.FontBold = True
     '...
   End If
   
    Dim AppLoc$ 'dimensionalize(declare) AppLoc as string data type
    'note: for C++ programmers unfamiliar with the String structure,
    'the String 'data type' is a variable-length CHAR/TCHAR data type wrapper
    'similar to C++ MFC's CString class, allthough its not a class type decleration and so unlike CString class there are no methods derived from its root class
     AppLoc = IIf(Left$(App.Path, 1) = "\", App.Path & App.EXEName & ".exe", App.Path & "\" & App.EXEName & ".exe")
     'initialize AppLoc with the full path to this module
     'IIf operator conditionally returns a value based upon the statement it evaluates
     'var = IIf(Statement to evaluate, TruePart, FalsePart)
      lblDateInfo.Caption = "Compiled on " & FileDateTime(AppLoc)
      'update the DateInfo label's caption with the formatted File Data time of this module
       TransPrep Me.hwnd 'see TransPrep function for more info...
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
 Label3.FontUnderline = False
 'set the label control's FontUnderline property to false
End Sub

Private Sub Label1_Click()
 Unload Me 'unload this dialog
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 TransWin Me.hwnd, False 'see TransWin for more info...
End Sub

Private Sub Label2_Click()
 Unload Me 'unload this dialog
End Sub

Private Sub Label3_Click()
 ShellExecute Me.hwnd, "open", "mailto:salex_software@shaw.ca?subject=Yahtzee XP", vbNullString, vbNullString, vbNormal
 'use Shell's ShellExecute function to execute the specified document
 'paramater lpOperation(verb) specifies the context in which to perform the documents execution
   'other valid verbs: open, edit, print, properties, ect...
  Unload Me 'unload this dialog
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
 Label3.FontUnderline = True '...
End Sub

Private Sub Label4_Click()
 Unload Me '...
End Sub

Private Sub Label5_Click()
 Unload Me '...
End Sub

Private Sub Label6_Click()
 Unload Me '...
End Sub

Private Sub Label7_Click()
 Unload Me '...
End Sub

Private Sub Label8_Click()
 Unload Me '...
End Sub

Private Sub Label9_Click()
 Unload Me '...
End Sub

Private Sub lblDateInfo_Click()
 Unload Me '...
End Sub

Private Sub lblName_Click()
 Unload Me '...
End Sub

Private Sub lblOrg_Click()
Unload Me '...
End Sub

Private Sub lblRegInf_Click()
 Unload Me '...
End Sub

Private Sub lblversion_Click()
 Unload Me '...
End Sub

Private Sub Picture1_Click()
 Unload Me '...
End Sub

Private Sub Picture2_Click()
 Unload Me '...
End Sub

Private Sub Picture3_Click()
 Unload Me '...
End Sub

Private Sub Picture4_Click()
 Unload Me '...
End Sub

Private Sub Text1_GotFocus()
 HideCaret Text1.hwnd '...
End Sub

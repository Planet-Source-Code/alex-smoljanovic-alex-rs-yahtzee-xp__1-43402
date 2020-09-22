VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmHighScores 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "High Scores"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4275
   HelpContextID   =   200
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   4275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ListView lstScores 
      Height          =   2775
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4155
      _ExtentX        =   7329
      _ExtentY        =   4895
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "#"
         Object.Width           =   741
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Player"
         Object.Width           =   3281
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Score"
         Object.Width           =   1191
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Date"
         Object.Width           =   2090
      EndProperty
   End
End
Attribute VB_Name = "frmHighScores"
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

'Private Sub testcolumwidths()
'debug.print lstScores.ColumnHeaders(1).Width & ":" & lstScores.ColumnHeaders(2).Width & ":" & lstScores.ColumnHeaders(3).Width & ":" & lstScores.ColumnHeaders(4).Width
'End Sub

Private Sub Form_Activate()
 TransWin Me.hwnd 'see TransWin for more info...
End Sub

Private Sub Form_Load()
On Error Resume Next 'on the event of an error resume execution on the next line of this procedure
Dim itmX As ListItem, i% 'declare itmX as ListItem structure, i as integer data type
 For i = 1 To 10
 'enumerate through each of the ten high scores
  Set itmX = lstScores.ListItems.Add(, , Str$(i))
  'initialize itmX with the instance of the listitem type object returned by the ListView control's Add method
   itmX.SubItems(1) = OldHighs(i).PlayerName
    itmX.SubItems(2) = OldHighs(i).TotalScore
     itmX.SubItems(3) = OldHighs(i).Date
     'update the listview items sub-items
 Next i
  TransPrep Me.hwnd 'see TransPrep function for more info...
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 TransWin Me.hwnd, False 'see TransWin function for more info...
End Sub

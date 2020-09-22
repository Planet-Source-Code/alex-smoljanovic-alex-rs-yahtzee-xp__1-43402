VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmPlayers 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add/Remove Players"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2970
   HelpContextID   =   400
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   2970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "&Rename"
      Height          =   315
      Left            =   1980
      TabIndex        =   4
      ToolTipText     =   "Rename selected player"
      Top             =   1620
      Width           =   795
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Remove"
      Height          =   315
      Left            =   960
      TabIndex        =   2
      ToolTipText     =   "Remove selected player"
      Top             =   1620
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Add"
      Height          =   315
      Left            =   180
      TabIndex        =   1
      ToolTipText     =   "Add a player"
      Top             =   1620
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "Players"
      Height          =   1935
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   2835
      Begin MSComctlLib.ListView lstCurrent 
         Height          =   1155
         Left            =   120
         TabIndex        =   3
         Top             =   300
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   2037
         View            =   2
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
   End
End
Attribute VB_Name = "frmPlayers"
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


Private Sub Command1_Click()
If MsgBox("By adding a player you reset the game." & vbCrLf & vbCrLf & "Are you sure you wish to continue?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
'request user confirmation...
 If lstCurrent.ListItems.Count < 6 Then
 'if there are less than 6 items in the listview control lstCurrent then...
  Dim NewPlayer$: NewPlayer$ = Trim$(InputBox("Please enter the new players name:", "New Player"))
  'dimensionalize NewPlayer as string data type, initialize NewPlayer with the input returned by the InputBox function
  'note: LTrim function removes the leading and trailing white-spaces
   NewPlayer$ = Mid$(NewPlayer$, 1, 9) 'return the first 9 characters of the string
   If NewPlayer = vbNullString Then Exit Sub 'if newplayer evaluates to vbNullString(Null String[\0]) then...
    lstCurrent.ListItems.Add , , NewPlayer 'add a new listitem to the ListView control representing the new player
     frmMain.AddPlayer NewPlayer 'see frmMain's AddPlayer method for more info...
 Else
  MsgBox "There is a maximum of 6 players allowed at a time.", vbExclamation, "Can't add player"
  'Inform user that no more users can be added to the game...
 End If
End Sub

Private Sub Command2_Click()
If MsgBox("By removing a player you reset the game." & vbCrLf & vbCrLf & "Are you sure you wish to continue?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
'request user confirmation
 If lstCurrent.SelectedItem.Text = "" Then Exit Sub
 'if no item is selected then...
  If lstCurrent.ListItems.Count = 1 Then
  'if there is only one player in the game, then the player can not be removed...
   MsgBox "There must be atleast one player to play the game." & vbCrLf & vbCrLf & "Action was cancelled.", vbExclamation, "Can't remove player"
    Exit Sub 'exit this procedure
  End If
   Say "goodbye " & lstCurrent.SelectedItem.Text
    Say "player " & lstCurrent.SelectedItem.Text & " has been removed from the game."
    'see Say function for more info...
     frmMain.RemoveAndCopyPlayerInf (lstCurrent.SelectedItem.Index + 1)
     'see frmMain's RemoveAndCopyPlayerInf for more information...
      lstCurrent.ListItems.Remove (lstCurrent.SelectedItem.Index)
      'remove the listitem specified by its index property
End Sub

Private Sub Command3_Click()
 lstCurrent.SetFocus 'set focus to the control, otherwise the objects StartLabelEdit method will not succeed
  lstCurrent.StartLabelEdit 'invoke the label edit event on the currently selected listitem
End Sub

Private Sub Form_Activate()
 TransWin Me.hwnd 'see TransWin function for more info...
End Sub

Private Sub Form_Load()
On Error Resume Next 'on the event of an error resume execution on the next line of this procedure
Dim i& 'dimensionalize i as long data type
 For i = 1 To NumOfPlayers
 'enumerate through each player
  lstCurrent.ListItems.Add , , PlayerScore(i).PlayerName
  'add a listitem representing the player
 Next i
  TransPrep Me.hwnd 'see TransPrep function for more info...
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 TransWin Me.hwnd, False 'see TransWin for more info...
End Sub

Private Sub lstCurrent_AfterLabelEdit(Cancel As Integer, NewString As String)
 Say "Player " & lstCurrent.SelectedItem.Text & "'s name has been changed to " & NewString
 'see function Say for more info...
  PlayerScore(lstCurrent.SelectedItem.Index).PlayerName = NewString
  'update the player's name
  'the index of the selected item will coincide with the player element index in the PlayerScore array
   Select Case NumOfPlayers
    Case 1: 'if NumOfPlayers evaluates to 1 then...
     frmMain.lblPlayer1 = PlayerScore(1).PlayerName
     'update the label control which displays the players name on the player icon picture box
    Case 2:
     frmMain.lblPlayer1 = PlayerScore(1).PlayerName
      frmMain.lblPlayer2 = PlayerScore(2).PlayerName
    Case 3:
     frmMain.lblPlayer1 = PlayerScore(1).PlayerName
      frmMain.lblPlayer2 = PlayerScore(2).PlayerName
       frmMain.lblPlayer3 = PlayerScore(3).PlayerName
    Case 4:
     frmMain.lblPlayer1 = PlayerScore(1).PlayerName
      frmMain.lblPlayer2 = PlayerScore(2).PlayerName
       frmMain.lblPlayer3 = PlayerScore(3).PlayerName
        frmMain.lblPlayer4 = PlayerScore(4).PlayerName
    Case 5:
     frmMain.lblPlayer1 = PlayerScore(1).PlayerName
      frmMain.lblPlayer2 = PlayerScore(2).PlayerName
       frmMain.lblPlayer3 = PlayerScore(3).PlayerName
        frmMain.lblPlayer4 = PlayerScore(4).PlayerName
         frmMain.lblPlayer5 = PlayerScore(5).PlayerName
    Case 6:
     frmMain.lblPlayer1 = PlayerScore(1).PlayerName
      frmMain.lblPlayer2 = PlayerScore(2).PlayerName
       frmMain.lblPlayer3 = PlayerScore(3).PlayerName
        frmMain.lblPlayer4 = PlayerScore(4).PlayerName
         frmMain.lblPlayer5 = PlayerScore(5).PlayerName
          frmMain.lblPlayer6 = PlayerScore(6).PlayerName
   End Select
    frmMain.Frame1.Caption = PlayerScore(CurrentPlayer).PlayerName & "'s Score"
    'update the main frame's caption to contain the current player's player name
End Sub


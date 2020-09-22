VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Yahtzee XP"
   ClientHeight    =   5655
   ClientLeft      =   4155
   ClientTop       =   2925
   ClientWidth     =   6795
   HelpContextID   =   100
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   6795
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox tmplogo2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   645
      Left            =   1500
      Picture         =   "frmMain.frx":08CA
      ScaleHeight     =   39
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   152
      TabIndex        =   74
      Top             =   0
      Visible         =   0   'False
      Width           =   2340
   End
   Begin VB.PictureBox tmplogo1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   645
      Left            =   1020
      Picture         =   "frmMain.frx":65AE
      ScaleHeight     =   39
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   152
      TabIndex        =   73
      Top             =   0
      Visible         =   0   'False
      Width           =   2340
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   2040
      Top             =   2940
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSComctlLib.ImageList imglstPlayer 
      Left            =   2220
      Top             =   1980
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   126
      ImageHeight     =   29
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C292
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":FBFE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Player6Ico 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   120
      Picture         =   "frmMain.frx":1356A
      ScaleHeight     =   435
      ScaleWidth      =   1890
      TabIndex        =   62
      ToolTipText     =   "Player 6"
      Top             =   2940
      Visible         =   0   'False
      Width           =   1890
      Begin VB.Label score6 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1500
         TabIndex        =   70
         ToolTipText     =   "Player 6's score"
         Top             =   60
         Width           =   375
      End
      Begin VB.Label lblPlayer6 
         BackStyle       =   0  'Transparent
         Caption         =   "Player 6"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00D5D5D5&
         Height          =   315
         Left            =   240
         TabIndex        =   63
         ToolTipText     =   "Player 6's name"
         Top             =   60
         Width           =   1155
      End
   End
   Begin VB.PictureBox Player5Ico 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   120
      Picture         =   "frmMain.frx":16EC6
      ScaleHeight     =   435
      ScaleWidth      =   1890
      TabIndex        =   60
      ToolTipText     =   "Player 5"
      Top             =   2460
      Visible         =   0   'False
      Width           =   1890
      Begin VB.Label score5 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1500
         TabIndex        =   69
         ToolTipText     =   "Player 5's score"
         Top             =   60
         Width           =   375
      End
      Begin VB.Label lblPlayer5 
         BackStyle       =   0  'Transparent
         Caption         =   "Player 5"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00D5D5D5&
         Height          =   315
         Left            =   240
         TabIndex        =   61
         ToolTipText     =   "Player 5's name"
         Top             =   60
         Width           =   1155
      End
   End
   Begin VB.PictureBox Player4Ico 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   120
      Picture         =   "frmMain.frx":1A822
      ScaleHeight     =   435
      ScaleWidth      =   1890
      TabIndex        =   58
      ToolTipText     =   "Player 4"
      Top             =   1980
      Visible         =   0   'False
      Width           =   1890
      Begin VB.Label score4 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1500
         TabIndex        =   68
         ToolTipText     =   "Player 4's score"
         Top             =   60
         Width           =   375
      End
      Begin VB.Label lblPlayer4 
         BackStyle       =   0  'Transparent
         Caption         =   "Player 4"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00D5D5D5&
         Height          =   315
         Left            =   240
         TabIndex        =   59
         ToolTipText     =   "Player 4's name"
         Top             =   60
         Width           =   1155
      End
   End
   Begin VB.PictureBox Player2Ico 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   120
      Picture         =   "frmMain.frx":1E17E
      ScaleHeight     =   435
      ScaleWidth      =   1890
      TabIndex        =   54
      ToolTipText     =   "Player 2"
      Top             =   1020
      Visible         =   0   'False
      Width           =   1890
      Begin VB.Label score2 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1500
         TabIndex        =   66
         ToolTipText     =   "Player 2's score"
         Top             =   60
         Width           =   375
      End
      Begin VB.Label lblPlayer2 
         BackStyle       =   0  'Transparent
         Caption         =   "Player 2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00D5D5D5&
         Height          =   255
         Left            =   240
         TabIndex        =   56
         ToolTipText     =   "Player 2's name"
         Top             =   60
         Width           =   1155
      End
   End
   Begin VB.PictureBox Player3Ico 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   120
      Picture         =   "frmMain.frx":21ADA
      ScaleHeight     =   435
      ScaleWidth      =   1890
      TabIndex        =   53
      ToolTipText     =   "Player 3"
      Top             =   1500
      Visible         =   0   'False
      Width           =   1890
      Begin VB.Label score3 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1500
         TabIndex        =   67
         ToolTipText     =   "Player 3's score"
         Top             =   60
         Width           =   375
      End
      Begin VB.Label lblPlayer3 
         BackStyle       =   0  'Transparent
         Caption         =   "Player 3"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00D5D5D5&
         Height          =   315
         Left            =   240
         TabIndex        =   57
         ToolTipText     =   "Player 3's name"
         Top             =   60
         Width           =   1155
      End
   End
   Begin VB.PictureBox Player1Ico 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   120
      Picture         =   "frmMain.frx":25436
      ScaleHeight     =   435
      ScaleWidth      =   1890
      TabIndex        =   52
      ToolTipText     =   "Player 1"
      Top             =   540
      Width           =   1890
      Begin VB.Label score1 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1500
         TabIndex        =   65
         ToolTipText     =   "Player 1's score"
         Top             =   60
         Width           =   375
      End
      Begin VB.Label lblPlayer1 
         BackStyle       =   0  'Transparent
         Caption         =   "Player 1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   55
         ToolTipText     =   "Player 1's name"
         Top             =   60
         Width           =   1155
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Player 1's Score"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3450
      HelpContextID   =   600
      Left            =   2580
      TabIndex        =   21
      Top             =   240
      Width           =   4110
      Begin VB.PictureBox picLogo 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   585
         Left            =   1620
         Picture         =   "frmMain.frx":28D92
         ScaleHeight     =   39
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   152
         TabIndex        =   72
         Top             =   2820
         Width           =   2280
      End
      Begin VB.TextBox txtScoreBoxG1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   285
         HelpContextID   =   600
         Index           =   0
         Left            =   900
         Locked          =   -1  'True
         MouseIcon       =   "frmMain.frx":2EA76
         MousePointer    =   1  'Arrow
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   300
         Width           =   555
      End
      Begin VB.TextBox txtScoreBoxG1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   285
         HelpContextID   =   600
         Index           =   1
         Left            =   900
         Locked          =   -1  'True
         MouseIcon       =   "frmMain.frx":2F340
         MousePointer    =   1  'Arrow
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   660
         Width           =   555
      End
      Begin VB.TextBox txtScoreBoxG1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   285
         HelpContextID   =   600
         Index           =   2
         Left            =   900
         Locked          =   -1  'True
         MouseIcon       =   "frmMain.frx":2FC0A
         MousePointer    =   1  'Arrow
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   1020
         Width           =   555
      End
      Begin VB.TextBox txtScoreBoxG1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   285
         HelpContextID   =   600
         Index           =   3
         Left            =   900
         Locked          =   -1  'True
         MouseIcon       =   "frmMain.frx":304D4
         MousePointer    =   1  'Arrow
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   1380
         Width           =   555
      End
      Begin VB.TextBox txtScoreBoxG1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   285
         HelpContextID   =   600
         Index           =   4
         Left            =   900
         Locked          =   -1  'True
         MouseIcon       =   "frmMain.frx":30D9E
         MousePointer    =   1  'Arrow
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   1740
         Width           =   555
      End
      Begin VB.TextBox txtScoreBoxG1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   285
         HelpContextID   =   600
         Index           =   5
         Left            =   900
         Locked          =   -1  'True
         MouseIcon       =   "frmMain.frx":31668
         MousePointer    =   1  'Arrow
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   2100
         Width           =   555
      End
      Begin VB.TextBox txtTotalScore 
         Height          =   285
         HelpContextID   =   600
         Left            =   900
         MousePointer    =   1  'Arrow
         TabIndex        =   30
         Top             =   2640
         Width           =   555
      End
      Begin VB.TextBox txtBonus 
         Height          =   285
         HelpContextID   =   600
         Left            =   900
         MousePointer    =   1  'Arrow
         TabIndex        =   29
         Top             =   3000
         Width           =   555
      End
      Begin VB.TextBox txtScoreBoxG2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   285
         HelpContextID   =   600
         Index           =   0
         Left            =   3240
         Locked          =   -1  'True
         MouseIcon       =   "frmMain.frx":31F32
         MousePointer    =   1  'Arrow
         TabIndex        =   28
         Top             =   300
         Width           =   795
      End
      Begin VB.TextBox txtScoreBoxG2 
         Alignment       =   2  'Center
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   285
         HelpContextID   =   600
         Index           =   1
         Left            =   3240
         Locked          =   -1  'True
         MouseIcon       =   "frmMain.frx":327FC
         MousePointer    =   1  'Arrow
         TabIndex        =   27
         Top             =   660
         Width           =   795
      End
      Begin VB.TextBox txtScoreBoxG2 
         Alignment       =   2  'Center
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   285
         HelpContextID   =   600
         Index           =   2
         Left            =   3240
         Locked          =   -1  'True
         MouseIcon       =   "frmMain.frx":330C6
         MousePointer    =   1  'Arrow
         TabIndex        =   26
         Top             =   1020
         Width           =   795
      End
      Begin VB.TextBox txtScoreBoxG2 
         Alignment       =   2  'Center
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   285
         HelpContextID   =   600
         Index           =   3
         Left            =   3240
         Locked          =   -1  'True
         MouseIcon       =   "frmMain.frx":33990
         MousePointer    =   1  'Arrow
         TabIndex        =   25
         Top             =   1380
         Width           =   795
      End
      Begin VB.TextBox txtScoreBoxG2 
         Alignment       =   2  'Center
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   285
         HelpContextID   =   600
         Index           =   4
         Left            =   3240
         Locked          =   -1  'True
         MouseIcon       =   "frmMain.frx":3425A
         MousePointer    =   1  'Arrow
         TabIndex        =   24
         Top             =   1740
         Width           =   795
      End
      Begin VB.TextBox txtScoreBoxG2 
         Alignment       =   2  'Center
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   285
         HelpContextID   =   600
         Index           =   5
         Left            =   3240
         Locked          =   -1  'True
         MouseIcon       =   "frmMain.frx":34B24
         MousePointer    =   1  'Arrow
         TabIndex        =   23
         Top             =   2100
         Width           =   795
      End
      Begin VB.TextBox txtScoreBoxG2 
         Alignment       =   2  'Center
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   285
         HelpContextID   =   600
         Index           =   6
         Left            =   3240
         Locked          =   -1  'True
         MouseIcon       =   "frmMain.frx":353EE
         MousePointer    =   1  'Arrow
         TabIndex        =   22
         Top             =   2460
         Width           =   795
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00D5D5D5&
         X1              =   1500
         X2              =   4440
         Y1              =   2775
         Y2              =   2775
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "One's:"
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
         Left            =   60
         TabIndex        =   51
         Top             =   360
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Two's:"
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
         Left            =   60
         TabIndex        =   50
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Three's:"
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
         Left            =   60
         TabIndex        =   49
         Top             =   1080
         Width           =   795
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Four's:"
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
         Left            =   60
         TabIndex        =   48
         Top             =   1440
         Width           =   675
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Five's:"
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
         Left            =   60
         TabIndex        =   47
         Top             =   1800
         Width           =   645
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Six's:"
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
         Left            =   60
         TabIndex        =   46
         Top             =   2160
         Width           =   525
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00D5D5D5&
         X1              =   1500
         X2              =   1500
         Y1              =   300
         Y2              =   3300
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00D5D5D5&
         X1              =   60
         X2              =   4140
         Y1              =   2415
         Y2              =   2415
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
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
         Left            =   60
         TabIndex        =   45
         Top             =   2700
         Width           =   555
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bonus:"
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
         Left            =   60
         TabIndex        =   44
         Top             =   3060
         Width           =   645
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Three of a Kind:"
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
         Left            =   1620
         TabIndex        =   43
         Top             =   360
         Width           =   1545
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Four of a Kind:"
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
         Left            =   1620
         TabIndex        =   42
         Top             =   720
         Width           =   1425
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Full House:"
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
         Left            =   1620
         TabIndex        =   41
         Top             =   1080
         Width           =   1080
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Low Straight:"
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
         Left            =   1620
         TabIndex        =   40
         Top             =   1440
         Width           =   1275
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "High Straight:"
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
         Left            =   1620
         TabIndex        =   39
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Yahtzee:"
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
         Left            =   1620
         TabIndex        =   38
         Top             =   2160
         Width           =   840
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Chance:"
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
         Left            =   1620
         TabIndex        =   37
         Top             =   2520
         Width           =   765
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00D5D5D5&
         X1              =   60
         X2              =   4080
         Y1              =   615
         Y2              =   615
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00D5D5D5&
         X1              =   60
         X2              =   4080
         Y1              =   975
         Y2              =   975
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00D5D5D5&
         X1              =   60
         X2              =   4080
         Y1              =   1335
         Y2              =   1335
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00D5D5D5&
         X1              =   60
         X2              =   4080
         Y1              =   1695
         Y2              =   1695
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00D5D5D5&
         X1              =   60
         X2              =   4080
         Y1              =   2055
         Y2              =   2055
      End
   End
   Begin VB.PictureBox picDice6 
      Height          =   570
      Left            =   1500
      ScaleHeight     =   510
      ScaleWidth      =   585
      TabIndex        =   20
      Top             =   3300
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.PictureBox picDice5 
      Height          =   570
      Left            =   1980
      ScaleHeight     =   510
      ScaleWidth      =   645
      TabIndex        =   19
      Top             =   780
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.PictureBox picDice4 
      Height          =   570
      Left            =   1920
      ScaleHeight     =   510
      ScaleWidth      =   645
      TabIndex        =   18
      Top             =   3240
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.PictureBox picDice3 
      Height          =   570
      Left            =   2040
      ScaleHeight     =   510
      ScaleWidth      =   645
      TabIndex        =   17
      Top             =   1140
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.PictureBox picDice2 
      Height          =   570
      Left            =   2040
      ScaleHeight     =   510
      ScaleWidth      =   645
      TabIndex        =   16
      Top             =   1860
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.PictureBox picDice1 
      Height          =   570
      Left            =   -645
      ScaleHeight     =   510
      ScaleWidth      =   645
      TabIndex        =   15
      Top             =   2595
      Visible         =   0   'False
      Width           =   705
   End
   Begin MSComctlLib.ImageList imglstclearstand 
      Left            =   2100
      Top             =   2460
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   279
      ImageHeight     =   19
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":35CB8
            Key             =   "nutral"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3AFE0
            Key             =   "clear"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":40308
            Key             =   "stand"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imglstHold 
      Left            =   1020
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   50
      ImageHeight     =   27
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":45630
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4668A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imglstDice 
      Left            =   1860
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   50
      ImageHeight     =   50
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   24
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":476E4
            Key             =   "b1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":494E6
            Key             =   "b2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4B220
            Key             =   "b3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4D022
            Key             =   "b4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4ED5C
            Key             =   "b5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":50AEE
            Key             =   "b6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":528F0
            Key             =   "gld1"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":546F2
            Key             =   "gld2"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5642C
            Key             =   "gld3"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":582F6
            Key             =   "gld4"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5A030
            Key             =   "gld5"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5BDC2
            Key             =   "gld6"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5DBC4
            Key             =   "g1"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5F9C6
            Key             =   "g2"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":61700
            Key             =   "g3"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":63502
            Key             =   "g4"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6523C
            Key             =   "g5"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":66F3E
            Key             =   "g6"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":68D40
            Key             =   "r1"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6AB42
            Key             =   "r2"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6C87C
            Key             =   "r3"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6E67E
            Key             =   "r4"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":703B8
            Key             =   "r5"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7214A
            Key             =   "r6"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picClearStand 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   165
      Picture         =   "frmMain.frx":73F4C
      ScaleHeight     =   285
      ScaleWidth      =   4185
      TabIndex        =   12
      Top             =   5280
      Width           =   4185
      Begin VB.Label lblstatus 
         BackStyle       =   0  'Transparent
         Caption         =   "Please Roll Dice!!"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1440
         TabIndex        =   71
         Top             =   60
         Width           =   1995
      End
      Begin VB.Label btnclearholds 
         BackStyle       =   0  'Transparent
         Height          =   210
         Left            =   30
         MouseIcon       =   "frmMain.frx":79264
         MousePointer    =   99  'Custom
         TabIndex        =   14
         Top             =   30
         Width           =   1215
      End
      Begin VB.Label btnstand 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   3495
         MouseIcon       =   "frmMain.frx":7956E
         MousePointer    =   99  'Custom
         TabIndex        =   13
         Top             =   0
         Width           =   690
      End
   End
   Begin VB.PictureBox hold5 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   3525
      MouseIcon       =   "frmMain.frx":79878
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":79B82
      ScaleHeight     =   405
      ScaleWidth      =   750
      TabIndex        =   11
      Top             =   4740
      Width           =   750
   End
   Begin VB.PictureBox hold4 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   2670
      MouseIcon       =   "frmMain.frx":7ABCC
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":7AED6
      ScaleHeight     =   405
      ScaleWidth      =   750
      TabIndex        =   10
      Top             =   4740
      Width           =   750
   End
   Begin VB.PictureBox hold3 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   1845
      MouseIcon       =   "frmMain.frx":7BF20
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":7C22A
      ScaleHeight     =   405
      ScaleWidth      =   750
      TabIndex        =   9
      Top             =   4740
      Width           =   750
   End
   Begin VB.PictureBox hold2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   1020
      MouseIcon       =   "frmMain.frx":7D274
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":7D57E
      ScaleHeight     =   405
      ScaleWidth      =   750
      TabIndex        =   8
      Top             =   4740
      Width           =   750
   End
   Begin VB.PictureBox hold1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   165
      MouseIcon       =   "frmMain.frx":7E5C8
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":7E8D2
      ScaleHeight     =   405
      ScaleWidth      =   750
      TabIndex        =   7
      Top             =   4740
      Width           =   750
   End
   Begin VB.PictureBox btnRollDice 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1665
      Left            =   4740
      MouseIcon       =   "frmMain.frx":7F91C
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":7FC26
      ScaleHeight     =   111
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   125
      TabIndex        =   6
      Top             =   3750
      Width           =   1875
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1485
      Left            =   75
      ScaleHeight     =   1485
      ScaleWidth      =   4440
      TabIndex        =   0
      Top             =   3750
      Width           =   4440
      Begin VB.PictureBox dice5 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   870
         Left            =   3450
         MouseIcon       =   "frmMain.frx":8D536
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":8D840
         ScaleHeight     =   870
         ScaleWidth      =   750
         TabIndex        =   5
         Top             =   225
         Width           =   750
      End
      Begin VB.PictureBox dice4 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   870
         Left            =   2625
         MouseIcon       =   "frmMain.frx":8F5C2
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":8F8CC
         ScaleHeight     =   870
         ScaleWidth      =   750
         TabIndex        =   4
         Top             =   225
         Width           =   750
      End
      Begin VB.PictureBox dice3 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   870
         Left            =   1800
         MouseIcon       =   "frmMain.frx":916BE
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":919C8
         ScaleHeight     =   870
         ScaleWidth      =   750
         TabIndex        =   3
         Top             =   225
         Width           =   750
      End
      Begin VB.PictureBox dice2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   870
         Left            =   960
         MouseIcon       =   "frmMain.frx":936F2
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":939FC
         ScaleHeight     =   870
         ScaleWidth      =   750
         TabIndex        =   2
         Top             =   210
         Width           =   750
      End
      Begin VB.PictureBox dice1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   870
         Left            =   120
         MouseIcon       =   "frmMain.frx":95726
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":95A30
         ScaleHeight     =   870
         ScaleWidth      =   750
         TabIndex        =   1
         Top             =   210
         Width           =   750
      End
      Begin VB.Shape box5 
         BackColor       =   &H000080FF&
         BorderColor     =   &H00D5D5D5&
         BorderWidth     =   2
         Height          =   1275
         Left            =   3420
         Top             =   165
         Width           =   855
      End
      Begin VB.Shape box4 
         BackColor       =   &H000080FF&
         BorderColor     =   &H00D5D5D5&
         BorderWidth     =   2
         Height          =   1275
         Left            =   2565
         Top             =   165
         Width           =   855
      End
      Begin VB.Shape box3 
         BackColor       =   &H000080FF&
         BorderColor     =   &H00D5D5D5&
         BorderWidth     =   2
         Height          =   1275
         Left            =   1740
         Top             =   165
         Width           =   855
      End
      Begin VB.Shape box2 
         BackColor       =   &H000080FF&
         BorderColor     =   &H00D5D5D5&
         BorderWidth     =   2
         Height          =   1275
         Left            =   915
         Top             =   165
         Width           =   855
      End
      Begin VB.Shape box1 
         BackColor       =   &H000080FF&
         BorderColor     =   &H00D5D5D5&
         BorderWidth     =   2
         Height          =   1275
         Left            =   60
         Top             =   165
         Width           =   885
      End
   End
   Begin MSComctlLib.ImageList imglstRollDice 
      Left            =   0
      Top             =   3060
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   125
      ImageHeight     =   111
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":97822
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A5142
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B2A62
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00D5D5D5&
      X1              =   900
      X2              =   1920
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "Players"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   64
      Top             =   240
      Width           =   1395
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNewGame 
         Caption         =   "&New Game"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuLoadGame 
         Caption         =   "&Load Game"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSaveGame 
         Caption         =   "&Save Game"
         Shortcut        =   ^S
      End
      Begin VB.Menu l2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHighScores 
         Caption         =   "&High Scores"
         Shortcut        =   ^H
      End
      Begin VB.Menu l1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuSound 
         Caption         =   "&Sound"
      End
      Begin VB.Menu mnuPrefs 
         Caption         =   "&Preferences"
      End
   End
   Begin VB.Menu mnuDice 
      Caption         =   "&Dice"
      Begin VB.Menu mnuBlue 
         Caption         =   "&Blue"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuRed 
         Caption         =   "&Red"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuGreen 
         Caption         =   "&Green"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuGold 
         Caption         =   "G&old"
      End
   End
   Begin VB.Menu mnuPLayers 
      Caption         =   "&Players"
      Begin VB.Menu mnuAddPLayer 
         Caption         =   "&Add/Remove Players"
         Shortcut        =   +{INSERT}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuContents 
         Caption         =   "&Contents"
      End
      Begin VB.Menu l3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
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

Private Declare Function UpdateColors Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)


Dim BF As BLENDFUNCTION, lngBF&

Private Dice1Stat As DiceStatistics
Private Dice2Stat As DiceStatistics
Private Dice3Stat As DiceStatistics
Private Dice4Stat As DiceStatistics
Private Dice5Stat As DiceStatistics

Public ChangePlayerflg%, dss As DirectSS
Attribute ChangePlayerflg.VB_VarHelpID = -1
Dim LogoFlag%

Private Function DBlend(Dest As PictureBox, Source As PictureBox)
'Main Dialog Logo Animated Blend
On Error Resume Next 'on the event of an error resume execution on the next line of this procedure
If GraphTrans = False Then Exit Function 'if global variable GraphTrans evaluates to false then exit this procedure
Dim i As Byte 'dimensionalize i as byte data type
 'initialize variable BF(BlendFunction Struct.)'s members
 BF.BlendOp = AC_SRC_OVER 'this member specifies the source operations,
 'currently the only flag supported is AC_SRC_OVER
  BF.BlendFlags = 0 'this member must be initialized to zero
   BF.AlphaFormat = 0 'this member controls the way the source and destination bitmaps are interpreted
    For i = 0 To 255 Step 1
    'for next loop statement, initialize i to 0, loop until i evaluates to 255 incrementing i by one each iteration
     DoEvents 'yeild execution to other asynchronously processing procedures
      Dest.Cls 'call the PictureBox's controls Cls method to delete its currently selected GDI object
       BF.SourceConstantAlpha = i 'initialize this member(transparency value) with the value of i(BYTE data type range 0-255), where 255 equals opaque
        RtlMoveMemory lngBF, BF, 4 'RtlMoveMemory or CopyMemory moves data in the specified memory block to the specified destination memory block
         AlphaBlend Dest.hdc, 0, 0, Source.ScaleWidth, Source.ScaleHeight, Source.hdc, 0, 0, Source.ScaleWidth, Source.ScaleHeight, lngBF
         'Call AlphaBlend to blend the source Device Context's GDI object to the destination Device Context
    Next i
     Dest.Picture = Source.Picture
     'Since the Picture Box control Dest's AutoRedraw(Persistant Bitmap)'s value is true, we will update it's bitmap...
End Function

Private Sub CheckForEndOfGame()
Dim i%, j% 'dimensionalize or declare variable i and j to be of integer data type
  For i = 1 To NumOfPlayers
  'for next loop; i is initialized to 1, looping until i evaluates to the value of NumOfPlayers, since the step keyword is omitting, "step 1" is assumed incrementing i by one each iteration
   If PlayerScore(i).uChance = False Then GoTo NEOG
   'if element i in array PlayerScore evaluates to false then jump to label NEOG(Not End Of Game), since this item on the scorecard hasn't yet been filled...
    If PlayerScore(i).uFives = False Then GoTo NEOG '...
     If PlayerScore(i).uFourKind = False Then GoTo NEOG '...
      If PlayerScore(i).uFours = False Then GoTo NEOG
       If PlayerScore(i).uFullHouse = False Then GoTo NEOG
        If PlayerScore(i).uHightStraight = False Then GoTo NEOG
         If PlayerScore(i).uLowStraight = False Then GoTo NEOG
          If PlayerScore(i).uOnes = False Then GoTo NEOG
           If PlayerScore(i).uSixes = False Then GoTo NEOG
            If PlayerScore(i).uThreeKind = False Then GoTo NEOG
             If PlayerScore(i).uThrees = False Then GoTo NEOG
              If PlayerScore(i).uTwos = False Then GoTo NEOG
               If PlayerScore(i).uYahtzee = False Then GoTo NEOG
 Next i 'evaluate loops conditions, increment i, loop
  For i = 1 To NumOfPlayers '...
   ScoresToAdd(i).PlayerName = PlayerScore(i).PlayerName
   'copy element i's PlayerName member value, to element i in ScoresToAdd array
    ScoresToAdd(i).TotalScore = PlayerScore(i).Total
    '...
  Next i
   
   If NumOfPlayers > 1 Then
   'if NumOfPlayers is greater than 1 then...
    Dim Cflg&, winner% 'dimensionalize Cflg as long data type, winner as integer data type
     For i = 1 To NumOfPlayers
     'for next loop; initialize i to 1; loop until i evaluates to constant NumOfPlayers
     'Loop Prototype
     '
     '        |-High Score 1
     '        |-High Score 2
     'Player1 |-High Score 3
     '        |-High Score 4
     '        |-High Score 5
     '
     '
     '        |-High Score 1
     '        |-High Score 2
     'Player2 |-High Score 3
     '        |-High Score 4
     '        |-High Score 5
     '
      For j = 1 To NumOfPlayers
       If PlayerScore(i).Total > PlayerScore(j).Total Then
        Cflg& = Cflg& + 1
       Else
        Cflg& = 0
       End If
        If Cflg& = (NumOfPlayers - 1) Then winner = i: GoTo dwin
        'determine the insertion position for the new score on the high score list
      Next j
     Next i
dwin:
     Say PlayerScore(winner).PlayerName 'See function Say for more info...
      Say " wins with a score of " & Str$(PlayerScore(winner).Total)
       Say "Game Over"
    Else
     Say "You finish the game with a score of " & PlayerScore(CurrentPlayer).Total
      Say "Game Over"
    End If
    ResetPlayers 'see ResetPlayers for more info...
     AddHighScores 'see AddHighScores for more info...
      SaveHighScores 'see SaveHighScores for more info...
       Load frmHighScores 'load the dialog into memory
        frmHighScores.Show vbModal, Me 'Show frmHighScores as a modal dialog
NEOG: Exit Sub 'exit this procedure
End Sub

Private Sub btnclearholds_Click()
If OldDice = True Then Exit Sub
'if OldDice evaluates to True as it will when the user has rolled the dice the maximum ammount of times allow(Frozen Dice)
 Dice1Stat.Holding = False 'set structure Dice1Stat's Holding member to false
  hold1.Picture = imglstHold.ListImages.Item(1).Picture
  'update picturebox hold1's picture property with the gdi object returned from the image list imglstHold's listImages collection
   box1.BorderColor = &HD5D5D5 'update the Shape controls BorderColor property (nutral state)
   
    Dice2Stat.Holding = False '...
     hold2.Picture = imglstHold.ListImages.Item(1).Picture '...
      box2.BorderColor = &HD5D5D5 '...
        
        Dice3Stat.Holding = False
         hold3.Picture = imglstHold.ListImages.Item(1).Picture
          box3.BorderColor = &HD5D5D5
          
            Dice4Stat.Holding = False
             hold4.Picture = imglstHold.ListImages.Item(1).Picture
               box4.BorderColor = &HD5D5D5
               
                Dice5Stat.Holding = False
                  hold5.Picture = imglstHold.ListImages.Item(1).Picture
                    box5.BorderColor = &HD5D5D5
                     
End Sub

Private Sub btnclearholds_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button <> 1 Then Exit Sub 'if Button is unequal to 1 then exit this procedure
 picClearStand.Picture = imglstclearstand.ListImages.Item(2).Picture
 'update picture(HOT state)
End Sub

Private Sub btnclearholds_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button <> 1 Then Exit Sub 'if Button is unequal to 1 then exit this procedure
 picClearStand.Picture = imglstclearstand.ListImages.Item(1).Picture
 'update picture(nutral state)
End Sub

Private Sub btnRollDice_Click()
Dim i% 'dimensionalize i as integer data type
If RollDiceBtnDis = True Then Exit Sub
'if RollDiceBtnDis(Roll Dice Button Disabled) evaluates to true then exit this procedure
 If OldDice = True Then 'if OldDice evaluates to true(Frozen Dice) then
  hold1.Picture = imglstHold.ListImages.Item(1).Picture 'update picture(nutral state)
   box1.BorderColor = &HD5D5D5 'update shape control box1's bordercolor (nutral state)
   
    Dice2Stat.Holding = False '...
     hold2.Picture = imglstHold.ListImages.Item(1).Picture '...
      box2.BorderColor = &HD5D5D5 '...
        
        Dice3Stat.Holding = False
         hold3.Picture = imglstHold.ListImages.Item(1).Picture
          box3.BorderColor = &HD5D5D5
          
            Dice4Stat.Holding = False
             hold4.Picture = imglstHold.ListImages.Item(1).Picture
               box4.BorderColor = &HD5D5D5
               
                Dice5Stat.Holding = False
                  hold5.Picture = imglstHold.ListImages.Item(1).Picture
                    box5.BorderColor = &HD5D5D5
                     If ChangePlayerflg = 2 Then ChangeCurrentPlayer: ChangePlayerflg = 0
 End If
  RollDie 'see RollDie for more info...
   If PlaySounds = True Then 'if PlaySounds evaluates to true(set by Options) then play the appropriate sound
    If Not (Dir(App.Path & "\snd\dice.snd") = "") Then
    'if function Dir(%FilePath%) returns a null string then the specified file doesn't exist
    'if the "dice.snd" file exists, then...
     PlaySound App.Path & "\snd\dice.snd", frmMain.hwnd, SND_FILENAME Or SND_ASYNC
     'Call PlaySound to play the sound
     'flag SND_FILENAME specifies that the first paramater is specifying a File Path rather than a memory address
     'flag SND_ASYNC is passed so the function performs asynchronously
    End If
   End If
    
    If OldDice = True Then OldDice = False:   lblstatus.Caption = "First of three draws": Exit Sub
    'refresh the 'Frozen Dice' status
    'update label lblstatus's caption property
     CurrentDraw = CurrentDraw + 1 'increment CurrentDraw
     
      Select Case CurrentDraw
      'Select Case Statement
       Case 1: 'if CurrentDraw evaluates to 1 then...
        lblstatus.Caption = "Second of three draws"
        'update label lblstatus's caption property
       Case 2: '...
        lblstatus.Caption = "Choose score board" '...
       End Select 'escape the Select Case statement
     
      If CurrentDraw >= 2 Then 'if CurrentDraw is greater than or equal to 2(User is on second or third draw)
       EnableRollDiceBtn False  'Disable the RollDice button(Updates picture, and the Enabled property of the control)
        CurrentDraw = 0 'refresh CurrentDraw
         OldDice = True 'initialize OldDice to true(Frozen Dice)
          ChoosingScoreBox = True 'update ChoosingScoreBox flag
          
           For i = 0 To txtScoreBoxG1.UBound
            txtScoreBoxG1(i).MousePointer = 99
            'Change the MousePointer property of each textbox control in the control array to Custom(99)[Use MouseIcon property]
           Next i
            For i = 0 To txtScoreBoxG2.UBound
             txtScoreBoxG2(i).MousePointer = 99 '...
            Next i

            CheckForYahtzee 'see CheckForYahtzee for more info...
             ChangePlayerflg = 1 'update flag ChangePlayerflg to true
              Dice1Stat.Holding = False
                
                
  hold1.Picture = imglstHold.ListImages.Item(1).Picture 'update picture(nutral state)
   box1.BorderColor = &HD5D5D5 'update shape's border color(nutral state)
   
    Dice2Stat.Holding = False '...
     hold2.Picture = imglstHold.ListImages.Item(1).Picture '...
      box2.BorderColor = &HD5D5D5 '...
        
        Dice3Stat.Holding = False
         hold3.Picture = imglstHold.ListImages.Item(1).Picture
          box3.BorderColor = &HD5D5D5
          
            Dice4Stat.Holding = False
             hold4.Picture = imglstHold.ListImages.Item(1).Picture
               box4.BorderColor = &HD5D5D5
               
                Dice5Stat.Holding = False
                  hold5.Picture = imglstHold.ListImages.Item(1).Picture
                    box5.BorderColor = &HD5D5D5
      End If
       
End Sub

Private Sub EnableRollDiceBtn(Enable As Boolean)
 RollDiceBtnDis = Not (Enable) 'perform a logical not operation to invert the boolean value of Enable
  If Enable = False Then 'if Enable evaluates to False then...
   btnRollDice.Picture = imglstRollDice.ListImages.Item(3).Picture 'disabled
    btnRollDice.MousePointer = 1 'default cursor icon
  Else
   btnRollDice.Picture = imglstRollDice.ListImages.Item(1).Picture 'enabled
    btnRollDice.MousePointer = 99 'custom cursor icon(MouseIcon property is used[Hand])
     btnclearholds_Click 'see sub routine btnclearholds_click for more info...
  End If
End Sub

Private Sub btnRollDice_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
 If Button <> 1 Or RollDiceBtnDis = True Then Exit Sub
 'if button is unequal to 1, or the RollDiceButton is disabled then exit procedure
  btnRollDice.Picture = imglstRollDice.ListImages.Item(2).Picture 'HOT state
End Sub

Private Sub btnRollDice_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
 If Not (Button = 1) Or RollDiceBtnDis = True Then Exit Sub
 'if button is unequal to 1, or the RollDiceButton is disabled then exit procedure
  btnRollDice.Picture = imglstRollDice.ListImages.Item(1).Picture 'Nutral State
End Sub

Private Sub btnstand_Click()
Dim i% 'dimensionalize i as integer
If OldDice = True Then Exit Sub 'if OldDice(Frozen Dice) evaluates to true then exit this procedure
 hold1.Picture = imglstHold.ListImages.Item(1).Picture 'update picture (Nutral[Non Holding])
   box1.BorderColor = &HD5D5D5 'Nutral state
   
    Dice2Stat.Holding = False '...
     hold2.Picture = imglstHold.ListImages.Item(1).Picture '...
      box2.BorderColor = &HD5D5D5 '...
        
        Dice3Stat.Holding = False
         hold3.Picture = imglstHold.ListImages.Item(1).Picture
          box3.BorderColor = &HD5D5D5
          
            Dice4Stat.Holding = False
             hold4.Picture = imglstHold.ListImages.Item(1).Picture
               box4.BorderColor = &HD5D5D5
               
                Dice5Stat.Holding = False
                  hold5.Picture = imglstHold.ListImages.Item(1).Picture
                    box5.BorderColor = &HD5D5D5

                     Say "standing, choose a score box"
                     'see function Say for more info...
                      CurrentDraw = 0 'refresh users Current Draw flag
                       OldDice = True 'frozen dice
                        ChoosingScoreBox = True
                         For i = 0 To txtScoreBoxG1.UBound
                          txtScoreBoxG1(i).MousePointer = 99
                          'set each control in the control array to 99(Custom [Uses MouseIcon properties Cursor handle])
                         Next i
                          For i = 0 To txtScoreBoxG2.UBound
                           txtScoreBoxG2(i).MousePointer = 99 '...
                          Next i
                           EnableRollDiceBtn False 'see EnableRollDiceBtn for more info...
                            CheckForYahtzee 'see CheckForYahtzee for more info...
                             ChangePlayerflg = 1 'update flag
                              lblstatus.Caption = "Choose score box"
                              'update label controls caption property
End Sub

Private Sub btnstand_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button <> 1 Then Exit Sub 'if the mouse button which invoked the mouse related window message is unequal to 1 then exit this procedure
 picClearStand.Picture = imglstclearstand.ListImages.Item(3).Picture
 'update picture to HOT state
End Sub

Private Sub btnstand_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button <> 1 Then Exit Sub '...
 picClearStand.Picture = imglstclearstand.ListImages.Item(1).Picture 'nutral state
End Sub

Private Sub dice1_Click()
 If OldDice = True Then Exit Sub 'If OldDice(Frozen Dice) evaluates to true then exit this procedure
  HoldDice 1 'see HoldDice function for more info... (inverts the specified dice's holding status)
   If Dice1Stat.Holding = True Then 'if Dice 1's status is holding then...
    If PlaySounds = True Then 'if PlaySounds flag(set by options) evaluates to true then...
     If Not (Dir(App.Path & "\snd\hold.wav") = "") Then
     'if the sound file exists then...
      PlaySound App.Path & "\snd\hold.wav", frmMain.hwnd, SND_FILENAME Or SND_ASYNC
      'call PlaySound to play the specified sound file asynchronously
     End If
    End If
   Else
   'if dice 1's holding status is nutral then...
    If PlaySounds = True Then 'if PlaySounds flag evaluates to true then...
     If Not (Dir(App.Path & "\snd\hold1.wav") = "") Then
     'if the specified file exists then...
      PlaySound App.Path & "\snd\hold1.wav", frmMain.hwnd, SND_FILENAME Or SND_ASYNC
      'call playsound to play the specified file...
     End If
    End If
   End If
End Sub

Private Sub dice2_Click()
'for more information see sub routine dice1_Click
 If OldDice = True Then Exit Sub
  HoldDice 2
   If Dice2Stat.Holding = True Then
    If PlaySounds = True Then
     If Not (Dir(App.Path & "\snd\hold.wav") = "") Then
      PlaySound App.Path & "\snd\hold.wav", frmMain.hwnd, SND_FILENAME Or SND_ASYNC
     End If
    End If
   Else
    If PlaySounds = True Then
     If Not (Dir(App.Path & "\snd\hold1.wav") = "") Then
      PlaySound App.Path & "\snd\hold1.wav", frmMain.hwnd, SND_FILENAME Or SND_ASYNC
     End If
    End If
   End If
End Sub

Private Sub dice3_Click()
'for more information see sub routine dice1_Click
 If OldDice = True Then Exit Sub
  HoldDice 3
   If Dice3Stat.Holding = True Then
    If PlaySounds = True Then
     If Not (Dir(App.Path & "\snd\hold.wav") = "") Then
      PlaySound App.Path & "\snd\hold.wav", frmMain.hwnd, SND_FILENAME Or SND_ASYNC
     End If
    End If
   Else
    If PlaySounds = True Then
     If Not (Dir(App.Path & "\snd\hold1.wav") = "") Then
      PlaySound App.Path & "\snd\hold1.wav", frmMain.hwnd, SND_FILENAME Or SND_ASYNC
     End If
    End If
   End If
End Sub

Private Sub dice4_Click()
'for more information see sub routine dice1_Click
 If OldDice = True Then Exit Sub
  HoldDice 4
   If Dice4Stat.Holding = True Then
    If PlaySounds = True Then
     If Not (Dir(App.Path & "\snd\hold.wav") = "") Then
      PlaySound App.Path & "\snd\hold.wav", frmMain.hwnd, SND_FILENAME Or SND_ASYNC
     End If
    End If
   Else
    If PlaySounds = True Then
     If Not (Dir(App.Path & "\snd\hold1.wav") = "") Then
      PlaySound App.Path & "\snd\hold1.wav", frmMain.hwnd, SND_FILENAME Or SND_ASYNC
     End If
    End If
   End If
End Sub

Private Sub dice5_Click()
'for more information see sub routine dice1_Click
 If OldDice = True Then Exit Sub
  HoldDice 5
   If Dice5Stat.Holding = True Then
    If PlaySounds = True Then
     If Not (Dir(App.Path & "\snd\hold.wav") = "") Then
      PlaySound App.Path & "\snd\hold.wav", frmMain.hwnd, SND_FILENAME Or SND_ASYNC
     End If
    End If
   Else
    If PlaySounds = True Then
     If Not (Dir(App.Path & "\snd\hold1.wav") = "") Then
      PlaySound App.Path & "\snd\hold1.wav", frmMain.hwnd, SND_FILENAME Or SND_ASYNC
     End If
    End If
   End If
End Sub

Private Sub Form_Load()
On Error Resume Next 'on the event of an error resume execution on the next line of this procedure
 Unload frmSplash 'unload dialog frmSplash from memory
  Dim PrefDieCol$ 'dimensionalize PrefDieCol(Preferred Dice Color) as string data type
   Set dss = New DirectSS 'initialize object dss with a new class instance of class DirectSS(Direct Speech Synthesis[Learnout & Haupsie Speech Synthesization Engine])
    InitDefPlayerInfo 'see InitDefPlayerInfo for more info...
     PrefDieCol = GetSetting("YAH", "PREF", "DIECOL", "b")
     'initialize PrefDieCol with the specified registry key value, if the registry key does not exist "b" (blue) is returned
      Select Case PrefDieCol
      'select case statement
       Case "b": 'if PrefDieCol evaluates to "b" then...
        CurDiceCol.DiceColor = Blue 'update flag
         mnuBlue.Checked = True 'set menu mnuBlue's checked(Checkmark Bitmap) property to true
       Case "r": '...
        CurDiceCol.DiceColor = Red '...
         mnuRed.Checked = True '...
       Case "g":
        CurDiceCol.DiceColor = Green
        mnuGreen.Checked = True
       Case "gld":
        CurDiceCol.DiceColor = Gold
         mnuGold.Checked = True
      End Select
       preloadDice CurDiceCol.DiceColor  'see preloadDic function for more info(Updates all the PictureBox's representing Dice with the appropriate dice picture)
        OldDice = True 'Frozen Dice
         RollDie 'see RollDie for more info...
           If Not (UserName = "") Then 'See function UserName for more info...
            Say "Welcome to Yahtzee X P " & UserName & ",,,,, roll the dice to start a new game"
            'see Say function for more info...
           Else
            Say "Welcome to Yahtzee X P,,,,, roll the dice to start a new game"
           End If
            Dim sndflg%: sndflg = Val(GetSetting("YAH", "SNDOPT", "PSflag", "1"))
            'dimensionalize sndflg as integer data type
            'initialize sndflg with the numerical value of the value of the specified registry key
            'note: function Val seeks the first substring of a string which represents numerical characters, it considers the first non-numerical character to be the white-space delimeter
            'similar to Shlwapi's StrToInt function
             PlaySounds = CBool(sndflg) 'initialize PlaySounds flag with the return of CBool(Boolean Conversion)
              If Len(App.HelpFile) = 0 Then mnuContents.Enabled = False
              'if the length of object App's HelpFile property is 0 then disable the Help>>Contents menu item
              'When this application initializes it tests the existance of the help file, if then updates that property accordingly, so if its length is zero then the help file wasn't found
               If Command <> "" Then
               'if the return of function Command(function returns the parsed command line)
               'is not equal to "" then...
                If Dir(Command) = "" Then Exit Sub
                'if the command line isn't specifying a file or it is specifying a file which doesnt exist then exit this procedure
                 LoadGame False, Command
                 'see LoadGame function for more info...
               End If
End Sub

Private Sub InitDefPlayerInfo()
'Initialize Default Player Information
 NumOfPlayers = 1 'by default set the number of players to 1
  CurrentPlayer = 1 'since there is only one player, set the number of players to one
   PlayerScore(CurrentPlayer).PlayerName = Mid$(UserName, 1, 9)
   'to prevent the label control from displaying text outside of the rectangular dimensions of the Player Icon return only the first 9 characters
    Frame1.Caption = PlayerScore(CurrentPlayer).PlayerName & "'s Score"
    'update the caption of the frame with the current players name
     ResetPlayers 'see resetplayers function for more info...
End Sub

Public Sub ChangeCurrentPlayer(Optional Inc As Integer = 1)
 CurrentPlayer = CurrentPlayer + Inc
 'update flag
  If CurrentPlayer > NumOfPlayers Then CurrentPlayer = 1
  'if CurrentPlayer is greater than the ammount of players, then reset the variable
   Select Case CurrentPlayer
    Case 1: 'if CurrentPlayer evaluates to 1 then...
     Player1Ico.Picture = imglstPlayer.ListImages(1).Picture
     'update Player1 player icon with the HOT player icon
      Player2Ico.Picture = imglstPlayer.ListImages(2).Picture
      'update the other player icons with the nutral play icon bitmaps
       Player3Ico.Picture = imglstPlayer.ListImages(2).Picture
        Player4Ico.Picture = imglstPlayer.ListImages(2).Picture
         Player5Ico.Picture = imglstPlayer.ListImages(2).Picture
          Player6Ico.Picture = imglstPlayer.ListImages(2).Picture
           lblPlayer1.ForeColor = vbBlack
            lblPlayer2.ForeColor = &H8000000F
             lblPlayer3.ForeColor = &H8000000F
              lblPlayer4.ForeColor = &H8000000F
               lblPlayer5.ForeColor = &H8000000F
                lblPlayer6.ForeColor = &H8000000F
    Case 2: '...
     Player1Ico.Picture = imglstPlayer.ListImages(2).Picture
      Player2Ico.Picture = imglstPlayer.ListImages(1).Picture
       Player3Ico.Picture = imglstPlayer.ListImages(2).Picture
        Player4Ico.Picture = imglstPlayer.ListImages(2).Picture
         Player5Ico.Picture = imglstPlayer.ListImages(2).Picture
          Player6Ico.Picture = imglstPlayer.ListImages(2).Picture
           lblPlayer1.ForeColor = &H8000000F
            lblPlayer2.ForeColor = vbBlack
             lblPlayer3.ForeColor = &H8000000F
              lblPlayer4.ForeColor = &H8000000F
               lblPlayer5.ForeColor = &H8000000F
                lblPlayer6.ForeColor = &H8000000F
    Case 3: '...
     Player1Ico.Picture = imglstPlayer.ListImages(2).Picture
      Player2Ico.Picture = imglstPlayer.ListImages(2).Picture
       Player3Ico.Picture = imglstPlayer.ListImages(1).Picture
        Player4Ico.Picture = imglstPlayer.ListImages(2).Picture
         Player5Ico.Picture = imglstPlayer.ListImages(2).Picture
          Player6Ico.Picture = imglstPlayer.ListImages(2).Picture
           lblPlayer1.ForeColor = &H8000000F
            lblPlayer2.ForeColor = &H8000000F
             lblPlayer3.ForeColor = vbBlack
              lblPlayer4.ForeColor = &H8000000F
               lblPlayer5.ForeColor = &H8000000F
                lblPlayer6.ForeColor = &H8000000F
    Case 4: '...
     Player1Ico.Picture = imglstPlayer.ListImages(2).Picture
      Player2Ico.Picture = imglstPlayer.ListImages(2).Picture
       Player3Ico.Picture = imglstPlayer.ListImages(2).Picture
        Player4Ico.Picture = imglstPlayer.ListImages(1).Picture
         Player5Ico.Picture = imglstPlayer.ListImages(2).Picture
          Player6Ico.Picture = imglstPlayer.ListImages(2).Picture
           lblPlayer1.ForeColor = &H8000000F
            lblPlayer2.ForeColor = &H8000000F
             lblPlayer3.ForeColor = &H8000000F
              lblPlayer4.ForeColor = vbBlack
               lblPlayer5.ForeColor = &H8000000F
                lblPlayer6.ForeColor = &H8000000F
    Case 5:
     Player1Ico.Picture = imglstPlayer.ListImages(2).Picture
      Player2Ico.Picture = imglstPlayer.ListImages(2).Picture
       Player3Ico.Picture = imglstPlayer.ListImages(2).Picture
        Player4Ico.Picture = imglstPlayer.ListImages(2).Picture
         Player5Ico.Picture = imglstPlayer.ListImages(1).Picture
          Player6Ico.Picture = imglstPlayer.ListImages(2).Picture
           lblPlayer1.ForeColor = &H8000000F
            lblPlayer2.ForeColor = &H8000000F
             lblPlayer3.ForeColor = &H8000000F
              lblPlayer4.ForeColor = &H8000000F
               lblPlayer5.ForeColor = vbBlack
                lblPlayer6.ForeColor = &H8000000F
    Case 6:
     Player1Ico.Picture = imglstPlayer.ListImages(2).Picture
      Player2Ico.Picture = imglstPlayer.ListImages(2).Picture
       Player3Ico.Picture = imglstPlayer.ListImages(2).Picture
        Player4Ico.Picture = imglstPlayer.ListImages(2).Picture
         Player5Ico.Picture = imglstPlayer.ListImages(2).Picture
          Player6Ico.Picture = imglstPlayer.ListImages(1).Picture
           lblPlayer1.ForeColor = &H8000000F
            lblPlayer2.ForeColor = &H8000000F
             lblPlayer3.ForeColor = &H8000000F
              lblPlayer4.ForeColor = &H8000000F
               lblPlayer5.ForeColor = &H8000000F
                lblPlayer6.ForeColor = vbBlack
   End Select
    Frame1.Caption = PlayerScore(CurrentPlayer).PlayerName & "'s Score"
    'update the frame's caption property with the current players name
     txtScoreBoxG1(0).Text = "" 'reset the score box(text box)'s text property
      txtScoreBoxG1(1).Text = "" '...
       txtScoreBoxG1(2).Text = ""
        txtScoreBoxG1(3).Text = ""
         txtScoreBoxG1(4).Text = ""
          txtScoreBoxG1(5).Text = ""
           If PlayerScore(CurrentPlayer).uOnes = True Then txtScoreBoxG1(0).Text = PlayerScore(CurrentPlayer).Ones
            If PlayerScore(CurrentPlayer).uTwos = True Then txtScoreBoxG1(1).Text = PlayerScore(CurrentPlayer).Twos
             If PlayerScore(CurrentPlayer).uThrees = True Then txtScoreBoxG1(2).Text = PlayerScore(CurrentPlayer).Threes
              If PlayerScore(CurrentPlayer).uFours = True Then txtScoreBoxG1(3).Text = PlayerScore(CurrentPlayer).Fours
               If PlayerScore(CurrentPlayer).uFives = True Then txtScoreBoxG1(4).Text = PlayerScore(CurrentPlayer).Fives
                If PlayerScore(CurrentPlayer).uSixes = True Then txtScoreBoxG1(5).Text = PlayerScore(CurrentPlayer).Sixes
                'save the players score information
                 txtScoreBoxG2(0).Text = ""
                  txtScoreBoxG2(1).Text = ""
                   txtScoreBoxG2(2).Text = ""
                    txtScoreBoxG2(3).Text = ""
                     txtScoreBoxG2(4).Text = ""
                      txtScoreBoxG2(5).Text = ""
                       txtScoreBoxG2(6).Text = ""
                       'reset the scorebox's
                        If PlayerScore(CurrentPlayer).uThreeKind = True Then txtScoreBoxG2(0).Text = PlayerScore(CurrentPlayer).ThreeKind
                         If PlayerScore(CurrentPlayer).uFourKind = True Then txtScoreBoxG2(1).Text = PlayerScore(CurrentPlayer).FourKind
                          If PlayerScore(CurrentPlayer).uFullHouse = True Then txtScoreBoxG2(2).Text = PlayerScore(CurrentPlayer).FullHouse
                           If PlayerScore(CurrentPlayer).uLowStraight = True Then txtScoreBoxG2(3).Text = PlayerScore(CurrentPlayer).LowStraight
                            If PlayerScore(CurrentPlayer).uHightStraight = True Then txtScoreBoxG2(4).Text = PlayerScore(CurrentPlayer).HightStraight
                             If PlayerScore(CurrentPlayer).uYahtzee = True Then txtScoreBoxG2(5).Text = PlayerScore(CurrentPlayer).Yahtzee
                             'save the current players score information
If PlayerScore(CurrentPlayer).uChance = True Then txtScoreBoxG2(6).Text = PlayerScore(CurrentPlayer).Chance
 txtTotalScore.Text = PlayerScore(CurrentPlayer).Total
  txtBonus.Text = PlayerScore(CurrentPlayer).Bonus
   score1.Caption = Str$(PlayerScore(1).Total)
    score2.Caption = Str$(PlayerScore(2).Total)
     score3.Caption = Str$(PlayerScore(3).Total)
      score4.Caption = Str$(PlayerScore(4).Total)
       score5.Caption = Str$(PlayerScore(5).Total)
        score6.Caption = Str$(PlayerScore(6).Total)
        'update the score label(Score label in the Player Icon pictureboxes) caption
          
End Sub


Private Sub ResetPlayers(Optional Loading As Boolean = False)
If Loading = False Then CurrentDraw = 0
 Select Case NumOfPlayers
  Case 1 'if NumOfPlayers evaluates to 1 then
   Player1Ico.Visible = True
    lblPlayer1.Caption = PlayerScore(1).PlayerName
     Player2Ico.Visible = False
      Player3Ico.Visible = False
       Player4Ico.Visible = False
        Player5Ico.Visible = False
         Player6Ico.Visible = False
         'update the Player specific sections of the user interface
  Case 2
   Player1Ico.Visible = True
    lblPlayer1.Caption = PlayerScore(1).PlayerName
     Player2Ico.Visible = True
      lblPlayer2.Caption = PlayerScore(2).PlayerName
       Player3Ico.Visible = False
        Player4Ico.Visible = False
         Player5Ico.Visible = False
          Player6Ico.Visible = False
  Case 3
   Player1Ico.Visible = True
    lblPlayer1.Caption = PlayerScore(1).PlayerName
     Player2Ico.Visible = True
      lblPlayer2.Caption = PlayerScore(2).PlayerName
       Player3Ico.Visible = True
        lblPlayer3.Caption = PlayerScore(3).PlayerName
         Player4Ico.Visible = False
          Player5Ico.Visible = False
           Player6Ico.Visible = False
  Case 4
   Player1Ico.Visible = True
    lblPlayer1.Caption = PlayerScore(1).PlayerName
      Player2Ico.Visible = True
       lblPlayer2.Caption = PlayerScore(2).PlayerName
        Player3Ico.Visible = True
         lblPlayer3.Caption = PlayerScore(3).PlayerName
          Player4Ico.Visible = True
           lblPlayer4.Caption = PlayerScore(4).PlayerName
            Player5Ico.Visible = False
             Player6Ico.Visible = False
 Case 5
  Player1Ico.Visible = True
   lblPlayer1.Caption = PlayerScore(1).PlayerName
    Player2Ico.Visible = True
     lblPlayer2.Caption = PlayerScore(2).PlayerName
      Player3Ico.Visible = True
       lblPlayer3.Caption = PlayerScore(3).PlayerName
        Player4Ico.Visible = True
        lblPlayer4.Caption = PlayerScore(4).PlayerName
         Player5Ico.Visible = True
          lblPlayer5.Caption = PlayerScore(5).PlayerName
           Player6Ico.Visible = False
Case 6
 Player1Ico.Visible = True
  lblPlayer1.Caption = PlayerScore(1).PlayerName
   Player2Ico.Visible = True
    lblPlayer2.Caption = PlayerScore(2).PlayerName
     Player3Ico.Visible = True
      lblPlayer3.Caption = PlayerScore(3).PlayerName
       Player4Ico.Visible = True
        lblPlayer4.Caption = PlayerScore(4).PlayerName
         Player5Ico.Visible = True
          lblPlayer5.Caption = PlayerScore(5).PlayerName
           Player6Ico.Visible = True
            lblPlayer6.Caption = PlayerScore(6).PlayerName
 End Select
  Dim i& 'dimensionalize variable i as long data type
   For i = 1 To 6
    PlayerScore(i).Bonus = 0
     PlayerScore(i).Chance = 0
      PlayerScore(i).uChance = False
       PlayerScore(i).Fives = 0
        PlayerScore(i).uFives = False
         PlayerScore(i).FourKind = 0
          PlayerScore(i).uFourKind = False
           PlayerScore(i).Fours = 0
            PlayerScore(i).uFours = False
             PlayerScore(i).FullHouse = 0
              PlayerScore(i).uFullHouse = False
               PlayerScore(i).HightStraight = 0
                PlayerScore(i).uHightStraight = False
                 PlayerScore(i).LowStraight = 0
                  PlayerScore(i).uLowStraight = False
                   PlayerScore(i).Ones = 0
                    PlayerScore(i).uOnes = False
                     PlayerScore(i).Sixes = 0
                      PlayerScore(i).uSixes = False
                       PlayerScore(i).ThreeKind = 0
                        PlayerScore(i).uThreeKind = False
                         PlayerScore(i).Threes = 0
                          PlayerScore(i).uThrees = False
                           PlayerScore(i).Total = 0
                            PlayerScore(i).Twos = 0
                             PlayerScore(i).uTwos = False
                              PlayerScore(i).Yahtzee = 0
                               PlayerScore(i).uYahtzee = False
                                PlayerScore(i).YahtzeeOccur = 0
  Next i 'reset each players virtual score card
   score1.Caption = "0"
    score2.Caption = "0"
     score3.Caption = "0"
      score4.Caption = "0"
       score5.Caption = "0"
        score6.Caption = "0"
         Dice1Stat.Holding = False
          hold1.Picture = imglstHold.ListImages.Item(1).Picture
           box1.BorderColor = &HD5D5D5
            Dice2Stat.Holding = False
             hold2.Picture = imglstHold.ListImages.Item(1).Picture
              box2.BorderColor = &HD5D5D5
               Dice3Stat.Holding = False
                hold3.Picture = imglstHold.ListImages.Item(1).Picture
                 box3.BorderColor = &HD5D5D5
                  Dice4Stat.Holding = False
                   hold4.Picture = imglstHold.ListImages.Item(1).Picture
                    box4.BorderColor = &HD5D5D5
                     Dice5Stat.Holding = False
                      hold5.Picture = imglstHold.ListImages.Item(1).Picture
                       box5.BorderColor = &HD5D5D5
                       'reset all Hold pictureboxes to the non holding pictures
                        If Loading = False Then
                        'if the application is just loading then a new game has to be started...
                         For i = 0 To txtScoreBoxG1.UBound
                          txtScoreBoxG1(i).MousePointer = 1
                          'set each score boxes mouse pointer to default
                         Next i
                         For i = 0 To txtScoreBoxG2.UBound
                          txtScoreBoxG2(i).MousePointer = 1
                          'set each score boxes mouse pointer to default
                         Next i
                        End If
                         '...
                         'if the application is just loading then a new game has to be started...
                         If Loading = False Then EnableRollDiceBtn True
                          If Loading = False Then CurrentDraw = 0
                           If Loading = False Then OldDice = True
                            If Loading = False Then ChoosingScoreBox = False
                             If Loading = False Then lblstatus.Caption = "Please Roll Dice!!"
End Sub

Public Function RemoveAndCopyPlayerInf(PlayerIndex As Integer)
Dim i&, tmpNewSpace% 'dimensionalize i as long data type, tmpNewSpace as integer type
  If PlayerIndex < NumOfPlayers And PlayerIndex > 1 Then
  'if PlayerIndex is less that NumOfPlayers and PlayerIndex is greater than one then...
   For i = (PlayerIndex + 1) To NumOfPlayers
   'for next loop; initialize i to PlayerIndex + 1; looping until i is equal to NumOfPlayer; incrementing i by one each iteration
    tmpNewSpace = PlayerIndex 'initialize tmpNewSpace to PlayerIndex
     PlayerScore(tmpNewSpace).PlayerName = PlayerScore(i).PlayerName
      tmpNewSpace = tmpNewSpace + 1
   Next i
   'shift the elements in the PlayerScore array
  ElseIf PlayerIndex = 1 Then
   For i = 2 To NumOfPlayers
    tmpNewSpace = 1
     PlayerScore(tmpNewSpace).PlayerName = PlayerScore(i).PlayerName
      tmpNewSpace = tmpNewSpace + 1
   Next i
  End If
   NumOfPlayers = NumOfPlayers - 1 'decrement NumOfPlayers
    CurrentPlayer = NumOfPlayers 'CurrentPlayer is initialized to NumOfPlayers so that it will be incremented again but because it will be greater than the number of players(NumOfPlayers) it will be set to zero
     ResetPlayers 'See ResetPlayers for more info...
      ChangeCurrentPlayer 'see ChangeCurrentPlayer for more info....
End Function

Public Function AddPlayer(PlayerName As String)
ChangePlayerflg = 0
 NumOfPlayers = NumOfPlayers + 1 'increment NumOfPlayers
  PlayerScore(NumOfPlayers).PlayerName = PlayerName
  'Update the new elelement in the PlayerScore array
   ResetPlayers 'see ResetPlayers function for more info...
    CurrentPlayer = NumOfPlayers
     ChangeCurrentPlayer 'see ChangeCurrentPlayer for more info...
      ResetPlayers 'see ResetPlayers for more info...
       OldDice = True 'initialize OldDice(Frozen Dice) to true
        Say "welcome " & PlayerName 'see Say function for more info...
End Function


Sub preloadDice(Color As DicCol)
 Select Case Color
  Case DicCol.Blue 'if Color evaluates to the enumeration DicCol's Blue member
   picDice1.Picture = imglstDice.ListImages.Item("b1").Picture
   'set picture box control picDice1's picture property to the GDI Object returned by the item method derived from the ImageList's ListImages collection
    picDice2.Picture = imglstDice.ListImages.Item("b2").Picture
     picDice3.Picture = imglstDice.ListImages.Item("b3").Picture
      picDice4.Picture = imglstDice.ListImages.Item("b4").Picture
       picDice5.Picture = imglstDice.ListImages.Item("b5").Picture
        picDice6.Picture = imglstDice.ListImages.Item("b6").Picture
        'When a use rolls the dice, it copies the appropriate dice picture from one of the picDice#'s picture box controls
        'so, when the user changes his or her dice color preference the dice will be refreshed
  Case DicCol.Gold '...
   picDice1.Picture = imglstDice.ListImages.Item("gld1").Picture '...
    picDice2.Picture = imglstDice.ListImages.Item("gld2").Picture
     picDice3.Picture = imglstDice.ListImages.Item("gld3").Picture
      picDice4.Picture = imglstDice.ListImages.Item("gld4").Picture
       picDice5.Picture = imglstDice.ListImages.Item("gld5").Picture
        picDice6.Picture = imglstDice.ListImages.Item("gld6").Picture
  Case DicCol.Red
   picDice1.Picture = imglstDice.ListImages.Item("r1").Picture
    picDice2.Picture = imglstDice.ListImages.Item("r2").Picture
     picDice3.Picture = imglstDice.ListImages.Item("r3").Picture
      picDice4.Picture = imglstDice.ListImages.Item("r4").Picture
       picDice5.Picture = imglstDice.ListImages.Item("r5").Picture
        picDice6.Picture = imglstDice.ListImages.Item("r6").Picture
  Case DicCol.Green
   picDice1.Picture = imglstDice.ListImages.Item("g1").Picture
    picDice2.Picture = imglstDice.ListImages.Item("g2").Picture
     picDice3.Picture = imglstDice.ListImages.Item("g3").Picture
      picDice4.Picture = imglstDice.ListImages.Item("g4").Picture
       picDice5.Picture = imglstDice.ListImages.Item("g5").Picture
        picDice6.Picture = imglstDice.ListImages.Item("g6").Picture
 End Select
  Select Case Dice1Stat.FaceValue
   Case 1: 'if Dice 1's face value evaluates to 1 then
    dice1.Picture = picDice1.Picture
    'Refresh the Dice display with the properly colored dice
   Case 2:
    dice1.Picture = picDice2.Picture
   Case 3:
    dice1.Picture = picDice3.Picture
   Case 4:
    dice1.Picture = picDice4.Picture
   Case 5:
    dice1.Picture = picDice5.Picture
   Case 6:
    dice1.Picture = picDice6.Picture
  End Select
  
   Select Case Dice2Stat.FaceValue
    Case 1:
     dice2.Picture = picDice1.Picture
    Case 2:
     dice2.Picture = picDice2.Picture
    Case 3:
     dice2.Picture = picDice3.Picture
    Case 4:
     dice2.Picture = picDice4.Picture
    Case 5:
     dice2.Picture = picDice5.Picture
    Case 6:
     dice2.Picture = picDice6.Picture
   End Select
   '...
    Select Case Dice3Stat.FaceValue
     Case 1:
      dice3.Picture = picDice1.Picture
     Case 2:
      dice3.Picture = picDice2.Picture
     Case 3:
      dice3.Picture = picDice3.Picture
     Case 4:
      dice3.Picture = picDice4.Picture
     Case 5:
      dice3.Picture = picDice5.Picture
     Case 6:
      dice3.Picture = picDice6.Picture
    End Select
    '...
     Select Case Dice4Stat.FaceValue
      Case 1:
       dice4.Picture = picDice1.Picture
      Case 2:
       dice4.Picture = picDice2.Picture
      Case 3:
       dice4.Picture = picDice3.Picture
      Case 4:
       dice4.Picture = picDice4.Picture
      Case 5:
       dice4.Picture = picDice5.Picture
      Case 6:
       dice4.Picture = picDice6.Picture
    End Select
    '...
     Select Case Dice5Stat.FaceValue
      Case 1:
       dice5.Picture = picDice1.Picture
      Case 2:
       dice5.Picture = picDice2.Picture
      Case 3:
       dice5.Picture = picDice3.Picture
      Case 4:
       dice5.Picture = picDice4.Picture
      Case 5:
       dice5.Picture = picDice5.Picture
      Case 6:
       dice5.Picture = picDice6.Picture
     End Select
     '...
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim PrefDieCol$ 'dimensionalize PrefDieCol as string data type
 If MsgBox("Are you sure you want to quit?", vbQuestion + vbYesNo, "Confirm Exit") = vbNo Then Cancel = 1: Exit Sub
 'request user confirmation
  Set dss = Nothing 'terminate object dss of DirectSS class
  'when an instance of any class is no longer needed, its terminated so that
  'it's deconstructor(Terminate Sub Routine) is called to clean up the resources it has utilized...
   SaveSetting "YAH", "SNDOPT", "PSflag", IIf(PlaySounds = True, "1", "0")
   'save the Play Sound flag in the registry
   'note: IIf operator is a conditional statement return
   'return = IIf(Statement, TruePart, FalsePart)
    Select Case CurDiceCol.DiceColor
     Case Blue: 'if CurDiceCol's DiceColor member evaluates to constant Blue(Defined in its enumeration type) then...
      PrefDieCol = "b" 'update PrefDieCol(Preferred Dice Color)
     Case Red:
      PrefDieCol = "r"
     Case Green:
      PrefDieCol = "g"
     Case Gold:
      PrefDieCol = "gld"
    End Select
     SaveSetting "YAH", "PREF", "DIECOL", PrefDieCol$
     'save the Dice Color in the registry
     'SaveSetting function is User specific as it saves the registry key in the HKEY_CURRENT_USER key handle
       Dim Form As Form 'declare variable form as form class
        For Each Form In Forms 'for each loop; enumerates through each form in the Forms collection
        'note: the Forms collection consists of only the forms loaded in memory
         If LCase(Form.Name) <> "frmmain" Then Unload Form
         'if the current form in the forms collection is not this form, then unload it from memory
        Next Form 'select the next form element in the forms collection
         End 'terminate process(forces each loaded module to unload)
End Sub

Private Sub hold1_Click()
 If OldDice = True Then Exit Sub
 'if OldDice(Frozen Dice) evaluates to true then exit this procedure
  HoldDice 1 'see HoldDice function for more info...
End Sub

Public Sub HoldDice(Dice%)
 Select Case Dice
  Case 1 'if Dice evaluates to 1
   Dice1Stat.Holding = Not (Dice1Stat.Holding)
   'invert the boolean type Holding member of Dice1Stat structure...
    If Dice1Stat.Holding = False Then
    'if Dice1Stat's Holding member evaluates to false then...
     hold1.Picture = imglstHold.ListImages.Item(1).Picture
     'update picture (Nutral State)
      box1.BorderColor = &HD5D5D5
      'update shape control's BorderColor (Nutral State)
    Else
     hold1.Picture = imglstHold.ListImages.Item(2).Picture 'Holding State
      box1.BorderColor = 33023 'Holding State
    End If

   Case 2 '...
    Dice2Stat.Holding = Not (Dice2Stat.Holding)
     If Dice2Stat.Holding = False Then
      hold2.Picture = imglstHold.ListImages.Item(1).Picture
       box2.BorderColor = &HD5D5D5
     Else
      hold2.Picture = imglstHold.ListImages.Item(2).Picture
       box2.BorderColor = 33023
     End If
      
   Case 3
    Dice3Stat.Holding = Not (Dice3Stat.Holding)
     If Dice3Stat.Holding = False Then
      hold3.Picture = imglstHold.ListImages.Item(1).Picture
       box3.BorderColor = &HD5D5D5
     Else
      hold3.Picture = imglstHold.ListImages.Item(2).Picture
       box3.BorderColor = 33023
     End If
      
    Case 4
     Dice4Stat.Holding = Not (Dice4Stat.Holding)
      If Dice4Stat.Holding = False Then
       hold4.Picture = imglstHold.ListImages.Item(1).Picture
        box4.BorderColor = &HD5D5D5
      Else
       hold4.Picture = imglstHold.ListImages.Item(2).Picture
        box4.BorderColor = 33023
      End If
      
    Case 5
     Dice5Stat.Holding = Not (Dice5Stat.Holding)
      If Dice5Stat.Holding = False Then
       hold5.Picture = imglstHold.ListImages.Item(1).Picture
        box5.BorderColor = &HD5D5D5
      Else
       hold5.Picture = imglstHold.ListImages.Item(2).Picture
        box5.BorderColor = 33023
      End If
 End Select
End Sub

Private Sub hold2_Click()
 If OldDice = True Then Exit Sub
 'if OldDice(Frozen Dice) evaluates to True then exit this procedure
  HoldDice 2 'see HoldDice function for more info...
  '(Set Dice 2's holding state)
End Sub

Private Sub hold3_Click()
 If OldDice = True Then Exit Sub '...
  HoldDice 3 '...
End Sub

Private Sub hold4_Click()
 If OldDice = True Then Exit Sub
  HoldDice 4
End Sub

Private Sub hold5_Click()
 If OldDice = True Then Exit Sub
  HoldDice 5
End Sub

Private Sub RollDie()
Dim i&, DiceValues(1 To 5) As Integer
'dimensionalize i as long data type, one dimensional array DiceValues(6 elements) as integer type
 Randomize (Second(Time) * (Int((30 - 1 + 1) * Rnd + 1)))
 'initializes the random-number generator
 'Randomize function uses the number paramater to initialize the Rnd function's random-number generator setting a new seed value
  For i = 1 To 5
  'for next loop; initialize i to 1; loop until i evaluates to 5 incrementing i by one each iteration
   DiceValues(i) = Int((6 - 1 + 1) * Rnd + 1)
   'randomly gerate the face value for each of the 5 die
  Next i 'increment i; evaluate loop condition; next iteration
   If Dice1Stat.Holding = False Then
   'if the holding state of Dice 1 is false then...
    Dice1Stat.FaceValue = DiceValues(1)
    'update the Face Value of Dice 1 with the randomly generated face value
     Select Case Dice1Stat.FaceValue
      Case 1: 'if the face value(Dice1Stat's FaceValue member) evaluates to 1
       dice1.Picture = picDice1.Picture 'update picture
      Case 2:
       dice1.Picture = picDice2.Picture
      Case 3:
       dice1.Picture = picDice3.Picture
      Case 4:
       dice1.Picture = picDice4.Picture
      Case 5:
       dice1.Picture = picDice5.Picture
      Case 6:
       dice1.Picture = picDice6.Picture
     End Select
  End If

   If Dice2Stat.Holding = False Then '...
    Dice2Stat.FaceValue = DiceValues(2) '...
     Select Case Dice2Stat.FaceValue
      Case 1:
       dice2.Picture = picDice1.Picture
      Case 2:
      dice2.Picture = picDice2.Picture
      Case 3:
       dice2.Picture = picDice3.Picture
      Case 4:
       dice2.Picture = picDice4.Picture
      Case 5:
       dice2.Picture = picDice5.Picture
      Case 6:
       dice2.Picture = picDice6.Picture
     End Select
   End If
  
    If Dice3Stat.Holding = False Then
     Dice3Stat.FaceValue = DiceValues(3)
      Select Case Dice3Stat.FaceValue
       Case 1:
        dice3.Picture = picDice1.Picture
       Case 2:
        dice3.Picture = picDice2.Picture
       Case 3:
        dice3.Picture = picDice3.Picture
       Case 4:
        dice3.Picture = picDice4.Picture
       Case 5:
        dice3.Picture = picDice5.Picture
       Case 6:
        dice3.Picture = picDice6.Picture
      End Select
    End If
  
  
     If Dice4Stat.Holding = False Then
      Dice4Stat.FaceValue = DiceValues(4)
       Select Case Dice4Stat.FaceValue
        Case 1:
         dice4.Picture = picDice1.Picture
        Case 2:
         dice4.Picture = picDice2.Picture
        Case 3:
         dice4.Picture = picDice3.Picture
        Case 4:
         dice4.Picture = picDice4.Picture
        Case 5:
         dice4.Picture = picDice5.Picture
        Case 6:
         dice4.Picture = picDice6.Picture
       End Select
     End If
  
      If Dice5Stat.Holding = False Then
       Dice5Stat.FaceValue = DiceValues(5)
        Select Case Dice5Stat.FaceValue
         Case 1:
          dice5.Picture = picDice1.Picture
         Case 2:
          dice5.Picture = picDice2.Picture
         Case 3:
          dice5.Picture = picDice3.Picture
         Case 4:
          dice5.Picture = picDice4.Picture
         Case 5:
          dice5.Picture = picDice5.Picture
         Case 6:
          dice5.Picture = picDice6.Picture
        End Select
      End If
End Sub

Private Sub mnuAbout_Click()
 Load frmAbout 'load frmAbout dialog into memory
  frmAbout.Show vbModal, Me 'show the dialog as a modal dialog
End Sub

Private Sub mnuAddPLayer_Click()
 Load frmPlayers '...
  frmPlayers.Show vbModal, Me
End Sub

Private Sub mnuBlue_Click()
 CurDiceCol.DiceColor = Blue: preloadDice CurDiceCol.DiceColor
 'update CurDicCol's DiceColor member to Blue; see preloadDice function for more info(Refresh the Dice picture for each die)
  mnuBlue.Checked = True 'set the menu items class wrapper's Checked property to true
   mnuRed.Checked = False
    mnuGreen.Checked = False
     mnuGold.Checked = False
End Sub

Private Sub mnuContents_Click()
Dim nRet% 'dimensionalize nRet as integer type
StartSub: 'label StartSub
 If Dir(App.HelpFile) = "" Then
 'if the Help file doesn't exist(or has special file attributes which prevent reading it) then...
  GoTo NHF 'just to NHF(No Help File) label
 Else
  On Error Resume Next 'on the event of an error resume execution on the next line of this procedure
   nRet = OSWinHelp(Me.hwnd, App.HelpFile, 3, 0)
   'open this files help file
    If Err Then MsgBox Err.Description
    'an error occured; display the description of the error
 End If
  Exit Sub 'exit this procedure
NHF: 'label NHF(No Help File)
     If Dir(App.Path & "\YahtzeeXP.HLP") <> "" Then
     'If the file exists then...
      App.HelpFile = App.Path & "\YahtzeeXP.HLP"
      'update object App's HelpFile property
     Else
      If Dir(App.Path & "\YahtzeeXP.HLP") = "" Then
      'if the specified file doesn't exist then...
       If MsgBox("Can't locate help file." & vbCrLf & vbCrLf & "If you have moved the help file associated with this program(Yahtzee XP) please return it to the installation directory and click Retry, otherwise click Cancel", vbCritical + vbRetryCancel, "Help File Missing") = vbRetry Then
       'request user confirmation to retry to find the help file
        If Dir(App.Path & "\YahtzeeXP.HLP") <> "" Then
        'if the file exists then...
         App.HelpFile = App.Path & "\YahtzeeXP.HLP"
         'update the help file path
          MsgBox "Help file found." & vbCrLf & """" & App.HelpFile & """", vbInformation, "Help File"
          'inform user the help file was found
        Else
         If Dir(App.Path & "YahtzeeXP.HLP") <> "" Then
         'if the file exists then
          App.HelpFile = App.Path & "\YahtzeeXP.HLP"
          'update the help file path
           MsgBox "Help file found." & vbCrLf & """" & App.HelpFile & """", vbInformation, "Help File" '...
         End If
        End If
       Else
        Exit Sub
       End If
      Else
       App.HelpFile = App.Path & "\YahtzeeXP.HLP"
     End If
    End If
     GoTo StartSub 'jump to the beggining of this procedure
End Sub

Private Sub mnuExit_Click()
 Unload Me 'unload this dialog
End Sub

Private Sub mnuGold_Click()
'see mnuBlue_Click for more info...
CurDiceCol.DiceColor = Gold: preloadDice CurDiceCol.DiceColor
 mnuBlue.Checked = False
  mnuRed.Checked = False
   mnuGreen.Checked = False
    mnuGold.Checked = True
End Sub

Private Sub mnuGreen_Click()
'see mnuBlue_Click for more info...
CurDiceCol.DiceColor = Green: preloadDice CurDiceCol.DiceColor
 mnuBlue.Checked = False
  mnuRed.Checked = False
   mnuGreen.Checked = True
    mnuGold.Checked = False
End Sub

Private Sub mnuHighScores_Click()
 Load frmHighScores 'load dialog frmHighScores into memory
  frmHighScores.Show vbModal, Me
  'show the dialog as a modal dialog
End Sub

Private Sub mnuLoadGame_Click()
 LoadGame True 'see LoadGame function for more info...
End Sub

Private Function LoadGame(ShowUI As Boolean, Optional FileName As String)
On Error GoTo errh 'on the even of an error jump to label errh(error handler)
Dim FN$, i% 'dimensionalize FN as string type, i as integer type
 If ShowUI = True Then 'if ShowUI(Common Dialog) flag evaluates to true
  If MsgBox("Are you sure you want to quit you're current game and load an old game?", vbQuestion + vbYesNo, "Load New Game") = vbNo Then Exit Function
  'request user confirmation
   CD.Flags = cdlOFNFileMustExist
   'cdlOFNFileMustExist flag - User can only enter a file name of an existing file
   CD.Filter = "Yahtzee XP [*.yxp]|*.yxp"
   'Initialize Filter member with File Extension pattern
   'syntax: "Description|FileNamePattern|Description2|FileNamePattern;FileNamePattern2;FileNamePattern3"
    CD.ShowOpen 'call CD(Common Dialog)'s ShowOpen method to Show the Open File modal Dialog
     FN$ = CD.FileName 'initialize FN with the filename of the file selected by the user
 Else
  FN$ = FileName
  'since ShowUI(Show User Interface) evaluates to false, the optional paramater FileName specifies the file to load
 End If
  NumOfPlayers = Int(ReadFromINI("GameInfo", "NumberOfPlayers", FN))
  'see ReadFromINI for more info...
   CurrentPlayer = Int(ReadFromINI("GameInfo", "CurrentPlayer", FN))
    For i = 1 To NumOfPlayers
     PlayerScore(i).PlayerName = ReadFromINI(Str$(i), "PlayerName", FN)
     'update the player names...
    Next i
     ResetPlayers True 'see ResetPlayers for more info...
      CurrentDraw = Int(ReadFromINI("GameInfo", "CurrentDraw", FN))
      'update the CurrentDraw flag
       OldDice = Int(ReadFromINI("GameInfo", "OldDice", FN))
       'update OldDice(Frozen Dice) flag
        ChangePlayerflg = Int(ReadFromINI("GameInfo", "ChangePlayerflg", FN))
        'update flag...
         lblstatus.Caption = ReadFromINI("GameInfo", "StatusLbl", FN)
         'update the Status label caption...
          ChoosingScoreBox = Int(ReadFromINI("GameInfo", "ChoosingScoreBox", FN))
          'update flag...
           For i = 0 To txtScoreBoxG1.UBound
            txtScoreBoxG1(i).MousePointer = ReadFromINI("GameInfo", "ScoreBoxG1MP", FN)
            'MousePointer will evaluate to either 1(default) or 99(Custom[Uses MouseIcon property])
           Next i
            For i = 0 To txtScoreBoxG2.UBound
             txtScoreBoxG2(i).MousePointer = ReadFromINI("GameInfo", "ScoreBoxG2MP", FN)
             '...
            Next i
             If ChoosingScoreBox = False Then EnableRollDiceBtn True Else EnableRollDiceBtn False
             'see EnableRollDiceBtn function for more info...(Update Picture and Disabled property)
             Dice1Stat.FaceValue = Int(ReadFromINI("GameInfo", "Dice1FV", FN))
             'update Dice 1's face value
              Dice1Stat.Holding = Int(ReadFromINI("GameInfo", "Dice1H", FN))
              'update Dice 1's holding state
               Dice1Stat.Holding = Not (Dice1Stat.Holding): HoldDice 1
               'invert the holding state, because it will be inverted again when the HoldDice method is called
               'see HoldDice for more info
             Dice2Stat.FaceValue = Int(ReadFromINI("GameInfo", "Dice2FV", FN))
              Dice2Stat.Holding = Int(ReadFromINI("GameInfo", "Dice2H", FN))
               Dice2Stat.Holding = Not (Dice2Stat.Holding): HoldDice 2
             Dice3Stat.FaceValue = Int(ReadFromINI("GameInfo", "Dice3FV", FN))
              Dice3Stat.Holding = Int(ReadFromINI("GameInfo", "Dice3H", FN))
               Dice3Stat.Holding = Not (Dice3Stat.Holding): HoldDice 3
             Dice4Stat.FaceValue = Int(ReadFromINI("GameInfo", "Dice4FV", FN))
              Dice4Stat.Holding = Int(ReadFromINI("GameInfo", "Dice4H", FN))
               Dice4Stat.Holding = Not (Dice4Stat.Holding): HoldDice 4
             Dice5Stat.FaceValue = Int(ReadFromINI("GameInfo", "Dice5FV", FN))
              Dice5Stat.Holding = Int(ReadFromINI("GameInfo", "Dice5H", FN))
               Dice5Stat.Holding = Not (Dice5Stat.Holding): HoldDice 5
 
                Select Case Dice1Stat.FaceValue
                 Case 1: 'if Dice 1's face value evaluates to 1 then...
                  dice1.Picture = picDice1.Picture 'update dice picture
                 Case 2:
                  dice1.Picture = picDice2.Picture
                 Case 3:
                  dice1.Picture = picDice3.Picture
                 Case 4:
                  dice1.Picture = picDice4.Picture
                 Case 5:
                  dice1.Picture = picDice5.Picture
                 Case 6:
                  dice1.Picture = picDice6.Picture
                End Select
 
                 Select Case Dice2Stat.FaceValue
                  Case 1: '...
                   dice2.Picture = picDice1.Picture
                  Case 2:
                   dice2.Picture = picDice2.Picture
                  Case 3:
                   dice2.Picture = picDice3.Picture
                  Case 4:
                   dice2.Picture = picDice4.Picture
                  Case 5:
                   dice2.Picture = picDice5.Picture
                  Case 6:
                   dice2.Picture = picDice6.Picture
                End Select
 
                 Select Case Dice3Stat.FaceValue
                  Case 1:
                   dice3.Picture = picDice1.Picture
                  Case 2:
                   dice3.Picture = picDice2.Picture
                  Case 3:
                   dice3.Picture = picDice3.Picture
                  Case 4:
                   dice3.Picture = picDice4.Picture
                  Case 5:
                   dice3.Picture = picDice5.Picture
                  Case 6:
                   dice3.Picture = picDice6.Picture
                End Select
 
                 Select Case Dice4Stat.FaceValue
                  Case 1:
                   dice4.Picture = picDice1.Picture
                  Case 2:
                   dice4.Picture = picDice2.Picture
                  Case 3:
                   dice4.Picture = picDice3.Picture
                  Case 4:
                   dice4.Picture = picDice4.Picture
                  Case 5:
                   dice4.Picture = picDice5.Picture
                  Case 6:
                   dice4.Picture = picDice6.Picture
                End Select
 
                 Select Case Dice5Stat.FaceValue
                  Case 1:
                   dice5.Picture = picDice1.Picture
                  Case 2:
                   dice5.Picture = picDice2.Picture
                  Case 3:
                   dice5.Picture = picDice3.Picture
                  Case 4:
                   dice5.Picture = picDice4.Picture
                  Case 5:
                   dice5.Picture = picDice5.Picture
                  Case 6:
                   dice5.Picture = picDice6.Picture
                End Select
 

  For i = 1 To NumOfPlayers
  'enumerate through each player of the game, and update their virtual score card
   PlayerScore(i).Bonus = Int(ReadFromINI(Str$(i), "Bonus", FN))
   'int function returns the integral value of a string
   'similar to StrToInt API, see ReadFromINI for more info...
    PlayerScore(i).Chance = Int(ReadFromINI(Str$(i), "Chance", FN))
     PlayerScore(i).Fives = Int(ReadFromINI(Str$(i), "Fives", FN))
      PlayerScore(i).FourKind = Int(ReadFromINI(Str$(i), "FourKind", FN))
       PlayerScore(i).Fours = Int(ReadFromINI(Str$(i), "Fours", FN))
        PlayerScore(i).FullHouse = Int(ReadFromINI(Str$(i), "FullHouse", FN))
         PlayerScore(i).HightStraight = Int(ReadFromINI(Str$(i), "HightStraight", FN))
          PlayerScore(i).LowStraight = Int(ReadFromINI(Str$(i), "LowStraight", FN))
           PlayerScore(i).Ones = Int(ReadFromINI(Str$(i), "Ones", FN))
            PlayerScore(i).Sixes = Int(ReadFromINI(Str$(i), "Sixes", FN))
             PlayerScore(i).ThreeKind = Int(ReadFromINI(Str$(i), "ThreeKind", FN))
              PlayerScore(i).Threes = Int(ReadFromINI(Str$(i), "Threes", FN))
               PlayerScore(i).Total = Int(ReadFromINI(Str$(i), "Total", FN))
                PlayerScore(i).Twos = Int(ReadFromINI(Str$(i), "Twos", FN))
                 PlayerScore(i).Yahtzee = Int(ReadFromINI(Str$(i), "Yahtzee", FN))
                  PlayerScore(i).YahtzeeOccur = Int(ReadFromINI(Str$(i), "YahtzeeOccur", FN))
   
   PlayerScore(i).uChance = Int(ReadFromINI(Str$(i), "uChance", FN))
    PlayerScore(i).uFives = Int(ReadFromINI(Str$(i), "uFives", FN))
     PlayerScore(i).uFourKind = Int(ReadFromINI(Str$(i), "uFourKind", FN))
      PlayerScore(i).uFours = Int(ReadFromINI(Str$(i), "uFours", FN))
       PlayerScore(i).uFullHouse = Int(ReadFromINI(Str$(i), "uFullHouse", FN))
        PlayerScore(i).uHightStraight = Int(ReadFromINI(Str$(i), "uHightStraight", FN))
         PlayerScore(i).uLowStraight = Int(ReadFromINI(Str$(i), "uLowStraight", FN))
          PlayerScore(i).uOnes = Int(ReadFromINI(Str$(i), "uOnes", FN))
           PlayerScore(i).uSixes = Int(ReadFromINI(Str$(i), "uSixes", FN))
            PlayerScore(i).uThreeKind = Int(ReadFromINI(Str$(i), "uThreeKind", FN))
             PlayerScore(i).uThrees = Int(ReadFromINI(Str$(i), "uThrees", FN))
              PlayerScore(i).uTwos = Int(ReadFromINI(Str$(i), "uTwos", FN))
               PlayerScore(i).uYahtzee = Int(ReadFromINI(Str$(i), "uYahtzee", FN))
                ChangeCurrentPlayer 0 'see ChangeCurrentPlayer function for more info...
  Next i
   MsgBox "Game has been resumed from" & vbCrLf & FN, vbInformation, "Game Resumed"
   'inform user that the game has been loaded
    Exit Function 'exit this procedure
errh: 'label errh
  If Err.Number = 32755 Then Exit Function 'object defined error 32755 occurs when the user cancelled the dialog shown by CD's ShowOpen method(User clicked the Cancel button)
   MsgBox "The following error has occured while loading the game." & vbCrLf & vbCrLf & Err.Description, vbCritical, "Error [" & Err.Number & "]"
   'notify the user of the un-expected error
End Function


Private Sub mnuNewGame_Click()
 If MsgBox("Are you sure you wish to start a new game?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
 'request user confirmation
  ResetPlayers 'see ResetPlayers for more info
   CurrentPlayer = NumOfPlayers 'CurrentPlayer is set to the Number of players as the ChangeCurrentPlayer function will increment this variable, since it will be greater than the number of players it will be reset to 0
    ChangeCurrentPlayer 'see ChangeCurrentPlayer function for more info...
End Sub

Private Sub mnuPrefs_Click()
 Load frmPref 'load frmPref dialog resource as a form object
  frmPref.Show vbModal, Me 'show the dialog as a modal dialog
End Sub

Private Sub mnuRed_Click()
'see mnuBlue_Click for more info...
 CurDiceCol.DiceColor = Red: preloadDice CurDiceCol.DiceColor
  mnuBlue.Checked = False
   mnuRed.Checked = True
    mnuGreen.Checked = False
     mnuGold.Checked = False
End Sub

Private Sub mnuSaveGame_Click()
On Error GoTo errh 'on the even of an error jump to label errh
 Dim FN$, i% 'dimensionalize FN as string type, i as integer type
  CD.Filter = "Yahtzee XP [*.yxp]|*.yxp|All Files [*.*]|*.*"
  'update Filter member(File Name/Extension pattern)
   CD.Flags = cdlOFNOverwritePrompt 'requests user confirmation to overwrite a file that allready exists
   CD.ShowSave 'show CD's ShowSave modal dialog
    If Dir(CD.FileName) <> "" Then
    'if the specified file allready exists, and since the user has allready confirmed that they desire the file to be overwritten, then purge it from disk
     Kill CD.FileName 'purge the specified file from disk
    End If
     FN$ = CD.FileName 'initialize FN with the filename of the selected file
      WriteToINI "GameInfo", "CurrentDraw", Str$(CurrentDraw), FN
      'see WriteToINI function for more info...
      'Save all the current game flags...
       WriteToINI "GameInfo", "NumberOfPlayers", Str$(NumOfPlayers), FN
        WriteToINI "GameInfo", "CurrentPlayer", Str$(CurrentPlayer), FN
         WriteToINI "GameInfo", "OldDice", Str$(ConvertBoolean(OldDice)), FN
          WriteToINI "GameInfo", "ChangePlayerflg", Str$(ChangePlayerflg), FN
           WriteToINI "GameInfo", "StatusLbl", lblstatus.Caption, FN
            WriteToINI "GameInfo", "ScoreBoxG1MP", Str$(txtScoreBoxG1.Item(0).MousePointer), FN
             WriteToINI "GameInfo", "ScoreBoxG2MP", Str$(txtScoreBoxG2.Item(0).MousePointer), FN
              WriteToINI "GameInfo", "ChoosingScoreBox", Str$(ConvertBoolean(ChoosingScoreBox)), FN

                WriteToINI "GameInfo", "Dice1FV", Str$(Dice1Stat.FaceValue), FN
                 WriteToINI "GameInfo", "Dice1H", Str$(ConvertBoolean(Dice1Stat.Holding)), FN
                  WriteToINI "GameInfo", "Dice2FV", Str$(Dice2Stat.FaceValue), FN
                   WriteToINI "GameInfo", "Dice2H", Str$(ConvertBoolean(Dice2Stat.Holding)), FN
                    WriteToINI "GameInfo", "Dice3FV", Str$(Dice3Stat.FaceValue), FN
                     WriteToINI "GameInfo", "Dice3H", Str$(ConvertBoolean(Dice3Stat.Holding)), FN
                      WriteToINI "GameInfo", "Dice4FV", Str$(Dice4Stat.FaceValue), FN
                       WriteToINI "GameInfo", "Dice4H", Str$(ConvertBoolean(Dice4Stat.Holding)), FN
                        WriteToINI "GameInfo", "Dice5FV", Str$(Dice5Stat.FaceValue), FN
                         WriteToINI "GameInfo", "Dice5H", Str$(ConvertBoolean(Dice5Stat.Holding)), FN

  
  For i = 1 To NumOfPlayers
   WriteToINI Str$(i), "PlayerName", PlayerScore(i).PlayerName, FN
    WriteToINI Str$(i), "Bonus", Str$(PlayerScore(i).Bonus), FN
     WriteToINI Str$(i), "Chance", Str$(PlayerScore(i).Chance), FN
      WriteToINI Str$(i), "Fives", Str$(PlayerScore(i).Fives), FN
       WriteToINI Str$(i), "FourKind", Str$(PlayerScore(i).FourKind), FN
        WriteToINI Str$(i), "Fours", Str$(PlayerScore(i).Fours), FN
         WriteToINI Str$(i), "FullHouse", Str$(PlayerScore(i).FullHouse), FN
          WriteToINI Str$(i), "HightStraight", Str$(PlayerScore(i).HightStraight), FN
           WriteToINI Str$(i), "LowStraight", Str$(PlayerScore(i).LowStraight), FN
            WriteToINI Str$(i), "Ones", Str$(PlayerScore(i).Ones), FN
             WriteToINI Str$(i), "Sixes", Str$(PlayerScore(i).Sixes), FN
              WriteToINI Str$(i), "ThreeKind", Str$(PlayerScore(i).ThreeKind), FN
               WriteToINI Str$(i), "Threes", Str$(PlayerScore(i).Threes), FN
                WriteToINI Str$(i), "Total", Str$(PlayerScore(i).Total), FN
                 WriteToINI Str$(i), "Twos", Str$(PlayerScore(i).Twos), FN
                  WriteToINI Str$(i), "uChance", Str$(ConvertBoolean(PlayerScore(i).uChance)), FN
                   WriteToINI Str$(i), "uFives", Str$(ConvertBoolean(PlayerScore(i).uFives)), FN
                    WriteToINI Str$(i), "uFourKind", Str$(ConvertBoolean(PlayerScore(i).uFourKind)), FN
                     WriteToINI Str$(i), "uFours", Str$(ConvertBoolean(PlayerScore(i).uFours)), FN
                      WriteToINI Str$(i), "uFullHouse", Str$(ConvertBoolean(PlayerScore(i).uFullHouse)), FN
                       WriteToINI Str$(i), "uHightStraight", Str$(ConvertBoolean(PlayerScore(i).uHightStraight)), FN
                        WriteToINI Str$(i), "uLowStraight", Str$(ConvertBoolean(PlayerScore(i).uLowStraight)), FN
                         WriteToINI Str$(i), "uOnes", Str$(ConvertBoolean(PlayerScore(i).uOnes)), FN
                          WriteToINI Str$(i), "uSixes", Str$(ConvertBoolean(PlayerScore(i).uSixes)), FN
                           WriteToINI Str$(i), "uThreeKind", Str$(ConvertBoolean(PlayerScore(i).uThreeKind)), FN
                            WriteToINI Str$(i), "uThrees", Str$(ConvertBoolean(PlayerScore(i).uThrees)), FN
                             WriteToINI Str$(i), "uTwos", Str$(ConvertBoolean(PlayerScore(i).uTwos)), FN
                              WriteToINI Str$(i), "uYahtzee", Str$(ConvertBoolean(PlayerScore(i).uYahtzee)), FN
                               WriteToINI Str$(i), "Yahtzee", Str$(PlayerScore(i).Yahtzee), FN
                                WriteToINI Str$(i), "YahtzeeOccur", Str$(PlayerScore(i).YahtzeeOccur), FN
  Next i
  'all of the game flags have now been copied to disk, somewhat like a game snapshot
   MsgBox "Game has been saved to" & vbCrLf & FN, vbInformation, "Game Saved"
   'inform user the file has been saved
    Exit Sub 'exit this procedure
errh: 'label errh(error handler)
  If Err.Number = 32755 Then Exit Sub 'if the current error number evaluates to 32755(User cancelled the ShowSave dialog either by pressing the Cancel button or by using the windows system menu)
   MsgBox "The following error has occured while saving the game." & vbCrLf & vbCrLf & Err.Description, vbCritical, "Error [" & Err.Number & "]"
   'inform user of the un-expected error
End Sub


Private Function ConvertBoolean(Value As Boolean) As Integer
 Select Case Value
  Case True: 'if value evaluates to true then return 1
   ConvertBoolean = 1
  Case False:
   ConvertBoolean = 0
 End Select
 'this function was designed to return an integer converted from a boolean value
 'since True can also evaluate to -1 since true means non-zero
End Function

Private Function Convert2Boolean(ByVal Value As String) As Boolean
'function returns a boolean value depending on the value of a string
 Value = LCase(Value) 'return the LowerCase version of the string
  Select Case Value
   Case "true": 'if Value evaluates to true
    Convert2Boolean = True 'return True
   Case "false": '...
    Convert2Boolean = False
  End Select
End Function

Private Sub mnuSound_Click()
 Load frmSndOpt 'load frmSndOpt into memory
  frmSndOpt.Show vbModal, Me 'show the dialog as a modal dialog
End Sub

Private Sub picLogo_Click()
If LogoFlag = 0 Then LogoFlag = 2
 If LogoFlag >= 3 Then LogoFlag = 1
  If LogoFlag = 1 Then
  'Determine the source bitmap of the Blend operation
   DBlend picLogo, tmplogo1
    'see DBlend function for more info...
    LogoFlag = LogoFlag + 1 'increment LogoFlag flag
  ElseIf LogoFlag = 2 Then
   DBlend picLogo, tmplogo2
    LogoFlag = LogoFlag + 1
  End If
End Sub

Private Sub txtBonus_GotFocus()
 HideCaret txtBonus.hwnd
 'call HideCaret function to hide the caret in the specified window
 'The caret represents the insertion point in a text box...
End Sub

Private Sub txtScoreBoxG1_Click(Index As Integer)
Dim CountingFlag%, DiceFaceValues(1 To 5) As Integer, i&, j&
'dimensionalize CountingFlag as integer data type, one dimensional array DiceFaceValues(6 elements) as integer type, i as long data type, j as long type
 DiceFaceValues(1) = Dice1Stat.FaceValue
  DiceFaceValues(2) = Dice2Stat.FaceValue
   DiceFaceValues(3) = Dice3Stat.FaceValue
    DiceFaceValues(4) = Dice4Stat.FaceValue
     DiceFaceValues(5) = Dice5Stat.FaceValue
     'initialize each element in the DiceFaceValues with the coinciding Dice face value
      If ChoosingScoreBox = False Then Exit Sub
      'if ChoosingScoreBox flag evaluates to false then exit this procedure...
       For i = 0 To txtScoreBoxG1.UBound
        txtScoreBoxG1(i).MousePointer = 1
        'enumerate through each text box control in the control array, setting each controls MousePointer property to 1(Default)
       Next i
        For i = 0 To txtScoreBoxG2.UBound
         txtScoreBoxG2(i).MousePointer = 1 '...
        Next i
        'NOTE: When a control in a control array is raising an event, the index paramater of that event specifies the controls index in the control array
        Select Case Index
         Case 0 'if index evaluates to 0 then...
         'ones score box
          If Not (txtScoreBoxG1(0).Text) = "" Then Exit Sub
          'if txtScoreBoxG1(0)[Ones Text box])'s text property does not evaluate to ""(null string) then exit this sub procedure
           CountingFlag = 0
            For i = 1 To 5
             If DiceFaceValues(i) = 1 Then CountingFlag = CountingFlag + 1
             'enumerate through each dice face value, for each dice face value which
             'evaluates to 1, increment CountingFlag by one
            Next i
             txtScoreBoxG1(0).Text = CStr(CountingFlag)
             'update the Ones score box textbox with the Ones Score(Sum of all Dice Face values whose value is 1)
              PlayerScore(CurrentPlayer).Ones = CountingFlag 'update virtual score box
               PlayerScore(CurrentPlayer).uOnes = True 'update virtual score boxes uOnes(Used Ones) member to true to prevent the user from using this score box again
                PlayerScore(CurrentPlayer).Total = PlayerScore(CurrentPlayer).Total + CountingFlag
                'calculate the users total score
                Say Str$(CountingFlag) & " points", True
                'see Say function for more info...
                 txtTotalScore.Text = PlayerScore(CurrentPlayer).Total
                 'update TotalScore textbox
                  EnableRollDiceBtn True 'see EnableRollDiceBtn for more info...
                   ChoosingScoreBox = False 'update action restriction flag
                    GoTo ExitSub 'jump to ExitSub label
       Case 1 'if index evaluates to 1 then..
       'twos score box
        If Not (txtScoreBoxG1(1).Text) = "" Then Exit Sub
         CountingFlag = 0
          For i = 1 To 5
           If DiceFaceValues(i) = 2 Then CountingFlag = CountingFlag + 1
          Next i
           CountingFlag = 2 * CountingFlag
            txtScoreBoxG1(1).Text = Str$(CountingFlag)
             PlayerScore(CurrentPlayer).Twos = CountingFlag
              PlayerScore(CurrentPlayer).uTwos = True
               PlayerScore(CurrentPlayer).Total = PlayerScore(CurrentPlayer).Total + CountingFlag
                Say Str$(CountingFlag) & " points", True
                 txtTotalScore.Text = PlayerScore(CurrentPlayer).Total
                  EnableRollDiceBtn True
                   ChoosingScoreBox = False
                    GoTo ExitSub
                    'see Case 0: remarks
       Case 2
       'threes score box
        If Not (txtScoreBoxG1(2).Text) = "" Then Exit Sub
         CountingFlag = 0
          For i = 1 To 5
           If DiceFaceValues(i) = 3 Then CountingFlag = CountingFlag + 1
          Next i
           CountingFlag = 3 * CountingFlag
            txtScoreBoxG1(2).Text = Str$(CountingFlag)
             PlayerScore(CurrentPlayer).Threes = CountingFlag
              PlayerScore(CurrentPlayer).uThrees = True
               PlayerScore(CurrentPlayer).Total = PlayerScore(CurrentPlayer).Total + CountingFlag
                Say Str$(CountingFlag) & " points", True
                 txtTotalScore.Text = PlayerScore(CurrentPlayer).Total
                  EnableRollDiceBtn True
                   ChoosingScoreBox = False
                    GoTo ExitSub
                    'see Case 0: remarks
       Case 3
       'fours score box
        If Not (txtScoreBoxG1(3).Text) = "" Then Exit Sub
         CountingFlag = 0
          For i = 1 To 5
           If DiceFaceValues(i) = 4 Then CountingFlag = CountingFlag + 1
          Next i
           CountingFlag = 4 * CountingFlag
            txtScoreBoxG1(3).Text = Str$(CountingFlag)
             PlayerScore(CurrentPlayer).Fours = CountingFlag
              PlayerScore(CurrentPlayer).uFours = True
               PlayerScore(CurrentPlayer).Total = PlayerScore(CurrentPlayer).Total + CountingFlag
                Say Str$(CountingFlag) & " points", True
                 txtTotalScore.Text = PlayerScore(CurrentPlayer).Total
                  EnableRollDiceBtn True
                   ChoosingScoreBox = False
                    GoTo ExitSub
                    'see Case 0: remarks
       Case 4
       'fives score box
        If Not (txtScoreBoxG1(4).Text) = "" Then Exit Sub
         CountingFlag = 0
          For i = 1 To 5
           If DiceFaceValues(i) = 5 Then CountingFlag = CountingFlag + 1
          Next i
           CountingFlag = 5 * CountingFlag
            txtScoreBoxG1(4).Text = Str$(CountingFlag)
             PlayerScore(CurrentPlayer).Fives = CountingFlag
              PlayerScore(CurrentPlayer).uFives = True
               PlayerScore(CurrentPlayer).Total = PlayerScore(CurrentPlayer).Total + CountingFlag
                Say Str$(CountingFlag) & " points", True
                 txtTotalScore.Text = PlayerScore(CurrentPlayer).Total
                  EnableRollDiceBtn True
                   ChoosingScoreBox = False
                    GoTo ExitSub
                    'see Case 0: remarks
       Case 5
       'sixes score box
        If Not (txtScoreBoxG1(5).Text) = "" Then Exit Sub
         CountingFlag = 0
          For i = 1 To 5
           If DiceFaceValues(i) = 6 Then CountingFlag = CountingFlag + 1
          Next i
           CountingFlag = 6 * CountingFlag
            txtScoreBoxG1(5).Text = Str$(CountingFlag)
             PlayerScore(CurrentPlayer).Sixes = CountingFlag
              PlayerScore(CurrentPlayer).uSixes = True
               PlayerScore(CurrentPlayer).Total = PlayerScore(CurrentPlayer).Total + CountingFlag
                Say Str$(CountingFlag) & " points", True
                 txtTotalScore.Text = PlayerScore(CurrentPlayer).Total
                  EnableRollDiceBtn True
                   ChoosingScoreBox = False
                    GoTo ExitSub
                    'see Case 0: remarks
      End Select 'escape select case statement
ExitSub: 'label ExitSub
 Dim onetosix As Boolean 'dimensionalize onetosize as boolean data type
  If PlayerScore(CurrentPlayer).uOnes = True And _
   PlayerScore(CurrentPlayer).uTwos = True And _
    PlayerScore(CurrentPlayer).uThrees = True And _
     PlayerScore(CurrentPlayer).uFours = True And _
      PlayerScore(CurrentPlayer).uFives = True And _
       PlayerScore(CurrentPlayer).uSixes = True Then onetosix = True
       'if each score box in the first score box group(ones to sixes)
       'has allready been used then initialize onetosize to true(non-zero)
  
        If onetosix = True Then 'if onetosix(Each score box in first score box group has been used) evaluates to true then...
         Dim Sum16%   '35 bonus for sum 1-6 [63]
         'dimensionalize Sum16 as integer data type
          If NumOfPlayers = 1 Then 'if NumOfPlayers evaluates to 1 then...
           Sum16 = Sum16 + PlayerScore(1).Ones
            Sum16 = Sum16 + PlayerScore(1).Twos
             Sum16 = Sum16 + PlayerScore(1).Threes
              Sum16 = Sum16 + PlayerScore(1).Fours
               Sum16 = Sum16 + PlayerScore(1).Fives
                Sum16 = Sum16 + PlayerScore(1).Sixes
                'calculate the sum of each of the score boxes value in the first score box group(ones to sixes)
                 If Sum16 >= 63 Then 'if the sum qualifies for the 35 point bonus then
                  PlayerScore(1).Bonus = PlayerScore(1).Bonus + 35
                   PlayerScore(1).Total = PlayerScore(1).Total + 35
                    txtTotalScore.Text = PlayerScore(1).Total
                     txtBonus.Text = PlayerScore(1).Bonus
                     'update the total score, and the bonus score
                      Say "You have been awarded a bonus of 35 points"
                      'see Say function for more information...
                 End If
          Else
           For i = 1 To NumOfPlayers
            Sum16 = Sum16 + PlayerScore(i).Ones
             Sum16 = Sum16 + PlayerScore(i).Twos
              Sum16 = Sum16 + PlayerScore(i).Threes
               Sum16 = Sum16 + PlayerScore(i).Fours
                Sum16 = Sum16 + PlayerScore(i).Fives
                 Sum16 = Sum16 + PlayerScore(i).Sixes
                  If Sum16 >= 63 Then
                   PlayerScore(i).Bonus = PlayerScore(i).Bonus + 35
                    PlayerScore(i).Total = PlayerScore(i).Total + 35
                     txtTotalScore.Text = PlayerScore(i).Total
                      txtBonus.Text = PlayerScore(i).Bonus
                       Say PlayerScore(i).PlayerName
                        Say " has been awarded a bonus of 35 points"
                  End If '...
                   Sum16 = 0 'reset Sum16 to zero
           Next i
          End If            '//35 bonus for sum 1-6 [63]
        End If
         CheckForEndOfGame 'see CheckForEndOfGame function for more info...
          If ChangePlayerflg = 1 Then ChangePlayerflg = 2
          'conditionally increment changePlayerflg flag
            Dim tmpCurrentPlayer% 'dimensionalize tmpCurrentPlayer as integer type
             If NumOfPlayers > 1 Then 'if there is more than one players
              tmpCurrentPlayer = CurrentPlayer + 1 'initialize tmpCurrentPlayer with the incremented product of CurrentPlayer
               If tmpCurrentPlayer > NumOfPlayers Then tmpCurrentPlayer = 1
               'if tmpCurrentPlayer evaluates greater than the number of players then set tmpCurrentplayer to 1
                Say PlayerScore(tmpCurrentPlayer).PlayerName
                 Say ", please roll the dice"
                 'see Say function for more info...
             End If
              lblstatus.Caption = "Please Roll Dice!!"
              'update status label's caption
End Sub

Private Sub txtScoreBoxG1_GotFocus(Index As Integer)
 HideCaret txtScoreBoxG1(Index).hwnd
 'hide the caret(insertion point representitive) in the specified textbox
End Sub


Private Sub txtScoreBoxG2_Click(Index As Integer)
Dim LowNum%, LowNumtmp%, LowNumInd%, CurNum%, NumOfStraight%
'dimensionalize LowNum as integer data type, LowNumtmp as integer type, LowNumInd as integer type, CurNum as integer data type, NumOfStraight as integer data type
Dim CountingFlag%, tmpFlag% 'dimensionalize CountingFlag as integer data type, tmpFlag as integer type
Dim DiceFaceValues(1 To 5) As Integer 'dimensionalize DiceFaceValues as a one dimensional array(6 elements) as integer data type
Dim i&, j& 'dimensionalize i and j as long data type
 DiceFaceValues(1) = Dice1Stat.FaceValue
  DiceFaceValues(2) = Dice2Stat.FaceValue
   DiceFaceValues(3) = Dice3Stat.FaceValue
    DiceFaceValues(4) = Dice4Stat.FaceValue
     DiceFaceValues(5) = Dice5Stat.FaceValue
     'initialize each element in the DiceFaceValues array to the coinciding Dice Face value...

      If ChoosingScoreBox = False Then Exit Sub
       For i = 0 To txtScoreBoxG1.UBound
        txtScoreBoxG1(i).MousePointer = 1
        'set each control's MousePointer to 1(default) in the control array
       Next i
        For i = 0 To txtScoreBoxG2.UBound
         txtScoreBoxG2(i).MousePointer = 1 '...
        Next i
        'note: index specifies the control index in the control array who raised this event
         Select Case Index
          Case 0 'if index evaluates to 0
          'three of a kind score box
           If Not (txtScoreBoxG2(0).Text) = "" Then Exit Sub
           'if the score box has allready been used then exit this sub procedure
            CountingFlag = 0 'initialize CountingFlag to zero
             For i = 1 To 5
              For j = 1 To 5
               If DiceFaceValues(i) = DiceFaceValues(j) Then CountingFlag = CountingFlag + 1
               'Element index i in array DiceFaceValues is being compared to each of the five dice face values including its self
               'if 3 of the comparisons returned true then there are three dice with the same face values
              Next j
              If CountingFlag >= 3 Then Exit For 'Counts its own occurence
              'if the of the same face values were found then exit this loop
               CountingFlag = 0 'reset CountingFlag to zero
             Next i
              If CountingFlag < 3 Then
               CountingFlag = 0
              Else
              'if three of a kind was found...
               CountingFlag = 0
                For i = 1 To 5
                 CountingFlag = CountingFlag + DiceFaceValues(i)
                 'calculate the sum of all the dice face values
                Next i
              End If
            
               txtScoreBoxG2(0).Text = Str$(CountingFlag)
                PlayerScore(CurrentPlayer).ThreeKind = CountingFlag
                 PlayerScore(CurrentPlayer).uThreeKind = True
                  PlayerScore(CurrentPlayer).Total = PlayerScore(CurrentPlayer).Total + CountingFlag
                  'update score information...
                   Say Str$(CountingFlag) & " points", True 'see Say for more info...
                    txtTotalScore.Text = PlayerScore(CurrentPlayer).Total
                     EnableRollDiceBtn True 'see EnableRollDiceBtn for more info...
                      ChoosingScoreBox = False 'update action restriction flag...
                       GoTo ExitSub 'jump to ExitSub label
            
        Case 1 'if index evaluates to 1 then....
         'four of a kind score box
         If Not (txtScoreBoxG2(1).Text) = "" Then Exit Sub
         'if the score box has allready been used then exit this sub procedure
          CountingFlag = 0 'initialize CountingFlag to zero
          
          'SEE remarks for Case 0 for more details...
          
           For i = 1 To 5
            For j = 1 To 5
             If DiceFaceValues(i) = DiceFaceValues(j) Then CountingFlag = CountingFlag + 1
            Next j
             If CountingFlag >= 4 Then Exit For 'Counts its own occurence
              CountingFlag = 0
           Next i
            If CountingFlag < 4 Then
             CountingFlag = 0
            Else
             CountingFlag = 0
              For i = 1 To 5
               CountingFlag = CountingFlag + DiceFaceValues(i)
              Next i
            End If
             txtScoreBoxG2(1).Text = Str$(CountingFlag)
              PlayerScore(CurrentPlayer).FourKind = CountingFlag
               PlayerScore(CurrentPlayer).uFourKind = True
                PlayerScore(CurrentPlayer).Total = PlayerScore(CurrentPlayer).Total + CountingFlag
                 Say Str$(CountingFlag) & " points", True
                  txtTotalScore.Text = PlayerScore(CurrentPlayer).Total
                   EnableRollDiceBtn True
                    ChoosingScoreBox = False
                     GoTo ExitSub
            
       Case 2: 'if index evaluates to 2 then...
       'the algorithm which determines if each of the 5 dice represent a full house(where three of the dice's face values are identical, and the remaining two dice face values are also the same)
       Dim threek As Boolean, twok As Boolean, threekindnum%
       'dimensionalize threek(three of a kind) and twok(two of a kind) as boolean
       'this algorithm will search for three of a kind, as in the previous case
       'it will then search for two of a kind, of both of these flags are set to true by the algorithm then a full house has been found
        'full house score box
        If Not (txtScoreBoxG2(2).Text) = "" Then Exit Sub
         If RetYahtzee = True Then CountingFlag = 25: GoTo yahtzee1
          CountingFlag = 0
           For i = 1 To 5
            For j = 1 To 5
             If DiceFaceValues(i) = DiceFaceValues(j) Then CountingFlag = CountingFlag + 1
            Next j
             If CountingFlag >= 3 Then threekindnum = DiceFaceValues(i): Exit For 'Counts its own occurence
              CountingFlag = 0
           Next i
           If CountingFlag < 3 Then
            CountingFlag = 0
           Else
            threek = True
           End If
            CountingFlag = 0
             For i = 1 To 5
              For j = 1 To 5
               If DiceFaceValues(i) = threekindnum Then GoTo skfh2
                If DiceFaceValues(i) = DiceFaceValues(j) Then CountingFlag = CountingFlag + 1
               Next j
                If CountingFlag >= 2 Then threekindnum = DiceFaceValues(i): Exit For 'Counts its own occurence
skfh2:
                 CountingFlag = 0
             Next i
              If CountingFlag < 2 Then
                CountingFlag = 0
              Else
               twok = True
              End If

               If threek = True And twok = True Then
                 CountingFlag = 25
               Else
                CountingFlag = 0
               End If
yahtzee1:
                txtScoreBoxG2(2).Text = Str$(CountingFlag)
                 PlayerScore(CurrentPlayer).FullHouse = CountingFlag
                  PlayerScore(CurrentPlayer).uFullHouse = True
                   PlayerScore(CurrentPlayer).Total = PlayerScore(CurrentPlayer).Total + CountingFlag
                    Say Str$(CountingFlag) & " points", True
                     txtTotalScore.Text = PlayerScore(CurrentPlayer).Total
                      EnableRollDiceBtn True
                       ChoosingScoreBox = False
                        GoTo ExitSub
     
     
        Case 3: 'if index evaluates to 3 then...
        'Low Straight is where 4 of the 5 dice's face values are in
        'chronological order starting at the lowest number in the
        'sequence(1,2,3,4 or 2,3,4,5 or 3,4,5,6), the remainding dice
        'face value is irrelevant
        
        'low straight score box
        If Not (txtScoreBoxG2(3).Text) = "" Then Exit Sub
         CountingFlag = 0 'DiceFaceValues(i)
          If RetYahtzee = True Then CountingFlag = 30: GoTo yahtzee2
           LowNum = DiceFaceValues(1): LowNumInd = 1
            For i = 2 To 5
             CurNum = DiceFaceValues(i)
              If LowNum > CurNum Then
               LowNum = CurNum: LowNumInd = i
              End If
            Next i
          '1+
          For i = 1 To 5
           If Not (i = LowNumInd) Then ' not same dice as lownum
            LowNumtmp = DiceFaceValues(i)
             If (LowNumtmp - 1) = LowNum Then
              NumOfStraight = NumOfStraight + 1
               Exit For
             End If
           End If
          Next i
          
          If NumOfStraight < 1 Then GoTo HIGHNUM
          
          '2+
          For i = 1 To 5
           If Not (i = LowNumInd) Then ' not same dice as lownum
            LowNumtmp = DiceFaceValues(i)
             If (LowNumtmp - 2) = LowNum Then
              NumOfStraight = NumOfStraight + 1
               Exit For
             End If
           End If
          Next i
          
          '3+
          For i = 1 To 5
           If Not (i = LowNumInd) Then ' not same dice as lownum
            LowNumtmp = DiceFaceValues(i)
             If (LowNumtmp - 3) = LowNum Then
              NumOfStraight = NumOfStraight + 1
               Exit For
             End If
           End If
          Next i
          
          
          If NumOfStraight >= 3 Then
           CountingFlag = 30
          Else
          
          'SECOND HIGH NUM
HIGHNUM:
          
         NumOfStraight = 0
         LowNum = DiceFaceValues(1): LowNumInd = 1
          For i = 1 To 5
           CurNum = DiceFaceValues(i)
            If LowNum < CurNum Then
             LowNum = CurNum: LowNumInd = i
            End If
          Next i
          
          '1+
          For i = 1 To 5
           If Not (i = LowNumInd) Then ' not same dice as highnum
            LowNumtmp = DiceFaceValues(i)
             If LowNumtmp = (LowNum - 1) Then
              NumOfStraight = NumOfStraight + 1
               Exit For
             End If
           End If
          Next i
          
          
          '2+
          For i = 1 To 5
           If Not (i = LowNumInd) Then ' not same dice as lownum
            LowNumtmp = DiceFaceValues(i)
             If LowNumtmp = (LowNum - 2) Then
              NumOfStraight = NumOfStraight + 1
               Exit For
             End If
           End If
          Next i
          
          '3+
          For i = 1 To 5
           If Not (i = LowNumInd) Then ' not same dice as lownum
            LowNumtmp = DiceFaceValues(i)
             If LowNumtmp = (LowNum - 3) Then
              NumOfStraight = NumOfStraight + 1
               Exit For
             End If
           End If
          Next i
          
          
          If NumOfStraight >= 3 Then
           CountingFlag = 30
          Else
           CountingFlag = 0
          End If
        End If
          
yahtzee2:
           txtScoreBoxG2(3).Text = Str$(CountingFlag)
            PlayerScore(CurrentPlayer).LowStraight = CountingFlag
             PlayerScore(CurrentPlayer).uLowStraight = True
              PlayerScore(CurrentPlayer).Total = PlayerScore(CurrentPlayer).Total + CountingFlag
               Say Str$(CountingFlag) & " points", True
                txtTotalScore.Text = PlayerScore(CurrentPlayer).Total
                 EnableRollDiceBtn True
                  ChoosingScoreBox = False
                   GoTo ExitSub


        Case 4: 'if index evaluates to 4 then...
        'high straight is where each of the dice's face values are
        'in chronological order (1,2,3,4,5 or 2,3,4,5,6)
        
        'high straight score box
        If Not (txtScoreBoxG2(4).Text) = "" Then Exit Sub
         CountingFlag = 0
          If RetYahtzee = True Then CountingFlag = 40: GoTo yahtzee3
          
          LowNum = DiceFaceValues(1): LowNumInd = 1
          For i = 2 To 5
           CurNum = DiceFaceValues(i)
            If LowNum > CurNum Then
             LowNum = CurNum: LowNumInd = i
            End If
          Next i
          
          '1+
          For i = 1 To 5
           If Not (i = LowNumInd) Then ' not same dice as lownum
            LowNumtmp = DiceFaceValues(i)
             If (LowNumtmp - 1) = LowNum Then
              NumOfStraight = NumOfStraight + 1
               Exit For
             End If
           End If
          Next i
          
          
          '2+
          For i = 1 To 5
           If Not (i = LowNumInd) Then ' not same dice as lownum
            LowNumtmp = DiceFaceValues(i)
             If (LowNumtmp - 2) = LowNum Then
              NumOfStraight = NumOfStraight + 1
               Exit For
             End If
           End If
          Next i
          
          '3+
          For i = 1 To 5
           If Not (i = LowNumInd) Then ' not same dice as lownum
            LowNumtmp = DiceFaceValues(i)
             If (LowNumtmp - 3) = LowNum Then
              NumOfStraight = NumOfStraight + 1
               Exit For
             End If
           End If
          Next i
          
          
          '4+
          For i = 1 To 5
           If Not (i = LowNumInd) Then ' not same dice as lownum
            LowNumtmp = DiceFaceValues(i)
             If (LowNumtmp - 4) = LowNum Then
              NumOfStraight = NumOfStraight + 1
               Exit For
             End If
           End If
          Next i
                  
          If NumOfStraight >= 4 Then
           CountingFlag = 40
          Else
           CountingFlag = 0
          End If
yahtzee3:
    
           txtScoreBoxG2(4).Text = Str$(CountingFlag)
            PlayerScore(CurrentPlayer).HightStraight = CountingFlag
             PlayerScore(CurrentPlayer).uHightStraight = True
              PlayerScore(CurrentPlayer).Total = PlayerScore(CurrentPlayer).Total + CountingFlag
               Say Str$(CountingFlag) & " points", True
                txtTotalScore.Text = PlayerScore(CurrentPlayer).Total
                 EnableRollDiceBtn True
                  ChoosingScoreBox = False
                   GoTo ExitSub
    
      Case 5: 'if index evaluates to 5...
      'Yahtzee is where each of the 5 dice face values are identical,
      '(1,1,1,1,1 or 4,4,4,4,4 ect...)
      
      
      'Yahtzee score box
       If Not (txtScoreBoxG2(5).Text) = "" Then Exit Sub
        CountingFlag = 0: CurNum = DiceFaceValues(1)
         For i = 1 To 5
          If DiceFaceValues(i) = CurNum Then CountingFlag = CountingFlag + 1
         Next i
          If CountingFlag < 5 Then
           CountingFlag = 0
          Else
           CountingFlag = 50
          End If
        
           txtScoreBoxG2(5).Text = Str$(CountingFlag)
            PlayerScore(CurrentPlayer).Yahtzee = CountingFlag
             PlayerScore(CurrentPlayer).uYahtzee = True
             If CountingFlag = 50 Then PlayerScore(CurrentPlayer).YahtzeeOccur = PlayerScore(CurrentPlayer).YahtzeeOccur + 1
              PlayerScore(CurrentPlayer).Total = PlayerScore(CurrentPlayer).Total + CountingFlag
               Say Str$(CountingFlag) & " points", True
                txtTotalScore.Text = PlayerScore(CurrentPlayer).Total
                 EnableRollDiceBtn True
                  ChoosingScoreBox = False
                   GoTo ExitSub
   
        Case 6: 'if index evaluates to 6
        'chance, the sum of every dice face value is calculated and added to the total score
        
        'Chance score box
         If Not (txtScoreBoxG2(6).Text) = "" Then Exit Sub
           For i = 1 To 5
            CountingFlag = CountingFlag + DiceFaceValues(i)
           Next i
  
            txtScoreBoxG2(6).Text = Str$(CountingFlag)
             PlayerScore(CurrentPlayer).Chance = CountingFlag
              PlayerScore(CurrentPlayer).uChance = True
               PlayerScore(CurrentPlayer).Total = PlayerScore(CurrentPlayer).Total + CountingFlag
                Say Str$(CountingFlag) & " points", True
                 txtTotalScore.Text = PlayerScore(CurrentPlayer).Total
                  EnableRollDiceBtn True
                   ChoosingScoreBox = False
                    GoTo ExitSub
  
  
    End Select 'escape select case statement
ExitSub: 'label ExitSub
 CheckForEndOfGame 'see CheckForEndOfGame function for more info...
  If ChangePlayerflg = 1 Then ChangePlayerflg = 2
  'conditionally increment ChangePlayerflg by one
   Dim tmpCurrentPlayer% 'dimensionalize tmpCurrentPlayer as integer data type
    If NumOfPlayers > 1 Then 'if there is more than one player then...
     tmpCurrentPlayer = CurrentPlayer + 1 'initialize tmpCurrentPlayer with the incremented product of CurrentPlayer
      If tmpCurrentPlayer > NumOfPlayers Then tmpCurrentPlayer = 1
      'if tmpCurrentPlayer evaluates to greater than the number of players then reset tmpCurrentPlayer to 1
       Say PlayerScore(tmpCurrentPlayer).PlayerName
        Say ", please roll the dice"
         'see Say function for more info...
    End If
     lblstatus.Caption = "Please Roll Dice!!"
     'update the status labels caption
End Sub

Private Sub txtScoreBoxG2_GotFocus(Index As Integer)
 HideCaret txtScoreBoxG2(Index).hwnd
 'hide the caret(insertion point representitive in input windows)
End Sub

Private Sub txtTotalScore_GotFocus()
 HideCaret txtTotalScore.hwnd '...
End Sub

Public Sub CheckForYahtzee()
Dim DiceFaceValues(1 To 5) As Integer, CountingFlag%, CurNum%, i&, j&
'dimensionalize one dimensional array DiceFaceValues with 6 elements as integer data type,
'CountingFlag as integer data type, CurNum as integer data type, i as long data type, j as long data type

 DiceFaceValues(1) = Dice1Stat.FaceValue
  DiceFaceValues(2) = Dice2Stat.FaceValue
   DiceFaceValues(3) = Dice3Stat.FaceValue
    DiceFaceValues(4) = Dice4Stat.FaceValue
     DiceFaceValues(5) = Dice5Stat.FaceValue
     'initialize each element in the DiceFaceValues array with the coinciding Dice face value
      
      'Query the exisistance of a yahtzee(5 of a kind)
      CountingFlag = 0: CurNum = DiceFaceValues(1)
       For i = 1 To 5
        If DiceFaceValues(i) = CurNum Then CountingFlag = CountingFlag + 1
       Next i
        If CountingFlag < 5 Then
         CountingFlag = 0
        Else
         If PlayerScore(CurrentPlayer).YahtzeeOccur > 0 Then
          txtBonus.Text = PlayerScore(CurrentPlayer).Bonus + 100
           PlayerScore(CurrentPlayer).Bonus = PlayerScore(CurrentPlayer).Bonus + 100
            PlayerScore(CurrentPlayer).YahtzeeOccur = PlayerScore(CurrentPlayer).YahtzeeOccur + 1
             PlayerScore(CurrentPlayer).Total = PlayerScore(CurrentPlayer).Total + PlayerScore(CurrentPlayer).Bonus
              txtTotalScore.Text = PlayerScore(CurrentPlayer).Total
               If NumOfPlayers = 1 Then
                Say "You have been awarded a bonus of 100"
               Else
                Say PlayerScore(CurrentPlayer).PlayerName
                 Say " has been awarded a bonus of 100 points"
               End If
        End If
         Load frmYahtzee 'load dialog frmYahtzee
          frmYahtzee.Show vbModal, Me 'show the dialog as a modal dialog
           Say "You rolled a Yahtzey"
            'see Say function for more info...
       End If
End Sub

Private Function RetYahtzee() As Boolean
Dim DiceFaceValues(1 To 5) As Integer, CountingFlag%, CurNum%, i&, j&
 DiceFaceValues(1) = Dice1Stat.FaceValue
  DiceFaceValues(2) = Dice2Stat.FaceValue
   DiceFaceValues(3) = Dice3Stat.FaceValue
    DiceFaceValues(4) = Dice4Stat.FaceValue
     DiceFaceValues(5) = Dice5Stat.FaceValue
     'initialize each element in the DiceFaceValues array to the coinciding dice face value
     
      CountingFlag = 0: CurNum = DiceFaceValues(1)
      'initialize CountingFlag to zero, CurNum to the first dice face value
       For i = 1 To 5
       'for next loop, initialize i to 1, loops until i evaluates to 5 incrementing i by one each iteration(since the step keyword is omitting, 1 is assumed)
        If DiceFaceValues(i) = CurNum Then CountingFlag = CountingFlag + 1
       Next i
        If CountingFlag < 5 Then
        'if countingflag is less than 5 then return true as no yahtzee has been discovered
         RetYahtzee = False
        Else
         RetYahtzee = True 'the dice represent a yahtzee...
        End If
End Function




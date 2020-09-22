Attribute VB_Name = "globals"
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
Public Declare Function InitCommonControls Lib "Comctl32.dll" () As Long


Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long



Public Const GWL_EXSTYLE = (-20)
Public Const WS_EX_LAYERED = &H80000
Public Const WS_EX_TRANSPARENT = &H20&
Public Const LWA_ALPHA = &H2&

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long


Global wndTransSpeed As Integer
Global wndTrans As Boolean
Global GraphTrans As Boolean


Public Declare Function AlphaBlend _
  Lib "msimg32" ( _
  ByVal hDestDC As Long, _
  ByVal x As Long, ByVal y As Long, _
  ByVal nWidth As Long, _
  ByVal nHeight As Long, _
  ByVal hSrcDC As Long, _
  ByVal xSrc As Long, _
  ByVal ySrc As Long, _
  ByVal widthSrc As Long, _
  ByVal heightSrc As Long, _
  ByVal dreamAKA As Long) _
  As Boolean


Public Declare Sub RtlMoveMemory Lib "kernel32.dll" (Destination As Any, Source As Any, ByVal Length As Long)
Public Type BLENDFUNCTION
  BlendOp As Byte
  BlendFlags As Byte
  SourceConstantAlpha As Byte
  AlphaFormat As Byte
End Type
Public Const AC_SRC_OVER = &H0


Global PrefferedName$


Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function HideCaret Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Public Const SND_FILENAME = &H20000
Public Const SND_ASYNC = &H1

Global PlaySounds As Boolean

Public Type DiceCol
 DiceColor As DicCol
End Type

Public Enum DicCol
 Blue = 1
 Red = 2
 Green = 3
 Gold = 4
End Enum

Global CurDiceCol As DiceCol

Public Type DiceStatistics
 Holding As Boolean
 FaceValue As Integer
End Type


Public RollDiceBtnDis As Boolean

Public CurrentDraw As Integer '3 max
Public OldDice As Boolean

Public ChoosingScoreBox As Boolean

Public NumOfPlayers As Integer
Public CurrentPlayer As Integer
Public PlayerScore(1 To 6) As ScoreInfo

Public Type ScoreInfo
'score box group 1
 Ones As Integer
 uOnes As Boolean
 
 Twos As Integer
 uTwos As Boolean
 
 Threes As Integer
 uThrees As Boolean
 
 Fours As Integer
 uFours As Boolean
 
 
 Fives As Integer
 uFives As Boolean
 
 Sixes As Integer
 uSixes As Boolean
 
 
  'score box group 2
   ThreeKind As Integer
   uThreeKind As Boolean
   
   FourKind As Integer
   uFourKind As Boolean
   
   FullHouse As Integer
   uFullHouse As Boolean
   
   LowStraight As Integer
   uLowStraight As Boolean
   
   HightStraight As Integer
   uHightStraight As Boolean
   
   Yahtzee As Integer
   uYahtzee As Boolean
   
   YahtzeeOccur As Integer
   
   Chance As Integer
   uChance As Boolean
   
     Total As Integer
     Bonus As Integer
      '++info
       PlayerName As String
End Type

Public OldHighs(1 To 10) As HighScoreInfo
Public ScoresToAdd(1 To 6) As HighScoreInfo

Public Type HighScoreInfo
 PlayerName As String
 TotalScore As Integer
 Date As String
End Type


Public Sub LoadHighScores()
Dim i%, istr$ 'dimensionalize i as integer type, istr as string type
 For i = 1 To 10
 'for next loop, initialize i to 1, loop until i evaluates to 10 incrementing i by one each iteration
  istr = CStr(i) 'initialize istr with the string conversion of integer i
   OldHighs(i).TotalScore = Int(GetSetting("YAH", "SC", istr, "0"))
    OldHighs(i).PlayerName = GetSetting("YAH", "SC", istr & "n", "Not Set")
     OldHighs(i).Date = GetSetting("YAH", "SC", istr & "d", "Not Set")
     'enumerate through each high score in the registry, and store them in the OldHighs array
 Next i 'increment i, evaluate loop conditions(1<=10), next iteration
End Sub


Public Sub AddHighScores()
Dim tmpHighScores(1 To 10) As HighScoreInfo
Dim i%, j%, NewSpot%, istr$, k%, l%
'dimensionalize one dimensional array tmpHighScores(10 elements) as HighScoreInfo structure
'dimensionalize i, j, NewSpot(New Position), k and l as integer data types, istr as string data type
 For i = 1 To NumOfPlayers
 'loop until i evaluates to NumOfPlayers+1
 'enumerate through each new score to add
  For j = 1 To 10
  'loop until j evaluates to 10+1
  'enumerate through each high score comparing it to the current score to add
   If ScoresToAdd(i).TotalScore > OldHighs(j).TotalScore Then
   'if the score to add specified by its element index in the ScoresToAdd Array
   'is greater than the High Score specified by its element index in the OldHighs array
   'then...
     NewSpot = j
     'initialize NewSpot to the value of j
      For k = NewSpot To 10
      'initialize k to NewSpot(New Score Insertion Position), loop until k evaluates to 10
       tmpHighScores(k) = OldHighs(k)
      Next k
       OldHighs(NewSpot).PlayerName = ScoresToAdd(i).PlayerName
        OldHighs(NewSpot).TotalScore = ScoresToAdd(i).TotalScore
         OldHighs(NewSpot).Date = Date
          For k = NewSpot To 9
           OldHighs(k + 1) = tmpHighScores(k)
          Next k
           Exit For
    'shift the new score into the High Scores Array, preserving the higher High Scores,
    'insert the new high score, replace the lower scores below the new high score removing the lowest score
    
    'Initial High Scores (only 5 are show in this example)
    '1 - 500
    '2 - 456
    '3 - 355
    '4 - 250
    '5 - 100

    'New Score to be added : 260
    
    'New High Scores Array
    '1 - 500
    '2 - 456
    '3 - 255
    '4 - 260  NEW SCORE
    
    'Temporary High Scores
    '1 - 250
    '2 - 100
    
    'Shift every Temporary High Score minus 1 to the New High Scores Array
    'New High Scores Array
    '1 - 500
    '2 - 456
    '3 - 255
    '4 - 260  NEW SCORE
    '5 - 250
    
   End If
  Next j
 Next i
End Sub

Public Sub SaveHighScores()
Dim i%, istr$
'dimensionalize i as integer data type, istr as string data type
 For i = 1 To 10
  istr = CStr(i)
   SaveSetting "YAH", "SC", istr, Str$(OldHighs(i).TotalScore)
    SaveSetting "YAH", "SC", istr & "n", OldHighs(i).PlayerName
     SaveSetting "YAH", "SC", istr & "d", OldHighs(i).Date
     'enumerate through each high score, and save it to the registry
 Next i
End Sub

Public Function UserName() As String
Dim buffer As String * 256, bufferLen&, ret&
'dimensionalize buffer as string data type[* byte memory allocation],
'bufferlen as long data type, r as long data type
 bufferLen& = Len(buffer) 'initialize bufferlen with the length of variable buffer
  ret& = GetUserName(buffer, bufferLen&)
  'initialize ret with the return of GetUserName
  'GetUserName function copies the length specified by bufferLen of the
  'current user name to variable buffer
   If ret& = 0 Or bufferLen& = 0 Then Exit Function
   'if ret(return) evaluates to zero or bufferLen evaluates to zero then exit this procedure
   'note: the GetUserName function also initialize bufferLen to the ammount of characters copied...
    UserName = Trim$(Left$(buffer, bufferLen& - 1))
    'return the username
    'Trim version removes the leading and trailing white-spaces of a string
    'Left$(String Return) function returns the specified amount of characters from Left to Right of the specified string
    'In this case, the left function is used to return only the amount of characters of buffer which were copied by the GetUserName function
End Function

Public Sub Say(sWhat As String, Optional reset As Boolean = False)
On Error Resume Next 'on the event of an error resume next
 If frmMain.dss Is Nothing Then Exit Sub
 'if frmMain's dss(DirectSS class) has not been initialized then exit this sub procedure
  If reset = True Then frmMain.dss.AudioReset
  'if reset paramater evaluates to true then call dss's AudioReset method to purge any qeued audio output operations including any operation it's currently performing
   frmMain.dss.Speak sWhat 'call dss's Speak method to 'say' the specified text
   'for documentation on Learnout & Haupsie's TruVoice Speech Synthesization engine refer to the Speech SDK available for download from MSDN Online(http://msdn.microsoft.com)
End Sub

Public Function WriteToINI(Section$, Entry$, Value$, FileName$)
'Calling WritePrivateProfileString directly would also suffice, however since I wish
'to leave remarks on this function I have inserted this function prototype...
 WritePrivateProfileString Section, Entry, Value, FileName
 'the WritePrivateProfileString function saves strings to specific keys of specific sections to a windows initialization file(INI)
 'This API function is only provided for compatibility on 16-bit versions of windows, however,
 'since we are using this file structure to save and load games for better organization and portability we will use it anyway
 'I suggest that instead of adopting this method of saving data, you save settings in the registry instead for you're own applications...
 
  'Syntax of INI files:
  '[section1]
  ' key1 = string
  ' key2 = string
  '[section2]
  ' key 1 = string
  '....
'For more detailed documentation on Windows Initialization file refer to the following on-line document:
'http://msdn.microsoft.com/library/en-us/sysinfo/base/writeprivateprofilestring.asp?frame=true
End Function


Public Function ReadFromINI(Section$, Entry$, FileName$) As String
Dim buffer$, buflen& 'dimensionalize buffer as string data type, buflen as long data type
 buffer$ = Space(500) 'initialize buffer with the string returned by Space function
 'Space function returns the specified amount of spaces
 'NOTE: That 500 is not a special length, but this number was used because no more than 500
 'characters would ever have to be retrieved by this application...
  buflen& = Len(buffer$) 'initialize buflen with the length of the variable buffer
   GetPrivateProfileString Section, Entry, vbNullString, buffer$, buflen&, FileName
   'retreive the specified key's value of the specified section
    ReadFromINI = Trim(buffer$) 'return buffer
    'Trim function removes leading and trailing spaces...
End Function

Public Sub TransWin(Handle As Long, Optional Inc As Boolean = True)
On Error Resume Next 'on the event of an error resume execution on the next line
 If wndTrans = False Then Exit Sub
 'if wndTrans flag evaluates to false then exit sub
 Dim i& 'dimensionalize i as long data type
  If Inc = True Then
  'if paramater Inc evaluates to true then...
   For i = 0 To 255 Step wndTransSpeed
   'The step keyword in the For statement specifies the amount to increase the Flag variable(i) each iteration, negative values are also supported
    DoEvents 'yeild execution to other asynchronously processing procedures
     SetLayeredWindowAttributes Handle, 0, i, LWA_ALPHA
     'Set the opacity(specified by i) of this window
   Next i
    SetLayeredWindowAttributes Handle, 0, 255, LWA_ALPHA
    '...
  Else
   For i = 255 To 0 Step -wndTransSpeed '...
    DoEvents '...
     SetLayeredWindowAttributes Handle, 0, IIf(i > 0, i, 0), LWA_ALPHA '...
   Next i '..
  End If
End Sub

Public Sub TransPrep(Handle As Long)
On Error Resume Next 'on the event of an error resume execution on the next line
Dim NormalWindowStyle& 'dimensionalize NormalWindowStyle as long type
 If wndTrans = False Then Exit Sub 'if wndTrans flag evaluates to false then exit this procedure
  NormalWindowStyle = GetWindowLong(Handle, GWL_EXSTYLE)
  'initialize NormalWindowStyle with the current extended window style of the window specified by its window handle
   SetWindowLong Handle, GWL_EXSTYLE, NormalWindowStyle Or WS_EX_LAYERED
   'set the windows new extended window style, the new style passed is the product of the initial window style on which a logical OR[operand B = WS_EX_LAYERED constant] operation is performed
    SetLayeredWindowAttributes Handle, 0, 0, LWA_ALPHA
    'set the window opacity to 100% transparent
End Sub

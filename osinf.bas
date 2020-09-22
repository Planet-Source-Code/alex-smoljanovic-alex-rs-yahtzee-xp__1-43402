Attribute VB_Name = "osinf"
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



Private Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Declare Function GetVersionEx& Lib "kernel32" Alias _
    "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) 'As Long
    Private Const VER_PLATFORM_WIN32_NT = 2
    Private Const VER_PLATFORM_WIN32_WINDOWS = 1
    Private Const VER_PLATFORM_WIN32s = 0


Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128 'Maintenance string For PSS usage
End Type


Public Function RetOSInf(Optional ByRef oStr$) As Boolean
On Error GoTo errh
'on the event of an error jump to label errh
Dim osvi As OSVERSIONINFO: osvi.dwOSVersionInfoSize = 148 'initialize variable
 If GetVersionEx(osvi) <> 0 Then
 'if the function returned succesfully then...
   Select Case osvi.dwPlatformId
    Case VER_PLATFORM_WIN32s: 'if osvi(operating system version information)'s dwPlatformID evaluates to the value of the VER_PLATFORM_WIN32s constant then...
     RetOSInf = False 'return false(incompatible OS)
      oStr = "Win32s" 'update output
    Case VER_PLATFORM_WIN32_WINDOWS:
     If osvi.dwMajorVersion = 4 And osvi.dwMinorVersion = 0 Then
     'if the OS(operating system)'s MajorVersion evaluates to 4,
     'and the OS's MinorVersion evaluates to 0 then OS is Win95...
      RetOSInf = False 'incompatible OS
       oStr$ = "Windows 95" 'update operating system description
        If LCase(osvi.szCSDVersion) = "c" Or LCase(osvi.szCSDVersion) = "b" Then oStr$ = oStr$ & " OSR2"
        'if szCSDVersion evaluates to "c" or "b" then the OS is Windows 95 OSR2
     ElseIf osvi.dwMajorVersion = 4 And osvi.dwMinorVersion = 10 Then
      RetOSInf = False '..
       oStr$ = "Windows 98" '..
        If LCase(osvi.szCSDVersion) = "a" Then oStr$ = "Windows 98 Second Edition"
        'if szCSDVersion evaluates to "a" then the OS is Window 98 Second Edition
     ElseIf osvi.dwMajorVersion = 4 And osvi.dwMinorVersion = 90 Then
     'Windows Mellenium edition
      RetOSInf = False '..
       oStr$ = "Windows Mellenium edition" '..
     End If
    Case VER_PLATFORM_WIN32_NT:
     If osvi.dwMajorVersion = 4 Then RetOSInf = False: oStr = "Windows NT 4"
     'Windows NT 4(<)
      If osvi.dwMajorVersion = 5 And osvi.dwMinorVersion = 0 Then RetOSInf = True
      'Windows 2000, OS is compatible
       If osvi.dwMajorVersion = 5 And osvi.dwMinorVersion >= 1 Or osvi.dwMajorVersion > 5 Then RetOSInf = True
       'Windows XP(>), OS is compatible
   End Select
 End If
  oStr$ = oStr$ & " build: " & osvi.dwBuildNumber  'build number of the Operating System version information
   Exit Function 'discontinue execution of this procedure
errh: 'label errh
 RetOSInf = False 'an error occured, return false
  oStr$ = "Undetermined"
End Function

Sub Main()
On Error GoTo errh ' on the event of an error jump to label errh
 If RetOSInf = False Then 'RetOsInf returns false if the OS is incompatible see RetOsInf for more info...
  Load frmSB 'load frmSB dialog into memory
   frmSB.Show 'show the dialog
 Else
  Load frmSplash 'load frmSplash dialog
 End If
  Exit Sub 'exit this procedure
errh: 'label errh
 If Err.Number = 364 Then Exit Sub
  MessageBox GetDesktopWindow, "The following error occured:" & vbCrLf & Err.Description & vbCrLf & vbCrLf & "Yahtzee XP has been terminated.", "Error " & Err.Number, vbCritical: End
End Sub

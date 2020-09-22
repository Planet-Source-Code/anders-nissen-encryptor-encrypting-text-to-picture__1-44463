Attribute VB_Name = "modAPI"
Option Explicit

'Const's & functions for making the form transparent
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'Used for the FormMove methods:
Private Const LP_HT_CAPTION = 2
Private Const WM_NCLBUTTONDOWN = &HA1
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'Round form edges:
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

'For finding color depth:
Const BITSPIXEL = 12
Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, _
  ByVal lpDeviceName As String, ByVal lpOutput As String, ByVal lpInitData As Any) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long

'For detecting the OS version:
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

'---------------------------------------------------------------------------------------
' Procedure : FormFadeIn
' DateTime  : 01-04-2003 23:03 CET
' Author    : Anders Nissen, IcySoft
' Purpose   : Increases the transparency of the form from 0 to 255
'---------------------------------------------------------------------------------------
Public Sub FormFadeIn(frmForm As Form, Optional FadeStep = 255 / 8)

  'Cannot use the "SetLayeredWindowAttributes"-API in win 3.11/95/98
  If isRuningWinNT = False Then
    'Disables the function in the "Settings"-page
    frmMain.chkFade.Value = vbGrayed
    '...and explains why it is disabled
    frmMain.chkFade.Tag = "Fading is not possible in Windows 3.11, 95 && 98"
    Exit Sub
  End If

  Dim ret As Long
  'Sets the form the act as a layer
  ret = GetWindowLong(frmForm.hwnd, GWL_EXSTYLE)
  ret = ret Or WS_EX_LAYERED
  SetWindowLong frmForm.hwnd, GWL_EXSTYLE, ret
  
  Dim ix As Double
  For ix = 0 To 255 Step FadeStep
    'Sets the transparency of the form
    SetLayeredWindowAttributes frmForm.hwnd, 0, ix, LWA_ALPHA
    DoEvents
  Next
End Sub

'---------------------------------------------------------------------------------------
' Procedure : FormFadeOut
' DateTime  : 01-04-2003 23:00 CET
' Author    : Anders Nissen, IcySoft
' Purpose   : Decreases the transparency of the form from 255 to 0
'---------------------------------------------------------------------------------------
Public Sub FormFadeOut(frmForm As Form, Optional FadeStep = 255 / 8)

  'Cannot use the "SetLayeredWindowAttributes"-API in win 3.11/95/98
  If isRuningWinNT = False Then Exit Sub

  Dim ret As Long
  'Sets the form the act as a layer
  ret = GetWindowLong(frmForm.hwnd, GWL_EXSTYLE)
  ret = ret Or WS_EX_LAYERED
  SetWindowLong frmForm.hwnd, GWL_EXSTYLE, ret
  
  Dim ix As Double
  For ix = 255 To 0 Step -FadeStep
    'Sets the transparency of the form
    SetLayeredWindowAttributes frmForm.hwnd, 0, ix, LWA_ALPHA
    DoEvents
  Next
End Sub


'---------------------------------------------------------------------------------------
' Procedure : MoveForm
' DateTime  : 01-04-2003 23:00 CET
' Author    : Anders Nissen, IcySoft
' Purpose   : Uses API to move a form without border
'---------------------------------------------------------------------------------------
Public Sub MoveForm(frmForm As Form)
  Dim rc As Long
  rc = ReleaseCapture
  rc = SendMessage(frmForm.hwnd, WM_NCLBUTTONDOWN, LP_HT_CAPTION, ByVal 0&)
End Sub

'---------------------------------------------------------------------------------------
' Procedure : RoundEdges
' DateTime  : 02-04-2003 21:01 CET
' Author    : Anders Nissen, IcySoft
' Purpose   : Rounds the edges of the form width a given radius
'---------------------------------------------------------------------------------------
Public Sub RoundEdges(frmForm As Form, Optional Radius As Integer = 13)
  Dim hRgn As Long
  hRgn = CreateRoundRectRgn(0, 0, frmForm.Width / 15, frmForm.Height / 15, Radius, Radius)
  SetWindowRgn frmForm.hwnd, hRgn, True
End Sub

'---------------------------------------------------------------------------------------
' Procedure : ColorDepth
' DateTime  : 04-04-2003 20:12 CET
' Author    : Anders Nissen, IcySoft
' Purpose   : Returns the color depth of the desktop
'---------------------------------------------------------------------------------------
Public Function ColorDepth() As Integer

  Dim nDC As Long
  'Creates a device context that equels the screen
  nDC = CreateDC("DISPLAY", vbNullString, vbNullString, ByVal 0&)
  'Find & returns the number of bits pr. pixel
  ColorDepth = GetDeviceCaps(nDC, BITSPIXEL)
  'Deletes the device context so it dosn't use memory
  DeleteDC nDC

End Function


'---------------------------------------------------------------------------------------
' Procedure : isRuningWinNT
' DateTime  : 04-04-2003 20:56 CET
' Author    : Anders Nissen, IcySoft
' Purpose   : Checks the Windows OS version and returns "true" if NT (meaning 2000, NT, XP)
'---------------------------------------------------------------------------------------
Public Function isRuningWinNT() As Boolean

  Dim OSInfo As OSVERSIONINFO
  OSInfo.dwOSVersionInfoSize = Len(OSInfo) 'Set the structure size
  Dim ret As Long
  ret& = GetVersionEx(OSInfo) 'Get the Windows version
  'Check for errors
  If ret& = 0 Then MsgBox "Error extracting Windows version information": Exit Function
  'Evaluates the OSinfomation ("2" is Windows NT)
  If OSInfo.dwPlatformId = 2 Then isRuningWinNT = True

End Function

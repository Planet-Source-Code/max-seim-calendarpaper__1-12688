Attribute VB_Name = "ScreenShot32"
Option Explicit


'Define constants used in the project
Public Const SRCCOPY = &HCC0020
Public Const SWP_NOSIZE = &H1
Public Const WM_ACTIVATE = &H6
Public Const WM_SETFOCUS = &H7
Public Const BACKDROPCOLOR = &HFFFF80

'Define functions used in the project
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hSrcDC As Integer, ByVal xSrc As Integer, ByVal ySrc As Integer, ByVal dwRop As Long) As Integer
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long


Sub Window_Reposition(win As Long)
'Reposition the window that we're taking a screenshot of
'Use SWP_NOSIZE so we don't accidentally resize the window
Call SetWindowPos(win, 0&, 0, 0, 0&, 0&, SWP_NOSIZE)
End Sub

Sub Window_SetFocus(win As Long)
'Make sure the window that we're taking a screenshot of
'has the top focu
Call SendMessage(win, WM_ACTIVATE, 0&, 0&)
Call SendMessage(win, WM_SETFOCUS, 0&, 0&)
End Sub

Sub Copy_Form(picBox, destPicBox As PictureBox, color As Long)
Dim X As Integer, Y As Integer
'Reset X and Y coordinates to 0
X = 0
Y = 0

'Locate where the form ends and the bright blue
'backdrop begins
Do
    X = X + 1
    DoEvents
Loop Until GetPixel(picBox.hdc, X, 1) = color

Do
    Y = Y + 1
    DoEvents
Loop Until GetPixel(picBox.hdc, 1, Y) = color


'Resize the destination picturebox
destPicBox.Width = X
destPicBox.Height = Y

'Copy JUST the window (not the bright blue back drop)
'to the destination picturebox
destPicBox.PaintPicture picBox.Image, 0, 0, X, Y, 0, 0, X, Y
End Sub

Sub Pause(interval)
'Pause for ___ seconds
Dim current
current = Timer
Do While Timer - current < Val(interval)
DoEvents
Loop
End Sub

Sub GrabScreenShot(TheHDC As Integer)
Dim DeskhWnd As Long, DeskDC

'Get the hWnd of the desktop
DeskhWnd = GetDesktopWindow()

'BitBlt needs the DC to copy the image. So, we
'need the GetDC API.
DeskDC = GetDC(DeskhWnd)

'Copy the screen shot to the HDC of a picturebox/form
BitBlt TheHDC, 0, 0, Screen.Width, Screen.Height, DeskDC, 0, 0, SRCCOPY
End Sub
Public Sub endall()
Unload frmTake
Unload Calendar
End
End Sub

Attribute VB_Name = "mdlJSS"
' NOTE: ScaleMode of this module is pixels.
' It means all of parameters must be in pixels.

' All of variables in this module must be declared.
Option Explicit

'Structure to pass mouse pointer info to and from DLLs.
Public Type usrPOINTAPI
X As Long
Y As Long
End Type

'Structure to pass rectangle info to and from DLLs.
Public Type usrRECT
Left As Long
Top As Long
Right As Long
Bottom As Long
End Type

' Windows API declarations
Public Declare Function EnableWindow Lib "user32" (ByVal hWnd As Integer, ByVal aBOOL As Integer) As Integer
Public Declare Function IsWindowEnabled Lib "user32" (ByVal hWnd As Integer) As Integer
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Declare Sub ClipCursor Lib "user32" (lpRect As usrRECT)
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Sub GetWindowRect Lib "user32" Alias "GetWindowRECT" (ByVal hWnd As Long, lpRect As usrRECT)

Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long

Public Declare Function ShowCursor Lib "user32" (ByVal blnSHOW As Long) As Long

' Variable declarations
Public lngTASKBARHWND As Long ' Taskbar Handler
Public intISTASKBARENABLED As Integer ' Determines Windows taskbar is enable or disable
Public intS_1 As Integer ' Will be used to hide/show mouse cursor

' This procedure enables the following
' keys in Windows:
' 1. Ctrl+Alt+Del [Close Program]
' 2. Alt+Tab [Quick Program Select]
' 3. Ctrl+Esc [Open Start Menu]
Public Sub KeysOn()
Dim lngA As Long, lngDISABLED As Long

lngDISABLED = False
lngA = SystemParametersInfo(97, lngDISABLED, CStr(1), 0)
End Sub

' This procedure disables the following
' keys in Windows:
' 1. Ctrl+Alt+Del [Close Program]
' 2. Alt+Tab [Quick Program Select]
' 3. Ctrl+Esc [Open Start Menu]
Public Sub KeysOff()
Dim lngA As Long, lngDISABLED As Long

lngDISABLED = True
lngA = SystemParametersInfo(97, lngDISABLED, CStr(1), 0)
End Sub

' This procedure disables Windows taskBar,
' but taskbar will be visible.
Public Sub DisableTaskBar()
Dim EWindow As Integer

lngTASKBARHWND = FindWindow("Shell_traywnd", "")
If lngTASKBARHWND <> 0 Then
EWindow = IsWindowEnabled(lngTASKBARHWND)
If EWindow = 1 Then _
 intISTASKBARENABLED = EnableWindow(lngTASKBARHWND, 0)
End If
End Sub

' This procedure enables Windows taskBar.
Public Sub EnableTaskBar()
If intISTASKBARENABLED = 0 Then _
 intISTASKBARENABLED = EnableWindow(lngTASKBARHWND, 1)
End Sub

' Pass a set of points as a rectangle and
' the mouse cursor will be limited to
' move only in that region.
Public Sub LimitCursor(Left, Top, Right, Bottom As Long)
Dim rctBox As usrRECT

rctBox.Left = Left
rctBox.Top = Top
rctBox.Right = Right
rctBox.Bottom = Bottom
ClipCursor rctBox
End Sub

' This procedure resets the cursor limit back to
' entire screen (turns limiting off)
Public Sub LimitCursorOff()
Dim rctBox As usrRECT
Dim hwndDesktop As Long
   
hwndDesktop = GetDesktopWindow()
GetWindowRect hwndDesktop, rctBox
ClipCursor rctBox
End Sub

' This procedure makes mouse cursor visible.
Public Sub CursorOn()
Dim intS_2 As Integer
intS_2 = ShowCursor(True)
Do While intS_2 < intS_1
intS_2 = ShowCursor(True)
Loop
End Sub

' This procedure makes mouse cursor unvisible.
Public Sub CursorOff()
Dim intS_2 As Integer
intS_2 = ShowCursor(False)
intS_1 = intS_2 + 1
Do While intS_2 > -1
intS_2 = ShowCursor(False)
Loop
End Sub

' This procedure moves the mouse cursor
' to a new place.
Public Sub MoveCursor(X As Long, Y As Long)
Dim lngA As Long
Dim lngNEWX As Long
Dim lngNEWY As Long
    
lngNEWX = X
lngNEWY = Y
lngA = SetCursorPos(lngNEWX, lngNEWY)
End Sub

' This is the startup procedure.
Public Sub Main()
' Place your extra code here!
End Sub


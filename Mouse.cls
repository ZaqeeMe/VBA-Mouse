Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Sub mouse_event Lib "user32.dll" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal dwData As Long, ByVal dwExtraInfo As Long)
    Private Declare PtrSafe Function SetCursorPos Lib "user32.dll" (ByVal X As Long, ByVal Y As Long) As Long
    Private Declare PtrSafe Function GetCursorPos Lib "user32.dll" (lpPoint As POINTAPI) As Long
#Else
    Private Declare Sub mouse_event Lib "user32.dll" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal dwData As Long, ByVal dwExtraInfo As Long)
    Private Declare Function SetCursorPos Lib "user32.dll" (ByVal X As Long, ByVal Y As Long) As Long
    Private Declare Function GetCursorPos Lib "user32.dll" (lpPoint As POINTAPI) As Long
#End If

Private Const MOUSEEVENTF_LEFTDOWN As Long = &H2
Private Const MOUSEEVENTF_LEFTUP As Long = &H4
Private Const MOUSEEVENTF_RIGHTDOWN As Long = &H8
Private Const MOUSEEVENTF_RIGHTUP As Long = &H10

Private Type POINTAPI
    X As Long
    Y As Long
End Type

' Method to simulate left mouse click
Public Sub SimulateLeftClick(X As Long, Y As Long)
    ' Set cursor position
    SetCursorPos X, Y
    ' Simulate left mouse button down and up
    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
    mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
End Sub

' Method to simulate right mouse click
Public Sub SimulateRightClick(X As Long, Y As Long)
    ' Set cursor position
    SetCursorPos X, Y
    ' Simulate right mouse button down and up
    mouse_event MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0
    mouse_event MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0
End Sub

' Method to set the mouse position without clicking
Public Sub SetMousePosition(X As Long, Y As Long)
    ' Set cursor position
    SetCursorPos X, Y
End Sub

' Method to get the current mouse position
Public Sub GetMousePosition(ByRef X As Long, ByRef Y As Long)
    Dim pt As POINTAPI
    GetCursorPos pt
    X = pt.X
    Y = pt.Y
End Sub

' Method to display the current mouse position in a message box
Public Sub ShowMousePosition()
    Dim X As Long
    Dim Y As Long
    GetMousePosition X, Y
    MsgBox "Current Mouse Position: (" & X & ", " & Y & ")", vbInformation, "Mouse Position"
End Sub


Option Explicit

' Example usage of the MouseController class
Sub TestMouseController()
    Dim Mouse As Mouse
    Set Mouse = New Mouse
    
    Dim xPos As Long
    Dim yPos As Long
    
    ' Show current mouse position
    Mouse.ShowMousePosition
    
    ' Coordinates for setting the mouse position
    xPos = 500
    yPos = 300
    
    ' Set mouse position without clicking
    Mouse.SetMousePosition xPos, yPos
    
    ' Optional: Wait for a while before the next action
    Application.Wait (Now + TimeValue("0:00:02")) ' Waits for 2 seconds
    
    ' Simulate left click at specified coordinates
    Mouse.SimulateLeftClick xPos, yPos
    
    ' Optional: Wait for a while before the next action
    Application.Wait (Now + TimeValue("0:00:02")) ' Waits for 2 seconds
    
    ' Simulate right click at specified coordinates
    Mouse.SimulateRightClick xPos, yPos
End Sub


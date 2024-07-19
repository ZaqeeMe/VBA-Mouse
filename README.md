
# VBA Mouse

A VBA class module for simulating mouse events and controlling cursor positioning in Windows.

## Features

- **Simulate Left Click**: Programmatically simulate a left mouse click at any screen coordinate.
- **Simulate Right Click**: Programmatically simulate a right mouse click at any screen coordinate.
- **Set Mouse Position**: Move the mouse cursor to a specified position without clicking.
- **Get Mouse Position**: Retrieve the current coordinates of the mouse cursor.
- **Display Mouse Position**: Show a message box with the current mouse cursor coordinates.

`** More Features to be added in the future **`


## Installation

1. Downlaod the **Mouse.cls** file.

2. Open your VBA project in the VBA editor.

3. Import the **Mouse.cls** class module into the project

## Usage

1. **Import the `Mouse` Class Module** into your VBA project.
2. **Instantiate the `Mouse` Class** in your VBA code.
3. **Call the Methods** to control the mouse.
## Example

```vba
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



```

## Support Me

[!["Buy Me A Coffee"](https://www.buymeacoffee.com/assets/img/custom_images/orange_img.png)](https://www.buymeacoffee.com/zaqee)

[![ko-fi](https://ko-fi.com/img/githubbutton_sm.svg)](https://ko-fi.com/O4O2PKT0A)

# eVB-Taskbar-Icon

This module enables your eVB application to have an icon in the Windows CE taskbar.

## Usage

Usage is very simple, just add the module and an ImageList to your project and add the image(s) to the ImageList as per normal. Add a form-level variable to hold the icon(s), and then call the `TaskbarIcon_Set` function from Form_Load or anytime you want to change the icon.

A complete application could look like this:

```vb
Option Explicit

Private LargeIcon As Long

Private SmallIcon As Long

Private Sub Form_Load()
    ImageList.Add "icon_small.bmp"
    ImageList.Add "icon_large.bmp"
    
    SmallIcon = TaskbarIcon_Let(hwnd, ImageList.hImageList, 0, tbSmallIcon)
    LargeIcon = TaskbarIcon_Let(hwnd, ImageList.hImageList, 1, tbLargeIcon)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    TaskbarIcon_Destroy LargeIcon
    TaskbarIcon_Destroy SmallIcon
End Sub
```

Currently this is bound to the same limitations as other icons represented by the eVB ImageList, most notably, transparency is not allowed.

## Screenshots

![Screenshot showing the example application. To demonstrate functionality, the application icon is visible in the Windows taskbar.](https://raw.githubusercontent.com/WinCEDev/eVB-Taskbar-Icon/main/Screenshots/CAPT0000.png?raw=1)
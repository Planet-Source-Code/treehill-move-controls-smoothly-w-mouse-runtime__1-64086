<div align="center">

## Move Controls Smoothly w/ Mouse \(Runtime\)


</div>

### Description

This code demostrates how to move controls smoothly on a form without them redrawing slowly or chunky looking.
 
### More Info
 
Just click and drag the picturebox to move it anywhere on the form.

I recycled this code that was moving a form, which when calling the SendMessage function, it would mimic the form's menu being clicked and moved, but I changed the handle (hWnd) to the pictureboxes, and it worked.

As of yet, I am unaware if I am using the correct flag (HTCAPTION) as a SendMessage param. Obviously it works, but I wasn't sure if there was a different flag that was more appropriate. Reason being, that there is no caption in a picturebox. Be aware of this and any side effects this might case in your code.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Treehill](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/treehill.md)
**Level**          |Intermediate
**User Rating**    |4.3 (13 globes from 3 users)
**Compatibility**  |VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/treehill-move-controls-smoothly-w-mouse-runtime__1-64086/archive/master.zip)

### API Declarations

```
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function ReleaseCapture Lib "user32" () As Long
```


### Source Code

```
'Copy this code into a form, and add a picturebox
Option Explicit
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = vbLeftButton Then
  ReleaseCapture
  Call SendMessage(Picture1.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
 End If
End Sub
```


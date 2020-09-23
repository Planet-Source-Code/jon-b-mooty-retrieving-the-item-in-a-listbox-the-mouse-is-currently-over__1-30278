<div align="center">

## Retrieving the Item in a Listbox the mouse is currently over\.


</div>

### Description

As you move your mouse pointer over top of a listbox this code will return the index and select the item underneath it.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Jon B\. Mooty](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/jon-b-mooty.md)
**Level**          |Advanced
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/jon-b-mooty-retrieving-the-item-in-a-listbox-the-mouse-is-currently-over__1-30278/archive/master.zip)

### API Declarations

```
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const LB_GETITEMRECT = &H198
Private Type RECT
 Left As Long
 Top As Long
 Right As Long
 Bottom As Long
End Type
'You will also need a ListBox named List1
```


### Source Code

```

Private Sub Form_Load()
  Dim i As Integer
  ' populate the listbox with the
  ' indexes for each entry
  For i = 0 To 50
    List1.AddItem i, i
  Next i
End Sub
Private Sub List1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
 Dim iSelected As Integer
  ' retrieve the index the mouse
  ' pointer is over (-1 represents
  ' mouse is over nothing)
  iSelected = GetListItemFromPt(List1, x, y)
  ' select the item the mouse
  ' is currently over
  If Not iSelected = -1 Then
    List1.ListIndex = iSelected
  End If
End Sub
Public Function GetListItemFromPt(plist As ListBox, x As Single, y As Single) As Integer
  Dim bFound As Boolean
  Dim rCur As RECT
  Dim iIndex As Integer
  Dim iPixX As Integer, iPixY As Integer
  ' convert the coordinates to
  ' pixels (remove the next two
  ' lines if you will be passing
  ' the coordinates in pixels)
  iPixX = Me.ScaleX(x, vbTwips, vbPixels)
  iPixY = Me.ScaleY(y, vbTwips, vbPixels)
  For iIndex = 0 To plist.ListCount - 1
    ' get the coordinates for each
    ' item in the listbox in its current
    ' state, if Top is less than 0 or Bottom
    '
    ' is greater than the height of the
    ' listbox then the item is currently off
    '   screen
    SendMessage plist.hwnd, LB_GETITEMRECT, iIndex, rCur
    ' if passed corrdinates are within
    ' the bounds of the current item than
    ' exit the loop and return the index
    If iPixY >= rCur.Top And iPixY <= rCur.Bottom Then
      bFound = True
      Exit For
    End If
  Next iIndex
  If bFound Then
    GetListItemFromPt = iIndex
  Else
    GetListItemFromPt = -1
  End If
End Function
```


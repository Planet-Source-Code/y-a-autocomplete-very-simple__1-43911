<div align="center">

## a AutoComplete Very Simple\!


</div>

### Description

VERY SIMPLE cut and paste funtion for the Keypress event of a combobox. Just paste this code into a module or form and call the function from the KeyPress event. KeyAscii = AutoComplete(cboCombobox, KeyAscii,Optional UpperCase)
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[\( \. Y \. \)](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/y.md)
**Level**          |Beginner
**User Rating**    |4.9 (93 globes from 19 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/y-a-autocomplete-very-simple__1-43911/archive/master.zip)





### Source Code

```
Option Explicit
Public Const CB_FINDSTRING = &H14C
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Function AutoComplete(cbCombo As ComboBox, sKeyAscii As Integer, Optional bUpperCase As Boolean = True) As Integer
 Dim lngFind As Long, intPos As Integer, intLength As Integer
 Dim tStr As String
 With cbCombo
 If sKeyAscii = 8 Then
 If .SelStart = 0 Then Exit Function
 .SelStart = .SelStart - 1
 .SelLength = 32000
 .SelText = ""
 Else
 intPos = .SelStart '// save intial cursor position
 tStr = .Text '// save string
 If bUpperCase = True Then
 .SelText = UCase(Chr(sKeyAscii)) '// change string. (uppercase only)
 Else
 .SelText = UCase(Chr(sKeyAscii)) '// change string. (leave case alone)
 End If
 End If
 lngFind = SendMessage(.hwnd, CB_FINDSTRING, 0, ByVal .Text) '// Find string in combobox
 If lngFind = -1 Then '// if string not found
 .Text = tStr '// set old string (used for boxes that require charachter monitoring
 .SelStart = intPos '// set cursor position
 .SelLength = (Len(.Text) - intPos) '// set selected length
 AutoComplete = 0 '// return 0 value to KeyAscii
 Exit Function
 Else '// If string found
 intPos = .SelStart '// save cursor position
 intLength = Len(.List(lngFind)) - Len(.Text) '// save remaining highlighted text length
 .SelText = .SelText & Right(.List(lngFind), intLength) '// change new text in string
 '.Text = .List(lngFind)'// Use this instead of the above .Seltext line to set the text typed to the exact case of the item selected in the combo box.
 .SelStart = intPos '// set cursor position
 .SelLength = intLength '// set selected length
 End If
 End With
End Function
```


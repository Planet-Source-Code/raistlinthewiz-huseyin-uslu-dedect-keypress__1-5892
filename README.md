<div align="center">

## Dedect keypress


</div>

### Description

This code finds if any key pressed. This will be helpfull for example you're programming a screensaver. But you can not get keypresses through forms because of some components which does not support this keypress event. So you must use some API's to do this.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[\-raistlinthewiz\- Huseyin Uslu](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/raistlinthewiz-huseyin-uslu.md)
**Level**          |Intermediate
**User Rating**    |3.3 (10 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/raistlinthewiz-huseyin-uslu-dedect-keypress__1-5892/archive/master.zip)

### API Declarations

```
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
```


### Source Code

```
sKeyStat = 0
For i = 0 To 255
 KeyResult = GetAsyncKeyState(i)
  If KeyResult = -32767 Then
   sKeyStat = 1
   Exit For
  End If
Next i
 If sKeyStat = 1 Then
  msgbox "Key pressed!!!"
 End If
```


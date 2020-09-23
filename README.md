<div align="center">

## Pointer Manipulation


</div>

### Description

Ever wanted to pass a String to a callback function, only to find out it only accepts Long's as inputs?

Now you can pass in the memory address (A long), then use the string as normal within the function.

Also included are pointer conversions to Two, Four, Eight and Twenty Two bit variables
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2006-10-05 06:50:02
**By**             |[Rob](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/rob.md)
**Level**          |Advanced
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Pointer\_Ma2023271052006\.zip](https://github.com/Planet-Source-Code/rob-pointer-manipulation__1-66711/archive/master.zip)

### API Declarations

```
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
```






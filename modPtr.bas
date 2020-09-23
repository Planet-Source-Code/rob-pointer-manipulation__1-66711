Attribute VB_Name = "Module1"
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

Public Function testMyConversionFunction(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    MsgBox ptrToStr(uMsg)
End Function

Public Function ptrToStr(memAddress As Long, Optional isUnicode As Boolean = True) As String
Dim lenField As Long
Dim charArr() As Byte
Dim returnStr As String

CopyMemory lenField, ByVal memAddress - 4, 4
ReDim charArr(lenField)
CopyMemory charArr(0), ByVal memAddress, lenField

returnStr = charArr
If Not isUnicode Then returnStr = StrConv(returnStr, vbUnicode)
ptrToStr = returnStr

End Function

Public Function ptrToTwoBytes(memAddress As Long) As Integer
Dim returnVal As Integer
CopyMemory returnVal, ByVal memAddress, 2
ptrToTwoBytes = returnVal
End Function

Public Function ptrToFourBytes(memAddress As Long) As Long
Dim returnVal As Long
CopyMemory returnVal, ByVal memAddress, 4
ptrToFourBytes = returnVal
End Function

Public Function ptrToEightBytes(memAddress As Long) As Currency
Dim returnVal As Currency
CopyMemory returnVal, ByVal memAddress, 8
ptrToEightBytes = returnVal
End Function

Public Function ptrToTwentyTwoBytes(memAddress As Long) As Variant
Dim returnVal As Variant
CopyMemory returnVal, ByVal memAddress, 22
ptrToFourteenBytes = returnVal
End Function

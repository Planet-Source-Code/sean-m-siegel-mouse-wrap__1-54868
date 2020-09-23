Attribute VB_Name = "modMouse"
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Public Type POINTAPI
    x As Long
    y As Long
End Type

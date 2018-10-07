Attribute VB_Name = "mActiveForm"
'Option Explicit
'Private x As Long
'Private mForm As Form
'Private mElastic As cElastic
'
'
'' SetWindowLong / GetWindowLong
'Public Const GWL_WNDPROC = -4
'Public Const GWL_HINSTANCE = -6
'Public Const GWL_ID = -12
'Public Const GWL_USERDATA = -21
'Public Const GWL_NEWUSERDATA = -8
'
'Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
'Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
'Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPRevWndFunc As Long, ByVal hwnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'
'
'

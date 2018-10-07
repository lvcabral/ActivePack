Attribute VB_Name = "modActiveStatus"
Option Explicit
Public mStatus As cStatusBar
Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Function GetCapslock() As Boolean
' Return or set the Capslock toggle.
GetCapslock = CBool(GetKeyState(vbKeyCapital) And 1)
End Function

Function GetNumlock() As Boolean
' Return or set the Numlock toggle.
GetNumlock = CBool(GetKeyState(vbKeyNumlock) And 1)
End Function

Function GetScrollLock() As Boolean
    ' Return or set the ScrollLock toggle.
    GetScrollLock = CBool(GetKeyState(vbKeyScrollLock) And 1)
End Function

Function GetInsertKey() As Boolean
    ' Return or set the ScrollLock toggle.
    GetInsertKey = CBool(GetKeyState(vbKeyInsert) And 1)
End Function


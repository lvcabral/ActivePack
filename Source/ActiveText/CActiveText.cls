VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CActiveText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private strID As String
Event Change(SourceControl As Object)
Event Click(SourceControl As Object)
Event DblClick(SourceControl As Object)
Event GotFocus(SourceControl As Object)
Event LostFocus(SourceControl As Object)
Event KeyDown(SourceControl As Object, KeyCode As Integer, Shift As Integer)
Event KeyPress(SourceControl As Object, KeyAscii As Integer)
Event KeyUp(SourceControl As Object, KeyCode As Integer, Shift As Integer)
Event MouseDown(SourceControl As Object, Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(SourceControl As Object, Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(SourceControl As Object, Button As Integer, Shift As Integer, X As Single, Y As Single)

Private Sub Class_Initialize()
    On Error Resume Next
    If colSuperClass Is Nothing Then
        Set colSuperClass = New Collection
    End If
    lngCounter = lngCounter + 1
    strID = "ID" & lngCounter
    colSuperClass.Add Me, strID
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    colSuperClass.Remove strID
End Sub

Friend Sub RaiseChange(SourceControl As Object)
    RaiseEvent Change(SourceControl)
End Sub

Friend Sub RaiseClick(SourceControl As Object)
    RaiseEvent Click(SourceControl)
End Sub

Friend Sub RaiseDblClick(SourceControl As Object)
    RaiseEvent DblClick(SourceControl)
End Sub

Friend Sub RaiseGotFocus(SourceControl As Object)
    RaiseEvent GotFocus(SourceControl)
End Sub

Friend Sub RaiseLostFocus(SourceControl As Object)
    RaiseEvent LostFocus(SourceControl)
End Sub

Friend Sub RaiseKeyDown(SourceControl As Object, KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(SourceControl, KeyCode, Shift)
End Sub

Friend Sub RaiseKeyUp(SourceControl As Object, KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(SourceControl, KeyCode, Shift)
End Sub

Friend Sub RaiseKeyPress(SourceControl As Object, KeyAscii As Integer)
    RaiseEvent KeyPress(SourceControl, KeyAscii)
End Sub

Friend Sub RaiseMouseDown(SourceControl As Object, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(SourceControl, Button, Shift, X, Y)
End Sub

Friend Sub RaiseMouseMove(SourceControl As Object, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(SourceControl, Button, Shift, X, Y)
End Sub

Friend Sub RaiseMouseUp(SourceControl As Object, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(SourceControl, Button, Shift, X, Y)
End Sub


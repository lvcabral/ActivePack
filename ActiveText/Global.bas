Attribute VB_Name = "basGlobal"
Option Explicit
Public colSuperClass As Collection
Public lngCounter As Long

Sub RaiseChange(SourceControl As Object)
Dim obj As CActiveText
    On Error Resume Next
    If Not colSuperClass Is Nothing Then
        For Each obj In colSuperClass
            obj.RaiseChange SourceControl
        Next
    End If
End Sub

Sub RaiseClick(SourceControl As Object)
Dim obj As CActiveText
    On Error Resume Next
    If Not colSuperClass Is Nothing Then
        For Each obj In colSuperClass
            obj.RaiseClick SourceControl
        Next
    End If
End Sub

Sub RaiseDblClick(SourceControl As Object)
Dim obj As CActiveText
    On Error Resume Next
    If Not colSuperClass Is Nothing Then
        For Each obj In colSuperClass
            obj.RaiseDblClick SourceControl
        Next
    End If
End Sub

Sub RaiseGotFocus(SourceControl As Object)
Dim obj As CActiveText
    On Error Resume Next
    If Not colSuperClass Is Nothing Then
        For Each obj In colSuperClass
            obj.RaiseGotFocus SourceControl
        Next
    End If
End Sub

Sub RaiseLostFocus(SourceControl As Object)
Dim obj As CActiveText
    On Error Resume Next
    If Not colSuperClass Is Nothing Then
        For Each obj In colSuperClass
            obj.RaiseLostFocus SourceControl
        Next
    End If
End Sub

Sub RaiseKeyDown(SourceControl As Object, KeyCode As Integer, Shift As Integer)
Dim obj As CActiveText
    On Error Resume Next
    If Not colSuperClass Is Nothing Then
        For Each obj In colSuperClass
            obj.RaiseKeyDown SourceControl, KeyCode, Shift
        Next
    End If
End Sub

Sub RaiseKeyUp(SourceControl As Object, KeyCode As Integer, Shift As Integer)
Dim obj As CActiveText
    On Error Resume Next
    If Not colSuperClass Is Nothing Then
        For Each obj In colSuperClass
            obj.RaiseKeyUp SourceControl, KeyCode, Shift
        Next
    End If
End Sub

Sub RaiseKeyPress(SourceControl As Object, KeyAscii As Integer)
Dim obj As CActiveText
    On Error Resume Next
    If Not colSuperClass Is Nothing Then
        For Each obj In colSuperClass
            obj.RaiseKeyPress SourceControl, KeyAscii
        Next
    End If
End Sub

Sub RaiseMouseDown(SourceControl As Object, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim obj As CActiveText
    On Error Resume Next
    If Not colSuperClass Is Nothing Then
        For Each obj In colSuperClass
            obj.RaiseMouseDown SourceControl, Button, Shift, X, Y
        Next
    End If
End Sub

Sub RaiseMouseMove(SourceControl As Object, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim obj As CActiveText
    On Error Resume Next
    If Not colSuperClass Is Nothing Then
        For Each obj In colSuperClass
            obj.RaiseMouseMove SourceControl, Button, Shift, X, Y
        Next
    End If
End Sub

Sub RaiseMouseUp(SourceControl As Object, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim obj As CActiveText
    On Error Resume Next
    If Not colSuperClass Is Nothing Then
        For Each obj In colSuperClass
            obj.RaiseMouseUp SourceControl, Button, Shift, X, Y
        Next
    End If
End Sub


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInstances"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"Members"
Option Explicit
Public Key As String
Public oType As String
Public Description As String
Public nName As String

Public dType As String
Public Arguments As String
Public MoreInfo As String
Public Example As String
Public CLSID As String
Public Value As String

Private mvarMembers As clsMembers

Public Property Get Members() As clsMembers
    If mvarMembers Is Nothing Then
        Set mvarMembers = New clsMembers
    End If
 
    Set Members = mvarMembers
End Property

Public Property Set Members(vData As clsMembers)
    Set mvarMembers = vData
End Property

Private Sub Class_Terminate()
    Set mvarMembers = Nothing
End Sub




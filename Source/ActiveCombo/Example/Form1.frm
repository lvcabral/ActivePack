VERSION 5.00
Object = "{ED442B9F-ADE2-11D4-B868-00606E3BC2C9}#1.0#0"; "ActiveCbo.ocx"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ActiveCombo Example"
   ClientHeight    =   750
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   750
   ScaleWidth      =   2505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin rdActiveCombo.ActiveCombo ActiveCombo1 
      Height          =   330
      Left            =   1095
      TabIndex        =   0
      Top             =   120
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   582
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IconColorDepth  =   3
      ShowIcons       =   -1  'True
      Style           =   1
   End
   Begin VB.Label Label1 
      Caption         =   "Drives:"
      Height          =   240
      Left            =   195
      TabIndex        =   1
      Top             =   135
      Width           =   840
   End
   Begin VB.Image imgIcons 
      Height          =   240
      Index           =   3
      Left            =   1650
      Picture         =   "Form1.frx":0000
      Top             =   510
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgIcons 
      Height          =   240
      Index           =   2
      Left            =   1185
      Picture         =   "Form1.frx":058A
      Top             =   510
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgIcons 
      Height          =   240
      Index           =   1
      Left            =   765
      Picture         =   "Form1.frx":0B14
      Top             =   510
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgIcons 
      Height          =   240
      Index           =   0
      Left            =   405
      Picture         =   "Form1.frx":109E
      Top             =   510
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim ico
    For Each ico In imgIcons
        ActiveCombo1.AddIcon ico
        ActiveCombo1.AddItem "Drive " & Chr(67 + ico.Index), , ico.Index + 1
    Next
    ActiveCombo1.ListIndex = 0
End Sub

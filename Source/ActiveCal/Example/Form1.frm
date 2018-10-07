VERSION 5.00
Object = "{703944EE-9203-11D2-8865-AD1268A0A52F}#1.0#0"; "ActiveCal.ocx"
Begin VB.Form Form1 
   Caption         =   "ActiveCalendar Example"
   ClientHeight    =   3525
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3690
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3525
   ScaleWidth      =   3690
   StartUpPosition =   3  'Windows Default
   Begin rdActiveCal.ActiveCalendar ActiveCalendar1 
      Height          =   3540
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3720
      _ExtentX        =   6562
      _ExtentY        =   6244
      Date            =   36623
      TodayCaption    =   "&Today"
      BorderStyle     =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ActiveCalendar1_DblClick()
    MsgBox Format(ActiveCalendar1.SelectedDate, "Long Date")
End Sub

Private Sub Form_Resize()
    ActiveCalendar1.Move 0, 0, ScaleWidth, ScaleHeight
End Sub

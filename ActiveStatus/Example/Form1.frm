VERSION 5.00
Object = "{ABFD3451-F684-11D2-9905-006008CDEC24}#1.0#0"; "ActiveStatus.ocx"
Begin VB.Form Form1 
   Caption         =   "ActiveStatus Example"
   ClientHeight    =   5250
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7545
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5250
   ScaleWidth      =   7545
   StartUpPosition =   3  'Windows Default
   Begin rdActiveStatus.ActiveStatus ActiveStatus1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   4920
      Width           =   7545
      _ExtentX        =   13309
      _ExtentY        =   582
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image imgDate 
      Height          =   240
      Left            =   180
      Picture         =   "Form1.frx":030A
      Top             =   150
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
    ActiveStatus1.Panels.Add , , "Raindrops ActiveStatus", sbrText
    ActiveStatus1.Panels.Add , , , sbrIns
    ActiveStatus1.Panels.Add , , , sbrNum
    ActiveStatus1.Panels.Add , , , sbrCaps
    ActiveStatus1.Panels.Add , , , sbrScrl
    ActiveStatus1.Panels.Add , , , sbrDate, imgDate.Picture
    ActiveStatus1.Panels.Add , , , sbrTime
    ActiveStatus1.Panels(1).AutoSize = sbrSpring
    ActiveStatus1.Panels(5).Bevel = sbrRaised
    ' Add a Panel to make a Border
    ActiveStatus1.Panels.Add
    ActiveStatus1.Panels(8).Width = 7
    ActiveStatus1.Panels(8).MinWidth = 7
    ActiveStatus1.Panels(8).Bevel = sbrNoBevel
End Sub

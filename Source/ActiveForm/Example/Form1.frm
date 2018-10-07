VERSION 5.00
Object = "{8E22FD0B-91ED-11D2-8865-EAF032485D5B}#1.3#0"; "ActiveForm.ocx"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "ActiveForm Example"
   ClientHeight    =   2640
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2640
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1500
      Left            =   1890
      TabIndex        =   8
      Top             =   45
      Width           =   1980
      Begin VB.CheckBox chkBackground 
         Caption         =   "Background"
         Height          =   315
         Left            =   180
         TabIndex        =   11
         Top             =   1035
         Value           =   1  'Checked
         Width           =   1485
      End
      Begin VB.CheckBox chkMinToTray 
         Caption         =   "MinimizeToTray"
         Height          =   315
         Left            =   180
         TabIndex        =   10
         Top             =   645
         Value           =   1  'Checked
         Width           =   1485
      End
      Begin VB.CheckBox chkTopMost 
         Caption         =   "AllwaysOnTop"
         Height          =   315
         Left            =   180
         TabIndex        =   9
         Top             =   270
         Width           =   1350
      End
   End
   Begin VB.Frame fraGradient 
      Caption         =   "BackGradient"
      Height          =   2460
      Left            =   60
      TabIndex        =   1
      Top             =   45
      Width           =   1725
      Begin VB.OptionButton optGradient 
         Caption         =   "Diagonal Right"
         Height          =   330
         Index           =   5
         Left            =   180
         TabIndex        =   7
         Top             =   2025
         Width           =   1365
      End
      Begin VB.OptionButton optGradient 
         Caption         =   "Diagonal Left "
         Height          =   330
         Index           =   4
         Left            =   180
         TabIndex        =   6
         Top             =   1680
         Width           =   1275
      End
      Begin VB.OptionButton optGradient 
         Caption         =   "Rectangle"
         Height          =   330
         Index           =   3
         Left            =   180
         TabIndex        =   5
         Top             =   1320
         Width           =   1275
      End
      Begin VB.OptionButton optGradient 
         Caption         =   "Circle"
         Height          =   330
         Index           =   2
         Left            =   180
         TabIndex        =   4
         Top             =   975
         Width           =   1275
      End
      Begin VB.OptionButton optGradient 
         Caption         =   "Horizontal"
         Height          =   330
         Index           =   1
         Left            =   180
         TabIndex        =   3
         Top             =   630
         Width           =   1275
      End
      Begin VB.OptionButton optGradient 
         Caption         =   "Vertical"
         Height          =   330
         Index           =   0
         Left            =   180
         TabIndex        =   2
         Top             =   270
         Value           =   -1  'True
         Width           =   1275
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   360
      Left            =   3405
      TabIndex        =   0
      Top             =   2115
      Width           =   1155
   End
   Begin rdActiveForm.ActiveForm ActiveForm1 
      Left            =   4125
      Top             =   60
      _ExtentX        =   794
      _ExtentY        =   688
      MinWidth        =   4800
      MinHeight       =   3330
      ResizeControls  =   -1  'True
      ResizeFonts     =   -1  'True
      CloseButton     =   0   'False
      Background      =   1
      MinimizeToTray  =   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   4215
      Picture         =   "Form1.frx":0442
      Top             =   630
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Menu miPopUp 
      Caption         =   "PopUp"
      Begin VB.Menu miRestore 
         Caption         =   "&Restore"
      End
      Begin VB.Menu miMove 
         Caption         =   "&Move"
         Enabled         =   0   'False
      End
      Begin VB.Menu miSize 
         Caption         =   "Re&size"
         Enabled         =   0   'False
      End
      Begin VB.Menu miMin 
         Caption         =   "&Minimize"
         Enabled         =   0   'False
      End
      Begin VB.Menu miMax 
         Caption         =   "Ma&ximize"
      End
      Begin VB.Menu miLine0 
         Caption         =   "-"
      End
      Begin VB.Menu miIcon 
         Caption         =   "&Change Icon"
      End
      Begin VB.Menu miTooltip 
         Caption         =   "Change &Tooltip"
      End
      Begin VB.Menu miLine1 
         Caption         =   "-"
      End
      Begin VB.Menu miAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu miLine2 
         Caption         =   "-"
      End
      Begin VB.Menu miClose 
         Caption         =   "&Close"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ActiveForm1_TrayRightClick()
    PopupMenu miPopUp, , , , miRestore
End Sub

Private Sub chkBackground_Click()
    If chkBackground.Value = 0 Then
        ActiveForm1.Background = bgNone
        Me.Cls
    Else
        ActiveForm1.Background = bgGradient
        Me.Move Left, Top, Width + Screen.TwipsPerPixelX
        Me.Move Left, Top, Width - Screen.TwipsPerPixelX
    End If
End Sub

Private Sub chkMinToTray_Click()
    ActiveForm1.MinimizeToTray = chkMinToTray.Value
End Sub

Private Sub chkTopMost_Click()
    ActiveForm1.AllwaysOnTop = chkTopMost.Value
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    miPopUp.Visible = False
End Sub

Private Sub miAbout_Click()
    ActiveForm1.AboutBox
End Sub

Private Sub miClose_Click()
    Unload Me
End Sub

Private Sub miIcon_Click()
    ActiveForm1.SetTrayIcon Image1.Picture
End Sub

Private Sub miMax_Click()
    ActiveForm1.Restore vbMaximized
End Sub

Private Sub miRestore_Click()
    ActiveForm1.Restore
End Sub

Private Sub miTooltip_Click()
    ActiveForm1.SetTrayToolTip "Click here now!"
End Sub

Private Sub optGradient_Click(Index As Integer)
    ActiveForm1.BackGradient = Index
    Me.Move Left, Top, Width + Screen.TwipsPerPixelX
    Me.Move Left, Top, Width - Screen.TwipsPerPixelX
End Sub

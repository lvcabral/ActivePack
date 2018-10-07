VERSION 5.00
Object = "{34CE1DC5-0F06-11D4-8632-00E07D813CFC}#1.0#0"; "ActiveScroll.ocx"
Begin VB.Form Form1 
   Caption         =   "ActiveScroll Example"
   ClientHeight    =   2790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4230
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2790
   ScaleWidth      =   4230
   StartUpPosition =   3  'Windows Default
   Begin rdActiveScroll.ActiveScroll ActiveScroll1 
      Height          =   540
      Left            =   3660
      TabIndex        =   1
      Top             =   2220
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   953
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   2520
      Left            =   165
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


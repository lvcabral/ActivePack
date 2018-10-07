VERSION 5.00
Object = "{F454059D-91FE-11D2-8865-AD1268A0A52F}#2.0#0"; "ActiveDate.ocx"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ActiveDate Example"
   ClientHeight    =   1170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3000
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1170
   ScaleWidth      =   3000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin rdActiveDate.ActiveDate ActiveDate1 
      Height          =   315
      Left            =   1335
      TabIndex        =   0
      Top             =   165
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TodayCaption    =   "&Today"
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin rdActiveDate.ActiveDate ActiveDate2 
      Height          =   315
      Left            =   1335
      TabIndex        =   1
      Top             =   630
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TodayCaption    =   "&Today"
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin VB.Label Label2 
      Caption         =   "Last Date:"
      Height          =   300
      Left            =   105
      TabIndex        =   3
      Top             =   645
      Width           =   1125
   End
   Begin VB.Label Label1 
      Caption         =   "First Date:"
      Height          =   300
      Left            =   105
      TabIndex        =   2
      Top             =   180
      Width           =   1125
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


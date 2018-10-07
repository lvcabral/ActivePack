VERSION 5.00
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.5#0"; "ACTIVE~1.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ActiveText Example"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6000
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin rdActiveText.ActiveText ActiveText4 
      Height          =   330
      Left            =   1440
      TabIndex        =   7
      Top             =   1500
      Width           =   1275
      _ExtentX        =   2249
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
      MaxLength       =   10
      Text            =   "52.020-010"
      TextMask        =   9
      RawText         =   9
      Mask            =   "##.###-###"
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
      Locked          =   -1  'True
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel"
      Height          =   360
      Index           =   1
      Left            =   4665
      TabIndex        =   9
      Top             =   1980
      Width           =   1200
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   360
      Index           =   0
      Left            =   3360
      TabIndex        =   8
      Top             =   1980
      Width           =   1200
   End
   Begin rdActiveText.ActiveText ActiveText3 
      Height          =   330
      Left            =   1440
      TabIndex        =   5
      Top             =   1050
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      Alignment       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   10
      Text            =   "R$ 0,00"
      TextMask        =   4
      RawText         =   4
      FloatFormat     =   2
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText ActiveText2 
      Height          =   330
      Left            =   1440
      TabIndex        =   3
      Top             =   585
      Width           =   1275
      _ExtentX        =   2249
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
      MaxLength       =   10
      TextMask        =   1
      RawText         =   1
      Mask            =   "##/##/####"
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText ActiveText1 
      Height          =   330
      Left            =   1440
      TabIndex        =   1
      Top             =   150
      Width           =   4380
      _ExtentX        =   7726
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
      MaxLength       =   60
      Text            =   "Marcelo Leal Limaverde Cabral"
      TextCase        =   3
      RawText         =   0
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin VB.Label Label1 
      Caption         =   "ZIP Code:"
      Height          =   270
      Index           =   3
      Left            =   180
      TabIndex        =   6
      Top             =   1530
      Width           =   1155
   End
   Begin VB.Label Label1 
      Caption         =   "Income:"
      Height          =   270
      Index           =   2
      Left            =   195
      TabIndex        =   4
      Top             =   1065
      Width           =   1155
   End
   Begin VB.Label Label1 
      Caption         =   "Birth Date:"
      Height          =   270
      Index           =   1
      Left            =   195
      TabIndex        =   2
      Top             =   630
      Width           =   1155
   End
   Begin VB.Label Label1 
      Caption         =   "Full Name:"
      Height          =   270
      Index           =   0
      Left            =   195
      TabIndex        =   0
      Top             =   210
      Width           =   1155
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click(Index As Integer)
    Unload Me
End Sub

VERSION 5.00
Object = "{DB44097B-F1FF-11D3-8DCD-F6D387C1003E}#1.0#0"; "ActivePanel.ocx"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ActivePanel Example"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6225
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   6225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin rdActivePanel.ActivePanel ActivePanel6 
      Height          =   2190
      Left            =   4290
      TabIndex        =   5
      Top             =   2835
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   3863
      BorderStyle     =   2
      Caption         =   "ActivePanel6"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
      CaptionPos      =   7
   End
   Begin rdActivePanel.ActivePanel ActivePanel5 
      Height          =   2190
      Left            =   2265
      TabIndex        =   4
      Top             =   2850
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   3863
      BorderStyle     =   1
      Caption         =   "ActivePanel5"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
      CaptionPos      =   5
   End
   Begin rdActivePanel.ActivePanel ActivePanel4 
      Height          =   2190
      Left            =   210
      TabIndex        =   3
      Top             =   2850
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   3863
      BorderStyle     =   9
      Caption         =   "ActivePanel4"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
      CaptionPos      =   0
   End
   Begin rdActivePanel.ActiveLine ActiveLine1 
      Height          =   195
      Left            =   195
      Top             =   2535
      Width           =   5850
      _ExtentX        =   10319
      _ExtentY        =   344
      Caption         =   "ActiveLine1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin rdActivePanel.ActivePanel ActivePanel3 
      Height          =   2190
      Left            =   4245
      TabIndex        =   2
      Top             =   165
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   3863
      BorderStyle     =   7
      Caption         =   "ActivePanel3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
      CaptionPos      =   8
   End
   Begin rdActivePanel.ActivePanel ActivePanel2 
      Height          =   2190
      Left            =   2220
      TabIndex        =   1
      Top             =   165
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   3863
      BorderStyle     =   5
      Caption         =   "ActivePanel2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
      CaptionPos      =   6
   End
   Begin rdActivePanel.ActivePanel ActivePanel1 
      Height          =   2190
      Left            =   180
      TabIndex        =   0
      Top             =   165
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   3863
      Caption         =   "ActivePanel1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


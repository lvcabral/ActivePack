VERSION 5.00
Object = "{A7D577E5-E727-11D3-8DCD-92EC485CE63E}#1.0#0"; "ActiveLink.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ActiveLink Example"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3645
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   3645
   StartUpPosition =   2  'CenterScreen
   Begin rdActiveLink.ActiveLink ActiveLink4 
      Height          =   285
      Left            =   165
      Top             =   1500
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   503
      ForeColor       =   192
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Link            =   "Notepad.exe"
      MousePointer    =   99
      MouseIcon       =   "Form1.frx":08CA
      Caption         =   "Run Notepad..."
   End
   Begin rdActiveLink.ActiveLink ActiveLink3 
      Height          =   300
      Left            =   1020
      Top             =   1020
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   529
      ForeColor       =   49152
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Link            =   "c:\"
      MousePointer    =   99
      MouseIcon       =   "Form1.frx":0BE4
      Caption         =   "Drive C:"
      BackColor       =   -2147483643
   End
   Begin rdActiveLink.ActiveLink ActiveLink2 
      Height          =   255
      Left            =   150
      Top             =   540
      Width           =   2640
      _ExtentX        =   4657
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Link            =   "mailto:activepack@xoommail.com"
      MousePointer    =   99
      MouseIcon       =   "Form1.frx":0EFE
      Caption         =   "activepack@xoommail.com"
   End
   Begin rdActiveLink.ActiveLink ActiveLink1 
      Height          =   330
      Left            =   150
      Top             =   135
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   582
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   99
      MouseIcon       =   "Form1.frx":1218
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "Try to open Drive C:  to view files."
      Height          =   300
      Left            =   150
      TabIndex        =   0
      Top             =   1020
      Width           =   3105
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


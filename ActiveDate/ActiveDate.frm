VERSION 5.00
Begin VB.Form frmActiveDate 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2490
   ClientLeft      =   6930
   ClientTop       =   2265
   ClientWidth     =   2325
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "ActiveDate.frx":0000
   KeyPreview      =   -1  'True
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2490
   ScaleWidth      =   2325
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2490
      Left            =   0
      ScaleHeight     =   2490
      ScaleWidth      =   2325
      TabIndex        =   0
      Top             =   0
      Width           =   2325
      Begin VB.CommandButton btToday 
         Appearance      =   0  'Flat
         Caption         =   "&Hoje"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1020
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   2100
         Width           =   900
      End
      Begin VB.CommandButton BotaoProx 
         Appearance      =   0  'Flat
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   2100
         Width           =   300
      End
      Begin VB.CommandButton BotaoAnt 
         Appearance      =   0  'Flat
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   720
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   2100
         Width           =   300
      End
      Begin VB.Label TitMesAno 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "Fevereiro 1999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   105
         TabIndex        =   5
         Top             =   105
         Width           =   2115
      End
      Begin VB.Label Casa 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   37
         Left            =   405
         TabIndex        =   6
         Top             =   2085
         Width           =   315
      End
      Begin VB.Label Casa 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Index           =   36
         Left            =   105
         TabIndex        =   7
         Top             =   2085
         Width           =   315
      End
      Begin VB.Label Casa 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   35
         Left            =   1905
         TabIndex        =   8
         Top             =   1800
         Width           =   315
      End
      Begin VB.Label Casa 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   34
         Left            =   1605
         TabIndex        =   9
         Top             =   1800
         Width           =   315
      End
      Begin VB.Label Casa 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   33
         Left            =   1305
         TabIndex        =   10
         Top             =   1800
         Width           =   315
      End
      Begin VB.Label Casa 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   32
         Left            =   1005
         TabIndex        =   11
         Top             =   1800
         Width           =   315
      End
      Begin VB.Label Casa 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   31
         Left            =   705
         TabIndex        =   12
         Top             =   1800
         Width           =   315
      End
      Begin VB.Label Casa 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   30
         Left            =   405
         TabIndex        =   13
         Top             =   1800
         Width           =   315
      End
      Begin VB.Label Casa 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Index           =   29
         Left            =   105
         TabIndex        =   14
         Top             =   1800
         Width           =   315
      End
      Begin VB.Label Casa 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   28
         Left            =   1905
         TabIndex        =   15
         Top             =   1515
         Width           =   315
      End
      Begin VB.Label Casa 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   27
         Left            =   1605
         TabIndex        =   16
         Top             =   1515
         Width           =   315
      End
      Begin VB.Label Casa 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   26
         Left            =   1305
         TabIndex        =   17
         Top             =   1515
         Width           =   315
      End
      Begin VB.Label Casa 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   25
         Left            =   1005
         TabIndex        =   18
         Top             =   1515
         Width           =   315
      End
      Begin VB.Label Casa 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   24
         Left            =   705
         TabIndex        =   19
         Top             =   1515
         Width           =   315
      End
      Begin VB.Label Casa 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   23
         Left            =   405
         TabIndex        =   20
         Top             =   1515
         Width           =   315
      End
      Begin VB.Label Casa 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Index           =   22
         Left            =   105
         TabIndex        =   21
         Top             =   1515
         Width           =   315
      End
      Begin VB.Label Casa 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   21
         Left            =   1905
         TabIndex        =   22
         Top             =   1230
         Width           =   315
      End
      Begin VB.Label Casa 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   20
         Left            =   1605
         TabIndex        =   23
         Top             =   1230
         Width           =   315
      End
      Begin VB.Label Casa 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   19
         Left            =   1305
         TabIndex        =   24
         Top             =   1230
         Width           =   315
      End
      Begin VB.Label Casa 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   18
         Left            =   1005
         TabIndex        =   25
         Top             =   1230
         Width           =   315
      End
      Begin VB.Label Casa 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   17
         Left            =   705
         TabIndex        =   26
         Top             =   1230
         Width           =   315
      End
      Begin VB.Label Casa 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   16
         Left            =   405
         TabIndex        =   27
         Top             =   1230
         Width           =   315
      End
      Begin VB.Label Casa 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Index           =   15
         Left            =   105
         TabIndex        =   28
         Top             =   1230
         Width           =   315
      End
      Begin VB.Label Casa 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   14
         Left            =   1905
         TabIndex        =   29
         Top             =   945
         Width           =   315
      End
      Begin VB.Label Casa 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   13
         Left            =   1605
         TabIndex        =   30
         Top             =   945
         Width           =   315
      End
      Begin VB.Label Casa 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   12
         Left            =   1305
         TabIndex        =   31
         Top             =   945
         Width           =   315
      End
      Begin VB.Label Casa 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   11
         Left            =   1005
         TabIndex        =   32
         Top             =   945
         Width           =   315
      End
      Begin VB.Label Casa 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   10
         Left            =   705
         TabIndex        =   33
         Top             =   945
         Width           =   315
      End
      Begin VB.Label Casa 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   9
         Left            =   405
         TabIndex        =   34
         Top             =   945
         Width           =   315
      End
      Begin VB.Label Casa 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Index           =   8
         Left            =   105
         TabIndex        =   35
         Top             =   945
         Width           =   315
      End
      Begin VB.Label Casa 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   7
         Left            =   1905
         TabIndex        =   36
         Top             =   660
         Width           =   315
      End
      Begin VB.Label Casa 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   6
         Left            =   1605
         TabIndex        =   37
         Top             =   660
         Width           =   315
      End
      Begin VB.Label Casa 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   5
         Left            =   1305
         TabIndex        =   38
         Top             =   660
         Width           =   315
      End
      Begin VB.Label Casa 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   4
         Left            =   1005
         TabIndex        =   39
         Top             =   660
         Width           =   315
      End
      Begin VB.Label Casa 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   3
         Left            =   705
         TabIndex        =   40
         Top             =   660
         Width           =   315
      End
      Begin VB.Label Casa 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   2
         Left            =   405
         TabIndex        =   41
         Top             =   660
         Width           =   315
      End
      Begin VB.Label Casa 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Index           =   1
         Left            =   105
         TabIndex        =   1
         Top             =   660
         Width           =   315
      End
      Begin VB.Label lbDiaSem 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   7
         Left            =   1905
         TabIndex        =   42
         Top             =   420
         Width           =   315
      End
      Begin VB.Label lbDiaSem 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   6
         Left            =   1605
         TabIndex        =   43
         Top             =   420
         Width           =   315
      End
      Begin VB.Label lbDiaSem 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Q"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   5
         Left            =   1305
         TabIndex        =   44
         Top             =   420
         Width           =   315
      End
      Begin VB.Label lbDiaSem 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Q"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   1005
         TabIndex        =   45
         Top             =   420
         Width           =   315
      End
      Begin VB.Label lbDiaSem 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "T"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   705
         TabIndex        =   46
         Top             =   420
         Width           =   315
      End
      Begin VB.Label lbDiaSem 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   405
         TabIndex        =   47
         Top             =   420
         Width           =   315
      End
      Begin VB.Label lbDiaSem 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   105
         TabIndex        =   48
         Top             =   420
         Width           =   315
      End
   End
End
Attribute VB_Name = "frmActiveDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public oComboDate As TextBox

Dim IndMes As Integer
Dim NomeMes As String
Dim NumDias As Integer
Dim Ano As Integer
Dim DiaDaSemana As Integer

Private Sub BotaoAnt_Click()
    MesAnt
    picBack.SetFocus
End Sub

Private Sub BotaoProx_Click()
    MesProx
    picBack.SetFocus
End Sub

Private Sub btToday_Click()
   MesAtual
   picBack.SetFocus
End Sub

Private Sub Casa_Click(Index As Integer)
   oComboDate.Text = Casa(Index).Tag
   Unload Me
End Sub

Private Sub ErroAno()
   Dim m$
   m$ = "Ano inválido. O Ano deve estar no" + Chr$(13) + Chr$(10)
   m$ = m$ + "intervalo de 1753 a 2078, inclusive."
   MsgBox m$, 48
End Sub

Private Sub EscreveMes()
   Dim n%, p%
   TitMesAno.Caption = StrConv(NomeMes, vbProperCase) + " " + Str$(Ano)
   For n% = 1 To 37
      Casa(n%).Caption = " "
      Casa(n%).BackColor = vbWhite
      Casa(n%).Enabled = False
   Next
   p% = DiaDaSemana
   For n% = 1 To NumDias
      Casa(p%).Caption = Format$(n%)
      Casa(p%).Tag = DateSerial(Ano, IndMes, n%)
      Casa(p%).Enabled = True
      If n% = Day(Now) And IndMes = Month(Now) And Ano = Year(Now) Then
         Casa(p%).BackColor = vbYellow
      Else
         Casa(p%).BackColor = vbWhite
      End If
      p% = p% + 1
   Next
   If IsDate(oComboDate.Text) Then
      If IndMes = Month(oComboDate.Text) And Ano = Year(oComboDate.Text) Then
         Casa(DiaDaSemana + Day(oComboDate.Text) - 1).BackColor = QBColor(7)
      End If
   End If
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Deactivate()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Or KeyCode = vbKeyF4 Then
        Unload Me
    ElseIf KeyCode = vbKeyReturn Then
        If Len(oComboDate.Text) = 0 Then oComboDate.Text = Date
        Unload Me
    End If
End Sub

Private Sub MesAnt()
    If IndMes = 1 Then
       If Ano = 1753 Then
          ErroAno
          Exit Sub
       Else
          Ano = Ano - 1
          IndMes = 12
       End If
    Else
       IndMes = IndMes - 1
    End If
    FillCalendar DateSerial(Ano, IndMes, 1)
End Sub

Private Sub MesAtual()
    FillCalendar Now
End Sub

Private Sub MesProx()
    If IndMes = 12 Then
       If Ano = 2078 Then
          ErroAno
          Exit Sub
       Else
          Ano = Ano + 1
          IndMes = 1
       End If
    Else
       IndMes = IndMes + 1
    End If
    FillCalendar DateSerial(Ano, IndMes, 1)
End Sub

Private Function Bissexto(ByVal Ano As Integer)
    If (Ano Mod 4 = 0 And Ano Mod 100 <> 0) Or (Ano Mod 400 = 0) Then
       Bissexto = -1
    Else
       Bissexto = 0
    End If
End Function

Public Sub FillCalendar(ByVal vDate As Date)
   NomeMes = Format$(vDate, "mmmm", vbUseSystemDayOfWeek)
   IndMes = Month(vDate)
   Ano = Year(vDate)
   Select Case IndMes
   Case 1
      NumDias = 31
   Case 2
      If Bissexto(Ano) Then
         NumDias = 29
      Else
         NumDias = 28
      End If
   Case 3
      NumDias = 31
   Case 4
      NumDias = 30
   Case 5
      NumDias = 31
   Case 6
      NumDias = 30
   Case 7
      NumDias = 31
   Case 8
      NumDias = 31
   Case 9
      NumDias = 30
   Case 10
      NumDias = 31
   Case 11
      NumDias = 30
   Case 12
      NumDias = 31
   End Select
   DiaDaSemana = Weekday(DateSerial(Ano, IndMes, 1))
   EscreveMes
End Sub

Private Sub Form_Load()
Dim s%
    For s% = 1 To 7
        lbDiaSem(s%) = UCase$(Left$(WeekdayName(s%, True, vbSunday), 1))
    Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    oComboDate.SetFocus
    Set oComboDate = Nothing
End Sub

Private Sub TitMesAno_Click()
    Unload Me
End Sub

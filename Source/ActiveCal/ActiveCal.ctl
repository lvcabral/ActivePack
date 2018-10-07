VERSION 5.00
Begin VB.UserControl ActiveCalendar 
   ClientHeight    =   3540
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3720
   ForwardFocus    =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   3540
   ScaleWidth      =   3720
   ToolboxBitmap   =   "ActiveCal.ctx":0000
   Begin VB.CommandButton btMove 
      Height          =   405
      Index           =   1
      Left            =   3075
      Picture         =   "ActiveCal.ctx":00FA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2970
      UseMaskColor    =   -1  'True
      Width           =   465
   End
   Begin VB.CommandButton btHoje 
      Caption         =   "Hoje"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1635
      TabIndex        =   1
      Top             =   2970
      Width           =   1425
   End
   Begin VB.CommandButton btMove 
      Height          =   405
      Index           =   0
      Left            =   1155
      Picture         =   "ActiveCal.ctx":0204
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2970
      UseMaskColor    =   -1  'True
      Width           =   465
   End
   Begin VB.Label Casa 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   37
      Left            =   645
      TabIndex        =   47
      Top             =   2940
      Width           =   495
   End
   Begin VB.Label Casa 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Index           =   36
      Left            =   165
      TabIndex        =   46
      Top             =   2940
      Width           =   495
   End
   Begin VB.Label Casa 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Index           =   35
      Left            =   3045
      TabIndex        =   45
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Casa 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   34
      Left            =   2565
      TabIndex        =   44
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Casa 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   33
      Left            =   2085
      TabIndex        =   43
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Casa 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   32
      Left            =   1605
      TabIndex        =   42
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Casa 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   31
      Left            =   1125
      TabIndex        =   41
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Casa 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   30
      Left            =   645
      TabIndex        =   40
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Casa 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Index           =   29
      Left            =   165
      TabIndex        =   39
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Casa 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Index           =   28
      Left            =   3045
      TabIndex        =   38
      Top             =   2100
      Width           =   495
   End
   Begin VB.Label Casa 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   27
      Left            =   2565
      TabIndex        =   37
      Top             =   2100
      Width           =   495
   End
   Begin VB.Label Casa 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   26
      Left            =   2085
      TabIndex        =   36
      Top             =   2100
      Width           =   495
   End
   Begin VB.Label Casa 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   25
      Left            =   1605
      TabIndex        =   35
      Top             =   2100
      Width           =   495
   End
   Begin VB.Label Casa 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   24
      Left            =   1125
      TabIndex        =   34
      Top             =   2100
      Width           =   495
   End
   Begin VB.Label Casa 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   23
      Left            =   645
      TabIndex        =   33
      Top             =   2100
      Width           =   495
   End
   Begin VB.Label Casa 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Index           =   22
      Left            =   165
      TabIndex        =   32
      Top             =   2100
      Width           =   495
   End
   Begin VB.Label Casa 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Index           =   21
      Left            =   3045
      TabIndex        =   31
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Casa 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   20
      Left            =   2565
      TabIndex        =   30
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Casa 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   19
      Left            =   2085
      TabIndex        =   29
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Casa 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   18
      Left            =   1605
      TabIndex        =   28
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Casa 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   17
      Left            =   1125
      TabIndex        =   27
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Casa 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   16
      Left            =   645
      TabIndex        =   26
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Casa 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Index           =   15
      Left            =   165
      TabIndex        =   25
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Casa 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Index           =   14
      Left            =   3045
      TabIndex        =   24
      Top             =   1260
      Width           =   495
   End
   Begin VB.Label Casa 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   13
      Left            =   2565
      TabIndex        =   23
      Top             =   1260
      Width           =   495
   End
   Begin VB.Label Casa 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   12
      Left            =   2085
      TabIndex        =   22
      Top             =   1260
      Width           =   495
   End
   Begin VB.Label Casa 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   11
      Left            =   1605
      TabIndex        =   21
      Top             =   1260
      Width           =   495
   End
   Begin VB.Label Casa 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   10
      Left            =   1125
      TabIndex        =   20
      Top             =   1260
      Width           =   495
   End
   Begin VB.Label Casa 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   9
      Left            =   645
      TabIndex        =   19
      Top             =   1260
      Width           =   495
   End
   Begin VB.Label Casa 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Index           =   8
      Left            =   165
      TabIndex        =   18
      Top             =   1260
      Width           =   495
   End
   Begin VB.Label Casa 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Index           =   7
      Left            =   3045
      TabIndex        =   17
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Casa 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   6
      Left            =   2565
      TabIndex        =   16
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Casa 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   5
      Left            =   2085
      TabIndex        =   15
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Casa 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   4
      Left            =   1605
      TabIndex        =   14
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Casa 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   3
      Left            =   1125
      TabIndex        =   13
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Casa 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Index           =   2
      Left            =   645
      TabIndex        =   12
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Casa 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   435
      Index           =   1
      Left            =   165
      TabIndex        =   11
      Top             =   840
      Width           =   495
   End
   Begin VB.Label lbDiaSem 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SAB"
      BeginProperty Font 
         Name            =   "Arial"
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
      Left            =   3045
      TabIndex        =   10
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lbDiaSem 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SEX"
      BeginProperty Font 
         Name            =   "Arial"
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
      Left            =   2565
      TabIndex        =   9
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lbDiaSem 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "QUI"
      BeginProperty Font 
         Name            =   "Arial"
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
      Left            =   2085
      TabIndex        =   8
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lbDiaSem 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "QUA"
      BeginProperty Font 
         Name            =   "Arial"
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
      Left            =   1605
      TabIndex        =   7
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lbDiaSem 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TER"
      BeginProperty Font 
         Name            =   "Arial"
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
      Left            =   1125
      TabIndex        =   6
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lbDiaSem 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SEG"
      BeginProperty Font 
         Name            =   "Arial"
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
      Left            =   645
      TabIndex        =   5
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lbDiaSem 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DOM"
      BeginProperty Font 
         Name            =   "Arial"
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
      Left            =   165
      TabIndex        =   4
      Top             =   600
      Width           =   495
   End
   Begin VB.Label TitMesAno 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "Fevereiro 1993"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   150
      TabIndex        =   3
      Top             =   120
      Width           =   3390
   End
End
Attribute VB_Name = "ActiveCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Const AMARELO = &HFFFF&
Const BRANCO = &HFFFFFF
Const CINZA = &HC0C0C0

'Elastic
Private bInitializing As Boolean
Private lWidth As Long, lHeight As Long
Private iFormHeight As Integer, iFormWidth As Integer, iNumOfControls As Integer
Private iTop() As Integer, iLeft() As Integer, iHeight() As Integer, iWidth() As Integer, iFontSize() As Integer, iRightMargin() As Integer
Private bFirstTime As Boolean

Enum BorderStyleOptions
    None
    FixedSingle
End Enum
    
Dim IndMes As Integer
Dim NomeMes As String
Dim NumDias As Integer
Dim Ano As Integer
Dim AnoMom As Integer
Dim MesMom As Integer
Dim DiaDaSemana As Integer
'Property Variables:
Dim m_Date As Date
'Properties Default Constants
Const m_def_TodayCaption = "&Hoje"
'Event Declarations:
Event Click() 'MappingInfo=Casa(1),Casa,1,DblClick
Event DblClick() 'MappingInfo=Casa(1),Casa,1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."

Private Function Bissexto(ByVal Ano As Integer)
    If (Ano Mod 4 = 0 And Ano Mod 100 <> 0) Or (Ano Mod 400 = 0) Then
       Bissexto = -1
    Else
       Bissexto = 0
    End If
End Function

Public Property Let BorderStyle(ByVal mNewBorderStyle As BorderStyleOptions)
    UserControl.BorderStyle = mNewBorderStyle
    PropertyChanged "BorderStyle"
End Property

Public Property Get BorderStyle() As BorderStyleOptions
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Get TodayCaption() As String
    TodayCaption = btHoje.Caption
End Property

Public Property Let TodayCaption(ByVal New_TodayCaption As String)
    btHoje.Caption = New_TodayCaption
    PropertyChanged "TodayCaption"
End Property

Private Sub btHoje_Click()
    MesAtual
End Sub

Private Sub btMove_Click(Index As Integer)
   If Index = 0 Then MesAnt Else MesProx
End Sub

Private Sub CalculaMes(ByVal vDate As Date)
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

Private Sub ErroAno()
   Dim m$
   m$ = "Ano inválido. O Ano deve estar no" + Chr$(13) + Chr$(10)
   m$ = m$ + "intervalo de 1753 a 2078, inclusive."
   MsgBox m$, 48
End Sub

Private Sub EscreveMes()
   Dim n%, p%, pp%
   TitMesAno.Caption = StrConv(NomeMes, vbProperCase) + " " + Str$(Ano)
   For n% = 1 To 37
      Casa(n%).Caption = " "
      Casa(n%).BackColor = BRANCO
   Next
   p% = DiaDaSemana
   pp% = p% - 1
   For n% = 1 To NumDias
      Casa(p%).Caption = Trim$(Str$(n%))
      If n% = Day(Now) And IndMes = Month(Now) And Ano = Year(Now) Then
         Casa(n% + pp%).BackColor = AMARELO
      ElseIf n% = Day(m_Date) And IndMes = Month(m_Date) And Ano = Year(m_Date) Then
         Casa(n% + pp%).BackColor = CINZA
      Else
         Casa(n% + pp%).BackColor = BRANCO
      End If
      p% = p% + 1
   Next
   MesMom = IndMes
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
    CalculaMes DateSerial(Ano, IndMes, 1)
End Sub

Private Sub MesAtual()
    CalculaMes Now
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
    CalculaMes DateSerial(Ano, IndMes, 1)
End Sub

Private Sub Casa_Click(Index As Integer)
Static mOldIndex As Integer
    If Val(Casa(Index).Caption) > 0 Then
        If mOldIndex >= 1 Then
            If Val(Casa(mOldIndex).Caption) = Day(Now) And IndMes = Month(Now) And Ano = Year(Now) Then
                Casa(mOldIndex).BackColor = AMARELO
            Else
                Casa(mOldIndex).BackColor = BRANCO
            End If
        End If
        Casa(Index).BackColor = CINZA
        mOldIndex = Index
        SelectedDate = DateSerial(Ano, IndMes, Val(Casa(Index).Caption))
        RaiseEvent Click
    End If
End Sub

Private Sub UserControl_Initialize()
Dim s%
    For s% = 1 To 7
        lbDiaSem(s%) = UCase$(WeekdayName(s%, True, vbSunday))
    Next
    iFormHeight = UserControl.Height
    iFormWidth = UserControl.Width
    bFirstTime = True
    MesAtual
End Sub

Private Sub Casa_DblClick(Index As Integer)
    If Val(Casa(Index).Caption) > 0 Then
        'SelectedDate = DateSerial(Ano, IndMes, Val(Casa(Index).Caption))
        RaiseEvent DblClick
    End If
End Sub

Public Property Get SelectedDate() As Date
Attribute SelectedDate.VB_MemberFlags = "2c"
    SelectedDate = m_Date
End Property

Public Property Let SelectedDate(ByVal New_Date As Date)
    If UserControl.CanPropertyChange("SelectedDate") Then
        m_Date = New_Date
        CalculaMes m_Date
        PropertyChanged "SelectedDate"
    End If
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    bInitializing = True
    lWidth = Width
    lHeight = Height
    m_Date = CDate(Int(Now))
    btHoje.Caption = m_def_TodayCaption
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Date = PropBag.ReadProperty("Date", CDate(CLng(Now)))
    btHoje.Caption = PropBag.ReadProperty("TodayCaption", m_def_TodayCaption)
    btHoje.ToolTipText = Format(Now, "Long Date")
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", FixedSingle)
End Sub

Private Sub UserControl_Resize()
    If bInitializing Then
        Width = lWidth
        Height = lHeight
    Else
        ControlResize
    End If
End Sub

Private Sub UserControl_Show()
    bInitializing = False
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Date", m_Date, CDate(CLng(Now)))
    Call PropBag.WriteProperty("TodayCaption", btHoje.Caption, m_def_TodayCaption)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, FixedSingle)
End Sub

Private Sub ControlResize()
Dim I As Integer, Inc As Integer
Dim RatioX As Double, RatioY As Double
Dim SaveRedraw%
    On Error Resume Next
    SaveRedraw% = UserControl.AutoRedraw
    
    UserControl.AutoRedraw = True
    
    If bFirstTime Then
        Init
        bFirstTime = False
        'Exit Sub
    End If
    
    If UserControl.Height < iFormHeight / 2 Then UserControl.Height = iFormHeight / 2
    If UserControl.Width < iFormWidth / 2 Then UserControl.Width = iFormWidth / 2
    RatioY = 1# * iFormHeight / UserControl.Height
    RatioX = 1# * iFormWidth / UserControl.Width
    On Error Resume Next ' for comboboxes, timeres and other nonsizible controls

    For I = 0 To iNumOfControls
        UserControl.Controls(I).Visible = False
        UserControl.Controls(I).Top = Int(iTop(I) / RatioY)
        UserControl.Controls(I).Left = Int(iLeft(I) / RatioX)
        UserControl.Controls(I).Height = Int(iHeight(I) / RatioY)
        UserControl.Controls(I).Width = Int(iWidth(I) / RatioX)
        UserControl.Controls(I).Font.Size = Int(iFontSize(I) / RatioX) + Int(iFontSize(I) / RatioX) Mod 2
        UserControl.Controls(I).Visible = True
    Next

    UserControl.AutoRedraw = SaveRedraw%
    Exit Sub
End Sub

Private Sub Init()
    Dim I As Integer
    iNumOfControls = UserControl.Controls.Count - 1
    ReDim iTop(iNumOfControls)
    ReDim iLeft(iNumOfControls)
    ReDim iHeight(iNumOfControls)
    ReDim iWidth(iNumOfControls)
    ReDim iFontSize(iNumOfControls)
    ReDim iRightMargin(iNumOfControls)
    On Error Resume Next
    For I = 0 To iNumOfControls
        If TypeOf UserControl.Controls(I) Is Line Then
            iTop(I) = UserControl.Controls(I).Y1
            iLeft(I) = UserControl.Controls(I).X1
            iHeight(I) = UserControl.Controls(I).Y2
            iWidth(I) = UserControl.Controls(I).X2
        Else
            iTop(I) = UserControl.Controls(I).Top
            iLeft(I) = UserControl.Controls(I).Left
            iHeight(I) = UserControl.Controls(I).Height
            iWidth(I) = UserControl.Controls(I).Width
            iFontSize(I) = UserControl.Controls(I).Font.Size
            iRightMargin(I) = UserControl.Controls(I).RightMargin
        End If
    Next
End Sub

Public Sub AboutBox()
Attribute AboutBox.VB_UserMemId = -552
Attribute AboutBox.VB_MemberFlags = "40"
    frmAbout.Show vbModal
End Sub

VERSION 5.00
Begin VB.Form frmFilesFilter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Filtro de Arquivos"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4935
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FilesFilter.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   3570
      TabIndex        =   8
      Top             =   1965
      Width           =   1245
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   2250
      TabIndex        =   7
      Top             =   1965
      Width           =   1260
   End
   Begin VB.Frame fraOptions 
      Caption         =   "Opções"
      Height          =   1770
      Left            =   105
      TabIndex        =   9
      Top             =   90
      Width           =   4710
      Begin VB.TextBox txtOther 
         Height          =   300
         Left            =   2415
         TabIndex        =   6
         Top             =   1072
         Width           =   2130
      End
      Begin VB.CheckBox chkOther 
         Appearance      =   0  'Flat
         Caption         =   "Outros:"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   1500
         TabIndex        =   5
         Top             =   1110
         Width           =   915
      End
      Begin VB.CheckBox chkOLB 
         Appearance      =   0  'Flat
         Caption         =   "*.OLB"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   1500
         TabIndex        =   3
         Top             =   750
         Width           =   915
      End
      Begin VB.CheckBox chkEXE 
         Appearance      =   0  'Flat
         Caption         =   "*.EXE"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   1500
         TabIndex        =   1
         Top             =   405
         Width           =   915
      End
      Begin VB.CheckBox chkOCX 
         Appearance      =   0  'Flat
         Caption         =   "*.OCX"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   225
         TabIndex        =   0
         Top             =   405
         Width           =   915
      End
      Begin VB.CheckBox chkTLB 
         Appearance      =   0  'Flat
         Caption         =   "*.TLB"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   225
         TabIndex        =   4
         Top             =   1095
         Width           =   915
      End
      Begin VB.CheckBox chkDLL 
         Appearance      =   0  'Flat
         Caption         =   "*.DLL"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   225
         TabIndex        =   2
         Top             =   750
         Width           =   915
      End
   End
End
Attribute VB_Name = "frmFilesFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private sTemp As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    sTemp = ""
    If chkOCX.Value = vbChecked Then
        sTemp = sTemp & "*.ocx;"
    End If
    If chkDLL.Value = vbChecked Then
        sTemp = sTemp & "*.dll;"
    End If
    If chkTLB.Value = vbChecked Then
        sTemp = sTemp & "*.tlb;"
    End If
    If chkEXE.Value = vbChecked Then
        sTemp = sTemp & "*.exe;"
    End If
    If chkOLB.Value = vbChecked Then
        sTemp = sTemp & "*.olb;"
    End If
    If chkOther.Value = vbChecked Then
        If InStr(txtOther.Text, ".") > 0 Then
            sTemp = sTemp & txtOther.Text
        ElseIf sTemp <> "" Then
            MsgBox LoadResString(201), vbInformation
            txtOther.SetFocus
            Exit Sub
        End If
    End If
    If Len(sTemp) = 0 Then
        MsgBox LoadResString(212), vbExclamation
    Else
        sTemp = Replace(sTemp, ";;", ";")
        sTemp = IIf(Right(sTemp, 1) = ";", Left(sTemp, Len(sTemp) - 1), sTemp)
        frmActiveTools.sFiles = sTemp
        Unload Me
    End If
End Sub

Private Sub Form_Load()
Dim pos As Integer
    Caption = LoadResString(208)
    fraOptions.Caption = LoadResString(309)
    chkOther.Caption = LoadResString(310)
    cmdOK.Caption = LoadResString(311)
    cmdCancel.Caption = LoadResString(312)
    sTemp = LCase(frmActiveTools.sFiles) & ";"
    pos = InStr(";" & sTemp, ";*.ocx;")
    If pos > 0 Then
        chkOCX.Value = vbChecked
        sTemp = Replace(sTemp, "*.ocx;", "")
    End If
    pos = InStr(";" & sTemp, ";*.dll;")
    If pos > 0 Then
        chkDLL.Value = vbChecked
        sTemp = Replace(sTemp, "*.dll;", "")
    End If
    pos = InStr(";" & sTemp, ";*.tlb;")
    If pos > 0 Then
        chkTLB.Value = vbChecked
        sTemp = Replace(sTemp, "*.tlb;", "")
    End If
    pos = InStr(";" & sTemp, ";*.exe;")
    If pos > 0 Then
        chkEXE.Value = vbChecked
        sTemp = Replace(sTemp, "*.exe;", "")
    End If
    pos = InStr(";" & sTemp, ";*.olb;")
    If pos > 0 Then
        chkOLB.Value = vbChecked
        sTemp = Replace(sTemp, "*.olb;", "")
    End If
    If InStr(sTemp, ".") > 0 Then
        chkOther.Value = vbChecked
        sTemp = Replace(sTemp, ";;", ";")
        sTemp = IIf(Right(sTemp, 1) = ";", Left(sTemp, Len(sTemp) - 1), sTemp)
        txtOther.Text = sTemp
    End If
End Sub

Private Sub txtOther_Change()
    If Len(Trim(txtOther.Text)) > 0 Then
        chkOther.Value = vbChecked
    End If
End Sub

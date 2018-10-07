VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmActiveTools 
   Caption         =   "ActiveX Tools"
   ClientHeight    =   4920
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7425
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ActiveTools.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4920
   ScaleWidth      =   7425
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picImages256 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      HasDC           =   0   'False
      Height          =   375
      Left            =   960
      ScaleHeight     =   375
      ScaleWidth      =   2670
      TabIndex        =   5
      Top             =   3900
      Visible         =   0   'False
      Width           =   2670
      Begin VB.Image imgExplore 
         Height          =   240
         Left            =   2235
         Picture         =   "ActiveTools.frx":0442
         Top             =   60
         Width           =   240
      End
      Begin VB.Image imgPrinter 
         Height          =   240
         Left            =   1500
         Picture         =   "ActiveTools.frx":09CC
         Top             =   60
         Width           =   240
      End
      Begin VB.Image imgLupa 
         Height          =   240
         Left            =   60
         Picture         =   "ActiveTools.frx":0F56
         Top             =   60
         Width           =   240
      End
      Begin VB.Image imgDisk 
         Height          =   240
         Left            =   1140
         Picture         =   "ActiveTools.frx":14E0
         Top             =   60
         Width           =   240
      End
      Begin VB.Image imgFiles 
         Height          =   240
         Left            =   780
         Picture         =   "ActiveTools.frx":1A6A
         Top             =   60
         Width           =   240
      End
      Begin VB.Image imgFolder 
         Height          =   240
         Left            =   420
         Picture         =   "ActiveTools.frx":1FF4
         Top             =   60
         Width           =   240
      End
      Begin VB.Image imgInfo 
         Height          =   240
         Left            =   1860
         Picture         =   "ActiveTools.frx":257E
         Top             =   60
         Width           =   240
      End
   End
   Begin VB.PictureBox picAnim 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   930
      Left            =   2940
      ScaleHeight     =   930
      ScaleWidth      =   1035
      TabIndex        =   3
      Top             =   1935
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSComctlLib.Toolbar barMain 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   7425
      _ExtentX        =   13097
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   20
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   200
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Find"
            Object.ToolTipText     =   "Iniciar a Procura"
            ImageKey        =   "Find"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Folder"
            Object.ToolTipText     =   "Selecionar a Pasta de Origem"
            ImageKey        =   "FillFolder"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Files"
            Object.ToolTipText     =   "Selecionar Filtro de Arquivos"
            ImageKey        =   "WinFile"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Save"
            Object.ToolTipText     =   "Salvar Resultado da Pesquisa"
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Print"
            Object.ToolTipText     =   "Imprimir Resultado da Pesquisa"
            ImageKey        =   "Print"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Prop"
            Object.ToolTipText     =   "Propriedades do Arquivo"
            ImageKey        =   "Prop"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "ObjBrowser"
            Object.ToolTipText     =   "Object Browser"
            ImageKey        =   "BROWSER"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Explore"
            Object.ToolTipText     =   "Explorar pasta do arquivo"
            ImageKey        =   "Explorer"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "CheckAll"
            Object.ToolTipText     =   "Marcar Todos os Arquivos"
            ImageKey        =   "Check"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "UnCheckAll"
            Object.ToolTipText     =   "Desmarcar Todos os Arquivos"
            ImageKey        =   "UnCheck"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Register"
            Object.ToolTipText     =   "Registrar os Componentes"
            ImageKey        =   "Register"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Unregister"
            Object.ToolTipText     =   "Desregistrar os Componentes"
            ImageKey        =   "UnRegister"
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "Delete"
            Object.ToolTipText     =   "Deletar Arquivos Selecionados"
            ImageKey        =   "Delete"
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "About"
            Object.ToolTipText     =   "Sobre o ActiveSearch"
            ImageKey        =   "HLP"
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   60
         Picture         =   "ActiveTools.frx":2B08
         ScaleHeight     =   315
         ScaleWidth      =   135
         TabIndex        =   4
         Top             =   0
         Width           =   135
      End
   End
   Begin MSComctlLib.StatusBar stbMain 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   4605
      Width           =   7425
      _ExtentX        =   13097
      _ExtentY        =   556
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   225
      Top             =   3780
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   128
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   24
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ActiveTools.frx":2BEA
            Key             =   "DLL"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ActiveTools.frx":2D02
            Key             =   "OCX"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ActiveTools.frx":2E1A
            Key             =   "EXE"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ActiveTools.frx":2F32
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ActiveTools.frx":3046
            Key             =   "FillFolder"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ActiveTools.frx":315A
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ActiveTools.frx":326E
            Key             =   "BMP"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ActiveTools.frx":3386
            Key             =   "TXT"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ActiveTools.frx":349E
            Key             =   "CPL"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ActiveTools.frx":35B2
            Key             =   "HLP"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ActiveTools.frx":36C6
            Key             =   "ClosedFolder"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ActiveTools.frx":37DA
            Key             =   "OpenedFolder"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ActiveTools.frx":38EE
            Key             =   "Explorer"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ActiveTools.frx":3A02
            Key             =   "Register"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ActiveTools.frx":3B1A
            Key             =   "UnRegister"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ActiveTools.frx":3C32
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ActiveTools.frx":3D4A
            Key             =   "Prop"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ActiveTools.frx":3E5E
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ActiveTools.frx":3F72
            Key             =   "WinFile"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ActiveTools.frx":4092
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ActiveTools.frx":41A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ActiveTools.frx":42BA
            Key             =   "BROWSER"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ActiveTools.frx":470E
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ActiveTools.frx":486A
            Key             =   "UnCheck"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwFiles 
      Height          =   4155
      Left            =   15
      TabIndex        =   0
      Top             =   420
      Width           =   7290
      _ExtentX        =   12859
      _ExtentY        =   7329
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Name"
         Text            =   "Nome"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Path"
         Text            =   "Na Pasta"
         Object.Width           =   5997
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Tipo"
         Object.Width           =   952
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "Version"
         Text            =   "Versão"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Key             =   "Size"
         Text            =   "Tamanho"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Key             =   "DateTime"
         Text            =   "Modificado"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Key             =   "GUID"
         Text            =   "GUID"
         Object.Width           =   6703
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Key             =   "Interface"
         Text            =   "Interface"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Key             =   "SortBySize"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Key             =   "SortByDate"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuBrowser 
         Caption         =   "&Object Browser"
      End
      Begin VB.Menu mnuProp 
         Caption         =   "&Propriedades"
      End
      Begin VB.Menu mnuExplore 
         Caption         =   "&Explorar Pasta do Arquivo"
      End
      Begin VB.Menu mnuLine 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&Sobre..."
      End
   End
End
Attribute VB_Name = "frmActiveTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sFiles As String
Private FT As Boolean
Private spath As String
Private oCheck As New Collection

Private Sub barMain_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim sKey As String
Dim sTemp As String
Dim lItem As ListItem
Dim oFileDlg As New clsFileDialog
Dim oTypeLib As New TLI.TypeLibInfo
On Error GoTo Erro
    sKey = Button.Key
    Select Case sKey
    Case "Find"
        Dim oCol As New Collection
        Dim Anim As clsAnimate
        Dim oItem As clsFileInfo
        Screen.MousePointer = vbHourglass
        Set oCheck = New Collection
        Set Anim = New clsAnimate
        With Anim
           .AutoPlay = True
           .Center = False
           .Parent = picAnim.hWnd
           .ResourceID = 10
        End With
        picAnim.Visible = True
        Anim.AniPlay
        lvwFiles.ListItems.Clear
        Set oCol = FilesSearch(spath, sFiles)
        DoEvents
        For Each oItem In oCol
            If InStr("OCX;DLL;EXE;BMP;TXT;BMP;HLP", oItem.Extension) > 0 Then
                Set lItem = lvwFiles.ListItems.Add(, oItem.Path & oItem.Name, oItem.Name, , oItem.Extension)
            ElseIf InStr("OLB", oItem.Extension) > 0 Then
                Set lItem = lvwFiles.ListItems.Add(, oItem.Path & oItem.Name, oItem.Name, , "DLL")
            Else
                Set lItem = lvwFiles.ListItems.Add(, oItem.Path & oItem.Name, oItem.Name, , "WinFile")
            End If
            lItem.SubItems(1) = oItem.Path
            lItem.SubItems(2) = oItem.Extension
            lItem.SubItems(3) = oItem.Version
            lItem.SubItems(4) = CStr(oItem.Size) & " bytes"
            lItem.SubItems(5) = CStr(oItem.DateTime)
            If InStr("OCX;DLL;TLB;OLB;EXE", oItem.Extension) Then
                On Error Resume Next
                Set oTypeLib = TLI.TypeLibInfoFromFile(oItem.Path & oItem.Name)
                If Err = 0 Then
                    lItem.SubItems(6) = oTypeLib.Guid
                    lItem.SubItems(7) = Format(oTypeLib.MajorVersion) & "." & Format(oTypeLib.MinorVersion)
                End If
                On Error GoTo Erro
            End If
            lItem.SubItems(8) = Format(oItem.Size, "0000000000000000")
            lItem.SubItems(9) = Format(CDbl(oItem.DateTime))
        Next
        Set oCol = Nothing
        Set oTypeLib = Nothing
        Anim.AniStop
        picAnim.Visible = False
        bCancel = False
        Screen.MousePointer = vbDefault
        'Habilita Ferramentas
        If lvwFiles.ListItems.Count > 0 Then
            Set lvwFiles.SelectedItem = lvwFiles.ListItems(1)
            barMain.Buttons("Save").Enabled = True
            barMain.Buttons("Print").Enabled = True
            barMain.Buttons("Prop").Enabled = True
            barMain.Buttons("ObjBrowser").Enabled = True
            barMain.Buttons("Explore").Enabled = True
            barMain.Buttons("CheckAll").Enabled = True
            barMain.Buttons("UnCheckAll").Enabled = True
        Else
            barMain.Buttons("Save").Enabled = False
            barMain.Buttons("Print").Enabled = False
            barMain.Buttons("Prop").Enabled = False
            barMain.Buttons("ObjBrowser").Enabled = False
            barMain.Buttons("Explore").Enabled = True
            barMain.Buttons("CheckAll").Enabled = False
            barMain.Buttons("UnCheckAll").Enabled = False
        End If
        barMain.Buttons("Register").Enabled = False
        barMain.Buttons("Unregister").Enabled = False
        barMain.Buttons("Delete").Enabled = False
        lvwAutofitColumnWidth lvwFiles
    Case "Folder"
        sTemp = BrowseForFolder(hWnd, LoadResString(209), spath)
        If sTemp <> "" Then spath = sTemp
    Case "Files"
        frmFilesFilter.Show vbModal, Me
    Case "Save"
        oFileDlg.DialogTitle = LoadResString(106)
        oFileDlg.Filter = LoadResString(210)
        oFileDlg.InitialDir = CurDir
        oFileDlg.Flags = cdlOverWritePrompt Or _
                         cdlPathMustExist Or _
                         cdlHideReadOnly Or _
                         cdlExplorer
        oFileDlg.hWndParent = Me.hWnd
        If Not oFileDlg.ShowSave Then
            Exit Sub
        End If
        Open oFileDlg.Filename For Output As #1
        For Each lItem In lvwFiles.ListItems
            Print #1, lItem.Text _
                    , lItem.ListSubItems(1).Text _
                    , lItem.ListSubItems(2).Text _
                    , lItem.ListSubItems(3).Text _
                    , lItem.ListSubItems(4).Text _
                    , lItem.ListSubItems(5).Text _
                    , lItem.ListSubItems(6).Text _
                    , lItem.ListSubItems(7).Text
        Next
        Close #1
    Case "Print"
        If MsgBox(LoadResString(202), vbQuestion + vbOKCancel) = vbCancel Then
          Exit Sub
        End If
        Printer.FontName = "Courier New"
        Printer.FontSize = 8
        Printer.Orientation = vbPRORPortrait
        For Each lItem In lvwFiles.ListItems
            Printer.Print lItem.Text _
                        , Tab(2), lItem.ListSubItems(1).Text _
                        , Tab(2), lItem.ListSubItems(3).Text _
                        , Tab(2), lItem.ListSubItems(4).Text _
                        , Tab(2), lItem.ListSubItems(5).Text _
                        , Tab(2), lItem.ListSubItems(6).Text _
                        , Tab(2), lItem.ListSubItems(7).Text
        Next
        Printer.EndDoc
    Case "Prop"
        If lvwFiles.SelectedItem Is Nothing Then
            Exit Sub
        End If
        Set lItem = lvwFiles.SelectedItem
        sTemp = lItem.ListSubItems(1).Text + lItem.Text
        ShowProperties sTemp, Me.hWnd
    Case "ObjBrowser"
        If lvwFiles.SelectedItem Is Nothing Then
            Exit Sub
        End If
        Set lItem = lvwFiles.SelectedItem
        frmObjBrowser.Tag = lItem.ListSubItems(1).Text + lItem.Text
        frmObjBrowser.Show
    Case "Explore"
        If lvwFiles.SelectedItem Is Nothing Then
            Exit Sub
        End If
        Set lItem = lvwFiles.SelectedItem
        sTemp = lItem.ListSubItems(1).Text
        Shell "Explorer " & sTemp, vbNormalFocus
    Case "CheckAll"
        Set oCheck = New Collection
        For Each lItem In lvwFiles.ListItems
            lItem.Checked = True
            oCheck.Add lItem, lItem.Key
        Next
        If oCheck.Count > 0 Then
            barMain.Buttons("Register").Enabled = True
            barMain.Buttons("Unregister").Enabled = True
            barMain.Buttons("Delete").Enabled = True
        End If
    Case "UnCheckAll"
        For Each lItem In lvwFiles.ListItems
            If lItem.Checked Then
                lItem.Checked = False
                oCheck.Remove lItem.Key
            End If
        Next
        barMain.Buttons("Register").Enabled = False
        barMain.Buttons("Unregister").Enabled = False
        barMain.Buttons("Delete").Enabled = False
    Case "Register"
        For Each lItem In oCheck
            Select Case lItem.SubItems(2)
            Case "OCX", "DLL", "OLB"
                Call ShellExecute(Me.hWnd, "open", "regsvr32.exe ", """" & lItem.Key & """", CurDir$, SW_SHOW)
            Case "EXE"
                Call ShellExecute(Me.hWnd, "open", lItem.Key, "/regserver", CurDir$, SW_SHOW)
            Case Else
                MsgBox lItem.Text & LoadResString(203), vbExclamation
            End Select
        Next
    Case "Unregister"
        For Each lItem In oCheck
            Select Case lItem.SubItems(2)
            Case "OCX", "DLL", "OLB"
                Call ShellExecute(Me.hWnd, "open", "regsvr32.exe", "/u """ & lItem.Key & """", CurDir$, SW_SHOW)
            Case "EXE"
                Call ShellExecute(Me.hWnd, "open", lItem.Key, "/unregserver", CurDir$, SW_SHOW)
            Case Else
                MsgBox lItem.Text & LoadResString(203), vbExclamation
            End Select
        Next
    Case "Delete"
        sTemp = ""
        For Each lItem In oCheck
            sTemp = sTemp & lItem.Key & Chr$(0)
            lItem.Ghosted = True
        Next
        If sTemp <> "" Then
            Dim oDel As Collection
            Call ShellRecycleFile(Me.hWnd, sTemp)
            For Each lItem In oCheck
                lItem.Checked = False
                If Not FileExists(lItem.Key) Then
                    'Deletado
                    lvwFiles.ListItems.Remove lItem.Key
                Else
                    lItem.Ghosted = False
                End If
            Next
            Set oCheck = Nothing
            barMain.Buttons("Register").Enabled = False
            barMain.Buttons("Unregister").Enabled = False
            barMain.Buttons("Delete").Enabled = False
        End If
    Case "About"
        frmAbout.Show vbModal
    End Select
Sair:
    stbMain.SimpleText = LoadResString(206) & spath & " (" & sFiles & ")"
    Exit Sub
Erro:
    Screen.MousePointer = vbDefault
    MsgBox LoadResString(204) & "(" & Err.Number & ") " & Err.Description, vbExclamation, LoadResString(205)
    Resume Sair
End Sub

Private Sub Form_Activate()
    If FT = True Then
        If Me.Tag <> "" Then
            spath = Me.Tag
            barMain_ButtonClick barMain.Buttons("Find")
        End If
        FT = False
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then bCancel = True
End Sub

Private Sub Form_Load()
    On Error Resume Next
    FT = True
    Move 300, 300, Screen.Width - 600, Screen.Height - 900
    spath = GetSetting(App.Title, "Options", "Path", Left$(App.Path, 3))
    sFiles = GetSetting(App.Title, "Options", "Files", "*.ocx;*.dll;*.tlb;*.olb;*.exe")
    Me.Caption = App.Title
    TranslateForm
    stbMain.SimpleText = LoadResString(206) & spath & " (" & sFiles & ")"
    mnuBrowser.Caption = LoadResString(151)
    mnuProp.Caption = LoadResString(152)
    mnuExplore.Caption = LoadResString(153)
    mnuAbout.Caption = LoadResString(154)
    'Configura Toolbar se >= HiColor
    If GetColorDepth() >= 16 Then
        ImageList1.ListImages.Add , "LUPA", imgLupa.Picture
        ImageList1.ListImages.Add , "FOLDER", imgFolder.Picture
        ImageList1.ListImages.Add , "FILES", imgFiles.Picture
        ImageList1.ListImages.Add , "DISK", imgDisk.Picture
        ImageList1.ListImages.Add , "PRINTER", imgPrinter.Picture
        ImageList1.ListImages.Add , "INFO", imgInfo.Picture
        ImageList1.ListImages.Add , "EXPLORE", imgExplore.Picture
        barMain.Buttons("Find").Image = "LUPA"
        barMain.Buttons("Folder").Image = "FOLDER"
        barMain.Buttons("Files").Image = "FILES"
        barMain.Buttons("Save").Image = "DISK"
        barMain.Buttons("Explore").Image = "EXPLORE"
        barMain.Buttons("Print").Image = "PRINTER"
        barMain.Buttons("About").Image = "INFO"
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    lvwFiles.Move 30, lvwFiles.Top, Width - 180, Height - 1200
    CenterOnParent picAnim, Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting App.Title, "Options", "Path", spath
    SaveSetting App.Title, "Options", "Files", sFiles
    Unload frmObjBrowser
    End
End Sub

Private Sub lvwFiles_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Static LastCol As Long
    If LastCol = 0 Then LastCol = 1
    If LastCol = ColumnHeader.Index Then
        If lvwFiles.SortOrder = lvwAscending Then
            lvwFiles.SortOrder = lvwDescending
        Else
            lvwFiles.SortOrder = lvwAscending
        End If
    Else
        If ColumnHeader.Key = "Size" Then
            lvwFiles.SortKey = 8
        ElseIf ColumnHeader.Key = "DateTime" Then
            lvwFiles.SortKey = 9
        Else
            lvwFiles.SortKey = ColumnHeader.Index - 1
        End If
        lvwFiles.SortOrder = lvwAscending
    End If
    LastCol = ColumnHeader.Index
    lvwFiles.Sorted = True
End Sub

Private Sub lvwFiles_DblClick()
Dim lItem As Object
    If lvwFiles.SelectedItem Is Nothing Then
        Exit Sub
    End If
    Set lItem = lvwFiles.SelectedItem
    frmObjBrowser.Tag = lItem.ListSubItems(1).Text + lItem.Text
    frmObjBrowser.Show
End Sub

Private Sub lvwFiles_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If Item.Checked Then
        oCheck.Add Item, Format(Item.Key)
        If oCheck.Count = 1 Then
            barMain.Buttons("Register").Enabled = True
            barMain.Buttons("Unregister").Enabled = True
            barMain.Buttons("Delete").Enabled = True
        End If
    Else
        oCheck.Remove Format(Item.Key)
        If oCheck.Count = 0 Then
            barMain.Buttons("Register").Enabled = False
            barMain.Buttons("Unregister").Enabled = False
            barMain.Buttons("Delete").Enabled = False
        End If
    End If
End Sub
Sub TranslateForm()
Dim lColumn As ColumnHeader
Dim lButton As Button
    For Each lButton In barMain.Buttons
        If lButton.Style = tbrDefault Then
            lButton.ToolTipText = LoadResString(100 + lButton.Index)
        End If
    Next
    For Each lColumn In lvwFiles.ColumnHeaders
        If lColumn.Index <= 8 Then
            lColumn.Text = LoadResString(300 + lColumn.Index)
        End If
    Next
End Sub

Private Sub lvwFiles_KeyPress(KeyAscii As Integer)
Dim lItem As Object
    If KeyAscii = vbKeyReturn Then
        If lvwFiles.SelectedItem Is Nothing Then
            Exit Sub
        End If
        Set lItem = lvwFiles.SelectedItem
        frmObjBrowser.Tag = lItem.ListSubItems(1).Text + lItem.Text
        frmObjBrowser.Show
    End If
End Sub

Private Sub lvwFiles_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If lvwFiles.SelectedItem Is Nothing Then
        Exit Sub
    ElseIf Button = vbRightButton Then
        PopupMenu mnuPopUp, , , , mnuBrowser
    End If
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show vbModal
End Sub

Private Sub mnuBrowser_Click()
    If lvwFiles.SelectedItem Is Nothing Then
        Exit Sub
    End If
    frmObjBrowser.Tag = lvwFiles.SelectedItem.ListSubItems(1).Text + _
                        lvwFiles.SelectedItem.Text
    frmObjBrowser.Show
End Sub

Private Sub mnuExplore_Click()
    Shell "Explorer " & lvwFiles.SelectedItem.ListSubItems(1).Text, vbNormalFocus
End Sub

Private Sub mnuProp_Click()
Dim sTemp As String
    If lvwFiles.SelectedItem Is Nothing Then
        Exit Sub
    End If
    sTemp = lvwFiles.SelectedItem.ListSubItems(1).Text + _
            lvwFiles.SelectedItem.Text
    ShowProperties sTemp, Me.hWnd
End Sub


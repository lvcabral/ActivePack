VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmObjBrowser 
   Caption         =   "Object Browser"
   ClientHeight    =   5655
   ClientLeft      =   1995
   ClientTop       =   2880
   ClientWidth     =   6960
   Icon            =   "ObjBrowser.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   377
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   464
   StartUpPosition =   2  'CenterScreen
   Tag             =   "c:\projetos\ActivePack\ActiveText\ActiveText.ocx"
   Begin VB.PictureBox picInfo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   6330
      MouseIcon       =   "ObjBrowser.frx":0442
      MousePointer    =   99  'Custom
      Picture         =   "ObjBrowser.frx":0594
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   4
      ToolTipText     =   "Sobre ActiveX Tools"
      Top             =   4695
      Width           =   240
   End
   Begin VB.PictureBox splitter 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      HasDC           =   0   'False
      Height          =   4455
      Left            =   2700
      MousePointer    =   9  'Size W E
      ScaleHeight     =   297
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   5
      TabIndex        =   3
      Top             =   60
      Width           =   75
   End
   Begin VB.TextBox txtDescPanel 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      ToolTipText     =   "Description"
      Top             =   4620
      Width           =   6810
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   180
      Top             =   3900
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ObjBrowser.frx":0B1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ObjBrowser.frx":0BC6
            Key             =   "Class"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ObjBrowser.frx":0C7E
            Key             =   "Enum"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ObjBrowser.frx":0D1E
            Key             =   "Event"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ObjBrowser.frx":0DAA
            Key             =   "Global"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ObjBrowser.frx":0E3A
            Key             =   "Unknown"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ObjBrowser.frx":0ED6
            Key             =   "Method"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ObjBrowser.frx":0F76
            Key             =   "Project"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ObjBrowser.frx":102A
            Key             =   "Property"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ObjBrowser.frx":10CA
            Key             =   "Type"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ObjBrowser.frx":116E
            Key             =   "Const"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      CausesValidation=   0   'False
      Height          =   4515
      Left            =   2760
      TabIndex        =   1
      Top             =   60
      Width           =   4110
      _ExtentX        =   7250
      _ExtentY        =   7964
      SortKey         =   1
      View            =   3
      LabelEdit       =   1
      SortOrder       =   -1  'True
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Type"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "DataType"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Parameters"
         Object.Width           =   26458
      EndProperty
   End
   Begin MSComctlLib.TreeView TreeView1 
      CausesValidation=   0   'False
      Height          =   4515
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   7964
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   0
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "ImageList1"
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
   End
End
Attribute VB_Name = "frmObjBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FT As Boolean
Public Prop As clsProps

Public CurInstance As String
Public CurProp As String

Public strReg As String

Private strProgID As String

'variable to hold the width of the splitter bar
 Private Const SPLT_WDTH As Integer = 4

'variable to hold the last-sized position
 Private currSplitPosX As Long

'variable to hold the horizontal
'and vertical offsets of the 2 controls
 Dim CTRL_OFFSET As Integer

'variable to hold the Splitter bar colour
 Dim SPLT_COLOUR As Long

Public Function GetType(sVarType As TliVarType) As String
    Dim strType As String
    
   
    Select Case sVarType
    Case VT_BLOB
        strType = strType & "Blob"
    Case VT_BLOB_OBJECT
        strType = strType & "Blob Object"
    Case VT_BOOL
        strType = strType & "Boolean"
    Case VT_BSTR
        strType = strType & "String"
    Case VT_CARRAY
        strType = strType & "cArray"
    Case VT_CF
        strType = strType & "CF"
    Case VT_CLSID
        strType = strType & "Class ID"
    Case VT_CY
        strType = strType & "Currency"
    Case VT_DATE
        strType = strType & "Date / Time"
    Case VT_DECIMAL
        strType = strType & "Decimal"
    Case VT_DISPATCH
        strType = strType & "Object"
    Case VT_EMPTY
        strType = strType & "Object"
    Case VT_ERROR
        strType = strType & "Error"
    Case VT_FILETIME
        strType = strType & "File Time"
    Case VT_HRESULT
        strType = strType & "hResult"
    Case VT_I1
        strType = strType & "Integer"
    Case VT_I2
        strType = strType & "Integer"
    Case VT_I4
        strType = strType & "Long"
    Case VT_I8
        strType = strType & "Long"
    Case VT_INT
        strType = strType & "Integer"
    Case VT_LPSTR
        strType = strType & "String"
    Case VT_LPWSTR
        strType = strType & "Wide String"
    Case VT_NULL
        strType = strType & "Null"
    Case VT_PTR
        strType = strType & "Pointer"
    Case VT_R4
        strType = strType & "Single"
    Case VT_R8
        strType = strType & "Double"
    Case VT_RECORD
        strType = strType & "Record"
    Case VT_RESERVED
        strType = strType & "Reserved"
    Case VT_SAFEARRAY
        strType = strType & "Safe Array"
    Case VT_STORAGE
        strType = strType & "Storage"
    Case VT_STORED_OBJECT
        strType = strType & "Stored Object"
    Case VT_STREAM
        strType = strType & "Stream"
    Case VT_STREAMED_OBJECT
        strType = strType & "Streamed Object"
    Case VT_UI1
        strType = strType & "Byte"
    Case VT_UI2
        strType = strType & "ui2"
    Case VT_UI4
        strType = strType & "OLE_Color"
    Case VT_UI8
        strType = strType & "ui8"
    Case VT_UINT
        strType = strType & "uInteger"
    Case VT_UNKNOWN
        strType = strType & "Unknown"
    Case VT_USERDEFINED
        strType = strType & "User Defined"
    Case VT_VARIANT
        strType = strType & "Variant"
    Case VT_VECTOR
        strType = strType & "Vector"
    Case VT_VOID
        strType = strType & "Empty"
    Case Else
        strType = "Unknown: (" & sVarType & ")"
    End Select
End1:

    GetType = strType
End Function

Public Function IsProp(InvokKind As InvokeKinds) As String
    Dim StrS As String
    Select Case InvokKind
    Case INVOKE_CONST
        StrS = "Const"
    Case INVOKE_EVENTFUNC
        StrS = "Event"
    Case INVOKE_FUNC
        StrS = "Method"
    Case INVOKE_PROPERTYGET
        StrS = "Property"
    Case INVOKE_PROPERTYPUT
        StrS = "Property"
    Case INVOKE_PROPERTYPUTREF
        StrS = "Property"
    Case INVOKE_UNKNOWN
        StrS = "Global"
    End Select
    IsProp = StrS
End Function

Public Sub LoadStuff()
    Dim Filename As String
    Dim objTLBInfo As TypeLibInfo
    Dim objTLBTemp As TypeLibInfo
    Dim objCoCls As CoClasses
    
    Dim objInterface As InterfaceInfo
    Dim objTypeInfo As TypeInfo
    Dim objMember As MemberInfo
    Dim objInstances As clsInstances
    Dim colInts As Collection
    Dim strName As String
    Dim strCLSID As String
    Dim I As Long
    Dim j As Long
    Dim k As Long
    Dim n As Long
    Dim argv As String
    Dim vtype As Long
    Dim sbyref As String
    Dim sarray As String
    Dim stype As String
    Dim sretype As String
    Dim slixo As String
    
    If Me.Tag = "" Then Exit Sub
    Filename = Me.Tag
 
    On Error Resume Next
    Err.Clear
    Set objTLBInfo = TypeLibInfoFromFile(Filename)
    If Err Then
        Select Case Err.Number
        Case -2147220990
            Screen.MousePointer = vbDefault
            MsgBox LoadResString(211), vbCritical
        Case Else
            Screen.MousePointer = vbDefault
            MsgBox LoadResString(204) & Err.Number & vbCrLf & Err.Description, vbCritical
        End Select
        Set objTLBInfo = Nothing
        Me.Hide
        Unload Me
        DoEvents
        Exit Sub
    End If
    'Verify if its registered
    Set objTLBTemp = TypeLibInfoFromRegistry(objTLBInfo.Guid, objTLBInfo.MajorVersion, objTLBInfo.MinorVersion, objTLBInfo.LCID)
    strReg = objTLBTemp.ContainingFile
    Set objTLBTemp = Nothing
    
    Set objCoCls = objTLBInfo.CoClasses
    Set colInts = New Collection
    On Error GoTo 0
 
    ' Processa as CoClasses
    
    For I = 1 To objCoCls.Count
        strName = objCoCls.Item(I).Name
        strCLSID = objCoCls.Item(I).Guid
        Set objInstances = Prop.Add(strName, objCoCls.Item(I).TypeKindString, objCoCls.Item(I).HelpString, "", "", "", "", "", strCLSID)
        For k = 1 To objCoCls.Item(I).Interfaces.Count
            On Error Resume Next
            Err.Clear
            Set objInterface = objCoCls.Item(I).Interfaces.Item(k)
            If Err <> 0 Then
                MsgBox LoadResString(204) & Err.Number & vbCrLf & Err.Description, vbCritical
            Else
                colInts.Add objInterface.Guid, objInterface.Name
            End If
            For j = 1 To objInterface.Members.Count
                Set objMember = objInterface.Members.Item(j)
                If LCase$(objMember.Name) <> "queryinterface" And _
                   LCase$(objMember.Name) <> "gettypeinfo" And _
                   LCase$(objMember.Name) <> "gettypeinfocount" And _
                   LCase$(objMember.Name) <> "getidsofnames" And _
                   LCase$(objMember.Name) <> "invoke" And _
                   LCase$(objMember.Name) <> "addref" And _
                   LCase$(objMember.Name) <> "release" Then
                    argv = ""
                    If objMember.Parameters.Count > 0 Then
                        For n = 1 To objMember.Parameters.Count
                            vtype = objMember.Parameters.Item(n).VarTypeInfo.VarType
                            'Verifica se é ByRef
                            If vtype And VT_BYREF Then
                                sbyref = "ByRef "
                                vtype = (vtype Xor VT_BYREF)
                            Else
                                sbyref = ""
                            End If
                            
                            'Verifica se é Array
                            If vtype And VT_ARRAY Then
                                sarray = "()"
                                vtype = (vtype Xor VT_ARRAY)
                            Else
                                sarray = ""
                            End If
                            argv = argv & sbyref & objMember.Parameters.Item(n).Name & sarray & " As " & GetType(vtype) & ", "
                        Next
                        argv = Left$(argv, Len(argv) - 2)
                    End If
                    stype = ""
                    If Not objCoCls.Item(I).DefaultEventInterface Is Nothing Then
                        If objCoCls.Item(I).DefaultEventInterface.Name = objInterface.Name And _
                           objMember.InvokeKind = INVOKE_FUNC Then
                            stype = "Event"
                        End If
                    End If
                    If stype = "" Then stype = IsProp(objMember.InvokeKind)
                    If stype <> "Property" Or objMember.ReturnType.VarType <> VT_VOID Then
                        vtype = objMember.ReturnType.VarType
                        'Verifica se é Array
                        If vtype And VT_ARRAY Then
                            sarray = "()"
                            vtype = (vtype Xor VT_ARRAY)
                        Else
                            sarray = ""
                        End If
                        Call objInstances.Members.Add(objMember.Name, stype, objMember.HelpString, _
                                                      GetType(vtype) & sarray, argv, "", "", "", "")
                    End If
                End If
            Next
        Next
    Next
    'Processa TypeInfos do type DISPATCH, sem Underscore e que não tenham
    'sido processados como interfaces das CoClasses
    For k = 1 To objTLBInfo.TypeInfoCount
        On Error Resume Next
        Err.Clear
        Set objTypeInfo = objTLBInfo.TypeInfos.Item(k)
        If Err Then
            MsgBox LoadResString(204) & Err.Number & vbCrLf & Err.Description, vbCritical
        End If
        strName = objTypeInfo.Name
        strCLSID = objTypeInfo.Guid
        slixo = colInts(strName)
        If Left(strName, 1) <> "_" And objTypeInfo.TypeKind = TKIND_DISPATCH And Err <> 0 Then
            Set objInstances = Prop.Add(strName, objTypeInfo.TypeKindString, objTypeInfo.HelpString, "", "", "", "", "", strCLSID)
            For j = 1 To objTypeInfo.Members.Count
                Set objMember = objTypeInfo.Members.Item(j)
                If LCase$(objMember.Name) <> "queryinterface" And _
                   LCase$(objMember.Name) <> "gettypeinfo" And _
                   LCase$(objMember.Name) <> "gettypeinfocount" And _
                   LCase$(objMember.Name) <> "getidsofnames" And _
                   LCase$(objMember.Name) <> "invoke" And _
                   LCase$(objMember.Name) <> "addref" And _
                   LCase$(objMember.Name) <> "release" Then
                    argv = ""
                    If objMember.Parameters.Count > 0 Then
                        For n = 1 To objMember.Parameters.Count
                            vtype = objMember.Parameters.Item(n).VarTypeInfo.VarType
                            'Verifica se é ByRef
                            If vtype And VT_BYREF Then
                                sbyref = "ByRef "
                                vtype = (vtype Xor VT_BYREF)
                            Else
                                sbyref = ""
                            End If

                            'Verifica se é Array
                            If vtype And VT_ARRAY Then
                                sarray = "()"
                                vtype = (vtype Xor VT_ARRAY)
                            Else
                                sarray = ""
                            End If
                            argv = argv & sbyref & objMember.Parameters.Item(n).Name & sarray & " As " & GetType(vtype) & ", "
                        Next
                        argv = Left$(argv, Len(argv) - 2)
                    End If
                    stype = IsProp(objMember.InvokeKind)
                    If stype <> "Property" Or objMember.ReturnType.VarType <> VT_VOID Then
                        Call objInstances.Members.Add(objMember.Name, stype, objMember.HelpString, _
                                              GetType(objMember.ReturnType.VarType), argv, "", "", "", "")
                    End If
                End If
            Next
        End If
    Next
    'Processa as Constantes
    For I = 1 To objTLBInfo.Constants.Count
        strName = objTLBInfo.Constants.Item(I).Name
        Set objInstances = Prop.Add(strName, objTLBInfo.Constants.Item(I).TypeKindString, objTLBInfo.Constants.Item(I).HelpString, "", "", "", "", "", "")
        For k = 1 To objTLBInfo.Constants.Item(I).Members.Count
            Set objMember = objTLBInfo.Constants.Item(I).Members.Item(k)

            Call objInstances.Members.Add(objMember.Name, "Const", objMember.HelpString & vbCrLf & objMember.Name & "=" & objMember.Value, GetType(objMember.ReturnType.VarType), "", "", CStr(objMember.Value), "", "")
        Next
    Next
    strProgID = objTLBInfo.Name
    LoadStuff2 objTLBInfo.HelpString
 
    Set objTLBInfo = Nothing
    Set objCoCls = Nothing
    Set objTypeInfo = Nothing
    Set objInterface = Nothing
    Set objMember = Nothing
    Set objInstances = Nothing
 
End Sub


Public Sub LoadStuff2(StrHelp As String)
    Dim A As Long
    
    Dim NodX As Node
    On Error Resume Next
    Set NodX = TreeView1.Nodes.Add(, , "Root", StrHelp, "Project", "Project")
    NodX.Expanded = True
    
    For A = 1 To Prop.Count
        If Prop.Item(A).oType = "coclass" Then Prop.Item(A).oType = "Class"
        If Prop.Item(A).oType = "dispinterface" Then Prop.Item(A).oType = "Class"
        If Prop.Item(A).oType = "enum" Then Prop.Item(A).oType = "Enum"
        If Prop.Item(A).nName <> "" Then
        
            Set NodX = TreeView1.Nodes.Add("Root", tvwChild, Prop.Item(A).nName, Prop.Item(A).nName, Prop.Item(A).oType, Prop.Item(A).oType)
        End If
    Next
    Set NodX = Nothing
    
    On Error GoTo 0
End Sub

Private Sub Form_Activate()
On Error GoTo ExitNow
    If FT = True Then
        Me.Caption = Caption & " - " & Me.Tag
        strReg = ""
        DoEvents
        FT = False
        Screen.MousePointer = vbHourglass
        If Me.Tag <> "" Then LoadStuff
        TreeView1.FullRowSelect = True
        Set TreeView1.SelectedItem = TreeView1.Nodes("Root")
        TreeView1_NodeClick TreeView1.Nodes("Root")
        lvwAutofitColumnWidth ListView1
        Screen.MousePointer = vbDefault
    End If
ExitNow:
    Exit Sub
End Sub

Private Sub Form_Load()
 
    'set the startup variables
    CTRL_OFFSET = 4
    SPLT_COLOUR = &H808080
    
    'set the current splitter bar position to an
    'arbitrary value that will always be outside
    'the possible range.  This allows us to check
    'for movement of the splitter bar in subsequent
    'mousexxx subs.
    currSplitPosX = &H7FFFFFFF
    
    Move 900, 900, Screen.Width - 1200, Screen.Height - 1500
    
    Set Prop = New clsProps
    
    picInfo.ToolTipText = LoadResString(120)
    
    FT = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
    Set Prop = Nothing
End Sub

Private Sub Form_Resize()
  Dim x1 As Integer
  Dim x2 As Integer
  Dim height1 As Integer
  Dim width1 As Integer
  Dim width2 As Integer
    
  On Error Resume Next

 'set the height of the controls
  height1 = ScaleHeight - (CTRL_OFFSET * 3)
  
  x1 = CTRL_OFFSET
  width1 = TreeView1.Width
  
  x2 = x1 + TreeView1.Width + SPLT_WDTH - 1
  width2 = ScaleWidth - x2 - CTRL_OFFSET
  
 'move the left list
  TreeView1.Move x1 - 1, CTRL_OFFSET, width1, height1 - (txtDescPanel.Height + CTRL_OFFSET)
 
 'move the right list
  ListView1.Move x2, CTRL_OFFSET, width2 + 1, height1 - (txtDescPanel.Height + CTRL_OFFSET)
 
 'move the down panel
  txtDescPanel.Move x1 - 1, (CTRL_OFFSET * 2) + ListView1.Height, width1 + width2 + CTRL_OFFSET, txtDescPanel.Height
  picInfo.Top = txtDescPanel.Top + 4
  picInfo.Left = txtDescPanel.Left + txtDescPanel.Width - picInfo.Width - 20
 
 'reposition the splitter bar
  splitter.Move x1 + TreeView1.Width - 1, _
               CTRL_OFFSET, SPLT_WDTH, height1 - (txtDescPanel.Height + CTRL_OFFSET)
End Sub

Private Sub Form_Terminate()
    Set Prop = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Set Prop = Nothing
    Screen.MousePointer = vbDefault
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ListView1.Sorted = False
    
    If ListView1.SortKey = ColumnHeader.Index - 1 Then
        If ListView1.SortOrder = lvwAscending Then ListView1.SortOrder = lvwDescending Else ListView1.SortOrder = lvwAscending
        
    Else
        ListView1.SortKey = ColumnHeader.Index - 1
        
        ListView1.SortOrder = lvwAscending
    End If
    
    ListView1.Sorted = True
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    CurProp = Item.Text
    If Trim(Prop.Item(CurInstance).Members.Item(CurProp).Description) <> "" Then
        txtDescPanel.Text = Replace(Prop.Item(CurInstance).Members.Item(CurProp).Description, "\par", vbCrLf) & vbCrLf & _
                     Prop.Item(CurInstance).Members.Item(CurProp).Arguments
    Else
        txtDescPanel.Text = Prop.Item(CurInstance).Members.Item(CurProp).Arguments
    End If
End Sub


Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim ItmX As ListItem
    
 
    If Button = vbRightButton Then
    
        Set ItmX = ListView1.HitTest(x, y)
        If ItmX Is Nothing Then Exit Sub
        
'        PopupMenu frmMain.Editmnu
    End If
    
End Sub

Private Sub picInfo_Click()
    frmAbout.Show vbModal
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
    If Node.Key = "Root" Then
        Dim x As TypeLibInfo
        ListView1.ListItems.Clear
        CurInstance = ""
        CurProp = ""
        Set x = TypeLibInfoFromFile(Me.Tag)
        txtDescPanel.Text = "File Path : " & Me.Tag & vbCrLf & _
                     "File Info : " & GetFileSize(Me.Tag) & " bytes " & _
                     " " & GetFileTime(Me.Tag, ftt_Modified, True) & _
                     " v" & GetFileVersion(Me.Tag) & vbCrLf & _
                     "TypeLib   : " & x.Guid & _
                     " v" & x.MajorVersion & "." & x.MinorVersion & _
                     IIf(strReg <> "", IIf(Trim(strReg) <> Me.Tag, vbCrLf & "Registered as " & strReg, " ** Registered **"), " ** Not Registered **")
        Set x = Nothing
        Exit Sub
    End If
    Dim f As clsInstances
    Dim A As Long
    Dim ItmX As ListItem
  
    CurInstance = Node.Text
    CurProp = ""
    If Len(Prop.Item(CurInstance).Description) > 0 Then
        txtDescPanel.Text = Prop.Item(CurInstance).Description & vbCrLf
    Else
        txtDescPanel.Text = ""
    End If
    If Len(Prop.Item(CurInstance).CLSID) > 0 Then
        txtDescPanel.Text = txtDescPanel.Text & _
                            "CLSID : " & Prop.Item(CurInstance).CLSID & vbCrLf
    End If
    txtDescPanel.Text = txtDescPanel.Text & _
                        "ProgID: " & strProgID & "." & Prop.Item(CurInstance).Key
    Set f = Prop.Item(Node.Text)
    ListView1.ListItems.Clear
    For A = 1 To f.Members.Count
        If f.Members.Item(A).oType = "" Then f.Members.Item(A).oType = "Unknown"
        Set ItmX = ListView1.ListItems.Add(, CurInstance + f.Members.Item(A).nName + "Ctrl", f.Members.Item(A).nName, f.Members.Item(A).oType, f.Members.Item(A).oType)
        ItmX.SubItems(1) = f.Members.Item(A).oType
        ItmX.SubItems(2) = f.Members.Item(A).dType
        ItmX.SubItems(3) = f.Members.Item(A).Arguments
        If Left(ItmX.Text, 1) = "_" Then
           ItmX.Ghosted = True
           ItmX.ForeColor = &H808080
        End If
    Next
    Set ItmX = Nothing
    Set f = Nothing
    lvwAutofitColumnWidth ListView1
    ListView1.Sorted = False
    ListView1.SortKey = 0
    
    ListView1.SortOrder = lvwAscending
    ListView1.Sorted = True
    ListView1.Sorted = False
    ListView1.SortKey = 1
    ListView1.SortOrder = lvwDescending
    ListView1.Sorted = True
 
End Sub

Private Sub splitter_MouseMove(Button As Integer, _
                               Shift As Integer, _
                               x As Single, y As Single)
    
 'if the splitter has been moved...
  If currSplitPosX <> &H7FFFFFFF Then
  
   'if the current position <> default, reposition
   'the splitter and set this as the current value
    
    If CLng(x) <> currSplitPosX Then
        splitter.Move splitter.Left + x, _
                      CTRL_OFFSET, SPLT_WDTH, _
                      ScaleHeight - (CTRL_OFFSET * 2)
        currSplitPosX = CLng(x)
    End If
    
End If

End Sub


Private Sub splitter_MouseUp(Button As Integer, _
                             Shift As Integer, _
                             x As Single, y As Single)
    
 'if the splitter has been moved...
  If currSplitPosX <> &H7FFFFFFF Then
      
     'if the current position <> the last
     'position do a final move of the splitter
      If CLng(x) <> currSplitPosX Then
        splitter.Move splitter.Left + x, _
                     CTRL_OFFSET, SPLT_WDTH, _
                     ScaleHeight - (CTRL_OFFSET * 2)
      End If
      
     'call this the default position
      currSplitPosX = &H7FFFFFFF
      
     'restore the normal splitter colour
      splitter.BackColor = &H8000000F
     
     'and check for valid sizing.
     'Either enforce the default minimum & maximum widths for
     'the left list, or, if within range, set the width
     
      If splitter.Left > 60 And splitter.Left < (ScaleWidth - 60) Then
            'the pane is within range
             TreeView1.Width = splitter.Left - TreeView1.Left
      
      ElseIf splitter.Left < 60 Then
            'the pane is too small
             TreeView1.Width = 60
      
      Else: 'the pane is too wide
             TreeView1.Width = ScaleWidth - 60
      End If
      
     'reposition both lists, and the splitter bar
      Form_Resize
  
  End If
  
End Sub

Private Sub splitter_MouseDown(Button As Integer, _
                               Shift As Integer, _
                               x As Single, y As Single)
    
    If Button = vbLeftButton Then
    
       'change the splitter colour
        splitter.BackColor = SPLT_COLOUR
       
       'set the current position to x
        currSplitPosX = CLng(x)
    
    Else
    
       'not the left button, so...
       'if the current position <> default, cause a MouseUp
        If currSplitPosX <> &H7FFFFFFF Then
           splitter_MouseUp Button, Shift, x, y
        End If
       
       'set the current position to the default value
        currSplitPosX = &H7FFFFFFF
    
    End If
End Sub


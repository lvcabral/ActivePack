Attribute VB_Name = "modActiveTools"
Option Explicit
Public Const SW_SHOW = 5
Public Declare Function ShellExecute& Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long)
Public bCancel As Boolean
'Then declare these constants. LVM_FIRST is base code
Private Const LVM_FIRST As Long = &H1000

'LVM_SETCOLUMNWIDTH constant tells Windows to handle column width
Private Const LVM_SETCOLUMNWIDTH As Long = LVM_FIRST + 30

'LVSCW_AUTOSIZE_USEHEADER is used to tell Windows that the column heads are to be resized
Private Const LVSCW_AUTOSIZE_USEHEADER As Long = -2

'Then decleare SendMessage for sending Windows message in order to autofit Listview Control
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'To get the color depth
Public Declare Function GetDeviceCaps Lib "gdi32" _
   (ByVal hdc As Long, _
    ByVal nIndex As Long) As Long
Const BITSPIXEL As Long = 12
Sub Main()
    If Command = "" Then
        frmActiveTools.Show
    Else
        If FileExists(Command) Then
            frmObjBrowser.Tag = Command
            frmObjBrowser.Show
        ElseIf FolderExists(Command) Then
            frmActiveTools.Tag = Command
            frmActiveTools.Show
        End If
    End If
End Sub

Function FilesSearch(ByVal DrivePath As String, Ext As String) As Collection
Static iLevel As Long
Static oFiles As New Collection
Dim oFile As New clsFileInfo
Dim XDir() As String
Dim XExt() As String
Dim TmpDir As String
Dim FFound As String
Dim DirCount As Integer
Dim x As Integer
    'Initialises Variables
    If iLevel = 0 Then
        Set oFiles = Nothing
    End If
    iLevel = iLevel + 1
    DirCount = 0
    ReDim XDir(0) As String
    XDir(DirCount) = ""

    If Right(DrivePath, 1) <> "\" Then
       DrivePath = DrivePath & "\"
    End If

    'Enter here the code for showing the path being
    'search. Example: Form1.label2 = DrivePath
    frmActiveTools.stbMain.SimpleText = LoadResString(207) & " - " & DrivePath
    'Search for all directories and store in the
    'XDir() variable

    DoEvents
    If bCancel Then GoTo Off
   
    TmpDir = Dir(DrivePath, vbDirectory)

    Do While TmpDir <> ""
        If TmpDir <> "." And TmpDir <> ".." Then
            On Error Resume Next
            If (GetAttr(DrivePath & TmpDir) And vbDirectory) = vbDirectory Then
                If Err = 0 Then
                XDir(DirCount) = DrivePath & TmpDir & "\"
                DirCount = DirCount + 1
                ReDim Preserve XDir(DirCount) As String
                End If
            End If
            On Error GoTo 0
        End If
        TmpDir = Dir
    Loop
    'Searches for the files given by extension Ext
    XExt = Split(Ext, ";")
    For x = 0 To UBound(XExt)
        FFound = Dir(DrivePath & XExt(x))
        Do Until FFound = ""
            'Code in here for the actions of the files found.
            'Files found stored in the variable FFound.
            'Example: Form1.list1.AddItem DrivePath & FFound
            Set oFile = New clsFileInfo
            oFile.Name = FFound
            oFile.Path = DrivePath
            oFile.Version = GetFileVersion(oFile.Path & oFile.Name)
            oFile.Size = GetFileSize(oFile.Path & oFile.Name)
            oFile.DateTime = GetFileTime(oFile.Path & oFile.Name, ftt_Modified, True)
            oFile.Extension = UCase(Right(Trim(oFile.Name), 3))
            oFiles.Add oFile
            FFound = Dir
            DoEvents
            If bCancel Then GoTo Off
        Loop
    Next
    'Recursive searches through all sub directories
    For x = 0 To (UBound(XDir) - 1)
        FilesSearch XDir(x), Ext
    Next
Off: If bCancel Then
        iLevel = 0
        Set oFiles = New Collection
        Exit Function
    End If
    Set FilesSearch = oFiles
    iLevel = iLevel - 1
End Function

Public Sub CenterOnParent(mdiChild As Object, mdiParent As Object)
    mdiChild.Left = (mdiParent.ScaleWidth - mdiChild.Width) / 2
    mdiChild.Top = (mdiParent.ScaleHeight - mdiChild.Height) / 2
End Sub


Function FileExists(ByVal Arquivo As String) As Boolean
    Dim nArq%
    nArq% = FreeFile
    On Error Resume Next
    Open Arquivo$ For Input As #nArq%
    If Err = 0 Then
       FileExists = True
    Else
       FileExists = False
    End If
    Close #nArq%
End Function

Public Function FolderExists(Pasta As String) As Boolean
    On Error Resume Next
    FolderExists = Len(Dir$(Pasta & "\.", vbDirectory)) > 0
End Function

'Then all we have to do is to write a public sub like this...
Public Sub lvwAutofitColumnWidth(ByVal lvw As MSComctlLib.ListView)
 
    'Declare a counter variable iCounter
    Dim iCounter As Long
 
    'Turn on error trapping
    On Error Resume Next
 
 
    'Check the current view for the ListView Control. If not in report view then exit here
    If lvw.View <> lvwReport Then Exit Sub '*** FAILSAFE REM     If lvw.View <> lvwReport Then Exit Sub
 
    'Now turn off the autoredrawing for the Listview control
    lvw.Visible = False
 
 
    'OK, so far so good. Now iterate through all columns and set column width to the widest
    For iCounter = 1 To lvw.ColumnHeaders.Count
        If lvw.ColumnHeaders.Item(iCounter).Width > 10 Then
            Call SendMessage(lvw.hwnd, LVM_SETCOLUMNWIDTH, iCounter - 1, ByVal LVSCW_AUTOSIZE_USEHEADER)
        End If
    Next
    
 
    'Now turn on the autoredrawing for the Listview control
    lvw.Visible = True
 
End Sub

Public Function GetColorDepth() As Long
    ' 8 - 256 cores
    '16 - HiColor - 64k cores
    '24 - TrueColor - 16m cores
   GetColorDepth = GetDeviceCaps(frmActiveTools.hdc, BITSPIXEL)

End Function

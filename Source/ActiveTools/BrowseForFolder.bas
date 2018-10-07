Attribute VB_Name = "basBrowseForFolder"
Option Explicit

'common to both methods
Public Type BROWSEINFO
  hOwner As Long
  pidlRoot As Long
  pszDisplayName As String
  lpszTitle As String
  ulFlags As Long
  lpfn As Long
  lParam As Long
  iImage As Long
End Type

Public Declare Function SHBrowseForFolder Lib _
   "shell32.dll" Alias "SHBrowseForFolderA" _
   (lpBrowseInfo As BROWSEINFO) As Long

Public Declare Function SHGetPathFromIDList Lib _
   "shell32.dll" Alias "SHGetPathFromIDListA" _
   (ByVal pidl As Long, _
   ByVal pszPath As String) As Long

Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)

Public Declare Function SendMessage Lib "user32" _
   Alias "SendMessageA" _
   (ByVal hWnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   lParam As Any) As Long
   
Public Declare Sub MoveMemory Lib "kernel32" _
   Alias "RtlMoveMemory" _
   (pDest As Any, _
    pSource As Any, _
    ByVal dwLength As Long)
    
Public Const MAX_PATH = 260
Public Const WM_USER = &H400
Public Const BFFM_INITIALIZED = 1

'Constants ending in 'A' are for Win95 ANSI
'calls; those ending in 'W' are the wide Unicode
'calls for NT.

'Sets the status text to the null-terminated
'string specified by the lParam parameter.
'wParam is ignored and should be set to 0.
Public Const BFFM_SETSTATUSTEXTA As Long = (WM_USER + 100)
Public Const BFFM_SETSTATUSTEXTW As Long = (WM_USER + 104)

'If the lParam  parameter is non-zero, enables the
'OK button, or disables it if lParam is zero.
'(docs erroneously said wParam!)
'wParam is ignored and should be set to 0.
Public Const BFFM_ENABLEOK As Long = (WM_USER + 101)

'Selects the specified folder. If the wParam
'parameter is FALSE, the lParam parameter is the
'PIDL of the folder to select , or it is the path
'of the folder if wParam is the C value TRUE (or 1).
'Note that after this message is sent, the browse
'dialog receives a subsequent BFFM_SELECTIONCHANGED
'message.
Public Const BFFM_SETSELECTIONA As Long = (WM_USER + 102)
Public Const BFFM_SETSELECTIONW As Long = (WM_USER + 103)
   
Private Const BIF_RETURNONLYFSDIRS = &H1

'specific to the PIDL method
'Undocumented call for the example. IShellFolder's
'ParseDisplayName member function should be used instead.
Public Declare Function SHSimpleIDListFromPath Lib _
   "shell32" Alias "#162" _
   (ByVal szPath As String) As Long


'windows-defined type OSVERSIONINFO
Public Type OSVERSIONINFO
  OSVSize         As Long
  dwVerMajor      As Long
  dwVerMinor      As Long
  dwBuildNumber   As Long
  PlatformID      As Long
  szCSDVersion    As String * 128
End Type
Public Const VER_PLATFORM_WIN32_NT = 2
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
  (lpVersionInformation As OSVERSIONINFO) As Long

        
Public Function BrowseCallbackProc(ByVal hWnd As Long, _
                                   ByVal uMsg As Long, _
                                   ByVal lParam As Long, _
                                   ByVal lpData As Long) As Long
 
  'Callback for the Browse PIDL method.
 
  'On initialization, set the dialog's
  'pre-selected folder using the pidl
  'set as the bi.lParam, and passed back
  'to the callback as lpData param.
 
   Select Case uMsg
      Case BFFM_INITIALIZED
      
         Call SendMessage(hWnd, BFFM_SETSELECTIONA, _
                          False, ByVal lpData)
                          
         Case Else:
         
   End Select

End Function


Public Function FARPROC(pfn As Long) As Long
  
  'A dummy procedure that receives and returns
  'the value of the AddressOf operator.
 
  'Obtain and set the address of the callback
  'This workaround is needed as you can't assign
  'AddressOf directly to a member of a user-
  'defined type, but you can assign it to another
  'long and use that (as returned here)
 
  FARPROC = pfn

End Function

Public Function ShellBrowseForFolder(lhWnd As Long, strMessage As String, Optional sSelPath As String) As String

   Dim BI As BROWSEINFO
   Dim pidl As Long
   Dim spath As String * MAX_PATH
  
   sSelPath = UnqualifyPath(sSelPath)
  
   With BI
      .hOwner = lhWnd
      .pidlRoot = 0
      .lpszTitle = strMessage
      .lpfn = FARPROC(AddressOf BrowseCallbackProc)
      .lParam = GetPIDLFromPath(sSelPath) 'replaces '= SHSimpleIDListFromPath(sSelPath)'
      .ulFlags = BIF_RETURNONLYFSDIRS
   End With
  
   pidl = SHBrowseForFolder(BI)
  
   If pidl Then
      If SHGetPathFromIDList(pidl, spath) Then
         ShellBrowseForFolder = Left$(spath, InStr(spath, vbNullChar) - 1)
      End If
     
     'free the pidl returned by call to SHBrowseForFolder
      Call CoTaskMemFree(pidl)
  End If
  
 'free the pidl set in call to GetPIDLFromPath
  Call CoTaskMemFree(BI.lParam)
  
End Function


Public Function GetPIDLFromPath(spath As String) As Long

  'return the pidl to the path supplied by calling the
  'undocumented API #162 (our name SHSimpleIDListFromPath).
  'This function is necessary as, unlike documented APIs,
  'the API is not implemented in 'A' or 'W' versions.

  If IsNT Then
    GetPIDLFromPath = SHSimpleIDListFromPath(StrConv(spath, vbUnicode))
  Else
    GetPIDLFromPath = SHSimpleIDListFromPath(spath)
  End If

End Function


Public Function IsNT() As Boolean

   #If Win32 Then
  
      Dim OSV As OSVERSIONINFO
   
      OSV.OSVSize = Len(OSV)
   
     'API returns 1 if a successful call
      If GetVersionEx(OSV) = 1 Then
   
        'PlatformId contains a value representing
        'the OS, so if its VER_PLATFORM_WIN32_NT,
        'return true
         IsNT = OSV.PlatformID = VER_PLATFORM_WIN32_NT
      End If

   #End If

End Function


Public Function UnqualifyPath(spath As String) As String

  'qualifying a path usually involves assuring
  'that its format is valid, including a trailing slash
  'ready for a filename. Since the SHBrowseForFolder API
  'will pre-select the path if it contains the trailing
  'slash, I call stripping it 'unqualifying the path'.
   If Len(spath) > 0 Then
   
      If Right$(spath, 1) = "\" Then
      
         UnqualifyPath = Left$(spath, Len(spath) - 1)
         Exit Function
      
      End If
   
   End If
   
   UnqualifyPath = spath
   
End Function


Attribute VB_Name = "modAPIFilesAndPaths"

' Module      : modAPIFilesAndPaths
' Description : Code for working with files and paths using the Windows API
' Source      : Total VB SourceBook 5
'
Public Enum EnumFileTimeType
    ftt_Created = 0
    ftt_Accessed = 1
    ftt_Modified = 2
End Enum

Private Type FILETIME
    lngLowDateTime As Long
    lngHighDateTime As Long
End Type

Private Type SYSTEMTIME
    intYear As Integer
    intMonth As Integer
    intDayOfWeek As Integer
    intDay As Integer
    intHour As Integer
    intMinute As Integer
    intSecond As Integer
    intMilliseconds As Integer
End Type

Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const OPEN_EXISTING = 3
Private Const FILE_FLAG_RANDOM_ACCESS = &H10000000

Private Declare Function CreateFile _
    Lib "kernel32" _
    Alias "CreateFileA" _
    (ByVal lpFileName As String, _
        ByVal dwDesiredAccess As Long, _
        ByVal dwShareMode As Long, _
        lpSecurityAttributes As Any, _
        ByVal dwCreationDisposition As Long, _
        ByVal dwFlagsAndAttributes As Long, _
        ByVal hTemplateFile As Long) _
    As Long

Private Declare Function CloseHandle _
    Lib "kernel32" _
    (ByVal hObject As Long) _
    As Long

Private Declare Function CGetFileSize _
    Lib "kernel32" _
    Alias "GetFileSize" _
    (ByVal hFile As Long, _
        lpFileSizeHigh As Long) _
    As Long
        
Private Declare Function CGetFileTime _
    Lib "kernel32" _
    Alias "GetFileTime" _
    (ByVal hFile As Long, _
        lpCreationTime As FILETIME, _
        lpLastAccessTime As FILETIME, _
        lpLastWriteTime As FILETIME) _
    As Long

Private Declare Function CSetFileTime _
    Lib "kernel32" _
    Alias "SetFileTime" _
    (ByVal hFile As Long, _
        lpCreationTime As FILETIME, _
        lpLastAccessTime As FILETIME, _
        lpLastWriteTime As FILETIME) _
    As Long

Private Declare Function CompareFileTime _
    Lib "kernel32" _
    (lpFileTime1 As FILETIME, _
        lpFileTime2 As FILETIME) _
    As Long
        
Private Declare Function SystemTimeToFileTime _
    Lib "kernel32" _
    (lpSystemTime As SYSTEMTIME, _
        lpFileTime As FILETIME) _
    As Long
    
Private Declare Function FileTimeToLocalFileTime _
    Lib "kernel32" _
    (lpFileTime As FILETIME, _
        lpLocalFileTime As FILETIME) _
    As Long
        
Private Declare Function LocalFileTimeToFileTime _
    Lib "kernel32" _
    (lpLocalFileTime As FILETIME, _
        lpFileTime As FILETIME) _
        As Long
        
Private Declare Function FileTimeToSystemTime _
    Lib "kernel32" _
    (lpFileTime As FILETIME, _
        lpSystemTime As SYSTEMTIME) _
    As Long
    
Private Declare Function GetShortPathName _
    Lib "kernel32" _
    Alias "GetShortPathNameA" _
    (ByVal lpszLongPath As String, _
        ByVal lpszShortPath As String, _
        ByVal cchBuffer As Long) _
    As Long

Private Const MAX_PATH As Long = 260

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

Private Declare Function FindFirstFile _
    Lib "kernel32" _
    Alias "FindFirstFileA" _
    (ByVal lpFileName As String, _
        lpFindFileData As WIN32_FIND_DATA) _
    As Long
    
Private Declare Function WNetGetConnection _
    Lib "mpr.dll" _
    Alias "WNetGetConnectionA" _
    (ByVal lpszLocalName As String, _
        ByVal lpszRemoteName As String, _
        cbRemoteName As Long) _
    As Long

Private Declare Function GetComputerName _
    Lib "kernel32" _
    Alias "GetComputerNameA" _
    (ByVal lpBuffer As String, _
        nSize As Long) _
    As Long

Private Const MAX_COMPUTERNAME_LENGTH = 255

Private Type VS_FILEVERSION
    FileVersion As String
    FileVersionMSl As Integer
    FileVersionMSh As Integer
    FileVersionLSl As Integer
    FileVersionLSh As Integer
    ProductVersion As String
    ProductVersionMSl As Integer
    ProductVersionMSh As Integer
    ProductVersionLSl As Integer
    ProductVersionLSh As Integer
End Type

Private Type VS_FIXEDFILEINFO
        dwSignature As Long
        dwStrucVersionl As Integer
        dwStrucVersionh As Integer
        dwFileVersionMSl As Integer
        dwFileVersionMSh As Integer
        dwFileVersionLSl As Integer
        dwFileVersionLSh As Integer
        dwProductVersionMSl As Integer
        dwProductVersionMSh As Integer
        dwProductVersionLSl As Integer
        dwProductVersionLSh As Integer
        dwFileFlagsMask As Long
        dwFileFlags As Long
        dwFileOS As Long
        dwFileType As Long
        dwFileSubtype As Long
        dwFileDateMS As Long
        dwFileDateLS As Long
End Type

Private Declare Function GetFileVersionInfo _
    Lib "Version.dll" _
    Alias "GetFileVersionInfoA" _
    (ByVal lptstrFilename As String, _
        ByVal dwHandle As Long, _
        ByVal dwLen As Long, _
        lpData As Any) _
As Long

Private Declare Function GetFileVersionInfoSize _
    Lib "Version.dll" _
    Alias "GetFileVersionInfoSizeA" _
    (ByVal lptstrFilename As String, _
        lpdwHandle As Long) _
    As Long
    
Private Declare Function VerQueryValue _
    Lib "Version.dll" _
    Alias "VerQueryValueA" _
    (pBlock As Any, _
        ByVal lpSubBlock As String, _
        lplpBuffer As Any, _
        puLen As Long) _
    As Long
    
Private Declare Sub MoveMemory _
    Lib "kernel32" _
    Alias "RtlMoveMemory" _
    (dest As Any, _
        ByVal Source As Long, _
        ByVal length As Long)

Public Function ConvertPathToUNC(strPath As String) As String
    ' Comments  : Converts the named path to its UNC representation
    ' Parameters: strPath - path (filename and extension are optional)
    ' Returns   : string UNC path
    ' Source    : Total VB SourceBook 5
    '
    Dim strTmp As String
    Dim strTmp2 As String
    Dim intPos As Integer
    Dim intRet As Integer
    Dim intTmp As Integer
    Dim astrTmp() As String
    Dim strDriveLetter As String
    Dim strComputerName As String
    Dim lngComputerNameLength As Long
    Const cstrNetHeader As String = "\\"
    Const cintBufLen As Integer = 2048
        
    On Error GoTo PROC_ERR
        
    ' Get the passed name into a temporary variable
    strTmp = strPath
    
    ' Get the Drive part of the path
    strDriveLetter = DriveFromPath(strTmp)
    
    If InStr(strTmp, cstrNetHeader) = 0 Then
    
        ' It doesn't have the \\, so its a remote drive
        
        ' Build a buffer
        strTmp2 = String$(cintBufLen, 0)
        
        ' Call the API to get the name
        intRet = WNetGetConnection(strDriveLetter, strTmp2, cintBufLen - 1)
        
        ' Did the call succeed?
        If intRet = 0 Then
            
            ' The API call returned a null-terminated string, so let's
            ' convert it to an array.
            intTmp = NullTerminatedStringToArray(strTmp2, astrTmp())
            
            ' Format the return value
            If Right(astrTmp(0), 1) <> "\" Then
                strTmp = astrTmp(0) & "\" & PathFromFullPath(strTmp)
            Else
                strTmp = astrTmp(0) & PathFromFullPath(strTmp)
            End If
        
        Else
            ' It's a local drive. Get the computer name as the root of the UNC path
            strComputerName = String(MAX_COMPUTERNAME_LENGTH + 1, 0)
            lngComputerNameLength = MAX_COMPUTERNAME_LENGTH
            
            ' Call the API
            Call GetComputerName(strComputerName, lngComputerNameLength)
            
            ' Massage the return value
            strComputerName = Mid(strComputerName, 1, lngComputerNameLength)
            strTmp = cstrNetHeader & strComputerName & "\" & strTmp
            
        End If
        
        ' Look for the colon
        intPos = InStr(strTmp, ":")
        If intPos > 0 Then
            ' Its there, so strip it off
            strTmp = Left(strTmp, intPos - 1) & Mid(strTmp, intPos + 1)
        End If
        
    End If
            
PROC_EXIT:
    ConvertPathToUNC = strTmp
    Exit Function
    
PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
        "ConvertPathToUNC"
    Resume PROC_EXIT
        
End Function

Private Function DriveFromPath(strPath As String) As String
    ' Comments  : Returns the drive letter part of the path
    ' Parameters: strPath - path containing the drive letter
    ' Returns   : String drive letter
    ' Source    : Total VB SourceBook 5
    '
    Dim intPos As Integer
    Dim strTmp As String
    Const cstrDelimiter As String = ":\"
    
    On Error GoTo PROC_ERR
    
    ' Initialize the return value
    strTmp = ""
    
    ' See of the colon and backslash exist
    intPos = InStr(strPath, cstrDelimiter)
    
    If intPos > 0 Then
        ' They exist, so return the remainder
        strTmp = Left(strPath, intPos)
    Else
        ' Look for the colon
        intPos = InStr(strPath, ":")
        If intPos > 0 Then
            ' It exists so return the remainder
            strTmp = Left(strPath, intPos)
        Else
            ' No drive letter information, so return a zero-length string
            strTmp = ""
        End If
    End If
    
PROC_EXIT:
    DriveFromPath = strTmp
    Exit Function
    
PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
        "DriveFromPath"
    Resume PROC_EXIT
    
End Function

Private Function GetFileHandle(strFile As String) As Long
    ' Comments  : Open a file to get a handle to it. We then use
    '             this handle for other calls.
    ' Parameters: strFile - Full path and name of the file
    ' Returns   : Long integer handle or -1 if an error occurred
    ' Source    : Total VB SourceBook 5
    '
    Const clngReadMode As Long = &H80000000

    On Error GoTo PROC_ERR
    
    ' Call the API to get a handle
    GetFileHandle = CreateFile( _
        strFile, _
        clngReadMode, _
        FILE_SHARE_READ Or FILE_SHARE_WRITE, _
        ByVal 0&, OPEN_EXISTING, _
        FILE_ATTRIBUTE_NORMAL Or _
        FILE_FLAG_RANDOM_ACCESS, 0&)
        
PROC_EXIT:
    Exit Function
    
PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
        "GetFileHandle"
    Resume PROC_EXIT
                
End Function

Public Function GetFileSize(strFile As String) As Long
    ' Comments  : Returns the size of the specified file
    ' Parameters: strFile - Path and name of the file to check
    ' Returns   : Long integer size in bytes, or -1 if an error occurs
    ' Source    : Total VB SourceBook 5
    '
    Dim lngFileHandle As Long
    Dim lngSizeLower As Long
    Dim lngSizeUpper As Long
        
    On Error GoTo PROC_ERR
    
    ' Get a handle to the file
    lngFileHandle = GetFileHandle(strFile)
    
    ' Make sure we got a valid handle
    If lngFileHandle > 0 Then
    
        ' Call the API to get the size
        lngSizeLower = CGetFileSize(lngFileHandle, lngSizeUpper)
        
        GetFileSize = lngSizeLower
        
        ' Close the file
        Call CloseHandle(lngFileHandle)
    Else
            GetFileSize = -1
    End If
    
PROC_EXIT:
    Exit Function
    
PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
        "GetFileSize"
    Resume PROC_EXIT
    
End Function

Public Function GetFileTime( _
    strFile As String, _
    eType As EnumFileTimeType, _
    fLocalTime As Boolean) _
    As Date
    ' Comments  : Returns the specified file time for the specified file.
    ' Parameters: strFile - Full path and name of the file
    '             eType - Time to return according to then
    '             EnumFileTimeType enumerated type
    '             fLocalTime - True to return local time as defined by the
    '             current operating system settings, False to return UTC
    '             time (also known as GMT)
    ' Returns   : Date/time value
    ' Source    : Total VB SourceBook 5
    '
    Dim ftFTCreate As FILETIME
    Dim ftFTAccess As FILETIME
    Dim ftFTWrite As FILETIME
    Dim lngFileHandle As Long
    Dim fOK As Boolean
    
    On Error GoTo PROC_ERR
        
    ' Get a handle to the file
    lngFileHandle = GetFileHandle(strFile)
    
    ' Make sure we got a valid handle
    If lngFileHandle > 0 Then
    
        ' Call the API to get the data
        fOK = CBool(CGetFileTime( _
            lngFileHandle, _
            ftFTCreate, _
            ftFTAccess, _
            ftFTWrite))
                
        If fOK Then
            
            ' Convert the Windows time format to something we can use
            Select Case eType
            
                Case ftt_Accessed
                    GetFileTime = WinTimeToVBATime(ftFTAccess, fLocalTime)
                
                Case ftt_Modified
                    GetFileTime = WinTimeToVBATime(ftFTWrite, fLocalTime)
                
                Case ftt_Created
                    GetFileTime = WinTimeToVBATime(ftFTCreate, fLocalTime)
                
            End Select
            
            ' Be sure to close the file when we're done
            CloseHandle lngFileHandle
        End If
    End If
PROC_EXIT:
    Exit Function
    
PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
        "GetFileTime"
    Resume PROC_EXIT
    
End Function

Public Function GetFileVersion(strFilename As String) As String
    ' Comments  : Returns version information from the specified file
    ' Parameters: strFileName - name of the file
    ' Returns   : string version number
    ' Source    : Total VB SourceBook 5
    '
    Dim lngRet As Long
    Dim lngTmp  As Long
    Dim astrBuf() As Byte
    Dim lngBufLen As Long
    Dim lngVerPointer As Long
    Dim VerBuffer As VS_FIXEDFILEINFO
    Dim lVerbufferLen As Long
    Dim strTmp As String

    On Error GoTo PROC_ERR
    
    ' Initialize the return value
    strTmp = ""
    
    ' Get the size needed for the buffer
    lngBufLen = GetFileVersionInfoSize(strFilename, lngTmp)
    
    If lngBufLen >= 1 Then
        ' Call succeeded, grow the array
        ReDim astrBuf(lngBufLen)
        
        ' Get the version information
        lngRet = GetFileVersionInfo(strFilename, 0&, lngBufLen, astrBuf(0))
        lngRet = VerQueryValue(astrBuf(0), "\", lngVerPointer, lVerbufferLen)
        MoveMemory VerBuffer, lngVerPointer, Len(VerBuffer)
    
        ' Get the constituent parts into the string
        strTmp = Format$(VerBuffer.dwFileVersionMSh) & "." & _
            Format$(VerBuffer.dwFileVersionMSl, "00") & "." & _
            Format$(VerBuffer.dwFileVersionLSh) & "."
        If VerBuffer.dwFileVersionLSl > 0 Then
            strTmp = strTmp & Format$(VerBuffer.dwFileVersionLSl, "0000")
        Else
            strTmp = strTmp & "0"
        End If
    End If
    
PROC_EXIT:
    GetFileVersion = strTmp
    Exit Function
    
PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
        "GetFileVersion"
    Resume PROC_EXIT
    
End Function

Public Function GetLongFileName( _
    ByVal strShortPath As String) _
    As String
    ' Comments  : Returns the full path and name of a file in long-file name
    '             format, given a path in short-file-name format.
    ' Parameters: strShortPath - the path to the file or folder
    '             in short-file-name format
    ' Returns   : The full long-file-name version of the path,
    '             if it exists, otherwise ""
    ' Source    : Total VB SourceBook 5
    '
    Dim FindData As WIN32_FIND_DATA
    Dim strResult As String
    Dim strPathTmp As String
    Dim strFileDir As String
    Dim lngResult As Long
    Dim lngSlashPos As Long

    On Error GoTo PROC_ERR

    ' Begin with the original value
    strPathTmp = strShortPath
    
    ' When no more matches are found, FindFirstFile returns -1
    Do Until lngResult = -1
        
        ' Find file matching search path
        lngResult = FindFirstFile(strPathTmp, FindData)
        
        If lngResult <> -1 Then
            ' File exists
            
            ' Get final file or path name portion of the file
            strFileDir = FindData.cFileName
            
            ' Remove trailing null characters
            strFileDir = Left$(strFileDir, InStr(strFileDir, vbNullChar) - 1)
            
            ' Find path portion of the file
            For lngSlashPos = Len(strPathTmp) To 0 Step -1
                Debug.Print lngSlashPos 'temp
                If Mid$(strPathTmp, lngSlashPos, 1) = "\" Then
                    Exit For
                End If
            Next lngSlashPos
            
            If lngSlashPos <> 0 Then
                strPathTmp = Left$(strPathTmp, lngSlashPos - 1)
            End If
            
            If strResult = "" Then
                strResult = strFileDir
            Else
                strResult = strFileDir & "\" & strResult
            End If
            
        End If
    
        If InStr(strPathTmp, "\") = 0 Then
            lngResult = -1
        End If
    
    Loop
    
    ' Add remaining drive or UNC prefix to path
    If strShortPath <> strPathTmp Then
        strResult = strPathTmp & "\" & strResult
    End If
    
    GetLongFileName = strResult

PROC_EXIT:
    Exit Function

PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
        "GetLongFileName"
    Resume PROC_EXIT

End Function

Public Function GetShortFileName( _
    ByVal strLongPath As String) _
    As String
    ' Comments  : Gets the DOS-compatible short-file-name version of a
    '             long file name
    ' Parameters: strLongPath - path to check. May be either
    '             a file name or a directory
    ' Returns   : The short file name version, if the
    '             file exists, otherwise ""
    ' Source    : Total VB SourceBook 5
    '
    On Error GoTo PROC_ERR
    
    Dim strShortFileName As String
    Dim lngSize As Long
    Dim lngResult As Long
    
    ' Allocate size in buffer to hold result
    strShortFileName = Space$(128)
    
    ' Determine length of test string
    lngSize = Len(strShortFileName)

    ' Return length of resulting string if found, and fill in
    ' the strShortfileName variable
    lngResult = GetShortPathName(strLongPath, strShortFileName, lngSize)

    ' Trim off trailing nulls
    strShortFileName = Left$(strShortFileName, lngResult)
    
    ' Return the value
    GetShortFileName = strShortFileName

PROC_EXIT:
    Exit Function

PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
        "GetShortFileName"
    Resume PROC_EXIT

End Function

Private Function NullTerminatedStringToArray( _
    strBuf As String, _
    astrRet() As String) _
    As Integer
    ' Comments  : Converts a null-terminated string to the passed array
    ' Parameters: strBuf - null terminated string
    '             astrRet() - Set: array of strings to fill
    ' Returns   : Number of elements
    ' Source    : Total VB SourceBook 5
    '
    Dim intIndex As Integer
    Dim strTmpBuf As String
    Dim intPos As Integer
    Dim intRet As Integer
    
    On Error GoTo PROC_ERR
    
    ' Assign the input value to a temporary variable
    strTmpBuf = strBuf
        
    ' Initialize the index pointer
    intIndex = -1
    
    ' Assume failure
    intRet = 0
        
    ' Loop through the input string
    Do
        intPos = InStr(strTmpBuf, vbNullChar)

        ' If we are still in bounds
        If intPos > 1 Then
            ' Grow the array
            intIndex = intIndex + 1
            ReDim Preserve astrRet(intIndex + 1)
            
            ' Assign the values
            astrRet(intIndex) = Left(strBuf, intPos - 1)
            strTmpBuf = Mid(strBuf, intPos + 1)
        End If
    Loop Until intPos <= 1
        
    ' Get the return value
    intRet = intIndex + 1
        
PROC_EXIT:
    NullTerminatedStringToArray = intRet
    Exit Function
    
PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
        "NullTerminatedStringToArray"
    Resume PROC_EXIT
    
End Function

Public Function PathFromFullPath(strPath As String)
    ' Comments  : Returns the remainder of a path after the drive letter
    ' Parameters: strPath - path to parse
    ' Returns   : Path without the drive letter
    ' Source    : Total VB SourceBook 5
    '
    Dim intPos As Integer
    Dim strTmp As String
    
    On Error GoTo PROC_ERR
    
    ' Initialize return value
    strTmp = ""
    
    ' Try to find the beginning
    intPos = InStr(strPath, ":\")
    
    If intPos > 0 Then
        ' It exists, so grab the remainder
        strTmp = Mid(strPath, intPos + 2)
    Else
        ' Try looking for just the colon part
        intPos = InStr(strPath, ":")
        If intPos > 0 Then
            ' It exists, so grab the remainder
            strTmp = Mid(strPath, intPos + 1)
        Else
            ' No drive letter information, so just return the input
            strTmp = strPath
        End If
    End If
    
PROC_EXIT:
    PathFromFullPath = strTmp
    Exit Function
    
PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
        "PathFromFullPath"
    Resume PROC_EXIT
    
End Function

Public Function SetFileTime( _
    strFile As String, _
    datIn As Date, _
    eType As EnumFileTimeType) _
    As Date
    ' Comments  : Sets the specified time on the specified file
    ' Parameters: strFile - Path and name of the file to change
    '             datIn - date/time value to change to
    '             eType - which time to change as defined by the
    '             EnumFileTimeType enumerated type
    ' Returns   : New time set (should be equal to the passed time)
    ' Source    : Total VB SourceBook 5
    '
    Dim ftFileTemp As FILETIME
    Dim ftFileCreated As FILETIME
    Dim ftFileAccessed As FILETIME
    Dim ftFileModified As FILETIME
    Dim lngFileHandle As Long
    Dim fOK As Boolean
    
    On Error GoTo PROC_ERR
                
    ' Get a handle to the file
    lngFileHandle = GetFileHandle(strFile)
    
    ' Make sure we got a valid handle
    If lngFileHandle > 0 Then
        
        ' Convert callers time to the correct format
        Call VBATimeToFileTime(datIn, ftFileTemp, True)
        
        ' What are the file's current date/time values?
        If CBool(CGetFileTime( _
            lngFileHandle, _
            ftFileCreated, _
            ftFileAccessed, _
            ftFileModified)) Then

            ' Decide which time to change
            Select Case eType
            
                Case ftt_Created
                    ftFileCreated = ftFileTemp
                
                Case ftt_Accessed
                    ftFileAccessed = ftFileTemp
                
                Case ftt_Modified
                    ftFileModified = ftFileTemp
            
            End Select
            
            ' Change the date/time
            fOK = CBool(CSetFileTime( _
                lngFileHandle, _
                ftFileCreated, _
                ftFileAccessed, _
                ftFileModified))
            
            ' Set the return value
            If fOK Then
                SetFileTime = datIn
            Else
                SetFileTime = vbNull
            End If
        End If
        
        ' Close the file
        Call CloseHandle(lngFileHandle)
        
    End If
    
PROC_EXIT:
    Exit Function
    
PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
        "SetFileTime"
    Resume PROC_EXIT
        
End Function

Private Function SystemTimeToVBATime( _
    stSystemTime As SYSTEMTIME) _
    As Date
    ' Comments  : Converts the Windows SYSTEMTIME struct to a VBA Date type
    ' Parameters: stSystemTime - Windows SYSTEMTIME struct to convert
    ' Returns   : VBA Date type
    ' Source    : Total VB SourceBook 5
    '
    On Error GoTo PROC_ERR
    
    ' Use native VBA calls to build the date type
    SystemTimeToVBATime = _
        DateSerial(stSystemTime.intYear, _
        stSystemTime.intMonth, _
        stSystemTime.intDay) + _
        TimeSerial(stSystemTime.intHour, _
        stSystemTime.intMinute, _
        stSystemTime.intSecond)

PROC_EXIT:
    Exit Function
    
PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
        "SystemTimeToVBATime"
    Resume PROC_EXIT
    
End Function

Private Sub VBATimeToFileTime( _
    datIn As Date, _
    ftFileTime As FILETIME, _
    fLocalTime As Boolean)
    ' Comments  : Converts the passed date value to the Windows FILETIME
    '             structure.
    ' Parameters: datIn - value to change
    '             ftFileTime - Windows FILETIME struct to convert into
    '             fLocalTime - True to return local time as defined by the
    '             current operating system settings, False to return UTC
    '             time (also known as GMT)
    ' Returns   : Nothing
    ' Source    : Total VB SourceBook 5
    '
    On Error GoTo PROC_ERR
        
    Dim stSystemTime As SYSTEMTIME
    Dim ftFileTmp As FILETIME
    Dim fOK As Boolean
        
    ' First convert to system time
    Call VBATimeToSystemTime(datIn, stSystemTime)
        
    ' Next convert the system time to a file time
    fOK = CBool(SystemTimeToFileTime(stSystemTime, ftFileTime))
    
    ' Did it work?
    If fOK Then
        ' Handle local time
        If fLocalTime Then
            Call LocalFileTimeToFileTime(ftFileTime, ftFileTmp)
            ftFileTime = ftFileTmp
        End If
    End If
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
        "VBATimeToFileTime"
    Resume PROC_EXIT
    
End Sub

Private Sub VBATimeToSystemTime( _
    datIn As Date, _
    stSystemTime As SYSTEMTIME)
    ' Comments  : Converts the supplied VBA date/time value to the
    '             Windows SYSTEMTIME struct
    ' Parameters: datIn - date to convert
    '             stSystemTime - Windows SYSTEMTIME struct to convert into
    ' Returns   : Nothing
    ' Source    : Total VB SourceBook 5
    '
    On Error GoTo PROC_ERR
    
    ' Convert using VB intrinsic functions
    With stSystemTime
        .intMonth = Month(datIn)
        .intDay = Day(datIn)
        .intYear = Year(datIn)
        .intHour = Hour(datIn)
        .intMinute = Minute(datIn)
        .intSecond = Second(datIn)
    End With
    
PROC_EXIT:
    Exit Sub
    
PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
        "VBATimeToSystemTime"
    Resume PROC_EXIT
    
End Sub

Private Function WinTimeToVBATime( _
    ftWinFileTime As FILETIME, _
    fLocalTime As Boolean) _
    As Date
    ' Comments  : Converts the Windows FILETIME struct to a VBA Date type
    ' Parameters: ftWinFileTime - Windows FILETIME struct to convert
    '             fLocalTime - True to return local time as defined by the
    '             current operating system settings, False to return UTC
    '             time (also known as GMT)
    ' Returns   : VBA Date type
    ' Source    : Total VB SourceBook 5
    '
    Dim stSystemFileTime As SYSTEMTIME
    Dim ftLocalFileTime As FILETIME
        
    On Error GoTo PROC_ERR
    
    ' Was local time requested?
    If fLocalTime Then
        ' Need to convert the file time to a local time
        FileTimeToLocalFileTime ftWinFileTime, ftLocalFileTime
        ftWinFileTime = ftLocalFileTime
    End If
    
    ' Call the API to convert the time format
    If CBool(FileTimeToSystemTime(ftWinFileTime, stSystemFileTime)) Then
        WinTimeToVBATime = SystemTimeToVBATime(stSystemFileTime)
    End If
    
PROC_EXIT:
    Exit Function
    
PROC_ERR:
    MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
        "WinTimeToVBATime"
    Resume PROC_EXIT
    
End Function


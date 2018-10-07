Attribute VB_Name = "modFileProperties"
Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ©1996-2001 VBnet, Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' You are free to use this code within your own applications,
' but you are expressly forbidden from selling or otherwise
' distributing this source code without prior written consent.
' This includes both posting free demo projects made from this
' code as well as reproducing the code in text or html format.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Type SHELLEXECUTEINFO
    cbSize        As Long
    fMask         As Long
    hwnd          As Long
    lpVerb        As String
    lpFile        As String
    lpParameters  As String
    lpDirectory   As String
    nShow         As Long
    hInstApp      As Long
    lpIDList      As Long     'Optional parameter
    lpClass       As String   'Optional parameter
    hkeyClass     As Long     'Optional parameter
    dwHotKey      As Long     'Optional parameter
    hIcon         As Long     'Optional parameter
    hProcess      As Long     'Optional parameter
End Type

Public Const SEE_MASK_INVOKEIDLIST = &HC
Public Const SEE_MASK_NOCLOSEPROCESS = &H40
Public Const SEE_MASK_FLAG_NO_UI = &H400

Declare Function ShellExecuteEx Lib "shell32.dll" (SEI As SHELLEXECUTEINFO) As Long

'****************************************************************
' Name: Implementing the API File Property Page
' Description:Code to invoke the property page information f
'     or any file on the system. This is identical to the properti
'     es displayed by right-clicking a file in Explorer. Using the
'      SHELLEXECUTEINFO type structure, and the API ShellExecuteEx
'     (), VB4 (32 bit) and VB5 applications can display the Window
'     s 95 or NT4 file property page using this simple routine. As
'      long as the path to the file is known, this routine can be
'     invoked. It works for both registered and unregistered Windo
'     ws file types, as well as bringing up the DOS property sheet
'      for DOS applications or files (try pointing the app to auto
'     exec.bat.
' By: VB Net (Randy Birch)
'
' Inputs:None
' Returns:None
' Assumes:Thanks go out to Ian Land for providing the Delphi code, and Roy Meyer.
' Start a new project, and to the form add a command button (cmdProperties), and a text box (Text1). Add a BAS module (Module1), and the following code to the project.
' Side Effects:None
'
'Code provided by Planet Source Code(tm) 'as is', without
'     warranties as to performance, fitness, merchantability,
'     and any other warranty (whether expressed or implied).
'****************************************************************

Public Sub ShowProperties(Filename As String, OwnerhWnd As Long)

On Error GoTo ErrorHandler

    'open a file properties property page for
    'specified file if return value
    
    Dim SEI As SHELLEXECUTEINFO
    
    'Fill in the SHELLEXECUTEINFO structure
    With SEI
       .cbSize = Len(SEI)
       .fMask = SEE_MASK_NOCLOSEPROCESS Or _
                SEE_MASK_INVOKEIDLIST Or _
                SEE_MASK_FLAG_NO_UI
       .hwnd = OwnerhWnd
       .lpVerb = "properties"
       .lpFile = Filename
       .lpParameters = vbNullChar
       .lpDirectory = vbNullChar
       .nShow = 0
       .hInstApp = 0
       .lpIDList = 0
    End With
    
    'call the API to display the property sheet
    Call ShellExecuteEx(SEI)
    
    Exit Sub
ErrorHandler:
    Exit Sub
End Sub





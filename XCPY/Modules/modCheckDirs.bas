Attribute VB_Name = "modCheckDirs"
Option Explicit
'~modCheckDirs.bas;modExpandEnvStrings.bas;modGetFileAndDirData.bas;modAddSlash.bas;
'Directory/File verification module
'*************************************************************************
' modCheckDirs: Directory/File verification module
'
' This module provides the following functions:
' FileExists(): check for a file's existence
' FileInUse():  Determines whether the specified file is currently in use
' IsFile():     see if the path is an existing file
' DirExists():  check for a directory's existence
' IsDir():      see if the path is an existing directory
' CheckDirs():  see if a file's directory exists. Build a path up to it if not.
'               Note that if you are checking the path TO a directory, the
'               path up TO it, NOT INCLUDING it, will be built. You should use
'               the MkDir command to create it. ***NOTE***: If you need just this
'               function from the module, see modVerifyPathExists for the
'               VerifyPathExists() function that uses a single API call and does
'               not require you to reference the Scrrun.dll and other modules.
'
' NOTE: This routine expects a project reference to
'       "Microsoft Scripting Runtime" (scrrun.dll)
'
' NOTE: This routine uses "modExpandEnvStrings.bas"
' NOTE: This routine uses "modGetFileAndDirData.bas"
' NOTE: This routine uses "modAddSlash.bas"
'*************************************************************************

'*************************************************
' Private API calls and data used by the support routines
'*************************************************
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

'*************************************************
' check for first match to specification
'*************************************************
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long

'*************************************************
' close an open check
'*************************************************
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
'
' File creation attributes
'
Private Const GENERIC_WRITE As Long = &H40000000
Private Const OPEN_EXISTING As Long = 3
Private Const FILE_ATTRIBUTE_NORMAL As Long = &H80
Private Const FILE_FLAG_WRITE_THROUGH As Long = &H80000000
Private Const INVALID_HANDLE_VALUE As Long = -1
Private Const ERROR_SHARING_VIOLATION As Long = 32

Private Type FILETIME
  dwLowDateTime As Long
  dwHighDateTime As Long
End Type
'
' text attributes
'
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const MAX_PATH = 260
'
' attribute structure
'
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
'
' flag for checking directory types
'
Public Const Normal As Long = vbSystem Or vbVolume Or vbDirectory

'***************************************************************
' FileInUse(): Determines whether the specified file is currently in use
'***************************************************************
Public Function FileInUse(ByVal FileName As String) As Boolean
  Dim hFile As Long, fname As String
  Dim Attr As VbFileAttribute
  
  On Error Resume Next
  fname = Trim$(FileName)
'
' if file does not exist then it should be OK
'
  hFile = Len(Dir$(fname, vbDirectory Or vbHidden Or vbSystem))
  If hFile = 0 Or CBool(Err.Number) Then Exit Function
'
' see if the file is actually a directory, read-only, or a volume name
'
  On Error Resume Next
  Attr = GetAttr(fname)
  If CBool(Err.Number) Then Exit Function
  On Error GoTo 0
  If (Attr And vbDirectory) Or (Attr And vbReadOnly) Or (Attr And vbVolume) Then Exit Function
'
' try to open the file. Check for a sharing violation
'
  hFile = CreateFile(fname, GENERIC_WRITE, 0&, 0&, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL Or FILE_FLAG_WRITE_THROUGH, 0&)
  If hFile = INVALID_HANDLE_VALUE Then
    FileInUse = (Err.LastDllError = ERROR_SHARING_VIOLATION)
  Else
    CloseHandle hFile
  End If
End Function

'***************************************************************
' check for a file's existence
'***************************************************************
Public Function FileExists(FileName As String) As Boolean
  FileExists = (CheckDFExists(FileName) = 1)
End Function

'***************************************************************
' see if the path is an existing file
'***************************************************************
Public Function IsFile(FileName As String) As Boolean
  IsFile = FileExists(FileName) 'wrap FileExists() function (for symantics use)
End Function

'***************************************************************
' check for a directory's existence
'***************************************************************
Public Function DirExists(DirName As String) As Boolean
  Dim S As String, flg As Integer, DName As String
  Dim fso As FileSystemObject, drv As Drive
  
  DName = Trim$(DirName)
  If Right$(DName, 1) = ":" Then DName = DName & "\"
  flg = CheckDFExists(DName)          'get result
  If flg = -1 Then                    'does not exist
    If LCase$(DName) = AddSlash(LCase$(GetDriveName(DName))) Then 'checking for drive?
      Set fso = New FileSystemObject                              'yes, is drive ready?
      Set drv = fso.GetDrive(fso.GetDriveName(DName))
      If drv.IsReady Then DirExists = True                        'if so, then OK
      Set fso = Nothing
      Exit Function
    End If
    On Error Resume Next
    S = Dir$(DName, vbDirectory)      'see if root (only time this will happen)
    If CBool(Err.Number) Then Exit Function 'drive does not exist
    On Error GoTo 0
    If Len(S) Then DirExists = True   'set to true if so
  Else
    DirExists = (flg = 0)             'else set flag according to type
  End If
End Function

'***************************************************************
' see if the path is an existing directory
'***************************************************************
Public Function IsDirectory(DirName As String) As Boolean
  IsDirectory = DirExists(DirName)
End Function

'***************************************************************
' see if a file's directory exists. Build a path up to it if not
'***************************************************************
Public Sub CheckDirs(Src As String)
  Dim Fname1 As String, Fname2 As String
  Dim Path1 As String, Path2 As String
  Dim Level As Integer
  
  Src = ExpandEnvStrings(Src)           'expand environment strings
  Fname1 = GetFileName(Src)
  Path1 = GetFilePath(Src)
'
' ssaj
'
  Do
    Level = 0
    Path2 = Src                         'init with copy of source
    Do
      Fname2 = GetFileName(Path2)       'trim off a path from right
      If Not CBool(Len(Fname2)) Then Exit Sub
      Path2 = GetFilePath(Path2)        'get path to path2, set as current
      If DirExists(Path2) Then Exit Do  'exit if path found
      Level = Level + 1
    Loop
    If Fname1 <> Fname2 Or Level > 0 Then
      On Error Resume Next
      MkDir AddSlash(Path2) & Fname2    'create as needed
      If CBool(Err.Number) Then
'        MsgBox "Could not create path '" & AddSlash(Path2) & Fname2 & "'." & vbCrLf & _
'               "Error: " & Err.Description, vbOKOnly Or vbExclamation, "Path Creation Error"
        Exit Sub
      End If
      On Error GoTo 0
    End If
    If DirExists(Path1) Then Exit Do    'if full path exists, then done
  Loop
End Sub

'***************************************************************
' see if a file or directory exists. 0=dir, 1=file, -1=nothing exists
'***************************************************************
Private Function CheckDFExists(Path As String) As Integer
  Dim hFind As Long
  Dim wFind As WIN32_FIND_DATA
  
  CheckDFExists = -1                    'init to not exists
  hFind = FindFirstFile(Path, wFind)    'get first match
  If hFind <> INVALID_HANDLE_VALUE Then 'does it exist?
    FindClose hFind                     'yes, close handle
    CheckDFExists = 1                   'init result to file, then check for DIR
    If wFind.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY Then CheckDFExists = 0
  End If
End Function

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************


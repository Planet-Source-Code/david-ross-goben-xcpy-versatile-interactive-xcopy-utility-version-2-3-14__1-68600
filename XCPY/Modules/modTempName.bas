Attribute VB_Name = "modTempName"
Option Explicit
'~modTempName.bas;
'Obtain a unique temp filename or temp directory
'**********************************************************************
' modTempName - Obtain a unique temp filename or temp directory
'
'The following functions are provided:
' GetTempName(): Get a temporary file/dir name in the system's temp directory
' GetTempDir():  Get the path to the TEMP directory
'**********************************************************************

'***************************************************************
' API calls and declarations
'***************************************************************
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long

Private Const MAX_PATH = 260

'**********************************************************************
' GetTempName(): get a temporary file/dir name in the system's temp directory
'**********************************************************************
' strHeader is a header string for the new name. For example,
' supplying "DBW" will create a file named DBWxxx.TMP, where "xxx" is
' system-supplied info that will make this name unique.
'
Public Function GetTempName(strHeader As String) As String
  Dim strName As String, strTemp As String
  Dim intLen As Long
    
  strTemp = String$(MAX_PATH + 1, 0)              'init path string
  strName = GetTempDir()                          'get temp directory
  If Len(strName) Then                            'if it was found
    intLen = GetTempFileName(strName, strHeader, 0, strTemp)
    If intLen Then                                'if temp file made
      intLen = InStr(1, strTemp, vbNullChar) - 1  'find terminator
      strName = Left$(strTemp, intLen)            'grab path
      Call DeleteFile(strName)                    'delete the file created
      GetTempName = strName                       'return the name found
    End If
  End If
End Function

'**********************************************************************
' GetTempDir(): get the path to the TEMP directory
'**********************************************************************
Public Function GetTempDir() As String
  Dim strPath As String
  Dim intLen As Long
  
  strPath = String$(MAX_PATH + 1, 0)              'init the path target
  intLen = GetTempPath(MAX_PATH, strPath)         'get the temp dir
  If intLen Then                                  'found?
    intLen = InStr(1, strPath, vbNullChar) - 1    'yes, find terminator
    GetTempDir = Left$(strPath, intLen)           'get path text w/o nulls
  End If
End Function

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************


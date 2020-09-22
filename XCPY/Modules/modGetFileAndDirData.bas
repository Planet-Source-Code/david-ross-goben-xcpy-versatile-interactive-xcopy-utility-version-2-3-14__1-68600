Attribute VB_Name = "modGetFileAndDirData"
Option Explicit
'~modGetFileAndDirData.bas;scrrun.dll;
'Various file/dir I/O routines
'*************************************************
' modGetFileAndDirData:
' These functions, provided a full path to a directory or file, return a string
' containing the target directory/file name, path beneath the target
' directory/file, or the drive for the target directory/file
'
' Provided function are:
' GetFileName():  get the target filename from a full path to the file
' GetDirName():   get the target dirname from a full path to the directory
' GetFilePath():  get a dirpath below a file from a full path to the target file
' GetDirPath():   get a dirpath below a directory from a full path to the target directory
' GetDriveName(): get drive name from a full path
'
' NOTE: This routine expects a project reference to
'       "Microsoft Scripting Runtime" (scrrun.dll)
'*************************************************

'*************************************************
' GetFileName(): get the target filename from a full path to the file
'*************************************************
Public Function GetFileName(Path As String) As String
  Dim fso As FileSystemObject
  
  Set fso = New FileSystemObject
  GetFileName = fso.GetFileName(Path)
  Set fso = Nothing
End Function

'*************************************************
' GetDirName(): get the target dirname from a full path to the directory
'*************************************************
Public Function GetDirName(Path As String) As String
  Dim fso As FileSystemObject, S As String
  
  Set fso = New FileSystemObject
  GetDirName = fso.GetFileName(Path)
  Set fso = Nothing
End Function

'*************************************************
' GetFilePath(): get a dirpath below a file from a full path to the target file
'*************************************************
Public Function GetFilePath(Path As String) As String
  GetFilePath = Left$(Path, Len(Path) - Len(GetFileName(Path)) - 1)
End Function

'*************************************************
' GetDirPath(): get a dirpath below a directory from a full path to the target directory
'*************************************************
Public Function GetDirPath(Path As String) As String
  If LCase$(GetDriveName(Path)) = LCase$(Path) Then
    GetDirPath = Path
  Else
    GetDirPath = Left$(Path, Len(Path) - Len(GetDirName(Path)) - 1)
  End If
End Function

'*************************************************
' GetDriveName(): get drive name from a full path
'*************************************************
Public Function GetDriveName(Path As String) As String
  Dim fso As FileSystemObject
  
  Set fso = New FileSystemObject
  GetDriveName = fso.GetDriveName(Path)
  Set fso = Nothing
End Function

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************


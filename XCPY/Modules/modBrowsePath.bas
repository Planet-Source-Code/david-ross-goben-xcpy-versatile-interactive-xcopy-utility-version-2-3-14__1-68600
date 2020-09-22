Attribute VB_Name = "modBrowsePath"
Option Explicit
'~modBrowsePath.bas;
'Browse path, Edit/Print File
'*************************************************************************
' modBrowsePath: This module supplies the following functions:
'
' BrowsePath():    Open an Explorer browser on the specified directory path.
' EditFilePath():  Edit the specified filepath file using the program associated with it.
' OpenFilePath():  Open the specified filepath file using the program associated with it.
' PrintFilePath(): Print the specified filepath file using the program associated with it.
' FindFilePath():  Find a file starting from a specified starting filepath using the Find file Dialog
'
'NOTE: Each function on exit returns a long error code:
' 0 = OK. Operation successful
' 1 = The OS is out of memory resources
' 2 = specified file was not found
' 3 = The specified path was not found
' 5 = Access is denied to the specified file
' 8 = There is not enough memory to complete the operation
'11 = The .EXE file is invalid (non-win32 or erro in .exe)
'26 = Sharing violation
'28 = The DDE transaction could not be completed because the request timed out
'29 = The DDE (Dynamic Data Exchange) transaction failed
'31 = There is no application associated with the specified file
'32 = If the specified file was a DLL and the DLL was not found
'*************************************************************************

'*************************************************
' API call used by BrowsePath
'*************************************************
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Const SW_NORMAL = 1

'*************************************************
' Open an Explorer browser on the specified directory path.
' Simply supply the Me.hWnd from your form for the hWndParent value.
'*************************************************
Public Function BrowsePath(hwndParent As Long, Path As String) As Long
  BrowsePath = ShellPath(hwndParent, "explore", Path)
End Function

'*************************************************
' Open the specified filepath file using the program associated with it.
' For example, *.TXT files are normally opened by NOTEPAD.EXE, and
' *.DOC files are opened by Word or WordPad.
'
' Simply supply the Me.hWnd from your form for the hWndParent value.
'*************************************************
Public Function OpenFilePath(hwndParent As Long, Path As String) As Long
  OpenFilePath = ShellPath(hwndParent, "open", Path)
End Function

'*************************************************
' Edit the specified filepath file using the program associated with it.
' For example, *.TXT files are normally opened by NOTEPAD.EXE, and
' *.DOC files are opened by Word or WordPad.
'
' Simply supply the Me.hWnd from your form for the hWndParent value.
'*************************************************
Public Function EditFilePath(hwndParent As Long, Path As String) As Long
  EditFilePath = ShellPath(hwndParent, "edit", Path)
End Function

'*************************************************
' Print the specified filepath file using the program associated with it.
' For example, *.TXT files are normally printed by NOTEPAD.EXE, and
' *.DOC files are printed by Word or WordPad.
'
' Simply supply the Me.hWnd from your form for the hWndParent value.
'*************************************************
Public Function PrintFilePath(hwndParent As Long, Path As String) As Long
  PrintFilePath = ShellPath(hwndParent, "print", Path)
End Function

'*************************************************
' Find a file starting from a specified starting filepath using the Find file Dialog
'
' Simply supply the Me.hWnd from your form for the hWndParent value.
'*************************************************
Public Function FindFilePath(hwndParent As Long, Path As String) As Long
  Dim Pth As String
  
  Pth = Trim$(Path)
  If Right$(Pth, 1) <> "\" Then Pth = Pth & "\"
  FindFilePath = ShellPath(hwndParent, "find", Pth)
End Function

'*************************************************
' support routine for module
'*************************************************
Private Function ShellPath(hwnd As Long, Cmd As String, Path As String) As Long
  Dim I As Long
  I = ShellExecute(hwnd, Cmd, Path, "", "", SW_NORMAL)
  If I = 0 Then I = 1
  If I > 32 Then I = 0
  ShellPath = I
End Function

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************


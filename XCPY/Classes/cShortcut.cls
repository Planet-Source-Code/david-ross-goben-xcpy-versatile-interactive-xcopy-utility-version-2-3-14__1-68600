VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cShortcut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'~cShortcut.cls;modAddSlash.bas;modExpandEnvStrings.bas;modGetSpecialFolder.bas;@wshom.ocx;
'Create desktop shortcuts to files/URLs
'***********************************************************
' cShortcut: Create desktop shortcuts to files/URLs. This calls
'            supports the following functions:
'
' CreateShortcut():           Create Windows Desktop Shortcut.
' CreateShortcutAt():         Create Windows Shortcut to a specified folder.
' LinkExists():               See if a target link already exists.
'                             Return True if so.
' CreateInternetShortcut():   Create link to a URL on the Windows Desktop.
' CreateInternetShortcutAt(): Create link to a URL to a specified folder.
' URLExists():                See if a target URL link already exists.
'                             Return True if so.
'---------------------------
' Also include the following lines in your code to instantiate the Shortcut class:
'   Dim sc As cShortcut
'   Set sc = New cShortcut
'
' Remove the object before you close your apps using:
'   Set sc = Nothing
'
' Create a shortcut using: sc.CreateShortcut "Shortcut Name" "Target Path"
'---------------------------
' NOTE: This Class requires that "Windows Scripting Host Object Model"
'       [wshom.ocx] be included under Project\References
'       in your VB project.
' NOTE: modAddSlash.bas module used
' NOTE: modExpandEnvStrings.bas module used
' NOTE: modGetSpecialFolder.bas module used
'***********************************************************

'''Private oshell As IWshRuntimeLibrary.IWshShell_Class   'testing...
Private oshell As Object

'*******************************************************************************
' Subroutine Name   : Class_Initialize
' Purpose           : Initialize resources
'*******************************************************************************
Private Sub Class_Initialize()
  On Error Resume Next
  Set oshell = New IWshRuntimeLibrary.IWshShell_Class
  If CBool(Err.Number) Then
    Err.Clear
    Set oshell = New IWshRuntimeLibrary.WshShell
    If CBool(Err.Number) Then Set oshell = CreateObject("Wscript.Shell")
  End If
End Sub

'*******************************************************************************
' Subroutine Name   : Class_Terminate
' Purpose           : Remove resources
'*******************************************************************************
Private Sub Class_Terminate()
  Set oshell = Nothing
End Sub

'***********************************************************
' CreateShortcut(): Create Windows Desktop Shortcut.
'    sName   = Shortcut Name for link file
'    sPath   = path to object of shortcut (ie, executable)
'    sArg    = optional Argument list for sPath
'    sDesc   = Optional description of shortcut for properties
'    sDir    = Optional working directory path
'    sIcon   = Optional path to an Icon file
'    DestLoc = Optional Destination special folder (default=Desktop)
'    WinStyle= App display style (default is WhsMinimizedFocus)
'***********************************************************
Public Sub CreateShortcut(sName As String, _
           sPath As String, _
           Optional sArg As String = vbNullString, _
           Optional sDesc As String = vbNullString, _
           Optional sDir As String = vbNullString, _
           Optional sIcon As String = vbNullString, _
           Optional DestLoc As String = "AllUsersDeskTop", _
           Optional WinStyle As WshWindowStyle = WshNormalFocus)
  Dim sDesk As String
'''  Dim oShortCut As IWshRuntimeLibrary.IWshShortcut_Class   'testing
  Dim oShortCut As Object

'''  sDesk = oshell.SpecialFolders.Item("AllUsersDesktop")    'testing
  sDesk = GetSpecialFolder(DestLoc)
  Set oShortCut = oshell.CreateShortcut(AddSlash(sDesk) & sName & ".lnk")
  oShortCut.TargetPath = ExpandEnvStrings(sPath)
  If Len(sDesc) Then oShortCut.Description = sDesc
  If Len(sArg) Then oShortCut.Arguments = ExpandEnvStrings(sArg)
  If Len(sDir) Then oShortCut.WorkingDirectory = ExpandEnvStrings(sDir)
  If Len(sIcon) Then oShortCut.IconLocation = ExpandEnvStrings(sIcon)
  oShortCut.WindowStyle = WinStyle
  oShortCut.Save
  Set oShortCut = Nothing
End Sub

'***********************************************************
' CreateShortcutAt(): Create Windows Shortcut to a specified path.
'    sFolder=Path to folder to place link
'    sName = Shortcut Name for link file
'    sPath = path to object of shortcut (ie, executable)
'    sArg  = optional Argument list for sPath
'    sDesc = Optional description of shortcut for properties
'    sDir  = Optional working directory path
'    sIcon = Optional path to an Icon file
'***********************************************************
Public Sub CreateShortcutAt(sFolder As String, _
           sName As String, _
           sPath As String, _
           Optional sArg As String = vbNullString, _
           Optional sDesc As String = vbNullString, _
           Optional sDir As String = vbNullString, _
           Optional sIcon As String = vbNullString, _
           Optional WinStyle As WshWindowStyle = WshNormalFocus)
  Dim sDesk As String
'''  Dim oShortCut As IWshRuntimeLibrary.IWshShortcut_Class   'testing
  Dim oShortCut As Object
  
  Set oShortCut = oshell.CreateShortcut(AddSlash(sFolder) & sName & ".lnk")
  oShortCut.TargetPath = ExpandEnvStrings(sPath)
  If Len(sDesc) Then oShortCut.Description = sDesc
  If Len(sArg) Then oShortCut.Arguments = ExpandEnvStrings(sArg)
  If Len(sDir) Then oShortCut.WorkingDirectory = ExpandEnvStrings(sDir)
  If Len(sIcon) Then oShortCut.IconLocation = ExpandEnvStrings(sIcon)
  oShortCut.WindowStyle = WinStyle
  oShortCut.Save
  Set oShortCut = Nothing
End Sub

'***********************************************************
' LinkExists(): see if a target desktop link already exists. Return True if so
'               Default folder to check is "Desktop"
'***********************************************************
Public Function LinkExists(LinkName As String, _
       Optional SpcFolder As String = vbNullString) As Boolean
  Dim sDesk As String, File As String
  
  If Len(SpcFolder) Then
    sDesk = Trim$(SpcFolder)
    sDesk = AddSlash(sDesk)
    sDesk = sDesk & LinkName
  Else
    sDesk = AddSlash(oshell.SpecialFolders.Item("DeskTop")) & LinkName
  End If
  If LCase$(Right$(sDesk, 4)) <> ".lnk" Then sDesk = sDesk & ".lnk"
  LinkExists = (Len(Dir$(sDesk)) > 0)
End Function

'***********************************************************
' CreateInternetShortcut(): Create link to a URL on the Windows Desktop.
'    sName = Name of Link file
'    sPath = path to object of shortcut (ie, executable)
'    sType = optional Filetype. Default = file. Other = http
'***********************************************************
Public Sub CreateInternetShortcut(sName As String, _
           sPath As String, _
           Optional sType As String = "file")
  Dim sDesk As String, fn As Integer, File As String
  
  sDesk = oshell.SpecialFolders.Item("Desktop")
  fn = FreeFile
  Open AddSlash(sDesk) & sName & ".url" For Output As #fn
  Print #fn, "[InternetShortcut]"
  If LCase$(sType) = "file" Then
    Print #fn, "URL=file://" & sPath
  Else
    Print #fn, "URL=http://" & sPath
  End If
  Close #fn
End Sub

'***********************************************************
' CreateInternetShortcutAt(): Create link to a URL to a specified path.
'    sFolder=Path to folder to place link
'    sName = Name of Link file
'    sPath = path to object of shortcut (ie, executable)
'    sType = optional Filetype. Default = file. Other = http
'***********************************************************
Public Sub CreateInternetShortcutAt(sFolder As String, _
           sName As String, _
           sPath As String, _
           Optional sType As String = "file")
  Dim sDest As String, fn As Integer, File As String
  
  sDest = Trim$(sFolder$)
  fn = FreeFile
  Open AddSlash(sDest) & sName & ".url" For Output As #fn
  Print #fn, "[InternetShortcut]"
  If LCase$(sType) = "file" Then
    Print #fn, "URL=file://" & sPath
  Else
    Print #fn, "URL=http://" & sPath
  End If
  Close #fn
End Sub

'***********************************************************
' URLExists(): see if a target URL link already exists. Return True if so.
'***********************************************************
Public Function URLExists(URLName As String) As Boolean
  Dim sDesk As String, fn As Integer, File As String
  
  sDesk = AddSlash(oshell.SpecialFolders.Item("Desktop")) & URLName
  If LCase$(Right$(sDesk, 4)) <> ".url" Then sDesk = sDesk & ".url"
  URLExists = (Len(Dir$(sDesk)) > 0)
End Function

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************


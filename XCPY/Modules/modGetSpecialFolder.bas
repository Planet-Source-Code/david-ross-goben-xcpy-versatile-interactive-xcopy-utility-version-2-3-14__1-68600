Attribute VB_Name = "modGetSpecialFolder"
Option Explicit
'~modGetSpecialFolder.bas;
'Get the long path to a windows system special folder
'***************************************************************
' modGetSpecialFolder:
' Get the long path to a windows system special folder (ignores character case)
'
' Available folders you can access:
'
'    AllUsersDesktop
'    AllUsersStartMenu
'    AllUsersPrograms
'    AllUsersStartup
'    AppData          'user folder for application-specific data
'    Desktop          'user
'    Favorites        'user screen
'    Fonts
'    MyDocuments      'user
'    NetHood          'user
'    PrintHood        'user
'    Programs         'user
'    QuickLaunch      'user explorer quicklaunch icon bar (if present)
'    Recent           'user
'    SendTo           'user
'    StartMenu        'user Start Button
'    Startup          'User StartUp folder
'    Templates        'user
'
'EXAMPLE:
'    DeskTopPath = GetSpecialFolder("desktop")
'
' NOTE: This module requires that "Windows Scripting Host Object Model"
' [wshom.ocx] be included under Project\References in your
' VB project.
'***************************************************************
Public Function GetSpecialFolder(fldr As String) As String
  Dim oshell As Object
  Dim S As String, I As Integer, tfldr As String
  
  tfldr = LCase$(Trim$(fldr))
  Set oshell = CreateObject("Wscript.Shell")
  Select Case tfldr
    Case "allusersdesktop"
      S = oshell.SpecialFolders.Item("AllUsersDesktop")
      If Len(S) = 0 Then S = oshell.SpecialFolders.Item("Desktop")
    Case "allusersstartmenu"
      S = oshell.SpecialFolders.Item("AllUsersStartMenu")
      If Len(S) = 0 Then S = oshell.SpecialFolders.Item("StartMenu")
    Case "allusersprograms"
      S = oshell.SpecialFolders.Item("AllUsersPrograms")
      If Len(S) = 0 Then S = oshell.SpecialFolders.Item("Programs")
    Case "allusersstartup"
      S = oshell.SpecialFolders.Item("AllUsersStartup")
    Case "appdata"
      S = oshell.SpecialFolders.Item("AppData")
    Case "desktop"
      S = oshell.SpecialFolders.Item("Desktop")
    Case "favorites"        'screen
      S = oshell.SpecialFolders.Item("Favorites")
    Case "fonts"
      S = oshell.SpecialFolders.Item("Fonts")
    Case "mydocuments"
      S = oshell.SpecialFolders.Item("MyDocuments")
    Case "nethood"
      S = oshell.SpecialFolders.Item("NetHood")
    Case "printhood"
      S = oshell.SpecialFolders.Item("PrintHood")
    Case "programs"
      S = oshell.SpecialFolders.Item("Programs")
    Case "recent"
      S = oshell.SpecialFolders.Item("Recent")
    Case "sendto"
      S = oshell.SpecialFolders.Item("SendTo")
    Case "startmenu"        'Start Button menu
      S = oshell.SpecialFolders.Item("StartMenu")
    Case "startup"          'User StartUp folder
      S = oshell.SpecialFolders.Item("Startup")
    Case "templates"
      S = oshell.SpecialFolders.Item("Templates")
    Case "quicklaunch"
      S = oshell.SpecialFolders.Item("AppData")
      S = S & "\Microsoft\Internet Explorer\Quick Launch"
      If Len(Dir$(S & "\*.*")) = 0 Then S = vbNullString
  End Select
  Set oshell = Nothing
  GetSpecialFolder = S
End Function

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************


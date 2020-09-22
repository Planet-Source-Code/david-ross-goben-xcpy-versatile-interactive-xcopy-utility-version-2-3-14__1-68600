Attribute VB_Name = "modSysMenu"
Option Explicit
'~modSysMenu.bas;
'Remove/disable system menu options
'***************************************************************
' modSysMenu: Functions to remove/disable system menu options on
'             an application. An additional function allows adding
'             an option to the system menu.

' The system Menu is the drop=down menu at the top-left of 95/98/Nt
' windows, and the special control buttons at the top right of the
' application window's title banner.
'
' The following functions are provided:
'
' VBRemoveMenuItem(): This routine removes the specified menu item from
'                     the control menu and the corresponding functionality
'                     from the form.
' VBAppendMenuItem(): This routine append a menu item to the control menu.
'
'***************************************************************

'***************************************************************
' API declarations necessary to work with the Control Box menus
'***************************************************************
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal revert As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal lIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal lIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long

Private Const MF_STRING = &H0&
Private Const MF_BYPOSITION = &H400
Private Const MAXIMIZE_BUTTON = &HFFFEFFFF
Private Const MINIMIZE_BUTTON = &HFFFDFFFF
Private Const GWL_STYLE = (-16)

' enumeration used when calling VBRemoveMenuItem
Public Enum RemoveMenuEnum
  rmRestore = 0
  rmMove = 1
  rmSize = 2
  rmMinimize = 3
  rmMaximize = 4
  rmClose = 6       'use rmClose-1 to remove the Separator
End Enum
Private Const SCOFFSET = 2000
Private CurrentID As Long

'***************************************************************
' VBRemoveMenuItem
' this routine removes the specified menu item from the control menu and
' the corresponding functionality from the form
'
' Parameters
' TargetForm - the form to perform the operation on
' MenuToRemove - Enum specifying which menu to remove
'***************************************************************
Public Sub VBRemoveMenuItem(ByVal TargetForm As Form, _
  ByVal MenuToRemove As RemoveMenuEnum)
  Dim hSysMenu As Long
  Dim lStyle As Long
  
  hSysMenu = GetSystemMenu(TargetForm.hwnd, 0&)
  RemoveMenu hSysMenu, MenuToRemove, MF_BYPOSITION

  Select Case MenuToRemove
    Case rmClose
      ' when removing the Close menu, also
      ' remove the separator over it
      RemoveMenu hSysMenu, MenuToRemove - 1, MF_BYPOSITION
    Case rmMinimize, rmMaximize
      ' get the current window style
      lStyle = GetWindowLong(TargetForm.hwnd, GWL_STYLE)

      If MenuToRemove = rmMaximize Then
        ' turn off bits for Maximize arrow
        ' button
        lStyle = lStyle And MAXIMIZE_BUTTON
      Else
        ' turn off bits for Minimize arrow
        ' button
        lStyle = lStyle And MINIMIZE_BUTTON
      End If
      ' set the new window style
      SetWindowLong TargetForm.hwnd, GWL_STYLE, lStyle
  End Select
End Sub

'***************************************************************
' VBAppendMenuItem(): This routine append a menu item to the
'        control menu. The id for the added item is returned.
'
' Parameters
' TargetForm - the form to perform the operation on
' ItemToAppend - test of menu to add
'***************************************************************
Public Function VBAppendMenuItem(ByVal TargetForm As Form, _
  ByVal ItemToAppend As String) As Integer
  Dim hSysMenu As Long
  Dim lStyle As Long
  
  VBAppendMenuItem = 0
  hSysMenu = GetSystemMenu(TargetForm.hwnd, 0&)
  If hSysMenu Then
    Call AppendMenu(hSysMenu, MF_STRING, SCOFFSET + CurrentID, ItemToAppend)
    CurrentID = CurrentID + 1
    Call DrawMenuBar(TargetForm.hwnd)
    VBAppendMenuItem = CurrentID
  End If
  End Function

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************


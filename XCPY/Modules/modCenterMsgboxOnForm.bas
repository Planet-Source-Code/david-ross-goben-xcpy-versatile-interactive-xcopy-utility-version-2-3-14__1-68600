Attribute VB_Name = "modCenterMsgBoxOnForm"
Option Explicit
'~modCenterMsgBoxOnForm.bas;
'*************************************************************************************
' modCenterMsgBoxOnForm - The CenterMsgBoxOnForm() will center a message box on a
'                         specified form, instead of the center of the screen. Simply
'                         call the CenterMsgBoxOnForm() function as you would the
'                         MsgBox() form, except you additionally specify the form
'                         that the message should be centered over.
'*************************************************************************************

Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Private Const GWL_HINSTANCE = (-6)
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOACTIVATE = &H10
Private Const HCBT_ACTIVATE = 5
Private Const WH_CBT = 5

Private hHook As Long
Private FrmhWnd As Long

Public Function CenterMsgBoxOnForm(ParentForm As Form, Msg As String, Optional Buttons As VbMsgBoxStyle = vbOKOnly, Optional Title As String = vbNullString) As VbMsgBoxResult
  Dim hInst As Long
  Dim Thread As Long

  'Set up the CBT hook
  FrmhWnd = ParentForm.hwnd
  hInst = GetWindowLong(ParentForm.hwnd, GWL_HINSTANCE)
  Thread = GetCurrentThreadId()
  hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProc1, hInst, Thread)
  'Display the message box
  CenterMsgBoxOnForm = MsgBox(Msg, Buttons, Title)
End Function

Private Function WinProc1(ByVal lMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Dim rectForm As RECT, rectMsg As RECT
  Dim X As Long, Y As Long

  'On HCBT_ACTIVATE, show the MsgBox centered over Form1
  If lMsg = HCBT_ACTIVATE Then
    'Get the coordinates of the form and the message box so that
    'you can determine where the center of the form is located
    GetWindowRect FrmhWnd, rectForm
    GetWindowRect wParam, rectMsg
    X = (rectForm.Left + (rectForm.Right - rectForm.Left) / 2) - ((rectMsg.Right - rectMsg.Left) / 2)
    Y = (rectForm.Top + (rectForm.Bottom - rectForm.Top) / 2) - ((rectMsg.Bottom - rectMsg.Top) / 2)
    'Position the msgbox
    SetWindowPos wParam, 0, X, Y, 0, 0, SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE
    'Release the CBT hook
    UnhookWindowsHookEx hHook
  End If
  WinProc1 = False
End Function

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************


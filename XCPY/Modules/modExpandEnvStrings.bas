Attribute VB_Name = "modExpandEnvStrings"
Option Explicit
'~modExpandEnvStrings.bas;
'expand environment variables
'*************************************************
' modExpandEnvStrings:
' expand environment variables declared within in a dir/file path
'*************************************************

'*************************************************
' API call used
'*************************************************
Private Declare Function ExpandEnvironmentStrings Lib "kernel32" Alias "ExpandEnvironmentStringsA" (ByVal lpSrc As String, ByVal lpDst As String, ByVal nSize As Long) As Long

Public Function ExpandEnvStrings(src As String) As String
  Dim ln As Integer, dst As String
  Dim S As String

  S = Trim$(src)                                          'clean up act
  dst = String$(500, 0)                                   'init destination length
  If Len(S) Then                                          'if something there
    ln = ExpandEnvironmentStrings(S, dst, CLng(Len(dst))) 'expand string
    ln = InStr(1, dst, vbNullChar)
    If ln > 0 And ln <= Len(dst) Then
      ExpandEnvStrings = Left$(dst, ln - 1)               'grab valid string data
    End If
  End If
End Function

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************


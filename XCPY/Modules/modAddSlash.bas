Attribute VB_Name = "modAddSlash"
Option Explicit
'~modAddSlash.bas;
'Add a terminating backslash to a drive/path if required. Also remove
'********************************************************************************
' modAddSlash: The following functions are provided:
'
' AddSlash():    Add a terminating backslash to a drive/path if required. This function
'                is useful for building paths, and the string you are working with may
'                or may not already have a backslash appended to it.
' RemoveSlash(): Remove any existing terminating backslash from a path.
'********************************************************************************

Public Function AddSlash(str As String) As String
  Dim S As String
  
  S = Trim$(str)
  If Right$(S, 1) <> "\" Then S = S & "\"
  AddSlash = S
End Function

Public Function RemoveSlash(str As String) As String
  Dim S As String
  
  S = Trim$(str)
  If Right$(S, 1) = "\" Then S = Left$(S, Len(S) - 1)
  RemoveSlash = S
End Function

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************


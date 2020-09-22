Attribute VB_Name = "modGetOSType"
Option Explicit
'~modGetOSType.bas;
'Gets the local operating system type
'***************************************************************
' modGetOSType:
' The GetOSType() function gets the local operating system type.
' Function Result: 1=NT351       2=NT40
'                  3=Win311      4=Win95
'                  5=Win98       6=Win2000
'                  7=WinME       8=WinXP
'                  9=.NET Server
'
' The IsNT() function returns TRUE if the OS is type 1,2,6, or 8
'***************************************************************

'***************************************************************
' API support routines and types and constants
'***************************************************************
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

'***************************************************************
' structure for system version information
'***************************************************************
Private Type OSVERSIONINFO
  dwOSVersionInfoSize As Long
  dwMajorVersion      As Long
  dwMinorVersion      As Long
  dwBuildNumber       As Long
  dwPlatformId        As Long
  szCSDVersion        As String * 128 'Maintenance string for PSS usage
End Type

'***************************************************************
' constants used for system version information
'***************************************************************
Private Const VER_PLATFORM_WIN32_NT = 2
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32s = 0

'***************************************************************
' get the local operating system type
' Result: 1=NT351
'         2=NT40
'         3=Win311
'         4=Win95
'         5=Win98
'         6=Win2000
'         7=WinME
'         8=WinXP
'         9=.NET Server
'***************************************************************
Public Function GetOSType() As Integer
  Dim OSV As OSVERSIONINFO
  
  OSV.dwOSVersionInfoSize = Len(OSV)                  'set size of info block
  If GetVersionEx(OSV) Then
    Select Case OSV.dwPlatformId
      Case VER_PLATFORM_WIN32s: GetOSType = 3         'win 3.11
      Case VER_PLATFORM_WIN32_WINDOWS:
        If OSV.dwMinorVersion = 0 Then
          GetOSType = 4                                'win 95
        Else
          If OSV.dwMinorVersion > 10 Then
            GetOSType = 7                              'win ME
          Else
            GetOSType = 5                              'win 98
          End If
        End If
      Case VER_PLATFORM_WIN32_NT
        Select Case OSV.dwMajorVersion
          Case 3: GetOSType = 1                        'NT 3.51
          Case 4: GetOSType = 2                        'NT 4.0
          Case 5
            If OSV.dwMinorVersion Then
              GetOSType = 8                           'WinXP (Minor = 1)
              If OSV.dwMinorVersion = 2 Then GetOSType = 9  '.NET SERVER
            Else
              GetOSType = 6                           'Win2000
            End If
          Case Else: GetOSType = 8                    'WinXP (default)
        End Select
    End Select
  End If
End Function

'***************************************************************
' get the exact local operating system type
' Result: 1=NT351
'         2=NT40
'         3=Win311
'         4=Win95
'         5=Win98
'         6=Win2000
'         7=WinME
'         8=WinXP
'         9=.NET Server
'        14=Win95 OSR2
'        15=Win98 SE
'***************************************************************
Public Function GetExactOSType() As Integer
  Dim OSV As OSVERSIONINFO
  
  OSV.dwOSVersionInfoSize = Len(OSV)                  'set size of info block
  If GetVersionEx(OSV) Then
    Select Case OSV.dwPlatformId
      Case VER_PLATFORM_WIN32s: GetExactOSType = 3    'win 3.11
      Case VER_PLATFORM_WIN32_WINDOWS:
        If OSV.dwMinorVersion = 0 Then
          If GetLoWord(OSV.dwBuildNumber) = 950 Then
            GetExactOSType = 4                        'win 95
          Else
            GetExactOSType = 14                       'win 95 OSR2
          End If
        Else
          Select Case GetLoWord(OSV.dwBuildNumber)
            Case 1998
              GetExactOSType = 5                      'win 98
            Case 2222
              GetExactOSType = 15                     'Win 98 SE
            Case 3000
              GetExactOSType = 7                      'Win ME
          End Select
        End If
      Case VER_PLATFORM_WIN32_NT
        Select Case OSV.dwMajorVersion
          Case 3: GetExactOSType = 1                  'NT 3.51
          Case 4: GetExactOSType = 2                  'NT 4.0
          Case 5
            If OSV.dwMinorVersion Then
              GetExactOSType = 8                      'WinXP
              If OSV.dwMinorVersion = 2 Then GetExactOSType = 9  '.NET SERVER
            Else
              GetExactOSType = 6                      'Win2000
            End If
          Case Else: GetExactOSType = 8               'WinXP (future expansion)
        End Select
    End Select
  End If
End Function

'*******************************************************************************
' Function Name     : IsNT
' Purpose           : Return TRUE if WinNT platform
'*******************************************************************************
Public Function IsNT() As Boolean
  Select Case GetOSType()
    Case 1, 2, 6, 8
      IsNT = True
    Case Else
      IsNT = False
  End Select
End Function

'***************************************************************
' GetLoWord(): Get low word (16 bits) from a Long
'***************************************************************
Private Function GetLoWord(ByRef LongIn As Long) As Integer
  Call CopyMemory(GetLoWord, LongIn, 2)
End Function

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************


Attribute VB_Name = "modNetPathFunctions"
Option Explicit
'~modNetPathFunctions.bas;
'Various functions to support adding/removing net connections
'********************************************************************************
' modNetPathFunctions - Various functions to support adding/removing net
'                       connections. The following functions are supported:
' GetDriveFromNetPath():This function returns the local drive path to a drive/
'                       directory/file if a network path is supplied. If the returned
'                       string still contains "\\" at its left, you may want to use
'                       the GetFreeDriveLetter() and ConnectNetDrive() functions if you
'                       need the path to contain a local drive letter.
' GetFreeDriveLetter(): This function returns a string defining an unused drive
'                       letter ("L:", for example). Blank is returned on an error.
' ConnectNetDrive():    This function connects a drive letter to a network drive path.
' DisconnectNetDrive(): This function disconnects a specified local drive letter
'                       from its network connection. True is returned if it is successful.
' GetDriveLetterType(): This function returns the drive type of the specified local drive
'                       letter (A-Z).
'                       DRIVE TYPES: 0=Unknown/invalid
'                                    1=Removable (Cartridge/Floppy)
'                                    2=Fixed     (Local hard disk)
'                                    3=Remote    (Network drive or network CDROM)
'                                    4=CDROM     (local)
'                                    5=RAMDisk
'
'EXAMPLES: if \\Source\Root is connected to X:
'  Dim NewDrv As String
'''  This example would print "X:\DavidG\Setup.Exe".
'  Debug.Print GetDriveFromNetPath("\\Source\Root\DavidG\Setup.exe")
'''This example would print "X:".
'  Debug.Print GetDriveFromNetPath("\\Source\Root")
'''get a free drive letter
'    NewDrv = GetFreeDriveLetter()
'    Debug.Print "Free Drive = " & NewDrv
'''connect it to \\source\root
'    If ConnectNetDrive(NewDrv, "\\Source\Root") Then
'      Debug.Print "Connected " & NewDrv & " to \\Source\Root"
'''show type
'      Debug.Print "Drive type is " & CStr(GetDriveLetterType(NewDrv))
'''Disconnect from the local drive share
'      If DisconnectNetDrive(NewDrv) Then
'        Debug.Print "Disconnected " & NewDrv & " from \\Source\Root"
'      End If
'    End If
'
'NOTE: At least the full drive path must be suppied. In the above example,
'      A minimum of \\Source\Root would be considered valid. \\Source would not.
'      Also, some of these functions reflect similar functions in
'      modNetworkDrives.bas, but these work without needing scrrun.dll. Those
'      featured here were designed to support the SMP/IS Setup program, but
'      can be used anywhere.
'********************************************************************************
Private Type NETRESOURCE
  dwScope As Long
  dwType As Long
  dwDisplayType As Long
  dwUsage As Long
  lpLocalName As String
  lpRemoteName As String
  lpComment As String
  lpProvider As String
End Type

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function WNetAddConnection2 Lib "mpr.dll" Alias "WNetAddConnection2A" (lpNetResource As NETRESOURCE, ByVal lpPassword As String, ByVal lpUserName As String, ByVal dwFlags As Long) As Long
Private Declare Function WNetCancelConnection2 Lib "mpr.dll" Alias "WNetCancelConnection2A" (ByVal lpName As String, ByVal dwFlags As Long, ByVal fForce As Long) As Long
Private Declare Function WNetGetConnection Lib "mpr.dll" Alias "WNetGetConnectionA" (ByVal lpszLocalName As String, ByVal lpszRemoteName As String, cbRemoteName As Long) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

Private Const MAX_COMPUTERNAME_LENGTH = 15
Private Const MAX_PATH = 260
Private Const DRIVE_REMOTE = 4
Private Const CONNECT_UPDATE_PROFILE = &H1
Private Const RESOURCE_PUBLICNET = &H2
Private Const RESOURCETYPE_DISK = &H1
Private Const RESOURCEDISPLAYTYPE_SHARE = &H3
Private Const RESOURCEUSAGE_CONNECTABLE = &H1

'*******************************************************************************
' Function Name     : GetDriveFromNetPath
' Purpose           : Extract the shared drive letter from a netdrive
'*******************************************************************************
Public Function GetDriveFromNetPath(NetPath As String) As String
  Dim Idx As Integer, Idy As Integer, I As Integer
  Dim Rlen As Long
  Dim RemoteName As String, drv As String, Npath As String, FPath As String
  Dim LocalComputerName As String
'
' check for valid data
'
  Npath = Trim$(NetPath)                        'get net path
  If Left$(Npath, 2) = "\\" Then                'if a netpath specified...
'
' get local computer name in case the netpath is also a local drive
'
    drv = String$(MAX_COMPUTERNAME_LENGTH + 1, 32)
    Call GetComputerName(drv, MAX_COMPUTERNAME_LENGTH + 1)
    LocalComputerName = "\\" & Left$(drv, InStr(1, drv, vbNullChar) - 1) & "\"
    
    Idx = InStr(3, Npath, "\")                    'find root\drive sep
    If Idx = 0 Then Exit Function                 'nada, so invalid
    Idy = InStr(Idx + 1, Npath, "\")              'find possible slash beyond drive
    FPath = vbNullString                          'init to no data dribbling
    If Idy Then
      FPath = Mid$(Npath, Idy)                    'save extra dribble text
      Npath = Left$(Npath, Idy - 1)               'if found, strip anything else off
    End If
'
' scan local system for a local drive connected to the netpath
'
    For Idx = 2 To 26                                   'check drives B thru Z
      drv = Chr$(Idx + 65) & ":"                        'build drive
      If GetDriveType(drv) > 1 Then                     'valid? (0-1 are unknown)
        Rlen = MAX_PATH                                 'init to max
        RemoteName = String$(MAX_PATH, " ")
        Call WNetGetConnection(drv, RemoteName, Rlen)   'get connection name for drive
        I = InStr(1, RemoteName, vbNullChar)            'find null
        If I Then                                       'found it?
          RemoteName = Left$(RemoteName, I - 1)         'yes, so grab netpath
        Else
          RemoteName = Trim$(RemoteName)                'else give it a trim
        End If
        ' get local drive info if data exists
        If Len(RemoteName) = 0 Then RemoteName = LocalComputerName & Left$(drv, 1)
        If StrComp(RemoteName, Npath, vbTextCompare) = 0 Then 'match?
          GetDriveFromNetPath = drv & FPath             'yes, so return drive with any extra data
          Exit Function
        End If
      End If
    Next Idx                                            'scan all possible drives
  End If
  GetDriveFromNetPath = Npath                         'default to master
End Function                                          'return blank if we get this far

'*******************************************************************************
' Function Name     : GetFreeDriveLetter
' Purpose           : Get a free drive letter. Return blank if none
'*******************************************************************************
Public Function GetFreeDriveLetter(Optional StartFrom As String = "C") As String
  Dim StartDrv As Integer, Idx As Integer, drv As String, Rlen As Long, str As String
  Dim I As Integer
'
' compute start drive for search
'
  If Len(StartFrom) Then StartDrv = Asc(UCase$(StartFrom))  'get asc of start drive
  StartDrv = StartDrv - 65                                  'drop "A" offset
  If StartDrv < 2 Or StartDrv > 25 Then StartDrv = 2        '<"C:", or > "Y:"?
  For Idx = StartDrv To 26                                  ' check to drive Z
    drv = Chr$(Idx + 65) & ":"                              'build drive
    I = GetDriveType(drv)
    If GetDriveType(drv) < 2 Then                           'invalid?
      GetFreeDriveLetter = drv                              'yes, grab it
      Exit Function
    End If
  Next Idx
End Function

'*******************************************************************************
' Function Name     : ConnectNetDrive
' Purpose           : Connect a local drive letter to a network drive
'*******************************************************************************
Public Function ConnectNetDrive(LocalDrive As String, NetPath As String) As Boolean
  Dim drv As String, Idx As Integer, Npath As String, Idy As Integer
  Dim nR As NETRESOURCE
  
  drv = UCase$(Left$(LocalDrive, 1))              'build drive mask
  If drv < "A" Or drv > "Z" Then Exit Function
  drv = drv & ":"
  If GetDriveType(drv) > 1 Then Exit Function     'drive already assigned
  Npath = Trim$(NetPath)                          'clean up input
  If Left$(Npath, 2) <> "\\" Then Exit Function   'netpath not valid
  Idx = InStr(3, Npath, "\")                      'find root\drive sep
  If Idx = 0 Then Exit Function                   'nada, so invalid
  Idy = InStr(Idx + 1, Npath, "\")
  If Idy Then Npath = Left$(Npath, Idy - 1)
  If Len(Dir(Npath & "\*.*")) = 0 And Len(Dir(Npath & "\*.*", vbDirectory)) = 0 Then Exit Function
  nR.dwScope = RESOURCE_PUBLICNET
  nR.dwType = RESOURCETYPE_DISK
  nR.dwUsage = RESOURCEUSAGE_CONNECTABLE
  nR.lpComment = vbNullString
  nR.lpProvider = vbNullString
  nR.lpLocalName = drv
  nR.lpRemoteName = Npath
  nR.lpProvider = vbNullString
  ConnectNetDrive = (WNetAddConnection2(nR, vbNullString, vbNullString, CONNECT_UPDATE_PROFILE) = 0)
End Function

'*******************************************************************************
' Function Name     : DisconnectNetDrive
' Purpose           : Disconnect a local drive letter from a network drive
'*******************************************************************************
Public Function DisconnectNetDrive(LocalDrive As String) As Boolean
  Dim drv As String
  
  drv = UCase$(Left$(LocalDrive, 1))              'build drive mask
  If drv < "A" Or drv > "Z" Then Exit Function
  drv = drv & ":"
  If GetDriveType(drv) <> DRIVE_REMOTE Then Exit Function
  DisconnectNetDrive = (WNetCancelConnection2(drv, CONNECT_UPDATE_PROFILE, 0&) = 0)
End Function

'*******************************************************************************
' Function Name     : GetDriveLetterType
' Purpose           : Return the drive type of a drive
'*******************************************************************************
Public Function GetDriveLetterType(LocalDrive As String) As Integer
  Dim drv As String, I As Integer
  
  If Left$(LocalDrive, 2) = "\\" Then           'if netspec
    drv = GetDriveFromNetPath(LocalDrive)       'derive netpath to a shared drive
    If Len(drv) = 0 Then                        'shared on a drive?
      GetDriveLetterType = 3                    'no, so assume remote
      Exit Function
    End If
    drv = UCase$(Left$(drv, 1))                 'build drive mask
  Else
    drv = UCase$(Left$(LocalDrive, 1))          'build drive mask
  End If
  If drv < "A" Or drv > "Z" Then Exit Function  'bad info
  drv = drv & ":"
  I = GetDriveType(drv)                         'grab type
  If I Then GetDriveLetterType = I - 1          'allow 0 to remain 0. Ajust others
End Function

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************


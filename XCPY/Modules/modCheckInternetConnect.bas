Attribute VB_Name = "modCheckInternetConnect"
Option Explicit
'~modCheckInternetConnect.bas;modExpandEnvStrings.bas;modGetWinSysCurDir.bas;
'Check internet connections on the local system
'**************************************************************************
' modCheckInternetConnect: Check internet connections on the local system
'
'This module provides the following functions:
'
' CheckInternetConnect():     Check if connection exists. Returns TRUE if so, else FALSE
' CheckInternetConnectType(): Get connection type flags as integer
' CheckInternetConnectName(): get connection name as a string (i.e., LAN)
'
'CheckInternetConnectType() return flags (do binary AND on values to verify):
'  1=Modem
'  2=Lan
'  4=Proxy
'  8=RAS
' 16=Connection configured (should always be set on valid connections)
' 32=Offline
'
' NOTE: This module requires modExpandEnvStrings.bas.
' NOTE: This module requires modGetWinSysCurDir.bas.
'**************************************************************************

'**************************************************************************
' API call and structure required to gather information for an internet connection
'**************************************************************************
Public Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef lpdwFlags As Long, ByVal dwReserved As Long) As Long

Private Enum EIGCInternetConnectionState
   INTERNET_CONNECTION_MODEM = &H1&
   INTERNET_CONNECTION_LAN = &H2&
   INTERNET_CONNECTION_PROXY = &H4&
   INTERNET_RAS_INSTALLED = &H10&
   INTERNET_CONNECTION_MODEM_BUSY = &H8
   INTERNET_CONNECTION_OFFLINE = &H20&
   INTERNET_CONNECTION_CONFIGURED = &H40&
End Enum

'*********************************************************
' Check if connection exists. Returns TRUE if so, else FALSE
'*********************************************************
Public Function CheckInternetConnect() As Boolean
  Dim eR As EIGCInternetConnectionState
  Dim sName As String

' Determine whether we have a connection:
  CheckInternetConnect = InternetConnected(eR, sName)

End Function

'*********************************************************
' Get connection type flags
'
' Returns a binary integer. By ANDing 1, 2, 4, 8, or 16, you can check specific states
' A result of 0 means no connection. You might want to check for an internet connection
' first using CheckInternetConnect() above.
'
' Flags: 1=Modem
'        2=Lan
'        4=Proxy
'        8=RAS
'       16=Connection configured (should always be set on valid connections)
'
'*********************************************************
Public Function GetInternetConnectType() As Integer
  Dim eR As EIGCInternetConnectionState
  Dim sName As String
  Dim Result As Integer

  Result = 0
  If InternetConnected(eR, sName) Then
    If (eR And INTERNET_CONNECTION_MODEM) = INTERNET_CONNECTION_MODEM Then Result = 1
    If (eR And INTERNET_CONNECTION_LAN) = INTERNET_CONNECTION_LAN Then Result = Result Or 2
    If (eR And INTERNET_CONNECTION_PROXY) = INTERNET_CONNECTION_PROXY Then Result = Result Or 4
    If (eR And INTERNET_RAS_INSTALLED) = INTERNET_RAS_INSTALLED Then Result = Result Or 8
    If (eR And INTERNET_CONNECTION_CONFIGURED) = INTERNET_CONNECTION_CONFIGURED Then Result = Result Or 16
    If (eR And INTERNET_CONNECTION_OFFLINE) = INTERNET_CONNECTION_OFFLINE Then Result = Result Or 32
  End If
  GetInternetConnectType = Result

End Function

'*********************************************************
' get connection name
'
' You might want to check for an internet connection
' first using CheckInternetConnect() above.
'*********************************************************
Public Function GetInternetConnectName() As String
  Dim eR As EIGCInternetConnectionState
  Dim sName As String

' Determine whether we have a connection:
  If InternetConnected(eR, sName) Then
    GetInternetConnectName = sName
  Else
    GetInternetConnectName = vbNullString
  End If

End Function

'********************************************************
' internal function to check the internet and grab information about the connection
'********************************************************
Private Property Get InternetConnected(ByRef eConnectionInfo As EIGCInternetConnectionState, Optional ByRef sConnectionName As String = vbNullString) As Boolean
  Dim dwFlags As Long
  Dim sNameBuf As String, S As String
  Dim lR As Long
'
' see if wininet.dll file is available
'
  S = GetSystemDir()
  If CBool(Len(Dir$(S & "\wininet.dll"))) Then
    On Error Resume Next
    lR = InternetGetConnectedState(dwFlags, 0&)
    If Err.Number Then lR = 0
    On Error GoTo 0
  End If
'
' check results
'
  sNameBuf = vbNullString
  If lR = 1 Then
    If dwFlags And INTERNET_RAS_INSTALLED Then sNameBuf = sNameBuf & "RAS "
    If dwFlags And INTERNET_CONNECTION_MODEM Then sNameBuf = sNameBuf & "Modem "
    If dwFlags And INTERNET_CONNECTION_LAN Then sNameBuf = sNameBuf & "LAN "
    If dwFlags And INTERNET_CONNECTION_PROXY Then sNameBuf = sNameBuf & "Proxy "
    If dwFlags And INTERNET_CONNECTION_OFFLINE Then sNameBuf = sNameBuf & "- Offline "
    If (dwFlags And INTERNET_CONNECTION_CONFIGURED) <> INTERNET_CONNECTION_CONFIGURED Then sNameBuf = sNameBuf & "- Not Configured"
  End If
  If Len(sNameBuf) = 0 Then sNameBuf = "NOT CONFIGURED"
  If Len(sConnectionName) = 0 Then sConnectionName = Trim$(sNameBuf)
  eConnectionInfo = dwFlags
  InternetConnected = (lR = 1)

End Property

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************


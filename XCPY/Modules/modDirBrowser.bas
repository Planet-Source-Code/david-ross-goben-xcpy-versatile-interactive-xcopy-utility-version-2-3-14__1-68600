Attribute VB_Name = "modDirBrowser"
Option Explicit
'~modDirBrowser.bas;modGetOsType.bas;
'Opens a system dialog for browsing for a folder
'**************************************************************************
' modDirBrowser - The DirBrowser() function opens a system dialog for browsing
'                 for a folder without using a large OCX file.
'EXAMPLE:
'  Dim S As String
'  S = DirBrowser(frmMain.hWnd, ViewDirsOnly, "Select Install Path", CurDir$)
'  If CBool(Len(S)) Then
'    Debug.Print "Install Path " & S
'  Else
'    Debug.Print "No Install Path selected"
'  End If
'
' hWndOwner can normally be: Me.Hwnd
' sPrompt is a 'memory-jogger' prompt to display on the browser.
'
' NOTE: This modules used the GetOsType.bas module in order to enable the
'       new user interface available on W2K and XP.
'**************************************************************************

' Frees a block of task memory previously allocated through a call to the
'   CoTaskMemAlloc or CoTaskMemRealloc function.
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
' Appends one string to another.
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" _
       (ByVal lpString1 As String, ByVal lpString2 As String) As Long
' Displays a dialog box enabling the user to select a shell folder.
'   The calling application is responsible for freeing the returned
'   item identifier list by using the shell's task allocator.
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
' Converts an item identifier list to a file system path.
Private Declare Function SHGetPathFromIDList Lib "shell32" _
       (ByVal pidList As Long, ByVal lpBuffer As String) As Long
'copy memory
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
       (pDest As Any, pSource As Any, ByVal dwLength As Long)

Public Enum BrowseOption
  ViewAll                'windows 2000/XP/98B only (otherwise like ViewDirsOnly)
  ViewDirsOnly
End Enum

Private Type BrowseInfo
  hwndOwner      As Long 'Handle to the owner window
  pidlRoot       As Long 'Address of an ITEMIDLIST structure specifying the location
                         'of the root folder from which to browse. Only the specified
                         'folder and its subfolders appear in the dialog box. This member
                         'can be NULL; in that case, the namespace root (the desktop folder) is used.
  pszDisplayName As Long 'Address of a buffer to receive the display name of the folder
                         'selected by the user. The size of this buffer is assumed to be MAX_PATH bytes.
  lpszTitle      As Long 'Address of a null-terminated string that is displayed above the
                         'tree view control in the dialog box. This string can be used to
                         'specify instructions to the user.
  ulFlags        As Long 'Flags specifying the options for the dialog box.  This can include
                         'zero or a combination of the below values:
  lpfnCallback   As Long 'Address of an application-defined function that the dialog box calls
                         'when an event occurs. This member can be NULL.
  lParam         As Long 'Application-defined value that the dialog box passes to the callback
                         'function, if one is specified.
  iImage         As Long 'Variable to receive the image associated with the selected folder.
                         'The image is specified as an index to the system image list.
End Type

Private Const MAX_PATH = 260                  'max file path length
Private Const BIF_EDITBOX = &H10              'INCLUDE FOLDER EDIT BOX
Private Const BIF_BROWSEINCLUDEFILES = &H4000 'The browse dialog will display files as well as folders.
Private Const BIF_RETURNONLYFSDIRS = &H1      'Only return file system directories. If the user selects
                                              'folders that are not part of the file system, the OK
                                              'button is grayed.
'
' options available to w2k/xp
'
Private Const BIF_NEWDIALOGSTYLE = &H40
Private Const BIF_USENEWUI = BIF_EDITBOX Or BIF_NEWDIALOGSTYLE
'
' set startup path support
'
Private Const LMEM_FIXED = &H0
Private Const LMEM_ZEROINIT = &H40
Private Const Lptr = (LMEM_FIXED Or LMEM_ZEROINIT)
Private Const BFFM_INITIALIZED = 1
Private Const WM_USER = &H400
Private Const BFFM_SETSELECTIONA As Long = (WM_USER + 102)

Private Declare Function LocalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal uBytes As Long) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
       (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private PathStart As String                   'start path for browsing

'**************************************************************************
' Opens the system dialog for browsing for a folder.
'
' hWndOwner can normally be: Me.Hwnd
' sPrompt is a 'memory-jogger' prompt to display on the browser.
'**************************************************************************
Public Function DirBrowser(hwndOwner As Long, BrowseType As BrowseOption, _
                           sPrompt As String, Optional StartPath As String) As String
  Dim iNull    As Integer
  Dim lpIDList As Long
  Dim lResult  As Long
  Dim lpPath   As Long
  Dim ulFlg    As Long
  Dim lPmt     As Long
  Dim sPath    As String
  Dim Pmt      As String
  Dim udtBI    As BrowseInfo
'
' set up information block
'
  PathStart = Trim$(StartPath)
  If CBool(Len(PathStart)) Then
    If CBool(Len(Dir$(PathStart, vbDirectory))) Then
      If Not CBool(GetAttr(PathStart) And vbDirectory) Then PathStart = vbNullString
    Else
      PathStart = vbNullString
    End If
  End If
  
  With udtBI
    .hwndOwner = hwndOwner                    'owner handle
    .pidlRoot = 0
    Pmt = sPrompt & Chr$(0)                   'set local prompt with null terminator
    lPmt = LocalAlloc(Lptr, Len(Pmt))         'set aside space
    CopyMemory ByVal lPmt, ByVal Pmt, Len(Pmt) 'copy data
    .lpszTitle = lPmt                         'set pointer
    If BrowseType = ViewAll Then
      ulFlg = BIF_BROWSEINCLUDEFILES Or BIF_RETURNONLYFSDIRS
    Else
      ulFlg = BIF_RETURNONLYFSDIRS
    End If
    
    Select Case GetOSType()
      Case 6, Is > 7                          'w2k, xp
        ulFlg = ulFlg Or BIF_USENEWUI         'allow using new interface
      Case Else
        ulFlg = ulFlg Or BIF_EDITBOX          'add edit box for everyone else
    End Select
    .ulFlags = ulFlg
    
    If CBool(Len(PathStart)) Then
      PathStart = PathStart & Chr$(0)         'set local copy with null terminator
      lpPath = LocalAlloc(Lptr, Len(PathStart))
      CopyMemory ByVal lpPath, ByVal PathStart, Len(PathStart)
      .lParam = lpPath
      .lpfnCallback = FARPROC(AddressOf BrowseCallbackProcStr)
    Else
      lpPath = 0
      .lParam = 0
      .lpfnCallback = 0
    End If
  End With
'
' browse...
'
  lpIDList = SHBrowseForFolder(udtBI)         'browse and allocate resources for data
'
' free set-aside memory
'
  Call LocalFree(lPmt)
  If CBool(lpPath) Then Call LocalFree(lpPath)
'
' get selection path
'
  If CBool(lpIDList) Then
    sPath = String$(MAX_PATH, 0)              'set up receiving buffer
    lResult = SHGetPathFromIDList(lpIDList, sPath)
    Call CoTaskMemFree(lpIDList)              'free resources
    iNull = InStr(sPath, vbNullChar)          'put path to VB string
    If iNull Then sPath = Left$(sPath, iNull - 1)
  Else
    sPath = vbNullString
  End If
'
' stuff path as return value
'
  DirBrowser = sPath
End Function

'*******************************************************************************
' Function Name     : BrowseCallbackProcStr
' Purpose           :   Callback for the Browse STRING method.
'                   :
'                   :   On initialization, set the dialog's
'                   :   pre-selected folder from the pointer
'                   :   to the path allocated as bi.lParam,
'                   :   passed back to the callback as lpData param.
'*******************************************************************************
Public Function BrowseCallbackProcStr(ByVal hwnd As Long, _
                                      ByVal uMsg As Long, _
                                      ByVal lParam As Long, _
                                      ByVal lpData As Long) As Long
   Select Case uMsg
     Case BFFM_INITIALIZED
       Call SendMessage(hwnd, BFFM_SETSELECTIONA, 1&, ByVal lpData)
   End Select
End Function

'*******************************************************************************
' Function Name     : FARPROC
' Purpose           :   A dummy procedure that receives and returns
'                   :   the value of the AddressOf operator.
'                   :
'                   :   This workaround is needed because you can't assign
'                   :   AddressOf directly to a member of a user-
'                   :   defined type, but you can assign it to another
'                   :   long and use that (as returned here)
'                   :
'*******************************************************************************
Private Function FARPROC(pfn As Long) As Long
  FARPROC = pfn
End Function

'******************************************************************************
' Copyright 1990-2007 David Ross Goben. All rights reserved.
'******************************************************************************

